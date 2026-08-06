# 招投标公告语义清洗（固定 29 字段）

将招标、采购、中标、成交或废标公告从 JSON、CSV、TSV、XLSX、XLSM、XLS 等原始数据中提取为固定的 29 个字段。优先依据公告正文提取可核验事实，支持直接处理本地文件，也支持从已有八爪鱼任务启动、等待、全量导出后再清洗。

> 仓库不包含 API Key，也不会保存或输出密钥。仅在“已有任务采集”路径中需要你在本地提供 `BAZHUAYU_API_KEY`。

## 适用场景

- 将一份或多份招投标公告转换为统一字段，便于筛选、入库和后续分析。
- 合并不同网站或不同采集任务的公告，并按来源链接或内容特征去重。
- 从正文中提取项目规模、范围、工期、金额、文件获取时间、投标/开标时间、招标人与代理机构联系方式等信息。
- 对八爪鱼账户内已经存在的任务执行“启动 → 等待 → 全量导出 → 清洗”的完整流程。

不适用于根据常识补写公告缺失内容。没有原文或原始列依据的字段一律留空。

## 输出：固定 29 字段

| 类别 | 字段 |
| --- | --- |
| 基本信息 | 原标题、来源链接、项目名称、招标编号、项目执行地址、项目规模、招标范围、项目工期、招标金额 |
| 时间与开标 | 获取招标文件时间、投标截止时间、开标时间、开标方式、开标地点 |
| 招标人 | 招标人、招标联系人、招标人电话、招标单位、招标人邮箱、招标人地址 |
| 代理机构 | 招标代理机构名称、招标代理机构地址、招标代理机构联系人、招标代理机构电话、招标代理机构邮箱 |
| 结果信息 | 中标人、中标候选人、中标金额、中标公告时间 |

金额只提取预算、最高限价、控制价或招标控制价，并统一换算为人民币“元”。标书费、保证金、代理费和数量不会被当作招标金额。招标/采购公告不会推断中标人或中标金额；只有中标/成交公告才填写对应字段。

## 安装

需要 Python 3.10+。JSON、CSV、TSV 的处理只使用标准库；处理 Excel 时安装依赖：

```powershell
python -m pip install -r requirements.txt
```

- `openpyxl`：读取 `.xlsx` / `.xlsm`
- `xlrd`：读取旧版 `.xls`

主入口始终是：

```text
scripts/tender_workflow.py
```

## 用法一：清洗本地文件（不需要 API Key）

```powershell
python scripts/tender_workflow.py files `
  --input "D:\data\notices.xlsx" "D:\data\extra.csv" `
  --output-dir "D:\data\cleaned" `
  --name "招标公告_29字段" `
  --quiet
```

常用参数：

| 参数 | 说明 |
| --- | --- |
| `--input` / `-i` | 一个或多个 JSON、CSV、TSV、XLSX、XLSM、XLS 文件（必填） |
| `--output-dir` / `-o` | 输出目录；默认当前目录 |
| `--name` | 输出文件名，不含扩展名 |
| `--sheet` | Excel 工作表名称；默认使用活动工作表 |
| `--keep-duplicates` | 保留重复公告；默认自动去重 |
| `--quiet` | 不输出逐条处理进度 |

默认优先用来源链接去重；链接缺失时，使用原标题、发布时间和正文前缀组合去重。输出为 UTF-8-BOM CSV（可直接用 Excel 打开）和 UTF-8 JSON；两者均固定为 29 列/字段。

## 用法二：从已有任务采集并清洗（需要 API Key）

此路径只操作账户中已存在的任务，不会创建新模板。请先在当前 PowerShell 会话或密钥管理服务中安全配置密钥：

```powershell
$env:BAZHUAYU_API_KEY = "你的真实密钥（不要写入文件或提交到 Git）"
```

然后按任务名或 `taskId` 执行：

```powershell
python scripts/tender_workflow.py collect `
  --task-name "任务名一" `
  --task-name "任务名二" `
  --output-dir "D:\data\tender-run" `
  --quiet
```

也可按 ID：

```powershell
python scripts/tender_workflow.py collect `
  --task-id "task-id-1" `
  --output-dir "D:\data\tender-run"
```

采集流程会：

1. 初始化 Streamable HTTP MCP 连接，使用 `x-api-key` 鉴权。
2. 分页读取任务列表并在本地做唯一精确匹配；不会依赖可能误判中文关键词的服务端模糊搜索。
3. 所有任务均成功解析后才启动；已经运行的任务不重复启动。
4. 轮询任务状态；仅对完成的任务发起 JSON 全量导出。
5. 下载 `exportFileUrl` 指向的完整文件。预览 `sampleData` 不会被误当作全量数据。
6. 分别清洗每个成功任务，并把成功导出文件与可选本地文件合并、去重和清洗。
7. 在输出目录写入 `manifest.json`、`raw/` 原始导出文件和 `clean/` 清洗结果。

`collect` 的补充参数包括：

| 参数 | 说明 |
| --- | --- |
| `--task-name` | 已有任务名，可重复传入 |
| `--task-id` | 已有任务 ID，可重复传入 |
| `--input` | 与导出结果一同合并清洗的本地文件 |
| `--resolve-only` | 只解析任务，不启动或导出 |
| `--task-timeout` | 等待任务完成时间，默认 1800 秒 |
| `--export-timeout` | 等待导出完成时间，默认 600 秒 |
| `--poll-interval` | 轮询间隔，默认 10 秒 |
| `--mcp-url` | MCP 服务地址；默认 `https://mcp.bazhuayu.com` |

请避免使用 `--api-key` 直接把密钥写进命令历史；优先采用 `BAZHUAYU_API_KEY` 环境变量。不要把 API Key 放在 `Authorization: Bearer`，直连接口使用 `x-api-key`。

## 清洗规则与数据边界

- 公告正文优先于标题和列表元数据；同一行有多个正文列时，选取最长正文并将短文本列作为补充上下文。
- 标题、编号、发布时间与截止时间等结构化列仅用于补充，不能覆盖正文中更明确的值。
- 来源链接仅保留输入中已有的链接，或正文明确出现的 URL；绝不按平台名称猜测链接。
- 单条记录失败会被隔离，不会导致其他公告被丢弃。
- 输出结束时会核验 CSV/JSON 行数一致、字段数为 29、金额单位为元，并报告输入行数、去重数和字段覆盖率。

完整字段口径见 [references/29-fields.md](references/29-fields.md)，MCP 连接或导出故障的处理见 [references/mcp-troubleshooting.md](references/mcp-troubleshooting.md)。

## 快速验证

仓库带有一份示例输入。运行以下命令会生成示例的 CSV 和 JSON：

```powershell
python scripts/tender_workflow.py files `
  --input assets/sample_input.json `
  --output-dir .\example-output `
  --name sample_clean `
  --quiet
```

预期生成 15 条清洗记录，每条固定 29 个字段。

## 目录结构

```text
SKILL.md                    供 AI 使用的任务工作流和边界
agents/openai.yaml          Skill 的界面信息
scripts/tender_workflow.py  统一入口：collect / files
scripts/tender_clean.py     读取、语义提取、去重、校验与 CSV/JSON 输出
references/29-fields.md     29 字段的语义口径
references/mcp-troubleshooting.md  MCP 协议与故障处理
assets/sample_input.json    本地文件清洗示例
assets/sample_output.csv    对应的参考输出
requirements.txt            Excel 支持所需依赖
```

## 安全说明

- 本仓库不含 API Key、授权令牌或账户信息。
- 不要提交 `.env`、本地运行目录、原始导出数据或包含联系人信息的结果文件。
- MCP 连接报错、HTTP 401/403、异步导出等待等场景，请参考故障处理文档；反馈问题时务必对密钥、会话 ID 和临时下载链接脱敏。

## 许可

当前项目未附带开源许可证。若计划公开分发或接受外部贡献，请先补充适用的许可证。
