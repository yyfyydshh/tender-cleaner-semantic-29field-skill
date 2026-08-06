---
name: tender-cleaner
description: 将招投标公告按公告正文语义清洗为固定29字段，支持两条路径：(1) 给出一个或多个八爪鱼现有任务名/taskId，启动、等待、全量导出后清洗；(2) 提供 JSON、CSV、TSV、XLSX、XLSM 或 XLS，直接语义清洗。用于招标公告结构化、项目筛选前的数据准备、跨站点公告合并去重。优先提取正文中的项目规模、范围、工期、文件获取、截止/开标、联系人及地址；仅保留可核验信息，不补造未披露字段。采集路径需用户自行提供 API Key，Skill 不保存密钥。
---

# 招投标公告语义清洗

使用 `scripts/tender_workflow.py` 作为唯一入口。它提供：

- `collect`：八爪鱼既有任务 → 启动 → 轮询 → 全量 JSON 导出 → 29字段语义清洗。
- `files`：本地 JSON/CSV/TSV/Excel → 29字段语义清洗。

先读取 `references/29-fields.md`，了解字段口径、事实边界和覆盖率检查。仅当 MCP 连接、鉴权、任务解析或导出失败时，读取 `references/mcp-troubleshooting.md`。

## 处理原则

1. 以公告正文为第一事实来源；同一行有“正文”“文本”“文本1/文本2”等多列时，优先最长的正文列，并把其余短元数据合并为上下文。
2. 优先提取显式标签及其所在段落，例如“工程地点”“总建筑面积”“招标范围”“计划工期”“招标文件获取时间”“联系方式”。
3. 标题、编号、发布时间和截止时间等结构化列只作补充，不能覆盖正文中更明确的值。
4. 招标金额仅取预算、最高限价、控制价或招标控制价；统一换算为人民币元。不得把标书费、保证金、代理费或数量当作招标金额。
5. 来源链接只保留原始输入链接，或正文中明确出现的 URL；不得根据平台名称推测详情页链接。
6. 招标/采购公告不推断中标人、中标候选人和中标金额；缺失字段填空字符串。
7. 对每条记录隔离失败。不要因单条公告字段不完整而丢弃其他可清洗记录。

## 路径选择

- 用户提供任务名或 taskId：使用 `collect`。
- 用户提供本地文件：使用 `files`。
- 两者同时提供：用 `collect --input ...`，把本地文件与导出数据合并后去重清洗。

除非用户明确要求创建模板，任务名默认指其账户中的既有任务；不要调用模板创建接口。

## 本地文件清洗

接受 `.json`、`.csv`、`.tsv`、`.xlsx`、`.xlsm`、`.xls`，可同时处理多个文件。

```powershell
python "<skill-root>\scripts\tender_workflow.py" files `
  --input "<absolute-a.xlsx>" "<absolute-b.csv>" `
  --output-dir "<absolute-output-dir>" `
  --name "招标公告_语义清洗" `
  --quiet
```

- 指定工作表：增加 `--sheet "工作表名"`。
- 默认按来源链接去重；没有链接时按标题、发布时间和正文前缀去重。
- 若 Excel 依赖缺失，运行 `python -m pip install -r "<skill-root>\requirements.txt"`。

## 采集后清洗

仅在此路径需要 API Key。API Key 必须来自用户或当前进程的 `BAZHUAYU_API_KEY`，不得写入脚本、输出文件、命令记录或回复正文。

```powershell
python "<skill-root>\scripts\tender_workflow.py" collect `
  --task-name "任务名一" `
  --task-name "任务名二" `
  --output-dir "<absolute-output-dir>" `
  --quiet
```

执行状态机：先分页读取并本地精确匹配全部任务名；全部匹配成功后才启动；已运行任务不重复启动；通过 `search_tasks` 轮询终态；只在成功后下载 `exportFileUrl` 的全量文件并清洗。不要把预览 `sampleData` 当成全量结果，也不要对 `start_or_stop_task` 使用 `poll` 动作。

## 输出与验收

- 输出 UTF-8-BOM CSV（Excel 可直接打开）和 UTF-8 JSON；均为固定29字段。
- 完成前检查：CSV/JSON 行数一致、列数为29、金额字段为元、输入行数/去重数已记录。
- 报告输出路径、原始行数、去重后行数和字段覆盖情况。
- 若“原标题”“项目名称”或“来源链接”为空，先检查源列映射与正文；来源链接确实未披露时可为空，但必须说明原因。
