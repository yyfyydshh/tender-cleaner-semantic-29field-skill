# 八爪鱼 MCP 协议与故障处理

## 目录

1. 正确连接参数
2. 任务解析与状态机
3. 导出规则
4. 已知故障与处理
5. 研发反馈模板

## 1. 正确连接参数

- URL：`https://mcp.bazhuayu.com`
- Transport：Streamable HTTP
- API Key 请求头：`x-api-key: <用户提供的真实Key>`
- `Content-Type`：`application/json`
- `Accept`：`application/json, text/event-stream`
- MCP protocol：`2025-03-26`

初始化顺序：

1. `initialize`
2. 保存响应头 `Mcp-Session-Id`
3. 后续请求携带 `Mcp-Session-Id` 与 `MCP-Protocol-Version`
4. 发送 `notifications/initialized`
5. 调用 `tools/call`

不要把直接 API Key 放在 `Authorization: Bearer`。不要期待命令字符串中的 `<API_KEY>` 自动替换。

## 2. 任务解析与状态机

任务名默认是账户现有任务。`search_tasks(keyword=中文)` 在部分服务版本可能被错误判定为通配符或恶意内容，因此使用：

1. `search_tasks(page=N, size=50)` 遍历 `totalPages`
2. 客户端对 `taskName` 做 NFKC、空白归一化
3. 必须唯一精确匹配
4. 全部任务名解析成功后再启动

启动：

```json
{"name":"start_or_stop_task","arguments":{"taskId":"...","action":"start"}}
```

轮询：

```json
{"name":"search_tasks","arguments":{"taskIds":["..."],"page":1,"size":10}}
```

常见状态：`运行中/Running`、`已完成/Completed/Finished`、`失败/Failed`、`已停止/Stopped`、`就绪`。

## 3. 导出规则

任务完成后调用：

```json
{
  "name": "export_data",
  "arguments": {"taskId":"...","exportFileType":"JSON","previewRows":5}
}
```

- `executing`、`collecting`、`exporting`：等待后重试。
- `exported`：下载 `exportFileUrl`。
- `no_data`：记录无数据，不进入成功清洗。
- `sampleData` 仅用于预览；除非 `dataTotal <= sampleData.length` 且没有文件链接，否则不能当成全量文件。
- 导出链接可能是临时链接，应立即下载到运行目录，同时在 manifest 中保留原链接。

## 4. 已知故障与处理

| 现象 | 原因 | 处理 |
|---|---|---|
| `-32603 Internal error` | Agent 内置连接器链路异常 | 切换 `scripts/tender_workflow.py collect` 标准直连 |
| HTTP 401/403 | Key 无效、过期或鉴权头错误 | 确认真实 Key；直连使用 `x-api-key` |
| “only wildcards / malicious content” | 服务端中文 keyword 校验误判 | 不传 keyword；分页全量读取、本地匹配 |
| 参数 expected record, received array | 客户端序列化参数错误 | `arguments` 必须是 JSON 对象，避免使用 PowerShell 保留变量 `$args` |
| 启动后仍显示旧“已完成” | 状态切换延迟或任务很快完成 | 启动请求成功后短暂等待，再用 taskId 轮询 |
| 导出一直 collecting/exporting | 异步导出尚未完成 | 按间隔轮询，直至超时；不要重启任务 |
| CSV 乱码 | Excel 编码识别 | 使用脚本生成的 UTF-8-BOM CSV |
| 项目名/联系人吞入整段正文 | 网页正文被压成单行 | 清洗器自动补标签边界并限制实体长度 |
| `50万` 未识别 | 公告使用“项目预算/预算控制价” | 清洗器已覆盖项目预算、采购预算、控制价并换算元 |

## 5. 研发反馈模板

```text
环境：<Agent/客户端名称与版本>
MCP URL：https://mcp.bazhuayu.com
协议：2025-03-26 Streamable HTTP
鉴权：x-api-key（密钥已脱敏）
initialize：成功/失败，serverInfo=<...>
失败工具：<search_tasks/start_or_stop_task/export_data>
请求参数（脱敏）：<JSON>
返回：HTTP <code> / MCP <code>，message=<...>
复现时间：<时区与时间>
对照结果：标准直连脚本成功/失败；其他 Agent 成功/失败
```

禁止在反馈中粘贴完整 API Key、Mcp-Session-Id 或临时导出中的敏感数据。
