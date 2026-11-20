# 智能体接入

microsoft-wps-live mcp 使用标准 stdio MCP 协议，与模型厂商无关。客户端必须能在安装 Office/WPS 的 Windows 用户桌面会话中启动本地进程。云端聊天网页本身不等于具备本机 MCP 客户端能力。

连接后直接调用工具，结果实时出现在原有 Office/WPS 窗口；服务无需单独控制页、接管按钮或录屏。客户端自身的工具审批及进度界面由客户端管理，服务不会代为关闭。即使客户端不显示 MCP 进度通知，原生文档窗口也会更新。

## 生成配置

安装后运行：

```powershell
.\.venv\Scripts\python.exe -m microsoft_wps_live_mcp --write-configs .\config\clients
```

也可只打印某个配置：

```powershell
.\.venv\Scripts\python.exe -m microsoft_wps_live_mcp --config codex
.\.venv\Scripts\python.exe -m microsoft_wps_live_mcp --config claude-code
.\.venv\Scripts\python.exe -m microsoft_wps_live_mcp --config deepseek-harness
```

配置中的 Python 路径来自当前安装环境。源码和公共示例不包含作者电脑路径。生成的配置是本机文件，不应提交 GitHub。

## 客户端对应关系

下表中的格式已对照官方文档或客户端配置核对。协议测试使用官方 MCP SDK，覆盖当前和 legacy 协商模式；没有逐家安装客户端进行界面验收，不能把配置格式核对写成全部客户端实测通过。

| 客户端 | profile / 文件 | 添加方式与依据 |
| --- | --- | --- |
| Codex | `codex` / `codex.toml` | 合并到客户端配置的 `mcp_servers`；使用当前安装 Python 的绝对路径 |
| Claude Desktop | `claude` / `claude.json` | 合并 `mcpServers` 配置，使用桌面版本地 MCP 设置 |
| Claude Code | `claude-code` / `claude-code.json` | 项目 `.mcp.json` 或客户端 MCP 配置，参见 [官方接入文档](https://code.claude.com/docs/en/mcp) |
| Google Gemini CLI | `gemini` / `gemini.json` | 合并到 `settings.json` 的 `mcpServers`，参见 [官方文档](https://geminicli.com/docs/tools/mcp-server/) |
| Google Antigravity | `antigravity` / `antigravity.json` | 从客户端打开 `mcp_config.json`，合并配置；参见 [官方文档](https://antigravity.google/docs/ide-mcp) |
| Kimi Code | `kimi` / `kimi.json` | 导入或合并其 MCP 配置；参见 [官方文档](https://www.kimi.com/code/docs/en/kimi-code-cli/customization/mcp) |
| ZCode | `zcode` / `zcode.json` | 设置 → MCP Servers → 新建 → 完整 JSON 配置；参见 [官方文档](https://zcode.z.ai/en/docs/mcp-services) |
| WorkBuddy | `workbuddy` / `workbuddy.json` | 添加自定义本地 stdio 服务，配置包含 120 秒超时；参见 [官方连接器文档](https://open.workbuddy.cn/docs/connector) |
| DeepSeek Harness | `deepseek-harness` / `deepseek-harness.yml` | 将生成的 Cordis 插件条目合并到目标配置作用域，需有 `@deepseek-ai/dsh-mcp-client`；参见 [官方包文档](https://github.com/deepseek-ai/deepseek-harness/blob/master/packages/mcp/mcp-client/README.md) |
| Trae | `trae` / `trae.json` | 添加自定义 MCP 时使用 JSON 配置；参见 [官方 MCP 入口](https://docs.trae.ai/ide/model-context-protocol)，具体导入界面需按安装版本确认 |
| Qoder | `qoder` / `qoder.json` | 添加自定义连接器或合并 `mcpServers`；参见 [官方配置参考](https://docs.qoder.com/cli/mcp-reference) |
| 其他支持本地 stdio MCP 的智能体 | `json` / `json.json` | 采用标准 `command`、`args`、`env` 条目；按客户端要求调整外层配置结构 |

DeepSeek Harness 的原生配置格式与 `mcpServers` JSON 不同，因此提供独立 YAML 片段。不能把通用 JSON 直接当作 Harness 的原生配置。

配置生成器只生成文件，不修改智能体审批策略，不覆盖用户现有客户端设置，不自动发布连接器。

## 检查连接

启动新客户端会话后检查 MCP 服务是否连接、是否列出 `office_status`、`office_api_batch` 等工具。先调用 `office_status`，再调用 `office_open_document` 或 `office_create_document`。

公开安装不需要现有 ONLYOFFICE 网页服务、WSL 或 Docker。请使用 Windows Python，不能用 WSL/Linux Python 进程直接调用 Windows COM。远程工作区运行在另一台主机时，也不会自动获得用户当前电脑上的桌面 Office。

同一文档应交给一个 MCP 服务实例管理。多个客户端各自启动实例、同时修改同一路径时可能产生版本冲突；冲突会暂停覆盖，不能通过重复调用强行解决。

自动保存观察器仅在 MCP 进程运行期间工作，且只管理明确连接的文档。结束修改前调用 `office_create_copy`，并向用户返回原文件、独立副本和恢复副本路径。
