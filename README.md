# microsoft-wps-live mcp

![microsoft-wps-live-mcp](assets/cover.png)

通用的本机 Office MCP 服务，让智能体在 **Microsoft Office 和 WPS Office** 中创建、读取和编辑 Word、Excel、PPT。文档已打开时连接当前原文档，未打开时由本机软件载入。适用于工作、学习、科研和个人使用，功能不按行业或任务设白名单。

**双系统原生支持：Windows & macOS。**
- **Windows**: 通过原生 COM 自动化接口无缝连接 Microsoft Office 桌面版与开放 COM 的 WPS Office。
- **macOS**: 通过 AppleScript 及桌面应用桥接，支持 Microsoft Office (Word, Excel, PowerPoint) 与 macOS 版 WPS Office。

## 已实现

- 标准 stdio MCP，官方 Python SDK，35 个工具，支持当前与 legacy 协议协商。
- **所有文档任务默认可见、分步执行**：显示真实编辑窗口，跟随当前段落、单元格区域或幻灯片，内容逐渐成型。复杂任务用命名阶段执行，每阶段保存并报告进度。
- 智能体直接调用 MCP，用户在原有 Office/WPS 窗口观看；无需额外控制页或“让智能体操作”按钮。原生编辑不依赖电脑接管或录屏；客户端自身工具授权仍由客户端管理。
- 通用对象模型：查询对象结构和枚举，批量读取属性、写入属性、调用方法，返回可继续使用的原生对象引用。
- Word 文本、表格、批注、修订、脚注尾注、书签、域、交叉引用、公式、图片、快照差异等快捷操作。
- Excel 区域、公式、工作表和格式；PPT 幻灯片、文本框与格式；三类文件的组件检查与 PDF 导出。
- 通用接口访问更多原生组件，包括形状、连接线、图表、主题、母版、表格、图片和其他公开对象。示例包含可编辑流程图与 Excel 图表。
- **成功修改后自动保存**并核验磁盘文件。MCP 运行期间，明确连接文档的手动修改也由观察器检测保存；空闲轮询约 1 秒，软件忙碌时会延后。
- **每次保存更新独立恢复副本**，修改前保留版本备份；结束时调用 `office_create_copy` 生成额外完整副本，保留原文档路径。
- 冲突、只读、保存失败和部分执行失败明确报告，不将未确认编辑说成已保存。

接口开放范围与逐项验收不同，详见 [兼容性记录](docs/COMPATIBILITY.md)、[功能核对清单](docs/FEATURE_AUDIT.md)。

## 安装

需要 Windows 或 macOS、Python 3.11+ 和本机 Office/WPS。源码可放在任意有权限的目录，不依赖作者电脑路径。原生 MCP **无需 WSL、Docker 或 ONLYOFFICE**。

### Windows (PowerShell)

在项目根目录运行：

```powershell
powershell -NoProfile -File .\scripts\setup-mcp.ps1
```

也可手动安装：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install .
.\.venv\Scripts\python.exe -m microsoft_wps_live_mcp --doctor
.\.venv\Scripts\python.exe -m microsoft_wps_live_mcp --write-configs .\config\clients
```

### macOS / Linux (Bash)

在项目根目录运行：

```bash
chmod +x ./scripts/setup-mcp.sh
./scripts/setup-mcp.sh
```

也可手动安装：

```bash
python3 -m venv .venv
./.venv/bin/python -m pip install .
./.venv/bin/python -m microsoft_wps_live_mcp --doctor
./.venv/bin/python -m microsoft_wps_live_mcp --write-configs ./config/clients
```

普通 pip/uv 安装也可以。安装后以 `microsoft-wps-live-mcp` 或 `python -m microsoft_wps_live_mcp` 启动，默认 stdio 通信，不监听端口。

## 接入智能体

为 Codex、Claude Desktop / Code、Google Gemini CLI / Antigravity、Kimi Code、ZCode、WorkBuddy、DeepSeek Harness、Trae、Qoder 和通用 stdio 客户端提供配置生成器。

```powershell
.\.venv\Scripts\python.exe -m microsoft_wps_live_mcp --config codex
.\.venv\Scripts\python.exe -m microsoft_wps_live_mcp --config claude-code
.\.venv\Scripts\python.exe -m microsoft_wps_live_mcp --config deepseek-harness
```

合并输出到客户端 MCP 设置，再启动新会话。DeepSeek Harness 使用独立插件 YAML 格式。详见 [接入步骤与官方依据](docs/CLIENTS.md)。配置生成不等于已在每个客户端完成实测。

## 使用

可以向智能体提出任意合适的 Office 任务，例如家庭收支表、项目流程图、科研图表、学习笔记、业务报告或文档审校。这些只是示例，不是功能白名单。

先调用 `office_status`，再通过 `office_open_document` / `office_create_document` 连接文件。`engine` 可选 `microsoft`、`wps`、`auto`：已打开文件沿用当前软件，未打开文件优先 Microsoft，其次 WPS。

后续 `document` 可使用 ID 或绝对路径；重启 MCP 后用路径重新连接。段落、工作表、幻灯片和形状索引从 **1** 开始，Word 字符偏移为原生 **0** 起始偏移。

`office_open_document` 和 `office_create_document` 默认 `visible=true`。Word 较长文本会按段落/文本块写入，普通 Excel 多行数据会按行块写入；所有快捷修改都会更新原生窗口。只有用户明确要求后台工作时才使用 `visible=false`。

复杂任务使用 `office_live_steps`，把工作拆成“建立结构 → 填入内容/数据 → 制作组件 → 调整格式”等短阶段。可以用 `office_show_document` 定位当前内容。此行为适用于三类文档的所有任务，不限于流程图或科研场景。见 [边做边看的使用说明](docs/LIVE_EDITING.md)。

快捷工具未覆盖的功能，通过 `office_api_describe`、`office_api_constants` 和 `office_api_batch` 查询并组合原生操作。见 [对象模型说明](docs/OBJECT_MODEL.md) 与 [examples](examples/)。

## 自动保存与副本

只管理明确连接的文档，不接管仅在列表中发现的其他文件。观察器需要 MCP 进程持续运行；电脑关机、软件挂起、文件只读、未命名或服务退出时无法保证自动保存。未命名文档先指定保存路径。

工具返回 `disk_saved`、SHA256、自动保存状态和恢复副本路径。软件内部 `saved` 状态与磁盘核验分别报告。外部磁盘修改或部分执行失败会暂停自动覆盖；检查后显式保存可恢复观察。

默认数据目录为 `%LOCALAPPDATA%\microsoft-wps-live-mcp`，可用 `MICROSOFT_WPS_LIVE_MCP_DATA` 或 `--data-dir` 调整。

从旧名称升级时会继续使用已经存在的旧数据目录，并兼容旧环境变量，保留原有备份和恢复副本。

| 位置 | 用途 |
| --- | --- |
| `backups/` | 修改前磁盘版本；每个连接保留首版与最近 49 版，不含当时未保存的用户修改 |
| `recovery/` | 每次保存后更新的独立恢复副本 |
| `office_create_copy` 返回路径 | 结束时的完整副本，默认在原文件旁生成唯一文件名，不覆盖已有文件 |

智能体结束编辑前应调用 `office_create_copy`，向用户提供原文件与副本位置。恢复副本用于原文件误删后找回，历史备份用于查看修改前版本；它们不能替代异地备份。

## 验证与开发

无需 Office 的自动测试：

```powershell
python -B -m unittest discover -s tests -v
```

真实桌面软件测试使用新文件，输出保存在被 Git 忽略的 `data/`：

```powershell
python .\scripts\smoke_native.py --engine all
python .\scripts\smoke_mcp.py --engine all
python .\scripts\smoke_mcp.py --engine all --mode legacy
python .\scripts\smoke_features.py --engine all
python .\scripts\smoke_live.py --engine all
```

仅在安装对应软件的 Windows 交互式桌面运行真实测试。GitHub Actions 只运行无需 Office 的测试。实际范围和缺口见 [COMPATIBILITY.md](docs/COMPATIBILITY.md)。

## GitHub 源码交付

```powershell
python .\scripts\export_source.py
```

生成 `dist/microsoft-wps-live-mcp-source.zip`，只收录源码、文档、示例和测试，排除配置、凭据、文档、备份、日志、Python 环境和可选运行镜像。用 GitHub 网页上传源码时先解压，不要拖入整个本机工作目录。

## 可选网页模式

原 ONLYOFFICE 网页编辑与静态预览代码仍保留，使用独立入口，不负责控制桌面 WPS/Microsoft Office。见 [网页模式](docs/WEB_EDITOR.md)。其 PPT 索引从 0 开始，与原生 MCP 不同。

仓库不附带 Microsoft Office、WPS、ONLYOFFICE 安装包或许可证。第三方软件和依赖遵守各自上游许可。
