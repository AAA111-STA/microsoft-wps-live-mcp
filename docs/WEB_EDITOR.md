# 可选的 ONLYOFFICE 网页编辑模式

本文件保留可选网页模式的使用说明。以下验证是开发环境记录，不代表其他电脑已经安装运行环境。

让用户和智能体在同一个本地网页会话中编辑 Word、Excel、PPT，并保存回本机原文件。采用 **一个文件夹、源码交付**：技能、代码、页面和本机运行环境放在一起，无 EXE。

面向能读取技能说明、执行本机命令或调用本地接口的智能体。支持网页嵌入的客户端可接入侧栏；其他客户端使用浏览器。当前验证使用浏览器，未逐一测试各家智能体侧栏。技能说明通用，当前安装脚本面向 **Windows + WSL 2**。

## 验证范围

| 功能 | 结果 |
| --- | --- |
| 网页编辑器与智能体桥接 | Word、Excel、PPT 均实际加载和连接 |
| 智能体实时修改并保存 | DOCX 追加段落、XLSX 修改单元格、PPTX 修改文本框均通过；独立读取磁盘文件确认内容 |
| 页面“保存到原文件”按钮 | Word 实际点击、等待写入、重新读取通过 |
| 用户手动键盘输入 | 编辑组件已接入；当前测试工具的键盘焦点错误阻止了验收，尚未标为验证通过 |
| 认证、冲突、文件边界、保存顺序、源码导出 | 自动测试覆盖 |
| 静态预览 | Word 1 页、PPT 2 页、Excel 2 个工作表已验证 |

这是可检查和继续开发的初版。ONLYOFFICE 与 Microsoft Office 的格式功能有差异；复杂排版、图表、公式等应先在副本上检查。

## 目录

```text
microsoft-wps-live-mcp/
├── SKILL.md              # 智能体入口
├── README.md
├── scripts/              # 服务、CLI、安装和源码导出脚本
├── assets/               # 页面及编辑器插件
├── tests/                # 完整性和冲突测试
├── config/               # 私有配置、令牌和原文件绑定，不上传
├── data/                 # 会话、文档修订和备份，不上传
├── logs/                 # 日志，不上传
├── downloads/            # 安装包，不上传
├── wsl/                  # 独立 WSL 运行环境，不上传
└── dist/                 # 生成的源码压缩包，不提交 Git
```

## 安装与启动

需要 Windows、可工作的 WSL 2、Python 3.11+、首次下载所需的网络，以及足够的内存和磁盘。文档服务建议至少预留 4 GB 内存和 40 GB 磁盘。依赖从官方下载，运行镜像不放入仓库。

解压源码到本地磁盘，在项目目录运行：

```powershell
powershell -NoProfile -File .\scripts\setup-runtime.ps1 -Python python
```

脚本创建独立的 `OfficeLiveEditor` WSL 发行版，将虚拟磁盘放在本目录的 `wsl/` 中，安装 Docker 并启动 ONLYOFFICE 和保存服务。同名发行版若属于别处，安装会停止，不覆盖它。无需重置已有 WSL 或 Docker Desktop 的默认配置。

开发机曾完成独立环境安装；在全新电脑完整运行安装脚本尚未另行验收。WSL 本身尚未启用时，需要先完成 Windows 的 WSL 安装。

以后启动、停止：

```powershell
python .\scripts\office_live.py start
python .\scripts\office_live.py stop
```

停止保留会话、备份和运行磁盘。导入 WSL 后不要直接改名或移动总目录：WSL 注册位置和生成配置包含路径，移动需先迁移 WSL 并重新生成配置。

## 打开与手动编辑

```powershell
python .\scripts\office_live.py open 'C:\Documents\报告.docx'
```

命令返回 `session`、`url`、`source`。打开编辑链接，然后编辑。**使用页面最上方“保存到原文件”按钮并等待磁盘写入确认；内部保存图标、Ctrl+S 和自动缓存不能单独证明已写回原文件。**

链接包含本地会话令牌，不要公开分享。编辑器和桥接未连接时，不要提交智能体修改。支持 DOCX、XLSX、PPTX，最大 100 MB；拒绝含宏文档，不自动转换旧格式或解除密码、保护。

## 智能体实时修改

让智能体读取 [SKILL.md](../SKILL.md)。普通文件夹不会自动注册到每个客户端；新项目或聊天可显式引用该文件，或按客户端自己的技能安装方式注册目录。

将操作写入 UTF-8 JSON 文件，例如：

```json
{"op":"word.append","text":"需要追加的段落。"}
```

```powershell
python .\scripts\office_live.py apply SESSION_ID --command-file 'C:\Temp\操作.json'
python .\scripts\office_live.py status SESSION_ID
```

`apply` 默认等待编辑器执行，再请求保存并等待对应保存序号的磁盘确认。`--no-save` 只修改当前会话。超时属于结果未确认，不能自动重放追加操作。

| 操作 | JSON 示例 |
| --- | --- |
| 检查内容或结构 | `{"op":"inspect"}` |
| Word 追加段落 | `{"op":"word.append","text":"新段落"}` |
| Word 查找替换 | `{"op":"word.replace","search":"原文","replace":"新文","matchCase":true}` |
| Excel 写入 | `{"op":"sheet.set","sheet":"Sheet1","range":"C3","values":"新值"}` |
| Excel 读取 | `{"op":"sheet.read","sheet":"Sheet1","range":"A1:B3"}` |
| PPT 文本框替换 | `{"op":"slide.set_text","slide":0,"shape":0,"text":"新标题"}` |

PPT 页号、形状索引从 0 开始；先检查结构再选择。文本替换重建该文本框的段落，原有局部样式可能改变。其他编辑功能尚无智能体接口，不要声称已支持。

`save SESSION_ID` 请求单独保存，`sessions` 列出会话。原文件被外部修改时暂停回写，保留会话修订；先比较版本，不能绕过检查覆盖。`open --new-session` 仅用于明确确认后的新快照。

## 数据与运行边界

- 文档由本地 ONLYOFFICE 服务处理；本项目未调用外部文档编辑或 AI API。首次安装会下载系统包与容器镜像。
- 服务仅发布到 `127.0.0.1:18820`、`127.0.0.1:18821`；API、会话与回调保留访问令牌或 JWT 认证。
- 保存前备份在 `data/backups/`。哈希冲突检查不能替代其他写入程序的文件锁协调，避免同时用其他程序编辑同一个原文件。
- 命令队列在内存中；服务重启、断线后可能需要人工确认，不自动重放未确认的修改。
- ONLYOFFICE 服务在运行时下载，遵守其上游许可；仓库不附带该服务镜像或第三方二进制。

## 静态预览

此模式单独运行，需要本机 Microsoft Office、Python 包 `pywin32` 和 `PyMuPDF`：

```powershell
python .\scripts\preview.py --input 'C:\Documents\报告.docx' --output-dir 'C:\Temp\报告预览'
```

支持 DOC/DOCX、PPT/PPTX、XLSX，输出 PDF 和包含翻页、缩放、工作表切换的静态 HTML。它没有编辑能力。转换前对应 Office 应用需关闭，不强制结束用户进程。

## 测试和 GitHub 源码导出

```powershell
python -B -m unittest discover -s .\tests -v
python .\scripts\export_source.py
```

输出：`dist/microsoft-wps-live-mcp-source.zip`。导出脚本只选取说明文件、脚本、页面和测试，不遍历运行数据。通过 GitHub 网页上传时，**先解压源码包，再上传解压出来的源码**；直接上传 ZIP 只会得到一个压缩文件。

`.gitignore` 同样排除运行环境、凭据、文档和日志。手动拖拽整个本机目录不受 `.gitignore` 保护，不要这样上传。是否配置远程仓库由使用者自行管理。

规范：[Agent Skills](https://agentskills.io/specification)。集成：[ONLYOFFICE Docs API](https://api.onlyoffice.com/docs/docs-api/get-started/installation/self-hosted/)、[插件 API](https://api.onlyoffice.com/docs/plugins/get-started/quick-start/)。
