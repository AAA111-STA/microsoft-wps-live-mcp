# 通用对象模型接口

本项目是通用 Office MCP。科研、工作、学习、家庭账本和个人文档都是使用场景，不存在按行业或任务设定的白名单。快捷工具用于减少常用操作的调用次数；`office_api_batch` 用于访问安装软件开放的公共对象、属性和方法。

## 工作方式

1. `office_status` 检测环境、查看已打开的文档。
2. `office_open_document` 连接原文档；`office_create_document` 新建文件。`engine` 可选 `microsoft`、`wps`、`auto`。
3. `office_api_describe` 查询对象支持的成员和参数名；`office_api_constants` 查询本机软件类型库中的枚举值。
4. `office_api_batch` 组合操作。结果中的 `object_id` 可以用于后续调用。
5. 检查 `state`、`disk_saved` 和实际内容；按需继续操作或使用 `office_save_document` 保存。

所有任务默认可见，复杂任务优先用 `office_live_steps` 组织阶段。阶段共享别名，每阶段自动保存并显示进度；普通批次的修改也更新原生窗口。用 `office_show_document` 定位当前内容，见 [实时编辑说明](LIVE_EDITING.md)。

每个工具中的 `document` 可以是已返回的 `document_id`，也可以是文件绝对路径。需要指定软件时先调用 `office_open_document`，后续使用它的 ID。`auto` 会沿用正在编辑该路径的软件；未打开的文件优先 Microsoft，其次 WPS。

## 操作语法

`steps` 是一个数组，每项可以包含：

| 字段 | 含义 |
| --- | --- |
| `action` | `get` 读取属性，`set` 写入属性，`call` 调用方法或带参数的索引属性 |
| `target` | `document`、`application`、当前批次的别名、之前返回的 `object_id`，可接点分属性路径 |
| `member` | 公开成员名称，按本机对象模型的命名填写 |
| `args` | `call` 的位置参数数组 |
| `kwargs` | `call` 的命名参数对象 |
| `value` | `set` 的新值 |
| `as` | 将结果保存为本批次内可用的别名 |
| `read_only` | 对真正只读的 `Item`、`Range` 等调用可填写 `true`；写操作不可标为只读 |

点分路径只表示属性导航，不解析 Python/JavaScript 表达式，不执行 `eval`。索引访问使用单独的 `call` 步骤。

参数中的 `{"$ref":"box"}` 引用本批次别名，`{"$ref":"obj_..."}` 引用之前返回的原生对象。`{"$missing":true}` 表示省略该位置的原生可选参数。二维数组会转换为 COM 数组，适用于 Range 数据与公式等接口。

例：向演示文稿插入原生形状。

```json
{
  "document": "上一步返回的document_id",
  "steps": [
    {"action":"call","target":"document.Slides","member":"Add","args":[1,12],"as":"slide"},
    {"action":"call","target":"slide.Shapes","member":"AddShape","args":[1,40,80,240,100],"as":"box"},
    {"action":"set","target":"box.TextFrame.TextRange","member":"Text","value":"任务与结果"},
    {"action":"set","target":"box.TextFrame.TextRange.Font","member":"Size","value":24}
  ]
}
```

数值枚举 `12`、`1` 的含义分别为 `ppLayoutBlank`、`msoShapeRectangle`。其他成员与枚举先查本机类型库，或查对应软件的官方对象模型文档。类型库只能显示接口声明；WPS 某些版本可能声明成员但不实现，实际调用会返回错误。

`get` 返回 COM 对象时得到 `object_id` 和类型名称，而不是把对象转换成文本。对象引用仅在本 MCP 进程和对应文档连接中有效；关闭或自动重新打开文档后要重新获取引用。最多保留最近 1000 个引用。

## 能力范围

同一接口可以访问文字、段落、样式、节、页眉页脚、表格、公式、域、批注、工作表、区域、名称、图表、透视表、幻灯片、母版、主题、图片、形状、连接线、动画、媒体、导出设置，以及其他已公开的原生对象。具体成员由安装软件决定，项目不会按任务名称限制功能。

这不是对任何软件版本、任何文件、每个界面按钮均可自动化的承诺。交互式对话框、加载项专有界面、未开放的组件、某些需要回调或特殊 COM 接口的功能，可能无法通过通用 get/set/call 调用。WPS 和 Microsoft 的对象模型也有差异。遇到错误应查询实际成员、调整操作或明确说明限制，不能把失败说成成功。

Windows 上本机交互式用户会话是当前原生后端的运行环境。Linux/macOS 不能用本项目的 Windows COM 后端控制当地桌面软件。

## 保存与并发

所有本机 COM 操作在一个持久 STA 线程上串行执行。读到的是原生应用中的当前文档，包括用户尚未保存的编辑。

成功的快捷修改和通用修改批次都会自动保存，公共 MCP 工具不提供跳过保存的开关。全是只读步骤的批次不主动保存；空闲观察器会约每秒检查明确连接文档的手动修改并保存。最终以 `disk_saved=true` 和磁盘核验为准，`saved` 只表示软件报告的原生状态。

每次保存同时更新独立恢复副本，修改前保留版本备份。结束编辑前调用 `office_create_copy`，生成额外完整副本并向用户返回位置。观察器只在 MCP 进程运行期间工作，不会接管仅在列表中发现的其他文档。

修改前备份当前磁盘版本；备份不包含当时仍未保存的用户编辑。备份位置由 `MICROSOFT_WPS_LIVE_MCP_DATA` 或 `--data-dir` 配置，默认放在当前用户的应用数据目录。检测到磁盘在连接后发生变化时暂停覆盖；可另存新路径或先解决冲突。原生保存前的哈希检查不能消除其他程序同时写盘的竞态，不应同时用两个程序编辑同一路径。

批次不是事务。第 N 步失败时前面的步骤可能已修改文档；返回 `partial_error`、失败位置和已完成步骤，不自动回滚，也不自动重放。原生调用超过等待期限时返回 `pending` 和 `operation_id`，使用 `office_operation_status` 查询完成状态。

优先使用专门的打开、创建、保存和关闭工具管理文件生命周期。通过 `application` 可以调用范围很广的桌面 API；仅在用户要求时操作其他文档、全局设置、宏或退出软件。备份和默认保存只管理当前选定文档；要编辑另一文件，应为它单独建立连接。

## 可复用示例

- [PPT 流程图](../examples/ppt_flowchart.json)：原生形状、文字、连接点与箭头，可继续在软件中编辑。
- [Excel 图表](../examples/excel_chart.json)：写数据、创建柱状图、设置标题。
- [Word 排版](../examples/word_layout.json)：页边距、插入文本、字体和段落间距。

这些文件的内容是 `steps` 数组，可调整为科研、业务、教学或个人使用的任何合适内容。

官方参考：[MCP Python SDK](https://github.com/modelcontextprotocol/python-sdk)、[Microsoft Office 对象模型](https://learn.microsoft.com/en-us/office/vba/api/overview/)、[WPS 开放平台](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wps/Application/obj)。
