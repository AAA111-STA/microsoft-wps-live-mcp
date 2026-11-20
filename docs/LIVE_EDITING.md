# 所有任务默认边做边看

从 0.3.0 开始，所有选定文档的任务默认在真实 Word、Excel、PowerPoint 或对应 WPS 窗口中执行。用户看到同一文档中的内容、数据、图表和布局逐渐形成，并可继续直接编辑。

这是三类文档的统一工作方式，适用于写文档、改稿、制作表格、处理数据、做 PPT、绘制图形等所有合适任务。科研流程图只是其中一个例子。

智能体直接调用 MCP，修改出现在原有 Office/WPS 编辑窗口。服务不创建额外控制页，不要求用户点击“让智能体操作”，也不依赖屏幕接管或录制。客户端自己显示的工具授权和进度界面由客户端管理；服务不会修改客户端审批策略。

## 默认行为

- 打开或新建文档默认 `visible=true`，必要时创建原生编辑窗口、恢复最小化窗口。
- 首次显示或明确定位时，原生窗口小到不足以查看内容（宽不足 800 或高不足 500 像素）会展开；较大的现有窗口布局保留。PPT 初次显示时尝试适应画布。
- 在选定文档中进行原生修改，保留可编辑组件，不以后台生成的静态图片替代制作过程。
- 修改时跟随段落、区域或幻灯片。Word/WPS 文字保留各自可用的定位实现；Excel 定位工作表与区域；PPT 定位幻灯片及可选形状。
- 长文字按段落/文本块写入，普通 Excel 多行数据按行块写入，单次调用最多分成约 24 组可见更新。公式块和合并单元格保持原有操作完整性。
- 快捷修改在结束时保存；可分块的文字、表格和数据写入也在块之间保存。恢复副本继续更新。
- 命名阶段执行完成后保存，报告阶段进度，并留出短暂显示时间。
- 仅列出文件的工具不接管其他文档窗口；观察器自动保存不反复抢焦点。
- 默认保留最终文档窗口，结束时生成独立副本。只有用户明确要求后台工作或关闭时才相应改变行为。

## 工具

### `office_show_document`

显示并定位选定文档。不修改内容，可将先前隐藏打开的原生文档显示出来。

```json
{"document":"返回的document_id","focus":{"slide":1,"shape":2}}
```

Word 的 `focus` 使用 `start`、`end` 原生字符偏移；Excel 使用 `sheet`、`address`；PPT 使用 1 起始的 `slide` 和可选 `shape`。

### `office_live_steps`

一个阶段包含 `label`、`steps` 和可选 `focus`。`steps` 使用通用对象模型的 get/set/call 语法。同一任务中前面阶段创建的别名，可以在后面的阶段继续引用。

```json
{
  "document":"返回的document_id",
  "pause_ms":450,
  "stages":[
    {
      "label":"建立幻灯片",
      "steps":[{"action":"call","target":"document.Slides","member":"Add","args":[1,12],"as":"slide"}],
      "focus":{"slide":1}
    },
    {
      "label":"加入第一个模块",
      "steps":[
        {"action":"call","target":"slide.Shapes","member":"AddShape","args":[5,60,190,220,100],"as":"box"},
        {"action":"set","target":"box.TextFrame.TextRange","member":"Text","value":"数据与研究问题"}
      ],
      "focus":{"slide":1}
    }
  ]
}
```

每个任务最多 100 个阶段，每阶段最多 500 个原生操作，建议保持阶段简短，避免长时间单次调用。阶段间显示时间 `pause_ms` 为 0–2000 毫秒，默认 450。原生图形或文字属性的修改也会让软件处理绘制消息，默认显示约 180 毫秒。

MCP 客户端支持进度通知时，会同时收到阶段编号、总数和描述；不支持进度面板的客户端仍可看到原生窗口的变化。阶段结果包含保存状态、视图状态、对象引用及可继续操作的信息。

用户中断仍在进行的调用时，在当前原生阶段结束后停止后续阶段；不强制打断不可分割的原生调用。超过等待期限返回 `pending` 时，使用 `office_operation_status` 查询，包括最新阶段，不能重放整项任务。

## 按任务组织步骤

| 任务 | 典型的可见阶段 |
| --- | --- |
| 撰写文档 | 建标题、写各节、插入表格/图片、调整样式、检查保存 |
| 修改文档 | 定位目标段落、修改文字、设置格式/修订、查看结果 |
| 表格和数据 | 建表头、填入数据块、设置公式、制作图表、调整格式 |
| 制作演示文稿 | 新建页、加入内容与组件、调整排版、检查各页 |
| 研究框架图 | 建画布/标题、逐组加入模块、逐条连接、统一样式与间距 |

这些是组织方式示例，不限制任务范围。应按用户实际材料和要求组织内容，不凭空添加研究结论。

## 结果判定

`live_view.state=ready` 表示已准备原生窗口；`window_open` 是窗口对象状态。`visibility_confirmed=true` 表示还核对了 Windows 原生可见窗口，`hwnd` 是该窗口的标识。`updated_at` 是最近一次检查时间，不能保证用户随后没有切换窗口。

`focus_error` 表示定位未确认，不能据此声称正在展示正确位置；`unavailable` 表示显示不可用。修改若已保存，其 `disk_saved=true` 与显示错误分别报告，不将显示失败误说成文件未保存。批次部分失败则保留已执行结果，停止后续步骤供检查。

Windows 可能限制前台抢焦点，软件本身也可能暂时忙碌。单个不可分割操作（例如某些导出或组件计算）可能一次完成，不能保证像逐字打字一样展示其全部内部过程。复杂任务应由智能体拆成有意义的小阶段，而非仅加长延时。

## 示例与复现

- [Word 分阶段撰写](../examples/live_word_document.json)
- [Excel 数据、公式与图表](../examples/live_excel_workbook.json)
- [PPT 形状和连接线逐步成型](../examples/live_ppt_flowchart.json)

```powershell
python .\scripts\smoke_live.py --engine all
python .\scripts\smoke_live.py --engine microsoft --kind ppt --keep-open
```

测试只创建 `data/live-smoke/` 中的新文件，记录每个阶段的磁盘哈希和结构变化，并通过真实 stdio MCP 调用执行。未指定 `--keep-open` 时，测试结束会关闭自己的测试文档；这不改变产品默认保留用户文档窗口的行为。

官方定位接口参考：[Word ScrollIntoView](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.window.scrollintoview)、[Excel Goto](https://learn.microsoft.com/en-us/office/vba/api/excel.application.goto)、[PPT GotoSlide](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.view.gotoslide)。
