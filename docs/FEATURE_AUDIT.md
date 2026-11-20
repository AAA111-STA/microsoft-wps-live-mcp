# 功能核对记录

本次对照当前会话中可用的 word-mcp-live 工具清单及本机安装实现，共 **120 个入口**，其中 **44 个 word_live_ 入口**。该数量包含离线/实时重复入口和多个脚注变体，不等于 120 种互不重复的能力。没有复制第三方实现代码。

快捷入口、通用接口可调用、逐项实测通过是三种不同状态。下面记录的是接口对应关系；“通用对象接口路线”表示存在原生对象模型操作路径，不表示 WPS/Microsoft 的所有版本已经验证通过。真实验收范围见 [兼容性记录](COMPATIBILITY.md)。

## 需要继续区分的能力

- 新项目已增加原生批注与修订、快照差异、脚注尾注、书签、域、交叉引用、公式/图片插入、PDF 导出等快捷入口。
- 复杂表格、样式、多级编号、页眉页脚、水印、目录更新和文档合并等仍主要通过通用接口组合，尚未提供旧工具的全部专用参数封装。
- 自动版式诊断、撤销历史摘要/工具级撤销分组、窗口截图、脚注 XML 完整性校验、实时 XML 导出等独立能力仍未实现。不能以“有通用接口”为由将其标为完成。
- 检查旧工具实现发现，其数字签名入口写入可见签名文字和哈希元数据；本项目没有将其表述为证书数字签名。
- 旧工具交叉引用中部分数字常量与本机实际类型库不符；新工具采用原生枚举，例如 wdContentText=-1、wdPageNumber=7，题注类型使用实际本地化标签。引用中文标题不依赖英文 Heading 前缀。

## Word 逐项对应

| 旧工具入口 | 新项目状态 | 对应与限制 |
| --- | --- | --- |
| `accept_tracked_changes` | 有对应快捷入口 | `office_word_revisions`；接口语义和参数不完全相同 |
| `add_bookmark` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `add_comment` | 有对应快捷入口 | `office_word_comments`；接口语义和参数不完全相同 |
| `add_digital_signature` | 待补 / 独立能力 | 未实现证书数字签名；旧实现为文字和哈希元数据 |
| `add_endnote_to_document` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `add_footnote_after_text` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `add_footnote_before_text` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `add_footnote_enhanced` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `add_footnote_robust` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `add_footnote_to_document` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `add_header_footer` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `add_heading` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `add_page_break` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `add_page_numbers` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `add_paragraph` | 有对应快捷入口 | `office_word_insert_text`；接口语义和参数不完全相同 |
| `add_picture` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `add_restricted_editing` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `add_section_break` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `add_table` | 有对应快捷入口 | `office_word_add_table`；接口语义和参数不完全相同 |
| `add_table_of_contents` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `add_watermark` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `apply_table_alternating_rows` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `auto_fit_table_columns` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `convert_to_pdf` | 有对应快捷入口 | `office_export_pdf`；接口语义和参数不完全相同 |
| `copy_document` | 有对应快捷入口 | `office_create_copy`；接口语义和参数不完全相同 |
| `create_custom_style` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `create_document` | 有对应快捷入口 | `office_create_document`；接口语义和参数不完全相同 |
| `customize_footnote_style` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `delete_footnote_from_document` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `delete_footnote_robust` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `delete_paragraph` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `find_text_in_document` | 有对应快捷入口 | `office_word_find`；接口语义和参数不完全相同 |
| `format_table` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `format_table_cell_text` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `format_text` | 有对应快捷入口 | `office_word_format_text`；接口语义和参数不完全相同 |
| `get_all_comments` | 有对应快捷入口 | `office_word_comments`；接口语义和参数不完全相同 |
| `get_comments_by_author` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `get_comments_for_paragraph` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `get_document_info` | 有对应快捷入口 | `office_inspect_document`；接口语义和参数不完全相同 |
| `get_document_outline` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `get_document_text` | 有对应快捷入口 | `office_word_read`；接口语义和参数不完全相同 |
| `get_document_xml` | 待补 / 独立能力 | 缺少实时原始 XML 导出工具；磁盘 XML 不代表未保存缓冲区 |
| `get_highlighted_text` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `get_paragraph_text_from_document` | 有对应快捷入口 | `office_word_read`；接口语义和参数不完全相同 |
| `highlight_table_header` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `insert_header_near_text` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `insert_line_or_paragraph_near_text` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `insert_numbered_list_near_text` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `list_available_documents` | 待补 / 独立能力 | 缺少独立目录扫描快捷工具；客户端可提供文件系统入口 |
| `list_tracked_changes` | 有对应快捷入口 | `office_word_revisions`；接口语义和参数不完全相同 |
| `manage_hyperlinks` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `merge_documents` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `merge_table_cells` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `merge_table_cells_horizontal` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `merge_table_cells_vertical` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `protect_document` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `reject_tracked_changes` | 有对应快捷入口 | `office_word_revisions`；接口语义和参数不完全相同 |
| `replace_block_between_manual_anchors` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `replace_paragraph_block_below_header` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `search_and_replace` | 有对应快捷入口 | `office_word_replace_text`；接口语义和参数不完全相同 |
| `set_page_layout` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `set_paragraph_spacing` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `set_table_alignment_all` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `set_table_cell_alignment` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `set_table_cell_padding` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `set_table_cell_shading` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `set_table_column_width` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `set_table_column_widths` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `set_table_width` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `track_delete` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `track_insert` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `track_replace` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `unprotect_document` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `validate_document_footnotes` | 待补 / 独立能力 | 缺少专用 OOXML 脚注完整性校验器 |
| `verify_document` | 待补 / 独立能力 | 未实现旧工具元数据签名验证，也不宣称证书签名验证 |
| `word_live_accept_revisions` | 有对应快捷入口 | `office_word_revisions`；接口语义和参数不完全相同 |
| `word_live_add_bookmark` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `word_live_add_comment` | 有对应快捷入口 | `office_word_comments`；接口语义和参数不完全相同 |
| `word_live_add_header_footer` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `word_live_add_page_numbers` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `word_live_add_section_break` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `word_live_add_table` | 有对应快捷入口 | `office_word_add_table`；接口语义和参数不完全相同 |
| `word_live_add_watermark` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `word_live_apply_list` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `word_live_delete_comment` | 有对应快捷入口 | `office_word_comments`；接口语义和参数不完全相同 |
| `word_live_delete_text` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `word_live_diagnose_layout` | 待补 / 独立能力 | 缺少自动版式问题诊断算法 |
| `word_live_find_text` | 有对应快捷入口 | `office_word_find`；接口语义和参数不完全相同 |
| `word_live_format_table` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `word_live_format_text` | 有对应快捷入口 | `office_word_format_text`；接口语义和参数不完全相同 |
| `word_live_get_comments` | 有对应快捷入口 | `office_word_comments`；接口语义和参数不完全相同 |
| `word_live_get_diff` | 有对应快捷入口 | `office_word_snapshot`；接口语义和参数不完全相同 |
| `word_live_get_info` | 有对应快捷入口 | `office_inspect_document`；接口语义和参数不完全相同 |
| `word_live_get_page_text` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `word_live_get_paragraph_format` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `word_live_get_text` | 有对应快捷入口 | `office_word_read`；接口语义和参数不完全相同 |
| `word_live_get_undo_history` | 待补 / 独立能力 | 缺少专用撤销历史摘要与工具操作分组记录 |
| `word_live_insert_cross_reference` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `word_live_insert_equation` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `word_live_insert_image` | 有对应快捷入口 | `office_word_insert_component`；接口语义和参数不完全相同 |
| `word_live_insert_paragraphs` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `word_live_insert_text` | 有对应快捷入口 | `office_word_insert_text`；接口语义和参数不完全相同 |
| `word_live_list_cross_reference_items` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `word_live_list_open` | 有对应快捷入口 | `office_list_open_documents`；接口语义和参数不完全相同 |
| `word_live_list_revisions` | 有对应快捷入口 | `office_word_revisions`；接口语义和参数不完全相同 |
| `word_live_modify_table` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `word_live_reject_revisions` | 有对应快捷入口 | `office_word_revisions`；接口语义和参数不完全相同 |
| `word_live_replace_text` | 有对应快捷入口 | `office_word_replace_text`；接口语义和参数不完全相同 |
| `word_live_reply_to_comment` | 有对应快捷入口 | `office_word_comments`；接口语义和参数不完全相同 |
| `word_live_resolve_comment` | 有对应快捷入口 | `office_word_comments`；接口语义和参数不完全相同 |
| `word_live_save` | 有对应快捷入口 | `office_save_document`；接口语义和参数不完全相同 |
| `word_live_set_core_properties` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `word_live_set_page_layout` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `word_live_set_paragraph_spacing` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `word_live_setup_heading_numbering` | 通用对象接口路线 | 可通过 office_api_describe + office_api_batch 组合原生调用；未逐项验收 |
| `word_live_snapshot_status` | 有对应快捷入口 | `office_word_snapshot`；接口语义和参数不完全相同 |
| `word_live_take_snapshot` | 有对应快捷入口 | `office_word_snapshot`；接口语义和参数不完全相同 |
| `word_live_toggle_track_changes` | 有对应快捷入口 | `office_word_revisions`；接口语义和参数不完全相同 |
| `word_live_undo` | 有对应快捷入口 | `office_word_undo`；接口语义和参数不完全相同 |
| `word_screen_capture` | 待补 / 独立能力 | 缺少专用窗口截图工具；PDF 导出不等于截图 |

## Excel 与 PPT

当前会话没有暴露可供逐项比较的 Excel/PPT 专用 MCP 工具。以下按原生对象模型核对，不声称已对照其他未连接的产品。

| 软件 | 已有快捷入口或实测内容 | 通用接口可进一步访问，但尚未逐项验收 |
| --- | --- | --- |
| Excel / WPS 表格 | 原生区域值/公式读取写入、工作表新增改名、单元格格式、图表创建；已连接文档自动保存和副本 | 数据验证、条件格式、结构化表、透视表、名称、筛选排序、图片、批注、页面设置及其他公开组件 |
| PPT / WPS 演示 | 新建文件/幻灯片、文字和字体、文本框、原生形状、连接点箭头、渲染导出；自动保存和副本 | 母版、布局、主题、表格、图表、图片、分组、动画、媒体、备注及其他公开组件 |

这些是核对分组，不是任务白名单。科研、工作、学习和个人使用共用同一套接口。软件专有界面、未开放接口及特殊 COM 回调能力可能仍需另外实现适配。

