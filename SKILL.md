---
name: microsoft-wps-live
description: Real-time creation, inspection, and live editing of Microsoft Office (Word, Excel, PowerPoint) and WPS Office documents via official stdio MCP on Windows and macOS. Supports visible step-by-step progress, auto-save with integrity verification, and recovery copies.
---

# microsoft-wps-live mcp 智能体入口

优先使用已配置的 `microsoft-wps-live` MCP。源码目录不会自动注册客户端；未连接时按 [README.md](README.md) 安装生成配置，不能假装已经取得工具访问能力。

连接后由智能体直接调用 MCP，用户在原有 Office/WPS 窗口观看和继续编辑，无需打开额外控制页或点击“让智能体操作”。原生编辑本身不依赖电脑接管、屏幕录制或浏览器预览。客户端自身的工具授权由客户端管理，服务不修改其审批设置。

## 执行流程

1. `office_status` 确认本机接口与文档。原生后端要求 Windows 桌面及 Office/WPS，其他系统不能声称已实时连接。
2. 按用户路径调用 `office_open_document` 或 `office_create_document`（先打开 PPT / Word / Excel 对应文档），并紧接着调用 `office_show_document` 将原生窗口**置顶到最前视图**；只有完成前置打开与置顶后，才开始执行任何绘图标注或文档写入。选择 `microsoft` / `wps`，未指定则 `auto`；连接已打开原文档或载入未打开文件。
   所有任务默认保持 `visible=true`，让用户在最上层视图实时看到整个编辑/画图过程。只有用户明确要求后台操作时才使用 `visible=false`。
3. 先读内容与结构再改。`document` 可用 ID 或绝对路径；段落、工作表、幻灯片、形状为 1 起始，Word 字符偏移为原生 0 起始。
4. 常用操作用快捷工具：
   - **Word 排版与三线表**：插入学术/商务表格优先传 `style="three_line"`（自动绘制 1.5pt 顶底粗线 + 0.75pt 栏目细线、去除竖线）；中英文混排通过 `font_name`（如 `"微软雅黑"`/`"宋体"`）配合 `ascii_font_name="Times New Roman"` 让英文、数字与公式符号自动应用新罗马字体。
   - **PPT 复杂矢量绘图**：科研架构图、流程图与 1:1 复刻优先使用 `office_ppt_draw_scene`（支持 `coord_space` 自动坐标映射、`prism3d`/`grid3d`/`shear_box`/`curve_arrow`/`junction`/`aligned_chain` 等）并配合 `office_ppt_export_image` 视觉校验（详细设计规范见 [PPT 架构图技能](skills/ppt-architecture-diagram/SKILL.md)）。
   - 其他高级原生能力用 `office_api_describe`、`office_api_constants`、`office_api_batch`。科研、办公、学习只是场景，不按示例限制功能。
   - **流畅分步节奏**：优先在一个 `office_live_steps` 或分节调用内按模块（如标题 → 背景 → 三线表 → 结论与批注）推进，减少碎片化往返等待，让用户在窗口前顺滑看到每一步过程。使用 `office_show_document` 定位编辑中的段落、区域、幻灯片。说明阶段进展，允许用户在过程中调整要求，不逐步索要确认。
5. 修改后自动保存，检查 `disk_saved`、`autosave_state`；原生缓存不能当作落盘。恢复副本失败应单独报告。
6. 结束编辑前调用 `office_create_copy`，报告原文件、独立副本、恢复副本位置，不覆盖已有副本。
   默认保留已完成文档的原生窗口供用户查看和继续编辑；未经要求不关闭窗口。

## 能力与结果

类型库声明、接口可调用、实际成功不同。某成员在类型库出现不代表当前版本实现它，不能把通用接口视为每项功能均已验收。参见 [对象模型说明](docs/OBJECT_MODEL.md) 与 [功能核对](docs/FEATURE_AUDIT.md)。

通过原生局部操作保留未修改内容、样式和动态域。不要直接重写已打开的 OOXML 文件，不要用整段重建代替局部格式调整。

批次不是事务。`partial_error` 保留已执行步骤并暂停自动覆盖，先检查再显式保存，不直接重放。`pending` 用 `office_operation_status` 查询。冲突、只读、保存错误时保留实时内容，不另起连接强行覆盖。

保存包含用户未保存的编辑。只操作用户指定的文档、组件和全局设置，不擅自退出应用或处理其他文件。未命名文档先指定保存路径；观察器仅在 MCP 运行期间自动保存已连接文档。

检查返回的 `live_view`。若显示失败或 `focus_error`，先解决视图问题，不能声称用户已经看到了正确位置。不关闭屏幕更新来加速，不把后台生成成品当成边做边看。详细行为见 [LIVE_EDITING.md](docs/LIVE_EDITING.md)。

## 安装与分发

原生 MCP 不依赖作者路径、WSL、Docker。见 [客户端接入](docs/CLIENTS.md)。不上传配置、文档、备份、运行环境、日志；使用 `scripts/export_source.py` 导出源码包。

可选旧网页模式见 [WEB_EDITOR.md](docs/WEB_EDITOR.md)，接口和索引规范与原生 MCP 不同，不能把网页会话称为桌面实时连接。
