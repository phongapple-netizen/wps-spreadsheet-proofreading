# Source provenance

本项目是 `phongapple-netizen/wps-proofreading` 的独立衍生项目，保留 GPL-3.0 许可证。

0.2.0 以文字版提交 `6c4602627e8ea1ef4a8f6ec54e241371ed12e946` 为只读参考，来源：
https://github.com/phongapple-netizen/wps-proofreading/tree/6c4602627e8ea1ef4a8f6ec54e241371ed12e946

迁移或改造的部分：

- `ui/taskpane.html`、`ui/taskpane.css`：沿用文字版布局与样式，改为表格范围、单元格结果和表格改写入口。
- `js/text-proofreading-core.js`：文字版纯校对逻辑，调整命名空间，修正重叠原文匹配计数和分类验证，并用于表格分段校对与一致性复核。
- `js/rewrite-core.js`：沿用纯改写提示、解析和事实风险检查。
- `js/rules-center.js`、`js/rules-ui.js`、`rules/`：迁移规则中心与内置规则包，使用独立存储命名空间，将规则测试改为逐单元格执行。
- `js/wps-et-api.js`、`js/ribbon.js`：沿用宿主兼容入口与任务窗格缓存设计，适配 ET。

表格的 `js/taskpane.js`、`js/spreadsheet-integration.js`、`js/settings-store.js` 和本地调试脚本为独立适配实现。没有迁入文字文档 Range、修订模式或整套文字宿主集成；两个仓库保持独立的版本、配置与发布流程。

文字版仓库仅查看，未写入或修改。
