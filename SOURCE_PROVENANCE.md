# Source provenance

本项目作为 `phongapple-netizen/wps-proofreading` 的独立衍生实验项目创建。

首版中以下部分沿用了原项目的架构思路并做了 ET 宿主改写：

- `js/wps-et-api.js`：由原项目 `js/wps-api.js` 的“Application / TaskPane / PluginStorage 兼容入口”思路改写；
- `js/ribbon.js`：由原项目同名文件的任务窗格创建与缓存模式改写；
- 安全写回原则：应用建议前重新核对原内容；
- 模型配置仍采用本机 OpenCode / OpenAI compatible 的方向。

没有直接迁入原项目巨大的 `proofreading-integration.js`、`taskpane.js`、规则中心和 WPS 文字 Range/修订模式逻辑。表格项目从最小实现重新开始，避免宿主逻辑交叉污染。

原项目采用 GPL-3.0 许可证；本 starter 同样保留 GPL-3.0 许可证。
