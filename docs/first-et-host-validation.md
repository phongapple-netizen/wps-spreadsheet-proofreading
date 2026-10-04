# v0.1.0 首轮 ET 验证记录

日期：2026-10-04。开发分支：`fix/first-et-host-validation`，从远程 main `a65e55f5f2edb3f805b7fe6f9ecacffde1d6eed7` 创建。未合并。

## 环境和范围

- 实际仓库目录：`C:\Users\phong\Desktop\WPS表格核对\wps-spreadsheet-proofreading`；最初父目录为空，已克隆指定的独立仓库。
- Windows，WPS `12.1.0.28505`，本地 npm 包 `wpsjs 2.2.3`，OpenCode `1.18.34`（健康接口与本机版本一致）。
- 初始 `git status` 干净，已检查最近十条提交、执行 `npm install` 和 `npm test`；基线 3 项测试通过。修复后完整 `npm test` 44 项通过、0 失败。
- `package.json` 原有 `addonType: et` 正确，无需修改；`npm run debug` 实际端口为 `3889`。已阅读本机 wpsjs 的参数和注册实现，没有凭印象改 ET 参数或端口。
- `publish.xml` 新增独立 `wps-spreadsheet-proofreading` / `et` 记录；原 `wps-text-proofreading` / `wps` 的 `3891` 记录保留。未修改文字版仓库。文字版 UI 完整回归未实测。

## 修改文件和问题

| 文件 | 修复或用途 |
|---|---|
| `js/util.js`、`js/ribbon.js` | 原代码缺少 GetUrlPath，任务窗格用了错误的相对地址，真机出现空白窗格。补绝对 URL，并按 URL 隔离窗格缓存。 |
| `js/wps-et-api.js` | 原身份仅依赖名称/路径，读取错误会被吞掉。捕获原工作表引用及工作簿路径与窗口句柄，校验单格地址；读取失败停止；写前重读精确原值及两种公式属性，拒绝公式形式的建议。 |
| `js/proofreading-core.js` | 原文 trim 破坏精确写回条件；只检查一个公式属性可能漏掉公式。保持原值空格并检查 Formula/FormulaR1C1/HasFormula。 |
| `js/spreadsheet-integration.js` | 严格确认单一连续选区、有效大小和 1000 格限制；传递原宿主上下文；定位失败显示提示；任一批次失败清空整轮建议。 |
| `js/model-client.js` | 原请求/响应字段及工具限制不可靠。按本机 API 修正 model 对象、text parts、权限与工具禁用、超时和清理，识别 info.error；参考文字版的 UTF-8 Basic、模型名称校验与无额外字符的文本拼接。 |
| `test/proofreading-core.test.js`、新增 `test/wps-et-api.test.js`、`test/ribbon.test.js`、`test/model-client.test.js`、`test/spreadsheet-integration.test.js` | 对上述错误及失败关闭路径补回归测试。 |
| `test/fixtures/et-validation.xlsx` | 指定四列表格，D2:D4 为原生公式，计算值 11/22/33。 |
| `.gitignore`、`package-lock.json` | 排除调试产物、临时诊断和构建目录；锁定依赖。 |
| `README.md`、本记录 | 更新实际端口、模型安全要求、测试表与验收状态。 |

## 真机结果

下面宿主测试使用标明“模拟模型”的临时脚本，仅返回指定三条建议；读取、定位和写回都调用真实 ET API。临时脚本已从正式页面移除，未提交。模拟建议不能替代真实模型验收。

移除临时脚本后已重新打开正式任务窗格；在 B2:D4 的 9 格选区点击校对，正式页面成功发出 OpenCode 请求并显示免费模型的 APIError，建议数保持 0、按钮恢复可用。该次未取得建议，不能视为模型闭环通过。调试过程中出现过页面状态重置；本机 wpsjs 对整个项目目录变化触发热刷新，因此正式复验期间应避免同时修改项目文件或 Git 元数据。

| 验收项目 | 结果和证据 |
|---|---|
| 加载与右侧窗格 | 已实测：“WPS 表格校改”标签显示，点击打开右侧完整面板，修复前为空白。 |
| B2 单格 | 已实测：工作簿/原表上下文、地址 B2、Value2 和两种公式属性正确，HasFormula=false。 |
| B2:B4 | 已实测：纳入 B2/B3/B4，模拟问题各自带正确地址。 |
| A2:D4 | 已实测：12 格中只纳入 B2/B3/B4，跳过 9 格；A2 Value2=1，D2 Value2=11、Formula=`=A2+C2`、FormulaR1C1=`=RC[-3]+RC[-1]`、HasFormula=true。 |
| 点击卡片定位 | 已实测：点击 B3 选中 B3；新增并切换另一表后点击卡片，返回“校改测试”并选中 B3。 |
| 单条安全写回 | 已实测：“修正”将 B3 改为“疏散通道堆放杂物。”，B2/B4 未被此操作修改。 |
| 原值变化保护 | 已实测：B2 内容改变后，旧建议显示“单元格内容已变化，请重新校对。”，保留当前内容。 |
| 工作簿切换保护 | 已实测：打开 et-other.xlsx 后点击原 B4 卡片不能定位，点击修正也拒绝写入，新工作簿保持原值。 |
| 同名不同工作簿 | 尝试实测时本机 WPS 拒绝同时打开两个同名文件。插件防护未实测；自动测试覆盖不同路径和相同路径/不同窗口句柄。 |
| 原工作表删除、无效地址 | 未实测；自动测试已覆盖拒绝定位与写入，包括同名替代工作表。 |
| 数字/公式以外的空格、空值、非文本；多 Areas、1000 格上限 | 自动测试通过；这些边界未全部真机实测。 |

## OpenCode 实测

健康、session 创建、model 参数、message 请求、APIError 响应结构、权限列表、工具 ID 枚举均已核验。实际中止和删除返回 HTTP 200、JSON `true`。本轮没有自动启动 OpenCode，也没有修改供应商配置或密钥。

| 模型 | 实测结果 |
|---|---|
| `opencode/mimo-v2.6-flash-free` | 真实请求失败：免费服务提示只能从 OpenCode 内使用。 |
| `zhipuai/glm-5.3-flash` | 真实请求失败：余额不足或无可用资源包。先前短超时未得到结果，后续正式请求取得此 APIError。 |
| MiniMax | 真实探测失败：余额不足。 |

没有取得真实校对 JSON，所以“WPS 选区 → 真实模型 → 建议 → 写回”的完整闭环尚未通过。当前缺少可调用的模型服务资源；不能把健康检查或模拟建议视为通过，也不能为兼容供应商放开工具。

安全语义按 [OpenCode v1.18.34 prompt 实现](https://github.com/anomalyco/opencode/blob/v1.18.34/packages/opencode/src/session/prompt.ts) 与 [permission 实现](https://github.com/anomalyco/opencode/blob/v1.18.34/packages/opencode/src/permission/index.ts) 核实：message.tools 会替换 session permission，因此除所有枚举工具 false 外，必须保留 `tools["*"]=false`。创建 session 使用 wildcard deny，返回未确认 deny、权限监测失败或返回 tool part 都拒绝结果；先确认服务中止，再删除 session。

## 下一次真机复验

1. 在独立仓库运行 `npm run debug`，检查实际输出端口与 ET 注册；不要覆盖文字版注册或使用其 3891 端口。
2. 打开测试工作簿的新副本，设置一个确实可用的 OpenCode `provider/model`。连接检测只证明 health。
3. 依次选择 B2、B2:B4、A2:D4、B2:D4；确认只检查 B 文本，并收到带 B2/B3/B4 地址的真实建议。
4. 切换到另一张表，点击卡片验证返回原表；逐条修正，检查 A/C 数字及 D 公式未变。
5. 校对后先修改原文或改为公式，再点击旧建议修正，必须拒绝覆盖。切换工作簿后同样拒绝定位/写回。
6. 在测试副本删除原表并创建同名表，旧建议仍应拒绝；此项本轮未实测。

保持单连续区域、最多 1000 格、逐条人工确认。不扩展整表扫描、批量修改、并发、安装包或文字版宿主逻辑。
