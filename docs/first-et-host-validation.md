# v0.1.0 首轮 ET 验证记录

日期：2026-10-04。开发分支：`fix/first-et-host-validation`，从远程 main `a65e55f5f2edb3f805b7fe6f9ecacffde1d6eed7` 创建。未合并。

## 环境和范围

- 实际仓库目录：`C:\Users\phong\Desktop\WPS表格核对\wps-spreadsheet-proofreading`；最初父目录为空，已克隆指定的独立仓库。
- Windows，WPS `12.1.0.28505`，本地 npm 包 `wpsjs 2.2.3`，OpenCode `1.18.34`（健康接口与本机版本一致）。
- 初始 `git status` 干净，已检查最近十条提交、执行 `npm install` 和 `npm test`；基线 3 项测试通过。首轮修复后 44 项通过；本轮复用工作完成后 `npm test` 为 66 项通过、0 失败。
- `package.json` 原有 `addonType: et` 正确，无需修改；`npm run debug` 实际端口为 `3889`。已阅读本机 wpsjs 的参数和注册实现，没有凭印象改 ET 参数或端口。
- `publish.xml` 新增独立 `wps-spreadsheet-proofreading` / `et` 记录；原 `wps-text-proofreading` / `wps` 的 `3891` 记录保留。未修改文字版仓库。文字版 UI 完整回归未实测。

## 第二轮：文字版代码复用核查

本轮按用户要求对照了本机 `wps-proofreading` 的实现；旧项目只读，没有修改。复用源是本机文字版提交 `057bc4f`：表格版新增独立 `js/opencode-client.js`，移入可通用的 OpenCode 地址校验、UTF-8 Basic 认证、模型目录解析、`provider/model` 拆分和安全错误映射。没有迁入文字版的 WPS 文字宿主调用、修订范围/字符偏移算法或权限 `ask` 实现。ET 继续使用自身的 session/tool 全禁用策略、表格寻址和精确原值写回检查。

任务窗格按文字版的卡片层级、结果状态、旧/新文本样式和设置页面重做；增加表格版自己的提供商设置存储、模型选择、手动模型输入和运行期凭据管理。兼容接口的 API Key、OpenCode 密码只保留当前面板会话内存，不写入工作簿或持久配置。原始文字版仓库未改动。

第二轮真机显示 WPS 表格功能区、打开右侧任务窗格、切换设置页均正常。点击“检测并读取模型”后，面板持续显示“正在检测并读取模型…”，未在任务窗格内确认目录返回；关闭再打开面板可恢复默认状态。本轮 WPS 的 JS 调试器连接的是 `http://127.0.0.1:3889/index.html` 主页面，未能看到独立任务窗格的请求。通过主页面调试控制台访问本机 `/config/providers` 收到 HTTP 200 并读到响应正文，但这不能证明任务窗格调用链完成。故模型目录按钮在真机**未验证成功**，需要下一轮从任务窗格 WebView 追踪并修复。

本轮没有在表格真机发起校对模型请求：默认 `opencode/big-pickle` 的同一 OpenCode 服务先前报告免费额度只允许在 OpenCode 内使用；表格版保留全部工具禁用，没有尝试降低限制。没有真实模型建议，因此本轮没有真机重测建议定位或安全写回；首轮使用临时模拟建议时已验证的 ET 行为仍按上表记录，自动化测试覆盖对应流程。

## 修改文件和问题

| 文件 | 修复或用途 |
|---|---|
| `js/util.js`、`js/ribbon.js` | 原代码缺少 GetUrlPath，任务窗格用了错误的相对地址，真机出现空白窗格。补绝对 URL，并按 URL 隔离窗格缓存。 |
| `js/wps-et-api.js` | 原身份仅依赖名称/路径，读取错误会被吞掉。捕获原工作表引用及工作簿路径与窗口句柄，校验单格地址；读取失败停止；写前重读精确原值及两种公式属性，拒绝公式形式的建议。 |
| `js/proofreading-core.js` | 原文 trim 破坏精确写回条件；只检查一个公式属性可能漏掉公式。保持原值空格并检查 Formula/FormulaR1C1/HasFormula。 |
| `js/spreadsheet-integration.js` | 严格确认单一连续选区、有效大小和 1000 格限制；传递原宿主上下文；定位失败显示提示；任一批次失败清空整轮建议。 |
| `js/model-client.js` | 原请求/响应字段及工具限制不可靠。按本机 API 修正 model 对象、text parts、权限与工具禁用、超时和清理，识别 info.error；参考文字版的 UTF-8 Basic、模型名称校验与无额外字符的文本拼接。 |
| `js/opencode-client.js`、`js/settings-store.js` | 本轮从文字版移入独立 OpenCode 通用辅助方法；新增 ET 专属设置存储，凭据只保存在会话内存，目录解析复用同一个客户端。 |
| `js/taskpane.js`、`ui/taskpane.html`、`ui/taskpane.css` | 沿用文字版任务窗格的结果卡片与设置布局，加入 ET 专属模型选项、手动输入、键盘定位和状态提示。模型目录在真机的最终返回仍待验证。 |
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
