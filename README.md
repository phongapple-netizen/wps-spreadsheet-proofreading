# WPS 表格校改

这是从 `phongapple-netizen/wps-proofreading` 的成熟经验中拆分出来的**独立实验项目**，目标是只服务 WPS 表格（ET），不与 WPS 文字版共用宿主代码和发布流程。

当前版本：`0.1.0`，定位为本地调试 starter。

## 第一版范围

代码提供的最小闭环（真机验收状态见 [首轮验证记录](docs/first-et-host-validation.md)）：

1. 在 WPS 表格中选择一个**连续单元格区域**；
2. 自动遍历选区；
3. 跳过空单元格、纯数字单元格和公式单元格；
4. 将文本单元格分批发送给模型；
5. 在任务窗格按 `A1/B2/...` 地址显示问题；
6. 点击问题卡片定位对应单元格；
7. 单条“修正”时再次核对原值；如果内容已变化或是公式，则拒绝写入；
8. 支持 OpenCode 和 OpenAI 兼容接口两种最小模型入口。

第一版**没有**整工作簿校对、规则中心、一键修正、流式输出、自动启动 OpenCode、安装包、复杂一致性复核、多区域选择。这些都留到本地真机跑通后再逐项增加。

## 项目结构

```text
.
├─ index.html
├─ main.js
├─ ribbon.xml
├─ js/
│  ├─ util.js
│  ├─ wps-et-api.js              # WPS 表格宿主适配层
│  ├─ ribbon.js
│  ├─ settings-store.js          # 表格版设置、会话凭据和模型目录
│  ├─ opencode-client.js         # OpenCode HTTP/API helpers
│  ├─ model-client.js            # OpenCode / OpenAI-compatible adapter
│  ├─ proofreading-core.js       # 纯逻辑：过滤、分批、Prompt、结果校验
│  ├─ spreadsheet-integration.js # 选区读取、定位、安全写回
│  └─ taskpane.js
├─ ui/
│  ├─ taskpane.html
│  └─ taskpane.css
└─ test/
   └─ *                            # core、宿主、模型、设置和任务窗格测试
```

## 本地运行

```bash
npm install
npm test
npm run debug
```

`package.json` 的 `addonType` 已设置为 `et`，即 WPS 表格加载项。

> 注意：不同 WPS 桌面版本的 `wpsjs debug` 注册行为可能不同。第一轮真机调试重点先确认功能区能否加载、`Application.Selection` 是否能正确返回 Range，以及任务窗格是否能访问 ET JSAPI。

## OpenCode

默认地址：

```text
http://127.0.0.1:4096
```

建议本地先手工启动：

```bash
opencode serve --hostname 127.0.0.1 --port 4096 --cors http://127.0.0.1:3889
```

具体调试端口以 `wpsjs debug` 实际输出为准；CORS 不匹配时，先把实际任务窗格来源加入 `--cors`。

当前模型值按 `provider/model` 形式填写，例如：

```text
opencode/big-pickle
```

本轮核验环境为 OpenCode `1.18.34`：`/global/health` 检测健康，创建独立 session，按 `/session/:id/message` 发送文本；最后先确认 abort 成功，再确认删除成功。客户端要求服务回显全部权限 `deny`，并禁用全部工具，发现工具结果或审批请求就中止整轮校对。权限或工具列表无法确认时也会停止。

“连接正常”只表示服务健康，不能证明模型能用。本轮默认免费模型返回服务限制，另外尝试的供应商返回余额不足；真实模型校对尚未成功。请填写本机已连接且可用的 `provider/model`，不要通过放开工具权限来解决供应商错误。

文字版复用说明：表格项目仍独立加载和发布。本轮从本机文字版仓库 `wps-proofreading` 的 `057bc4f` 复制了可独立的 OpenCode URL/凭据处理、模型目录解析和错误映射；任务窗格沿用文字版卡片、状态和设置布局。工作簿/工作表识别、选区读取、单元格定位和写回仍由表格宿主适配层独立实现。未修改文字版仓库。复用边界和本轮 WPS 实测结果见 [首轮验证记录](docs/first-et-host-validation.md)。

## 首轮真机验收建议

打开 [测试工作簿](test/fixtures/et-validation.xlsx)，依次选择 `B2`、`B2:B4`、`A2:D4`、`B2:D4`。B 列是三条有错误的文本，A/C 列为数字，D 列为公式。预期只发送 B 列文本，点击建议返回原表并选中对应地址，逐条修正只写目标单元格。再修改原文、切换工作簿或将目标改为公式，确认旧建议不能覆盖。完整步骤和未实测项目见验证记录。

## 安全边界

- 公式单元格永不写入；
- 非字符串值默认不送 AI；
- 每次应用建议前重新核对当前值；
- 绑定校对时的工作簿路径、窗口句柄和原工作表引用，拒绝跨工作簿和同名替代目标；
- 首版最多处理 1000 个选中单元格；
- 首版仅支持一个连续区域；
- AI 建议必须逐条人工确认。

## 与 WPS 文字版的关系

两个项目应保持独立：独立仓库、独立版本、独立测试、独立安装与发布。这里仅吸收已经验证过的设计原则，不计划把两个宿主重新合并。
