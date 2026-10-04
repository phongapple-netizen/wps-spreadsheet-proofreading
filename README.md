# WPS 表格校改

这是从 `phongapple-netizen/wps-proofreading` 的成熟经验中拆分出来的**独立实验项目**，目标是只服务 WPS 表格（ET），不与 WPS 文字版共用宿主代码和发布流程。

当前版本：`0.1.0`，定位为本地调试 starter。

## 第一版范围

已经实现的最小闭环：

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
│  ├─ model-client.js            # OpenCode / OpenAI compatible
│  ├─ proofreading-core.js       # 纯逻辑：过滤、分批、Prompt、结果校验
│  ├─ spreadsheet-integration.js # 选区读取、定位、安全写回
│  └─ taskpane.js
├─ ui/
│  ├─ taskpane.html
│  └─ taskpane.css
└─ test/
   └─ proofreading-core.test.js
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
opencode serve --hostname 127.0.0.1 --port 4096 --cors http://127.0.0.1:3888 --cors http://127.0.0.1:3891
```

具体调试端口以 `wpsjs debug` 实际输出为准；CORS 不匹配时，先把实际任务窗格来源加入 `--cors`。

当前模型值按 `provider/model` 形式填写，例如：

```text
opencode/mimo-v2.6-flash-free
```

首版 OpenCode 客户端按 `/session` 与 `/session/:id/message` 接口发送消息；如果你本机 OpenCode 版本的 API 已变化，优先在 `js/model-client.js` 做单点调整，不要动表格宿主逻辑。

## 首轮真机验收建议

先新建一个简单表格：

| A | B |
|---|---|
| 序号 | 隐患描述 |
| 1 | 现场发现在安全隐患。 |
| 2 | 疏散通到堆放杂物。 |
| 3 | `=A2+1`（公式示例） |

只框选 B2:B4。预期：文本被送审，公式单元格跳过；结果卡片显示单元格地址；点击卡片定位；修正只写入目标单元格。

## 安全边界

- 公式单元格永不写入；
- 非字符串值默认不送 AI；
- 每次应用建议前重新核对当前值；
- 首版最多处理 1000 个选中单元格；
- 首版仅支持一个连续区域；
- AI 建议必须逐条人工确认。

## 与 WPS 文字版的关系

两个项目应保持独立：独立仓库、独立版本、独立测试、独立安装与发布。这里仅吸收已经验证过的设计原则，不计划把两个宿主重新合并。
