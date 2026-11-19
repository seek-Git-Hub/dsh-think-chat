# dsh-better-chat
当前版本0.1.0
给 [DeepSeek Harness] 的两个「**仅对话模式**」：像 DeepSeek 网页版一样聊天，支持图片多模态识别，但**会自己多轮思考**。

一个 DSH 本地 bundle，向 Web 端注册两个 agent preset，并附带 `keep_thinking`、`glob`、`grep`（联网版另加 `tool-web`）工具，以及一组 **Office 文档查看**工具（`office_render` / `office_convert` / `office_recalculate`，需加载 Office 技能后激活）。插件清单里只显示**一个**插件：`dsh-better-chat`。

| 预设 | id | order | 联网 |
|---|---|---|---|
| 仅对话模式(不联网版) | `chat-offline` | 0 | 不联网 |
| 仅对话模式(联网版) | `chat-online` | 1 | 默认不联网，用户明确要求时才联网 |

---

## 它是什么

- **两个模式（agent preset）**，在 DSH 的 Agent 预设页出现，名字分别是「仅对话模式(不联网版)」与「仅对话模式(联网版)」
- **共用一份 `keep_thinking`**：纯文字的控制流工具，两个预设各挂一次，各自独立计数、互不影响
- **`glob` `grep`**：只查找和搜索文件，不修改任何文件
- **Office 文档查看**：挂载 DSH 自带的 `tool-office` + `skill-office` + `tool-skill`，加载 Office 技能后可用 `office_render` 把 `.docx` / `.xlsx` / `.pptx` / `.pdf` 的页面渲染成图片给模型看，也可以用 `office_convert` 转成 `txt`/`csv`，再用 `grep` 把文字读回对话；不写工作区文件
- **联网版独有的 `tool-web`**：通过 `persona` 系统提示词的强力约束实现**「按需联网」**
  - **默认状态**：纯对话，绝不擅自联网，基于内部知识库作答
  - **触发条件**：用户在提示词中明确提出“搜一下”、“联网查”、“看最新新闻”等要求
- **支持多模态**：用户可以直接上传图片，模型能识别图片内容
- 模型**自己决定**要想 1 轮还是 N 轮，**自己判断**什么时候该停
- 两个预设都只有一个硬上限防止无限循环
- **无文件修改、无产出物**：不挂文件写入工具，不挂 present

### 为什么合并成一个插件

两个预设共用同一份 `keep_thinking` 实现、同一套 persona 结构，拆成两个包只是为了让插件清单出现两行。合并后：

- 一份代码（`lib/index.js`）、一份 README、插件清单里一个条目
- 预设的 `id`、名称、描述、`order`、persona 提示词、工具行**全部不变**
- Agent 预设界面仍是原来的两张卡片，行为不变

### 「多轮思考」是怎么实现的

DSH 的 agent 循环规则是：**让轮次继续的唯一燃料是「工具调用」或「steering」**。没有工具调用，模型答完就结束。

所以这里给模型一个**只用来表达"我还没想完"的空工具**：

| 模型的行为 | 循环的行为 |
|---|---|
| 调用 `keep_thinking` | 执行工具 → **再跑一步**（又一次模型请求） |
| 不调用、直接输出答案 | **轮次结束** |

这就是 agent loop 原生的控制流，不需要监听轮次边界、不需要往对话正文里塞协议标记。**「停止」就等于「不再调用」**。

## 挂载了哪些插件

**全部内容就是 [`cordis.patch.yml`](./cordis.patch.yml) 里的两个 preset 声明。**

### 宿主层（profile 层）

**空的。** 本 bundle 不在 profile 层放任何工具行——所有插件都挂在 preset 内部，原因见下。

### preset 层（模式内部）

`chat-offline` 挂 6 行：

| 行 id | 包 | 配置 | 提供的工具 |
|---|---|---|---|
| `think-better-tool` | `dsh-better-chat`（本仓库，按包名引入） | `maxRounds: 8` | `keep_thinking` |
| `fs-search` | `@deepseek-ai/dsh-tool-fs-search` | `sampleOverCapGlobResults: false` | `glob、grep` |
| `tool-skill` | `@deepseek-ai/dsh-tool-skill` | — | `skill`（让模型/用户加载技能） |
| `office` | `@deepseek-ai/dsh-tool-office` | — | `office_render、office_convert、office_recalculate`（加载 Office 技能后激活） |
| `skill-office` | `@deepseek-ai/dsh-skill-office` | — | `office、office-docx、office-xlsx、office-pptx` 四个技能 |
| `persona` | `@deepseek-ai/dsh-persona` | 见 persona | —（系统提示词） |

> `tool-skill` 必须排在 `office` / `skill-office` **之前**，否则 Office 三件套不会正常激活——这是踩过的坑。

`chat-online` 挂 7 行：上面 6 行再加

| 行 id | 包 | 配置 | 提供的工具 |
|---|---|---|---|
| `tool-web` | `@deepseek-ai/dsh-tool-web` | — | 联网搜索、网页抓取 |

**合计：不联网版常驻 4 个工具（`keep_thinking`、`glob`、`grep`、`skill`），联网版 5 个；加载 Office 技能后再多出 3 个 `office_*`。** 宿主层其它 bundle 还会注入约 13 个工具。

> 因为工具行**就在 preset 里**，`ctx.tools.register()` 走的是 preset 的作用域：只有对应的「仅对话模式」agent 看得见它，其它模式**不需要被摘掉任何东西**。两个预设各挂一次本包，得到两个独立的 `keep_thinking` 实例。

### Office 文档查看怎么用

- 工具都由 DSH 自带包提供，本插件只负责在 preset 里挂三行：
  `@deepseek-ai/dsh-tool-office`（工具）、`@deepseek-ai/dsh-skill-office`（技能）、`@deepseek-ai/dsh-tool-skill`（`skill` 工具）。
- `office_*` 三个工具**默认不出现在工具列表里**：模型（或用户用 `/office`）先加载 Office 技能，从**下一个请求**起才激活。
- **两条读取路线，都不需要 shell：**

  | 想要什么 | 用什么 | 结果 |
  |---|---|---|
  | 文字内容（正文、标题、台词） | `office_convert` → `format: "txt"` | 返回 Harness 私有临时目录里的 `.txt` 路径；再对这个路径用 **`grep` 的 `pattern` 匹配行**即可把文字读进对话 |
  | 版面/图片/图文关系 | `office_render` → `pages` | 直接返回页面图片给模型看，不产生文件 |

  实测流程：`office_convert(某个 .docx → txt)` 得到临时路径，`grep` 以该路径为 `path`、`pattern: ".+"` 即可拿到逐行文本——**全程没有 shell**，`grep` 读的就是那个私有临时文件。
- **限制**：
  - 临时文件有保留期，对话结束即删；要在对话里复用，必须当轮读回。
  - 渲染保真度与分页取决于 LibreOffice Kit，可能与 Microsoft Office / WPS 不同。
  - Web GUI 的 Office 行只显示文字、不显示渲染出的图片（**模型能看到**）。
  - 单次渲染受 `maxImages`、像素与字节上限约束。
  - `office_convert` 不接受 PDF 输入（PDF 只能 `office_render`）。

### 刻意没有挂的

| 没挂 | 原因 |
|---|---|
| `dsh-tool-fs` | 只做 UTF-8 文本读写，读 `.docx` 只会得到乱码；而且它会带来写文件能力，与"无文件修改"冲突。Office 查看走 `dsh-tool-office`。 |
| `dsh-tool-present` | 不产出文件 |
| `dsh-tool-goal` + `dsh-command-goal` | `goal-round-driver` 在 base 中默认启用，模型一旦能建 goal 就会**自动续轮**，直接破坏"模型自己决定何时停" |
| `dsh-tool-ralph` | 天生自我循环 |
| `dsh-tool-subagent` / `dsh-tool-workflow` / agent-team | 请求数成倍放大 |
| `dsh-tool-todo` | 会把人格拉向"干活" |
| `dsh-tool-pwsh` / `dsh-tool-bash` / `dsh-tool-jobs` / `dsh-tool-terminal` | 重型，与纯对话定位冲突。Office 读取用 `office_convert` + `grep` 读回临时文件即可，不需要 shell |
| `dsh-agent-instructions` | **必须不挂**，否则会去读 `AGENTS.md`，模型一上来就进入编码 agent 人格 |
| `dsh-plan-mode` | 与纯对话模式无关 |

### 不用挂、base 已经有的

`dsh-compaction-basic`（长对话自动压缩）、`dsh-fs-observation-policy`（先读后写门禁）、`dsh-repeat-tool-reminder`、`dsh-tool-call-timeout-policy`、`dsh-session-title` 系列、`dsh-subprocess-local`（子进程后端）、`dsh-spill-local`（超限结果落盘后端）。

> ⚠️ `dsh-subprocess-local` / `dsh-spill-local` 已在宿主层提供 `subprocess` / `spillStore`。preset 里再挂这两行会重复注册同名服务，导致**整个模式加载失败**（DSH 0.2.2 起），所以它们不在上面的行列表里。

---

## 配置参数

### `maxRounds` —— 唯一需要理解的参数

两个 preset 各有一行，**改轮数要两处都改**：

```yaml
- id: preset-chat-offline      # 不联网版
  ...
          - id: think-better-tool
            name: 'dsh-better-chat'
            config:
              maxRounds: 8      # ← 必须是 ≥1 的整数

- id: preset-chat-online       # 联网版
  ...
          - id: think-better-tool
            name: 'dsh-better-chat'
            config:
              maxRounds: 8      # ← 同上
```

## 目录结构

```
dsh-better-chat/
├── package.json        # 名字、exports、peerDependencies、dsh.bundle.patch
├── cordis.patch.yml    # 两个 preset 声明（本插件的全部内容）
├── lib/
│   └── index.js        # keep_thinking 工具实现（两个 preset 共用）
├── locale/
│   ├── en.json
│   ├── zh.json
│   └── zh-CN.json
├── README.md
└── LICENSE
```

修改 `lib/index.js` 或 `cordis.patch.yml` 后，到 DSH 的插件清单里禁用再启用本插件，或重启 DSH 生效；已经开始的会话保持启动时的插件版本，需要新建会话验证。
