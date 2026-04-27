# word-format-skill

> 一个 Claude Code skill：把一份参考 Word (`.docx`) 的排版样式（字体 / 字号 / 缩进 / 行距 / 对齐 / **页面设置 / 样式表 / 主题 / 页眉页脚**）**视觉一致地**复刻到新内容上。

仅在 macOS 工作（依赖 Microsoft Word + 浏览器 + System Events 的 UI 自动化）。

---

## 它解决什么问题

让 AI 写一份 Word 文档时，最痛苦的不是写内容，而是排版——把 AI 输出的 Markdown / 纯文本粘进 Word，字体、字号、缩进总会乱。

这个 skill 不让 AI 凭空生成 HTML / OOXML，而是：**让 Word 自己导出一份 inline-CSS 的 Filtered HTML，AI 在这份"原件"上只改文字内容，所有样式标签原样保留。** 然后浏览器把它渲染成富文本，进系统剪贴板，粘进 Word。

```
参考 .docx
   │ ① docx_to_html.py（驱动 Word 自己另存为筛选过的网页）
   ▼
Word Filtered HTML（所有样式以 inline CSS 写入每个标签）
   │ ② Claude 在原件上增删改文字（保留所有 inline style）
   ▼
edited.html / append.html
   │ ③ render_and_paste.sh
   │   浏览器渲染 → 系统剪贴板（HTML+RTF）→ Word 粘贴 → 保存
   ▼
最终 .docx
```

不走「AI 生成 HTML 代码」这条路，AI 始终编辑的是 Word 自己产出的"标准模板"——这是当前所知 AI → Word 损失最小的链路。

---

## 两种工作模式

| 模式 | 命令 | 能保留 |
|---|---|---|
| **B. 续写（推荐）** | `--append-to <reference.docx>` | **100% 模板**：页面设置、样式表、主题、页眉页脚、字体表 + 字符级排版 |
| A. 新建 | （不带 `--append-to`） | 仅字符 / 段落级排版（字体 / 字号 / 缩进 / 行距 / 对齐） |

> 想 100% 保留模板格式 → **必须用模式 B**。诀窍是把原 `.docx` 副本作为承载文档，新内容只追加到末尾，原文档的页面设置 / 样式表 / 主题原封不动地继承下来——浏览器粘贴管线本身只能传字符级直接属性，不传 `@page` / 样式表 / theme1.xml。

---

## 环境要求

- **OS**：macOS
- **Microsoft Word for Mac**（`/Applications/Microsoft Word.app`）
- **浏览器**：Google Chrome（首选）或 Safari
- **Python 3**：系统自带，无第三方依赖

首次运行须开启系统授权：

- 系统设置 → 隐私与安全性 → **自动化**：勾选允许「终端宿主进程（Terminal / iTerm / Ghostty / Claude Code）」控制 Microsoft Word、Chrome（或 Safari）、System Events
- 系统设置 → 隐私与安全性 → **辅助功能**：启用同一个宿主进程

环境检查：

```bash
bash skills/word-format/scripts/verify_env.sh
```

---

## 安装

### 通过 Claude Code 插件商店

```
/plugin marketplace add yeap531/word-format-skill
/plugin install word-format@word-format-skill
```

### 或手动 clone

```bash
git clone https://github.com/yeap531/word-format-skill.git
```

---

## 快速开始

```bash
# 1. 把参考 docx 导出为 inline-CSS HTML
python3 skills/word-format/scripts/docx_to_html.py "/path/to/reference.docx"
# → 产物：~/Library/Caches/word-format-skill/<basename>.html

# 2. （Claude 这一步）从原件挑一段同类型段落作模板，复制完整 inline style，
#    在 ~/Library/Caches/word-format-skill/append.html 写要追加的新内容
#    所有 inline style 原样保留，只换文字。

# 3. 续写模式：复制原 docx 副本 → 末尾粘贴 → 保存
bash skills/word-format/scripts/render_and_paste.sh \
    --append-to "/path/to/reference.docx" \
    ~/Library/Caches/word-format-skill/append.html \
    ~/Desktop/最终输出.docx
```

第 2 步 Claude 必须遵守的 8 条硬约束（inline CSS / pt 单位 / 表格 `align="center"` / `<body>` 无 padding margin / 字体白名单 等）见 [`skills/word-format/SKILL.md`](skills/word-format/SKILL.md)。

⚠️ 第 3 步运行期间（约 5~7 秒）不要动鼠标键盘，UI 自动化在跑。脚本结束会把焦点还给运行前的前台应用。

---

## 文件结构

```
word-format-skill/
├── .claude-plugin/
│   └── marketplace.json            # Claude Code 插件商店配置
├── skills/
│   └── word-format/
│       ├── SKILL.md                # Skill 核心指令（8 条契约 + 完整流程 + 故障排查）
│       └── scripts/
│           ├── docx_to_html.py     # 步骤 ①：参考 docx → Filtered HTML
│           ├── render_and_paste.sh # 步骤 ③：渲染 + 粘贴 + 保存（支持续写模式）
│           └── verify_env.sh       # 环境检查
├── README.md
└── LICENSE                         # Apache 2.0
```

---

## 致谢

核心思路源自 linux.do 社区的这篇帖子：

> [**AI 快速排版 Word 文章的一个通用思路**](https://linux.do/t/topic/1217729)
>
> 作者提出「让 AI 生成带 inline CSS 的 HTML → 浏览器中转 → 粘贴进 Word」的排版方案，以及一套 8 条 prompt 硬约束（inline CSS / pt 单位 / 表格 `align="center"` 等）。

本 skill 在此基础上做了两点关键调整：

1. **不让 AI 凭空写 HTML**——而是用 Word 自己导出的 Filtered HTML 作为样式模板，AI 只改文字内容、保留所有 inline style 标签。这样能避免 AI 生成 HTML 时的字体名误用、单位混用等隐性问题。
2. **支持"续写"模式**——把原 `.docx` 副本作为承载文档，新内容粘到末尾，从而 100% 继承原文档的页面设置 / 样式表 / 主题 / 页眉页脚（这些信息浏览器粘贴管线传不过去）。

---

## License

Apache 2.0 — 详见 [LICENSE](./LICENSE)
