---
name: word-format
description: 把一份参考 Word (.docx) 的排版样式（字体/字号/缩进/行距/对齐/页面设置/样式表/主题）复刻到新内容上。当用户要求「按某个 .docx 模板排版」「在已有文档基础上续写」「修复后半部分排版混乱」「统一字体字号缩进」等场景时使用。仅在 macOS 工作（依赖 Microsoft Word + 浏览器 + System Events 的 UI 自动化）。
---

# word-format

把参考 Word 文档的排版**视觉一致地**复刻到新内容上。
工作方式：**直接在原件 HTML 上增删改文字**，再让浏览器渲染并把富文本送进 Word。

## 工作原理：HTML Bridge

```
参考 .docx
    │  ① docx_to_html.py
    ▼
Word Filtered HTML（原件，所有样式 inline）
    │  ② Claude 在原件上增删改文本（保留所有 inline style）
    ▼
edited.html
    │  ③ render_and_paste.sh
    │     浏览器渲染 → Cmd+A/C → Word 粘贴 → 保存
    ▼
最终 .docx
```

**步骤 ③ 有两种模式，决定能保留多少模板格式：**

| 模式 | 命令 | 能保留 |
|---|---|---|
| **B. 续写（推荐）** | `--append-to <reference.docx>` | **100% 模板格式**：页面设置、样式表、主题、页眉页脚、字体表 + 字符级排版 |
| A. 新建 | （不带 `--append-to`） | 仅字符/段落级排版（字体/字号/缩进/行距/对齐） |

> **想 100% 保留模板格式 → 必须用模式 B。**
> 浏览器粘贴管线本身只能传字符级直接属性，不传`@page`/样式表/theme1.xml。
> 模式 B 的诀窍：**不靠粘贴管线传这些**——直接复制原 .docx 副本作为承载文档，新内容用 Cmd+A → → → Cmd+V 续写到末尾，原文档的页面设置 / 样式表 / 主题原封不动地继承下来。

## 为什么不直接让 Word 打开生成的 HTML 再另存为 .docx

走「Word 打开 HTML → 另存 docx」的是 Word 的 **Open Web Page 遗留导入器**，对通用 HTML 行为保守诡异：

- 字体回退链评估方式与浏览器不同，`font-family: '宋体', SimSun, serif` 可能落到默认字体
- `text-indent`、`line-height` 在导入时会被 clamp 到 Word 默认范围
- `<body>` 的 padding/margin 不严格遵守 CSS spec

而**当前方案：浏览器渲染 → 系统剪贴板 → Word 粘贴**，走的是 Word 的「粘贴外部富文本」管线：

- 浏览器是严格的 CSS 渲染器，**所见即所得**
- 字体在浏览器渲染那一刻就钉死，剪贴板里携带的是确定的字体名
- Word 粘贴管线把剪贴板里的 HTML+RTF **直接展开成段落直接属性**，跳过 HTML 导入器

对 inline-CSS HTML 而言，这是中间损失最小的链路。

## 环境要求

- **OS**: macOS（脚本依赖 AppleScript / System Events）
- **Microsoft Word for Mac**（`/Applications/Microsoft Word.app`）
- **浏览器**: Google Chrome（首选）或 Safari
- **Python 3**：系统自带，无第三方依赖
- **首次运行须开启系统授权**：
  - 系统设置 → 隐私与安全性 → **自动化**：勾选允许「终端宿主进程（Terminal / iTerm / Ghostty / Claude Code）」控制 Microsoft Word、Chrome（或 Safari）、System Events
  - 系统设置 → 隐私与安全性 → **辅助功能**：启用同一个宿主进程

环境检查：
```bash
bash "${SKILL_DIR}/scripts/verify_env.sh"
```

## ⚠️ 给 Claude 的硬约束 prompt（编辑 HTML 时必须遵守）

> 请帮我生成一段**专门用于复制粘贴到 Word 文档**的 HTML 排版代码。
> 核心要求如下：
>
> 1. **必须采用"行内样式"（Inline CSS）的写法**：请把所有的样式规则（如字体、字号、间距）直接写在每一个 HTML 标签的 `style` 属性里面，不要使用 `<style>` 标签或者外部 CSS，以确保 Word 能够完整读取格式。
> 2. **严格使用"pt"（磅）作为单位**：请务必把字体大小的单位设定为 `pt`，绝对不要使用 `px`，以防止因屏幕缩放而导致字号出现误差。
> 3. **强制表格居中**：请不要使用 CSS 的 `margin: auto`，必须直接在 `<table>` 标签上添加 `align="center"` 属性（例如 `<table align="center" ...>`），这是 Word 唯一能识别的居中方式。
> 4. **防止页面偏移**：请确保 `<body>` 标签没有 padding 或 margin，防止复制后产生左侧缩进。
> 5. **宽度控制**：大表格请设定 `style="width:440pt"`（适应 A4 版心），小表格请设定 `style="width:auto"`。

### 在"就地编辑原件"模式下的实操含义

- **核心动作**：Claude 用 `Read` 读原件 → `cp` 复制为 `edited.html` → 用 `Edit` 工具**只改文字节点的内容**。所有 inline `style`、嵌套结构、`<p>`/`<span>`/`<table>` 标签**原样保留**。
- **第 1、2、5 条**：原件本来就是 Word 自己导出，已经符合（inline + pt + 表格宽度合理）。**只要不引入新的 `<style>` 块、不引入 px、不引入 `margin:auto`**，就是合规。
- **第 3 条**：如果原件里某个 `<table>` 视觉上居中但缺 `align="center"`，编辑时给它补上。
- **第 4 条**：原件 `<body>` 通常是 `<body lang=ZH-CN style='...'>`，**编辑时把 style 改成 / 补上 `margin:0;padding:0;`**。
- **额外禁止**：在 HTML 文本节点里写 Markdown 语法（`**加粗**` / `# 标题` / `- 列表`）；凭空写原件中没出现过的字体名。

## 完整流程

### 步骤 1：导出参考 .docx 为 HTML

```bash
python3 "${SKILL_DIR}/scripts/docx_to_html.py" "/path/to/reference.docx"
# 产物：~/Library/Caches/word-format-skill/<basename>.html
```

驱动 Word 自身用「另存为 → 筛选过的网页」导出。Filtered HTML 把字体、字号、缩进、行距等全部以 inline CSS 写进每个标签——这是唯一能 100% 保留 Word 排版信息的文本格式。

### 步骤 2：Claude 在原件上增删改

按场景选编辑方式：

#### 场景 ★ 续写（最常用，与模式 B 配套）

只产出**要追加的新内容**，从原件里复制一段同类型段落（含完整 inline style），改文字。

```bash
# Claude 选一段原件中已有的同类型段落作为模板段，只产出"待追加内容"：
cat > ~/Library/Caches/word-format-skill/append.html <<'HTML'
<p style="margin:0;font-family:'宋体';font-size:12.0pt;text-indent:24.0pt;line-height:150%;text-align:justify;">
新增的正文段落，照抄原件 inline style 写法，只换文字内容。
</p>
<p style="...">…</p>
HTML
```

> 这种 `append.html` 不需要完整 `<html>/<body>` 外壳——剪贴板复制的是渲染后的富文本，浏览器会把零散段落正常渲染。但若想严谨，可以包一层 `<html><body style="margin:0;padding:0;">…</body></html>`。

#### 场景 整文重写（与模式 A 配套）

```bash
cp ~/Library/Caches/word-format-skill/<basename>.html \
   ~/Library/Caches/word-format-skill/<basename>.edited.html
```

然后用 `Edit` 工具就地修改 `<basename>.edited.html`：**只动文本节点的内容**，所有标签 / `style` / 嵌套结构原样保留；同时把 `<body ...>` 的 style 改成包含 `margin:0;padding:0;`。

### 步骤 3：渲染 + 复制 + 粘贴 + 保存

#### 模式 B：续写（推荐，100% 保留模板）

```bash
bash "${SKILL_DIR}/scripts/render_and_paste.sh" \
    --append-to "/path/to/reference.docx" \
    ~/Library/Caches/word-format-skill/append.html \
    ~/Desktop/最终输出.docx
```

脚本顺序：
1. 把 `reference.docx` 复制为 `~/Desktop/最终输出.docx`
2. 浏览器加载 `append.html`，等 2.5s 字体加载完
3. 在浏览器里 Cmd+A / Cmd+C
4. Word 打开 `~/Desktop/最终输出.docx`（即副本，原模板设置全在）
5. 等 Word 成为前台进程
6. Cmd+A → 右方向键（光标塌缩到文档末尾）→ Cmd+V（粘到末尾）
7. `save active document`（不是 save as，副本已命名）
8. 把焦点还给运行前的前台应用

#### 模式 A：新建空白文档

```bash
bash "${SKILL_DIR}/scripts/render_and_paste.sh" \
    ~/Library/Caches/word-format-skill/<basename>.edited.html \
    ~/Desktop/最终输出.docx
```

脚本顺序：
1. 浏览器加载 `<basename>.edited.html`，等 2.5s
2. Cmd+A / Cmd+C
3. Word 新建空白文档
4. Cmd+V 粘贴
5. `save as` 为指定 `.docx`
6. 把焦点还给运行前的前台应用

⚠️ **脚本运行期间（约 5~7 秒）不要使用键盘 / 鼠标**，UI 自动化在跑。

## 故障排查

| 症状 | 原因 / 处理 |
|------|------|
| 步骤 1 报 `-1728` | Word 自动化授权未开 → 系统设置 → 隐私与安全性 → 自动化；或 Word 处于残留状态（脚本会自动 quit + retry 一次） |
| 步骤 3 键盘事件没生效（错误码 1002） | 辅助功能权限未开 → 系统设置 → 隐私与安全性 → 辅助功能 → 把宿主进程加入并打开开关 |
| 模式 B 续写后页面设置仍是 Word 默认 | `--append-to` 没指对源 docx；或 Word 把粘贴内容放进了新节并应用了节属性，检查目标 docx 是否仍是单节 |
| 粘贴后字体变 Word 默认中文 | 编辑时引入了原件没有的字体名 |
| 段落整体左侧出现空白 | `<body>` 没设 `margin:0;padding:0;`（违反契约 4） |
| 字号在 Word 里轻微偏移 | 出现了 `px` 单位（违反契约 2） |
| 表格不居中 | `<table>` 缺 `align="center"`（违反契约 3） |
| 出现 Markdown 文本 `**` `#` | 编辑时夹带了 Markdown，必须用 HTML 标签 |
| 粘贴后 Word 文档为空白 | Cmd+V 在 Word 还没拿到键盘焦点时落空；脚本里已通过"等待 frontmost = Microsoft Word"修复，若仍出现可加大 `delay` |

## 核心原则

> **不要擅自提取格式、不要重写片段。**
> Word Filtered HTML 是"全 inline 化的可渲染页面"，最忠实的复刻方式就是**保留它所有的样式标签，只换文字**。
>
> **想 100% 保留模板（含页面设置 / 样式表 / 主题）→ 用续写模式（`--append-to`），让原 .docx 副本承载，新内容粘到末尾。**
> 这是 AI 生成 HTML 进入 Word 时损失最小的链路。
