# word-format-skill

> **一句话描述**：告别"复制进 Word 格式全乱"——让 Claude 直接帮你生成格式完整的中文 `.docx` 文件，字体、字号、缩进、行距全部精确写入，一步到位。

---

## 这个东西解决了什么痛点？

你有没有遇到过这种情况：

- 让 AI 写了几千字的报告，复制粘贴进 Word 之后——字体乱、字号乱、缩进全没有
- 花在手动调格式上的时间，比让 AI 写内容还多
- 调好了这一段，下一段又乱了

这是个结构性问题，不是 AI 写得不好。**纯文本/Markdown 粘贴进 Word，Word 对格式的解析本来就不可靠。**

---

## 这个 Skill 的解法

不走"生成文本 → 复制粘贴"这条路，**直接让 Claude 用 [docx](https://github.com/dolanmiu/docx) 库生成原生 `.docx` 文件**。

格式被精确写进 Word 的内部 XML，而不是依赖剪贴板兼容性。

```
你描述排版要求（宋体小四/首行缩进/1.5倍行距...）
              ↓
   Claude 生成 docx-js 脚本
              ↓
     Node.js 执行 → output.docx
              ↓
         Word 自动打开，格式完整
```

**不需要浏览器。不需要手动复制粘贴。不需要再手动调整任何格式。**

---

## 功能一览

| 功能 | 说明 |
|------|------|
| 中文字号全覆盖 | 初号到八号，小四/小三/三号等全部支持 |
| 中文字体 | 宋体、黑体、楷体、仿宋、微软雅黑等 |
| 排版要素 | 首行缩进、行距、段间距、对齐方式 |
| 文档结构 | 多级标题、正文、有序/无序列表、表格、页眉页脚 |
| 参考匹配 | 给一份已有 `.docx`，自动提取其排版风格生成新内容 |
| 数学公式 | LaTeX 嵌入 + VBA 宏一键转为 Word 原生公式 |
| 开箱即用预设 | 课题报告、政府公文（国标）、学术论文、简历 |

---

## 快速开始

### 第一步：安装依赖

```bash
npm install -g docx
```

Python 3 为系统内置，无需额外安装。

### 第二步：安装 Skill

```bash
# 添加插件源
/plugin marketplace add <your-github-username>/word-format-skill

# 安装
/plugin install word-format@word-format-skill
```

或手动克隆：

```bash
git clone https://github.com/<your-github-username>/word-format-skill.git
```

### 第三步：直接用

在 Claude Code 里描述你的需求：

```
帮我生成一份课题报告：
- 标题：宋体三号加粗居中
- 一级标题：宋体小三加粗
- 正文：宋体小四，首行缩进2字符，1.5倍行距
- 输出到桌面

内容：
标题：基于机器学习的推荐系统研究

一、研究背景
近年来，推荐系统已成为电商、流媒体平台的核心基础设施……

二、研究方法
本文采用协同过滤与深度学习相结合的方案……
```

Claude 自动完成全部步骤，Word 文件直接弹出，格式完整。

---

## 进阶用法

### 匹配现有文档的排版风格

有一份已经排版好的 Word 文档，想让新内容和它格式一致：

```
我有参考文档：~/Documents/老报告.docx
请按照它的排版风格，帮我写一篇关于供应链优化的新报告。
```

Skill 会先读取参考文档的样式信息（字体/字号/缩进/行距），再生成风格完全匹配的新文档。

### 含数学公式的文档

```
帮我生成一份数学作业，包含以下内容：

题目1：设 $f(x) = x^3 - 3x^2$，求其极值点。

解：令 $f'(x) = 3x^2 - 6x = 0$，解得 $x = 0$ 或 $x = 2$。

因此极小值为 $$f(2) = 8 - 12 = -4$$
```

LaTeX 公式会被嵌入文档，Skill 同时输出一段 Word VBA 宏代码，在 Word 中按 `Alt+F11` 运行一次，所有公式自动转换为 Word 原生数学公式。

---

## 支持的文档预设

<details>
<summary>课题报告</summary>

- 标题：宋体三号（16pt）加粗居中
- 一级标题：宋体小三（15pt）加粗
- 二级标题：宋体四号（14pt）加粗
- 正文：宋体小四（12pt）首行缩进2字符，1.5倍行距

</details>

<details>
<summary>政府公文（国标）</summary>

- 标题：宋体二号（22pt）加粗居中
- 正文：仿宋四号（14pt）首行缩进2字符
- 页边距：上下 3.7cm，左右 2.8cm

</details>

<details>
<summary>学术论文</summary>

- 标题：宋体三号（16pt）加粗居中
- 摘要/关键词：宋体五号（10.5pt）
- 正文：宋体小四（12pt）首行缩进2字符，1.5倍行距
- 参考文献：宋体五号（10.5pt）

</details>

<details>
<summary>简历</summary>

- 姓名：黑体二号（22pt）居中
- 节标题：黑体四号（14pt）加粗
- 正文：宋体小四（12pt）

</details>

---

## 文件结构

```
word-format-skill/
├── .claude-plugin/
│   └── marketplace.json          # Claude Code 插件商店配置
├── skills/
│   └── word-format/
│       ├── SKILL.md              # Skill 核心指令
│       ├── scripts/
│       │   └── read_reference.py # 提取参考 .docx 排版样式
│       └── examples/
│           ├── academic_report.md
│           └── math_formula.md
├── README.md
└── LICENSE                       # Apache 2.0
```

---

## 致谢

本 Skill 的核心思路来自 linux.do 社区用户分享的这篇帖子：

> [**AI 快速排版 Word 文章的一个通用思路**](https://linux.do/t/topic/1217729)
>
> 作者发现 HTML 的富文本结构与 Word 高度兼容，提出了「让 AI 生成带 inline CSS 的 HTML → 浏览器中转 → 粘贴进 Word」的排版方案，以及通过 VBA 宏解决数学公式渲染问题的完整思路。

本 Skill 在此基础上进一步自动化：跳过浏览器和剪贴板环节，用 [docx](https://github.com/dolanmiu/docx)（MIT License）直接生成原生 `.docx` 文件，格式精度更高、操作步骤更少。

---

## License

Apache 2.0 — 详见 [LICENSE](./LICENSE)
