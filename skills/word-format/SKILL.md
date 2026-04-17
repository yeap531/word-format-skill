---
name: word-format
description: "Use this skill whenever the user wants to create a professionally formatted Word (.docx) document with Chinese typography standards — including 字体/字号/缩进/行距/表格 specifications. Triggers: mentions of 宋体/黑体/楷体, 字号 (小四/三号/etc.), 首行缩进, 课题报告/政府公文/学术论文, or any request to produce a .docx with specific formatting. Also use when the user provides a reference Word file and wants to generate matching-style content, or when content contains LaTeX math formulas that need to be pasted into Word. Do NOT use for PDF, PowerPoint, Excel, or plain Markdown output."
---

# Word Format Skill — 中文排版 Word 文档生成

## 概述

本 skill 负责将用户的内容和排版要求直接生成格式完整的 `.docx` 文件，全程无需浏览器或手动粘贴。

**技术路径：**
```
用户需求 → 分析排版规格 → 生成 docx-js 脚本 → Node.js 执行 → 输出 .docx → open 打开
```

---

## 工作流决策树

```
用户提供了参考 .docx 文件？
├─ 是 → 运行 scripts/read_reference.py 提取样式 → 匹配其排版生成新内容
└─ 否 → 用户描述了排版规格？
         ├─ 是 → 按规格生成
         └─ 否 → 按文档类型使用预设（见下方预设模板）
```

---

## 第一步：收集信息

向用户确认以下内容（未提供时询问）：

1. **文档类型**：课题报告 / 政府公文 / 学术论文 / 简历 / 其他
2. **排版规格**：字体、字号、缩进、行距（可参考下方中文字号表）
3. **是否有参考文件**：现有 .docx 文件路径（用于匹配已有排版）
4. **是否含数学公式**：含 LaTeX 公式时需附加 VBA 宏说明
5. **输出路径**：默认输出到 `~/Desktop/output.docx`

---

## 第二步：读取参考文件（如有）

```bash
python3 skills/word-format/scripts/read_reference.py <path/to/reference.docx>
```

脚本会提取：字体名称、字号、段落样式、缩进、行距，打印为结构化摘要供后续生成使用。

---

## 第三步：生成 docx-js 脚本

### 安装依赖（首次使用）

```bash
npm install -g docx
```

### 中文字号对照表（docx-js size 单位为半点：1pt = 2）

| 中文字号 | pt 值 | docx-js size |
|---------|-------|-------------|
| 初号    | 42pt  | 84  |
| 小初    | 36pt  | 72  |
| 一号    | 26pt  | 52  |
| 小一    | 24pt  | 48  |
| 二号    | 22pt  | 44  |
| 小二    | 18pt  | 36  |
| 三号    | 16pt  | 32  |
| 小三    | 15pt  | 30  |
| 四号    | 14pt  | 28  |
| 小四    | 12pt  | 24  |
| 五号    | 10.5pt| 21  |
| 小五    | 9pt   | 18  |

### 中文字体对照表

| 中文名称   | docx-js font 值         |
|-----------|------------------------|
| 宋体       | `"SimSun"`             |
| 黑体       | `"SimHei"`             |
| 楷体       | `"KaiTi"`              |
| 仿宋       | `"FangSong"`           |
| 微软雅黑   | `"Microsoft YaHei"`    |
| 华文宋体   | `"STSong"`             |
| 华文楷体   | `"STKaiti"`            |
| 华文仿宋   | `"STFangsong"`         |

### 首行缩进计算（首行缩进2字符 = 2 × pt × 20 DXA）

| 字号  | 首行缩进2字符 (DXA) |
|-------|-----------------|
| 小四 (12pt) | 480 |
| 五号 (10.5pt) | 420 |
| 四号 (14pt) | 560 |

### 行距设置

| 行距类型     | docx-js spacing            |
|------------|---------------------------|
| 单倍行距     | `{ line: 240, lineRule: "auto" }` |
| 1.5 倍行距   | `{ line: 360, lineRule: "auto" }` |
| 2 倍行距     | `{ line: 480, lineRule: "auto" }` |
| 固定值 xx 磅 | `{ line: xx * 20, lineRule: "exact" }` |

### A4 页面设置（中文标准）

```javascript
properties: {
  page: {
    size: { width: 11906, height: 16838 },  // A4: 11906 × 16838 DXA
    margin: {
      top: 1440,     // 2.54cm
      bottom: 1440,  // 2.54cm
      left: 1800,    // 3.17cm（政府公文标准）
      right: 1800,
    }
  }
}
```

---

## docx-js 脚本模板

运行方式：
```bash
NODE_PATH=$(npm root -g) node output_script.js
```

### 标准课题报告模板

```javascript
const { Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat } = require('docx');
const fs = require('fs');

const doc = new Document({
  styles: {
    paragraphStyles: [
      // 标题：宋体三号（16pt）加粗居中
      {
        id: "DocTitle", name: "Doc Title", basedOn: "Normal",
        run: { size: 32, bold: true, font: "SimSun" },
        paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 240, after: 240 } }
      },
      // 一级标题：宋体小三（15pt）加粗
      {
        id: "Heading1CN", name: "Heading 1 CN", basedOn: "Normal",
        run: { size: 30, bold: true, font: "SimSun" },
        paragraph: { spacing: { before: 240, after: 120 } }
      },
      // 二级标题：宋体四号（14pt）加粗
      {
        id: "Heading2CN", name: "Heading 2 CN", basedOn: "Normal",
        run: { size: 28, bold: true, font: "SimSun" },
        paragraph: { spacing: { before: 180, after: 80 } }
      },
      // 正文：宋体小四（12pt）首行缩进2字符，1.5倍行距
      {
        id: "BodyCN", name: "Body CN", basedOn: "Normal",
        run: { size: 24, font: "SimSun" },
        paragraph: {
          indent: { firstLine: 480 },
          spacing: { line: 360, lineRule: "auto" }
        }
      },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, bottom: 1440, left: 1800, right: 1800 }
      }
    },
    children: [
      new Paragraph({ style: "DocTitle", children: [new TextRun("课题报告标题")] }),
      new Paragraph({ style: "Heading1CN", children: [new TextRun("一、引言")] }),
      new Paragraph({ style: "BodyCN", children: [new TextRun("正文内容在此处填写。")] }),
      new Paragraph({ style: "Heading1CN", children: [new TextRun("二、研究方法")] }),
      new Paragraph({ style: "BodyCN", children: [new TextRun("第二段正文在此处填写。")] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(process.argv[2] || "output.docx", buf);
  console.log("生成成功：" + (process.argv[2] || "output.docx"));
});
```

---

## 第四步：执行并打开

```bash
NODE_PATH=$(npm root -g) node /tmp/word_format_gen.js ~/Desktop/output.docx
open ~/Desktop/output.docx
```

---

## 数学公式处理（含 LaTeX）

当文档包含数学公式时：

### 生成策略
在 docx-js 脚本中，将所有公式保留为 **未渲染的原始 LaTeX** 格式：
- 行内公式用 `$...$` 包裹
- 行间公式用 `$$...$$` 包裹

例如：
```javascript
new Paragraph({ style: "BodyCN", children: [
  new TextRun("设函数 $f(x) = x^2 + 2x + 1$，则其导数为 $f'(x) = 2x + 2$。"),
]})
```

### VBA 宏（用户在 Word 中运行一次）

生成文档后，告知用户在 Word 中执行以下操作：

1. 按 `Alt + F11` 打开 VBA 编辑器
2. 选择「插入」→「模块」
3. 粘贴以下代码，按 F5 运行：

```vba
Sub LatexToWordMath()
    Dim rng As Range
    Dim mathRng As Range
    Application.ScreenUpdating = False

    ' 处理双美元符号 $$ （行间公式）
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Text = "\$\$*\$\$"
        .MatchWildcards = True
        Do While .Execute
            Set mathRng = rng.Duplicate
            mathRng.End = mathRng.End
            mathRng.MoveEnd Unit:=wdCharacter, Count:=-2
            mathRng.Start = mathRng.Start + 2
            rng.Text = mathRng.Text
            ActiveDocument.OMaths.Add rng
            rng.OMaths(1).BuildUp
            rng.Collapse wdCollapseEnd
        Loop
    End With

    ' 处理单美元符号 $ （行内公式）
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Text = "\$*\$"
        .MatchWildcards = True
        Do While .Execute
            Set mathRng = rng.Duplicate
            If Len(mathRng.Text) > 2 Then
                Dim cleanText As String
                cleanText = Mid(mathRng.Text, 2, Len(mathRng.Text) - 2)
                rng.Text = cleanText
                ActiveDocument.OMaths.Add rng
                rng.OMaths(1).BuildUp
            End If
            rng.Collapse wdCollapseEnd
        Loop
    End With

    Application.ScreenUpdating = True
    MsgBox "公式转换完成！"
End Sub
```

---

## 预设模板：常见文档类型

### 政府公文（国标）
- 标题：宋体二号（22pt）加粗居中
- 正文：仿宋四号（14pt）首行缩进2字符
- 页边距：上下各 3.7cm，左右各 2.8cm

### 学术论文
- 标题：宋体三号（16pt）加粗居中
- 摘要/关键词：宋体五号（10.5pt）
- 正文：宋体小四（12pt）首行缩进2字符，1.5倍行距
- 参考文献：宋体五号（10.5pt）

### 课题报告
- 标题：宋体三号（16pt）加粗居中
- 一级标题：宋体小三（15pt）加粗
- 二级标题：宋体四号（14pt）加粗
- 正文：宋体小四（12pt）首行缩进2字符，1.5倍行距

### 简历
- 姓名：黑体二号（22pt）居中
- 节标题：黑体四号（14pt）加粗
- 正文：宋体小四（12pt）

---

## 使用注意

以下内容来自 [docx 官方 demo](https://github.com/dolanmiu/docx/tree/master/demo) 和 OOXML 规范的实测总结：

**段落与换行**
OOXML 的 `<w:p>` 是段落的基本单位，`<w:t>` 内的 `\n` 不会产生换行效果。
每个独立段落必须用单独的 `new Paragraph({...})` 表示（见 [demo/1-basic.ts](https://github.com/dolanmiu/docx/blob/master/demo/1-basic.ts)）。

**列表**
docx 的列表通过 `numbering` config 实现（见 [demo/8-header-footer.ts](https://github.com/dolanmiu/docx/blob/master/demo/8-header-footer.ts) 中的 `LevelFormat.DECIMAL` 用法）。
直接在 `TextRun` 里写 `•` 或 `※` 不会生成真正的 Word 列表结构，无法被 Word 识别为有序/无序列表。

**分页**
OOXML 规范要求 `<w:br w:type="page"/>` 必须在 `<w:r>` 内，而 `<w:r>` 必须在 `<w:p>` 内。
docx-js 对应写法：`new Paragraph({ children: [new PageBreak()] })`，不能单独使用 `new PageBreak()`。

**图片**
`ImageRun` 的 `type` 字段是 docx-js TypeScript 类型定义中的必填项（见 [demo/5-images.ts](https://github.com/dolanmiu/docx/blob/master/demo/5-images.ts)），需明确指定 `"png"` / `"jpg"` / `"jpeg"` 等。

**表格宽度**
官方 demo [4-basic-table.ts](https://github.com/dolanmiu/docx/blob/master/demo/4-basic-table.ts) 在 `Table` 上设置 `columnWidths`，同时在每个 `TableCell` 上设置 `width`。
单位统一使用 `WidthType.DXA`（1440 DXA = 1 英寸）。

**执行方式（macOS/Linux 全局安装场景）**
全局安装的 npm 包默认不在 `NODE_PATH` 内，需显式指定：
```bash
NODE_PATH=$(npm root -g) node your_script.js
```

---

## 依赖

| 工具 | 安装方式 | 用途 |
|------|---------|------|
| Node.js | [nodejs.org](https://nodejs.org) | 运行 docx-js |
| docx (npm) | `npm install -g docx` | Word 文档生成 |
| Python 3 | 系统内置 | 读取参考 .docx 样式 |
