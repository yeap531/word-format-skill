# 示例：含数学公式的文档

## 用户输入示例

```
帮我生成一份数学作业，包含以下内容和公式：

题目1：求 $f(x) = x^3 - 3x^2 + 2x$ 的导数。

解：由求导法则，

$$f'(x) = 3x^2 - 6x + 2$$

当 $f'(x) = 0$ 时，解方程得 $x = 1$ 或 $x = \frac{1}{3} + \frac{\sqrt{3}}{3}$。
```

## 处理策略

公式以 LaTeX 原文嵌入文档，用户在 Word 中运行 VBA 宏一键转换。

## Claude 生成的 docx-js 脚本

```javascript
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require('docx');
const fs = require('fs');

const doc = new Document({
  styles: {
    paragraphStyles: [
      {
        id: "Body", name: "Body", basedOn: "Normal",
        run: { size: 24, font: "SimSun" },
        paragraph: { indent: { firstLine: 480 }, spacing: { line: 360, lineRule: "auto" } }
      },
      {
        id: "FormulaBlock", name: "Formula Block", basedOn: "Normal",
        run: { size: 24, font: "SimSun" },
        paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 120, after: 120 } }
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
      // 公式用原始 LaTeX，用 $ 或 $$ 包裹
      new Paragraph({ style: "Body", children: [
        new TextRun("题目1：求 $f(x) = x^3 - 3x^2 + 2x$ 的导数。"),
      ]}),
      new Paragraph({ style: "Body", children: [new TextRun("解：由求导法则，")] }),
      // 行间公式居中显示
      new Paragraph({ style: "FormulaBlock", children: [
        new TextRun("$$f'(x) = 3x^2 - 6x + 2$$"),
      ]}),
      new Paragraph({ style: "Body", children: [
        new TextRun("当 $f'(x) = 0$ 时，解方程得 $x = 1$ 或 $x = \\frac{1}{3} + \\frac{\\sqrt{3}}{3}$。"),
      ]}),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  const out = process.argv[2] || "output.docx";
  fs.writeFileSync(out, buf);
  console.log("已生成：" + out);
  console.log("请在 Word 中按 Alt+F11 运行 VBA 宏转换公式。");
});
```

## 执行命令

```bash
NODE_PATH=$(npm root -g) node /tmp/gen.js ~/Desktop/math_homework.docx
open ~/Desktop/math_homework.docx
```

## VBA 宏（在 Word 中运行一次）

见 SKILL.md 中「数学公式处理」章节的完整宏代码。
