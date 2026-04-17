# 示例：课题报告

## 用户输入示例

```
帮我生成一份课题报告，要求：
- 标题：宋体三号加粗居中
- 一级标题：宋体小三加粗
- 二级标题：宋体四号加粗
- 正文：宋体小四首行缩进2字符1.5倍行距
- 输出到桌面 report.docx

内容：
标题：基于深度学习的图像识别研究

一、研究背景
近年来，深度学习技术在计算机视觉领域取得了突破性进展...

二、研究方法
（一）数据集构建
本研究使用了 ImageNet 数据集...
（二）模型设计
采用残差网络（ResNet）作为基础架构...
```

## Claude 生成的 docx-js 脚本

```javascript
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require('docx');
const fs = require('fs');

const doc = new Document({
  styles: {
    paragraphStyles: [
      {
        id: "DocTitle", name: "Doc Title", basedOn: "Normal",
        run: { size: 32, bold: true, font: "SimSun" },
        paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 240, after: 240 } }
      },
      {
        id: "H1", name: "H1", basedOn: "Normal",
        run: { size: 30, bold: true, font: "SimSun" },
        paragraph: { spacing: { before: 240, after: 120 } }
      },
      {
        id: "H2", name: "H2", basedOn: "Normal",
        run: { size: 28, bold: true, font: "SimSun" },
        paragraph: { spacing: { before: 180, after: 80 } }
      },
      {
        id: "Body", name: "Body", basedOn: "Normal",
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
      new Paragraph({ style: "DocTitle", children: [new TextRun("基于深度学习的图像识别研究")] }),
      new Paragraph({ style: "H1", children: [new TextRun("一、研究背景")] }),
      new Paragraph({ style: "Body", children: [new TextRun("近年来，深度学习技术在计算机视觉领域取得了突破性进展...")] }),
      new Paragraph({ style: "H1", children: [new TextRun("二、研究方法")] }),
      new Paragraph({ style: "H2", children: [new TextRun("（一）数据集构建")] }),
      new Paragraph({ style: "Body", children: [new TextRun("本研究使用了 ImageNet 数据集...")] }),
      new Paragraph({ style: "H2", children: [new TextRun("（二）模型设计")] }),
      new Paragraph({ style: "Body", children: [new TextRun("采用残差网络（ResNet）作为基础架构...")] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  const out = process.argv[2] || "output.docx";
  fs.writeFileSync(out, buf);
  console.log("生成成功：" + out);
});
```

## 执行命令

```bash
NODE_PATH=$(npm root -g) node /tmp/gen.js ~/Desktop/report.docx
open ~/Desktop/report.docx
```
