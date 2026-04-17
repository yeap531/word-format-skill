#!/usr/bin/env python3
"""
read_reference.py — 从现有 .docx 文件提取排版样式摘要

用法：
    python3 read_reference.py <path/to/reference.docx>

输出用于指导 docx-js 生成匹配排版的新文档。
"""

import sys
import zipfile
import re
import json
from pathlib import Path


# 字号对照（半点 → 中文名）
HALFPT_TO_CN = {
    84: "初号(42pt)", 72: "小初(36pt)", 52: "一号(26pt)", 48: "小一(24pt)",
    44: "二号(22pt)", 36: "小二(18pt)", 32: "三号(16pt)", 30: "小三(15pt)",
    28: "四号(14pt)", 24: "小四(12pt)", 21: "五号(10.5pt)", 18: "小五(9pt)",
}

# 字体中文名对照
FONT_CN = {
    "simsun": "宋体", "simhei": "黑体", "kaiti": "楷体",
    "fangsong": "仿宋", "microsoft yahei": "微软雅黑",
    "stsong": "华文宋体", "stkaiti": "华文楷体", "stfangsong": "华文仿宋",
    "arial": "Arial", "times new roman": "Times New Roman",
}


def _tag(name: str) -> str:
    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    }
    for prefix, uri in ns.items():
        if name.startswith(prefix + ":"):
            local = name[len(prefix) + 1:]
            return f"{{{uri}}}{local}"
    return name


def _find_val(element, tag: str) -> str | None:
    """在 element 子元素中找到 tag，返回 w:val 属性值。"""
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    full_tag = f"{{{w_ns}}}{tag}"
    val_attr = f"{{{w_ns}}}val"
    el = element.find(f".//{full_tag}")
    if el is not None:
        return el.get(val_attr)
    return None


def extract_styles(docx_path: str) -> dict:
    """解析 .docx 文件，返回段落样式摘要。"""
    try:
        import xml.etree.ElementTree as ET
    except ImportError:
        print("错误：需要 Python 标准库 xml.etree.ElementTree", file=sys.stderr)
        sys.exit(1)

    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    W = f"{{{w_ns}}}"

    results = {}

    with zipfile.ZipFile(docx_path, "r") as zf:
        # 读取 styles.xml
        if "word/styles.xml" not in zf.namelist():
            print("错误：未找到 word/styles.xml", file=sys.stderr)
            sys.exit(1)

        styles_xml = zf.read("word/styles.xml")
        root = ET.fromstring(styles_xml)

        for style in root.findall(f"{W}style"):
            style_id = style.get(f"{W}styleId", "")
            style_type = style.get(f"{W}type", "")
            name_el = style.find(f"{W}name")
            style_name = name_el.get(f"{W}val", "") if name_el is not None else ""

            if style_type not in ("paragraph", "character"):
                continue

            info = {"id": style_id, "name": style_name, "type": style_type}

            # 提取字体
            rPr = style.find(f".//{W}rPr")
            if rPr is not None:
                fonts_el = rPr.find(f"{W}rFonts")
                if fonts_el is not None:
                    font = (fonts_el.get(f"{W}ascii") or
                            fonts_el.get(f"{W}eastAsia") or
                            fonts_el.get(f"{W}hAnsi"))
                    if font:
                        cn_name = FONT_CN.get(font.lower(), font)
                        info["font"] = f"{font} ({cn_name})"

                sz_el = rPr.find(f"{W}sz")
                if sz_el is not None:
                    half_pt = int(sz_el.get(f"{W}val", 0))
                    pt = half_pt / 2
                    cn = HALFPT_TO_CN.get(half_pt, f"{pt}pt")
                    info["size"] = cn

                bold_el = rPr.find(f"{W}b")
                if bold_el is not None:
                    info["bold"] = True

            # 提取段落属性
            pPr = style.find(f".//{W}pPr")
            if pPr is not None:
                jc_el = pPr.find(f"{W}jc")
                if jc_el is not None:
                    info["alignment"] = jc_el.get(f"{W}val", "")

                ind_el = pPr.find(f"{W}ind")
                if ind_el is not None:
                    fl = ind_el.get(f"{W}firstLine")
                    if fl:
                        info["firstLine_dxa"] = int(fl)

                spacing_el = pPr.find(f"{W}spacing")
                if spacing_el is not None:
                    line = spacing_el.get(f"{W}line")
                    rule = spacing_el.get(f"{W}lineRule", "auto")
                    if line:
                        line_int = int(line)
                        if rule == "auto":
                            ratio = line_int / 240
                            info["lineSpacing"] = f"{ratio:.1f}倍行距 (line={line})"
                        else:
                            info["lineSpacing"] = f"固定值 {line_int/20:.1f}pt (line={line})"

            if len(info) > 3:  # 有实质内容才收录
                results[style_id] = info

    return results


def main():
    if len(sys.argv) < 2:
        print("用法：python3 read_reference.py <path/to/reference.docx>")
        sys.exit(1)

    path = sys.argv[1]
    if not Path(path).exists():
        print(f"错误：文件不存在 {path}", file=sys.stderr)
        sys.exit(1)

    styles = extract_styles(path)

    print("=" * 60)
    print(f"参考文件排版摘要：{path}")
    print("=" * 60)

    priority = ["Normal", "Title", "Heading1", "Heading2", "Heading3",
                "ListParagraph", "BodyText", "Caption"]
    shown = set()

    # 优先显示常见样式
    for sid in priority:
        if sid in styles:
            _print_style(styles[sid])
            shown.add(sid)

    # 其余样式
    for sid, info in styles.items():
        if sid not in shown:
            _print_style(info)

    print("=" * 60)
    print("请将以上样式信息提供给 Claude 以生成匹配排版的 docx-js 代码。")


def _print_style(info: dict):
    parts = [f"[{info['id']}] {info['name']}"]
    if "font" in info:
        parts.append(f"字体={info['font']}")
    if "size" in info:
        parts.append(f"字号={info['size']}")
    if info.get("bold"):
        parts.append("加粗")
    if "alignment" in info:
        parts.append(f"对齐={info['alignment']}")
    if "firstLine_dxa" in info:
        parts.append(f"首行缩进={info['firstLine_dxa']}DXA")
    if "lineSpacing" in info:
        parts.append(f"行距={info['lineSpacing']}")
    print("  " + " | ".join(parts))


if __name__ == "__main__":
    main()
