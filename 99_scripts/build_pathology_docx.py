#!/usr/bin/env python3
from __future__ import annotations

from html import escape
from pathlib import Path
import re
import shutil
import zipfile


ROOT = Path("/Users/daniellelan/Desktop/病理期末理论准备")
VAULT = ROOT / "病理学理论复习"
TEMPLATE = Path("/Users/daniellelan/Desktop/微生物最终复习资料_完整版 (1).docx")
OUT = ROOT / "病理学最终复习资料_完整版.docx"

FONT = "Noto Sans CJK SC"
PAGE_WIDTH_DXA = 10540


def clean_text(text: str) -> str:
    text = re.sub(r"\[\[([^|\]#]+)(?:#[^|\]]+)?\|([^\]]+)\]\]", r"\2", text)
    text = re.sub(r"\[\[([^|\]#]+)(?:#[^\]]+)?\]\]", r"\1", text)
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
    text = text.replace("<br>", "；").replace("<br/>", "；")
    text = text.replace("`", "")
    text = text.replace("&", "＆")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def rpr(size: int = 18, bold: bool = False, color: str | None = None) -> str:
    b = "<w:b/>" if bold else ""
    c = f'<w:color w:val="{color}"/>' if color else ""
    return (
        f"<w:rPr><w:rFonts w:ascii=\"{FONT}\" w:hAnsi=\"{FONT}\" "
        f"w:eastAsia=\"{FONT}\"/>{b}{c}<w:sz w:val=\"{size}\"/></w:rPr>"
    )


def run_xml(text: str, size: int = 18, bold: bool = False, color: str | None = None) -> str:
    text = clean_text(text)
    parts = re.split(r"(\*\*[^*]+\*\*)", text)
    out = []
    for part in parts:
        if not part:
            continue
        is_bold = bold
        if part.startswith("**") and part.endswith("**"):
            part = part[2:-2]
            is_bold = True
        out.append(f"<w:r>{rpr(size, is_bold, color)}<w:t>{escape(part)}</w:t></w:r>")
    return "".join(out)


def paragraph(
    text: str,
    style: str | None = None,
    size: int = 18,
    bold: bool = False,
    color: str | None = None,
    align: str | None = None,
    before: int | None = None,
    after: int | None = 80,
    num_id: int | None = None,
    keep_next: bool = False,
) -> str:
    ppr = []
    if style:
        ppr.append(f'<w:pStyle w:val="{style}"/>')
    if num_id is not None:
        ppr.append(f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>')
    if align:
        ppr.append(f'<w:jc w:val="{align}"/>')
    spacing = []
    if before is not None:
        spacing.append(f'w:before="{before}"')
    if after is not None:
        spacing.append(f'w:after="{after}"')
    if spacing:
        ppr.append(f'<w:spacing {" ".join(spacing)} w:line="260" w:lineRule="auto"/>')
    if keep_next:
        ppr.append("<w:keepNext/>")
    ppr_xml = f"<w:pPr>{''.join(ppr)}</w:pPr>" if ppr else ""
    return f"<w:p>{ppr_xml}{run_xml(text, size=size, bold=bold, color=color)}</w:p>"


def title_para(text: str) -> str:
    return paragraph(text, style="36", size=34, bold=True, color="1F4E79", align="center", after=120)


def subtitle_para(text: str) -> str:
    return paragraph(text, style="34", size=21, color="666666", align="center", after=240)


def heading(text: str, level: int) -> str:
    if level <= 1:
        return paragraph(text, style="2", size=26, bold=True, color="1F4E79", before=260, after=120, keep_next=True)
    if level == 2:
        return paragraph(text, style="3", size=22, bold=True, color="5B9BD5", before=180, after=80, keep_next=True)
    return paragraph(text, style="4", size=19, bold=True, color="333333", before=120, after=60, keep_next=True)


def column_widths(n: int) -> list[int]:
    if n == 1:
        return [PAGE_WIDTH_DXA]
    if n == 2:
        return [2600, PAGE_WIDTH_DXA - 2600]
    if n == 3:
        return [1900, 4450, PAGE_WIDTH_DXA - 6350]
    if n == 4:
        return [1700, 2950, 2950, PAGE_WIDTH_DXA - 7600]
    base = PAGE_WIDTH_DXA // n
    widths = [base] * n
    widths[-1] += PAGE_WIDTH_DXA - sum(widths)
    return widths


def cell(text: str, width: int, header: bool = False) -> str:
    fill = '<w:shd w:val="clear" w:color="auto" w:fill="D9EAF7"/>' if header else ""
    return (
        "<w:tc>"
        f'<w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{fill}<w:vAlign w:val="center"/></w:tcPr>'
        f"{paragraph(text, size=17, bold=header, after=0)}"
        "</w:tc>"
    )


def table(rows: list[list[str]], header: bool = True, fill: str = "D9EAF7") -> str:
    if not rows:
        return ""
    n = max(len(r) for r in rows)
    widths = column_widths(n)
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    tbl = [
        "<w:tbl>",
        '<w:tblPr><w:tblStyle w:val="35"/><w:tblW w:w="10540" w:type="dxa"/>'
        '<w:jc w:val="center"/><w:tblBorders>'
        '<w:top w:val="single" w:color="B7D7F4" w:sz="4" w:space="0"/>'
        '<w:left w:val="single" w:color="B7D7F4" w:sz="4" w:space="0"/>'
        '<w:bottom w:val="single" w:color="B7D7F4" w:sz="4" w:space="0"/>'
        '<w:right w:val="single" w:color="B7D7F4" w:sz="4" w:space="0"/>'
        '<w:insideH w:val="single" w:color="B7D7F4" w:sz="4" w:space="0"/>'
        '<w:insideV w:val="single" w:color="B7D7F4" w:sz="4" w:space="0"/>'
        '</w:tblBorders><w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
        '<w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr>',
        f"<w:tblGrid>{grid}</w:tblGrid>",
    ]
    for i, row in enumerate(rows):
        row = row + [""] * (n - len(row))
        tbl.append("<w:tr>")
        for j, val in enumerate(row):
            is_header = header and i == 0
            tbl.append(cell(val, widths[j], header=is_header))
        tbl.append("</w:tr>")
    tbl.append("</w:tbl>")
    return "".join(tbl)


def callout(text: str, fill: str = "E2F0D9") -> str:
    return (
        "<w:tbl><w:tblPr><w:tblStyle w:val=\"12\"/><w:tblW w:w=\"10540\" w:type=\"dxa\"/>"
        "<w:jc w:val=\"center\"/><w:tblBorders>"
        "<w:top w:val=\"single\" w:color=\"B7D7F4\" w:sz=\"4\" w:space=\"0\"/>"
        "<w:left w:val=\"single\" w:color=\"B7D7F4\" w:sz=\"4\" w:space=\"0\"/>"
        "<w:bottom w:val=\"single\" w:color=\"B7D7F4\" w:sz=\"4\" w:space=\"0\"/>"
        "<w:right w:val=\"single\" w:color=\"B7D7F4\" w:sz=\"4\" w:space=\"0\"/>"
        "</w:tblBorders><w:tblCellMar><w:top w:w=\"120\" w:type=\"dxa\"/><w:left w:w=\"160\" w:type=\"dxa\"/>"
        "<w:bottom w:w=\"120\" w:type=\"dxa\"/><w:right w:w=\"160\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr>"
        "<w:tblGrid><w:gridCol w:w=\"10540\"/></w:tblGrid><w:tr><w:tc>"
        f"<w:tcPr><w:tcW w:w=\"10540\" w:type=\"dxa\"/><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{fill}\"/></w:tcPr>"
        f"{paragraph(text, size=19, bold=True, after=0)}"
        "</w:tc></w:tr></w:tbl>"
    )


def parse_md(md: str) -> list[str]:
    blocks: list[str] = []
    lines = md.splitlines()
    i = 0
    in_code = False
    code_lines: list[str] = []
    while i < len(lines):
        raw = lines[i].rstrip()
        line = raw.strip()
        if line.startswith("```"):
            if in_code:
                if code_lines:
                    blocks.append(callout("；".join(clean_text(x) for x in code_lines if x.strip()), fill="F2F2F2"))
                code_lines = []
                in_code = False
            else:
                in_code = True
            i += 1
            continue
        if in_code:
            code_lines.append(raw)
            i += 1
            continue
        if not line or line == "---":
            i += 1
            continue
        if line.startswith(">"):
            quote_lines = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                q = lines[i].strip().lstrip(">").strip()
                if q and not q.startswith("[!"):
                    quote_lines.append(q)
                i += 1
            if quote_lines:
                blocks.append(callout(" ".join(quote_lines), fill="EAF4FF"))
            continue
        if line.startswith("|") and i + 1 < len(lines) and re.match(r"^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$", lines[i + 1]):
            rows = []
            rows.append([clean_text(c) for c in line.strip("|").split("|")])
            i += 2
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append([clean_text(c) for c in lines[i].strip().strip("|").split("|")])
                i += 1
            blocks.append(table(rows))
            continue
        m = re.match(r"^(#{1,6})\s+(.*)$", line)
        if m:
            hashes, text = m.groups()
            text = clean_text(text)
            if text and not text.startswith("病理学理论"):
                level = len(hashes)
                # Keep the final document compact: original H2 becomes major subsection.
                blocks.append(heading(text, max(1, level - 1)))
            i += 1
            continue
        m = re.match(r"^[-*]\s+\[.\]\s+(.*)$", line)
        if m:
            blocks.append(paragraph(clean_text(m.group(1)), size=18, num_id=2, after=40))
            i += 1
            continue
        m = re.match(r"^[-*]\s+(.*)$", line)
        if m:
            blocks.append(paragraph(clean_text(m.group(1)), size=18, num_id=2, after=40))
            i += 1
            continue
        m = re.match(r"^\d+\.\s+(.*)$", line)
        if m:
            blocks.append(paragraph(clean_text(m.group(1)), size=18, num_id=6, after=40))
            i += 1
            continue
        blocks.append(paragraph(clean_text(line), size=18, after=70))
        i += 1
    return blocks


def slice_file(path: Path, start: str | None = None, end: str | None = None) -> str:
    text = path.read_text(encoding="utf-8")
    if start and start in text:
        text = text[text.index(start):]
    if end and end in text:
        text = text[:text.index(end)]
    return text


def build_markdown() -> str:
    intro = """
背诵总策略：A档完整背“定义、形态、机制、临床联系”；B档会一眼认和关键鉴别；C档只背题眼。病理理论卷主观题分值高，先把名解、简答和病例模板背熟，再用选择题查漏。

| 等级 | 范围 | 要求 |
|---|---|---|
| A档：完整背 | 细胞损伤、修复、血栓栓塞梗死、炎症、肿瘤总论、心血管、呼吸、消化肝胆、传染病 | 定义、肉眼、镜下、机制、临床病理联系都要会写 |
| B档：会认会鉴别 | 泌尿、生殖乳腺、淋巴造血、内分泌、神经、骨关节、寄生虫 | 题眼 + 关键形态 + 常见鉴别 |
| C档：题眼 | 少见肿瘤亚型、很细的分子分型、低频系统边角 | 能把关键词和疾病对应起来 |

"""
    total = slice_file(VAULT / "01_高频病变核心速查.md", "## 一、总论必考病变")
    systems = slice_file(VAULT / "03_系统疾病速查.md", "## 一、心血管系统")
    subj = slice_file(VAULT / "02_机制对比名解.md", "## 名词解释高频库")
    cases = slice_file(VAULT / "04_题型病例真题.md", "## 答题模板")
    return "\n".join([intro, "\n# 一、总论核心\n", total, "\n# 二、系统病理\n", systems, "\n# 三、高频名解与简答题\n", subj, "\n# 四、题型病例与真题反查\n", cases])


def document_xml(blocks: list[str], sect_pr: str) -> str:
    ns = (
        '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
        'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
        'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData" '
        'mc:Ignorable="w14 wp14">'
    )
    body = "".join(blocks) + sect_pr
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ns + "<w:body>" + body + "</w:body></w:document>"


def footer_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:ftr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
        'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
        'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData" mc:Ignorable="w14 w15 wp14">'
        '<w:p><w:pPr><w:pStyle w:val="18"/><w:jc w:val="center"/></w:pPr>'
        f'<w:r>{rpr(16, False, "646464")}<w:t>病理学最终复习资料（完整版）</w:t></w:r></w:p></w:ftr>'
    )


def main() -> None:
    md = build_markdown()
    blocks = [title_para("病理学最终复习资料"), subtitle_para("完整版：总论 + 系统病理 + 高频名解 + 简答病例")]
    blocks.append(callout("背诵核心：选择题看题眼，名解写三句，简答写表格，病例写诊断依据。两天内先背A档，再用病例反查补漏洞。"))
    blocks.extend(parse_md(md))

    with zipfile.ZipFile(TEMPLATE) as zin:
        original_doc = zin.read("word/document.xml").decode("utf-8")
        m = re.search(r"<w:sectPr[\s\S]*?</w:sectPr>", original_doc)
        sect_pr = m.group(0) if m else '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="850" w:right="850" w:bottom="794" w:left="850" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
        temp = OUT.with_suffix(".tmp.docx")
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = document_xml(blocks, sect_pr).encode("utf-8")
                elif item.filename == "word/footer1.xml":
                    data = footer_xml().encode("utf-8")
                elif item.filename == "docProps/core.xml":
                    text = data.decode("utf-8", errors="ignore")
                    text = re.sub(r"<dc:title>.*?</dc:title>", "<dc:title>病理学最终复习资料_完整版</dc:title>", text)
                    data = text.encode("utf-8")
                zout.writestr(item, data)
    shutil.move(temp, OUT)
    print(OUT)


if __name__ == "__main__":
    main()
