#!/usr/bin/env python3
from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
import re
import shutil
import sys
import zipfile


ROOT = Path("/Users/daniellelan/Desktop/病理期末理论准备")
SCRIPT_DIR = ROOT / "99_scripts"
sys.path.insert(0, str(SCRIPT_DIR))

import build_pathology_docx as base  # noqa: E402


VAULT = ROOT / "病理学理论复习"
OUT = ROOT / "病理学期末复习提纲_自用整理.docx"
TITLE = "病理学期末复习提纲"
SUBTITLE = "自用整理：按教学大纲、教材目录和课堂复习顺序归纳"


REPLACEMENTS = [
    ("高频名解与简答题", "核心概念与问答整理"),
    ("名词解释高频库", "核心概念整理"),
    ("名词解释模板", "概念表述框架"),
    ("名词解释", "概念表述"),
    ("| 名词 |", "| 概念 |"),
    ("三句式答案", "整理表达"),
    ("最短答案", "整理要点"),
    ("直接答案", "直接对应"),
    ("精选答案骨架", "整理示例"),
    ("答案骨架", "整理示例"),
    ("答案", "表述"),
    ("简答题专项", "问答整理"),
    ("简答题模板", "问答表述框架"),
    ("简答题", "问答整理"),
    ("简答", "问答"),
    ("选择题易混点速查", "易混概念对照"),
    ("选择题快反", "易混概念对照"),
    ("选择题", "概念辨析"),
    ("题型 · 病例 · 真题反查", "临床病理联系整理"),
    ("题型病例与真题反查", "临床病理联系与应用思路"),
    ("病例题精选答案骨架", "临床病理联系示例"),
    ("病例题模板", "临床病理联系表述框架"),
    ("病例题", "临床病理联系"),
    ("病例反查总表", "临床病理线索总表"),
    ("病例反查", "临床病理线索"),
    ("病例写诊断依据", "临床病理联系写依据"),
    ("病例模板", "临床病理联系框架"),
    ("病例", "临床病理联系"),
    ("真题名解回收清单", "核心概念补充清单"),
    ("真题简答回收清单", "问答整理补充清单"),
    ("真题", "复习资料"),
    ("往年题高频主题", "复习主题整理"),
    ("往年题", "复习资料"),
    ("考场病例题小抄", "临床病理联系提示"),
    ("考场", "复习"),
    ("小抄", "提示"),
    ("高频答案", "整理要点"),
    ("高频特点", "主要特点"),
    ("高频点", "复习要点"),
    ("高频主题", "重点主题"),
    ("高频", "重点"),
    ("必考病变", "基础病变"),
    ("必考", "基础"),
    ("必背点", "需要掌握的要点"),
    ("必背", "需要掌握"),
    ("背诵核心", "整理思路"),
    ("背诵总策略", "复习整理思路"),
    ("背诵", "复习"),
    ("题眼", "识别线索"),
    ("答题模板", "表述框架"),
    ("答题顺序", "整理顺序"),
    ("答题要点", "整理要点"),
    ("答题", "表述"),
    ("常考一句话", "复习提示"),
    ("考点", "要点"),
    ("主观题", "书面表达"),
    ("客观题", "辨析题"),
    ("两天内", "复习时"),
    ("2天内", "复习时"),
    ("A 级：必须会写", "基础概念：需要准确表述"),
    ("B 级：常考但可压缩", "补充概念：理解后能概括"),
    ("A档：完整背", "基础框架"),
    ("B档：会认会鉴别", "系统补充"),
    ("C档：题眼", "查漏内容"),
    ("A档", "基础框架"),
    ("B档", "系统补充"),
    ("C档", "查漏内容"),
    ("完整背", "完整掌握"),
    ("会认会鉴别", "会识别和鉴别"),
    ("题眼", "识别线索"),
    ("题干", "具体材料"),
    ("必写依据/补充", "主要依据/补充"),
    ("死因可写", "死因可考虑"),
    ("押题", "复习"),
    ("OCR", "资料整理"),
]


def soften(text: str) -> str:
    for old, new in REPLACEMENTS:
        text = text.replace(old, new)
    text = re.sub(r"优先复习这些，因为.*?：", "可先整理这些概念：", text)
    text = text.replace("病理理论卷", "病理学理论复习")
    return text


def slice_file(path: Path, start: str | None = None, end: str | None = None) -> str:
    text = path.read_text(encoding="utf-8")
    if start and start in text:
        text = text[text.index(start):]
    if end and end in text:
        text = text[: text.index(end)]
    return text


def drop_leading_heading(text: str) -> str:
    lines = text.splitlines()
    while lines and not lines[0].strip():
        lines.pop(0)
    if lines and re.match(r"^#{1,6}\s+", lines[0].strip()):
        lines = lines[1:]
    return "\n".join(lines).lstrip()


def build_markdown() -> str:
    intro = """
整理说明：这份提纲按“总论概念 → 系统病理 → 临床病理联系”的顺序整理，用来复习时查漏补缺。复习时先把定义、形态学改变和机制连起来，再回到各系统疾病中看临床病理联系。

| 层次 | 内容范围 | 复习目标 |
|---|---|---|
| 基础框架 | 细胞损伤与修复、局部血液循环障碍、炎症、肿瘤总论 | 能准确说清定义、基本形态、发生机制和临床意义 |
| 系统整合 | 心血管、呼吸、消化肝胆、泌尿、生殖乳腺、淋巴造血、内分泌、神经、骨关节、传染病 | 按病因/机制、大体、镜下、临床病理联系四条线整理 |
| 查漏补缺 | 容易混淆的概念、疾病之间的鉴别、少见但教材有强调的病变 | 能根据关键词回到对应章节，不把相近疾病混在一起 |

"""
    total = drop_leading_heading(slice_file(VAULT / "01_高频病变核心速查.md", "## 一、总论必考病变"))
    systems = slice_file(VAULT / "03_系统疾病速查.md", "## 一、心血管系统")
    concepts = slice_file(VAULT / "02_机制对比名解.md", "## 名词解释高频库")
    case_templates = slice_file(VAULT / "04_题型病例真题.md", "## 答题模板", "## 往年题高频主题")
    case_links = slice_file(VAULT / "04_题型病例真题.md", "## 病例反查总表", "## 真题名解回收清单")

    md = "\n".join(
        [
            intro,
            "\n# 一、总论基础\n",
            total,
            "\n# 二、系统病理整理\n",
            systems,
            "\n# 三、核心概念与问答整理\n",
            concepts,
            "\n# 四、临床病理联系与应用思路\n",
            case_templates,
            case_links,
        ]
    )
    return soften(md)


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
        f'<w:r>{base.rpr(16, False, "646464")}<w:t>{TITLE}（自用整理）</w:t></w:r></w:p></w:ftr>'
    )


def core_xml() -> bytes:
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        f"<dc:title>{TITLE}</dc:title>"
        "<dc:creator>Lan Danielle</dc:creator>"
        "<cp:lastModifiedBy>Lan Danielle</cp:lastModifiedBy>"
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
        "<cp:revision>1</cp:revision>"
        "</cp:coreProperties>"
    ).encode("utf-8")


def app_xml(data: bytes) -> bytes:
    text = data.decode("utf-8", errors="ignore")
    text = re.sub(r"<Application>.*?</Application>", "<Application>Microsoft Word</Application>", text)
    text = re.sub(r"<Pages>.*?</Pages>", "<Pages>1</Pages>", text)
    return text.encode("utf-8")


def main() -> None:
    md = build_markdown()
    blocks = [
        base.title_para(TITLE),
        base.subtitle_para(SUBTITLE),
        base.callout("整理思路：以教学大纲和教材章节为主线，重点把“病变名称—病因机制—大体/镜下形态—临床病理联系”串起来，方便复习时查漏补缺。"),
    ]
    blocks.extend(base.parse_md(md))

    with zipfile.ZipFile(base.TEMPLATE) as zin:
        original_doc = zin.read("word/document.xml").decode("utf-8")
        m = re.search(r"<w:sectPr[\s\S]*?</w:sectPr>", original_doc)
        sect_pr = (
            m.group(0)
            if m
            else '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="850" w:right="850" w:bottom="794" w:left="850" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
        )
        temp = OUT.with_suffix(".tmp.docx")
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zout:
            seen_core = False
            seen_app = False
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = base.document_xml(blocks, sect_pr).encode("utf-8")
                elif item.filename == "word/footer1.xml":
                    data = footer_xml().encode("utf-8")
                elif item.filename == "docProps/core.xml":
                    data = core_xml()
                    seen_core = True
                elif item.filename == "docProps/app.xml":
                    data = app_xml(data)
                    seen_app = True
                zout.writestr(item, data)
            if not seen_core:
                zout.writestr("docProps/core.xml", core_xml())
            if not seen_app:
                zout.writestr(
                    "docProps/app.xml",
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Word</Application></Properties>'.encode(
                        "utf-8"
                    ),
                )
    shutil.move(temp, OUT)
    print(OUT)


if __name__ == "__main__":
    main()
