#!/usr/bin/env python3
from __future__ import annotations

from pathlib import Path
import re
import shutil
import sys
import zipfile


ROOT = Path("/Users/daniellelan/Desktop/病理期末理论准备")
SCRIPT_DIR = ROOT / "99_scripts"
sys.path.insert(0, str(SCRIPT_DIR))

import build_pathology_docx as base  # noqa: E402
import build_pathology_teacher_docx as teacher  # noqa: E402


OUT = ROOT / "病理学期末复习提纲_自用整理_一句话故事法.docx"
SUBTITLE = "自用整理：表格梳理为主，易混病变加入一句话故事法"


PNEUMONIA_STORY = """
复习方法：肺炎不要先硬背名字，先把病变范围和显微画面连起来，再看下面表格。
- 大叶性肺炎：一整个肺叶像肝一样实变，所以有红肝样、灰肝样。
- 小叶性肺炎：以细支气管为中心，一个个小灶化脓。
- 病毒性肺炎：不在肺泡里堆脓，而是肺泡间隔变厚，淋巴单核浸润。
"""


INTESTINE_STORY = """
复习方法：肠道感染先记“部位 + 溃疡方向/形状 + 典型表现”，再回到表格看机制和鉴别。
- 肠结核：回盲部，环形溃疡，长轴和肠管垂直。
- 伤寒：回肠末端，溃疡长轴和肠管平行。
- 阿米巴：烧瓶状溃疡，口小底大。
- 菌痢：直肠乙状结肠，假膜性炎，黏液脓血便。
"""


def add_story_blocks(md: str) -> str:
    md = md.replace("### 肺炎三类\n\n", "### 肺炎三类\n\n" + PNEUMONIA_STORY + "\n")
    md = md.replace("### 伤寒、菌痢、梅毒、麻风、真菌\n\n", "### 肠道感染/溃疡一句话故事法\n\n" + INTESTINE_STORY + "\n### 伤寒、菌痢、梅毒、麻风、真菌\n\n")
    md = md.replace("### 15. 肠结核、伤寒、菌痢、阿米巴病鉴别\n\n", "### 15. 肠结核、伤寒、菌痢、阿米巴病鉴别\n\n" + INTESTINE_STORY + "\n")
    return md


def blank_footer_xml() -> bytes:
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
        'xmlns:wpsCustomData="http://www.wps.cn/officeDocument/2013/wpsCustomData" '
        'mc:Ignorable="w14 w15 wp14"><w:p/></w:ftr>'
    ).encode("utf-8")


def main() -> None:
    md = add_story_blocks(teacher.build_markdown())
    blocks = [
        base.title_para(teacher.TITLE),
        base.subtitle_para(SUBTITLE),
        base.callout("整理思路：以教学大纲和教材章节为主线，能用画面记住的病变先写成一句话，其他内容继续用表格整理。"),
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
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    data = base.document_xml(blocks, sect_pr).encode("utf-8")
                elif item.filename.startswith("word/footer") and item.filename.endswith(".xml"):
                    data = blank_footer_xml()
                elif item.filename == "docProps/core.xml":
                    data = teacher.core_xml()
                elif item.filename == "docProps/app.xml":
                    data = teacher.app_xml(data)
                zout.writestr(item, data)
    shutil.move(temp, OUT)
    print(OUT)


if __name__ == "__main__":
    main()
