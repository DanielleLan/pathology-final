#!/usr/bin/env python3
from __future__ import annotations

import os
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


SRC = Path("/Users/daniellelan/Desktop/Copy of 病理学期末复习.docx")
OUT = Path("/Users/daniellelan/Desktop/病理学期末复习_WPS可编辑版.docx")
WORK = Path("/Users/daniellelan/Desktop/病理期末理论准备/_docx_table_work/wps_unpacked")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
W = f"{{{W_NS}}}"
NS = {"w": W_NS}


def el(tag: str, attrs: dict[str, str] | None = None) -> ET.Element:
    return ET.Element(W + tag, attrs or {})


def sub(parent: ET.Element, tag: str, attrs: dict[str, str] | None = None) -> ET.Element:
    node = el(tag, attrs)
    parent.append(node)
    return node


def first(parent: ET.Element, path: str) -> ET.Element | None:
    return parent.find(path, NS)


def ensure(parent: ET.Element, tag: str, index: int | None = None) -> ET.Element:
    node = parent.find(f"w:{tag}", NS)
    if node is None:
        node = el(tag)
        if index is None:
            parent.insert(0, node)
        else:
            parent.insert(index, node)
    return node


def remove_all(parent: ET.Element, tag: str) -> None:
    for child in list(parent):
        if child.tag == W + tag:
            parent.remove(child)


def grid_span(tc: ET.Element) -> int:
    node = first(tc, "w:tcPr/w:gridSpan")
    return int(node.attrib.get(W + "val", "1")) if node is not None else 1


def row_cols(tr: ET.Element) -> int:
    return sum(grid_span(tc) for tc in tr.findall("w:tc", NS))


def blank_cell(width: int, header: bool = False) -> ET.Element:
    tc = el("tc")
    tc_pr = sub(tc, "tcPr")
    sub(tc_pr, "tcW", {W + "w": str(width), W + "type": "dxa"})
    sub(tc_pr, "vAlign", {W + "val": "center"})
    margins = sub(tc_pr, "tcMar")
    for side in ("top", "bottom"):
        sub(margins, side, {W + "w": "80", W + "type": "dxa"})
    for side in ("left", "right"):
        sub(margins, side, {W + "w": "100", W + "type": "dxa"})
    if header:
        sub(tc_pr, "shd", {W + "fill": "EAF2F8"})
    p = sub(tc, "p")
    ppr = sub(p, "pPr")
    sub(ppr, "spacing", {W + "before": "0", W + "after": "0", W + "line": "220", W + "lineRule": "auto"})
    sub(p, "r")
    return tc


def style_run(run: ET.Element, header: bool) -> None:
    rpr = ensure(run, "rPr")
    rfonts = first(rpr, "w:rFonts")
    if rfonts is None:
        rfonts = el("rFonts")
        rpr.insert(0, rfonts)
    rfonts.attrib[W + "ascii"] = "Calibri"
    rfonts.attrib[W + "hAnsi"] = "Calibri"
    rfonts.attrib[W + "eastAsia"] = "宋体"
    rfonts.attrib[W + "cs"] = "Calibri"
    for tag in ("sz", "szCs"):
        node = first(rpr, f"w:{tag}")
        if node is None:
            node = sub(rpr, tag)
        node.attrib[W + "val"] = "18" if header else "17"
    if header and first(rpr, "w:b") is None:
        sub(rpr, "b")


def style_cell(tc: ET.Element, width: int, header: bool) -> None:
    tc_pr = ensure(tc, "tcPr")
    # WPS editing is much easier with plain rectangular tables.
    remove_all(tc_pr, "gridSpan")
    remove_all(tc_pr, "vMerge")

    tc_w = first(tc_pr, "w:tcW")
    if tc_w is None:
        tc_w = sub(tc_pr, "tcW")
    tc_w.attrib[W + "w"] = str(width)
    tc_w.attrib[W + "type"] = "dxa"

    valign = first(tc_pr, "w:vAlign")
    if valign is None:
        valign = sub(tc_pr, "vAlign")
    valign.attrib[W + "val"] = "center"

    remove_all(tc_pr, "tcMar")
    margins = sub(tc_pr, "tcMar")
    for side in ("top", "bottom"):
        sub(margins, side, {W + "w": "80", W + "type": "dxa"})
    for side in ("left", "right"):
        sub(margins, side, {W + "w": "100", W + "type": "dxa"})

    shd = first(tc_pr, "w:shd")
    if header:
        if shd is None:
            shd = sub(tc_pr, "shd")
        shd.attrib[W + "fill"] = "EAF2F8"

    for p in tc.findall(".//w:p", NS):
        ppr = ensure(p, "pPr")
        spacing = first(ppr, "w:spacing")
        if spacing is None:
            spacing = sub(ppr, "spacing")
        spacing.attrib[W + "before"] = "0"
        spacing.attrib[W + "after"] = "0"
        spacing.attrib[W + "line"] = "220"
        spacing.attrib[W + "lineRule"] = "auto"
        jc = first(ppr, "w:jc")
        if jc is None:
            jc = sub(ppr, "jc")
        jc.attrib[W + "val"] = "center" if header else "left"
    for run in tc.findall(".//w:r", NS):
        style_run(run, header)


def set_table_properties(tbl: ET.Element, cols: int, usable_width: int) -> int:
    width = max(900, usable_width // max(cols, 1))
    table_width = width * cols

    tbl_pr = ensure(tbl, "tblPr")
    for tag in ("tblStyle", "tblW", "tblInd", "jc", "tblLayout", "tblCellMar", "tblBorders", "tblLook"):
        remove_all(tbl_pr, tag)

    tbl_pr.insert(0, el("tblStyle", {W + "val": "TableGrid"}))
    sub(tbl_pr, "tblW", {W + "w": str(table_width), W + "type": "dxa"})
    sub(tbl_pr, "jc", {W + "val": "center"})
    sub(tbl_pr, "tblLayout", {W + "type": "fixed"})

    mar = sub(tbl_pr, "tblCellMar")
    for side in ("top", "bottom"):
        sub(mar, side, {W + "w": "80", W + "type": "dxa"})
    for side in ("left", "right"):
        sub(mar, side, {W + "w": "100", W + "type": "dxa"})

    borders = sub(tbl_pr, "tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        sub(borders, side, {W + "val": "single", W + "sz": "6", W + "space": "0", W + "color": "AEB8C2"})
    sub(tbl_pr, "tblLook", {W + "val": "04A0", W + "firstRow": "1", W + "lastRow": "0", W + "firstColumn": "0", W + "lastColumn": "0", W + "noHBand": "1", W + "noVBand": "1"})

    grid = first(tbl, "w:tblGrid")
    if grid is None:
        insert_at = 1 if first(tbl, "w:tblPr") is not None else 0
        grid = el("tblGrid")
        tbl.insert(insert_at, grid)
    else:
        remove_all(grid, "gridCol")
    for _ in range(cols):
        sub(grid, "gridCol", {W + "w": str(width)})
    return width


def normalize_table(tbl: ET.Element, usable_width: int) -> tuple[int, int]:
    rows = tbl.findall("w:tr", NS)
    if not rows:
        return (0, 0)

    original_cols = max(row_cols(row) for row in rows)
    target_cols = original_cols + 1
    width = set_table_properties(tbl, target_cols, usable_width)

    for row_index, tr in enumerate(rows):
        while row_cols(tr) < original_cols:
            tr.append(blank_cell(width, row_index == 0))
        tr.append(blank_cell(width, row_index == 0))
        for tc in tr.findall("w:tc", NS):
            style_cell(tc, width, row_index == 0)

    new_tr = el("tr")
    tr_pr = sub(new_tr, "trPr")
    sub(tr_pr, "cantSplit")
    for _ in range(target_cols):
        cell = blank_cell(width)
        style_cell(cell, width, False)
        new_tr.append(cell)
    tbl.append(new_tr)
    return (len(rows) + 1, target_cols)


def set_document_page(root: ET.Element) -> None:
    # A4 landscape gives WPS more horizontal room, especially after adding a column.
    for sect_pr in root.findall(".//w:sectPr", NS):
        pg_sz = first(sect_pr, "w:pgSz")
        if pg_sz is None:
            pg_sz = sub(sect_pr, "pgSz")
        pg_sz.attrib[W + "w"] = "16838"
        pg_sz.attrib[W + "h"] = "11906"
        pg_sz.attrib[W + "orient"] = "landscape"

        pg_mar = first(sect_pr, "w:pgMar")
        if pg_mar is None:
            pg_mar = sub(sect_pr, "pgMar")
        pg_mar.attrib[W + "top"] = "720"
        pg_mar.attrib[W + "bottom"] = "720"
        pg_mar.attrib[W + "left"] = "720"
        pg_mar.attrib[W + "right"] = "720"
        pg_mar.attrib[W + "header"] = "420"
        pg_mar.attrib[W + "footer"] = "420"
        pg_mar.attrib[W + "gutter"] = "0"


def copy_and_patch() -> None:
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)
    with zipfile.ZipFile(SRC) as zin:
        zin.extractall(WORK)

    doc_xml = WORK / "word/document.xml"
    tree = ET.parse(doc_xml)
    root = tree.getroot()
    set_document_page(root)

    tables = root.findall(".//w:tbl", NS)
    dims = [normalize_table(tbl, 15398) for tbl in tables]

    tree.write(doc_xml, encoding="UTF-8", xml_declaration=True, short_empty_elements=False)

    if OUT.exists():
        OUT.unlink()
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
        for folder, _, filenames in os.walk(WORK):
            for filename in filenames:
                full = Path(folder) / filename
                zout.write(full, full.relative_to(WORK).as_posix())

    print(OUT)
    print(f"tables={len(tables)}")
    print("first", dims[0] if dims else None, "last", dims[-1] if dims else None)


if __name__ == "__main__":
    copy_and_patch()
