#!/usr/bin/env python3
from __future__ import annotations

import copy
import os
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


SRC = Path("/Users/daniellelan/Desktop/Copy of 病理学期末复习.docx")
OUT = Path("/Users/daniellelan/Desktop/病理学期末复习_表格整理版.docx")
WORK = Path("/Users/daniellelan/Desktop/病理期末理论准备/_docx_table_work/unpacked")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
W = f"{{{W_NS}}}"


def el(tag: str, attrs: dict[str, str] | None = None) -> ET.Element:
    return ET.Element(W + tag, attrs or {})


def sub(parent: ET.Element, tag: str, attrs: dict[str, str] | None = None) -> ET.Element:
    child = el(tag, attrs)
    parent.append(child)
    return child


def wval(value: str) -> dict[str, str]:
    return {W + "val": value}


def first(parent: ET.Element, path: str) -> ET.Element | None:
    return parent.find(path, {"w": W_NS})


def ensure(parent: ET.Element, tag: str, index: int | None = None) -> ET.Element:
    node = parent.find(f"w:{tag}", {"w": W_NS})
    if node is None:
        node = el(tag)
        if index is None:
            parent.insert(0, node)
        else:
            parent.insert(index, node)
    return node


def clear_children(parent: ET.Element, tag: str) -> None:
    for child in list(parent):
        if child.tag == W + tag:
            parent.remove(child)


def cell_grid_span(tc: ET.Element) -> int:
    grid_span = first(tc, "w:tcPr/w:gridSpan")
    if grid_span is None:
        return 1
    return int(grid_span.attrib.get(W + "val", "1"))


def row_col_count(tr: ET.Element) -> int:
    return sum(cell_grid_span(tc) for tc in tr.findall("w:tc", {"w": W_NS}))


def set_cell_text_empty(tc: ET.Element) -> None:
    for child in list(tc):
        if child.tag != W + "tcPr":
            tc.remove(child)
    p = sub(tc, "p")
    ppr = sub(p, "pPr")
    sub(ppr, "jc", {W + "val": "left"})
    sub(p, "r")


def make_blank_cell(width: int, shaded: bool = False) -> ET.Element:
    tc = el("tc")
    tc_pr = sub(tc, "tcPr")
    sub(tc_pr, "tcW", {W + "w": str(width), W + "type": "dxa"})
    sub(tc_pr, "vAlign", {W + "val": "center"})
    margins = sub(tc_pr, "tcMar")
    for side in ("top", "bottom"):
        sub(margins, side, {W + "w": "70", W + "type": "dxa"})
    for side in ("left", "right"):
        sub(margins, side, {W + "w": "90", W + "type": "dxa"})
    if shaded:
        sub(tc_pr, "shd", {W + "fill": "EAF2F8"})
    set_cell_text_empty(tc)
    return tc


def set_run_font(run: ET.Element, size_half_points: str) -> None:
    rpr = ensure(run, "rPr")
    r_fonts = first(rpr, "w:rFonts")
    if r_fonts is None:
        r_fonts = el("rFonts")
        rpr.insert(0, r_fonts)
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        r_fonts.attrib[W + key] = "DengXian"
    sz = first(rpr, "w:sz")
    if sz is None:
        sz = sub(rpr, "sz")
    sz.attrib[W + "val"] = size_half_points
    sz_cs = first(rpr, "w:szCs")
    if sz_cs is None:
        sz_cs = sub(rpr, "szCs")
    sz_cs.attrib[W + "val"] = size_half_points


def paragraph_cleanup(p: ET.Element) -> None:
    ppr = ensure(p, "pPr")
    spacing = first(ppr, "w:spacing")
    if spacing is None:
        spacing = sub(ppr, "spacing")
    spacing.attrib[W + "before"] = "0"
    spacing.attrib[W + "after"] = "0"
    spacing.attrib[W + "line"] = "220"
    spacing.attrib[W + "lineRule"] = "auto"


def style_cell(tc: ET.Element, width: int, is_header: bool) -> None:
    tc_pr = ensure(tc, "tcPr")
    clear_children(tc_pr, "gridSpan")
    clear_children(tc_pr, "vMerge")
    tc_w = first(tc_pr, "w:tcW")
    if tc_w is None:
        tc_w = sub(tc_pr, "tcW")
    tc_w.attrib[W + "w"] = str(width)
    tc_w.attrib[W + "type"] = "dxa"
    v_align = first(tc_pr, "w:vAlign")
    if v_align is None:
        v_align = sub(tc_pr, "vAlign")
    v_align.attrib[W + "val"] = "center"

    clear_children(tc_pr, "tcMar")
    margins = sub(tc_pr, "tcMar")
    for side in ("top", "bottom"):
        sub(margins, side, {W + "w": "70", W + "type": "dxa"})
    for side in ("left", "right"):
        sub(margins, side, {W + "w": "90", W + "type": "dxa"})

    shd = first(tc_pr, "w:shd")
    if is_header:
        if shd is None:
            shd = sub(tc_pr, "shd")
        shd.attrib[W + "fill"] = "EAF2F8"
    elif shd is not None and shd.attrib.get(W + "fill") == "EAF2F8":
        tc_pr.remove(shd)

    for p in tc.findall(".//w:p", {"w": W_NS}):
        paragraph_cleanup(p)
        ppr = ensure(p, "pPr")
        jc = first(ppr, "w:jc")
        if jc is None:
            jc = sub(ppr, "jc")
        jc.attrib[W + "val"] = "center" if is_header else "left"
    for run in tc.findall(".//w:r", {"w": W_NS}):
        set_run_font(run, "18" if is_header else "17")
        rpr = ensure(run, "rPr")
        bold = first(rpr, "w:b")
        if is_header and bold is None:
            sub(rpr, "b")
        if not is_header and bold is not None:
            rpr.remove(bold)


def set_table_style(tbl: ET.Element, cols: int) -> None:
    tbl_pr = ensure(tbl, "tblPr")
    clear_children(tbl_pr, "tblStyle")
    tbl_style = el("tblStyle", {W + "val": "TableGrid"})
    tbl_pr.insert(0, tbl_style)

    tbl_w = first(tbl_pr, "w:tblW")
    if tbl_w is None:
        tbl_w = sub(tbl_pr, "tblW")
    tbl_w.attrib[W + "w"] = "9638"
    tbl_w.attrib[W + "type"] = "dxa"

    layout = first(tbl_pr, "w:tblLayout")
    if layout is None:
        layout = sub(tbl_pr, "tblLayout")
    layout.attrib[W + "type"] = "fixed"

    jc = first(tbl_pr, "w:jc")
    if jc is None:
        jc = sub(tbl_pr, "jc")
    jc.attrib[W + "val"] = "center"

    clear_children(tbl_pr, "tblCellMar")
    tbl_cell_mar = sub(tbl_pr, "tblCellMar")
    for side in ("top", "bottom"):
        sub(tbl_cell_mar, side, {W + "w": "70", W + "type": "dxa"})
    for side in ("left", "right"):
        sub(tbl_cell_mar, side, {W + "w": "90", W + "type": "dxa"})

    borders = first(tbl_pr, "w:tblBorders")
    if borders is None:
        borders = sub(tbl_pr, "tblBorders")
    clear_children(borders, "top")
    clear_children(borders, "left")
    clear_children(borders, "bottom")
    clear_children(borders, "right")
    clear_children(borders, "insideH")
    clear_children(borders, "insideV")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        sub(borders, side, {W + "val": "single", W + "sz": "6", W + "space": "0", W + "color": "B7C2CC"})

    tbl_grid = first(tbl, "w:tblGrid")
    if tbl_grid is None:
        insert_at = 0
        for i, child in enumerate(list(tbl)):
            if child.tag == W + "tblPr":
                insert_at = i + 1
        tbl_grid = el("tblGrid")
        tbl.insert(insert_at, tbl_grid)
    else:
        clear_children(tbl_grid, "gridCol")
    width = max(760, 9638 // max(cols, 1))
    for _ in range(cols):
        sub(tbl_grid, "gridCol", {W + "w": str(width)})


def normalize_table(tbl: ET.Element) -> tuple[int, int]:
    rows = tbl.findall("w:tr", {"w": W_NS})
    if not rows:
        return (0, 0)
    original_max_cols = max(row_col_count(row) for row in rows)
    target_cols = original_max_cols + 1
    width = max(760, 9638 // max(target_cols, 1))

    set_table_style(tbl, target_cols)

    for r_idx, tr in enumerate(rows):
        while row_col_count(tr) < original_max_cols:
            tr.append(make_blank_cell(width, shaded=(r_idx == 0)))
        tr.append(make_blank_cell(width, shaded=(r_idx == 0)))
        for tc in tr.findall("w:tc", {"w": W_NS}):
            style_cell(tc, width, r_idx == 0)

    new_tr = el("tr")
    tr_pr = sub(new_tr, "trPr")
    sub(tr_pr, "cantSplit")
    for _ in range(target_cols):
        new_tr.append(make_blank_cell(width, shaded=False))
    tbl.append(new_tr)
    for tc in new_tr.findall("w:tc", {"w": W_NS}):
        style_cell(tc, width, False)

    return (len(rows) + 1, target_cols)


def update_section_margins(root: ET.Element) -> None:
    for sect_pr in root.findall(".//w:sectPr", {"w": W_NS}):
        pg_sz = first(sect_pr, "w:pgSz")
        if pg_sz is None:
            pg_sz = sub(sect_pr, "pgSz")
        pg_sz.attrib[W + "w"] = "12240"
        pg_sz.attrib[W + "h"] = "15840"
        pg_mar = first(sect_pr, "w:pgMar")
        if pg_mar is None:
            pg_mar = sub(sect_pr, "pgMar")
        pg_mar.attrib[W + "top"] = "900"
        pg_mar.attrib[W + "right"] = "720"
        pg_mar.attrib[W + "bottom"] = "900"
        pg_mar.attrib[W + "left"] = "720"
        pg_mar.attrib[W + "header"] = "450"
        pg_mar.attrib[W + "footer"] = "450"
        pg_mar.attrib[W + "gutter"] = "0"


def main() -> None:
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)
    with zipfile.ZipFile(SRC) as zin:
        zin.extractall(WORK)

    doc_xml = WORK / "word/document.xml"
    tree = ET.parse(doc_xml)
    root = tree.getroot()
    tables = root.findall(".//w:tbl", {"w": W_NS})
    dims = []
    for tbl in tables:
        dims.append(normalize_table(tbl))
    update_section_margins(root)
    tree.write(doc_xml, encoding="UTF-8", xml_declaration=True, short_empty_elements=False)

    if OUT.exists():
        OUT.unlink()
    with zipfile.ZipFile(OUT, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for folder, _, files in os.walk(WORK):
            for filename in files:
                full = Path(folder) / filename
                arcname = full.relative_to(WORK).as_posix()
                zout.write(full, arcname)
    print(f"wrote {OUT}")
    print(f"tables formatted: {len(tables)}")
    for i, (rows, cols) in enumerate(dims, 1):
        print(f"{i}: {rows} rows x {cols} cols")


if __name__ == "__main__":
    main()
