#!/usr/bin/env python3
from pathlib import Path
import re
import sys
import zipfile
import xml.etree.ElementTree as ET


def slide_key(name: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", name)
    return int(match.group(1)) if match else 0


def iter_text(xml_bytes: bytes):
    root = ET.fromstring(xml_bytes)
    for node in root.iter():
        if node.tag.endswith("}t") and node.text:
            yield node.text


def main() -> int:
    if len(sys.argv) != 3:
        print("Usage: extract_pptx_text.py <input.pptx> <output.txt>", file=sys.stderr)
        return 1

    src = Path(sys.argv[1])
    dst = Path(sys.argv[2])

    with zipfile.ZipFile(src) as zf:
        slide_names = sorted(
            (name for name in zf.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")),
            key=slide_key,
        )
        chunks = []
        for idx, name in enumerate(slide_names, start=1):
            texts = list(iter_text(zf.read(name)))
            chunks.append(f"===== Slide {idx} =====\n" + "\n".join(texts))

    dst.write_text("\n\n".join(chunks), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
