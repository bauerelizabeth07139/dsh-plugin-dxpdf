#!/usr/bin/env python3
"""Build the `parity-sections.docx` fixture.

Two sections split by an **odd-page** section break, which Word honours by
starting section 2 on an odd page and inserting a blank page when section 1
ends on one. dxpdf treats the break as a plain `nextPage`, so the fixture is
what proves the pagination pass restores the blank page.

Section 1 is one page long, so the fix must add exactly one page.

Run from this directory:  python make-parity-sections.py
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from docx import Document
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

HERE = Path(__file__).resolve().parent
TARGET = HERE / "parity-sections.docx"

# Kept short: section 1 must fit on a single page for the parity to be testable.
SECTION_ONE = ("A", "B")
SECTION_TWO = ("C", "D")


def build(target: Path, kind: str = "oddPage") -> None:
    document = Document()
    for name in SECTION_ONE + SECTION_TWO:
        document.add_paragraph(f"{name} " + "filler " * 30)
    document.save(target)

    with zipfile.ZipFile(target) as archive:
        entries = [(info, archive.read(info.filename)) for info in archive.infolist()]

    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as outgoing:
        for info, data in entries:
            if info.filename == "word/document.xml":
                root = etree.fromstring(data)
                body = root.find(W + "body")
                section = etree.fromstring(etree.tostring(body.find(W + "sectPr")))
                kind_element = etree.Element(W + "type")
                kind_element.set(W + "val", kind)
                section.insert(0, kind_element)

                # Attach the break to the last paragraph of section 1.
                paragraph = body.findall(W + "p")[len(SECTION_ONE) - 1]
                ppr = paragraph.find(W + "pPr")
                if ppr is None:
                    ppr = etree.Element(W + "pPr")
                    paragraph.insert(0, ppr)
                ppr.append(section)

                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            outgoing.writestr(info.filename, data, compress_type=info.compress_type)


if __name__ == "__main__":
    build(TARGET)
    print(f"wrote {TARGET.name} ({TARGET.stat().st_size:,} bytes)")
