#!/usr/bin/env python3
"""Build the `wps-defects.docx` fixture from `single-page.docx`.

Injects the constructs WPS Office writes and dxpdf rejects, so the repair pass
is covered by a committed fixture instead of a private document:

* a second `<w:sz>` in a style's run-properties block;
* a second `<w:pPr>` on a paragraph;
* an `<m:oMath>` equation wrapped in a `<w:r>`, when it belongs to the paragraph.

Run from this directory:  python make-wps-defects.py
"""

from __future__ import annotations

import copy
import shutil
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

etree.register_namespace("m", M_NS)

HERE = Path(__file__).resolve().parent
SOURCE = HERE / "single-page.docx"
TARGET = HERE / "wps-defects.docx"


def add_duplicate_size(root) -> int:
    """Append a second w:sz to the first run-properties block that has one."""
    for style in root.iter(W + "style"):
        rpr = style.find(W + "rPr")
        if rpr is None:
            continue
        sizes = rpr.findall(W + "sz")
        if sizes:
            extra = copy.deepcopy(sizes[0])
            value = extra.get(W + "val")
            if value is not None and value.isdigit():
                extra.set(W + "val", str(int(value) - 1))
            rpr.append(extra)
            return 1
    return 0


def add_duplicate_paragraph_properties(body) -> int:
    """Give the first paragraph a second w:pPr, as WPS does."""
    for paragraph in body.iter(W + "p"):
        ppr = paragraph.find(W + "pPr")
        if ppr is None:
            continue
        paragraph.insert(list(paragraph).index(ppr) + 1, copy.deepcopy(ppr))
        return 1
    return 0


def add_misplaced_paragraph_properties(body) -> int:
    """Build a table-of-contents paragraph with its w:pPr after the run.

    WPS writes such an entry with the field runs first, which makes serde see
    the paragraph's flattened content twice and report `duplicate field
    `$value`` — an error that names no element.
    """
    paragraph = etree.Element(W + "p")
    run = etree.SubElement(paragraph, W + "r")
    text = etree.SubElement(run, W + "t")
    text.text = "toc entry"
    ppr = etree.SubElement(paragraph, W + "pPr")
    tabs = etree.SubElement(ppr, W + "tabs")
    tab = etree.SubElement(tabs, W + "tab")
    tab.set(W + "val", "right")
    tab.set(W + "leader", "dot")
    tab.set(W + "pos", "8312")

    # `w:sectPr`, when present, must remain the body's last child.
    section = body.find(W + "sectPr")
    if section is None:
        body.append(paragraph)
    else:
        body.insert(list(body).index(section), paragraph)
    return 1


def add_nested_run(body) -> int:
    """Wrap a run inside another run, which no reader should accept."""
    paragraphs = list(body.iter(W + "p"))
    host = paragraphs[-1] if paragraphs else body
    outer = etree.SubElement(host, W + "r")
    inner = etree.SubElement(outer, W + "r")
    text = etree.SubElement(inner, W + "t")
    text.text = "nested"
    return 1


def add_run_wrapped_math(body) -> int:
    """Attach an equation as a child of a run, which no reader should accept."""
    paragraphs = list(body.iter(W + "p"))
    host = paragraphs[-1] if paragraphs else body
    run = etree.SubElement(host, W + "r")
    math = etree.SubElement(run, M + "oMath")
    math_run = etree.SubElement(math, M + "r")
    text = etree.SubElement(math_run, M + "t")
    text.text = "x"
    return 1


def inject(input_path: Path, output_path: Path) -> dict[str, int]:
    counts = {
        "duplicate_sz": 0,
        "duplicate_pPr": 0,
        "misplaced_pPr": 0,
        "nested_run": 0,
        "wrapped_math": 0,
    }

    with zipfile.ZipFile(input_path) as archive:
        entries = [(info, archive.read(info.filename)) for info in archive.infolist()]

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as outgoing:
        for info, data in entries:
            if info.filename == "word/styles.xml":
                root = etree.fromstring(data)
                counts["duplicate_sz"] = add_duplicate_size(root)
                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            elif info.filename == "word/document.xml":
                root = etree.fromstring(data)
                body = root.find(W + "body")
                counts["duplicate_pPr"] = add_duplicate_paragraph_properties(body)
                counts["misplaced_pPr"] = add_misplaced_paragraph_properties(body)
                counts["nested_run"] = add_nested_run(body)
                counts["wrapped_math"] = add_run_wrapped_math(body)
                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            outgoing.writestr(info.filename, data, compress_type=info.compress_type)

    return counts


if __name__ == "__main__":
    if not SOURCE.exists():
        raise SystemExit(f"missing base fixture: {SOURCE}")
    shutil.copyfile(SOURCE, TARGET)
    print(f"wrote {TARGET.name}: {inject(SOURCE, TARGET)}")
