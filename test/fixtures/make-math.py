#!/usr/bin/env python3
"""Build the `math.docx` fixture from `single-page.docx`.

dxpdf lays out three OMML constructs — `m:r`, `m:sSup` and `m:f` — and drops the
whole subtree of every other one *without reporting an error*, so a document
full of equations converts to a PDF with blanks where the formulas were. This
fixture carries one paragraph per dropped construct plus one the engine already
renders, so the math pass can be tested at both ends: the constructs must be
absent from a plain conversion and present after lowering.

Every equation is written with digit markers, because the engine renders a
letter inside a math run as a mathematical-italic code point while a digit
survives as itself — digits make the assertions in `test/math.test.mjs` exact.

Run from this directory:  python make-math.py
"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

etree.register_namespace("w", W_NS)
etree.register_namespace("m", M_NS)

HERE = Path(__file__).resolve().parent
SOURCE = HERE / "single-page.docx"
TARGET = HERE / "math.docx"

# label -> the OMML body. The markers are digits so each one survives to the
# text layer unchanged, and each construct is given a marker no other uses.
CASES = [
    ("sSub", '<m:sSub><m:e><m:r><m:t>22</m:t></m:r></m:e>'
             '<m:sub><m:r><m:t>33</m:t></m:r></m:sub></m:sSub>'),
    ("sSubSup", '<m:sSubSup><m:e><m:r><m:t>17</m:t></m:r></m:e>'
                '<m:sub><m:r><m:t>18</m:t></m:r></m:sub>'
                '<m:sup><m:r><m:t>19</m:t></m:r></m:sup></m:sSubSup>'),
    ("rad", '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>'
            '<m:e><m:r><m:t>111</m:t></m:r></m:e></m:rad>'),
    ("rad.degree", '<m:rad><m:radPr/><m:deg><m:r><m:t>3</m:t></m:r></m:deg>'
                   '<m:e><m:r><m:t>444</m:t></m:r></m:e></m:rad>'),
    ("d", '<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr>'
          '<m:e><m:r><m:t>55</m:t></m:r></m:e></m:d>'),
    ("nary", '<m:nary><m:naryPr><m:chr m:val="\u2211"/></m:naryPr>'
             '<m:sub><m:r><m:t>66</m:t></m:r></m:sub>'
             '<m:sup><m:r><m:t>77</m:t></m:r></m:sup>'
             '<m:e><m:r><m:t>888</m:t></m:r></m:e></m:nary>'),
    ("acc", '<m:acc><m:accPr><m:chr m:val="\u0302"/></m:accPr>'
            '<m:e><m:r><m:t>99</m:t></m:r></m:e></m:acc>'),
    ("bar", '<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>'
            '<m:e><m:r><m:t>10</m:t></m:r></m:e></m:bar>'),
    ("groupChr", '<m:groupChr><m:groupChrPr><m:chr m:val="\u23df"/>'
                 '<m:pos m:val="bot"/></m:groupChrPr>'
                 '<m:e><m:r><m:t>11</m:t></m:r></m:e></m:groupChr>'),
    ("box", '<m:box><m:boxPr/><m:e><m:r><m:t>12</m:t></m:r></m:e></m:box>'),
    ("eqArr", '<m:eqArr><m:e><m:r><m:t>13</m:t></m:r></m:e>'
              '<m:e><m:r><m:t>14</m:t></m:r></m:e></m:eqArr>'),
    ("func", '<m:func><m:funcPr/><m:fName><m:r><m:t>sin</m:t></m:r></m:fName>'
             '<m:e><m:r><m:t>20</m:t></m:r></m:e></m:func>'),
    ("limLow", '<m:limLow><m:limLowPr/><m:e><m:r><m:t>21</m:t></m:r></m:e>'
               '<m:lim><m:r><m:t>23</m:t></m:r></m:lim></m:limLow>'),
    # A fraction whose numerator is a script. The engine stacks the fraction but
    # cannot position the script inside it, which is what `--fraction` decides
    # between: `native` keeps the stack and writes the script in Unicode,
    # `linear` gives up the stack so the script stays a real subscript.
    ("f.with-script", '<m:f><m:num><m:sSub><m:e><m:r><m:t>24</m:t></m:r></m:e>'
                      '<m:sub><m:r><m:t>25</m:t></m:r></m:sub></m:sSub></m:num>'
                      '<m:den><m:r><m:t>26</m:t></m:r></m:den></m:f>'),
    # Already inside the subset the engine renders: this one must survive the
    # pass byte for byte, which is what keeps a clean equation from churning.
    ("sSup-supported", '<m:sSup><m:e><m:r><m:t>15</m:t></m:r></m:e>'
                       '<m:sup><m:r><m:t>16</m:t></m:r></m:sup></m:sSup>'),
]


def build_paragraph(label: str, body: str):
    """One paragraph holding a marker label, the equation, and a terminator.

    The label carries no bracket or symbol an equation could also produce, so a
    test can assert on those symbols without the label answering for them.
    """
    xml = (
        f'<w:p xmlns:w="{W_NS}" xmlns:m="{M_NS}">'
        f"<w:r><w:t xml:space=\"preserve\">{label}=</w:t></w:r>"
        f"<m:oMath>{body}</m:oMath>"
        f'<w:r><w:t xml:space="preserve">|E</w:t></w:r>'
        f"</w:p>"
    )
    return etree.fromstring(xml)


def main() -> int:
    shutil.copyfile(SOURCE, TARGET)
    with zipfile.ZipFile(SOURCE) as archive:
        part = archive.read("word/document.xml")
        names = archive.namelist()
        infos = {info.filename: info for info in archive.infolist()}

    root = etree.fromstring(part)
    body = root.find(W + "body")
    for label, inner in CASES:
        body.append(build_paragraph(label, inner))

    rewritten = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                               standalone=True)
    with zipfile.ZipFile(TARGET, "w", zipfile.ZIP_DEFLATED) as outgoing:
        for name in names:
            data = rewritten if name == "word/document.xml" else None
            if data is None:
                with zipfile.ZipFile(SOURCE) as archive:
                    data = archive.read(name)
            outgoing.writestr(name, data, compress_type=infos[name].compress_type)

    print(f"wrote {TARGET} with {len(CASES)} equation(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
