#!/usr/bin/env python3
"""Give dxpdf the parity pages Word inserts at odd/even section breaks.

Word's `w:sectPr/w:type` of `oddPage` (or `evenPage`) means "start the next
section on an odd (even) page", inserting a blank page when the section before
it ended on the wrong parity. dxpdf documents these as *treated as nextPage*,
so it breaks to the very next page and those blank pages never appear.

dxpdf is a compiled wheel, so the fix is not to change its layout but to give
it a document whose plain `nextPage` breaks already land on the required
parity. An empty paragraph carrying `w:pageBreakBefore` reliably costs exactly
one page, so one filler per offending boundary reproduces Word's pagination.

Which boundaries offend depends on the page each section ends on, which is a
layout result — hence the measurement pass. The loop is not a search: adding a
filler shifts every later section by exactly one page and flips its parity, so
the state is a single running counter and one greedy pass settles it.

Usage:
    python paginate.py INPUT.docx --output FIXED.docx [--report report.json]
                                   [--python PATH] [--cli PATH]

It prints a JSON report and exits 0. Exit 2 means the document could not be
read; exit 3 means the engine could not convert a probe.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
PAGE = re.compile(rb"/Type\s*/Page(?![s])")

# Section break types that promise the next section a page of known parity.
PARITY = {"oddPage": 1, "evenPage": 0}

# An empty paragraph that starts a new page. Measured to cost exactly one page.
FILLER = f'<w:p xmlns:w="{W[1:-1]}"><w:pPr><w:pageBreakBefore/></w:pPr></w:p>'


@dataclass
class Boundary:
    """One section break, as stored on the last paragraph of its section."""

    index: int          # position of the carrying paragraph among body children
    kind: str           # the raw w:type value
    parity: int | None  # 1 for oddPage, 0 for evenPage, None otherwise
    ends_on: int = 0    # page the section ends on, once measured
    starts_next_on: int = 0
    filler: bool = False


@dataclass
class Report:
    boundaries: list[Boundary] = field(default_factory=list)
    probes: int = 0
    fillers: int = 0
    pages_before: int = 0
    pages_after: int = 0

    def as_dict(self) -> dict:
        return {
            "probes": self.probes,
            "fillers": self.fillers,
            "pagesBefore": self.pages_before,
            "pagesAfter": self.pages_after,
            "boundaries": [
                {
                    "paragraph": b.index,
                    "type": b.kind,
                    "endsOnPage": b.ends_on,
                    "nextStartsOnPage": b.starts_next_on,
                    "fillerInserted": b.filler,
                }
                for b in self.boundaries
            ],
        }


def read_parts(source: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(source) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def write_parts(target: Path, parts: dict[str, bytes], compress: dict[str, int]) -> None:
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data, compress_type=compress.get(name, zipfile.ZIP_DEFLATED))


def parse(source: Path) -> tuple[etree._Element, list[Boundary]]:
    parts = read_parts(source)
    root = etree.fromstring(parts["word/document.xml"])
    body = root.find(W + "body")
    boundaries: list[Boundary] = []
    for position, child in enumerate(list(body)):
        if child.tag != W + "p":
            continue
        ppr = child.find(W + "pPr")
        if ppr is None:
            continue
        section = ppr.find(W + "sectPr")
        if section is None:
            continue
        kind_element = section.find(W + "type")
        kind = kind_element.get(W + "val") if kind_element is not None else "nextPage"
        boundaries.append(
            Boundary(index=position, kind=kind, parity=PARITY.get(kind))
        )
    return root, boundaries


def build_probe(root: etree._Element, keep: int, final_section, fillers: set[int]) -> bytes:
    """Serialize the document truncated after body child `keep`, as one section.

    The trailing section break is detached from its paragraph and re-attached at
    body level. Appending the document's *own* final section instead would leave
    a second, empty section behind, and dxpdf gives that empty section a page of
    its own — which would make every probe read one page too long.
    """
    clone = etree.fromstring(etree.tostring(root))
    body = clone.find(W + "body")
    children = list(body)
    for child in children:
        body.remove(child)

    kept = [child for child in children[: keep + 1] if child.tag != W + "sectPr"]

    section_xml = None
    if kept and kept[-1].tag == W + "p":
        ppr = kept[-1].find(W + "pPr")
        if ppr is not None:
            section = ppr.find(W + "sectPr")
            if section is not None:
                section_xml = etree.tostring(section)
                ppr.remove(section)

    for position, child in enumerate(kept):
        if position in fillers:
            body.append(etree.fromstring(FILLER))
        body.append(child)

    if section_xml is not None:
        body.append(etree.fromstring(section_xml))
    elif final_section is not None:
        body.append(etree.fromstring(etree.tostring(final_section)))

    return etree.tostring(clone, xml_declaration=True, encoding="UTF-8", standalone=True)


def convert(python: str, cli: str, source: Path, target: Path) -> int:
    """Convert and return the produced page count."""
    proc = subprocess.run(
        [python, cli, str(source), "-o", str(target), "--quiet"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    if proc.returncode != 0:
        raise SystemExit(f"engine failed on {source}:\n{proc.stderr.strip()[:400]}")
    return len(PAGE.findall(target.read_bytes()))


def parity_boundaries(source: Path) -> list[Boundary]:
    """The boundaries that promise a parity, cheaply.

    Most documents break sections as plain `nextPage`, or have no sections at
    all, and paying a full XML parse for them is waste: a byte scan of the main
    part settles it before anything is parsed.
    """
    with zipfile.ZipFile(source) as archive:
        data = archive.read("word/document.xml")
    if b"oddPage" not in data and b"evenPage" not in data:
        return []
    _, boundaries = parse(source)
    return [boundary for boundary in boundaries if boundary.parity is not None]


def paginate(source: Path, destination: Path, python: str, cli: str) -> Report:
    report = Report()
    parts = read_parts(source)
    compress = {}
    with zipfile.ZipFile(source) as archive:
        compress = {info.filename: info.compress_type for info in archive.infolist()}

    root, boundaries = parse(source)
    report.boundaries = boundaries
    body = root.find(W + "body")
    final_section = body.find(W + "sectPr")

    work = Path(tempfile.mkdtemp(prefix="dxpdf-paginate-"))
    probe = work / "probe.docx"
    probe_pdf = work / "probe.pdf"

    def measure(keep: int, fillers: set[int]) -> int:
        """Page count of the document truncated after body child `keep`."""
        data = build_probe(root, keep, final_section, fillers)
        parts["word/document.xml"] = data
        write_parts(probe, parts, compress)
        report.probes += 1
        return convert(python, cli, probe, probe_pdf)

    try:
        last = len(list(body)) - 1
        report.pages_before = measure(last, set())

        # One greedy pass. Each probe measures the prefix ending at that
        # boundary with every filler decided so far already in place, so the
        # page the next section starts on is simply one past where this one
        # ended — the shifting is observed rather than modelled, and adding a
        # filler flips that parity, so a boundary never needs a second one.
        fillers: set[int] = set()
        for boundary in boundaries:
            boundary.ends_on = measure(boundary.index, fillers)
            start = boundary.ends_on + 1
            if boundary.parity is not None and start % 2 != boundary.parity:
                fillers.add(boundary.index)
                boundary.filler = True
                start += 1
            boundary.starts_next_on = start

        if fillers:
            parts["word/document.xml"] = build_probe(root, last, final_section, fillers)
            write_parts(destination, parts, compress)
            report.fillers = len(fillers)
            report.pages_after = measure(last, fillers)
        else:
            report.pages_after = report.pages_before

        return report
    finally:
        # The probes are whole copies of the document — tens of megabytes each
        # for a real book — so leaving them behind is not an option.
        shutil.rmtree(work, ignore_errors=True)


def _default_cli() -> str:
    """Locate a dxpdf CLI shim without assuming where it was installed.

    Order: the declared machine variable, then the copy shipped beside this
    helper. The plugin passes `--cli` explicitly, so this default only matters
    when the helper is run by hand.
    """
    declared = os.environ.get("DXPDF_CLI")
    if declared:
        return declared
    return str(Path(__file__).resolve().with_name("dxpdf_cli.py"))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="dxpdf-paginate",
        description="Insert the parity pages Word adds at odd/even section breaks.",
    )
    parser.add_argument("input", help="source .docx, never modified")
    parser.add_argument(
        "--output",
        help="path of the corrected .docx; required unless --check is given",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="report the parity section breaks and exit without converting "
        "anything, so a caller can skip the expensive measurement pass",
    )
    parser.add_argument("--python", default=sys.executable, help="interpreter running dxpdf_cli")
    parser.add_argument(
        "--cli",
        default=_default_cli(),
        help="path to dxpdf_cli.py",
    )
    parser.add_argument("--report", help="also write the JSON report here")
    args = parser.parse_args(argv)

    if args.check:
        try:
            found = parity_boundaries(Path(args.input))
        except (zipfile.BadZipFile, OSError, KeyError, etree.XMLSyntaxError) as exc:
            sys.stderr.write(f"dxpdf-paginate: {exc}\n")
            return 2
        json.dump(
            {
                "boundaries": [
                    {"paragraph": b.index, "type": b.kind} for b in found
                ]
            },
            sys.stdout,
            indent=2,
        )
        sys.stdout.write("\n")
        return 0

    if not args.output:
        parser.error("--output is required unless --check is given")

    try:
        report = paginate(Path(args.input), Path(args.output), args.python, args.cli)
    except (zipfile.BadZipFile, OSError, KeyError, etree.XMLSyntaxError) as exc:
        sys.stderr.write(f"dxpdf-paginate: {exc}\n")
        return 2

    payload = report.as_dict()
    if args.report:
        Path(args.report).write_text(json.dumps(payload, indent=2), encoding="utf-8")
    json.dump(payload, sys.stdout, indent=2)
    sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
