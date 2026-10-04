#!/usr/bin/env python3
"""Make a WPS-authored DOCX acceptable to strict OOXML readers.

WPS Office emits three constructs that Word tolerates but a strict schema-based
reader rejects. dxpdf is such a reader, so a document that opens perfectly in
an editor fails to convert.

1. **Repeated property elements.** A run-properties block is written twice::

       <w:rPr><w:rFonts .../><w:sz w:val="21"/><w:szCs w:val="24"/>
              <w:b/><w:color .../><w:sz w:val="20"/></w:rPr>

   ECMA-376 allows each of those at most once inside its parent, so dxpdf
   aborts with ``failed to deserialize XML: duplicate field `sz```.

2. **An equation wrapped in a run.** OMML math belongs to the paragraph, not
   to a run::

       <w:p><w:r><m:oMath>...</m:oMath></w:r></w:p>

   A run may not contain ``m:oMath``, so dxpdf aborts with ``unknown variant
   `oMath` `` while parsing run content — even though it renders paragraph
   level math correctly. Unwrapping the equation to its paragraph keeps every
   glyph and restores the structure Word itself would have written.

3. **A property element that is not first.** A paragraph's ``w:pPr`` must be
   its first child, but WPS emits a table-of-contents entry with the field
   runs ahead of it::

       <w:p><w:r>fldChar begin</w:r><w:r>instrText</w:r>
            <w:r>fldChar separate</w:r>
            <w:pPr>...</w:pPr>          <!-- too late -->
            <w:r>...</w:r></w:p>

   dxpdf flattens a paragraph's content into a single ``$value`` field, so the
   interruption makes it see that field twice and abort with ``duplicate field
   `$value` `` — an error that names no element and so cannot be repaired by
   name. Moving the block to the front restores the required order without
   altering a single property it carries.

Repair 1 removes the redundant element; repair 2 rewrites the run around the
equation without moving any content; repair 3 reorders. Every other ZIP entry
is copied through byte for byte, and the input file is never modified.

Usage:
    python normalize.py INPUT.docx --output FIXED.docx [--policy first|last]
                                   [--only sz,pPr]

`--only` restricts repair 1 to the property names given. The engine names the
offending field in its own error (``duplicate field `sz` ``), so feeding that
name back is both targeted and safe; without it the script sweeps a built-in
list of elements ECMA-376 allows at most once. Repairs 2 and 3 run either way:
both are structure-preserving, so neither needs the engine's permission.

It prints a JSON report on stdout and exits 0 whether or not anything changed.
Exit 2 means the input could not be read as a DOCX.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile

# The WordprocessingML main namespace. Only elements in it are considered, so
# a same-named element from the math or drawing namespaces is never touched.
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
# Office Math Markup Language, whose elements belong to the paragraph.
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# Elements ECMA-376 permits at most once inside their parent. Repeating
# containers (w:p, w:r, w:tbl, w:tr, w:tc, w:t, w:tab inside w:tabs, w:lvl
# inside w:abstractNum, m:oMath, ...) are deliberately absent: those are
# vectors and a repeat is legal.
SINGLETON = frozenset({
    # run properties
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
    "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
    "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
    "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish", "rPr",
    # paragraph properties
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr",
    "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
    "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl",
    "divId", "cnfStyle", "sectPr", "pPr",
    # style definition
    "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority",
    "semiHidden", "unhideWhenUsed", "qFormat", "locked", "personal",
    "personalCompose", "personalReply",
    # table, row and cell properties
    "tblStyle", "tblpPr", "tblOverlap", "tblW", "tblLayout", "tblCellMar",
    "tblLook", "tblCaption", "tblDescription", "trHeight", "tblHeader",
    "cantSplit", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders",
    "vAlign", "tcMar", "tcFitText", "hideMark",
    # section properties
    "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType",
    "cols", "formProt", "noEndnote", "titlePg", "bidi", "rtlGutter",
    "docGrid", "printerSettings",
    # numbering
    "start", "lvlRestart", "isLgl", "suff", "lvlJc", "lvlText",
    "lvlPicBulletId", "legacy", "abstractNumId", "numFmt",
})

# Elements that look like properties but legally repeat, recorded so the list
# above is never "fixed" back into including them.
#   tblStylePr  - a table style carries one per conditional region
#   rsid        - w:rsids holds the whole revision-id history
#   lvlOverride - w:num carries one per overridden numbering level
#   tab         - w:tabs holds the whole tab-stop list
#   lvl         - w:abstractNum carries one per level
#   p, r, tc, tr, tbl, t - the document body and table content
_REPEATS_LEGALLY = frozenset({
    "tblStylePr", "rsid", "lvlOverride", "tab", "lvl", "p", "r", "tc", "tr",
    "tbl", "t", "hyperlink", "bookmarkStart", "bookmarkEnd", "proofErr",
    "oMath", "oMathPara", "drawing", "pict", "AlternateContent", "sdt",
})

# Parents that may legally hold paragraph-level math, which is where an
# unwrapped equation is moved to.
MATH_HOSTS = frozenset({"p", "hyperlink", "sdtContent", "ins"})

# parent -> property elements ECMA-376 requires as its leading children, in
# this order. Anything else at the front means the block is out of position and
# is moved back; the order is schema-fixed, so no rendering decision is made.
#
# `tr` lists two: a row may carry `w:tblPrEx` before `w:trPr`, and promoting
# `trPr` ahead of an existing `tblPrEx` would break a row that was already
# valid — which is exactly what an unguarded "must be first" rule does.
LEADING_PROPS = {
    "p": ("pPr",),
    "r": ("rPr",),
    "tbl": ("tblPr",),
    "tr": ("tblPrEx", "trPr"),
    "tc": ("tcPr",),
}

# Repair rules feed each other, so one pass is not always enough. Six is far
# more than any document has needed; the loop stops early on a clean pass.
MAX_PASSES = 6

_ATTR = re.compile(r'xmlns:([\w.-]+)\s*=\s*"([^"]*)"|xmlns\s*=\s*"([^"]*)"')
_NAME = re.compile(r"</?\s*([^\s/>]+)")
_RUN_TAG = "</w:r>"


class Node:
    """One element, reduced to the spans it occupies in the source text."""

    __slots__ = (
        "local", "uri", "start", "open_end", "close_start", "end",
        "self_closing", "children",
    )

    def __init__(self, local: str, uri: str | None, start: int, self_closing: bool):
        self.local = local
        self.uri = uri
        self.start = start
        self.open_end = start
        self.close_start = start
        self.end = start
        self.self_closing = self_closing
        self.children: list[Node] = []


def tokens(text: str):
    """Yield `(kind, start, end, name, self_closing)` for every markup token.

    Everything between tokens is text and is never inspected, so a `<` or `>`
    inside a run's text cannot be mistaken for markup — valid XML escapes those.
    """
    index = 0
    length = len(text)
    while index < length:
        start = text.find("<", index)
        if start < 0:
            return
        if start > index:
            yield ("text", index, start, None, False)
            index = start

        if text.startswith("<!--", start):
            close = text.find("-->", start + 4)
            end = length if close < 0 else close + 3
            yield ("other", start, end, None, False)
            index = end
            continue
        if text.startswith("<![CDATA[", start):
            close = text.find("]]>", start + 9)
            end = length if close < 0 else close + 3
            yield ("other", start, end, None, False)
            index = end
            continue
        if text.startswith("<?", start):
            close = text.find("?>", start + 2)
            end = length if close < 0 else close + 2
            yield ("other", start, end, None, False)
            index = end
            continue
        if text.startswith("<!", start):
            close = text.find(">", start + 2)
            end = length if close < 0 else close + 1
            yield ("other", start, end, None, False)
            index = end
            continue

        cursor = start + 1
        quote = None
        while cursor < length:
            char = text[cursor]
            if quote is not None:
                if char == quote:
                    quote = None
            elif char in "\"'":
                quote = char
            elif char == ">":
                break
            cursor += 1
        end = cursor + 1
        raw = text[start:end]
        match = _NAME.match(raw)
        name = match.group(1) if match else None
        if raw.startswith("</"):
            yield ("close", start, end, name, False)
        else:
            yield ("open", start, end, name, raw.endswith("/>"))
        index = end


def namespace_map(text: str, limit: int = 65536) -> dict[str, str]:
    """Prefix -> URI declarations taken from the document's leading markup."""
    mapping: dict[str, str] = {}
    for prefixed, uri, default in _ATTR.findall(text[:limit]):
        if prefixed:
            mapping.setdefault(prefixed, uri)
        elif default and "" not in mapping:
            mapping[""] = default
    return mapping


def build_tree(text: str, namespaces: dict[str, str]) -> Node:
    """Parse the token stream into a span tree."""
    root = Node("", None, 0, False)
    root.open_end = 0
    stack = [root]
    for kind, start, end, name, self_closing in tokens(text):
        if kind == "open":
            if name is None or not stack:
                continue
            prefix, _, local = name.rpartition(":")
            node = Node(local, namespaces.get(prefix), start, self_closing)
            node.open_end = end
            stack[-1].children.append(node)
            if self_closing:
                node.end = end
            else:
                stack.append(node)
        elif kind == "close" and len(stack) > 1:
            node = stack.pop()
            if node.end <= node.start:
                node.close_start = start
                node.end = end
    return root


def find_duplicates(
    node: Node, policy: str, names: frozenset[str], edits: list, report: list[dict]
) -> None:
    """Queue a deletion for every redundant property element below `node`."""
    by_name: dict[str, list[Node]] = {}
    for child in node.children:
        by_name.setdefault(child.local, []).append(child)

    for local, kids in by_name.items():
        if len(kids) < 2 or local not in names:
            continue
        if any(kid.uri != W_NS for kid in kids):
            continue
        keep = 0 if policy == "first" else len(kids) - 1
        for position, kid in enumerate(kids):
            if position != keep:
                edits.append((kid.start, kid.end, ""))
        report.append({
            "kind": "duplicate-property",
            "parent": node.local,
            "element": local,
            "action": f"kept occurrence {keep + 1} of {len(kids)}",
        })

    for child in node.children:
        find_duplicates(child, policy, names, edits, report)


def rebuild_run(text: str, run: Node, math: set[int]) -> str:
    """Rewrite one run with its math children lifted out to the parent.

    The run's own start tag is reused, and the source text between children is
    carried across verbatim, so the result differs from the original only by
    where the equation sits.
    """
    open_tag = text[run.start:run.open_end]
    pieces: list[str] = []
    segment_start = run.open_end
    for child in run.children:
        if id(child) not in math:
            continue
        body = text[segment_start:child.start]
        if body.strip():
            pieces.append(f"{open_tag}{body}{_RUN_TAG}")
        pieces.append(text[child.start:child.end])
        segment_start = child.end
    tail = text[segment_start:run.close_start]
    if tail.strip():
        pieces.append(f"{open_tag}{tail}{_RUN_TAG}")
    return "".join(pieces)


def find_wrapped_math(node: Node, text: str, edits: list, report: list[dict]) -> None:
    """Queue a rewrite for every run that illegally contains OMML math."""
    for child in node.children:
        if (
            child.local == "r"
            and child.uri == W_NS
            and not child.self_closing
            and node.uri == W_NS
            and node.local in MATH_HOSTS
        ):
            math = [c for c in child.children if c.uri == MATH_NS]
            if math:
                edits.append((child.start, child.end, rebuild_run(text, child, {id(c) for c in math})))
                report.append({
                    "kind": "math-in-run",
                    "parent": node.local,
                    "element": "m:oMath",
                    "action": f"unwrapped {len(math)} equation(s) to the paragraph",
                })
        find_wrapped_math(child, text, edits, report)


def find_misplaced(
    node: Node, text: str, policy: str, edits: list, report: list[dict]
) -> None:
    """Queue a reorder for every leading-property block found out of position.

    The occurrence a duplicate sweep would keep is the one moved, so the two
    rules agree: with `first` the leading block stays, with `last` the trailing
    one is promoted. The whole surviving set is re-emitted in schema order,
    which keeps a `w:tblPrEx` ahead of the `w:trPr` that follows it.
    """
    sequence = LEADING_PROPS.get(node.local) if node.uri == W_NS else None
    if sequence and node.children:
        survivors: list[Node] = []
        for name in sequence:
            candidates = [c for c in node.children if c.local == name and c.uri == W_NS]
            if not candidates:
                continue
            keep = 0 if policy == "first" else len(candidates) - 1
            survivors.append(candidates[keep])

        if survivors:
            leading = [c for c in node.children if any(c is s for s in survivors)]
            ordered = sorted(survivors, key=lambda s: sequence.index(s.local))
            in_place = (
                node.children[: len(leading)] == leading and leading == ordered
            )
            if not in_place:
                # Insert the schema-ordered block after the start tag, then drop
                # each block where it stood: a move, expressed as disjoint edits.
                edits.append((
                    node.open_end,
                    node.open_end,
                    "".join(text[s.start:s.end] for s in ordered),
                ))
                for survivor in survivors:
                    edits.append((survivor.start, survivor.end, ""))
                report.append({
                    "kind": "misplaced-property",
                    "parent": node.local,
                    "element": "+".join(s.local for s in ordered),
                    "action": "moved to the front of <{}> in schema order".format(node.local),
                })

    for child in node.children:
        find_misplaced(child, text, policy, edits, report)


def unwrap_run(text: str, run: Node) -> str:
    """Replace a run with its children, handing its `w:rPr` down to child runs.

    WPS wraps a text-and-equation sequence in an extra `w:r`, which no reader
    may accept: a run cannot contain a run. Splicing the children into the
    parent restores what the document meant, and any run properties the outer
    element carried are pushed into child runs that lack their own.
    """
    rpr = None
    for child in run.children:
        if child.local == "rPr" and child.uri == W_NS:
            rpr = text[child.start:child.end]
            break

    pieces: list[str] = []
    cursor = run.open_end
    for child in run.children:
        if child.local == "rPr" and child.uri == W_NS:
            cursor = child.end
            continue
        pieces.append(text[cursor:child.start])
        is_run = child.local == "r" and child.uri == W_NS and not child.self_closing
        has_own_rpr = is_run and any(
            c.local == "rPr" and c.uri == W_NS for c in child.children
        )
        if rpr is not None and is_run and not has_own_rpr:
            open_tag = text[child.start:child.open_end]
            pieces.append(f"{open_tag}{rpr}{text[child.open_end:child.end]}")
        else:
            pieces.append(text[child.start:child.end])
        cursor = child.end
    pieces.append(text[cursor:run.close_start])
    return "".join(pieces)


def find_nested_runs(node: Node, text: str, edits: list, report: list[dict]) -> None:
    """Queue an unwrap for every run that contains another run."""
    for child in node.children:
        if (
            child.local == "r"
            and child.uri == W_NS
            and not child.self_closing
            and node.local == "r"
            and node.uri == W_NS
        ):
            edits.append((child.start, child.end, unwrap_run(text, child)))
            report.append({
                "kind": "nested-run",
                "parent": node.local,
                "element": "r",
                "action": f"unwrapped {len(child.children)} child element(s) into the paragraph",
            })
        find_nested_runs(child, text, edits, report)


def apply_edits(text: str, edits: list[tuple[int, int, str]]) -> tuple[str, int]:
    """Apply non-overlapping edits left to right, skipping any that collide."""
    edits.sort(key=lambda edit: (edit[0], -(edit[1] - edit[0])))
    pieces: list[str] = []
    cursor = 0
    applied = 0
    for start, end, replacement in edits:
        if start < cursor:
            continue
        pieces.append(text[cursor:start])
        pieces.append(replacement)
        cursor = end
        applied += 1
    pieces.append(text[cursor:])
    return "".join(pieces), applied


def effective_names(only: list[str] | None) -> tuple[frozenset[str], str]:
    """Decide which property elements this pass may deduplicate.

    An engine that reports ``duplicate field `sz` `` has already told us `sz`
    is modelled as a single field, which is stronger evidence than any list we
    could keep. So an explicit name set is honoured even for elements outside
    `SINGLETON` — minus the ones known to repeat legally, which no message
    should ever name.
    """
    if only:
        chosen = {name for name in only if name not in _REPEATS_LEGALLY}
        return frozenset(chosen), "targeted"
    return frozenset(SINGLETON - _REPEATS_LEGALLY), "bulk"


def normalize_part(
    text: str, policy: str, names: frozenset[str], report: list[dict]
) -> tuple[str, int]:
    """Return the repaired part and how many edits were applied.

    The rules feed each other: unwrapping a nested run exposes equations that
    then sit at paragraph level and need unwrapping themselves, and moving a
    property can reveal the next one out of order. Each pass re-parses what the
    previous one produced and stops as soon as a pass finds nothing, so the
    result is a fixed point rather than one rule's optimistic guess.
    """
    namespaces = namespace_map(text)
    if W_NS not in namespaces.values():
        return text, 0

    current = text
    applied_total = 0
    for _ in range(MAX_PASSES):
        root = build_tree(current, namespaces)
        edits: list[tuple[int, int, str]] = []
        find_duplicates(root, policy, names, edits, report)
        find_wrapped_math(root, current, edits, report)
        find_nested_runs(root, current, edits, report)
        find_misplaced(root, current, policy, edits, report)
        if not edits:
            break
        current, applied = apply_edits(current, edits)
        applied_total += applied
        if applied == 0:
            break
    return current, applied_total


def normalize_document(
    source: str, destination: str, policy: str, only: list[str] | None = None
) -> dict:
    names, mode = effective_names(only)
    report: list[dict] = []
    changed: dict[str, int] = {}

    with zipfile.ZipFile(source) as incoming:
        infos = {info.filename: info for info in incoming.infolist()}
        with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as outgoing:
            for name in incoming.namelist():
                data = incoming.read(name)
                if name.startswith("word/") and name.endswith(".xml"):
                    part_report: list[dict] = []
                    repaired, applied = normalize_part(
                        data.decode("utf-8"), policy, names, part_report
                    )
                    if applied:
                        data = repaired.encode("utf-8")
                        changed[name] = applied
                        for entry in part_report:
                            entry["part"] = name
                        report.extend(part_report)
                outgoing.writestr(name, data, compress_type=infos[name].compress_type)

    return {
        "changed": bool(changed),
        "policy": policy,
        "mode": mode,
        "targets": sorted(names) if mode == "targeted" else None,
        "edits": sum(changed.values()),
        "parts": changed,
        "repairs": report[:200],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="dxpdf-normalize",
        description="Repair WPS-authored OOXML that strict readers reject.",
    )
    parser.add_argument("input", help="source .docx, never modified")
    parser.add_argument("--output", required=True, help="path of the repaired .docx")
    parser.add_argument(
        "--policy",
        choices=("first", "last"),
        default="first",
        help="which occurrence of a duplicated property to keep (default: first)",
    )
    parser.add_argument(
        "--only",
        default=None,
        metavar="NAME[,NAME...]",
        help="deduplicate only these property element names. The engine names "
        "them in its `duplicate field` error, which is far safer than guessing; "
        "omit to sweep the built-in singleton list instead.",
    )
    parser.add_argument("--quiet", action="store_true", help="print nothing on success")
    args = parser.parse_args(argv)

    only = [part.strip() for part in args.only.split(",") if part.strip()] if args.only else None

    try:
        result = normalize_document(args.input, args.output, args.policy, only)
    except (zipfile.BadZipFile, OSError, UnicodeDecodeError) as exc:
        sys.stderr.write(f"dxpdf-normalize: {exc}\n")
        return 2

    if not args.quiet:
        json.dump(result, sys.stdout, ensure_ascii=False, indent=2)
        sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
