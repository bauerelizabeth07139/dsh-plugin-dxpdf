#!/usr/bin/env python3
"""Lower office math the dxpdf engine cannot lay out into content it can.

The engine implements three OMML constructs — ``m:r``, ``m:sSup`` and ``m:f`` —
and silently drops the whole subtree of every other one. That is not a parse
error, so a document full of equations converts "successfully" into a PDF with
blank spaces where the formulas were. Measured coverage of the constructs that
technical writing actually uses:

    renders:     m:r   m:sSup   m:f
    dropped:     m:sSub  m:sSubSup  m:sPre  m:rad  m:d  m:nary  m:func
                 m:limLow  m:limUpp  m:acc  m:bar  m:groupChr  m:box
                 m:borderBox  m:marg  m:phant  m:eqArr  m:m

The drop is total, not partial: ``<m:rad><m:e><m:sSup>…`` loses the superscript
as well, so the absence of an error is no evidence an equation survived. A
document whose formulas are all ``m:r`` and ``m:sSup`` converts correctly today,
which is why this pass is a per-equation decision and never a rewrite in place.

Two facts about run properties decide the whole design, both measured:

* **outside math** ``w:rPr`` is honoured — ``w:vertAlign="subscript"`` gives
  58 % of the size on a baseline 1.0 pt lower, ``w:sz="16"`` gives 8 pt;
* **inside math** it is ignored — the same run properties leave text at 11 pt on
  the math baseline.

So a subscript is *hoisted out of the equation* into an ordinary run, where the
engine positions it correctly, while a fraction stays a native ``m:f`` because
it is the one construct that stacks — its numerator and denominator are a
math-only slot, and there a script can only be spelled with Unicode characters
(``σ₀``) or brackets (``L_(D)``).

A third measurement fixes how the marks are written. The engine does no complex
text layout: a combining character keeps its zero advance but **no ink reaches
the page**, so ``G`` + U+0305 is a plain ``G`` whatever the run asks for.
Spacing characters do draw, so an accent or a rule follows its base as the
spacing modifier letter of the same shape (``xˆ``, ``v→``, ``x‾``) — exactly how
the same marks are written in plain text.

Everything else is written linearly in the engine's own vocabulary: a radical
becomes ``√`` with its radicand parenthesised, a delimiter pair becomes its own
characters, an n-ary operator keeps its limits as scripts, and the transparent
containers (``m:box``, ``m:borderBox``, ``m:marg``, ``m:phant``) disappear.

An equation whose every construct the engine already renders is left byte for
byte alone; every ZIP entry this pass does not change is copied through
verbatim, and the source file is never modified.

Usage:
    python omml.py INPUT.docx --output LOWERED.docx
                    [--fraction native|linear] [--radical parens|overline]
                    [--scripts unicode|brackets]

``--check`` scans and reports without writing anything, which is how the plugin
decides whether a rewrite is worth a second conversion.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from xml.sax.saxutils import escape

# Office Math Markup Language, whose elements belong to the paragraph.
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
# The WordprocessingML main namespace: used for the runs this pass emits, and to
# recognise a part that is a document at all.
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Constructs the engine lays out. Everything else loses its subtree.
RENDERS = frozenset({"r", "sSup", "f"})

# The constructs measured to be dropped. Naming them turns the pre-scan into a
# search for a known set instead of a guess about what an equation contains.
DROPPED = (
    "sSub", "sSubSup", "sPre", "rad", "d", "nary", "func", "limLow", "limUpp",
    "acc", "bar", "groupChr", "box", "borderBox", "marg", "phant", "eqArr", "m",
)
_DROPPED = frozenset(DROPPED)

# Containers that carry no glyphs of their own, so removing them loses nothing.
TRANSPARENT = frozenset({"box", "borderBox", "marg", "phant"})

# The math-only containers that hold children rather than text.
SLOTS = frozenset({"e", "num", "den", "sup", "sub", "lim", "fName", "deg"})

# The font a lowered run asks for, so the pieces of one equation stay in one
# typeface. The engine derives the script size itself.
SCRIPT_FONT = "Cambria Math"

# Accent and rule characters, keyed by what OMML stores in `m:chr`. Both the
# combining character and the spacing character it is normally written as are
# accepted, since documents contain either.
#
# Every value is a *spacing* character, and that is deliberate: measured on this
# engine, a combining mark keeps its zero advance but is never painted — the
# glyph reaches the text layer and nothing reaches the page. A spacing modifier
# letter draws correctly, so an accent follows its base instead of sitting on
# it, which is also how the same marks are written in plain text.
OVERLINE = "\u203e"          # ‾  over a whole group
UNDERLINE_MARK = "\u02cd"    # ˍ  below a whole group
MACRON = "\u00af"            # ¯
ACCENTS = {
    "\u0302": "\u02c6", "^": "\u02c6", "\u02c6": "\u02c6",      # hat
    "\u0300": "\u02cb", "`": "\u02cb", "\u02cb": "\u02cb",      # grave
    "\u0301": "\u02ca", "'": "\u02ca", "\u02ca": "\u02ca",      # acute
    "\u0303": "\u02dc", "~": "\u02dc", "\u02dc": "\u02dc",      # tilde
    "\u0304": MACRON, "\u00af": MACRON,                         # macron
    "\u0305": OVERLINE, "\u203e": OVERLINE, "\u00b0": OVERLINE,  # overline
    "\u0306": "\u02d8", "\u02d8": "\u02d8",                     # breve
    "\u0307": "\u02d9", "\u02d9": "\u02d9",                     # dot above
    "\u0308": "\u00a8", "\u00a8": "\u00a8",                     # diaeresis
    "\u030a": "\u02da", "\u02da": "\u02da",                     # ring above
    "\u030c": "\u02c7", "\u02c7": "\u02c7",                     # caron
    "\u20d7": "\u2192", "\u2192": "\u2192",                     # vector
    "\u20d6": "\u2190", "\u2190": "\u2190",                     # left vector
    "\u20db": "\u00a8",                                         # three dots
}

# Unicode sub/superscript forms. They are used only in a math-only slot, where
# the engine will not position a plain run.
SUBSCRIPT = str.maketrans({
    "0": "\u2080", "1": "\u2081", "2": "\u2082", "3": "\u2083", "4": "\u2084",
    "5": "\u2085", "6": "\u2086", "7": "\u2087", "8": "\u2088", "9": "\u2089",
    "+": "\u208a", "-": "\u208b", "=": "\u208c", "(": "\u208d", ")": "\u208e",
    "a": "\u2090", "e": "\u2091", "h": "\u2095", "i": "\u1d62", "j": "\u2c7c",
    "k": "\u2096", "l": "\u2097", "m": "\u2098", "n": "\u2099", "o": "\u2092",
    "p": "\u209a", "r": "\u1d63", "s": "\u209b", "t": "\u209c", "u": "\u1d64",
    "v": "\u1d65", "x": "\u2093",
})
SUPERSCRIPT = str.maketrans({
    "0": "\u2070", "1": "\u00b9", "2": "\u00b2", "3": "\u00b3", "4": "\u2074",
    "5": "\u2075", "6": "\u2076", "7": "\u2077", "8": "\u2078", "9": "\u2079",
    "+": "\u207a", "-": "\u207b", "=": "\u207c", "(": "\u207d", ")": "\u207e",
    "n": "\u207f", "i": "\u2071", "j": "\u02b2", "a": "\u1d43", "b": "\u1d47",
    "c": "\u1d9c", "d": "\u1d48", "e": "\u1d49", "f": "\u1da0", "g": "\u1d4d",
    "h": "\u02b0", "k": "\u1d4f", "l": "\u02e1", "m": "\u1d50", "o": "\u1d52",
    "p": "\u1d56", "r": "\u02b3", "s": "\u02e2", "t": "\u1d57", "u": "\u1d58",
    "v": "\u1d5b", "w": "\u02b7", "x": "\u02e3", "y": "\u02b8", "z": "\u1dbb",
})
_SUB_KEYS = frozenset(chr(key) for key in SUBSCRIPT)
_SUP_KEYS = frozenset(chr(key) for key in SUPERSCRIPT)

# Characters looser-binding than the fraction slash: a side containing one of
# them at bracket depth zero needs parentheses once the fraction is spelled out.
LOOSE = frozenset("+-=\u00b1\u2213<>\u2264\u2265\u2248\u2260,;\u00b7\u00d7\u00f7 ")

# A radicand may take a combining overline only when it is a run of simple
# glyphs: a mark cannot be drawn over a stacked fraction or a bracket pair.
_ATOM = re.compile(r"^[0-9A-Za-z\u0370-\u03ff\u1d00-\u1dff\u2070-\u209f"
                   r"\u2100-\u214f+\-*/=.,'\"(){}\[\]|]{1,24}$")

_ATTR = re.compile(r'([\w:.\-]+)\s*=\s*"([^"]*)"')
_TAG = re.compile(r'<(/?)([\w:.\-]+)((?:\s+[\w:.\-]+\s*=\s*"[^"]*")*)\s*(/?)>')
_NUMERIC = re.compile(r"&#(x?)([0-9a-fA-F]+);")
_ENTITIES = (("&lt;", "<"), ("&gt;", ">"), ("&quot;", '"'), ("&apos;", "'"), ("&amp;", "&"))


def unescape(text: str) -> str:
    """Turn XML character references back into the characters they denote."""
    if "&" not in text:
        return text
    text = _NUMERIC.sub(lambda m: chr(int(m.group(2), 16 if m.group(1) else 10)), text)
    for entity, char in _ENTITIES:
        if entity in text:
            text = text.replace(entity, char)
    return text


class El:
    """One element of a parsed math fragment, keeping its original prefix."""

    __slots__ = ("prefix", "local", "attrs", "children")

    def __init__(self, prefix: str, local: str, attrs: list[tuple[str, str]]):
        self.prefix = prefix
        self.local = local
        self.attrs = attrs
        self.children: list = []

    def first(self, local: str):
        for child in self.children:
            if isinstance(child, El) and child.local == local:
                return child
        return None

    def all(self, local: str):
        return [c for c in self.children if isinstance(c, El) and c.local == local]

    def prop(self, block: str, name: str) -> str | None:
        """The value of ``m:name`` inside this element's ``m:<block>``."""
        holder = self.first(block)
        if holder is None:
            return None
        stored = holder.first(name)
        if stored is None:
            return None
        for key, value in stored.attrs:
            if key in ("m:val", "val"):
                return value
        return None


def parse(fragment: str) -> list:
    """Parse an XML fragment into a list of ``El`` and text nodes."""
    root: list = []
    stack: list[list] = [root]
    cursor = 0
    for match in _TAG.finditer(fragment):
        if match.start() > cursor:
            text = unescape(fragment[cursor:match.start()])
            if text:
                stack[-1].append(text)
        cursor = match.end()
        closing, name, attrs, self_closing = match.groups()
        if closing:
            if len(stack) > 1:
                stack.pop()
            continue
        prefix, _, local = name.rpartition(":")
        element = El(prefix, local, _ATTR.findall(attrs))
        stack[-1].append(element)
        if not self_closing:
            stack.append(element.children)
    if cursor < len(fragment):
        text = unescape(fragment[cursor:])
        if text:
            stack[-1].append(text)
    return root


def walk(node):
    """Every element below `node`, depth first."""
    if isinstance(node, El):
        yield node
        for child in node.children:
            yield from walk(child)


def raw_text(node) -> str:
    """Every character below `node`, with markup removed."""
    if isinstance(node, str):
        return node
    return "".join(raw_text(child) for child in node.children)


def serialize(node) -> str:
    """Render a parsed node back to XML under its original prefixes."""
    if isinstance(node, str):
        return escape(node)
    name = f"{node.prefix}:{node.local}" if node.prefix else node.local
    attrs = "".join(f' {key}="{escape(value, {chr(34): "&quot;"})}"'
                    for key, value in node.attrs)
    if not node.children:
        return f"<{name}{attrs}/>"
    body = "".join(serialize(child) for child in node.children)
    return f"<{name}{attrs}>{body}</{name}>"


def native_ok(node) -> bool:
    """True when the engine lays this subtree out instead of dropping it."""
    for child in node.children:
        if not isinstance(child, El):
            continue
        local = child.local
        if local in ("r", "t") or local.endswith("Pr"):
            continue
        if local in SLOTS:
            if not native_ok(child):
                return False
            continue
        if local == "sSup":
            base, sup = child.first("e"), child.first("sup")
            if base is None or sup is None or not native_ok(base) or not native_ok(sup):
                return False
            continue
        if local == "f":
            num, den = child.first("num"), child.first("den")
            if num is None or den is None or not native_ok(num) or not native_ok(den):
                return False
            continue
        return False
    return True


class Lowering:
    """Rewrite one equation into the subset the engine implements.

    Two output languages exist because the engine accepts different things in
    different places: `math` stays inside ``m:oMath`` and may use only ``m:r``,
    ``m:sSup`` and ``m:f``; `para` leaves the equation as ordinary runs, where
    ``w:vertAlign`` works and a script becomes a real sub/superscript.
    """

    def __init__(self, ns: "Names", fraction: str = "native",
                 radical: str = "parens", scripts: str = "unicode"):
        self.ns = ns
        self.fraction = fraction
        self.radical = radical
        self.scripts = scripts
        self.counts: dict[str, int] = {}

    # -- emitters ----------------------------------------------------------
    @property
    def m(self) -> str:
        return self.ns.math_tag

    @property
    def w(self) -> str:
        return self.ns.word_tag

    def run(self, text: str, align: str | None = None, italic: bool = False) -> str:
        """An ordinary run, which is where ``w:vertAlign`` is honoured."""
        if not text:
            return ""
        word = self.w
        props = (f'<{word}rFonts w:ascii="{SCRIPT_FONT}" w:hAnsi="{SCRIPT_FONT}"'
                 f' w:cs="{SCRIPT_FONT}"/>')
        if italic:
            props += f"<{word}i/><{word}iCs/>"
        if align:
            props += f'<{word}vertAlign w:val="{align}"/>'
        return (f"<{word}r><{word}rPr>{props}</{word}rPr>"
                f'<{word}t xml:space="preserve">{escape(text)}</{word}t></{word}r>')

    def glyph(self, text: str, italic: bool = False) -> str:
        """A single operator glyph, kept in the math font via a math island."""
        extra = f'<{self.m}i/>' if italic else ""
        return (f'<{self.m}oMath><{self.m}r><{self.m}rPr>{extra}</{self.m}rPr>'
                f'<{self.m}t xml:space="preserve">{escape(text)}</{self.m}t>'
                f"</{self.m}r></{self.m}oMath>")

    def island(self, inner: str) -> str:
        return f"<{self.m}oMath>{inner}</{self.m}oMath>"

    # -- small helpers -----------------------------------------------------
    @staticmethod
    def script(text: str, table: dict, keys: frozenset, unicode_ok: bool) -> str:
        """Attach a script linearly, in Unicode if the whole run has a form."""
        if not text:
            return ""
        marker = "_" if table is SUBSCRIPT else "^"
        if unicode_ok and all(char in keys for char in text):
            return text.translate(table)
        return f"{marker}({text})"

    def sub(self, text: str) -> str:
        return self.script(text, SUBSCRIPT, _SUB_KEYS, self.scripts == "unicode")

    def sup(self, text: str) -> str:
        return self.script(text, SUPERSCRIPT, _SUP_KEYS, self.scripts == "unicode")

    @staticmethod
    def wrap(text: str) -> str:
        """Parenthesise a fraction side only when it binds more loosely."""
        if not text:
            return ""
        depth = 0
        for char in text:
            if char in "([":
                depth += 1
            elif char in ")]":
                depth -= 1
            elif depth <= 0 and char in LOOSE:
                return f"({text})"
        return text

    @staticmethod
    def accent_mark(node) -> str:
        stored = node.prop("accPr", "chr") or "\u0302"
        return ACCENTS.get(stored, stored)

    @staticmethod
    def bar_mark(node) -> str:
        position = node.prop("barPr", "pos") or "top"
        return UNDERLINE_MARK if position == "bot" else OVERLINE

    # -- linear text: the only thing a math-only slot can hold -------------
    def text(self, node) -> str:
        """Spell a subtree out as linear text."""
        if isinstance(node, str):
            return node
        if not isinstance(node, El):
            return ""
        local = node.local
        if local in ("r", "t", "rPr", "ctrlPr"):
            return raw_text(node)
        if local.endswith("Pr"):
            return ""
        if local in SLOTS:
            return "".join(self.text(child) for child in node.children)
        if local == "sSup":
            return self.text(node.first("e")) + self.sup(self.text(node.first("sup")))
        if local == "sSub":
            return self.text(node.first("e")) + self.sub(self.text(node.first("sub")))
        if local == "sSubSup":
            return (self.text(node.first("e")) + self.sub(self.text(node.first("sub")))
                    + self.sup(self.text(node.first("sup"))))
        if local == "sPre":
            return (self.sub(self.text(node.first("sub"))) + self.sup(self.text(node.first("sup")))
                    + self.text(node.first("e")))
        if local == "f":
            top = self.text(node.first("num"))
            bottom = self.text(node.first("den"))
            return f"{self.wrap(top)}/{self.wrap(bottom)}"
        if local == "rad":
            degree = self.text(node.first("deg"))
            head = "\u221a" + (f"({degree})" if degree else "")
            return head + f"({self.text(node.first('e'))})"
        if local == "d":
            return (self.begin_of(node) + self.text(node.first("e")) + self.end_of(node))
        if local == "nary":
            return (self.operator(node) + self.sub(self.text(node.first("sub")))
                    + self.sup(self.text(node.first("sup"))) + self.text(node.first("e")))
        if local in TRANSPARENT:
            return "".join(self.text(child) for child in node.children)
        if local == "func":
            return self.text(node.first("fName")) + self.text(node.first("e"))
        if local in ("limLow", "limUpp"):
            limit = self.text(node.first("lim"))
            return self.text(node.first("e")) + (self.sub(limit) if local == "limLow"
                                                 else self.sup(limit))
        if local == "acc":
            return self.text(node.first("e")) + self.accent_mark(node)
        if local == "bar":
            return self.text(node.first("e")) + self.bar_mark(node)
        if local == "groupChr":
            return self.text(node.first("e")) + (self.group_char(node) or "")
        if local == "eqArr":
            return "; ".join(self.text(row) for row in node.all("e"))
        if local == "m":
            return "; ".join(" ".join(self.text(cell) for cell in row.all("e"))
                             for row in node.all("mr"))
        return "".join(self.text(child) for child in node.children)

    # -- details read off a construct's properties -------------------------
    def begin_of(self, node) -> str:
        return node.prop("dPr", "begChr") or "("

    def end_of(self, node) -> str:
        return node.prop("dPr", "endChr") or ")"

    def operator(self, node) -> str:
        return node.prop("naryPr", "chr") or "\u2211"

    def group_char(self, node) -> str:
        return node.prop("groupChrPr", "chr") or "\u23df"

    # -- math-only output --------------------------------------------------
    def run_math(self, text: str) -> str:
        """A math run, which is the only leaf a numerator or denominator takes."""
        if not text:
            return ""
        return (f'<{self.m}r><{self.m}t xml:space="preserve">{escape(text)}</{self.m}t>'
                f"</{self.m}r>")

    def math(self, node) -> str:
        """Render a subtree using ``m:r``, ``m:sSup`` and ``m:f`` alone.

        This is the markup a numerator or a denominator accepts: the engine
        reads math elements there and ignores anything else, so a construct it
        does not implement is spelled out as text inside an ``m:r`` rather than
        left as a child it would skip.
        """
        if isinstance(node, str):
            return self.run_math(node)
        if not isinstance(node, El):
            return ""
        local = node.local
        if local in ("r", "t"):
            return self.run_math(raw_text(node))
        if local.endswith("Pr"):
            return ""
        if local in RENDERS and native_ok(node):
            return serialize(node)          # a nested fraction stays stacked
        if local in SLOTS:
            return "".join(self.math(child) for child in node.children)
        if local == "sSup":
            base = self.math(node.first("e"))
            script = self.sup(self.text(node.first("sup")))
            if not script:
                return base
            return (f"<{self.m}sSup><{self.m}e>{base}</{self.m}e>"
                    f"<{self.m}sup>{self.run_math(script)}</{self.m}sup>"
                    f"</{self.m}sSup>")
        # Every other construct is written linearly, which is valid here.
        return self.run_math(self.text(node))

    # -- paragraph-level output -------------------------------------------
    def para(self, node) -> list[tuple[str, str]]:
        """Lower one node into ``("math", inner)`` and ``("run", xml)`` items."""
        if isinstance(node, str) or not isinstance(node, El):
            return []
        if node.local in RENDERS and native_ok(node):
            return [("math", serialize(node))]

        local = node.local
        if local in _DROPPED:
            self.counts[local] = self.counts.get(local, 0) + 1

        if local in SLOTS or local.endswith("Pr"):
            items: list[tuple[str, str]] = []
            for child in node.children:
                items.extend(self.para(child))
            return items

        if local in ("r", "t"):
            return [("math", f'<{self.m}r><{self.m}t xml:space="preserve">'
                             f"{escape(raw_text(node))}</{self.m}t></{self.m}r>")]

        if local == "sSup":
            base, sup = node.first("e"), node.first("sup")
            base_items = self.para(base) if base is not None else []
            script = self.text(sup)
            if len(base_items) == 1 and base_items[0][0] == "math" and script:
                inner = (f"{base_items[0][1]}<{self.m}sup><{self.m}r>"
                         f'<{self.m}t xml:space="preserve">{escape(script)}</{self.m}t>'
                         f"</{self.m}r></{self.m}sup>")
                return [("math", f"<{self.m}sSup>{inner}</{self.m}sSup>")]
            return base_items + [("run", self.run(script, "superscript"))]

        if local == "sSub":
            base, sub = node.first("e"), node.first("sub")
            return ((self.para(base) if base is not None else [])
                    + [("run", self.run(self.text(sub), "subscript"))])

        if local == "sSubSup":
            base, sub, sup = node.first("e"), node.first("sub"), node.first("sup")
            return ((self.para(base) if base is not None else [])
                    + [("run", self.run(self.text(sub), "subscript")),
                       ("run", self.run(self.text(sup), "superscript"))])

        if local == "sPre":
            sub, sup, base = node.first("sub"), node.first("sup"), node.first("e")
            return ([("run", self.run(self.text(sub), "subscript")),
                     ("run", self.run(self.text(sup), "superscript"))]
                    + (self.para(base) if base is not None else []))

        if local == "f":
            num, den = node.first("num"), node.first("den")
            if self.fraction == "native":
                inner = (f"<{self.m}num>{self.math(num)}</{self.m}num>"
                         f"<{self.m}den>{self.math(den)}</{self.m}den>")
                return [("math", f"<{self.m}f>{inner}</{self.m}f>")]
            top, bottom = self.text(num), self.text(den)
            head = "(" if self.wrap(top) != top else ""
            tail = ")" if self.wrap(bottom) != bottom else ""
            return ([("run", self.run(head))]
                    + (self.para(num) if num is not None else [])
                    + [("run", self.run("/"))]
                    + (self.para(den) if den is not None else [])
                    + [("run", self.run(tail))])

        if local == "rad":
            deg, base = node.first("deg"), node.first("e")
            items = [("math", f'<{self.m}r><{self.m}t xml:space="preserve">\u221a'
                              f"</{self.m}t></{self.m}r>")]
            degree = self.text(deg)
            if degree:
                items.append(("run", self.run(degree, "superscript")))
            body = self.para(base) if base is not None else []
            ruled = self.ruled(base) if self.radical == "overline" else ""
            if ruled:
                return items + [("run", self.run(ruled, None, True))]
            return items + [("run", self.run("("))] + body + [("run", self.run(")"))]

        if local == "d":
            return ([("run", self.run(self.begin_of(node)))]
                    + (self.para(node.first("e")) if node.first("e") is not None else [])
                    + [("run", self.run(self.end_of(node)))])

        if local == "nary":
            sub, sup = node.first("sub"), node.first("sup")
            items = [("math", f'<{self.m}r><{self.m}t xml:space="preserve">'
                              f"{escape(self.operator(node))}</{self.m}t></{self.m}r>")]
            if sub is not None:
                items.append(("run", self.run(self.text(sub), "subscript")))
            if sup is not None:
                items.append(("run", self.run(self.text(sup), "superscript")))
            return items + (self.para(node.first("e")) if node.first("e") is not None else [])

        if local in TRANSPARENT:
            items = []
            for child in node.children:
                items.extend(self.para(child))
            return items

        if local == "func":
            return ([("run", self.run(self.text(node.first("fName"))))]
                    + (self.para(node.first("e")) if node.first("e") is not None else []))

        if local in ("limLow", "limUpp"):
            align = "subscript" if local == "limLow" else "superscript"
            return ((self.para(node.first("e")) if node.first("e") is not None else [])
                    + [("run", self.run(self.text(node.first("lim")), align))])

        if local in ("acc", "bar"):
            text = self.text(node.first("e"))
            if text:
                mark = self.accent_mark(node) if local == "acc" else self.bar_mark(node)
                return [("run", self.run(text + mark, None, True))]
            return (self.para(node.first("e")) if node.first("e") is not None else [])

        if local == "groupChr":
            return ((self.para(node.first("e")) if node.first("e") is not None else [])
                    + [("run", self.run(self.group_char(node)))])

        if local == "eqArr":
            items = []
            for position, row in enumerate(node.all("e")):
                if position:
                    items.append(("run", self.run("; ")))
                items.extend(self.para(row))
            return items

        if local == "m":
            items = []
            for position, row in enumerate(node.all("mr")):
                if position:
                    items.append(("run", self.run("; ")))
                for cell in row.all("e"):
                    items.extend(self.para(cell))
            return items

        items = []
        for child in node.children:
            items.extend(self.para(child))
        return items

    def ruled(self, node) -> str:
        """The radicand with a trailing overline, when it can carry one.

        The rule is a spacing character, so it follows the radicand rather than
        spanning it, and only a plain run of glyphs can take one: a stacked
        fraction keeps explicit parentheses instead.
        """
        if node is None:
            return ""
        for child in node.children:
            for element in walk(child):
                if element.local not in ("r", "t", "rPr"):
                    return ""
        text = raw_text(node)
        if not text or not _ATOM.match(text):
            return ""
        return text + OVERLINE

    # -- one equation ------------------------------------------------------
    def equation(self, node) -> str:
        """The replacement markup for one ``m:oMath`` element."""
        if native_ok(node):
            return serialize(node)
        merged: list[tuple[str, str]] = []
        for kind, payload in self.para(node):
            if not payload:
                continue
            if kind == "math" and merged and merged[-1][0] == "math":
                merged[-1] = ("math", merged[-1][1] + payload)
            else:
                merged.append((kind, payload))
        return "".join(self.island(payload) if kind == "math" else payload
                       for kind, payload in merged)


class Names:
    """The prefixes a part uses for the math and WordprocessingML namespaces."""

    __slots__ = ("math", "word")

    def __init__(self, math: str, word: str):
        self.math = math
        self.word = word

    @property
    def math_tag(self) -> str:
        return f"{self.math}:" if self.math else ""

    @property
    def word_tag(self) -> str:
        return f"{self.word}:" if self.word else ""


def namespace_map(text: str, limit: int = 65536) -> dict[str, str]:
    """Prefix -> URI for every declaration in the part's leading markup."""
    mapping: dict[str, str] = {}
    for match in re.finditer(r'xmlns:([\w.\-]+)\s*=\s*"([^"]*)"|xmlns\s*=\s*"([^"]*)"',
                             text[:limit]):
        prefixed, uri, default = match.groups()
        if prefixed:
            mapping.setdefault(prefixed, uri)
        elif default and "" not in mapping:
            mapping[""] = default
    return mapping


# --------------------------------------------------------------------------
# Part level
# --------------------------------------------------------------------------

def names_for(text: str) -> Names | None:
    """The prefixes to emit with, or None when the part cannot hold math."""
    mapping = namespace_map(text)
    if W_NS not in mapping.values():
        return None
    math = next((prefix for prefix, uri in mapping.items() if uri == M_NS), None)
    if math is None:
        return None
    word = next((prefix for prefix, uri in mapping.items() if uri == W_NS), "w")
    return Names(math, word)


def scan(text: str, names: Names) -> dict[str, int]:
    """Count the constructs the engine would drop, without parsing anything."""
    prefix = names.math_tag
    pattern = re.compile(rf"<{prefix}({'|'.join(DROPPED)})(?=[\s/>])")
    found: dict[str, int] = {}
    for match in pattern.finditer(text):
        name = match.group(1)
        found[name] = found.get(name, 0) + 1
    return found


def math_spans(text: str, names: Names) -> list[tuple[int, int]]:
    """Spans of the outermost ``m:oMath`` / ``m:oMathPara`` elements.

    Math never nests inside math, except that an ``m:oMathPara`` holds
    ``m:oMath`` children, so one depth counter finds every outermost element:
    anything opened while another is open is that one's child.
    """
    prefix = names.math_tag
    opener = re.compile(rf"<{prefix}(?:oMath|oMathPara)(?=[\s/>])")
    closer = re.compile(rf"</{prefix}(?:oMath|oMathPara)\s*>")
    spans: list[tuple[int, int]] = []
    depth = 0
    start = 0
    cursor = 0
    while True:
        opening = opener.search(text, cursor)
        closing = closer.search(text, cursor)
        if opening is None and closing is None:
            break
        if closing is None or (opening is not None and opening.start() < closing.start()):
            head = text[opening.start():text.find(">", opening.start()) + 1]
            if depth == 0:
                start = opening.start()
            if not head.endswith("/>"):
                depth += 1
            cursor = opening.end()
        else:
            if depth == 1:
                spans.append((start, closing.end()))
            depth = max(0, depth - 1)
            cursor = closing.end()
    return spans


def inside_run(text: str, position: int, names: Names) -> bool:
    """True when the element starting at `position` sits inside a ``w:r``.

    A run may not contain a run, so an equation that WPS wrapped in one cannot
    receive ordinary runs; such an equation is lowered to math-only output and
    the schema repair pass unwraps it afterwards.
    """
    word = names.word_tag
    opening = text.rfind(f"<{word}r", 0, position)
    if opening < 0:
        return False
    after = text[opening + len(word) + 1:opening + len(word) + 2]
    if after not in (" ", "\t", "\r", "\n", "/", ">"):
        return False                      # <w:rPr>, <w:rFonts>, <w:rsid …>
    return opening > text.rfind(f"</{word}r>", 0, position)


def lower_part(text: str, names: Names, lowering: Lowering) -> tuple[str, int, int, int]:
    """Rewrite every affected equation in one XML part.

    Returns the new text, the equations rewritten, the equations seen and how
    many ``m:oMathPara`` wrappers were dropped.
    """
    spans = math_spans(text, names)
    if not spans:
        return text, 0, 0, 0
    edits: list[tuple[int, int, str]] = []
    unwrapped = 0
    for start, end in spans:
        parsed = parse(text[start:end])
        node = next((child for child in parsed if isinstance(child, El)), None)
        if node is None:
            continue
        if node.local == "oMathPara":
            inner = [child for child in node.children
                     if isinstance(child, El) and child.local == "oMath"]
            if all(native_ok(child) for child in inner):
                continue
            replacement = "".join(lowering.equation(child) for child in inner)
            unwrapped += 1
        else:
            if native_ok(node):
                continue
            replacement = lowering.equation(node)
        if not replacement:
            continue
        if inside_run(text, start, names):
            # Keep the equation's own markup: ordinary runs cannot be nested.
            replacement = (f"<{names.math_tag}oMath>"
                           + "".join(lowering.math(child) for child in node.children)
                           + f"</{names.math_tag}oMath>")
        edits.append((start, end, replacement))
    if not edits:
        return text, 0, len(spans), 0
    edits.sort(key=lambda edit: edit[0])
    pieces: list[str] = []
    cursor = 0
    for start, end, replacement in edits:
        if start < cursor:
            continue
        pieces.append(text[cursor:start])
        pieces.append(replacement)
        cursor = end
    pieces.append(text[cursor:])
    return "".join(pieces), len(edits), len(spans), unwrapped


def lower_document(source: str, destination: str | None, options: dict) -> dict:
    """Scan, and optionally rewrite, every math-bearing part of a DOCX.

    The scan comes first and reads every part, because it is the only way to
    know whether a rewrite is worth writing at all: a document the engine
    already renders correctly must come out of this pass untouched, and the
    caller must not pay for a second conversion to discover that.
    """
    constructs: dict[str, int] = {}
    parts: dict[str, int] = {}
    equations = 0
    lowered = 0
    unwrapped = 0
    warnings: list[str] = []
    fraction = options.get("fraction", "native")
    radical = options.get("radical", "parens")
    scripts = options.get("scripts", "unicode")

    with zipfile.ZipFile(source) as incoming:
        infos = {info.filename: info for info in incoming.infolist()}

        # -- pass one: the parts carrying a construct the engine would drop --
        carried: set[str] = set()
        for name in incoming.namelist():
            if not (name.startswith("word/") and name.endswith(".xml")):
                continue
            text = incoming.read(name).decode("utf-8", errors="replace")
            names = names_for(text)
            if names is None:
                continue
            found = scan(text, names)
            if not found:
                continue
            carried.add(name)
            for key, count in found.items():
                constructs[key] = constructs.get(key, 0) + count
            equations += len(math_spans(text, names))

        # -- pass two: rewrite those parts, copy every other entry verbatim --
        if destination is not None and carried:
            with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as outgoing:
                for name in incoming.namelist():
                    data = incoming.read(name)
                    if name in carried:
                        text = data.decode("utf-8", errors="replace")
                        names = names_for(text)
                        lowering = Lowering(names, fraction=fraction, radical=radical,
                                            scripts=scripts)
                        rewritten, count, _, dropped = lower_part(text, names, lowering)
                        unwrapped += dropped
                        if count:
                            data = rewritten.encode("utf-8")
                            parts[name] = count
                            lowered += count
                    outgoing.writestr(name, data, compress_type=infos[name].compress_type)

    if unwrapped:
        warnings.append(
            f"{unwrapped} display equation(s) lost their m:oMathPara wrapper, so "
            "they take the paragraph's own alignment")
    return {
        "changed": bool(parts),
        "needed": bool(constructs),
        "fraction": fraction,
        "radical": radical,
        "scripts": scripts,
        "equations": equations,
        "lowered": lowered,
        "constructs": dict(sorted(constructs.items(), key=lambda kv: -kv[1])),
        "parts": parts,
        "warnings": warnings,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="dxpdf-omml",
        description="Lower OMML the dxpdf engine cannot lay out into content it renders.",
    )
    parser.add_argument("input", help="source .docx, never modified")
    parser.add_argument("--output", help="path of the lowered .docx")
    parser.add_argument("--check", action="store_true",
                        help="report what would be lowered without writing anything")
    parser.add_argument("--fraction", choices=("native", "linear"), default="native",
                        help="keep fractions stacked and lower only their sides "
                             "(default), or spell them out as num/den so every "
                             "script stays a real subscript")
    parser.add_argument("--radical", choices=("parens", "overline"), default="parens",
                        help="how a root draws its vinculum: a combining overline "
                             "over the radicand (default), or parentheses")
    parser.add_argument("--scripts", choices=("unicode", "brackets"), default="unicode",
                        help="how a script inside a fraction is written: Unicode "
                             "sub/superscript characters where they exist "
                             "(default), or _() / ^()")
    parser.add_argument("--quiet", action="store_true", help="print nothing on success")
    args = parser.parse_args(argv)

    if not args.check and not args.output:
        parser.error("--output is required unless --check is given")

    options = {
        "fraction": args.fraction,
        "radical": args.radical,
        "scripts": args.scripts,
    }
    try:
        result = lower_document(args.input, None if args.check else args.output, options)
    except (zipfile.BadZipFile, OSError, UnicodeDecodeError) as exc:
        sys.stderr.write(f"dxpdf-omml: {exc}\n")
        return 2

    if not args.quiet:
        json.dump(result, sys.stdout, ensure_ascii=False, indent=2)
        sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
