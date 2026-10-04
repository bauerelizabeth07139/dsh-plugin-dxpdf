#!/usr/bin/env python3
"""Command-line front end for the dxpdf DOCX->PDF engine.

The `dxpdf` wheel on PyPI ships only the compiled extension module
(`dxpdf/dxpdf.pyd`) and no console script, so `dxpdf <file.docx>` does not
work out of the box. This shim restores the documented CLI surface on top of
the module's `convert_file` API.

Usage:
    dxpdf input.docx                  # writes input.pdf next to the input
    dxpdf input.docx -o output.pdf    # explicit output path
    dxpdf input.docx --image-dpi 300  # embedded image resolution (default 220)
    dxpdf --version
    dxpdf --help
"""

from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path

PROG = "dxpdf"
DEFAULT_IMAGE_DPI = 220.0
DPI_MIN, DPI_MAX = 1.0, 2400.0


def _module_version() -> str | None:
    """Best-effort version from installed distribution metadata."""
    try:
        from importlib.metadata import version

        return version("dxpdf")
    except Exception:
        return None


def _load_engine():
    """Import the extension module, failing with an actionable message."""
    try:
        import dxpdf  # noqa: PLC0415 - deliberate lazy import
    except Exception as exc:  # pragma: no cover - environment dependent
        sys.stderr.write(
            f"{PROG}: cannot import the dxpdf module ({exc}).\n"
            f"       Install it with: {sys.executable} -m pip install dxpdf\n"
        )
        raise SystemExit(3) from exc

    if not hasattr(dxpdf, "convert_file"):
        sys.stderr.write(
            f"{PROG}: the installed dxpdf module has no convert_file(); "
            f"found {sorted(n for n in dir(dxpdf) if not n.startswith('_'))}\n"
        )
        raise SystemExit(3)
    return dxpdf


def _dpi(value: str) -> float:
    try:
        parsed = float(value)
    except ValueError:
        raise argparse.ArgumentTypeError(f"expected a number, got {value!r}") from None
    if not (DPI_MIN <= parsed <= DPI_MAX):
        raise argparse.ArgumentTypeError(
            f"image_dpi must be between {DPI_MIN:g} and {DPI_MAX:g}, got {parsed:g}"
        )
    return parsed


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=PROG,
        description="Convert a DOCX file to PDF without Office or LibreOffice.",
        epilog="Exits 0 on success, 1 on a conversion failure, 2 on a usage error, "
        "3 when the dxpdf module itself is unavailable.",
    )
    parser.add_argument("input", nargs="?", help="path to the source .docx file")
    parser.add_argument(
        "-o",
        "--output",
        help="path to the PDF to write (default: the input path with a .pdf suffix)",
    )
    parser.add_argument(
        "--image-dpi",
        type=_dpi,
        default=DEFAULT_IMAGE_DPI,
        metavar="N",
        help=f"resolution embedded raster images are downsampled to, "
        f"{DPI_MIN:g}-{DPI_MAX:g} (default: {DEFAULT_IMAGE_DPI:g})",
    )
    parser.add_argument(
        "--version", action="store_true", help="print the dxpdf version and exit"
    )
    parser.add_argument(
        "--quiet", "-q", action="store_true", help="suppress the success summary"
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.version:
        found = _module_version()
        print(f"{PROG} {found}" if found else f"{PROG} (version unknown)")
        return 0

    if not args.input:
        parser.error("the following arguments are required: input")

    source = Path(args.input).expanduser()
    if not source.is_file():
        sys.stderr.write(f"{PROG}: no such file: {source}\n")
        return 1
    if source.suffix.lower() != ".docx":
        sys.stderr.write(
            f"{PROG}: expected a .docx input, got {source.name!r}. "
            f"Only DOCX is supported.\n"
        )
        return 1

    destination = (
        Path(args.output).expanduser() if args.output else source.with_suffix(".pdf")
    )
    if destination.resolve() == source.resolve():
        sys.stderr.write(f"{PROG}: output path would overwrite the input: {source}\n")
        return 1

    if destination.parent and not destination.parent.exists():
        try:
            destination.parent.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            sys.stderr.write(
                f"{PROG}: cannot create output directory {destination.parent}: {exc}\n"
            )
            return 1

    engine = _load_engine()
    started = time.perf_counter()
    try:
        engine.convert_file(str(source), str(destination), image_dpi=args.image_dpi)
    except Exception as exc:
        sys.stderr.write(f"{PROG}: conversion failed: {exc}\n")
        return 1
    elapsed_ms = (time.perf_counter() - started) * 1000

    if not destination.is_file():
        sys.stderr.write(
            f"{PROG}: the engine reported success but wrote no file at {destination}\n"
        )
        return 1

    if not args.quiet:
        size = destination.stat().st_size
        sys.stderr.write(
            f"{source.name} -> {destination} "
            f"({size / 1024:.0f} KiB, {elapsed_ms:.0f} ms, "
            f"image_dpi={args.image_dpi:g})\n"
        )
    else:
        print(destination)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
