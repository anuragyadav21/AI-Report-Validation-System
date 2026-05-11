#!/usr/bin/env python3
"""Insert rubric columns (E–L) after LLM response (D) on existing workbooks without clearing D."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from openpyxl import load_workbook

from workbook_layout import ensure_rubric_columns_in_workbook

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_IN = ROOT / "prompt-worksheet.xlsx"


def migrate_workbook(src: Path, dest: Path) -> int:
    wb = load_workbook(src)
    try:
        ensure_rubric_columns_in_workbook(wb)
    except ValueError:
        return 1

    wb.save(dest)
    print(f"Saved: {dest}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", type=Path, default=DEFAULT_IN)
    ap.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output path (default: sibling *_with_rubric.xlsx)",
    )
    ap.add_argument(
        "--inplace",
        action="store_true",
        help="Overwrite --input (use with caution)",
    )
    args = ap.parse_args()

    src = args.input
    if not src.is_file():
        print(f"Not found: {src}", file=sys.stderr)
        return 1

    if args.inplace:
        dest = src
    else:
        dest = args.output or src.with_name(
            src.stem + "_with_rubric" + src.suffix
        )

    return migrate_workbook(src, dest)


if __name__ == "__main__":
    raise SystemExit(main())
