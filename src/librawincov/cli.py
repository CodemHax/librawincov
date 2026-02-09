"""Command-line interface for librawincov."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from librawincov.converter import convert_docx_to_pdf


def main() -> int:
    """Main entry point for the CLI."""
    parser = argparse.ArgumentParser(
        prog="librawincov",
        description="Convert DOCX files to PDF using LibreOffice",
    )
    parser.add_argument(
        "input",
        type=Path,
        help="Input DOCX file path",
    )
    parser.add_argument(
        "output",
        type=Path,
        help="Output PDF file path",
    )
    parser.add_argument(
        "-v", "--version",
        action="version",
        version="%(prog)s 0.1.0",
    )

    args = parser.parse_args()

    if not args.input.exists():
        print(f"Error: Input file not found: {args.input}", file=sys.stderr)
        return 1

    try:
        convert_docx_to_pdf(args.input, args.output)
        print(f"Successfully converted: {args.input} -> {args.output}")
        return 0
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except RuntimeError as e:
        print(f"Conversion error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
