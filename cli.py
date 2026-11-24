"""Command line: python cli.py report.pdf --to docx xlsx"""
import argparse
import sys
from pathlib import Path

from converter import FORMATS, ConversionError, convert, has_text_layer


def main(argv=None):
    p = argparse.ArgumentParser(description="Convert PDF files to editable formats.")
    p.add_argument("pdf", nargs="+", help="PDF file(s) to convert")
    p.add_argument("--to", nargs="+", default=["docx"], choices=list(FORMATS), metavar="FORMAT",
                   help="one or more of: " + ", ".join(FORMATS) + " (default: docx)")
    p.add_argument("--pages", help="page range, e.g. 1-3,5 (default: all)")
    p.add_argument("--out", help="output folder (default: next to the PDF)")
    p.add_argument("--password", help="password for protected PDFs")
    args = p.parse_args(argv)

    out_dir = Path(args.out) if args.out else None
    if out_dir:
        out_dir.mkdir(parents=True, exist_ok=True)

    failed = 0
    for pdf in map(Path, args.pdf):
        try:
            if not has_text_layer(pdf, args.password):
                print(f"warning: {pdf.name} looks scanned (no text layer) - output will be mostly empty",
                      file=sys.stderr)
        except ConversionError as e:
            print(f"FAILED {pdf}: {e}", file=sys.stderr)
            failed += 1
            continue
        for fmt in args.to:
            target = (out_dir / (pdf.stem + FORMATS[fmt]["ext"])) if out_dir else None
            try:
                out = convert(pdf, fmt, target, args.pages, args.password)
                print(f"OK     {pdf.name} -> {out}")
            except ConversionError as e:
                print(f"FAILED {pdf.name} -> {fmt}: {e}", file=sys.stderr)
                failed += 1
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
