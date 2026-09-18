#!/usr/bin/env python3
"""Financial Model Construction Agent - CLI entry point.

Usage:
  python main.py --package inputs/acme.yaml
  python main.py --documents inputs/documents/            # a folder of PDFs
  python main.py --package inputs/acme.yaml --out outputs/models --company "Acme Ltd"

The agent transforms company documents into a professional, historically
populated, formula-driven, source-traceable Excel financial model with a
3-year forecast framework. The analyst owns all forecast and valuation
assumptions; the agent never invents data.
"""

from __future__ import annotations

import argparse
import sys

from modules.pipeline import run


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="Financial Model Construction Agent")
    src = p.add_mutually_exclusive_group(required=True)
    src.add_argument("--package", help="Structured company package (.yaml/.json)")
    src.add_argument("--documents", help="Directory of source PDF documents")
    p.add_argument("--out", dest="output_dir", default=None, help="Output directory")
    p.add_argument("--company", default=None, help="Override company name")
    p.add_argument("--force", action="store_true",
                   help="Continue past a failed gate (records the failure).")
    args = p.parse_args(argv)

    result = run(package=args.package, documents=args.documents,
                 output_dir=args.output_dir, company=args.company, force=args.force)

    print(f"\nCompany: {result.company}")
    for g in result.gate_results:
        print("\n".join("  " + line for line in g.report_lines()))
    if result.stopped_reason:
        print(f"\nPIPELINE STOPPED: {result.stopped_reason}")
        if result.diligence_path:
            print(f"Diligence report: {result.diligence_path}")
        return 1
    print(f"\nWorkbook:        {result.workbook_path}")
    print(f"Build report:    {result.build_report_path}")
    print(f"Diligence:       {result.diligence_path}")
    print(f"Source registry: {result.source_registry_path}")
    print("\nDone. Open the workbook, review historicals/exceptions, then enter assumptions.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
