"""CLI entry point: build the premium sales report.

Usage
-----
    python main3.py
    python main3.py --input input/sales.xlsx --output output/report.xlsx
    python main3.py --currency EUR
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from core import (
    DEFAULT_CURRENCY,
    INPUT_FILE,
    OUTPUT_FILE,
    DataValidationError,
    build_report,
    money,
)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate a styled Excel sales dashboard from raw sales data.",
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=INPUT_FILE,
        help=f"input .xlsx file (default: {INPUT_FILE})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=OUTPUT_FILE,
        help=f"output .xlsx file (default: {OUTPUT_FILE})",
    )
    parser.add_argument(
        "--currency",
        default=DEFAULT_CURRENCY,
        help=f"currency symbol used in the report (default: {DEFAULT_CURRENCY})",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    print("Premium Excel Automation Started!")

    if not args.input.exists():
        print(f"ERROR: input file not found: {args.input}", file=sys.stderr)
        print("Hint: run 'python generate_data.py' to create sample data.", file=sys.stderr)
        return 1

    try:
        summary, _ = build_report(args.input, args.output, currency=args.currency)
    except DataValidationError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print("Premium dashboard created successfully!")
    print(f"Total Revenue:  {money(summary.total_revenue, args.currency)}")
    print(f"Total Quantity: {summary.total_quantity:,.0f}")
    print(f"Avg Unit Price: {money(summary.average_unit_price, args.currency)}")
    print(f"Transactions:   {summary.transactions:,}")
    print(f"Best Product:   {summary.best_product} ({money(summary.best_product_revenue, args.currency)})")
    print(f"Best Region:    {summary.best_region} ({money(summary.best_region_revenue, args.currency)})")
    print(f"Output:         {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
