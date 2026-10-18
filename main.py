"""Command-line entry point for the weekly sales report.

Example
-------
    python main.py sample_input/sales_data.csv --week 8

That writes output/weekly_report_week8.xlsx.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from report import build_report


def parse_args() -> argparse.Namespace:
    """Read the CSV path, the week to report, and an optional output path."""
    parser = argparse.ArgumentParser(
        description="Build a weekly Excel sales report from a CSV file."
    )
    parser.add_argument(
        "csv_path",
        help="Path to the sales CSV. Generate one with make_sample.py.",
    )
    parser.add_argument(
        "--week",
        type=int,
        required=True,
        help="Week number to highlight, for example 8.",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Workbook path. Default: output/weekly_report_weekN.xlsx",
    )
    return parser.parse_args()


def main() -> int:
    """Build the report. Returns 0 on success and 1 when the inputs are wrong."""
    args = parse_args()
    if args.week < 1:
        print("Error: --week must be 1 or greater.")
        return 1

    output = args.output
    if output is None:
        output = str(Path("output") / f"weekly_report_week{args.week}.xlsx")

    try:
        saved = build_report(args.csv_path, args.week, output)
    except FileNotFoundError as exc:
        print(f"Error: {exc}")
        return 1
    except ValueError as exc:
        print(f"Error: {exc}")
        return 1

    print(f"Wrote {saved}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
