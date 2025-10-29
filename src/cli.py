# """Command-line interface for the other income synchroniser."""

# from __future__ import annotations

# import argparse
# import sys

# from .runner import run_misc_income


# def main(argv: list[str] | None = None) -> int:
#     parser = argparse.ArgumentParser(
#         description="Synchronise misc income between Excel and QuickBooks"
#     )
#     parser.add_argument(
#         "--workbook",
#         required=True,
#         help="Excel workbook containing the other income worksheet",
#     )
#     parser.add_argument("--output", help="Optional JSON output path")

#     args = parser.parse_args(argv)

#     path = run_misc_income("", args.workbook, output_path=args.output)
#     print(f"Report written to {path}")
#     return 0


# if __name__ == "__main__":  # pragma: no cover - manual invocation
#     sys.exit(main())
