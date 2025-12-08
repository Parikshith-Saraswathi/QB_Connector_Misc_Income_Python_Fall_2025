# """Command-line interface for the other income synchroniser."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .runner import run_misc_income


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Synchronise misc income between Excel and QuickBooks"
    )
    parser.add_argument(
        "--workbook",
        required=True,
        help="Excel workbook containing the other income worksheet",
    )
    parser.add_argument(
        "--bank_account",
        required=False,
        help=(
            "When running the packaged exe: the bank account name (string). "
            "When running as Python: optional path to JSON file with bank account info "
            "(defaults to src/input_settings.json)."
        ),
    )
    parser.add_argument("--output", help="Optional JSON output path")

    args = parser.parse_args(argv)

    # Decide how to interpret --bank_account. If running as a frozen exe, the
    # user will pass the bank account name directly. When running as Python the
    # argument is an optional path to the JSON file.
    if getattr(sys, "frozen", False):
        # Running from packaged exe — require a bank account name string
        if not args.bank_account:
            parser.error(
                "--bank_account is required when running the packaged exe and must be the bank account name"
            )
        bank_account_arg = args.bank_account
    else:
        # Running from Python — use provided path or the default JSON
        bank_account_arg = args.bank_account or "src/input_settings.json"

    path = run_misc_income(
        Path(args.workbook),
        bank_account_json=bank_account_arg,
        output_path=args.output,
    )
    print(f"Report written to {path}")
    return 0


if __name__ == "__main__":  # pragma: no cover - manual invocation
    sys.exit(main())
