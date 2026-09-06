import argparse
import os
import sys

from backend.copilot import ExcelCopilot


def main():
    parser = argparse.ArgumentParser(
        description="Excel AI Copilot - Natural Language Spreadsheet Assistant"
    )
    parser.add_argument(
        "filepath",
        nargs="?",
        help="Path to an Excel workbook to open",
    )
    parser.add_argument(
        "-c", "--command",
        help="Run a single command instead of interactive mode",
    )
    parser.add_argument(
        "--sheet",
        help="Sheet name to use (default: active sheet)",
    )

    args = parser.parse_args()

    copilot = ExcelCopilot()

    # If a filepath is given, load the workbook
    if args.filepath:
        if not os.path.exists(args.filepath):
            print(f"Error: File not found: {args.filepath}")
            sys.exit(1)
        try:
            copilot.load_workbook(args.filepath)
        except Exception as e:
            print(f"Error loading workbook: {e}")
            sys.exit(1)

    # If a single command is given, run it and exit
    if args.command:
        result = copilot.process_command(args.command, sheet_name=args.sheet)
        print(result["message"])
        if result["success"] and args.filepath:
            try:
                copilot.excel.save()
                print(f"Workbook saved: {args.filepath}")
            except Exception as e:
                print(f"Note: could not save workbook: {e}")
        sys.exit(0)

    # Otherwise run interactive mode
    copilot.run_interactive(args.filepath)


if __name__ == "__main__":
    main()