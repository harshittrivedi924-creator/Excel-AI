import argparse
import os
import sys

from backend.copilot import ExcelCopilot
from backend.voice.listener import VoiceError, VoiceListener


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
        "-c",
        "--command",
        help="Run a single command instead of interactive mode",
    )
    parser.add_argument(
        "--sheet",
        help="Sheet name to use (default: active sheet)",
    )
    parser.add_argument(
        "--voice",
        action="store_true",
        help="Use voice input in interactive mode",
    )
    parser.add_argument(
        "--serve",
        action="store_true",
        help="Start the web UI instead of the terminal UI",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=5000,
        help="Port for the web UI (default: 5000)",
    )

    args = parser.parse_args()

    # Web UI mode
    if args.serve:
        from backend.api.app import create_app

        app_file = args.filepath
        app = create_app(initial_filepath=app_file)
        print(f"Excel AI Copilot web UI running on http://localhost:{args.port}")
        app.run(host="127.0.0.1", port=args.port, debug=False)
        return

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

    # Voice mode: wrap the interactive loop with speech-to-text input
    if args.voice:
        listener = VoiceListener()
        run_voice_loop(copilot, listener, args.filepath)
        return

    # Otherwise run interactive mode
    copilot.run_interactive(args.filepath)


def run_voice_loop(copilot, listener, filepath=None):
    """Interactive loop that takes commands with the microphone."""
    if filepath:
        copilot.load_workbook(filepath)

    print("=" * 60)
    print("Excel AI Copilot (Voice Mode)")
    print("Speak a command like 'Column D ka sum karo'.")
    print("Say 'history' to see past calculations, 'exit' to quit.")
    print("=" * 60)

    while True:
        print("\nListening...")
        try:
            command = listener.listen()
        except VoiceError as e:
            print(f"Voice error: {e.message}")
            continue
        except KeyboardInterrupt:
            print("\nGoodbye!")
            break

        print(f"You said: {command}")

        if not command:
            continue

        if command.lower() in ("exit", "quit", "bye"):
            print("Goodbye!")
            break

        if command.lower() in ("history", "hist"):
            copilot._print_history()
            continue

        result = copilot.process_interactive(command)
        if result is None:
            print("Goodbye!")
            break

        copilot._print_result(result)

        if result["success"] and copilot.excel.filepath:
            try:
                copilot.excel.save()
            except Exception as e:
                print(f"Note: could not save workbook: {e}")


if __name__ == "__main__":
    main()
