from backend.excel.handler import ExcelHandler
from backend.parser.parser import CommandParser, ParserError
from backend.validator.validator import Validator, ValidationError


class HistoryEntry:
    """Represents a single calculation in the history."""

    def __init__(self, command, instruction, result, destination=None):
        self.command = command
        self.instruction = instruction
        self.result = result
        self.destination = destination

    def __repr__(self):
        dest = f" to {self.destination}" if self.destination else ""
        return f"{self.instruction.get('operation')}(...){dest} = {self.result}"


class ExcelCopilot:
    """Main orchestrator that ties parser, validator, and excel together."""

    def __init__(self, filepath=None):
        self.excel = ExcelHandler(filepath)
        self.parser = CommandParser()
        self.validator = Validator(self.excel)
        self.history = []
        self.interactive_mode = False

    def load_workbook(self, filepath):
        """Load a workbook for processing."""
        self.excel.load(filepath)
        return self

    def process_command(self, command, sheet_name=None, allow_overwrite=False):
        """Process a natural language command end-to-end.
        Returns a dict with result, confirmation, and history.
        """
        try:
            # Step 1: Parse into structured instruction
            instruction = self.parser.parse(command)

            # Step 2: Validate
            operation = self.validator.validate_operation(
                instruction.get("operation")
            )
            instruction["operation"] = operation

            # Step 3: Resolve data and execute
            result = None
            destination = None

            if operation in ("SUM", "AVERAGE", "MIN", "MAX", "COUNT"):
                if "column" in instruction:
                    source = instruction["column"]
                    numeric_values, used_range = self.validator.validate_column(
                        source, sheet_name
                    )
                    instruction["used_range"] = used_range
                    instruction["values"] = numeric_values

                    destination = instruction.get("destination")

                    # Execute the calculation
                    instruction_for_execute = dict(instruction)
                    result = self.validator.execute(instruction_for_execute)

                    # Write to destination if specified
                    if destination:
                        self.validator.validate_destination(
                            destination, sheet_name, allow_overwrite
                        )
                        formula = f"={operation}({used_range})"
                        self.excel.write_formula(destination, formula, sheet_name)

                elif "inputs" in instruction:
                    # Operations with explicit cell refs
                    cell_refs = instruction["inputs"]
                    values = []
                    for ref in cell_refs:
                        val = self.excel.get_cell_value(ref, sheet_name)
                        if not isinstance(val, (int, float)):
                            try:
                                val = float(val)
                            except (ValueError, TypeError):
                                continue
                        values.append(val)
                    instruction["values"] = values
                    result = self.validator.execute(instruction)
                    destination = instruction.get("output")

                    if destination:
                        self.validator.validate_destination(
                            destination, sheet_name, allow_overwrite
                        )
                        formula = f"={operation}"
                        for i, ref in enumerate(cell_refs):
                            if i > 0:
                                op_symbol = {
                                    "ADD": "+",
                                    "SUBTRACT": "-",
                                    "MULTIPLY": "*",
                                    "DIVIDE": "/",
                                }.get(operation, "+")
                                formula += op_symbol
                            formula += ref
                        self.excel.write_formula(destination, formula, sheet_name)

            elif operation in ("ADD", "SUBTRACT", "MULTIPLY", "DIVIDE"):
                inputs = instruction.get("inputs", [])
                if len(inputs) >= 2:
                    a = self._get_numeric_value(inputs[0], sheet_name)
                    b = self._get_numeric_value(inputs[1], sheet_name)
                    instruction["a"] = a
                    instruction["b"] = b
                    result = self.validator.execute(instruction)

                    destination = instruction.get("output")
                    if destination:
                        self.validator.validate_destination(
                            destination, sheet_name, allow_overwrite
                        )
                        op_symbol = {
                            "ADD": "+",
                            "SUBTRACT": "-",
                            "MULTIPLY": "*",
                            "DIVIDE": "/",
                        }.get(operation, "+")
                        formula = f"={inputs[0]}{op_symbol}{inputs[1]}"
                        self.excel.write_formula(destination, formula, sheet_name)
                else:
                    raise ParserError(
                        "Arithmetic operations need two values, "
                        "like 'B2 + C2'."
                    )

            elif operation == "PERCENTAGE":
                raise ParserError("Percentage operations not yet implemented.")

            # Step 4: Record history
            entry = HistoryEntry(command, instruction, result, destination)
            self.history.append(entry)

            # Step 5: Build confirmation
            message = self._build_confirmation(instruction, result, destination)

            return {
                "success": True,
                "instruction": instruction,
                "result": result,
                "destination": destination,
                "message": message,
                "history": entry,
            }

        except (ParserError, ValidationError) as e:
            return {
                "success": False,
                "message": e.message,
                "instruction": None,
                "result": None,
                "destination": None,
                "history": None,
            }
        except ZeroDivisionError:
            return {
                "success": False,
                "message": "Cannot divide by zero. Please check your numbers.",
                "instruction": None,
                "result": None,
                "destination": None,
                "history": None,
            }
        except Exception as e:
            return {
                "success": False,
                "message": f"An unexpected error occurred: {e}",
                "instruction": None,
                "result": None,
                "destination": None,
                "history": None,
            }

    def _get_numeric_value(self, ref, sheet_name=None):
        """Get a numeric value from a cell reference."""
        val = self.excel.get_cell_value(ref, sheet_name)
        if isinstance(val, str):
            try:
                return float(val)
            except ValueError:
                raise ValidationError(
                    f"Cell {ref} doesn't contain a number."
                )
        if not isinstance(val, (int, float)):
            raise ValidationError(
                f"Cell {ref} doesn't contain a number."
            )
        return val

    def _build_confirmation(self, instruction, result, destination):
        """Build a confirmation message for the user."""
        op = instruction.get("operation")
        source = instruction.get("column") or instruction.get("source")
        used_range = instruction.get("used_range")

        if source and used_range:
            op_name = op.title()
            location = f"{op_name} of {used_range}"
        else:
            inputs = instruction.get("inputs", [])
            if inputs:
                op_symbol = {
                    "ADD": "+",
                    "SUBTRACT": "-",
                    "MULTIPLY": "*",
                    "DIVIDE": "/",
                }.get(op, "+")
                location = f" {op_symbol} ".join(inputs)
            else:
                location = "calculation"

        if destination:
            return (
                f"Done. I calculated {location} = {result} "
                f"and placed the result in {destination}."
            )
        else:
            return f"Done. The result of {location} is {result}."

    def run_interactive(self, filepath=None):
        """Run an interactive command loop."""
        if filepath:
            self.load_workbook(filepath)

        print("=" * 60)
        print("Excel AI Copilot")
        print("Type a command like 'Column D ka sum karo'.")
        print("Type 'history' to see past calculations.")
        print("Type 'exit' or 'quit' to leave.")
        print("=" * 60)

        while True:
            try:
                command = input("\n> ").strip()
            except (EOFError, KeyboardInterrupt):
                print("\nGoodbye!")
                break

            if not command:
                continue

            if command.lower() in ("exit", "quit", "bye"):
                print("Goodbye!")
                break

            if command.lower() in ("history", "hist"):
                self._print_history()
                continue

            result = self.process_command(command)
            self._print_result(result)

            if result["success"] and self.excel.filepath:
                try:
                    self.excel.save()
                except Exception as e:
                    print(f"Note: could not save workbook: {e}")

    def _print_result(self, result):
        """Print a result nicely."""
        print("-" * 40)
        print(result["message"])
        if not result["success"]:
            print("Try again with a clearer command.")
        print("-" * 40)

    def _print_history(self):
        """Print the calculation history."""
        if not self.history:
            print("No calculations yet.")
            return

        print("\nCalculation History")
        print("-" * 40)
        for entry in self.history:
            dest = f" to {entry.destination}" if entry.destination else ""
            print(f"{entry.instruction.get('operation')}(...){dest} = {entry.result}")