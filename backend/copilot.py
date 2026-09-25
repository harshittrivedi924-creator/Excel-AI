import re

from backend.excel.handler import ExcelHandler
from backend.parser.parser import CommandParser, ParserError
from backend.validator.validator import ValidationError, Validator


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

    BINARY_OPERATIONS = ("ADD", "SUBTRACT", "MULTIPLY", "DIVIDE")
    AGGREGATE_OPERATIONS = ("SUM", "AVERAGE", "MIN", "MAX", "COUNT")
    OP_SYMBOLS = {
        "ADD": "+",
        "SUBTRACT": "-",
        "MULTIPLY": "*",
        "DIVIDE": "/",
    }

    def __init__(self, filepath=None):
        self.excel = ExcelHandler(filepath)
        self.parser = CommandParser()
        self.validator = Validator(self.excel)
        self.history = []

    def load_workbook(self, filepath):
        """Load a workbook for processing."""
        self.excel.load(filepath)
        return self

    def process_command(self, command, sheet_name=None, allow_overwrite=False):
        """Process a natural language command end-to-end.
        Returns a dict with result, confirmation, and history.
        """
        try:
            # Step 1: Refresh the parser's view of the workbook's column
            # headers so arbitrary names (Salary, Orders, ...) are recognized.
            if self.excel.workbook:
                self.parser.set_headers(self.excel.get_headers(sheet_name))

            # Step 2: Parse into structured instruction
            instruction = self.parser.parse(command)

            # Step 3: Validate
            operation = self.validator.validate_operation(instruction.get("operation"))
            instruction["operation"] = operation

            # Step 3: Resolve data and execute
            result = None
            destination = None
            written_count = 0

            if operation in self.AGGREGATE_OPERATIONS:
                result = self._handle_aggregate(instruction, sheet_name, allow_overwrite)
                destination = instruction.get("destination")

            elif operation in (*self.BINARY_OPERATIONS, "DIFFERENCE", "GROWTH"):
                result, destination, written_count = self._handle_binary(
                    instruction, sheet_name, allow_overwrite
                )

            elif operation == "PERCENTAGE":
                result, destination, written_count = self._handle_percentage(
                    instruction, sheet_name, allow_overwrite
                )

            elif operation in ("SUMIF", "COUNTIF", "AVERAGEIF"):
                result, destination, written_count = self._handle_conditional(
                    instruction, sheet_name, allow_overwrite
                )

            elif operation in (
                "SORT",
                "DEDUPE",
                "FILTER",
                "FIND_EMPTY",
                "STANDARDIZE_DATES",
                "STANDARDIZE_NAMES",
                "DETECT_INVALID",
            ):
                result, written_count = self._handle_data_operation(
                    instruction, sheet_name, allow_overwrite
                )

            elif operation == "CHART":
                result, written_count = self._handle_chart(instruction, sheet_name)

            elif operation == "ANALYZE":
                result, written_count = self._handle_analyze(instruction, sheet_name)

            # Step 4: Record history
            entry = HistoryEntry(command, instruction, result, destination)
            self.history.append(entry)

            # Step 5: Build confirmation
            message = self._build_confirmation(instruction, result, destination, written_count)

            return {
                "success": True,
                "instruction": instruction,
                "result": result,
                "destination": destination,
                "written_count": written_count,
                "message": message,
                "history": entry,
            }

        except (ParserError, ValidationError) as e:
            return self._error(e.message)
        except ZeroDivisionError:
            return self._error("Cannot divide by zero. Please check your numbers.")
        except Exception as e:
            return self._error(f"An unexpected error occurred: {e}")

    # ------------------------------------------------------------------ #
    # Operation handlers
    # ------------------------------------------------------------------ #

    def _handle_aggregate(self, instruction, sheet_name, allow_overwrite=False):
        """Handle SUM/AVERAGE/MIN/MAX/COUNT on a column, named column,
        cell inputs, or whole sheet. Returns the computed result.
        """
        operation = instruction["operation"]

        if instruction.get("whole_sheet"):
            values, used_range = self._get_whole_sheet_values(sheet_name)
            if not values:
                raise ValidationError("There are no numeric values in this sheet.")
            instruction["used_range"] = used_range
            instruction["values"] = values
            result = self.validator.execute(instruction)
            self._maybe_write_formula(
                operation,
                used_range,
                instruction.get("destination"),
                sheet_name,
                instruction,
                allow_overwrite,
            )
            return result

        # Named source: "Sales ka total karo"
        if instruction.get("named_source"):
            column = self._resolve_named_column(instruction["named_source"], sheet_name)
            numeric_values, used_range = self.validator.validate_column(column, sheet_name)
            instruction["column"] = column
            instruction["used_range"] = used_range
            instruction["values"] = numeric_values
            result = self.validator.execute(instruction)
            self._maybe_write_formula(
                operation,
                used_range,
                instruction.get("destination"),
                sheet_name,
                instruction,
                allow_overwrite,
            )
            return result

        # Explicit column
        if instruction.get("column"):
            numeric_values, used_range = self.validator.validate_column(
                instruction["column"], sheet_name
            )
            instruction["used_range"] = used_range
            instruction["values"] = numeric_values
            result = self.validator.execute(instruction)
            self._maybe_write_formula(
                operation,
                used_range,
                instruction.get("destination"),
                sheet_name,
                instruction,
                allow_overwrite,
            )
            return result

        # Cell inputs
        if instruction.get("inputs"):
            values = []
            for ref in instruction["inputs"]:
                val = self._read_operand_value(ref, sheet_name)
                if isinstance(val, (int, float)):
                    values.append(val)
            instruction["values"] = values
            result = self.validator.execute(instruction)
            self._write_binary_formula_from_inputs(
                instruction["inputs"],
                instruction.get("output"),
                instruction.get("operation"),
                sheet_name,
                instruction,
                allow_overwrite,
            )
            return result

        raise ParserError(
            "I couldn't figure out which data to use. Specify a column, "
            "a name like 'Revenue', or cell references."
        )

    def _handle_binary(self, instruction, sheet_name, allow_overwrite):
        """Handle ADD/SUBTRACT/MULTIPLY/DIVIDE/DIFFERENCE/GROWTH."""
        operation = instruction["operation"]

        # Named-column binary operation (element-wise or aggregate)
        if instruction.get("named_inputs"):
            return self._handle_named_binary(instruction, sheet_name, allow_overwrite)

        # Growth with cells: "A2 se B2 kitni badhi" -> new=B2, old=A2
        if operation == "GROWTH":
            inputs = instruction["inputs"]
            if len(inputs) != 2:
                raise ParserError("Growth needs two values, like 'B2 vs A2'.")
            new_val = self._read_operand_value(inputs[0], sheet_name)
            old_val = self._read_operand_value(inputs[1], sheet_name)
            instruction["new"], instruction["old"] = new_val, old_val
            result = self.validator.execute(instruction)
            self._write_operand_formula(
                inputs[0],
                inputs[1],
                "subtract_symbol_for_growth",
                instruction.get("output"),
                sheet_name,
                instruction,
            )
            return result, instruction.get("output"), 0

        # Cell-based binary: read operand values and compute
        inputs = instruction["inputs"]
        if len(inputs) != 2:
            raise ParserError(f"I need two values for {operation}, like 'B2 + C2'.")
        a = self._read_operand_value(inputs[0], sheet_name)
        b = self._read_operand_value(inputs[1], sheet_name)

        if operation == "DIFFERENCE":
            instruction["a"], instruction["b"] = a, b
        elif operation in self.BINARY_OPERATIONS:
            instruction["a"], instruction["b"] = a, b
        else:
            raise ParserError(f"Operation {operation} requires named data.")

        result = self.validator.execute(instruction)
        destination = instruction.get("output")
        written_count = 0

        # Write formula when a destination is specified and inputs are cells
        if destination and self._is_cell_ref(inputs[0]):
            self.validator.validate_destination(destination, sheet_name, allow_overwrite)
            if operation == "DIFFERENCE":
                formula = f"=ABS({inputs[0]}-{inputs[1]})"
            elif operation in self.BINARY_OPERATIONS:
                formula = f"={inputs[0]}{self.OP_SYMBOLS[operation]}{inputs[1]}"
            self.excel.write_formula(destination, formula, sheet_name)
            instruction["formula"] = formula
            written_count = 1

        return result, destination, written_count

    def _handle_named_binary(self, instruction, sheet_name, allow_overwrite):
        """Handle binary operations on named columns.

        For ADD/SUBTRACT/MULTIPLY/DIVIDE with a named output column, write
        row-by-row formulas into the output column (element-wise).
        For GROWTH, compare column totals.
        """
        operation = instruction["operation"]
        named_inputs = instruction["named_inputs"]

        if operation == "GROWTH":
            if len(named_inputs) < 2:
                raise ParserError("Growth needs two named columns, like 'February vs January'.")
            try:
                new_col = self._resolve_named_column(named_inputs[0], sheet_name)
                old_col = self._resolve_named_column(named_inputs[1], sheet_name)
            except ValidationError:
                # Month names are not column headers here, which means the
                # workbook stores months as row values (month-per-row layout).
                raise ValidationError(
                    f"I couldn't find a column named '{named_inputs[0]}' in this sheet. "
                    f"This workbook stores months as rows (e.g. a 'Month' column). "
                    f"Month-over-month growth currently works when months are column "
                    f"headers. Try a cell-based version instead, like 'B3 se B2 kitni "
                    f"badhi'."
                ) from None
            new_total, _, _ = self.excel.get_column_values(new_col, sheet_name)
            old_total, _, _ = self.excel.get_column_values(old_col, sheet_name)
            instruction["new"] = sum(new_total)
            instruction["old"] = sum(old_total)
            result = self.validator.execute(instruction)
            instruction["used_range"] = (
                f"{self.excel.get_column_letter(self.excel.get_column_number(new_col))}"
                f" vs {self.excel.get_column_letter(self.excel.get_column_number(old_col))}"
            )
            return result, None, 0

        # Element-wise for the other binary ops
        if len(named_inputs) < 2:
            raise ParserError(f"{operation} needs two named columns, like 'Revenue and Expense'.")

        in1 = self._resolve_named_column(named_inputs[0], sheet_name)
        in2 = self._resolve_named_column(named_inputs[1], sheet_name)

        # Determine the data range (union of both columns)
        values1, max1, min1 = self.excel.get_column_values(in1, sheet_name)
        values2, max2, min2 = self.excel.get_column_values(in2, sheet_name)
        start_row = min(min1, min2) if (min1 and min2) else (min1 or min2)
        end_row = max(max1, max2)
        if not start_row or not end_row:
            raise ValidationError("I couldn't find numeric data for the named columns.")

        # Determine output column
        named_output = instruction.get("named_output")
        if named_output:
            out_col = self._resolve_named_column(named_output, sheet_name, create=True)
            instruction["output_column"] = out_col
        elif instruction.get("output"):
            out_col = instruction["output"][:1]
        else:
            raise ParserError("Please specify an output column, like 'profit calculate kar do'.")

        count = 0
        self._validate_column_destination(out_col, start_row, end_row, sheet_name, allow_overwrite)
        for row in range(start_row, end_row + 1):
            if operation in self.BINARY_OPERATIONS:
                formula = f"={in1}{row}{self.OP_SYMBOLS[operation]}{in2}{row}"
            else:
                formula = f"=ABS({in1}{row}-{in2}{row})"
            self.excel.write_formula(f"{out_col}{row}", formula, sheet_name)
            count += 1

        instruction["row_count"] = count
        instruction["out_col"] = out_col
        return count, None, count

    def _validate_column_destination(
        self, out_col, start_row, end_row, sheet_name, allow_overwrite
    ):
        """Guard a row-by-row column write before anything is written.

        A single-cell check is not enough here: the write touches every cell
        of the column range, so any occupied cell must raise the overwrite
        prompt instead of being silently replaced.
        """
        if allow_overwrite:
            return
        for row in range(start_row, end_row + 1):
            self.validator.validate_destination(f"{out_col}{row}", sheet_name, allow_overwrite)

    def _handle_percentage(self, instruction, sheet_name, allow_overwrite):
        """Handle percentage operations.

        Percentages are computed as each value's share of the column total:
          - with a destination column: row-by-row formulas are written
            (e.g. ``=B2/6600*100``);
          - without a destination: the per-row percentages are computed and
            returned so the command never silently reports the column total.

        A source column is required; if its total is zero percentages cannot
        be computed and a clear error is raised.
        """
        named_source = instruction.get("named_source")
        column = None
        if named_source:
            column = self._resolve_named_column(named_source, sheet_name)
        elif instruction.get("column") or instruction.get("source"):
            column = self._resolve_named_column(
                instruction.get("column") or instruction.get("source"), sheet_name
            )
        if not column:
            raise ParserError(
                "Please specify which column to calculate the percentage of, "
                "like 'Sales ka percentage karo'."
            )

        source_label = named_source if named_source else column
        values, max_row, min_row = self.excel.get_column_values(column, sheet_name)
        if not values:
            raise ValidationError(f"There are no numeric values in the '{source_label}' column.")
        total = sum(values)
        if total == 0:
            raise ValidationError(
                f"The total of the '{source_label}' column is 0, so "
                f"percentages cannot be calculated."
            )

        # A destination cell/column implies row-by-row share-of-total formulas.
        out_col = (
            instruction.get("output_column")
            or instruction.get("output")
            or instruction.get("destination")
        )
        if out_col:
            out_col = out_col[:1]
            self._validate_column_destination(
                out_col, min_row, max_row, sheet_name, allow_overwrite
            )
            for row in range(min_row, max_row + 1):
                formula = f"={column}{row}/{total}*100"
                self.excel.write_formula(f"{out_col}{row}", formula, sheet_name)
            instruction["total"] = total
            instruction["output_column"] = out_col
            return total, None, (max_row - min_row + 1)

        # No destination: report each value's share of the column total.
        percentages = [round(v / total * 100, 2) for v in values]
        instruction["total"] = total
        instruction["values"] = values
        instruction["percentages"] = percentages
        return percentages, None, 0

    def _handle_conditional(self, instruction, sheet_name=None, allow_overwrite=False):
        """Handle SUMIF / COUNTIF / AVERAGEIF.

        The instruction carries:
          - sum_range:   column to aggregate (letter or name)
          - criteria:    list of (column, operator, value)
        Evaluates all criteria against the data rows and aggregates the
        matching cells of the sum range.
        """
        operation = instruction["operation"]
        sum_range = instruction.get("sum_range")
        criteria = instruction.get("criteria") or []

        if not self.excel.workbook:
            raise ValidationError("No workbook loaded.")
        if not sum_range:
            raise ParserError("Please specify which column to calculate.")
        if not criteria:
            raise ParserError("I couldn't find any conditions for the calculation.")

        sheet = self.excel.get_sheet(sheet_name)
        rows = self.excel._read_rows(sheet)
        if not rows:
            raise ValidationError("This sheet has no data.")
        data = rows[1:]

        sum_col = self._resolve_named_column(sum_range, sheet_name)
        sum_idx = self.excel.get_column_number(sum_col)

        # Resolve criteria columns and values, building a combined mask.
        mask = [True] * len(data)
        used_letters = []
        for crit in criteria:
            crit_col, operator, value = crit
            if crit_col is None:
                continue
            crit_letter = self._resolve_named_column(crit_col, sheet_name)
            used_letters.append(crit_letter)
            idx = self.excel.get_column_number(crit_letter)

            if isinstance(value, str):
                try:
                    value = float(value)
                except ValueError:
                    value = value.lower()

            for i, row in enumerate(data):
                cell_value = row[idx - 1] if idx <= len(row) else None
                cv = self._coerce(cell_value)
                numeric = isinstance(cv, (int, float))
                if operator == ">":
                    ok = numeric and cv > value
                elif operator == "<":
                    ok = numeric and cv < value
                elif operator == ">=":
                    ok = numeric and cv >= value
                elif operator == "<=":
                    ok = numeric and cv <= value
                elif operator == "==":
                    ok = str(cv).lower() == str(value).lower()
                else:
                    ok = False
                mask[i] = mask[i] and ok

        # Aggregate matching values from the sum range.
        values = []
        for i, row in enumerate(data):
            if mask[i]:
                cell_value = row[sum_idx - 1] if sum_idx <= len(row) else None
                cv = self._coerce(cell_value)
                if isinstance(cv, (int, float)):
                    values.append(cv)

        matched = sum(1 for keep in mask if keep)

        from backend.calculator.engine import CalculationEngine

        if operation == "SUMIF":
            result = CalculationEngine.sum(values)
        elif operation == "AVERAGEIF":
            result = CalculationEngine.average(values)
        elif operation == "COUNTIF":
            result = CalculationEngine.count_if(mask)
        else:
            raise ParserError(f"Unsupported conditional operation: {operation}")

        instruction["values"] = values
        instruction["mask"] = mask
        instruction["criteria_letters"] = used_letters
        instruction["sum_range_letter"] = sum_col
        instruction["matched"] = matched

        destination = instruction.get("destination")
        if destination:
            self.validator.validate_destination(destination, sheet_name, allow_overwrite)
            crit = criteria[0]
            crit_letter = self._resolve_named_column(crit[0], sheet_name)
            if operation == "COUNTIF":
                formula = f"=COUNTIF({crit_letter}:{crit_letter},{self._formula_criteria(crit)})"
            else:
                formula = (
                    f"={operation}({crit_letter}:{crit_letter},"
                    f"{self._formula_criteria(crit)},"
                    f"{sum_col}:{sum_col})"
                )
            self.excel.write_formula(destination, formula, sheet_name)
            instruction["formula"] = formula

        return result, destination, matched

    @staticmethod
    def _formula_criteria(cond):
        """Render a parsed condition into an Excel criteria argument."""
        column, operator, value = cond
        sym = {"==": "=", ">": ">", "<": "<", ">=": ">=", "<=": "<="}[operator]
        if isinstance(value, str):
            return f'"{value}"'
        return f"{sym}{value}"

    @staticmethod
    def _coerce(v):
        """Coerce a cell value for comparisons."""
        if isinstance(v, str):
            try:
                return float(v)
            except ValueError:
                return v.lower()
        return v

    # ------------------------------------------------------------------ #
    # Helpers
    # ------------------------------------------------------------------ #

    def _handle_data_operation(self, instruction, sheet_name=None, allow_overwrite=False):
        """Handle SORT / DEDUPE / FILTER / FIND_EMPTY and data cleaning
        operations. Returns (result, count)."""
        operation = instruction["operation"]
        column = instruction.get("column")

        # Keep the parsed source (name or letter) for friendly confirmations.
        instruction["source_name"] = column

        # Resolve named column letters
        col_letter = None
        if column:
            col_letter = self._resolve_named_column(column, sheet_name)

        if operation == "SORT":
            ascending = not bool(instruction.get("descending"))
            if col_letter:
                count = self.excel.sort_sheet(
                    col_letter, ascending=ascending, sheet_name=sheet_name
                )
                return {"column": col_letter, "rows": count}, count
            # Sort by the first numeric-friendly column
            sheet = self.excel.get_sheet(sheet_name)
            target = 1
            for col in range(1, sheet.max_column + 1):
                ctx = self.excel.get_column_letter(col)
                vals, _, _ = self.excel.get_column_values(ctx, sheet_name)
                if vals:
                    target = col
                    break
            col_letter = self.excel.get_column_letter(target)
            count = self.excel.sort_sheet(col_letter, ascending=ascending, sheet_name=sheet_name)
            return {"column": col_letter, "rows": count}, count

        elif operation == "DEDUPE":
            count, removed_rows = self.excel.remove_duplicates(sheet_name)
            return {"removed": count, "preview": removed_rows[:3]}, count

        elif operation == "STANDARDIZE_DATES":
            col_letter = col_letter or None
            converted, total = self.excel.standardize_dates(col_letter, sheet_name)
            return {"converted": converted, "total": total}, converted

        elif operation == "STANDARDIZE_NAMES":
            converted, checked = self.excel.standardize_names(col_letter, sheet_name)
            return {"converted": converted, "checked": checked}, converted

        elif operation == "DETECT_INVALID":
            issues, candidates = self.excel.detect_invalid_values(col_letter, sheet_name)
            return {"count": len(issues), "issues": issues[:50]}, len(issues)

        elif operation == "FILTER":
            operator = instruction.get("operator")
            value = instruction.get("value")
            if not col_letter or operator is None or value is None:
                raise ParserError("I need a filter condition like 'Revenue 1000 se zyada'.")
            matched, total = self.excel.filter_rows(col_letter, operator, value, sheet_name)
            return {
                "matched": matched,
                "total": total,
                "column": col_letter,
                "value": value,
            }, matched

        elif operation == "FIND_EMPTY":
            empty = self.excel.find_empty_cells(col_letter, sheet_name)
            return {"empty_count": len(empty), "cells": empty[:50]}, len(empty)

        raise ParserError(f"Unsupported data operation: {operation}")

    def _handle_chart(self, instruction, sheet_name=None):
        """Handle CHART."""
        column = instruction.get("column")
        if not column:
            raise ParserError("Please specify a column to chart, like 'Sales ka chart bana do'.")
        col_letter = self._resolve_named_column(column, sheet_name)
        chart_type = "line" if instruction.get("line_chart") else "column"
        try:
            info = self.excel.create_chart(col_letter, sheet_name, chart_type)
        except ValueError as e:
            raise ValidationError(str(e)) from e
        return {"chart": info["title"], "source": info["source"]}, info["rows"]

    def _handle_analyze(self, instruction, sheet_name=None):
        """Handle ANALYZE. Returns a structured report + writes a summary sheet."""
        report = self.excel.analyze_workbook(sheet_name)

        summary = []
        total_insights = 0
        for sheet_title, data in report.items():
            metrics = data["metrics"]
            trends = data["trends"]
            anomalies = data.get("anomalies", [])
            summary.append(f"Sheet: {sheet_title}")
            if not metrics and not trends and not anomalies:
                summary.append("  No numeric data to analyze.")
                continue
            for m in metrics:
                summary.append(
                    f"  {m['column']}: total {m['total']}, average "
                    f"{m['average']}, max {m['max']}, min {m['min']} "
                    f"({m['count']} values)"
                )
                total_insights += 1
            for t in trends:
                summary.append(f"  Trend: {t}")
            for label, value, mean, stdev, _i in anomalies:
                summary.append(
                    f"  Anomaly: {label} {value} is "
                    f"{abs(value - mean) / stdev:.1f} std devs from the "
                    f"average ({mean:.2f})"
                )
                total_insights += 1

        self._write_analysis_sheet("\n".join(summary), sheet_name)
        return {
            "sheets": list(report.keys()),
            "insights": total_insights,
            "lines": summary,
        }, total_insights

    def _write_analysis_sheet(self, text, sheet_name=None, max_cols=6):
        """Write the analysis text into an 'Analysis Report' sheet."""
        if "Analysis Report" in self.excel.workbook.sheetnames:
            del self.excel.workbook["Analysis Report"]
        ws = self.excel.workbook.create_sheet("Analysis Report")
        lines = text.splitlines()
        for i, line in enumerate(lines[:max_cols]):
            ws.cell(row=i + 1, column=1).value = line

    def _resolve_named_column(self, name, sheet_name=None, create=False):
        """Resolve a name to a column letter via the header row.
        A single column letter (e.g. 'B') is used directly.
        """
        if isinstance(name, str) and len(name) == 1 and name.isalpha():
            column = name.upper()
            # Verify it exists when resolving an existing column
            if not create:
                sheet = self.excel.get_sheet(sheet_name)
                if self.excel.get_column_number(column) > sheet.max_column:
                    raise ValidationError(f"I couldn't find Column {column} in this sheet.")
            return column

        column = self.excel.get_column_by_name(name, sheet_name)
        if column is None and create:
            # Add a new column with the given name
            sheet = self.excel.get_sheet(sheet_name)
            col_idx = sheet.max_column + 1
            from openpyxl.utils import get_column_letter

            column = get_column_letter(col_idx)
            sheet.cell(row=1, column=col_idx).value = name.capitalize()
        if column is None:
            raise ValidationError(f"I couldn't find a column named '{name}' in this sheet.")
        return column

    def _get_whole_sheet_values(self, sheet_name=None):
        """Get all numeric values in the sheet, skipping the header row."""
        used_range = self.excel.get_used_range()
        if used_range is None:
            raise ValidationError("This sheet is empty.")
        sheet = self.excel.get_sheet(sheet_name)
        values = []
        for row in sheet.iter_rows():
            if row[0].row == 1:
                continue
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    values.append(cell.value)
                elif isinstance(cell.value, str):
                    try:
                        values.append(float(cell.value))
                    except ValueError:
                        pass
        return values, used_range

    def _maybe_write_formula(
        self, operation, used_range, destination, sheet_name, instruction, allow_overwrite=False
    ):
        """Write an aggregate formula to a destination cell if requested."""
        if destination:
            self.validator.validate_destination(destination, sheet_name, allow_overwrite)
            formula = f"={operation}({used_range})"
            self.excel.write_formula(destination, formula, sheet_name)
            instruction["formula"] = formula

    def _write_binary_formula_from_inputs(
        self, inputs, output, operation, sheet_name, instruction, allow_overwrite=False
    ):
        """Write a formula for an aggregate operation over explicit cells."""
        if output and self._is_cell_ref(inputs[0]) and self._is_cell_ref(inputs[1]):
            if operation == "SUM":
                formula = f"=SUM({inputs[0]}:{inputs[1]})"
            else:
                formula = f"={operation}({inputs[0]}:{inputs[1]})"
            self.validator.validate_destination(output, sheet_name, allow_overwrite)
            self.excel.write_formula(output, formula, sheet_name)
            instruction["formula"] = formula

    def _write_operand_formula(self, a, b, symbol, destination, sheet_name, instruction):
        """Write a growth/difference formula between two cell refs."""
        if destination and self._is_cell_ref(a) and self._is_cell_ref(b):
            if symbol == "subtract_symbol_for_growth":
                formula = f"=({a}-{b})/{b}*100"
            self.validator.validate_destination(destination, sheet_name, True)
            self.excel.write_formula(destination, formula, sheet_name)
            instruction["formula"] = formula

    def _read_operand_value(self, ref, sheet_name=None):
        """Read a value that may be a cell reference or a numeric literal."""
        if isinstance(ref, (int, float)):
            return ref
        if not isinstance(ref, str):
            raise ValidationError(f"Invalid operand: {ref}")
        if self._is_cell_ref(ref):
            return self._get_numeric_value(ref, sheet_name)
        try:
            return float(ref)
        except ValueError:
            raise ValidationError(f"I couldn't understand the value '{ref}'.") from None

    def _is_cell_ref(self, ref):
        """Check if a string looks like a cell reference (e.g. B2)."""
        return isinstance(ref, str) and bool(re.fullmatch(r"[A-Z]{1,2}\d+", ref))

    def _get_numeric_value(self, ref, sheet_name=None):
        """Get a numeric value from a cell reference."""
        val = self.excel.get_cell_value(ref, sheet_name)
        if isinstance(val, str):
            try:
                return float(val)
            except ValueError:
                raise ValidationError(f"Cell {ref} doesn't contain a number.") from None
        if not isinstance(val, (int, float)):
            raise ValidationError(f"Cell {ref} doesn't contain a number.")
        return val

    def _error(self, message):
        """Build an error result dict."""
        return {
            "success": False,
            "message": message,
            "instruction": None,
            "result": None,
            "destination": None,
            "written_count": 0,
            "history": None,
        }

    def _build_confirmation(self, instruction, result, destination, written_count=0):
        """Build a confirmation message for the user."""
        op = instruction.get("operation")

        # Data operation messages
        if op == "SORT":
            col = result.get("column") if isinstance(result, dict) else None
            source = self._display_source(instruction.get("source_name") or col)
            return f"Done. I sorted {written_count} rows" + (
                f" by column {source}." if source else "."
            )
        if op == "DEDUPE":
            if not written_count:
                return "No duplicate rows were found."
            r = result if isinstance(result, dict) else {}
            preview = r.get("preview", [])
            preview_text = ""
            if preview:
                sample = " | ".join(", ".join(str(v) for v in row) for row in preview[:2])
                preview_text = f" Example removed rows: {sample}."
            return f"Done. I removed {written_count} duplicate row(s).{preview_text}"
        if op == "FILTER":
            r = result if isinstance(result, dict) else {}
            col = self._display_source(instruction.get("source_name") or r.get("column", ""))
            return (
                f"Done. I kept {r.get('matched', 0)} of "
                f"{r.get('total', 0)} rows matching the filter on {col}."
            )
        if op == "FIND_EMPTY":
            if not written_count:
                return "No empty cells were found."
            r = result if isinstance(result, dict) else {}
            cells = r.get("cells", [])
            preview = ", ".join(cells[:10])
            more = f" and {written_count - 10} more" if len(cells) > 10 else ""
            return f"Found {written_count} empty cell(s): {preview}{more}."
        if op == "STANDARDIZE_DATES":
            r = result if isinstance(result, dict) else {}
            return f"Done. I standardized {r.get('converted', 0)} date(s) to DD-MM-YYYY format."
        if op == "STANDARDIZE_NAMES":
            r = result if isinstance(result, dict) else {}
            return f"Done. I title-cased {r.get('converted', 0)} name(s)."
        if op == "DETECT_INVALID":
            if not written_count:
                return "No invalid values were found."
            r = result if isinstance(result, dict) else {}
            issues = r.get("issues", [])
            preview = "; ".join(f"{ref} ({why})" for ref, why in issues[:8])
            more = f" and {written_count - 8} more" if len(issues) > 8 else ""
            return f"Found {written_count} invalid value(s): {preview}{more}."
        if op in ("SUMIF", "COUNTIF", "AVERAGEIF"):
            matched = instruction.get("matched", 0)
            agg = {"SUMIF": "summed", "COUNTIF": "counted", "AVERAGEIF": "averaged"}[op]
            crits = instruction.get("criteria", [])
            desc = " and ".join(f"{c[0]} {c[1]} {c[2]}" for c in crits)
            if destination:
                return (
                    f"Done. I {agg} {matched} matching cell(s) where {desc} "
                    f"and placed the result in {destination}."
                )
            return f"Done. I {agg} {matched} matching cell(s) where {desc} = {result}."
        if op == "CHART":
            r = result if isinstance(result, dict) else {}
            return (
                f"Done. I created a chart for '{r.get('chart')}' from {written_count} rows of data."
            )
        if op == "ANALYZE":
            r = result if isinstance(result, dict) else {}
            if r.get("lines"):
                top = "\n".join(r["lines"][:5])
                return f"Analysis complete across {len(r.get('sheets', []))} sheet(s).\n{top}"
            return "Analysis complete."
        if op == "GROWTH":
            named_inputs = instruction.get("named_inputs")
            inputs = instruction.get("inputs")
            if named_inputs and len(named_inputs) >= 2:
                old_name = self._display_source(named_inputs[1])
                new_name = self._display_source(named_inputs[0])
                location = f"Growth from {old_name} to {new_name}"
            elif inputs and len(inputs) >= 2:
                location = f"Growth from {inputs[1]} to {inputs[0]}"
            else:
                location = "Growth"
            percent = result
            if percent is not None:
                percent = round(percent, 2)
                if float(percent).is_integer():
                    percent = int(percent)
            return f"Done. {location} is {percent}%."
        if op == "PERCENTAGE":
            named = (
                instruction.get("named_source")
                or instruction.get("column")
                or instruction.get("source", "")
            )
            total = instruction.get("total", 0)
            out_col = instruction.get("output_column")
            if out_col and written_count:
                return (
                    f"Done. I wrote {written_count} percentage formula(s) to "
                    f"column {out_col} (each value of {named} as a % of its "
                    f"total {total})."
                )
            pcts = instruction.get("percentages", [])
            preview = ", ".join(self._format_percent(p) for p in pcts[:8])
            if len(pcts) > 8:
                preview += f" and {len(pcts) - 8} more"
            return f"Done. Each value of {named} as a % of its total ({total}): {preview}."

        source = instruction.get("column") or instruction.get("source")
        used_range = instruction.get("used_range")

        if source and used_range:
            location = f"{op.title()} of {used_range}"
        elif instruction.get("named_source"):
            location = f"{op.title()} of the {instruction['named_source']} column"
        elif instruction.get("whole_sheet"):
            location = f"{op.title()} of all the data"
        else:
            inputs = instruction.get("inputs", [])
            if inputs:
                display = " and ".join(str(i) for i in inputs[:2])
                location = f"{display} ({op.title()})"
            else:
                location = "calculation"

        if instruction.get("output_column") and written_count:
            return (
                f"Done. I wrote {written_count} formula(s) to "
                f"column {instruction['output_column']}."
            )

        if destination:
            return (
                f"Done. I calculated {location} = {result} and placed the result in {destination}."
            )
        else:
            return f"Done. The result of {location} is {result}."

    # ------------------------------------------------------------------ #
    # Interactive loop
    # ------------------------------------------------------------------ #

    @staticmethod
    def _display_source(value):
        """Return a friendly display name for a column source.

        Column letters are shown as-is; workbook header names are
        title-cased (e.g. 'salary' -> 'Salary').
        """
        if isinstance(value, str) and len(value) > 1:
            return value.title()
        return value

    @staticmethod
    def _format_percent(value):
        """Format a percentage number without trailing zero noise."""
        rounded = round(value, 2)
        if float(rounded).is_integer():
            return f"{int(rounded)}%"
        return f"{rounded}%"

    @staticmethod
    def _is_overwrite_confirmation(result):
        """Return True when a command only failed because a destination
        already contains data and an overwrite decision is needed."""
        if not isinstance(result, dict):
            return False
        message = result.get("message", "")
        return not result.get("success", True) and "already contains data" in message

    def process_interactive(self, command, sheet_name=None):
        """Process a command in an interactive session.

        If writing to a destination would overwrite existing data, the user
        is prompted for permission (just like the web UI confirms). Returns
        the final result dict, or ``None`` if the user chose to quit while
        answering the prompt.
        """
        result = self.process_command(command, sheet_name=sheet_name)

        if self._is_overwrite_confirmation(result):
            print("-" * 40)
            print(result["message"])
            try:
                choice = input("Replace it? [Yes/No]: ").strip().lower()
            except (EOFError, KeyboardInterrupt):
                print("\nGoodbye!")
                return None
            if choice in ("y", "ye", "yes", "ha", "haan"):
                result = self.process_command(command, sheet_name=sheet_name, allow_overwrite=True)
            else:
                print("OK, I didn't change anything.")

        return result

    def run_interactive(self, filepath=None):
        """Run an interactive command loop."""
        if filepath:
            self.load_workbook(filepath)

        print("=" * 60)
        print("Excel AI Copilot")
        print("Type a command like 'Column D ka sum karo' or 'Sales ka total karo'.")
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

            result = self.process_interactive(command)

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
