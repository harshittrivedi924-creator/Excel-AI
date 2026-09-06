from openpyxl.utils import get_column_letter, column_index_from_string

from backend.calculator.engine import CalculationEngine


class ValidationError(Exception):
    """Exception raised for validation failures."""

    def __init__(self, message):
        self.message = message
        super().__init__(message)


class Validator:
    """Validate Excel actions before execution."""

    SUPPORTED_OPERATIONS = {
        "SUM",
        "AVERAGE",
        "MIN",
        "MAX",
        "COUNT",
        "ADD",
        "SUBTRACT",
        "MULTIPLY",
        "DIVIDE",
        "PERCENTAGE",
        "DIFFERENCE",
        "GROWTH",
        "SORT",
        "DEDUPE",
        "FILTER",
        "FIND_EMPTY",
        "CHART",
        "ANALYZE",
        "SUMIF",
        "COUNTIF",
        "AVERAGEIF",
        "STANDARDIZE_DATES",
        "STANDARDIZE_NAMES",
        "DETECT_INVALID",
    }

    def __init__(self, excel_handler):
        self.excel = excel_handler

    def validate_sheet(self, sheet_name):
        """Check if the sheet exists."""
        sheet_names = self.excel.get_sheet_names()
        if sheet_name is not None and sheet_name not in sheet_names:
            raise ValidationError(
                f"I couldn't find the sheet '{sheet_name}' in this workbook. "
                f"Available sheets: {', '.join(sheet_names)}"
            )
        return True

    def validate_range(self, range_ref, sheet_name=None):
        """Validate that a range is valid and contains numeric data."""
        if not range_ref:
            raise ValidationError("No valid range was found.")

        values = self.excel.get_range_values(range_ref, sheet_name)
        numeric_values = [
            v for v in values
            if isinstance(v, (int, float))
        ]

        if not numeric_values:
            raise ValidationError(
                f"There are no numeric values in the selected range ({range_ref})."
            )

        return numeric_values

    def validate_column(self, column, sheet_name=None):
        """Validate a column letter and return used range."""
        if isinstance(column, str):
            try:
                col_index = column_index_from_string(column.upper())
            except Exception:
                raise ValidationError(f"I couldn't find Column {column} in this sheet.")

        used_range = self.excel.get_used_range(column, sheet_name)

        if used_range is None:
            raise ValidationError(f"I couldn't find any data in Column {column}.")

        return self.validate_range(used_range, sheet_name), used_range

    def validate_destination(self, cell_ref, sheet_name=None, allow_overwrite=False):
        """Validate a destination cell."""
        if not cell_ref:
            raise ValidationError("No destination cell was specified.")

        # Check if overwrite is needed
        if not allow_overwrite and self.excel.cell_has_data(cell_ref, sheet_name):
            raise ValidationError(
                f"{cell_ref} already contains data. Replace it?"
            )

        return True

    def validate_operation(self, operation):
        """Check if an operation is supported."""
        op = operation.upper()
        if op not in self.SUPPORTED_OPERATIONS:
            raise ValidationError(
                f"The operation '{operation}' is not supported. "
                f"Supported operations: {', '.join(sorted(self.SUPPORTED_OPERATIONS))}"
            )
        return op

    def execute(self, instruction):
        """Execute a validated instruction and return result."""
        operation = instruction.get("operation")
        result = None

        if operation == "SUM":
            values = instruction["values"]
            result = CalculationEngine.sum(values)
        elif operation == "AVERAGE":
            values = instruction["values"]
            result = CalculationEngine.average(values)
        elif operation == "MIN":
            values = instruction["values"]
            result = CalculationEngine.minimum(values)
        elif operation == "MAX":
            values = instruction["values"]
            result = CalculationEngine.maximum(values)
        elif operation == "COUNT":
            values = instruction["values"]
            result = CalculationEngine.count(values)
        elif operation in ("ADD", "SUBTRACT", "MULTIPLY", "DIVIDE"):
            a = instruction["a"]
            b = instruction["b"]
            if operation == "ADD":
                result = CalculationEngine.add(a, b)
            elif operation == "SUBTRACT":
                result = CalculationEngine.subtract(a, b)
            elif operation == "MULTIPLY":
                result = CalculationEngine.multiply(a, b)
            elif operation == "DIVIDE":
                result = CalculationEngine.divide(a, b)
        elif operation == "PERCENTAGE":
            value = instruction["value"]
            total = instruction["total"]
            result = CalculationEngine.percentage(value, total)
        elif operation == "DIFFERENCE":
            a = instruction["a"]
            b = instruction["b"]
            result = CalculationEngine.difference(a, b)
        elif operation == "GROWTH":
            new = instruction["new"]
            old = instruction["old"]
            result = CalculationEngine.growth(new, old)

        return result