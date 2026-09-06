import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string


class ExcelHandler:
    """Read and write Excel workbooks."""

    def __init__(self, filepath=None):
        self.filepath = filepath
        self.workbook = None
        if filepath:
            self.load(filepath)

    def load(self, filepath):
        """Load an Excel workbook from filepath."""
        self.filepath = filepath
        self.workbook = openpyxl.load_workbook(filepath)
        return self.workbook

    def create(self, filepath=None):
        """Create a new workbook."""
        self.workbook = openpyxl.Workbook()
        if filepath:
            self.filepath = filepath
        return self.workbook

    def save(self, filepath=None):
        """Save the workbook."""
        target = filepath or self.filepath
        if not target:
            raise ValueError("No filepath specified to save")
        self.workbook.save(target)
        return target

    def get_sheet_names(self):
        """Return list of sheet names."""
        return self.workbook.sheetnames

    def get_sheet(self, name=None):
        """Get a worksheet by name or active sheet."""
        if name is None:
            return self.workbook.active
        return self.workbook[name]

    def get_cell_value(self, cell_ref, sheet_name=None):
        """Get the value of a cell like 'B2'."""
        sheet = self.get_sheet(sheet_name)
        return sheet[cell_ref].value

    def get_column_values(self, column, sheet_name=None, start_row=2):
        """Get all numeric values in a column.
        Column can be letter like 'D' or index like 4.
        Returns (values, max_row, min_row).
        """
        sheet = self.get_sheet(sheet_name)

        if isinstance(column, str):
            col_index = column_index_from_string(column.upper())
        else:
            col_index = column

        values = []
        min_row = None
        max_row = start_row if start_row else 1

        for row in range(sheet.min_row, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=col_index)
            if cell.value is not None:
                max_row = row
                if isinstance(cell.value, (int, float)):
                    if min_row is None:
                        min_row = row
                    values.append(cell.value)
                elif isinstance(cell.value, str):
                    try:
                        num = float(cell.value)
                        if min_row is None:
                            min_row = row
                        values.append(num)
                    except ValueError:
                        pass

        return values, max_row, min_row

    def get_range_values(self, range_ref, sheet_name=None):
        """Get values from a range like 'D2:D20'."""
        sheet = self.get_sheet(sheet_name)
        cell_range = sheet[range_ref]
        values = []
        for row in cell_range:
            for cell in row:
                values.append(cell.value)
        return values

    def get_used_range(self, column=None, sheet_name=None):
        """Determine the used range for a column.
        Returns a range reference like 'D2:D20' or None if no data.
        """
        sheet = self.get_sheet(sheet_name)

        if column is None:
            # Determine range for the entire used portion
            if sheet.max_row < 1 or sheet.max_column < 1:
                return None
            start_col = get_column_letter(sheet.min_column)
            end_col = get_column_letter(sheet.max_column)
            return f"{start_col}{sheet.min_row}:{end_col}{sheet.max_row}"

        if isinstance(column, str):
            col_index = column_index_from_string(column.upper())
        else:
            col_index = column

        min_row = None
        max_row = sheet.max_row

        for row in range(sheet.min_row, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=col_index)
            if cell.value is not None:
                if min_row is None:
                    min_row = row
                max_row = row

        if min_row is None:
            return None

        # Detect a header row: if the first non-empty cell in the column
        # is non-numeric and there is at least one more row of data below it,
        # treat that row as a header and start the data range at the next row.
        first_cell = sheet.cell(row=min_row, column=col_index)
        if (
            first_cell.value is not None
            and not isinstance(first_cell.value, (int, float))
            and max_row > min_row
        ):
            min_row += 1

        if min_row > max_row:
            return None

        col_letter = get_column_letter(col_index)
        return f"{col_letter}{min_row}:{col_letter}{max_row}"

    def set_cell_value(self, cell_ref, value, sheet_name=None):
        """Write a value/formula to a cell."""
        sheet = self.get_sheet(sheet_name)
        sheet[cell_ref] = value

    def write_formula(self, cell_ref, formula, sheet_name=None):
        """Write an Excel formula to a cell."""
        sheet = self.get_sheet(sheet_name)
        if not formula.startswith("="):
            formula = "=" + formula
        sheet[cell_ref] = formula
        return formula

    def get_formula(self, cell_ref, sheet_name=None):
        """Get the formula of a cell (without '=').
        Returns None if the cell has no formula.
        """
        sheet = self.get_sheet(sheet_name)
        cell = sheet[cell_ref]
        if isinstance(cell.value, str) and cell.value.startswith("="):
            return cell.value[1:]
        return None

    def cell_has_data(self, cell_ref, sheet_name=None):
        """Check if a cell already contains data."""
        sheet = self.get_sheet(sheet_name)
        return sheet[cell_ref].value is not None

    def get_column_letter(self, num):
        """Get column letter from number."""
        return get_column_letter(num)

    def get_column_number(self, letter):
        """Get column number from letter."""
        return column_index_from_string(letter)

    def get_column_by_name(self, name, sheet_name=None):
        """Find a column letter by matching its header text.
        Matches case-insensitively and ignores trailing spaces/percent.
        Returns the column letter (e.g. 'D') or None if not found.
        """
        sheet = self.get_sheet(sheet_name)
        search = name.strip().lower()

        for row in range(sheet.min_row, min(sheet.max_row, 3) + 1):
            for col_index in range(sheet.min_column, sheet.max_column + 1):
                header = sheet.cell(row=row, column=col_index).value
                if header is None:
                    continue
                header_str = str(header).strip().lower()
                header_str = header_str.replace("%", "").strip()
                if header_str == search or header_str.startswith(search):
                    return get_column_letter(col_index)
        return None