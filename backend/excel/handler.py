import re
from datetime import date, datetime

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import column_index_from_string, get_column_letter


class ExcelHandler:
    """Read and write Excel workbooks."""

    def __init__(self, filepath=None):
        self.filepath = filepath
        self.workbook = None
        if filepath:
            self.load(filepath)

    def load(self, filepath):
        """Load an Excel workbook from filepath."""
        import os

        filepath_str = str(filepath)
        if not os.path.exists(filepath_str):
            raise FileNotFoundError(f"File not found: {filepath}")
        if not filepath_str.endswith((".xlsx", ".xlsm", ".xls")):
            raise ValueError("File must be an Excel workbook (.xlsx, .xlsm, or .xls)")
        self.filepath = filepath_str
        self.workbook = openpyxl.load_workbook(filepath_str)
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
        if not self.workbook:
            raise ValueError("No workbook loaded")
        return self.workbook.sheetnames

    def get_sheet(self, name=None):
        """Get a worksheet by name or active sheet."""
        if not self.workbook:
            raise ValueError("No workbook loaded")
        if name is None:
            return self.workbook.active
        if name not in self.workbook.sheetnames:
            raise ValueError(
                f"Sheet '{name}' not found. Available: {', '.join(self.workbook.sheetnames)}"
            )
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

    def get_headers(self, sheet_name=None):
        """Return the header values from the first row, lowercased."""
        sheet = self.get_sheet(sheet_name)
        headers = []
        for col in range(sheet.min_column, sheet.max_column + 1):
            value = sheet.cell(row=sheet.min_row, column=col).value
            if value is not None:
                headers.append(str(value).strip().lower())
        return headers

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

    # ------------------------------------------------------------------ #
    # Data operations
    # ------------------------------------------------------------------ #

    def _read_rows(self, sheet):
        """Read all used rows as lists of values."""
        rows = []
        for row in sheet.iter_rows(
            min_row=sheet.min_row,
            max_row=sheet.max_row,
            min_col=sheet.min_column,
            max_col=sheet.max_column,
        ):
            rows.append([cell.value for cell in row])
        return rows

    def _write_rows(self, sheet, rows):
        """Overwrite the used range with the given rows (clears leftovers)."""
        for row in sheet.iter_rows(
            min_row=sheet.min_row,
            max_row=sheet.max_row,
            min_col=sheet.min_column,
            max_col=sheet.max_column,
        ):
            for cell in row:
                cell.value = None

        for r_idx, row in enumerate(rows, start=1):
            for c_idx, value in enumerate(row, start=1):
                sheet.cell(row=r_idx, column=c_idx).value = value

    def sort_sheet(self, column, ascending=True, sheet_name=None):
        """Sort the sheet by a column, keeping the header row on top.
        Rows with a blank value in the sort column are always kept at the
        bottom (Excel-style) regardless of the sort direction.
        Returns the number of data rows sorted.
        """
        sheet = self.get_sheet(sheet_name)
        if isinstance(column, str):
            col_index = self.get_column_number(column)
        else:
            col_index = column

        rows = self._read_rows(sheet)
        if not rows:
            return 0
        header = rows[0]
        data = rows[1:]

        def value_of(row):
            return row[col_index - 1] if col_index <= len(row) else None

        with_value = []
        blank = []
        for row in data:
            value = value_of(row)
            if value is None or (isinstance(value, str) and not value.strip()):
                blank.append(row)
            else:
                with_value.append(row)

        if ascending:
            with_value.sort(key=lambda row: self._sort_key_asc(value_of(row)))
        else:
            with_value.sort(key=lambda row: self._sort_key_desc(value_of(row)))

        self._write_rows(sheet, [header] + with_value + blank)
        return len(data)

    @staticmethod
    def _sort_key_asc(value):
        """Ascending sort key: numbers first, then text, blanks handled outside."""
        try:
            return (0, float(value))
        except (TypeError, ValueError):
            return (1, str(value).lower())

    @staticmethod
    def _sort_key_desc(value):
        """Descending sort key: numbers descending first, then text, then blanks."""
        try:
            return (0, -float(value))
        except (TypeError, ValueError):
            return (1, str(value).lower())

    def remove_duplicates(self, sheet_name=None):
        """Remove fully-duplicate rows, keeping the first occurrence.
        Returns (removed_count, list_of_removed_rows).
        """
        sheet = self.get_sheet(sheet_name)
        rows = self._read_rows(sheet)
        if not rows:
            return 0, []
        header = rows[0]
        data = rows[1:]

        seen = set()
        unique = []
        removed = []
        for row in data:
            signature = tuple(repr(v) for v in row)
            if signature in seen:
                removed.append(row)
                continue
            seen.add(signature)
            unique.append(row)

        self._write_rows(sheet, [header] + unique)
        return len(data) - len(unique), removed

    def filter_rows(self, column, operator, value, sheet_name=None):
        """Filter rows by a condition on a column.
        Writes matching rows to a new 'Filtered' sheet.
        Returns (matched_count, total_rows).
        """
        sheet = self.get_sheet(sheet_name)
        if isinstance(column, str):
            col_index = self.get_column_number(column)
        else:
            col_index = column

        rows = self._read_rows(sheet)
        if not rows:
            return 0, 0
        header = rows[0]
        data = rows[1:]

        if isinstance(value, str):
            try:
                value = float(value)
            except ValueError:
                value = value.lower()

        matches = []
        # Count rows that hold data in the filtered column, so trailing padding
        # in the used range is not reported as part of the dataset.
        total = 0
        for row in data:
            cell_value = row[col_index - 1] if col_index <= len(row) else None
            if cell_value is not None and cell_value != "":
                total += 1

            def coerce(v):
                if isinstance(v, str):
                    try:
                        return float(v)
                    except ValueError:
                        return v.lower()
                return v

            cv = coerce(cell_value)
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
            if ok:
                matches.append(row)

        # Write to a new sheet
        if "Filtered" in self.workbook.sheetnames:
            del self.workbook["Filtered"]
        filtered = self.workbook.create_sheet("Filtered")
        self._write_rows(filtered, [header] + matches)
        return len(matches), total

    def find_empty_cells(self, column=None, sheet_name=None):
        """Find empty cells in the used range (optionally a single column).
        Returns a list of cell references like ['B3', 'D5'].
        """
        sheet = self.get_sheet(sheet_name)
        empty = []
        max_col = sheet.max_column if column is None else (self.get_column_number(column))
        if column is None:
            min_col = sheet.min_column
        else:
            min_col = max_col

        for row in range(sheet.min_row, sheet.max_row + 1):
            for col in range(min_col, max_col + 1):
                if sheet.cell(row=row, column=col).value is None:
                    empty.append(f"{get_column_letter(col)}{row}")
        return empty

    # ------------------------------------------------------------------ #
    # Data cleaning (Phase 11)
    # ------------------------------------------------------------------ #

    @staticmethod
    def _parse_date(text):
        """Best-effort parse of a date string into a datetime.date.
        Returns None if the text is not a recognizable date.
        """
        text = str(text).strip()
        for fmt in (
            "%d-%m-%Y",
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%Y/%m/%d",
            "%d-%m-%y",
            "%d/%m/%y",
            "%m/%d/%Y",
            "%d %b %Y",
            "%b %d, %Y",
            "%B %d, %Y",
            "%d %B %Y",
        ):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                pass
        m = re.match(r"(\d{1,2})\.(\d{1,2})\.(\d{2,4})", text)
        if m:
            day, month, year = map(int, m.groups())
            if 1 <= day <= 31 and 1 <= month <= 12:
                year = year + 2000 if year < 100 else year
                try:
                    return date(year, month, day)
                except ValueError:
                    return None
        return None

    def standardize_dates(self, column=None, sheet_name=None):
        """Rewrite date-like cells to a consistent DD-MM-YYYY format.
        Returns the number of cells converted.
        """
        sheet = self.get_sheet(sheet_name)
        count = 0
        cols = range(sheet.min_column, sheet.max_column + 1)
        if column is not None:
            cols = [self.get_column_number(column)]

        for col in cols:
            for row in range(sheet.min_row, sheet.max_row + 1):
                cell = sheet.cell(row=row, column=col)
                value = cell.value
                parsed = None
                if isinstance(value, datetime):
                    parsed = value.date()
                elif isinstance(value, date):
                    parsed = value
                elif isinstance(value, str):
                    parsed = self._parse_date(value)
                if parsed is None:
                    continue
                cell.value = parsed.strftime("%d-%m-%Y")
                count += 1
        return count, count

    @staticmethod
    def _is_name_header(text):
        header = str(text).strip().lower()
        return header in (
            "name",
            "names",
            "naam",
            "customer",
            "customername",
            "client",
            "product",
            "employee",
            "person",
            "user",
        )

    def standardize_names(self, column=None, sheet_name=None):
        """Title-case text cells in the first name-like column
        (or the requested column). Returns (converted, total_checked).
        """
        sheet = self.get_sheet(sheet_name)
        target_col = column

        if target_col is None:
            for col in range(sheet.min_column, sheet.max_column + 1):
                header = sheet.cell(row=sheet.min_row, column=col).value
                if header and self._is_name_header(header):
                    target_col = get_column_letter(col)
                    break
        if target_col is None:
            # Fall back to the first non-numeric, non-header column with text.
            for col in range(sheet.min_column, sheet.max_column + 1):
                if col == 1 and sheet.max_column > 1:
                    continue
                for row in range(sheet.min_row + 1, sheet.max_row + 1):
                    v = sheet.cell(row=row, column=col).value
                    if isinstance(v, str) and not self._parse_date(v):
                        try:
                            float(v)
                        except (ValueError, TypeError):
                            target_col = get_column_letter(col)
                            break
                if target_col:
                    break
        if target_col is None:
            return 0, 0

        col_index = self.get_column_number(target_col)
        converted = 0
        checked = 0
        for row in range(sheet.min_row + 1, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=col_index)
            value = cell.value
            if not isinstance(value, str):
                continue
            checked += 1
            title = " ".join(w.capitalize() if w else w for w in value.split())
            if title != value:
                cell.value = title
                converted += 1
        return converted, checked

    def detect_invalid_values(self, selected_column=None, sheet_name=None):
        """Detect invalid values in the used range: non-numeric cells inside
        otherwise-numeric columns, plus empty gaps in the middle of columns.
        Returns a list of (cell_ref, reason) and the raw candidate list.
        """
        sheet = self.get_sheet(sheet_name)
        issues = []

        for col in range(sheet.min_column, sheet.max_column + 1):
            col_letter = get_column_letter(col)
            values = []
            rows_with_value = []
            for row in range(sheet.min_row, sheet.max_row + 1):
                cell = sheet.cell(row=row, column=col)
                if cell.value is None:
                    continue
                if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                    values.append(cell.value)
                    rows_with_value.append(row)
                elif self._is_numeric_str(cell.value):
                    values.append(float(cell.value))
                    rows_with_value.append(row)
            if not values:
                continue

            # Numeric column: any non-empty cell that is not numeric is invalid.
            for row in range(sheet.min_row + 1, sheet.max_row + 1):
                cell = sheet.cell(row=row, column=col)
                if cell.value is None:
                    continue
                if not self._is_numeric_str(cell.value):
                    if isinstance(cell.value, str):
                        issues.append((f"{col_letter}{row}", "text in numeric column"))
                    elif isinstance(cell.value, bool):
                        issues.append((f"{col_letter}{row}", "TRUE/FALSE in numeric column"))

            # Empty gap in the middle of the column.
            if rows_with_value:
                top, bottom = min(rows_with_value), max(rows_with_value)
                for row in range(top, bottom + 1):
                    cell = sheet.cell(row=row, column=col)
                    if cell.value is None:
                        issues.append((f"{col_letter}{row}", "empty cell in data range"))

        return issues, issues

    @staticmethod
    def _is_numeric_str(value):
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return True
        if not isinstance(value, str):
            return False
        try:
            float(value.strip())
            return True
        except ValueError:
            return False

    def create_chart(self, column, sheet_name=None, chart_type="column"):
        """Create a chart for a numeric column and add it to the sheet.
        Returns a dict describing the chart.
        """
        sheet = self.get_sheet(sheet_name)
        if isinstance(column, str):
            col_index = self.get_column_number(column)
        else:
            col_index = column
        col_letter = get_column_letter(col_index)

        # Use the last row that actually has data in the charted column, so
        # trailing padding in the used range does not add empty categories.
        max_row = max(
            (
                r
                for r in range(2, sheet.max_row + 1)
                if sheet.cell(row=r, column=col_index).value not in (None, "")
            ),
            default=0,
        )
        if max_row < 2:
            raise ValueError("Not enough data to chart.")

        title = str(sheet.cell(row=1, column=col_index).value or f"Column {col_letter}")
        data_ref = Reference(sheet, min_col=col_index, min_row=1, max_row=max_row)

        # Use the first column as categories if it has data
        cats = None
        first_col_has = any(
            sheet.cell(row=r, column=1).value is not None for r in range(1, max_row + 1)
        )
        if first_col_has and col_index != 1:
            cats = Reference(sheet, min_col=1, min_row=2, max_row=max_row)

        chart = BarChart() if chart_type == "column" else LineChart()
        chart.add_data(data_ref, titles_from_data=True)
        if cats is not None:
            chart.set_categories(cats)
        chart.title = title
        chart.y_axis.title = title

        sheet.add_chart(chart, f"{col_letter}{max_row + 2}")
        return {"title": title, "source": f"{col_letter}", "rows": max_row - 1}

    # ------------------------------------------------------------------ #
    # Analysis
    # ------------------------------------------------------------------ #

    def analyze_workbook(self, sheet_name=None):
        """Produce a text analysis report of the workbook.
        Returns a dict of insights per sheet.
        """
        report = {}

        if sheet_name is not None:
            sheets = [sheet_name]
        else:
            sheets = self.workbook.sheetnames

        for name in sheets:
            sheet = self.workbook[name]
            insights = []
            max_col = sheet.max_column

            for col in range(1, max_col + 1):
                header = sheet.cell(row=sheet.min_row, column=col).value
                values, max_row, min_row = self.get_column_values(get_column_letter(col), name)
                if not values:
                    continue
                label = header or get_column_letter(col)
                total = sum(values)
                avg = sum(values) / len(values)
                high = max(values)
                low = min(values)
                insights.append(
                    {
                        "column": label,
                        "total": round(total, 2),
                        "average": round(avg, 2),
                        "max": high,
                        "min": low,
                        "count": len(values),
                    }
                )

            trends = self._detect_trends(sheet)
            anomalies = self._detect_anomalies(sheet)
            report[name] = {"metrics": insights, "trends": trends, "anomalies": anomalies}

        return report

    def _detect_anomalies(self, sheet):
        """Flag values more than 2 standard deviations from the column mean."""
        from backend.calculator.engine import CalculationEngine

        anomalies = []
        for col in range(sheet.min_column, sheet.max_column + 1):
            values, max_row, min_row = self.get_column_values(get_column_letter(col), sheet.title)
            if len(values) < 3:
                continue
            header = sheet.cell(row=sheet.min_row, column=col).value
            label = header or get_column_letter(col)
            mean = sum(values) / len(values)
            stdev = CalculationEngine.standard_deviation(values)
            if stdev == 0:
                continue
            for i, v in enumerate(values):
                if abs(v - mean) > 2 * stdev:
                    anomalies.append((label, v, mean, stdev, i))
        return anomalies

    def _detect_trends(self, sheet):
        """Detect simple trends across numeric columns vs. a label column."""
        trends = []
        label_col = None
        for col in range(sheet.min_column, sheet.max_column + 1):
            header = sheet.cell(row=sheet.min_row, column=col).value
            if header and str(header).lower() in (
                "month",
                "date",
                "name",
                "product",
                "year",
                "category",
            ):
                label_col = col
                break

        for col in range(sheet.min_column, sheet.max_column + 1):
            values, max_row, min_row = self.get_column_values(get_column_letter(col), sheet.title)
            if len(values) < 2:
                continue
            header = sheet.cell(row=sheet.min_row, column=col).value or "data"
            # Detect the biggest change between consecutive rows
            best_growth = None
            best_drop = None
            for i in range(1, len(values)):
                prev, curr = values[i - 1], values[i]
                if prev == 0:
                    continue
                pct = (curr - prev) / prev * 100
                if best_growth is None or pct > best_growth[2]:
                    best_growth = (i, curr, pct)
                if best_drop is None or pct < best_drop[2]:
                    best_drop = (i, curr, pct)
            if best_growth:
                label = ""
                if label_col:
                    label = str(
                        sheet.cell(row=min_row + best_growth[0], column=label_col).value or ""
                    )
                    label = f" ({label})"
                trends.append(f"{header} rose {best_growth[2]:.1f}% to {best_growth[1]}{label}")
            if best_drop and best_drop[2] < 0:
                label = ""
                if label_col:
                    label = str(
                        sheet.cell(row=min_row + best_drop[0], column=label_col).value or ""
                    )
                    label = f" ({label})"
                trends.append(f"{header} fell {abs(best_drop[2]):.1f}% to {best_drop[1]}{label}")
        return trends
