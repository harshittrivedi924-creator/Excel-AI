import openpyxl
import pytest

from backend.copilot import ExcelCopilot


@pytest.fixture
def sales_workbook(tmp_path):
    """Workbook with messy/duplicate rows for data ops."""
    filepath = tmp_path / "sales.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sales"

    ws["A1"] = "Product"
    ws["B1"] = "Revenue"
    ws["C1"] = "Margin"

    rows = [
        ("Pen", 100, 40),
        ("Pencil", 50, 20),
        ("Pen", 100, 40),
        ("Eraser", 200, 80),
        ("Pen", 100, 40),
        ("Marker", 300, 120),
    ]
    for i, (p, r, m) in enumerate(rows, start=2):
        ws[f"A{i}"] = p
        ws[f"B{i}"] = r
        ws[f"C{i}"] = m

    wb.save(filepath)
    return str(filepath)


class TestSort:
    def test_sort_column(self, sales_workbook):
        copilot = ExcelCopilot(sales_workbook)
        result = copilot.process_command("Column B sort karo")
        assert result["success"] is True
        assert result["written_count"] == 6

        copilot.excel.save()
        # First data row should now be the smallest revenue (50)
        wb = openpyxl.load_workbook(sales_workbook)
        ws = wb["Sales"]
        assert ws["A2"].value == "Pencil"
        assert ws["A7"].value == "Marker"


class TestDedupe:
    def test_remove_duplicates(self, sales_workbook):
        copilot = ExcelCopilot(sales_workbook)
        result = copilot.process_command("Duplicate rows hata do")
        assert result["success"] is True
        assert result["written_count"] == 2  # two extra "Pen" rows

        copilot.excel.save()
        wb = openpyxl.load_workbook(sales_workbook)
        ws = wb["Sales"]
        products = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
        assert products == ["Pen", "Pencil", "Eraser", "Marker"]
        assert len(products) == 4


class TestFilter:
    def test_filter_greater_than(self, sales_workbook):
        copilot = ExcelCopilot(sales_workbook)
        result = copilot.process_command(
            "Filter data jahan Revenue 100 se zyada"
        )
        assert result["success"] is True
        r = result["result"]
        assert r["matched"] == 2  # Eraser(200), Marker(300)

        copilot.excel.save()
        wb = openpyxl.load_workbook(sales_workbook)
        assert "Filtered" in wb.sheetnames
        ws = wb["Filtered"]
        revenues = [ws.cell(row=r, column=2).value for r in range(2, ws.max_row + 1)]
        assert revenues == [200, 300]


class TestFindEmpty:
    def test_find_empty_column(self, sales_workbook):
        copilot = ExcelCopilot(sales_workbook)
        result = copilot.process_command("Khali cells dhundo column B")
        assert result["success"] is True

    def test_find_empty_whole_sheet(self, sales_workbook):
        copilot = ExcelCopilot(sales_workbook)
        result = copilot.process_command("Empty cells find karo")
        assert result["success"] is True
        assert result["written_count"] >= 0


class TestChart:
    def test_create_chart(self, sales_workbook):
        copilot = ExcelCopilot(sales_workbook)
        result = copilot.process_command("Revenue ka chart bana do")
        assert result["success"] is True
        r = result["result"]
        assert r["chart"] == "Revenue"

        copilot.excel.save()
        wb = openpyxl.load_workbook(sales_workbook)
        ws = wb["Sales"]
        assert len(ws._charts) >= 1


class TestAnalyze:
    def test_analyze_workbook(self, sales_workbook):
        copilot = ExcelCopilot(sales_workbook)
        result = copilot.process_command("Is workbook ka analysis kar do")
        assert result["success"] is True
        r = result["result"]
        assert r["insights"] >= 2  # Revenue + Margin columns
        assert any("Revenue" in line for line in r["lines"])

        copilot.excel.save()
        wb = openpyxl.load_workbook(sales_workbook)
        assert "Analysis Report" in wb.sheetnames


class TestDataParser:
    def test_parse_sort(self):
        from backend.parser.parser import CommandParser
        parser = CommandParser()
        result = parser.parse("Column B sort karo")
        assert result["operation"] == "SORT"
        assert result["column"] == "B"

    def test_parse_dedupe(self):
        from backend.parser.parser import CommandParser
        parser = CommandParser()
        result = parser.parse("Duplicate rows hata do")
        assert result["operation"] == "DEDUPE"

    def test_parse_filter(self):
        from backend.parser.parser import CommandParser
        parser = CommandParser()
        result = parser.parse("Filter data jahan Revenue 100 se zyada")
        assert result["operation"] == "FILTER"
        assert result["column"].upper() == "REVENUE"
        assert result["operator"] == ">"
        assert result["value"] == 100

    def test_parse_chart(self):
        from backend.parser.parser import CommandParser
        parser = CommandParser()
        result = parser.parse("Sales ka chart bana do")
        assert result["operation"] == "CHART"
        assert result["column"] == "sales"

    def test_parse_analyze(self):
        from backend.parser.parser import CommandParser
        parser = CommandParser()
        result = parser.parse("Is workbook ka analysis kar do")
        assert result["operation"] == "ANALYZE"