import pytest

from backend.parser.parser import CommandParser, ParserError


class TestColumnCommands:
    def setup_method(self):
        self.parser = CommandParser()

    def test_column_sum_english(self):
        result = self.parser.parse("sum of column D")
        assert result["operation"] == "SUM"
        assert result["column"] == "D"

    def test_column_total_hinglish(self):
        result = self.parser.parse("Column D ka total karo")
        assert result["operation"] == "SUM"
        assert result["column"] == "D"

    def test_column_average_hinglish(self):
        result = self.parser.parse("D column ka average nikal do")
        assert result["operation"] == "AVERAGE"
        assert result["column"] == "D"

    def test_column_maximum_english(self):
        result = self.parser.parse("maximum of column E")
        assert result["operation"] == "MAX"
        assert result["column"] == "E"

    def test_column_minimum_english(self):
        result = self.parser.parse("minimum of column A")
        assert result["operation"] == "MIN"
        assert result["column"] == "A"

    def test_column_destination(self):
        result = self.parser.parse("Column D ka total karo aur D21 mein daal do")
        assert result["operation"] == "SUM"
        assert result["column"] == "D"
        assert result["destination"] == "D21"


class TestArithmeticCommands:
    def setup_method(self):
        self.parser = CommandParser()

    def test_multiply_cells(self):
        result = self.parser.parse("B2 aur C2 multiply karke D2 mein daalo")
        assert result["operation"] == "MULTIPLY"
        assert "B2" in result["inputs"]
        assert "C2" in result["inputs"]
        assert result["output"] == "D2"

    def test_add_cells(self):
        result = self.parser.parse("B2 + C2")
        assert result["operation"] == "ADD"
        assert "B2" in result["inputs"]

    def test_subtract(self):
        result = self.parser.parse("subtract B2 from C2")
        assert result["operation"] == "SUBTRACT"


class TestParserErrors:
    def setup_method(self):
        self.parser = CommandParser()

    def test_empty_command(self):
        with pytest.raises(ParserError):
            self.parser.parse("")

    def test_unknown_command(self):
        with pytest.raises(ParserError):
            self.parser.parse("hello world how are you")

    def test_no_data_specified(self):
        with pytest.raises(ParserError):
            self.parser.parse("sum karo")


class TestDetection:
    def setup_method(self):
        self.parser = CommandParser()

    def test_detect_sum_keyword(self):
        assert self.parser._detect_operation("total karo") == "SUM"

    def test_detect_average_keyword(self):
        assert self.parser._detect_operation("average nikalo") == "AVERAGE"

    def test_detect_max_keyword(self):
        assert self.parser._detect_operation("maximum batao") == "MAX"

    def test_detect_min_keyword(self):
        assert self.parser._detect_operation("minimum batao") == "MIN"

    def test_detect_count_keyword(self):
        assert self.parser._detect_operation("count karo") == "COUNT"