import pytest

from backend.calculator.engine import CalculationEngine


class TestSum:
    def test_sum_of_numbers(self):
        assert CalculationEngine.sum([1, 2, 3, 4, 5]) == 15

    def test_sum_empty(self):
        assert CalculationEngine.sum([]) == 0

    def test_sum_negative(self):
        assert CalculationEngine.sum([-1, -2, -3]) == -6


class TestAverage:
    def test_average_of_numbers(self):
        assert CalculationEngine.average([1, 2, 3, 4, 5]) == 3.0

    def test_average_empty(self):
        assert CalculationEngine.average([]) == 0

    def test_average_single(self):
        assert CalculationEngine.average([42]) == 42


class TestMinimum:
    def test_min_of_numbers(self):
        assert CalculationEngine.minimum([5, 2, 9, 1, 7]) == 1

    def test_min_empty(self):
        assert CalculationEngine.minimum([]) is None


class TestMaximum:
    def test_max_of_numbers(self):
        assert CalculationEngine.maximum([5, 2, 9, 1, 7]) == 9

    def test_max_empty(self):
        assert CalculationEngine.maximum([]) is None


class TestCount:
    def test_count_of_numbers(self):
        assert CalculationEngine.count([1, 2, 3]) == 3

    def test_count_empty(self):
        assert CalculationEngine.count([]) == 0


class TestAdd:
    def test_add(self):
        assert CalculationEngine.add(2, 3) == 5

    def test_add_negative(self):
        assert CalculationEngine.add(-2, 3) == 1


class TestSubtract:
    def test_subtract(self):
        assert CalculationEngine.subtract(5, 3) == 2

    def test_subtract_negative_result(self):
        assert CalculationEngine.subtract(3, 5) == -2


class TestMultiply:
    def test_multiply(self):
        assert CalculationEngine.multiply(4, 5) == 20

    def test_multiply_zero(self):
        assert CalculationEngine.multiply(4, 0) == 0


class TestDivide:
    def test_divide(self):
        assert CalculationEngine.divide(10, 2) == 5.0

    def test_divide_by_zero(self):
        with pytest.raises(ValueError):
            CalculationEngine.divide(10, 0)


class TestPercentage:
    def test_percentage(self):
        assert CalculationEngine.percentage(20, 100) == 20.0

    def test_percentage_zero_total(self):
        with pytest.raises(ValueError):
            CalculationEngine.percentage(20, 0)


class TestDifference:
    def test_difference(self):
        assert CalculationEngine.difference(10, 4) == 6

    def test_difference_absolute(self):
        assert CalculationEngine.difference(4, 10) == 6


class TestGrowth:
    def test_growth_positive(self):
        assert CalculationEngine.growth(120, 100) == 20.0

    def test_growth_negative(self):
        assert CalculationEngine.growth(80, 100) == -20.0

    def test_growth_zero_old(self):
        with pytest.raises(ValueError):
            CalculationEngine.growth(100, 0)
