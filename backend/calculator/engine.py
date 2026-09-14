class CalculationEngine:
    """Excel calculation engine for basic operations."""

    @staticmethod
    def sum(values):
        """Calculate sum of values."""
        if not values:
            return 0
        return sum(values)

    @staticmethod
    def average(values):
        """Calculate average of values."""
        if not values:
            return 0
        return sum(values) / len(values)

    @staticmethod
    def minimum(values):
        """Find minimum value."""
        if not values:
            return None
        return min(values)

    @staticmethod
    def maximum(values):
        """Find maximum value."""
        if not values:
            return None
        return max(values)

    @staticmethod
    def count(values):
        """Count number of values."""
        return len(values)

    @staticmethod
    def add(a, b):
        """Add two numbers."""
        if not isinstance(a, (int, float)) or not isinstance(b, (int, float)):
            raise ValueError("Both operands must be numbers")
        return a + b

    @staticmethod
    def subtract(a, b):
        """Subtract b from a."""
        if not isinstance(a, (int, float)) or not isinstance(b, (int, float)):
            raise ValueError("Both operands must be numbers")
        return a - b

    @staticmethod
    def multiply(a, b):
        """Multiply two numbers."""
        if not isinstance(a, (int, float)) or not isinstance(b, (int, float)):
            raise ValueError("Both operands must be numbers")
        return a * b

    @staticmethod
    def divide(a, b):
        """Divide a by b."""
        if not isinstance(a, (int, float)) or not isinstance(b, (int, float)):
            raise ValueError("Both operands must be numbers")
        if b == 0:
            raise ValueError("Cannot divide by zero")
        return a / b

    @staticmethod
    def percentage(value, total):
        """Calculate percentage."""
        if not isinstance(value, (int, float)) or not isinstance(total, (int, float)):
            raise ValueError("Both value and total must be numbers")
        if total == 0:
            raise ValueError("Total cannot be zero")
        return (value / total) * 100

    @staticmethod
    def difference(a, b):
        """Absolute difference between two numbers."""
        if not isinstance(a, (int, float)) or not isinstance(b, (int, float)):
            raise ValueError("Both operands must be numbers")
        return abs(a - b)

    @staticmethod
    def growth(new, old):
        """Percentage growth of new relative to old."""
        if not isinstance(new, (int, float)) or not isinstance(old, (int, float)):
            raise ValueError("Both values must be numbers")
        if old == 0:
            raise ValueError("Old value cannot be zero for growth")
        return ((new - old) / old) * 100

    @staticmethod
    def sum_if(values, mask):
        """Sum values where mask is True."""
        if not values or not mask:
            return 0
        return sum(v for v, keep in zip(values, mask, strict=False) if keep)

    @staticmethod
    def average_if(values, mask):
        """Average values where mask is True."""
        if not values or not mask:
            return 0
        matched = [v for v, keep in zip(values, mask, strict=False) if keep]
        if not matched:
            return 0
        return sum(matched) / len(matched)

    @staticmethod
    def count_if(mask):
        """Count positions where mask is True."""
        if not mask:
            return 0
        return sum(1 for keep in mask if keep)

    @staticmethod
    def standard_deviation(values):
        """Sample standard deviation of a list of numbers."""
        if not values or len(values) < 2:
            return 0.0
        mean = sum(values) / len(values)
        variance = sum((v - mean) ** 2 for v in values) / (len(values) - 1)
        return variance**0.5

    @staticmethod
    def median(values):
        """Calculate median of values."""
        if not values:
            return None
        sorted_vals = sorted(values)
        n = len(sorted_vals)
        mid = n // 2
        if n % 2 == 0:
            return (sorted_vals[mid - 1] + sorted_vals[mid]) / 2
        return sorted_vals[mid]

    @staticmethod
    def mode(values):
        """Calculate mode of values."""
        if not values:
            return None
        from collections import Counter

        counts = Counter(values)
        max_count = max(counts.values())
        modes = [v for v, c in counts.items() if c == max_count]
        return modes[0] if len(modes) == 1 else modes

    @staticmethod
    def variance(values):
        """Calculate population variance of values."""
        if not values or len(values) < 2:
            return 0.0
        mean = sum(values) / len(values)
        return sum((v - mean) ** 2 for v in values) / len(values)
