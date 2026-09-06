class CalculationEngine:
    """Excel calculation engine for basic operations."""
    
    @staticmethod
    def sum(values):
        """Calculate sum of values."""
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
        return a + b
    
    @staticmethod
    def subtract(a, b):
        """Subtract b from a."""
        return a - b
    
    @staticmethod
    def multiply(a, b):
        """Multiply two numbers."""
        return a * b
    
    @staticmethod
    def divide(a, b):
        """Divide a by b."""
        if b == 0:
            raise ValueError("Cannot divide by zero")
        return a / b
    
    @staticmethod
    def percentage(value, total):
        """Calculate percentage."""
        if total == 0:
            raise ValueError("Total cannot be zero")
        return (value / total) * 100

    @staticmethod
    def difference(a, b):
        """Absolute difference between two numbers."""
        return abs(a - b)

    @staticmethod
    def growth(new, old):
        """Percentage growth of new relative to old."""
        if old == 0:
            raise ValueError("Old value cannot be zero for growth")
        return ((new - old) / old) * 100

    @staticmethod
    def sum_if(values, mask):
        """Sum values where mask is True."""
        return sum(v for v, keep in zip(values, mask) if keep)

    @staticmethod
    def average_if(values, mask):
        """Average values where mask is True."""
        matched = [v for v, keep in zip(values, mask) if keep]
        if not matched:
            return 0
        return sum(matched) / len(matched)

    @staticmethod
    def count_if(mask):
        """Count positions where mask is True."""
        return sum(1 for keep in mask if keep)

    @staticmethod
    def standard_deviation(values):
        """Sample standard deviation of a list of numbers."""
        if len(values) < 2:
            return 0.0
        mean = sum(values) / len(values)
        variance = sum((v - mean) ** 2 for v in values) / (len(values) - 1)
        return variance ** 0.5