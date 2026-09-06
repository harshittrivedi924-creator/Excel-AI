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