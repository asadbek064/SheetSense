# SheetSense
A comprehensive TypeScript library for analyzing Excel workbooks that detects quality issues, validates formulas, and ensures data consistency.

## Features

### Formula Analysis
- Detects complex nested formulas and suggests simplification
- Identifies potential division by zero errors
- Analyzes formula complexity based on nested functions and operators
- Validates formula structure and syntax

### Data Validation
- Ensures data type consistency within columns
- Validates date formats (YYYY-MM-DD)
- Detects mixed data types and suggests standardization
- Tracks column-level data quality metrics

### Style Analysis
- Validates consistent number formatting
- Detects mixed format usage (1000.00 vs 1,000)
- Provides formatting standardization suggestions
- Identifies inconsistent styling patterns

### Metadata Collection
- Tracks total formula count
- Monitors sheet statistics
- Records named ranges
- Identifies volatile functions
- Detects external references