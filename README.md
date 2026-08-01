![PyPI](https://img.shields.io/pypi/v/magpyxl)
![Python](https://img.shields.io/pypi/pyversions/magpyxl)
![License](https://img.shields.io/github/license/ritik197/magPyXL)
![Downloads](https://img.shields.io/pypi/dm/magpyxl)

# MagpyXL

Spreadsheet-inspired formulas for Python, powered by Rust.

## Why MagpyXL?

MagpyXL brings familiar spreadsheet-style functions to Python with a consistent API and a high-performance Rust backend.

Use them across Python lists, NumPy arrays, pandas, polars, CSV files, and Excel files without changing your code. Whether you're building reports, performing data analysis, or automating spreadsheets, MagpyXL provides an intuitive, Spreadsheet-inspired experience for Python developers.

## Install

```bash
pip install magpyxl
```
## Supported Inputs

MagpyXL works with:

- Python lists and tuples
- NumPy arrays
- pandas Series and DataFrames
- polars Series and DataFrames
- CSV files
- Excel (.xlsx) files

The same formula syntax works across all supported data sources.

---

## Features

- Spreadsheet-inspired formula syntax
- Rust-powered execution
- Automatic vectorized operations
- Wildcard support (`*`, `?`)
- Spreadsheet-compatible criteria (`>`, `<`, `>=`, `<=`, `<>`)
- Case-insensitive text matching
- Supports scalar and vectorized lookups
- Works with pandas, polars, NumPy, CSV, and Excel

---

## Current Features

### Aggregate Functions
- SUM
- AVERAGE
- COUNT

### Conditional Functions
- COUNTIF
- SUMIF
- AVERAGEIF
- COUNTIFS
- SUMIFS

### Lookup Functions
- VLOOKUP
- XLOOKUP
- LOOKUPIFS

### Performance Features
- Rust-powered execution
- Automatic vectorized operations
- Batch lookup optimization
- Batch COUNTIF/SUMIF/COUNTIFS/SUMIFS evaluation
- Wildcard support (`*`, `?`)
- Spreadsheet-compatible comparison operators (`>`, `<`, `>=`, `<=`, `<>`)
- Case-insensitive text matching

### Supported Data Sources
- Python lists & tuples
- NumPy arrays
- pandas Series & DataFrames
- polars Series & DataFrames
- CSV files
- Excel (.xlsx) files

### Platform Support
- Python 3.8+
- Windows
- Linux
- macOS

## Contributing

Contributions, feature requests, and bug reports are welcome.

If you find a bug or have an idea for a new Spreadsheet function, please open an Issue or Pull Request.

---

## License

This project is licensed under the MIT License.
