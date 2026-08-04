# magpyxl

Spreadsheet-style formulas — `SUM`, `AVERAGE`, `COUNT`, `COUNTIF`, `SUMIF`, `AVERAGEIF`,
`COUNTIFS`, `SUMIFS`, `VLOOKUP`, `XLOOKUP`, `LOOKUPIFS` — backed by a compiled Rust core,
callable with the exact same syntax whether your data is a plain Python list,
a pandas DataFrame, a polars DataFrame, or a `.csv`/`.xlsx` file on disk.


## Install

```bash
pip install magpyxl
```

## Why magpyxl
magpyxl brings familiar Excel/Spreadsheet functions to Python with a consistent API, allowing the same function to work across multiple data sources.

## Features

- 🚀 Rust-powered computation
- 📊 Excel/Spreadsheet-compatible formulas
- 🐼 Native pandas support
- ⚡ Native Polars support
- 📄 CSV and Excel file support
- 🔄 Consistent API across data sources
- 🐍 Python 3.8+

## Quick start

```python
import magpyxl as mx

sales  = [1200, 800, 1500, 300, 950]
region = ["East", "West", "East", "South", "East"]

mx.SUM(sales)                              # 4750
mx.AVERAGE(sales)                          # 950.0
mx.COUNTIF(region, "East")                 # 3
mx.SUMIF(region, "East", sales)            # 3650
mx.SUMIFS(sales, region, "East", sales, ">900")   # AND across pairs
mx.VLOOKUP("Bob", [("Alice", 50000), ("Bob", 60000)], 2)   # 60000
mx.XLOOKUP("Carol", ["Alice", "Bob", "Carol"], [50000, 60000, 70000])  # 70000
```

### Works the same way with pandas

```python
import pandas as pd
df = pd.DataFrame({"Name": [...], "Dept": [...], "Salary": [...]})

mx.SUM(df["Salary"])
mx.AVERAGEIF(df["Dept"], "Sales", df["Salary"])
mx.VLOOKUP("Carol", df, "Salary")          # col_index can be a column name

# Vectorized: pass a whole column as lookup_value -> get a matching
# pandas Series back (same type in, same type out) -> chains right
# back into pandas code.
df["Price"] = mx.VLOOKUP(df["Key"], other_df, "Price", if_not_found=0)

```

### Works the same way with polars

```python
import polars as pl
pdf = pl.DataFrame({...})
mx.SUM(pdf["Salary"])
mx.VLOOKUP(pdf["Key"], other_pdf, "Price")   # returns a polars Series
mx.COUNTIF("Dept", "Eng")
```

### Standalone — no pandas/polars needed

```python 
mx.SUM(tbl["Revenue"])
mx.VLOOKUP("Dave", "sales.csv", "Salary")   # path works directly too
```

## Function reference

| Function | Signature | Notes |
|---|---|---|
| `SUM` | `SUM(range)` | Ignores text/blank cells |
| `AVERAGE` | `AVERAGE(range)` | Same |
| `COUNT` | `COUNT(range)` | Counts numeric cells only (Spreadsheet semantics) |
| `COUNTIF` | `COUNTIF(range, criteria)` | Criteria: `10`, `">10"`, `"<=5"`, `"<>0"`, `"ab*"`, `"a?c"` |
| `SUMIF` | `SUMIF(range, criteria, sum_range=None)` | |
| `AVERAGEIF` | `AVERAGEIF(range, criteria, average_range=None)` | |
| `COUNTIFS` | `COUNTIFS(range1, criteria1, range2, criteria2, ...)` | AND across all pairs |
| `SUMIFS` | `SUMIFS(sum_range, range1, criteria1, ...)` | AND across all pairs |
| `VLOOKUP` | `VLOOKUP(lookup_value, table, col_index, range_lookup=False, if_not_found=None)` | `col_index`: 1-based number or column name; `lookup_value` can be scalar or a whole column |
| `XLOOKUP` | `XLOOKUP(lookup_value, lookup_array, return_array, if_not_found=None)` | Same scalar-or-column behavior |
| `LOOKUPIFS` | `LOOKUPIFS(return_array, range1, criteria1, ...)` | AND across all pairs |

`table` (for VLOOKUP) accepts: pandas DataFrame, polars DataFrame, a
`magpyxl.Table`, a list of dicts, a list of lists, or a path to a
`.csv`/`.xlsx` file.

## Criteria syntax

Same as Spreadsheet: a bare number or string means equality; prefix with
`>`, `<`, `>=`, `<=`, `<>` for comparisons; use `*` (any run of characters)
or `?` (exactly one character) for text wildcards. Text matching is
case-insensitive, same as Spreadsheet.

## Design philosophy (current phase)

The primary goal of magpyxl is correctness and consistency. Every function aims to match Excel's behavior as closely as possible while providing a clean and predictable Python API.

Performance optimizations are added only after correctness is validated, ensuring speed never comes at the cost of reliability.

## What's Next

Our vision is to make magpyxl a practical, everyday toolkit for working with Excel-like operations in Python.

Future releases will introduce many more useful functions that solve common real-world data tasks while keeping the API simple and intuitive. We believe powerful tools shouldn't require complicated code, so simplicity, consistency, and Excel-like familiarity will remain our guiding principles.

Whether you're an experienced Python developer or someone with limited programming experience, our goal is to make magpyxl easy to learn, easy to use, and reliable for day-to-day data analysis and automation.

## Roadmap

- Additional Excel-compatible functions
- Text and string functions
- Date and time functions
- Mathematical and statistical functions
- Performance improvements
- SIMD acceleration
- Broader file format support

## Contributing

Issues and pull requests are welcome.
