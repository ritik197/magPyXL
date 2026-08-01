# MagpieXL

Excel-style formulas — `SUM`, `AVERAGE`, `COUNT`, `COUNTIF`, `SUMIF`, `AVERAGEIF`,
`COUNTIFS`, `SUMIFS`, `VLOOKUP`, `XLOOKUP` — backed by a compiled Rust core,
callable with the exact same syntax whether your data is a plain Python list,
a pandas DataFrame, a polars DataFrame, or a `.csv`/`.xlsx` file on disk.

## Design philosophy (current phase)

This release prioritizes **correctness, Excel-compatibility, and a clean,
maintainable architecture** over raw performance. Every function is a single,
straightforward implementation — normalize input → call the Rust core →
return. No dual fast/slow code paths yet.

Performance optimization (zero-copy numeric fast paths, SIMD, etc.) is a
planned **later phase**, once every function's behavior has been tested and
confirmed correct. It's fine if this version isn't the fastest possible — it
is the most correct and easiest to reason about.

## Install

```bash
pip install magpiexl-0.1.0-cp312-cp312-manylinux_2_34_x86_64.whl
```

(Only Python 3.12 on manylinux x86_64 is built in this package. To build for
another Python version/platform, install Rust + `pip install maturin`, then
run `maturin build --release` from the project root.)

## Quick start

```python
import magpiexl as mx

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

# Or stay in pandas method-chaining style:
df.mx.SUM("Salary")
df.mx.SUMIFS("Salary", "Dept", "Eng")
```

### Works the same way with polars

```python
import polars as pl
pdf = pl.DataFrame({...})
mx.SUM(pdf["Salary"])
mx.VLOOKUP(pdf["Key"], other_pdf, "Price")   # returns a polars Series
pdf.mx.COUNTIF("Dept", "Eng")
```

### Standalone — no pandas/polars needed

```python
tbl = mx.read_table("sales.csv")     # or sales.xlsx
mx.SUM(tbl["Revenue"])
mx.VLOOKUP("Dave", "sales.csv", "Salary")   # path works directly too
```

## Function reference

| Function | Signature | Notes |
|---|---|---|
| `SUM` | `SUM(range)` | Ignores text/blank cells |
| `AVERAGE` | `AVERAGE(range)` | Same |
| `COUNT` | `COUNT(range)` | Counts numeric cells only (Excel semantics) |
| `COUNTIF` | `COUNTIF(range, criteria)` | Criteria: `10`, `">10"`, `"<=5"`, `"<>0"`, `"ab*"`, `"a?c"` |
| `SUMIF` | `SUMIF(range, criteria, sum_range=None)` | |
| `AVERAGEIF` | `AVERAGEIF(range, criteria, average_range=None)` | |
| `COUNTIFS` | `COUNTIFS(range1, criteria1, range2, criteria2, ...)` | AND across all pairs |
| `SUMIFS` | `SUMIFS(sum_range, range1, criteria1, ...)` | AND across all pairs |
| `VLOOKUP` | `VLOOKUP(lookup_value, table, col_index, range_lookup=False, if_not_found=None)` | `col_index`: 1-based number or column name; `lookup_value` can be scalar or a whole column |
| `XLOOKUP` | `XLOOKUP(lookup_value, lookup_array, return_array, if_not_found=None)` | Same scalar-or-column behavior |

`table` (for VLOOKUP) accepts: pandas DataFrame, polars DataFrame, a
`magpiexl.Table`, a list of dicts, a list of lists, or a path to a
`.csv`/`.xlsx` file.

## Criteria syntax

Same as Excel: a bare number or string means equality; prefix with
`>`, `<`, `>=`, `<=`, `<>` for comparisons; use `*` (any run of characters)
or `?` (exactly one character) for text wildcards. Text matching is
case-insensitive, same as Excel.

## Project layout

```
magpiexl/
├── Cargo.toml              # Rust crate config
├── pyproject.toml          # maturin build config (mixed python/rust layout)
├── src/lib.rs               # Rust core: value model, criteria parsing, all 10 functions
└── python/magpiexl/
    └── __init__.py          # Python adapter: input normalization, public API, .mx accessor
```

## What's next (later phase, not yet started)

Once every function above is tested and confirmed correct in real use:
- Zero-copy numeric fast paths for array-backed data (numpy/pandas/polars)
- Benchmarking against pandas on realistic workloads
- Expanding the criteria/lookup edge cases (multi-column VLOOKUP keys, etc.)
