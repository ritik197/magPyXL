"""
magpyxl — Excel-style formulas (SUM, VLOOKUP, XLOOKUP, SUMIFS, COUNTIFS ...)
backed by a compiled Rust core.

Works the same way whether your data is a plain Python list, a pandas
DataFrame/Series, a polars DataFrame/Series, or a .csv/.xlsx file on disk:
same syntax everywhere, same result type as what came in.

This layer intentionally keeps ONE code path per function (no dual
fast/slow dispatch). It's correct, simple, and easy to reason about.
Performance tuning happens later, once every function's behaviour is
confirmed — see the project README for the release plan.
"""

from __future__ import annotations

import csv as _csv
import os as _os
from typing import Any, Iterable, List, Optional, Union

from magpyxl._core import (
    sum_values as _sum_values,
    average_values as _average_values,
    count_values as _count_values,
    countif_values as _countif_values,
    sumif_values as _sumif_values,
    averageif_values as _averageif_values,
    countifs_values as _countifs_values,
    sumifs_values as _sumifs_values,
    countif_vec_values as _countif_vec_values,
    sumif_vec_values as _sumif_vec_values,
    averageif_vec_values as _averageif_vec_values,
    countifs_vec_values as _countifs_vec_values,
    sumifs_vec_values as _sumifs_vec_values,
    vlookup_values as _vlookup_values,
    xlookup_values as _xlookup_values,
    vlookup_many_values as _vlookup_many_values,
    xlookup_many_values as _xlookup_many_values,
    xlookup_many_indices as _xlookup_many_indices,
    lookupifs_indices_values as _lookupifs_indices_values,
)

# Sentinel so `if_not_found=None` (user explicitly wants None back on a
# miss) can be told apart from "if_not_found wasn't passed at all" (should
# raise, Excel #N/A style). Plain `None` can't do this on its own, since
# "not given" and "given as None" would otherwise look identical.
_UNSET = object()
_NONE_FALLBACK = object()  # placeholder that safely survives the Rust boundary


def _rust_fallback(if_not_found):
    """Translate a user-facing if_not_found value into what the Rust core
    should receive: None -> raise on miss; a real None fallback needs the
    sentinel so it doesn't collide with "no fallback given"."""
    if if_not_found is _UNSET:
        return None
    if if_not_found is None:
        return _NONE_FALLBACK
    return if_not_found


def _unwrap_fallback(value):
    return None if value is _NONE_FALLBACK else value

__version__ = "0.1.0"
__all__ = [
    "SUM", "AVERAGE", "COUNT", "COUNTIF", "SUMIF", "AVERAGEIF",
    "COUNTIFS", "SUMIFS", "VLOOKUP", "XLOOKUP", "LOOKUPIFS",
    "Table", "read_table",
]


# ---------------------------------------------------------------------------
# Universal input layer: normalize plain lists, numpy arrays, pandas
# Series/columns, or polars Series into a plain Python list before the
# data reaches the Rust core.
# ---------------------------------------------------------------------------

def _to_list(data: Any) -> list:
    if data is None:
        return []
    if isinstance(data, list):
        return data
    if isinstance(data, tuple):
        return list(data)
    if isinstance(data, Table):
        raise TypeError(
            "Pass a single column (e.g. table['Sales']), not the whole Table, "
            "where a range/array is expected."
        )
    tolist = getattr(data, "tolist", None)  # pandas Series, numpy array
    if callable(tolist):
        try:
            return list(tolist())
        except Exception:
            pass
    to_list = getattr(data, "to_list", None)  # polars Series
    if callable(to_list):
        try:
            return list(to_list())
        except Exception:
            pass
    try:
        return list(data)
    except TypeError:
        return [data]


def _maybe_num(value):
    """Turn numeric-looking strings into real numbers; leave everything else as-is."""
    if value is None:
        return None
    if isinstance(value, str):
        s = value.strip()
        if s == "":
            return None
        try:
            if "." not in s and "e" not in s.lower():
                return int(s)
            return float(s)
        except ValueError:
            return value
    return value


# ---------------------------------------------------------------------------
# Lightweight standalone table — lets magpyxl load CSV/XLSX data with no
# pandas or polars dependency required.
# ---------------------------------------------------------------------------

class Table:
    """A minimal columnar table used for standalone CSV/XLSX loading."""

    def __init__(self, columns: dict):
        self.columns = columns
        self._names = list(columns.keys())

    def __getitem__(self, key: str) -> list:
        return self.columns[key]

    def __contains__(self, key: str) -> bool:
        return key in self.columns

    @property
    def column_names(self) -> list:
        return list(self._names)

    def __len__(self) -> int:
        return len(next(iter(self.columns.values()), []))

    def rows(self) -> List[list]:
        n = len(self)
        return [[self.columns[name][i] for name in self._names] for i in range(n)]

    def __repr__(self) -> str:
        return f"Table(columns={self._names}, rows={len(self)})"

    @classmethod
    def from_csv(cls, path: str) -> "Table":
        with open(path, newline="", encoding="utf-8") as f:
            rows = list(_csv.reader(f))
        if not rows:
            return cls({})
        header, *data_rows = rows
        cols = {
            name: [_maybe_num(r[i]) if i < len(r) else None for r in data_rows]
            for i, name in enumerate(header)
        }
        return cls(cols)

    @classmethod
    def from_xlsx(cls, path: str, sheet: Optional[str] = None) -> "Table":
        try:
            import openpyxl
        except ImportError as exc:
            raise ImportError(
                "Reading .xlsx files needs openpyxl. Install with: pip install openpyxl"
            ) from exc
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[sheet] if sheet else wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return cls({})
        header, *data_rows = rows
        cols = {name: [r[i] for r in data_rows] for i, name in enumerate(header)}
        return cls(cols)


def read_table(path: str, sheet: Optional[str] = None) -> Table:
    """Load a .csv or .xlsx file straight into a magpyxl Table (no pandas/polars needed)."""
    ext = _os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return Table.from_csv(path)
    if ext in (".xlsx", ".xlsm"):
        return Table.from_xlsx(path, sheet=sheet)
    raise ValueError(f"Unsupported file type for read_table: {path!r}")


# ---------------------------------------------------------------------------
# Table adapter for VLOOKUP — normalizes pandas/polars DataFrame, magpyxl
# Table, list-of-dicts, list-of-lists, or a csv/xlsx path into plain rows.
# ---------------------------------------------------------------------------

def _load_rows_and_columns(data: Any):
    """Returns (rows: list[list], column_names: list[str] | None)."""
    if isinstance(data, str):
        tbl = read_table(data)
        return tbl.rows(), tbl.column_names
    if isinstance(data, Table):
        return data.rows(), data.column_names
    module = type(data).__module__
    if module.startswith("pandas") and hasattr(data, "values"):
        return data.values.tolist(), list(data.columns)
    if module.startswith("polars") and hasattr(data, "rows"):
        return data.rows(), list(data.columns)
    if isinstance(data, list):
        if data and isinstance(data[0], dict):
            names = list(data[0].keys())
            return [[row.get(n) for n in names] for row in data], names
        return data, None
    raise TypeError(f"Unsupported table type for VLOOKUP: {type(data)!r}")


def _resolve_col_index(col_index: Union[int, str], column_names: Optional[list]) -> int:
    if isinstance(col_index, str):
        if column_names is None:
            raise ValueError(
                "col_index was given as a column name, but this table has no column names."
            )
        try:
            return column_names.index(col_index) + 1
        except ValueError:
            raise KeyError(f"Column {col_index!r} not found. Available: {column_names}")
    return col_index


def _is_array_like(x: Any) -> bool:
    """True for lists/tuples/numpy arrays/pandas Series/polars Series —
    i.e. 'do a whole column at once', as opposed to a single scalar value."""
    if isinstance(x, (list, tuple)):
        return True
    if isinstance(x, (str, bytes)):
        return False
    module = type(x).__module__
    if module.startswith("pandas") or module.startswith("polars") or module == "numpy":
        return hasattr(x, "__len__")
    return False


def _wrap_like(origin: Any, values: list, index=None):
    """Wrap a plain list of results back into the same ecosystem as `origin`,
    so the output can be chained straight back into pandas/polars code —
    whatever type went in comes back out."""
    module = type(origin).__module__
    if module.startswith("pandas"):
        import pandas as pd
        return pd.Series(values, index=index if index is not None else getattr(origin, "index", None))
    if module.startswith("polars"):
        import polars as pl
        return pl.Series(values)
    if module == "numpy":
        import numpy as np
        return np.array(values, dtype=object)
    if isinstance(origin, tuple):
        return tuple(values)
    return list(values)


# ---------------------------------------------------------------------------
# Multi-column return support (LOOKUPIFS' return_range, XLOOKUP's
# return_array can each be a single column OR a whole sub-table).
# ---------------------------------------------------------------------------

_NO_MATCH = object()  # internal marker: "this row had zero matching indices"


def _resolve_return_columns(data: Any):
    """Normalize a return_range/return_array argument into
    (is_multi_column, {column_name_or_None: [values...]}).

    A pandas/polars DataFrame (or a dict of columns) is multi-column; a
    Series/list/tuple/array is a single column, keyed under None.
    """
    if isinstance(data, dict):
        return True, {k: _to_list(v) for k, v in data.items()}
    module = type(data).__module__
    if (module.startswith("pandas") or module.startswith("polars")) and hasattr(data, "columns"):
        cols = list(data.columns)
        return True, {c: _to_list(data[c]) for c in cols}
    return False, {None: _to_list(data)}


def _pick_by_mode(indices: list, values: list, mode: str):
    """Given the matching row-indices for one output row and one column's
    values, apply mode ('first'/'last'/'all') — or _NO_MATCH if indices is
    empty. 'all' joins every matching value into one comma-separated string
    (per column — each requested column gets its own independent string)."""
    if not indices:
        return _NO_MATCH
    if mode == "first":
        return values[indices[0]]
    if mode == "last":
        return values[indices[-1]]
    if mode == "all":
        return ", ".join(str(values[i]) for i in indices)
    raise ValueError("mode must be 'first', 'last', or 'all'")


def _build_scalar_multi_result(row: dict):
    """Scalar lookup + multiple return columns -> a pandas Series (index =
    column names), matching 'one row of the vectorized DataFrame result'.
    Falls back to a plain dict if pandas isn't installed."""
    try:
        import pandas as pd
        return pd.Series(row)
    except ImportError:
        return row


def _build_vector_multi_result(columns: dict, origin: Any, index=None):
    """Vectorized lookup + multiple return columns -> a DataFrame in the
    same ecosystem as `origin` (pandas/polars), or a dict of lists if
    neither is available."""
    module = type(origin).__module__ if origin is not None else ""
    if module.startswith("pandas"):
        import pandas as pd
        return pd.DataFrame(columns, index=index)
    if module.startswith("polars"):
        import polars as pl
        return pl.DataFrame(columns)
    return columns


# ---------------------------------------------------------------------------
# Public Excel-style API
#
# Every function here is ONE simple, correct implementation: normalize
# input -> call the Rust core -> return. No fast/slow dual paths yet.
# ---------------------------------------------------------------------------

def SUM(values: Iterable) -> float:
    """SUM(range) — adds all numeric values, ignoring text/blank cells."""
    return _sum_values(_to_list(values))


def AVERAGE(values: Iterable) -> float:
    """AVERAGE(range) — mean of numeric values, ignoring text/blank cells."""
    return _average_values(_to_list(values))


def COUNT(values: Iterable) -> int:
    """COUNT(range) — counts numeric cells only (Excel semantics)."""
    return _count_values(_to_list(values))


def COUNTIF(range_: Iterable, criteria: Any):
    """COUNTIF(range, criteria) — e.g. COUNTIF(sales, '>100'), COUNTIF(city, 'Delhi*').

    `criteria` can be a single value (returns a single count) or a whole
    column/list/Series (returns a matching column of counts — one count
    per criteria value, same type in/same type out). This is what lets
    you write `COUNTIF(table2["ID"], table1["ID"])` directly instead of
    looping over table1's rows yourself. The whole batch runs in one
    native Rust call (cached via a frequency map when every criteria is
    a plain equality check).
    """
    if _is_array_like(criteria):
        crit_list = _to_list(criteria)
        results = _countif_vec_values(_to_list(range_), crit_list)
        return _wrap_like(criteria, results, index=getattr(criteria, "index", None))
    return _countif_values(_to_list(range_), criteria)


def SUMIF(range_: Iterable, criteria: Any, sum_range: Optional[Iterable] = None):
    """SUMIF(range, criteria, sum_range=None) — sums sum_range where range meets criteria.

    `criteria` can be a single value or a whole column (same in/out
    behavior as COUNTIF above).
    """
    if _is_array_like(criteria):
        crit_list = _to_list(criteria)
        sr_list = _to_list(sum_range) if sum_range is not None else None
        results = _sumif_vec_values(_to_list(range_), crit_list, sr_list)
        return _wrap_like(criteria, results, index=getattr(criteria, "index", None))
    sr = _to_list(sum_range) if sum_range is not None else None
    return _sumif_values(_to_list(range_), criteria, sr)


def AVERAGEIF(range_: Iterable, criteria: Any, average_range: Optional[Iterable] = None):
    """AVERAGEIF(range, criteria, average_range=None). Same scalar-or-column criteria behavior.

    A criteria row with no matching numeric values comes back as NaN
    (Excel's #DIV/0! for that row) rather than failing the whole batch.
    """
    if _is_array_like(criteria):
        crit_list = _to_list(criteria)
        ar_list = _to_list(average_range) if average_range is not None else None
        raw = _averageif_vec_values(_to_list(range_), crit_list, ar_list)
        results = [r if r is not None else float("nan") for r in raw]
        return _wrap_like(criteria, results, index=getattr(criteria, "index", None))
    ar = _to_list(average_range) if average_range is not None else None
    return _averageif_values(_to_list(range_), criteria, ar)


def _prepare_ifs_vector(ranges: list, criteria: list):
    """If any criteria in an *IFS call is a whole column, build the
    (pairs, origin, index) needed for ONE call into the native vectorized
    Rust function. Scalar criteria are broadcast to the batch length (a
    cheap, small-N operation — nowhere near the cost of the range data
    itself) so Rust only has to handle one shape. Returns None when
    every criteria is a plain scalar (no vectorization needed).
    """
    vector_flags = [_is_array_like(c) for c in criteria]
    if not any(vector_flags):
        return None
    crit_lists = [_to_list(c) if v else None for c, v in zip(criteria, vector_flags)]
    lengths = {len(cl) for cl, v in zip(crit_lists, vector_flags) if v}
    if len(lengths) > 1:
        raise ValueError("All vectorized criteria columns must be the same length")
    n = lengths.pop()
    origin = next(c for c, v in zip(criteria, vector_flags) if v)
    index = getattr(origin, "index", None)
    pairs = [
        (_to_list(r), cl if v else [c] * n)
        for r, c, cl, v in zip(ranges, criteria, crit_lists, vector_flags)
    ]
    return pairs, origin, index


def COUNTIFS(*args):
    """COUNTIFS(range1, criteria1, range2, criteria2, ...) — AND across all pairs.

    Any criteria can be a single value or a whole column; if any column is
    given, you get back a matching column of counts, one per row — all
    computed in a single native Rust call.
    """
    if len(args) < 2 or len(args) % 2 != 0:
        raise ValueError("COUNTIFS needs range/criteria pairs, e.g. COUNTIFS(r1, c1, r2, c2, ...)")
    ranges = [args[i] for i in range(0, len(args), 2)]
    criteria = [args[i + 1] for i in range(0, len(args), 2)]

    vec = _prepare_ifs_vector(ranges, criteria)
    if vec is not None:
        pairs, origin, index = vec
        return _wrap_like(origin, _countifs_vec_values(pairs), index=index)

    pairs = [(_to_list(ranges[i]), criteria[i]) for i in range(len(ranges))]
    return _countifs_values(pairs)


def SUMIFS(sum_range: Iterable, *args):
    """SUMIFS(sum_range, range1, criteria1, range2, criteria2, ...) — AND across all pairs.

    Any criteria can be a single value or a whole column (same
    row-per-column behavior as COUNTIFS above).
    """
    if len(args) < 2 or len(args) % 2 != 0:
        raise ValueError("SUMIFS needs range/criteria pairs after sum_range")
    ranges = [args[i] for i in range(0, len(args), 2)]
    criteria = [args[i + 1] for i in range(0, len(args), 2)]

    vec = _prepare_ifs_vector(ranges, criteria)
    if vec is not None:
        pairs, origin, index = vec
        results = _sumifs_vec_values(_to_list(sum_range), pairs)
        return _wrap_like(origin, results, index=index)

    pairs = [(_to_list(ranges[i]), criteria[i]) for i in range(len(ranges))]
    return _sumifs_values(_to_list(sum_range), pairs)


def VLOOKUP(
    lookup_value: Any,
    table: Any,
    col_index: Union[int, str],
    range_lookup: bool = False,
    if_not_found: Any = _UNSET,
):
    """VLOOKUP(lookup_value, table, col_index, range_lookup=False, if_not_found=<raise>).

    `table` can be a pandas DataFrame, a polars DataFrame, a magpyxl Table,
    a list of dicts, a list of lists, or a path to a .csv/.xlsx file.
    `col_index` can be a 1-based column number (Excel-style) or a column name.

    `lookup_value` can be a single value (returns a single result) or a
    whole column/list/Series (returns a matching column/list/Series —
    the same type that went in comes back out, so it chains straight
    back into pandas/polars code).

    If `if_not_found` isn't given, a miss raises ValueError (Excel's
    #N/A). Pass any value — including `None` — to get that back instead
    of raising; `None` here means "return None", not "no default given".
    """
    rows, names = _load_rows_and_columns(table)
    idx = _resolve_col_index(col_index, names)

    if _is_array_like(lookup_value):
        values = _to_list(lookup_value)
        index = getattr(lookup_value, "index", None)
        fallback = _rust_fallback(if_not_found)
        result = _vlookup_many_values(values, rows, idx, range_lookup, fallback)
        result = [_unwrap_fallback(v) for v in result]
        return _wrap_like(lookup_value, result, index=index)

    try:
        return _vlookup_values(lookup_value, rows, idx, range_lookup)
    except ValueError:
        if if_not_found is not _UNSET:
            return if_not_found
        raise


def XLOOKUP(
    lookup_value: Any,
    lookup_array: Iterable,
    return_array: Any,
    if_not_found: Any = _UNSET,
):
    """XLOOKUP(lookup_value, lookup_array, return_array, if_not_found=<raise>).

    Same scalar-or-column auto-detection as VLOOKUP: pass a single value
    for a single result, or a whole column/Series to get a matching
    column/Series back (same type in, same type out).

    `return_array` can also be MULTIPLE columns at once (e.g.
    `master[["Salary", "City", "Dept"]]`):
      - scalar `lookup_value` -> a pandas Series (index = column names)
      - vectorized `lookup_value` -> a DataFrame (one row per lookup
        value, one column per requested field) — this is exactly what
        lets `summary[["Salary","City","Dept"]] = mx.XLOOKUP(...)` work.

    If `if_not_found` isn't given, a scalar miss raises ValueError; a miss
    inside a vectorized batch fills that row with None instead (so one
    missing row doesn't fail the whole batch — same precedent as
    AVERAGEIF's vectorized path). Pass `None` explicitly to get `None`
    back instead of raising, even for a scalar call.
    """
    la = _to_list(lookup_array)
    is_multi, columns = _resolve_return_columns(return_array)

    if is_multi:
        vectorized = _is_array_like(lookup_value)
        values = _to_list(lookup_value) if vectorized else [lookup_value]
        idx_results = _xlookup_many_indices(values, la)

        resolved = []
        for idx in idx_results:
            if idx is None:
                if if_not_found is _UNSET and not vectorized:
                    raise ValueError(
                        "XLOOKUP: value not found. Use if_not_found=... to specify a default value."
                    )
                resolved.append(_NO_MATCH)
            else:
                resolved.append(idx)

        out = {}
        for col_name, col_values in columns.items():
            out[col_name] = [
                (None if if_not_found is _UNSET else if_not_found) if idx is _NO_MATCH else col_values[idx]
                for idx in resolved
            ]

        if vectorized:
            index = getattr(lookup_value, "index", None)
            return _build_vector_multi_result(out, lookup_value, index=index)
        row = {k: v[0] for k, v in out.items()}
        return _build_scalar_multi_result(row)

    # Single-column path (unchanged from before multi-column support existed).
    ra = _to_list(return_array)
    fallback = _rust_fallback(if_not_found)

    if _is_array_like(lookup_value):
        values = _to_list(lookup_value)
        index = getattr(lookup_value, "index", None)
        result = _xlookup_many_values(values, la, ra, fallback)
        result = [_unwrap_fallback(v) for v in result]
        return _wrap_like(lookup_value, result, index=index)

    result = _xlookup_values(lookup_value, la, ra, fallback)
    return _unwrap_fallback(result)


def LOOKUPIFS(
    return_range: Any,
    *args,
    mode: str = "first",
    if_not_found: Any = _UNSET,
):
    """LOOKUPIFS(return_range, range1, criteria1, range2, criteria2, ..., mode="first", if_not_found=<raise>).

    Multi-criteria lookup: same AND-across-pairs matching as SUMIFS/COUNTIFS,
    but returns the actual matching value(s) from `return_range` instead of
    an aggregate — e.g.
        mx.LOOKUPIFS(df["Salary"], df["Department"], "IT", df["City"], "Delhi")

    Any criteria can be a single value or a whole column (vectorized, same
    as SUMIFS/COUNTIFS — pass a whole criteria column to get one result per
    row back).

    `return_range` can be a single column OR multiple columns at once (a
    DataFrame/dict of columns, e.g. `master[["Salary","City","Manager"]]`) —
    scalar call -> pandas Series (index = column names); vectorized call ->
    a DataFrame (one row per input row, one column per requested field).

    mode:
      "first" (default) — first matching row's value(s)
      "last"  — last matching row's value(s)
      "all"   — every matching value, comma-joined into one string; with
                multiple return columns, each column gets its OWN
                independent comma-joined string, not one merged string

    If `if_not_found` isn't given, a scalar call with no match raises
    ValueError; inside a vectorized batch, an unmatched row is filled with
    None instead (so one missing row doesn't fail the whole batch).
    """
    if mode not in ("first", "last", "all"):
        raise ValueError("mode must be 'first', 'last', or 'all'")
    if len(args) < 2 or len(args) % 2 != 0:
        raise ValueError(
            "LOOKUPIFS needs range/criteria pairs, e.g. LOOKUPIFS(return_range, r1, c1, r2, c2, ...)"
        )

    ranges = [args[i] for i in range(0, len(args), 2)]
    criteria = [args[i + 1] for i in range(0, len(args), 2)]

    vec = _prepare_ifs_vector(ranges, criteria)
    if vec is not None:
        pairs, origin, index = vec
        vectorized = True
    else:
        pairs = [(_to_list(r), [c]) for r, c in zip(ranges, criteria)]
        origin, index = None, None
        vectorized = False

    match_indices = _lookupifs_indices_values(pairs)  # one list of matching row-indices per output row

    resolved = []
    for idxs in match_indices:
        if not idxs:
            if if_not_found is _UNSET and not vectorized:
                raise ValueError(
                    "LOOKUPIFS: no matching row found. Use if_not_found=... to specify a default value."
                )
            resolved.append(_NO_MATCH)
        else:
            resolved.append(idxs)

    is_multi, columns = _resolve_return_columns(return_range)

    out = {}
    for col_name, col_values in columns.items():
        out[col_name] = [
            (None if if_not_found is _UNSET else if_not_found)
            if idxs is _NO_MATCH
            else _pick_by_mode(idxs, col_values, mode)
            for idxs in resolved
        ]

    if is_multi:
        if vectorized:
            return _build_vector_multi_result(out, origin, index=index)
        row = {k: v[0] for k, v in out.items()}
        return _build_scalar_multi_result(row)

    (single_col_values,) = out.values()
    if vectorized:
        return _wrap_like(origin, single_col_values, index=index)
    return single_col_values[0]


# ---------------------------------------------------------------------------
# Optional syntax sugar: df.mx.SUM("Sales"), df.mx.SUMIFS(...) — same engine,
# just lets you stay inside pandas/polars method-chaining style if you want.
# ---------------------------------------------------------------------------

def _make_accessor(base_cls):
    class _magpyxlAccessor(base_cls):
        def __init__(self, obj):
            self._obj = obj

        def _col(self, name_or_data):
            if isinstance(name_or_data, str) and name_or_data in self._obj.columns:
                return self._obj[name_or_data]
            return name_or_data

        def SUM(self, col):
            return SUM(self._col(col))

        def AVERAGE(self, col):
            return AVERAGE(self._col(col))

        def COUNT(self, col):
            return COUNT(self._col(col))

        def COUNTIF(self, col, criteria):
            return COUNTIF(self._col(col), criteria)

        def SUMIF(self, col, criteria, sum_col=None):
            return SUMIF(self._col(col), criteria, self._col(sum_col) if sum_col is not None else None)

        def AVERAGEIF(self, col, criteria, average_col=None):
            return AVERAGEIF(self._col(col), criteria, self._col(average_col) if average_col is not None else None)

        def COUNTIFS(self, *args):
            resolved = [self._col(a) if i % 2 == 0 else a for i, a in enumerate(args)]
            return COUNTIFS(*resolved)

        def SUMIFS(self, sum_col, *args):
            resolved = [self._col(sum_col)] + [self._col(a) if i % 2 == 0 else a for i, a in enumerate(args)]
            return SUMIFS(*resolved)

        def VLOOKUP(self, lookup_col, table, col_index, range_lookup=False, if_not_found=_UNSET):
            return VLOOKUP(self._col(lookup_col), table, col_index, range_lookup, if_not_found)

        def XLOOKUP(self, lookup_col, lookup_array, return_array, if_not_found=_UNSET):
            return XLOOKUP(self._col(lookup_col), lookup_array, return_array, if_not_found)

        def LOOKUPIFS(self, return_col, *args, mode="first", if_not_found=_UNSET):
            resolved = [self._col(a) if i % 2 == 0 else a for i, a in enumerate(args)]
            return LOOKUPIFS(self._col(return_col), *resolved, mode=mode, if_not_found=if_not_found)

    return _magpyxlAccessor


def _register_accessors():
    try:
        import pandas as pd
        pd.api.extensions.register_dataframe_accessor("mx")(_make_accessor(object))
    except ImportError:
        pass
    except Exception:
        pass  # already registered, or pandas version mismatch — non-fatal

    try:
        import polars as pl
        pl.api.register_dataframe_namespace("mx")(_make_accessor(object))
    except ImportError:
        pass
    except Exception:
        pass


_register_accessors()
