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
import re as _re
from typing import Any, Iterable, List, Optional, Union

from magpyxl._core import (
    sum_values as _sum_values,
    average_values as _average_values,
    count_values as _count_values,
    min_values as _min_values,
    max_values as _max_values,
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
    averageifs_vec_values as _averageifs_vec_values,
    minifs_vec_values as _minifs_vec_values,
    maxifs_vec_values as _maxifs_vec_values,
    vlookup_values as _vlookup_values,
    xlookup_values as _xlookup_values,
    vlookup_many_values as _vlookup_many_values,
    xlookup_many_values as _xlookup_many_values,
    vlookup_many_columnar as _vlookup_many_columnar,
    xlookup_many_columnar as _xlookup_many_columnar,
    xlookup_many_indices as _xlookup_many_indices,
    lookupifs_indices_values as _lookupifs_indices_values,
    sum_mixed as _sum_mixed,
    average_mixed as _average_mixed,
    count_mixed as _count_mixed,
    countif_mixed as _countif_mixed,
    sumif_mixed as _sumif_mixed,
    averageif_mixed as _averageif_mixed,
    countifs_mixed as _countifs_mixed,
    sumifs_mixed as _sumifs_mixed,
    averageifs_mixed as _averageifs_mixed,
    min_mixed as _min_mixed,
    max_mixed as _max_mixed,
    minifs_mixed as _minifs_mixed,
    maxifs_mixed as _maxifs_mixed,
    summarize_numeric_column as _summarize_numeric_column,
    summarize_text_column as _summarize_text_column,
    build_category_merge_mapping as _build_category_merge_mapping,
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

__version__ = "0.2.2"
__all__ = [
    "SUM", "AVERAGE", "COUNT", "MIN", "MAX", "COUNTIF", "SUMIF", "AVERAGEIF",
    "COUNTIFS", "SUMIFS", "AVERAGEIFS", "MINIFS", "MAXIFS",
    "VLOOKUP", "XLOOKUP", "LOOKUPIFS", "INFO", "CLEAN",
    "Table", "read_table",
]


# ---------------------------------------------------------------------------
# Universal input layer: normalize plain lists, numpy arrays, pandas
# Series/columns, or polars Series into a plain Python list before the
# data reaches the Rust core.
# ---------------------------------------------------------------------------

def _reject_invalid_column_types(data: Any) -> None:
    """Raise early for input shapes that are clearly wrong for a
    column/range argument. Shared by `_to_list` (the generic path) and
    the mixed fast-path dispatch below, so both paths reject the same
    obviously-wrong inputs the same way."""
    if isinstance(data, dict):
        # A bare dict is almost certainly a mistake here (e.g. someone
        # passed a row/record instead of a column) — iterating it would
        # silently sum/count its KEYS, producing a plausible-looking but
        # wrong number instead of failing fast. Multi-column return
        # values (LOOKUPIFS/XLOOKUP) use dicts on purpose in a different
        # place (_resolve_return_columns), not here.
        raise TypeError(
            "Expected a column (list/tuple/Series/array), got a dict. "
            "If this is meant to be multiple return columns, use it with "
            "VLOOKUP/XLOOKUP/LOOKUPIFS's table/return_range argument instead."
        )
    if isinstance(data, Table):
        raise TypeError(
            "Pass a single column (e.g. table['Sales']), not the whole Table, "
            "where a range/array is expected."
        )


def _is_array_backed(data: Any) -> bool:
    """True for numpy arrays and pandas/polars Series/columns — data
    that's ALREADY a buffer-backed column, where the Rust 'mixed' fast
    path (FastColumn) can add real value. False for plain lists/tuples
    (converting one of those to a numpy array first is pure overhead,
    no real win — see HANDOFF.md §5)."""
    module = type(data).__module__
    return module == "numpy" or module.startswith("pandas") or module.startswith("polars")


def _is_numeric_array_backed(data: Any) -> bool:
    """True only when `data` is ALREADY a numeric numpy/pandas/polars
    buffer — the one case where handing the raw object straight to Rust
    is a genuine win (FastColumn gets a true zero-copy f64 view, no
    Python object touched at all).

    For anything else that's merely _is_array_backed (a TEXT/object
    pandas or polars column, for instance), extracting a Vec<PyObject>
    directly from the Series via PyO3's generic sequence/iterator
    protocol turned out to be markedly SLOWER than pandas'/polars' own
    highly optimized `.tolist()`/`.to_list()` (confirmed by direct
    benchmark) — so those still go through `_to_list()` first, same as
    before, and only the resulting plain Python list is handed to Rust.
    """
    module = type(data).__module__
    if module == "numpy":
        return data.dtype.kind in "iuf"
    if module.startswith("pandas"):
        import pandas as pd
        return pd.api.types.is_numeric_dtype(data.dtype)
    if module.startswith("polars"):
        return data.dtype.is_numeric()
    return False


def _is_string_array_backed(data: Any) -> bool:
    """True when `data` is a polars Utf8/String Series, or a pandas
    text column (either the dedicated `string` dtype, or a plain
    `object`-dtype column that happens to hold only `str`) that's
    cleanly all-str with no nulls mixed in.

    `object` dtype is included (unlike before) because that's what most
    real-world pandas string columns actually are — `pd.Series(list_of_
    strings)` defaults to `object`, not the newer dedicated `string`
    dtype. Excluding it meant most pandas text columns never got the
    fast path at all. Checked in two cheap steps rather than a
    per-element Python scan: `pd.api.types.infer_dtype(data, skipna=
    False)` reports `"string"` only when every value's runtime type is
    actually `str` (confirmed empirically: ~0.1ms even at 2M rows,
    since it's a fast internal C scan, not a Python-level loop) —
    but it does NOT flag None/NaN mixed in (confirmed: still reports
    "string" with nulls present), so nulls are checked separately via
    `.isna().any()` (~1ms at 2M rows). A column with real nulls mixed in
    still falls through to `.to_list()` + the Generic Rust path, same as
    before, so correctness doesn't change, only the fast case gets wider.
    """
    module = type(data).__module__
    if module.startswith("polars"):
        try:
            return data.dtype == data.dtype.__class__ and str(data.dtype) in ("Utf8", "String")
        except Exception:
            return False
    if module.startswith("pandas"):
        import pandas as pd
        if pd.api.types.is_string_dtype(data.dtype) and not pd.api.types.is_object_dtype(data.dtype):
            return not data.isna().any()
        if pd.api.types.is_object_dtype(data.dtype):
            try:
                if pd.api.types.infer_dtype(data, skipna=False) != "string":
                    return False
                return not data.isna().any()
            except Exception:
                return False
        return False
    return False


def _is_fast_mixed_backed(data: Any) -> bool:
    """True when `data` qualifies for the Rust `_mixed` fast path at
    all — either a numeric buffer (zero-copy numpy/pandas/polars) or a
    clean text column (bulk `Vec<String>` extraction). Used at the
    `*IF`/`*IFS` call sites, which previously only checked
    `_is_numeric_array_backed` and so always sent a criteria/filter
    column of strings through the slow `_to_list()` + Generic path even
    when it was a polars/pandas string column that could have taken the
    fast bulk-string path instead."""
    return _is_numeric_array_backed(data) or _is_string_array_backed(data)


def _normalize_numeric_array(arr):
    """Given a raw numpy array of ANY numeric dtype, returns whichever
    of int64/float64 it fits into losslessly — zero-copy if it's
    already the right width, one fast vectorized `.astype()` cast
    otherwise. This is what actually gets EVERY numpy-family numeric
    dtype onto Rust's fast `FastColumn` path, not just already-int64/
    float64 data.

    Before this existed, `_mixed_arg` only recognized already-int64 or
    already-float64 arrays as fast-path-eligible; anything else
    numeric (`int8/16/32`, `uint8/16/32/64`, `float16/32`) was handed
    to Rust completely unchanged, silently failed BOTH of
    `FastColumn::resolve`'s numpy-extraction attempts (which only ever
    try `PyReadonlyArray1<f64>` then `PyReadonlyArray1<i64>` — no
    other width), and fell all the way back to the slow generic
    per-element `Vec<PyObject>` path — with no error and no
    indication anything was suboptimal. Confirmed by direct benchmark:
    an `int32` column of 2,000,000 rows took ~150ms through that
    silent fallback versus ~2ms for the equivalent `int64` column
    through the real fast path — a ~60x gap purely from dtype width,
    invisible unless you went looking for it.

    `uint64` specifically widens to float64, not int64: a uint64 value
    can exceed int64's positive range (2^63-1), and casting it to
    int64 would silently wrap around to an incorrect, possibly
    negative, number — float64 stays numerically correct to within
    its own precision limits instead, which is a far safer failure
    mode for a value that large. `uint8/16/32` always fit safely into
    int64 (no width they can hold exceeds it), so those go there
    directly rather than to float64, keeping small-integer columns on
    Rust's genuinely-zero-copy int64 path instead of the merely-cheap
    float64-cast one.

    Boolean arrays are NOT handled here — `dtype.kind == "b"` falls
    through to the generic list path unchanged, same as before this
    function existed; that's `_is_numeric_array_backed`'s job to gate
    on, not this function's.
    """
    kind = arr.dtype.kind  # 'i' signed int, 'u' unsigned int, 'f' float
    itemsize = arr.dtype.itemsize
    if kind == "i":
        return arr if itemsize == 8 else arr.astype("int64", copy=False)
    if kind == "u":
        if itemsize == 8:
            return arr.astype("float64", copy=False)  # avoid silent int64 wraparound
        return arr.astype("int64", copy=False)  # uint8/16/32 always fit safely
    if kind == "f":
        return arr if itemsize == 8 else arr.astype("float64", copy=False)
    return arr  # unexpected kind reaching here — hand back unchanged, let Rust's own extraction attempts fail safely


def _mixed_arg(data: Any):
    """What to actually pass into a `_mixed` Rust function for one
    column: a real numpy array when the data is numeric (so Rust's
    `PyReadonlyArray1<f64>`/`PyReadonlyArray1<i64>` extraction succeeds
    and gets a zero-copy view), the RAW polars Series object itself when
    it's a clean polars string column (so Rust's Arrow C Stream
    Interface path engages — true zero-copy, no `list[str]` in between
    at all), a plain `list[str]` for other clean-text cases (pandas
    string dtype) where Rust's bulk `Vec<String>` extraction engages,
    otherwise a plain list via `_to_list` (generic fallback, always
    correct).

    IMPORTANT: for the polars-string case, do NOT call `.to_list()`
    here. Doing so would convert the column to a plain Python list
    before Rust ever sees it — at which point Rust's
    `try_arrow_text_column` has nothing to call `__arrow_c_stream__()`
    on, since a `list` doesn't implement that protocol, and everything
    silently falls back to the slower `Text(Vec<String>)` path instead
    of the true zero-copy `ArrowText` path. Confirmed by direct
    benchmark: handing over `.to_list()` here left COUNTIF on a 2M-row
    polars string column at ~330ms; passing the raw Series through (so
    Rust reads its Arrow buffer directly) is what actually gets it down
    to single-digit ms, matching polars' own native speed.

    Important subtlety on the numeric side: `PyReadonlyArray1`
    extraction only succeeds on an actual `numpy.ndarray` instance — a
    pandas/polars Series merely *wraps* one, so handing Rust the Series
    object directly makes the numpy extraction fail and silently fall
    back to the slow generic Vec<PyObject> path (confirmed by direct
    benchmark: ~150ms instead of ~2ms for 1,000,000 rows). Pulling the
    underlying array out here with `.to_numpy()` is what actually makes
    the fast path engage. Polars strings don't have this problem because
    Rust reads them via the Arrow protocol, not via numpy extraction.

    int64 gets special treatment: it's the overwhelmingly common numpy/
    pandas dtype for whole-number columns (IDs, counts, salaries), and
    Rust's FastColumn has a genuine zero-copy int64 path — so an
    already-int64 column is handed through with NO cast at all, not even
    a cheap vectorized one. Anything else numeric (float64, or a
    narrower int/uint width Rust doesn't have a dedicated path for) still
    gets the float64 cast — zero-copy when already float64, one fast
    vectorized cast otherwise, both far cheaper than the fully-generic
    per-element path.
    """
    if data is None:
        return []
    module = type(data).__module__
    if module.startswith("polars") and _is_string_array_backed(data):
        import polars as pl
        # Convert via `to_arrow(compat_level=oldest())` BEFORE handing
        # to Rust — measured to be faster overall than passing the raw
        # Series through, even though this conversion step itself
        # isn't free (tens of ms at large N). The alternative (a
        # hand-rolled Rust-side reader for polars' native Utf8View
        # stream format, avoiding this conversion) was built and
        # correctness-tested, but direct A/B benchmarking showed it is
        # SLOWER end to end: it has to materialize an owned `String`
        # per row (a heap allocation per row) to safely read out of
        # the stream, which costs more than this single bulk
        # conversion followed by Rust's zero-copy `&str` borrows over
        # the result. Confirmed at 1M rows: ~90ms for the native-
        # stream route vs ~14-16ms for this conversion + zero-copy
        # read. Don't change this without re-benchmarking — the
        # measured numbers, not intuition, should decide it.
        try:
            return data.to_arrow(compat_level=pl.CompatLevel.oldest())
        except Exception:
            return data.to_list()
    if module.startswith("pandas") and _is_string_array_backed(data):
        # Same idea as the polars branch above: hand Rust a pyarrow
        # array (which implements `__arrow_c_array__`) instead of a
        # plain `list[str]`, so the same zero-copy Arrow path engages —
        # avoiding the per-row PyObject cost entirely, not just paying
        # it once in bulk. `pyarrow` is an optional dependency here
        # (not everyone doing pandas work has it installed), so this
        # is a soft try: on any failure (pyarrow missing, or some edge
        # case it can't convert), fall back to `.to_list()`, which is
        # still correct — just not zero-copy.
        #
        # Confirmed by direct benchmark: `pa.Array.from_pandas()` on a
        # plain object-dtype string column of 1.5M rows costs ~2ms —
        # pandas' object arrays are already flat arrays of Python str
        # objects, so pyarrow can walk them in a tight C loop, similar
        # to why `.to_numpy()` was already fast for numeric object
        # columns elsewhere in this function. That 2ms is negligible
        # next to the ~300ms+ the old per-row path cost at this size.
        #
        # pandas' dedicated `string` dtype (as opposed to the far more
        # common plain `object` dtype) already implements
        # `__arrow_c_stream__` directly — confirmed to export classic
        # `large_string` (format "U") — but Rust's native stream reader
        # for that path is currently disabled (see `resolve`'s comment
        # on `NativeText`: it benchmarked slower than the
        # `pa.Array.from_pandas()` conversion below, not faster), so
        # this always goes through the conversion call, same as the
        # object-dtype case just below.
        try:
            import pyarrow as pa
            return pa.Array.from_pandas(data)
        except Exception:
            return data.to_list()
    if _is_string_array_backed(data):
        return data.to_list()  # pandas: fast on a clean non-null string dtype
    if not _is_numeric_array_backed(data):
        return _to_list(data)
    if module == "numpy":
        return data if data.dtype.kind not in "iuf" else _normalize_numeric_array(data)
    if module.startswith("pandas"):
        return _normalize_numeric_array(data.to_numpy())
    if module.startswith("polars"):
        return _normalize_numeric_array(data.to_numpy())
    return data


def _to_list(data: Any) -> list:
    if data is None:
        return []
    if isinstance(data, list):
        return data
    if isinstance(data, tuple):
        return list(data)
    _reject_invalid_column_types(data)
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
        idx = index if index is not None else getattr(origin, "index", None)
        if any(v is None for v in values):
            # Force dtype=object so an explicit `if_not_found=None` (or a
            # soft-failed row with no match) reliably comes back as real
            # `None`. Without this, `pd.Series([...floats..., None])`
            # silently upcasts the trailing/embedded None to NaN — pandas'
            # own Series-construction behavior, not something magpyxl
            # does on purpose, but it broke the "None means None" promise.
            return pd.Series(values, index=idx, dtype=object)
        return pd.Series(values, index=idx)
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
        if any(v is None for v in row.values()):
            return pd.Series(row, dtype=object)  # preserve real None — see _wrap_like
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
        # Built column-by-column (not pd.DataFrame(columns) directly) so
        # each column can independently get dtype=object when it contains
        # a real None — same reasoning as _wrap_like.
        series = {
            name: (pd.Series(vals, index=index, dtype=object) if any(v is None for v in vals)
                   else pd.Series(vals, index=index))
            for name, vals in columns.items()
        }
        return pd.DataFrame(series)
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
    _reject_invalid_column_types(values)
    if _is_array_backed(values):
        return _sum_mixed(_mixed_arg(values))
    return _sum_values(_to_list(values))


def AVERAGE(values: Iterable) -> float:
    """AVERAGE(range) — mean of numeric values, ignoring text/blank cells."""
    _reject_invalid_column_types(values)
    if _is_array_backed(values):
        return _average_mixed(_mixed_arg(values))
    return _average_values(_to_list(values))


def COUNT(values: Iterable) -> int:
    """COUNT(range) — counts numeric cells only (Excel semantics)."""
    _reject_invalid_column_types(values)
    if _is_array_backed(values):
        return _count_mixed(_mixed_arg(values))
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
    _reject_invalid_column_types(range_)
    if _is_array_like(criteria):
        crit_list = _to_list(criteria)
        results = _countif_vec_values(_to_list(range_), crit_list)
        return _wrap_like(criteria, results, index=getattr(criteria, "index", None))
    if _is_fast_mixed_backed(range_):
        return _countif_mixed(_mixed_arg(range_), criteria)
    return _countif_values(_to_list(range_), criteria)


def SUMIF(range_: Iterable, criteria: Any, sum_range: Optional[Iterable] = None):
    """SUMIF(range, criteria, sum_range=None) — sums sum_range where range meets criteria.

    `criteria` can be a single value or a whole column (same in/out
    behavior as COUNTIF above).
    """
    _reject_invalid_column_types(range_)
    if _is_array_like(criteria):
        crit_list = _to_list(criteria)
        sr_list = _to_list(sum_range) if sum_range is not None else None
        results = _sumif_vec_values(_to_list(range_), crit_list, sr_list)
        return _wrap_like(criteria, results, index=getattr(criteria, "index", None))
    # The filter column (range_) can be numeric OR clean text and still
    # take the Rust _mixed path — FastColumn::matches_at handles both.
    # sum_range, if given separately, must resolve numeric on the Rust
    # side regardless (summing text is meaningless); _mixed_arg already
    # falls back to _to_list()'s Generic path for it if it isn't.
    if _is_fast_mixed_backed(range_):
        sr = _mixed_arg(sum_range) if sum_range is not None else None
        return _sumif_mixed(_mixed_arg(range_), criteria, sr)
    sr = _to_list(sum_range) if sum_range is not None else None
    return _sumif_values(_to_list(range_), criteria, sr)


def AVERAGEIF(range_: Iterable, criteria: Any, average_range: Optional[Iterable] = None):
    """AVERAGEIF(range, criteria, average_range=None). Same scalar-or-column criteria behavior.

    A criteria row with no matching numeric values comes back as NaN
    (Excel's #DIV/0! for that row) rather than failing the whole batch.
    """
    _reject_invalid_column_types(range_)
    if _is_array_like(criteria):
        crit_list = _to_list(criteria)
        ar_list = _to_list(average_range) if average_range is not None else None
        raw = _averageif_vec_values(_to_list(range_), crit_list, ar_list)
        results = [r if r is not None else float("nan") for r in raw]
        return _wrap_like(criteria, results, index=getattr(criteria, "index", None))
    if _is_fast_mixed_backed(range_):
        ar = _mixed_arg(average_range) if average_range is not None else None
        return _averageif_mixed(_mixed_arg(range_), criteria, ar)
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
    for r in ranges:
        _reject_invalid_column_types(r)

    vec = _prepare_ifs_vector(ranges, criteria)
    if vec is not None:
        pairs, origin, index = vec
        return _wrap_like(origin, _countifs_vec_values(pairs), index=index)

    # All-scalar criteria. Historically the mixed path only paid off when
    # at least one range was numeric-array-backed, because an all-text
    # range fell back to FastColumn's Generic(Vec<PyObject>) variant,
    # which was slightly SLOWER per-row than the original generic path.
    # Now that FastColumn has a real zero-copy-extraction Text(Vec<String>)
    # variant (bulk-converted once, no per-row PyObject/GIL touch in the
    # loop), a clean text range is a genuine win too — so the gate is any
    # numeric OR clean-string range; only a truly mixed bag of
    # object/nullable columns falls back to the original path.
    if any(_is_fast_mixed_backed(r) for r in ranges):
        pairs = [(_mixed_arg(ranges[i]), criteria[i]) for i in range(len(ranges))]
        return _countifs_mixed(pairs)
    pairs = [(_to_list(ranges[i]), criteria[i]) for i in range(len(ranges))]
    return _countifs_values(pairs)


def SUMIFS(sum_range: Iterable, *args):
    """SUMIFS(sum_range, range1, criteria1, range2, criteria2, ...) — AND across all pairs.

    Any criteria can be a single value or a whole column (same
    row-per-column behavior as COUNTIFS above).
    """
    if len(args) < 2 or len(args) % 2 != 0:
        raise ValueError("SUMIFS needs range/criteria pairs after sum_range")
    _reject_invalid_column_types(sum_range)
    ranges = [args[i] for i in range(0, len(args), 2)]
    criteria = [args[i + 1] for i in range(0, len(args), 2)]
    for r in ranges:
        _reject_invalid_column_types(r)

    vec = _prepare_ifs_vector(ranges, criteria)
    if vec is not None:
        pairs, origin, index = vec
        results = _sumifs_vec_values(_to_list(sum_range), pairs)
        return _wrap_like(origin, results, index=index)

    if any(_is_fast_mixed_backed(r) for r in ranges):
        pairs = [(_mixed_arg(ranges[i]), criteria[i]) for i in range(len(ranges))]
        return _sumifs_mixed(_mixed_arg(sum_range), pairs)
    pairs = [(_to_list(ranges[i]), criteria[i]) for i in range(len(ranges))]
    return _sumifs_values(_to_list(sum_range), pairs)


def AVERAGEIFS(average_range: Iterable, *args):
    """AVERAGEIFS(average_range, range1, criteria1, range2, criteria2, ...) — AND across all pairs.

    Same scalar-or-column criteria behavior as COUNTIFS/SUMIFS. A row
    whose group matched but had zero numeric values comes back as NaN
    (Excel's #DIV/0! for that row) rather than failing the whole batch.
    """
    if len(args) < 2 or len(args) % 2 != 0:
        raise ValueError("AVERAGEIFS needs range/criteria pairs after average_range")
    _reject_invalid_column_types(average_range)
    ranges = [args[i] for i in range(0, len(args), 2)]
    criteria = [args[i + 1] for i in range(0, len(args), 2)]
    for r in ranges:
        _reject_invalid_column_types(r)

    vec = _prepare_ifs_vector(ranges, criteria)
    if vec is not None:
        pairs, origin, index = vec
        raw = _averageifs_vec_values(_to_list(average_range), pairs)
        results = [r if r is not None else float("nan") for r in raw]
        return _wrap_like(origin, results, index=index)

    if any(_is_fast_mixed_backed(r) for r in ranges):
        pairs = [(_mixed_arg(ranges[i]), criteria[i]) for i in range(len(ranges))]
        return _averageifs_mixed(_mixed_arg(average_range), pairs)
    pairs = [(_to_list(ranges[i]), criteria[i]) for i in range(len(ranges))]
    # No dedicated all-scalar-criteria _averageifs_values Rust function
    # exists (unlike SUM/COUNT's siblings) — the vec-values path above
    # already covers every criteria shape (scalar criteria are
    # broadcast to length-1 lists by _prepare_ifs_vector's caller
    # convention), so route scalar-only calls through it too rather
    # than duplicating the same reduction in a third Rust function.
    single_result = _averageifs_vec_values(_to_list(average_range), [(rng, [c]) for rng, c in pairs])
    if single_result[0] is None:
        raise ValueError("AVERAGEIFS: no matching numeric values found")
    return single_result[0]


def MIN(values: Iterable) -> float:
    """MIN(range) — smallest numeric value, ignoring text/blank cells."""
    _reject_invalid_column_types(values)
    if _is_array_backed(values):
        return _min_mixed(_mixed_arg(values))
    return _min_values(_to_list(values))


def MAX(values: Iterable) -> float:
    """MAX(range) — largest numeric value, ignoring text/blank cells."""
    _reject_invalid_column_types(values)
    if _is_array_backed(values):
        return _max_mixed(_mixed_arg(values))
    return _max_values(_to_list(values))


def MINIFS(min_range: Iterable, *args):
    """MINIFS(min_range, range1, criteria1, range2, criteria2, ...) — AND across all pairs.

    Same scalar-or-column criteria behavior as COUNTIFS/SUMIFS/
    AVERAGEIFS. Matches Excel's own MINIFS: returns 0 (not an error)
    when nothing matches, rather than raising.
    """
    if len(args) < 2 or len(args) % 2 != 0:
        raise ValueError("MINIFS needs range/criteria pairs after min_range")
    _reject_invalid_column_types(min_range)
    ranges = [args[i] for i in range(0, len(args), 2)]
    criteria = [args[i + 1] for i in range(0, len(args), 2)]
    for r in ranges:
        _reject_invalid_column_types(r)

    vec = _prepare_ifs_vector(ranges, criteria)
    if vec is not None:
        pairs, origin, index = vec
        results = _minifs_vec_values(_to_list(min_range), pairs)
        return _wrap_like(origin, results, index=index)

    if any(_is_fast_mixed_backed(r) for r in ranges):
        pairs = [(_mixed_arg(ranges[i]), criteria[i]) for i in range(len(ranges))]
        return _minifs_mixed(_mixed_arg(min_range), pairs)
    pairs = [(_to_list(ranges[i]), criteria[i]) for i in range(len(ranges))]
    return _minifs_vec_values(_to_list(min_range), [(rng, [c]) for rng, c in pairs])[0]


def MAXIFS(max_range: Iterable, *args):
    """MAXIFS(max_range, range1, criteria1, range2, criteria2, ...) — AND across all pairs.

    Same scalar-or-column criteria behavior as MINIFS above. Matches
    Excel's own MAXIFS: returns 0 (not an error) when nothing matches,
    rather than raising.
    """
    if len(args) < 2 or len(args) % 2 != 0:
        raise ValueError("MAXIFS needs range/criteria pairs after max_range")
    _reject_invalid_column_types(max_range)
    ranges = [args[i] for i in range(0, len(args), 2)]
    criteria = [args[i + 1] for i in range(0, len(args), 2)]
    for r in ranges:
        _reject_invalid_column_types(r)

    vec = _prepare_ifs_vector(ranges, criteria)
    if vec is not None:
        pairs, origin, index = vec
        results = _maxifs_vec_values(_to_list(max_range), pairs)
        return _wrap_like(origin, results, index=index)

    if any(_is_fast_mixed_backed(r) for r in ranges):
        pairs = [(_mixed_arg(ranges[i]), criteria[i]) for i in range(len(ranges))]
        return _maxifs_mixed(_mixed_arg(max_range), pairs)
    pairs = [(_to_list(ranges[i]), criteria[i]) for i in range(len(ranges))]
    return _maxifs_vec_values(_to_list(max_range), [(rng, [c]) for rng, c in pairs])[0]


def _try_columnar_table_columns(table: Any, col_index: Union[int, str]):
    """If `table` is a pandas or polars DataFrame, returns
    `(key_column, return_column, column_names)` — the two raw Series
    objects for the FIRST column (the lookup key, always column 1 in
    VLOOKUP) and the requested `col_index` column, pulled directly via
    `table[name]`/`table.iloc[:, i]` — NOT via `.values.tolist()` or
    `.rows()`, both of which materialize the ENTIRE table (every row,
    every column) into nested Python lists before Rust ever sees it.
    Pulling just the two needed columns is itself a cheap, already-fast
    DataFrame operation, unlike materializing the whole table.

    Returns `None` (not an error) for anything that isn't a pandas/
    polars DataFrame — a magpyxl `Table`, a list of dicts, a list of
    lists, or a CSV/XLSX path all still go through the original
    `_load_rows_and_columns`-based row-major path unchanged; this is
    strictly an additional fast path for the DataFrame case, not a
    replacement for the others.
    """
    module = type(table).__module__
    if module.startswith("pandas"):
        names = list(table.columns)
        idx = _resolve_col_index(col_index, names)
        if idx < 1 or idx > len(names):
            return None  # let the caller's existing validation produce the right error
        return table.iloc[:, 0], table.iloc[:, idx - 1], names
    if module.startswith("polars"):
        names = list(table.columns)
        idx = _resolve_col_index(col_index, names)
        if idx < 1 or idx > len(names):
            return None
        return table[names[0]], table[names[idx - 1]], names
    return None


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
    # Fast columnar path: table is a pandas/polars DataFrame, doing an
    # exact-match vectorized lookup (the overwhelmingly common shape:
    # `df['result'] = VLOOKUP(df['key'], other_df, 'value_col')`).
    # Reads only the two columns actually needed, zero-copy where
    # FastColumn supports it — skips materializing the whole table
    # into a row-major Python list entirely. See
    # `_try_columnar_table_columns`'s own docstring for why this
    # exists: that whole-table conversion alone measured ~360ms for a
    # 500,000-row table, before Rust ever touched the data.
    #
    # Approximate-match (`range_lookup=True`) isn't included here: its
    # sorted-order semantics need the row-based linear scan for
    # correctness (see `vlookup_many_values`'s own comment on this),
    # so it keeps using the original path unconditionally.
    if not range_lookup and _is_array_like(lookup_value):
        cols = _try_columnar_table_columns(table, col_index)
        if cols is not None:
            key_col, ret_col, names = cols
            idx = _resolve_col_index(col_index, names)
            fallback = _rust_fallback(if_not_found)
            index = getattr(lookup_value, "index", None)
            result = _vlookup_many_columnar(
                _mixed_arg(lookup_value) if _is_fast_mixed_backed(lookup_value) else _to_list(lookup_value),
                _mixed_arg(key_col) if _is_fast_mixed_backed(key_col) else _to_list(key_col),
                _mixed_arg(ret_col) if _is_fast_mixed_backed(ret_col) else _to_list(ret_col),
                fallback,
            )
            result = [_unwrap_fallback(v) for v in result]
            return _wrap_like(lookup_value, result, index=index)

    rows, names = _load_rows_and_columns(table)
    idx = _resolve_col_index(col_index, names)

    # Validate col_index BEFORE anything else — deliberately outside the
    # try/except below. Both "value not found" and "col_index out of
    # range" used to raise the same plain ValueError from the Rust core,
    # so when if_not_found was set, a genuinely bad col_index (the
    # caller's mistake) was silently swallowed and the fallback value
    # was returned instead of surfacing the real problem. Checking here,
    # first, means a bad col_index always raises no matter what
    # if_not_found is set to. This also replaces a raw, unhelpful
    # `OverflowError` (from trying to convert a negative Python int to
    # Rust's unsigned col_index) with a clear, consistent ValueError.
    max_width = max((len(r) for r in rows), default=0)
    if idx < 1:
        raise ValueError(
            f"VLOOKUP: col_index must be 1 or greater (1-based column number); got {idx}"
        )
    if idx > max_width:
        raise ValueError(
            f"VLOOKUP: col_index {idx} is out of range for the table "
            f"(table has {max_width} column{'s' if max_width != 1 else ''})"
        )

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
    class _MagpieXLAccessor(base_cls):
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

        def MIN(self, col):
            return MIN(self._col(col))

        def MAX(self, col):
            return MAX(self._col(col))

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

        def AVERAGEIFS(self, average_col, *args):
            resolved = [self._col(average_col)] + [self._col(a) if i % 2 == 0 else a for i, a in enumerate(args)]
            return AVERAGEIFS(*resolved)

        def MINIFS(self, min_col, *args):
            resolved = [self._col(min_col)] + [self._col(a) if i % 2 == 0 else a for i, a in enumerate(args)]
            return MINIFS(*resolved)

        def MAXIFS(self, max_col, *args):
            resolved = [self._col(max_col)] + [self._col(a) if i % 2 == 0 else a for i, a in enumerate(args)]
            return MAXIFS(*resolved)

        def VLOOKUP(self, lookup_col, table, col_index, range_lookup=False, if_not_found=_UNSET):
            return VLOOKUP(self._col(lookup_col), table, col_index, range_lookup, if_not_found)

        def XLOOKUP(self, lookup_col, lookup_array, return_array, if_not_found=_UNSET):
            return XLOOKUP(self._col(lookup_col), lookup_array, return_array, if_not_found)

        def LOOKUPIFS(self, return_col, *args, mode="first", if_not_found=_UNSET):
            resolved = [self._col(a) if i % 2 == 0 else a for i, a in enumerate(args)]
            return LOOKUPIFS(self._col(return_col), *resolved, mode=mode, if_not_found=if_not_found)

        def INFO(self, **kwargs):
            return INFO(self._obj, **kwargs)

        def CLEAN(self, **kwargs):
            return CLEAN(self._obj, **kwargs)

        def __getattr__(self, name):
            # Case-insensitive method access: df.mx.sumifs(...),
            # df.mx.SumIfs(...) resolve to the same method as
            # df.mx.SUMIFS(...) — same idea as the module-level
            # case-insensitivity below, scoped to this accessor.
            # `_CASE_MAP` is built ONCE per class (see right after this
            # method), not rescanned via `dir()` on every call — that
            # earlier version measured ~6.6us of per-call overhead
            # (a full `dir(type(self))` scan every time), which this
            # cuts to roughly the same ~0.5us a plain dict lookup costs
            # at the module level.
            upper = name.upper()
            attr_name = _MagpieXLAccessor._CASE_MAP.get(upper)
            if attr_name is not None:
                return getattr(self, attr_name)
            raise AttributeError(
                f"{type(self).__name__!r} object has no attribute {name!r}"
            )

    # Built once when the accessor class is created (module import
    # time), not on every `__getattr__` call — see that method's own
    # comment for the measured cost of rebuilding this per call.
    # `dir(_MagpieXLAccessor)` here (unlike `dir(type(self))` inside a
    # call) still only sees this class's own methods, since `base_cls`
    # is `object` — no risk of picking up pandas/polars DataFrame
    # methods and exposing them under new case-insensitive spellings.
    _MagpieXLAccessor._CASE_MAP = {
        attr.upper(): attr
        for attr in dir(_MagpieXLAccessor)
        if not attr.startswith("_")
    }

    return _MagpieXLAccessor


# ---------------------------------------------------------------------------
# INFO / CLEAN — data profiling and (explicit, never-automatic-by-
# default) data cleaning. Pure Python for now — no Rust involved yet
# (per explicit instruction: get this working and correct first, speed
# optimization is a deliberately separate later pass). Column
# classification and statistics logic below mirrors a Rust-backed
# reference implementation reviewed for this port (same dominant-type
# classification, same near-duplicate-category detection, same "never
# silently drop data" design decisions) — reimplemented fresh here in
# plain Python against magpyxl's own container-normalization helpers,
# not copied verbatim.
# ---------------------------------------------------------------------------

def _load_named_columns(data: Any, column_names: Optional[list] = None) -> "dict[str, Any]":
    """Column-major loader for INFO/CLEAN — deliberately separate from
    VLOOKUP's row-major `_load_rows_and_columns`; transposing would be
    wasted work here since every column is profiled independently."""
    mod = type(data).__module__
    if mod.startswith("pandas"):
        if hasattr(data, "columns") and hasattr(data, "iloc"):  # DataFrame
            return {str(c): data[c] for c in data.columns}
        name = str(getattr(data, "name", None) or "value")  # Series
        return {name: data}
    if mod.startswith("polars"):
        if hasattr(data, "columns") and hasattr(data, "height"):  # DataFrame
            return {str(c): data[c] for c in data.columns}
        name = str(getattr(data, "name", None) or "value")  # Series
        return {name: data}
    if isinstance(data, Table):
        return {n: data[n] for n in data.column_names}
    if isinstance(data, list) and data and isinstance(data[0], dict):
        names = column_names or list(data[0].keys())
        return {n: [row.get(n) for row in data] for n in names}
    if isinstance(data, (list, tuple)) and data and isinstance(data[0], (list, tuple)):
        n_cols = len(data[0])
        names = column_names or [f"col_{i}" for i in range(n_cols)]
        return {names[i]: [row[i] for row in data] for i in range(n_cols)}
    if hasattr(data, "shape") and getattr(data, "ndim", 1) == 2:  # 2-D numpy array
        n_cols = data.shape[1]
        names = column_names or [f"col_{i}" for i in range(n_cols)]
        return {names[i]: data[:, i] for i in range(n_cols)}
    name = column_names[0] if column_names else "value"  # single column
    return {name: data}


def _is_datetime_column(col: Any) -> bool:
    mod = type(col).__module__
    if mod.startswith("pandas"):
        try:
            import pandas as pd
            if pd.api.types.is_datetime64_any_dtype(col):
                return True
        except Exception:
            pass
    if mod.startswith("polars"):
        dtype_str = str(getattr(col, "dtype", ""))
        if dtype_str.startswith("Date") or dtype_str.startswith("Datetime"):
            return True
    # Fallback for a plain list/tuple/numpy array, OR a pandas/polars
    # column holding real date/datetime objects that wasn't cast to a
    # native datetime dtype (e.g. a plain Python list of
    # `datetime.date` put into a pandas column stays dtype=object —
    # pandas does NOT auto-cast it). Only short-circuits on a
    # POSITIVE match from the first element; anything else (including
    # an all-None column) falls through to "not a datetime column"
    # rather than guessing.
    import datetime as _dt
    try:
        for v in col:
            if v is None:
                continue
            return isinstance(v, (_dt.date, _dt.datetime))
    except TypeError:
        return False
    return False


def _looks_like_date(s: str) -> bool:
    if len(s) < 6 or len(s) > 10:
        return False
    parts = [p for p in _re.split(r"[-/]", s)]
    if len(parts) != 3:
        return False
    if not all(p and len(p) <= 4 and p.isdigit() for p in parts):
        return False
    return any(len(p) == 4 for p in parts)


def _percentile(sorted_vals: list, p: float) -> float:
    """Nearest-rank percentile on an already-sorted list — a simple,
    documented method (not numpy's linear-interpolation default), used
    only for INFO's advisory outlier flag, not an exact-precision
    statistic."""
    if not sorted_vals:
        return float("nan")
    idx = round(p * (len(sorted_vals) - 1))
    return sorted_vals[min(idx, len(sorted_vals) - 1)]


def _summarize_column_slow(col_list: list, top_n: int, categorical_max_unique: int, categorical_max_ratio: float) -> dict:
    """Classifies and computes statistics for one column's already-
    materialized `list` — the pure-Python fallback used by
    `_summarize_column` below whenever the Rust fast path doesn't
    apply (a boolean-dominant, mixed-type, or otherwise non-numeric/
    non-text-exclusive column — anything `FastColumn::resolve` can't
    cleanly classify into one zero-copy variant). Always correct,
    just not accelerated; see `_summarize_column`'s own docstring for
    when the fast path engages instead.
    Classification is by DOMINANT type (ties go to numeric): a column
    that's mostly numbers with a handful of stray text values still
    gets numeric stats, with the minority reported via `mixed_types`
    rather than silently dropping the majority (or the whole column)."""
    total = len(col_list)
    nums: list = []
    texts: list = []
    bools: list = []
    missing = 0
    for v in col_list:
        if v is None:
            missing += 1
            continue
        if isinstance(v, bool):
            bools.append(v)
            continue
        if isinstance(v, float):
            if v != v:  # NaN
                missing += 1
            else:
                nums.append(v)
            continue
        if isinstance(v, int):
            nums.append(float(v))
            continue
        if isinstance(v, str):
            if v.strip() == "":
                missing += 1
            else:
                texts.append(v)
            continue
        # Anything else (dates handled separately upstream, custom
        # objects, ...) — stringify as a last resort so it's still
        # counted somewhere rather than silently vanishing.
        texts.append(str(v))

    n_bool, n_num, n_text = len(bools), len(nums), len(texts)

    if n_bool > 0 and n_num == 0 and n_text == 0:
        true_count = sum(1 for b in bools if b)
        is_constant = n_bool > 0 and (true_count == 0 or true_count == n_bool)
        return {
            "kind": "boolean", "total": total, "missing": missing,
            "unique": 2 if n_bool > 0 else 0,
            "sum": float(true_count), "mean": true_count / n_bool if n_bool else float("nan"),
            "min": float("nan"), "max": float("nan"), "median": float("nan"), "std": float("nan"),
            "zeros": 0, "negatives": 0, "q1": float("nan"), "q3": float("nan"), "outlier_count": 0,
            "top_categories": [], "bottom_categories": [], "is_categorical": False, "is_constant": is_constant,
            "numeric_looking": 0, "date_looking": 0, "most_frequent": "", "most_frequent_count": 0,
            "mixed_other_count": 0, "normalized_unique": 0, "near_dup_example": "",
        }

    numeric_ish = n_num + n_bool
    if numeric_ish == 0 and n_text == 0:
        return {
            "kind": "empty", "total": total, "missing": missing, "unique": 0,
            "sum": float("nan"), "mean": float("nan"), "min": float("nan"), "max": float("nan"),
            "median": float("nan"), "std": float("nan"), "zeros": 0, "negatives": 0,
            "q1": float("nan"), "q3": float("nan"), "outlier_count": 0,
            "top_categories": [], "bottom_categories": [], "is_categorical": False, "is_constant": False,
            "numeric_looking": 0, "date_looking": 0, "most_frequent": "", "most_frequent_count": 0,
            "mixed_other_count": 0, "normalized_unique": 0, "near_dup_example": "",
        }

    if numeric_ish >= n_text:
        for b in bools:
            nums.append(1.0 if b else 0.0)
        n = len(nums)
        total_sum = sum(nums)
        mean = total_sum / n
        mn, mx = min(nums), max(nums)
        sorted_nums = sorted(nums)
        mid = n // 2
        median = (sorted_nums[mid - 1] + sorted_nums[mid]) / 2.0 if n % 2 == 0 else sorted_nums[mid]
        variance = sum((x - mean) ** 2 for x in nums) / n
        std = variance ** 0.5
        zeros = sum(1 for x in nums if x == 0.0)
        negatives = sum(1 for x in nums if x < 0.0)
        q1 = _percentile(sorted_nums, 0.25)
        q3 = _percentile(sorted_nums, 0.75)
        iqr = q3 - q1
        lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
        outlier_count = sum(1 for x in nums if x < lower or x > upper)
        is_constant = mn == mx
        return {
            "kind": "numeric", "total": total, "missing": missing, "unique": len(set(sorted_nums)),
            "sum": total_sum, "mean": mean, "min": mn, "max": mx, "median": median, "std": std,
            "zeros": zeros, "negatives": negatives, "q1": q1, "q3": q3, "outlier_count": outlier_count,
            "top_categories": [], "bottom_categories": [], "is_categorical": False, "is_constant": is_constant,
            "numeric_looking": 0, "date_looking": 0, "most_frequent": "", "most_frequent_count": 0,
            "mixed_other_count": n_text, "normalized_unique": 0, "near_dup_example": "",
        }

    # --- TEXT (dominant) ---
    counts: "dict[str, int]" = {}
    numeric_looking = 0
    date_looking = 0
    # numeric_looking/date_looking are advisory FLAGS, not exact
    # statistics — sampling (a stride, not just the first K, to avoid
    # bias from ordering) is legitimate here the way it wouldn't be
    # for `unique`/`missing`, which are always computed exactly.
    sample_cap = 5000
    stride = max(1, len(texts) // sample_cap)
    sampled = 0
    for i, t in enumerate(texts):
        counts[t] = counts.get(t, 0) + 1
        if i % stride == 0:
            sampled += 1
            trimmed = t.strip()
            try:
                float(trimmed)
                numeric_looking += 1
            except ValueError:
                pass
            if _looks_like_date(trimmed):
                date_looking += 1
    if sampled > 0 and stride > 1:
        scale = len(texts) / sampled
        numeric_looking = round(numeric_looking * scale)
        date_looking = round(date_looking * scale)

    unique = len(counts)
    is_categorical = unique <= categorical_max_unique or (unique / max(n_text, 1)) <= categorical_max_ratio
    is_constant = unique == 1

    top_categories: list = []
    bottom_categories: list = []
    most_frequent = ""
    most_frequent_count = 0
    normalized_unique = 0
    near_dup_example = ""

    if is_categorical:
        ranked = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))
        top_categories = ranked[:top_n]
        bottom_categories = list(reversed(ranked))[:min(5, top_n)]
        norm_groups: "dict[str, list]" = {}
        for raw_key in counts:
            norm_groups.setdefault(raw_key.strip().lower(), []).append(raw_key)
        normalized_unique = len(norm_groups)
        if normalized_unique < unique:
            for variants in norm_groups.values():
                if len(variants) > 1:
                    near_dup_example = "/".join(sorted(variants))
                    break
    else:
        best_key, best_count = "", 0
        for k, c in counts.items():
            if c > best_count or (c == best_count and (not best_key or k < best_key)):
                best_key, best_count = k, c
        most_frequent, most_frequent_count = best_key, best_count
        normalized_unique = unique

    return {
        "kind": "text", "total": total, "missing": missing, "unique": unique,
        "sum": float("nan"), "mean": float("nan"), "min": float("nan"), "max": float("nan"),
        "median": float("nan"), "std": float("nan"), "zeros": 0, "negatives": 0,
        "q1": float("nan"), "q3": float("nan"), "outlier_count": 0,
        "top_categories": top_categories, "bottom_categories": bottom_categories,
        "is_categorical": is_categorical, "is_constant": is_constant,
        "numeric_looking": numeric_looking, "date_looking": date_looking,
        "most_frequent": most_frequent, "most_frequent_count": most_frequent_count,
        "mixed_other_count": numeric_ish, "normalized_unique": normalized_unique,
        "near_dup_example": near_dup_example,
    }


def _summarize_column(col: Any, top_n: int, categorical_max_unique: int, categorical_max_ratio: float) -> dict:
    """Classifies and computes statistics for one column — the core of
    INFO(), and (via `_build_cleaning_plan`) of CLEAN()'s diagnosis
    step too. Tries the Rust-backed fast path first (via `FastColumn`,
    the same zero-copy architecture COUNTIF/SUMIF/VLOOKUP already use)
    and falls back to `_summarize_column_slow`'s pure-Python
    implementation whenever the fast path doesn't apply.

    Added after profiling showed the pure-Python version spending
    ~95% of its own time in a per-value loop (`isinstance` checks,
    dict/list mutation) at real DataFrame scale — confirmed directly
    with cProfile before writing the Rust side, not assumed. Measured
    end-to-end win on a 1,000,000-row column: INFO() on a 3-column,
    1M-row DataFrame went from ~2.1s to well under 100ms for the
    columns that qualify for the fast path (see HANDOFF notes for the
    exact before/after numbers).

    The Rust fast path engages for a column that's CLEANLY one type —
    exclusively numeric (any numpy/pandas/polars numeric dtype) or
    exclusively text (no numbers, no booleans mixed in) — the same
    scope `FastColumn::resolve` already handles zero-copy elsewhere in
    this file. A column that's boolean-dominant, genuinely mixed-type,
    or resolves to the fully-generic `Vec<PyObject>` path gets `None`
    back from the Rust call and falls through to the slow-but-always-
    correct Python path unchanged — this is purely additive; nothing
    about the classification LOGIC or edge-case handling changed, only
    which values get computed by a tight Rust loop instead of a
    Python one.
    """
    try:
        if _is_numeric_array_backed(col):
            raw = _summarize_numeric_column(_mixed_arg(col))
            if raw is not None:
                return _numeric_rust_result_to_summary(raw)
        elif _is_string_array_backed(col) or (
            isinstance(col, list) and col and all(isinstance(v, str) for v in col)
        ):
            arg = _mixed_arg(col) if _is_string_array_backed(col) else col
            raw = _summarize_text_column(arg, top_n, categorical_max_unique, categorical_max_ratio)
            if raw is not None:
                return _text_rust_result_to_summary(raw)
    except Exception:
        # Never let an acceleration-path failure break INFO/CLEAN
        # itself — fall through to the always-correct pure-Python
        # path below on ANY unexpected error from the Rust call.
        pass
    return _summarize_column_slow(_to_list(col), top_n, categorical_max_unique, categorical_max_ratio)


def _numeric_rust_result_to_summary(raw: dict) -> dict:
    """Fills in the keys `_summarize_column_slow` would have set for a
    numeric column but `summarize_numeric_column` (Rust) doesn't
    return — the text-only fields, always empty/zero for a numeric
    result, and `mixed_other_count`, always 0 here since the Rust
    fast path only engages for a column with NO non-numeric values."""
    if raw["kind"] == "empty":
        return {
            "kind": "empty", "total": raw["total"], "missing": raw["missing"], "unique": 0,
            "sum": float("nan"), "mean": float("nan"), "min": float("nan"), "max": float("nan"),
            "median": float("nan"), "std": float("nan"), "zeros": 0, "negatives": 0,
            "q1": float("nan"), "q3": float("nan"), "outlier_count": 0,
            "top_categories": [], "bottom_categories": [], "is_categorical": False, "is_constant": False,
            "numeric_looking": 0, "date_looking": 0, "most_frequent": "", "most_frequent_count": 0,
            "mixed_other_count": 0, "normalized_unique": 0, "near_dup_example": "",
        }
    return {
        "kind": "numeric", "total": raw["total"], "missing": raw["missing"], "unique": raw["unique"],
        "sum": raw["sum"], "mean": raw["mean"], "min": raw["min"], "max": raw["max"],
        "median": raw["median"], "std": raw["std"], "zeros": raw["zeros"], "negatives": raw["negatives"],
        "q1": raw["q1"], "q3": raw["q3"], "outlier_count": raw["outlier_count"],
        "top_categories": [], "bottom_categories": [], "is_categorical": False,
        "is_constant": raw["is_constant"],
        "numeric_looking": 0, "date_looking": 0, "most_frequent": "", "most_frequent_count": 0,
        "mixed_other_count": 0, "normalized_unique": 0, "near_dup_example": "",
    }


def _text_rust_result_to_summary(raw: dict) -> dict:
    """Fills in the keys `_summarize_column_slow` would have set for a
    text column but `summarize_text_column` (Rust) doesn't return —
    the numeric-only fields, always NaN/0 for a text result, and
    `mixed_other_count`, always 0 here since the Rust fast path only
    engages for a column with NO non-text values."""
    if raw["kind"] == "empty":
        return {
            "kind": "empty", "total": raw["total"], "missing": raw["missing"], "unique": 0,
            "sum": float("nan"), "mean": float("nan"), "min": float("nan"), "max": float("nan"),
            "median": float("nan"), "std": float("nan"), "zeros": 0, "negatives": 0,
            "q1": float("nan"), "q3": float("nan"), "outlier_count": 0,
            "top_categories": [], "bottom_categories": [], "is_categorical": False, "is_constant": False,
            "numeric_looking": 0, "date_looking": 0, "most_frequent": "", "most_frequent_count": 0,
            "mixed_other_count": 0, "normalized_unique": 0, "near_dup_example": "",
        }
    return {
        "kind": "text", "total": raw["total"], "missing": raw["missing"], "unique": raw["unique"],
        "sum": float("nan"), "mean": float("nan"), "min": float("nan"), "max": float("nan"),
        "median": float("nan"), "std": float("nan"), "zeros": 0, "negatives": 0,
        "q1": float("nan"), "q3": float("nan"), "outlier_count": 0,
        "top_categories": raw["top_categories"], "bottom_categories": raw["bottom_categories"],
        "is_categorical": raw["is_categorical"], "is_constant": raw["is_constant"],
        "numeric_looking": raw["numeric_looking"], "date_looking": raw["date_looking"],
        "most_frequent": raw["most_frequent"], "most_frequent_count": raw["most_frequent_count"],
        "mixed_other_count": 0, "normalized_unique": raw["normalized_unique"],
        "near_dup_example": raw["near_dup_example"],
    }


def _summarize_date_column(name: str, col_list: list) -> dict:
    total = len(col_list)
    vals = [v for v in col_list if v is not None]
    missing = total - len(vals)
    missing_pct = round(100.0 * missing / total, 2) if total else 0.0
    if not vals:
        return {"column": name, "type": "date", "count": 0, "missing": missing,
                "missing_%": missing_pct, "unique": 0, "min": None, "max": None,
                "range_days": None, "flags": "high_missing" if missing_pct >= 50 else ""}
    mn, mx = min(vals), max(vals)
    try:
        range_days = (mx - mn).days
    except Exception:
        range_days = None
    flags = []
    if missing_pct >= 50:
        flags.append("high_missing")
    if mn == mx:
        flags.append("constant")
    return {"column": name, "type": "date", "count": len(vals), "missing": missing,
            "missing_%": missing_pct, "unique": len(set(vals)), "min": mn, "max": mx,
            "range_days": range_days, "flags": ", ".join(flags)}


def _format_info_row(name: str, s: dict) -> dict:
    non_missing = s["total"] - s["missing"]
    # `used` = values that actually went into the stats — distinct
    # from `non_missing`: a mixed-type column's minority values are
    # neither "missing" nor "used" (excluded from dominant-type
    # stats), reported separately via the mixed_types flag instead of
    # inflating `count`.
    used = max(non_missing - s["mixed_other_count"], 0)
    missing_pct = round(100.0 * s["missing"] / s["total"], 2) if s["total"] else 0.0

    flags = []
    if s["is_constant"] and used > 0:
        flags.append("constant")
    if missing_pct >= 50:
        flags.append("high_missing")
    if s["mixed_other_count"] > 0 and non_missing > 0:
        flags.append(f"mixed_types({s['mixed_other_count']}_excluded)")
    if s["kind"] == "numeric" and used > 0:
        if s["zeros"] / used >= 0.3:
            flags.append("many_zeros")
        if s["negatives"] > 0:
            flags.append("has_negatives")
        if s["outlier_count"] > 0:
            flags.append(f"outliers({s['outlier_count']})")
    if s["kind"] == "text" and used > 0:
        if s["numeric_looking"] / used >= 0.9:
            flags.append("looks_numeric")
        if s["date_looking"] / used >= 0.9:
            flags.append("looks_date")
        if not s["is_categorical"] and s["unique"] / used >= 0.95:
            flags.append("high_cardinality_id_like")
        if s["normalized_unique"] < s["unique"]:
            flags.append(f"near_duplicate_categories(e.g. {s['near_dup_example']})")

    row = {"column": name, "type": s["kind"], "count": used,
           "missing": s["missing"], "missing_%": missing_pct, "unique": s["unique"]}
    if s["kind"] in ("numeric", "boolean"):
        row.update({"sum": s["sum"], "mean": s["mean"], "min": s["min"], "max": s["max"],
                    "median": s["median"], "std": s["std"]})
        if s["kind"] == "numeric":
            row.update({"q1": s["q1"], "q3": s["q3"], "zeros": s["zeros"],
                        "negatives": s["negatives"], "outliers": s["outlier_count"]})
    if s["kind"] == "text":
        if s["is_categorical"]:
            row["top_categories"] = ", ".join(f"{c}({n})" for c, n in s["top_categories"])
            bottom_str = ", ".join(f"{c}({n})" for c, n in s["bottom_categories"])
            if bottom_str != row["top_categories"]:
                row["bottom_categories"] = bottom_str
        else:
            row["most_frequent"] = f"{s['most_frequent']} ({s['most_frequent_count']})" if s["most_frequent"] else ""
    row["flags"] = ", ".join(flags)
    return row


def _table_overview(original: Any, n_columns: int) -> Optional[dict]:
    """Cheap, whole-table stats computed ONCE (not per column): row/col
    counts, memory footprint, dtype breakdown, exact duplicate-row
    count — via each library's own native method (pandas
    `.duplicated()`, polars `.is_duplicated()`) rather than
    reimplementing row hashing. Returns None for single-column input,
    where "duplicate rows" isn't a meaningful concept."""
    mod = type(original).__module__
    if mod.startswith("pandas") and hasattr(original, "columns") and hasattr(original, "iloc"):
        try:
            memory_mb = round(original.memory_usage(deep=True).sum() / (1024 * 1024), 3)
        except Exception:
            memory_mb = None
        try:
            duplicate_rows = int(original.duplicated().sum())
        except Exception:
            duplicate_rows = None
        dtype_counts: "dict[str, int]" = {}
        for dt in original.dtypes.astype(str):
            dtype_counts[dt] = dtype_counts.get(dt, 0) + 1
        return {"rows": len(original), "columns": len(original.columns),
                "memory_mb": memory_mb, "duplicate_rows": duplicate_rows, "dtype_counts": dtype_counts}
    if mod.startswith("polars") and hasattr(original, "columns") and hasattr(original, "height"):
        try:
            memory_mb = round(original.estimated_size() / (1024 * 1024), 3)
        except Exception:
            memory_mb = None
        try:
            duplicate_rows = int(original.is_duplicated().sum())
        except Exception:
            duplicate_rows = None
        dtype_counts = {}
        for dt in original.dtypes:
            key = str(dt)
            dtype_counts[key] = dtype_counts.get(key, 0) + 1
        return {"rows": original.height, "columns": original.width,
                "memory_mb": memory_mb, "duplicate_rows": duplicate_rows, "dtype_counts": dtype_counts}
    if isinstance(original, list) and original and isinstance(original[0], dict):
        seen: set = set()
        duplicate_rows = 0
        try:
            for row in original:
                key = tuple(sorted((k, str(v)) for k, v in row.items()))
                if key in seen:
                    duplicate_rows += 1
                else:
                    seen.add(key)
        except TypeError:
            duplicate_rows = None
        return {"rows": len(original), "columns": n_columns,
                "memory_mb": None, "duplicate_rows": duplicate_rows, "dtype_counts": None}
    if isinstance(original, (list, tuple)) and original and isinstance(original[0], (list, tuple)):
        seen = set()
        duplicate_rows = 0
        try:
            for row in original:
                key = tuple(row)
                if key in seen:
                    duplicate_rows += 1
                else:
                    seen.add(key)
        except TypeError:
            duplicate_rows = None
        return {"rows": len(original), "columns": n_columns,
                "memory_mb": None, "duplicate_rows": duplicate_rows, "dtype_counts": None}
    return None  # single-column input


def _print_info_overview(overview: Optional[dict]) -> None:
    if overview is None:
        return
    parts = [f"{overview['rows']:,} rows x {overview['columns']} columns"]
    if overview.get("memory_mb") is not None:
        parts.append(f"{overview['memory_mb']:.2f} MB")
    print("mx.INFO — " + "  |  ".join(parts))
    if overview.get("duplicate_rows"):
        print(f"  \u26a0 {overview['duplicate_rows']:,} duplicate rows")
    if overview.get("dtype_counts"):
        dtype_str = ", ".join(f"{k}({v})" for k, v in sorted(overview["dtype_counts"].items()))
        print(f"  dtypes: {dtype_str}")
    print()


def _wrap_info_result(original: Any, rows: list, overview: Optional[dict]):
    mod = type(original).__module__
    if mod.startswith("pandas"):
        import pandas as pd
        result = pd.DataFrame(rows).set_index("column")
        if overview is not None:
            result.attrs["overview"] = overview
        return result
    if mod.startswith("polars"):
        import polars as pl
        return pl.DataFrame(rows)
    return rows  # Table/list/tuple/numpy/Series input -> plain list of dicts


def INFO(
    table: Any,
    top_n: int = 10,
    categorical_max_unique: int = 20,
    categorical_max_ratio: float = 0.05,
    column_names: Optional[list] = None,
    print_overview: bool = True,
):
    """INFO(table) — one function, full column-by-column data profile.

    Works on a whole table (pandas/polars DataFrame, magpyxl Table,
    list-of-dicts, list-of-lists) or a single column (list/tuple/numpy
    array/pandas or polars Series) — same call either way.

    Prints a table-level header (row/column counts, memory, dtype
    breakdown, exact duplicate-row count) and returns a per-column
    profile: type (numeric/text/boolean/date/empty), count, missing
    (+ %), unique count, and type-appropriate statistics — plus a
    `flags` column surfacing data-quality issues proactively:
    constant columns, >=50% missing, mixed-type columns, many zeros,
    negative values, IQR outliers, text that looks numeric or
    date-shaped, ID-like high-cardinality columns, and near-duplicate
    categories from case/whitespace variation.

    Output type follows the input: pandas DataFrame in -> pandas
    DataFrame out (indexed by column name, overview attached via
    `.attrs["overview"]`), polars DataFrame in -> polars DataFrame
    out, anything else -> a list of dicts (one per input column).
    """
    columns = _load_named_columns(table, column_names)
    overview = _table_overview(table, len(columns))
    if print_overview:
        _print_info_overview(overview)
    rows = []
    for name, col in columns.items():
        if _is_datetime_column(col):
            rows.append(_summarize_date_column(name, _to_list(col)))
        else:
            s = _summarize_column(col, top_n, categorical_max_unique, categorical_max_ratio)
            rows.append(_format_info_row(name, s))
    return _wrap_info_result(table, rows, overview)


# ---------------------------------------------------------------------------
# CLEAN — companion to INFO(). Turns INFO's diagnostics into concrete,
# inspectable, EXPLICIT cleaning actions. Three modes, chosen by the
# caller, never guessed:
#
#   mx.CLEAN(table)              -> PLAN only. Data is never touched.
#   mx.CLEAN(table, plan=plan)   -> executes EXACTLY the given plan.
#   mx.CLEAN(table, auto=True)   -> executes only the safe, auto-eligible
#                                   subset (fill missing, merge near-dup
#                                   categories, drop exact duplicate
#                                   rows). NEVER drops a column. NEVER
#                                   guesses a dtype conversion.
# ---------------------------------------------------------------------------

def _build_cleaning_plan(
    table: Any,
    column_names: Optional[list] = None,
    top_n: int = 10,
    categorical_max_unique: int = 20,
    categorical_max_ratio: float = 0.05,
) -> list:
    """Diagnose `table` (reusing INFO's own classification logic) and
    turn the findings into a concrete, inspectable list of proposed
    actions. Purely read-only — never touches `table`'s data."""
    columns = _load_named_columns(table, column_names)
    overview = _table_overview(table, len(columns))
    plan: list = []

    if overview is not None and overview.get("duplicate_rows"):
        plan.append({
            "column": "__table__", "issue": "duplicate_rows",
            "action": "drop_duplicate_rows",
            "detail": f"{overview['duplicate_rows']:,} exact duplicate row(s) — keep first occurrence.",
            "risk": "low", "auto_eligible": True,
        })

    for name, col in columns.items():
        col_list = _to_list(col)
        total = len(col_list)

        if _is_datetime_column(col):
            n_missing = sum(1 for v in col_list if v is None)
            if n_missing:
                plan.append({
                    "column": name, "issue": "missing_values", "action": "review_only",
                    "detail": f"{n_missing:,} missing date value(s). Dates are never "
                              "auto-filled (a guessed date can be actively misleading) — "
                              "fill manually if you have a real value.",
                    "risk": "manual_review_required", "auto_eligible": False,
                })
            continue

        s = _summarize_column(col, top_n, categorical_max_unique, categorical_max_ratio)
        non_missing = total - s["missing"]

        if s["kind"] == "empty":
            if s["missing"]:
                plan.append({
                    "column": name, "issue": "missing_values", "action": "review_only",
                    "detail": f"All {s['missing']:,} value(s) are missing — nothing to compute a "
                              "fill value from. Needs a real data source, or consider dropping "
                              "this column (your decision, never automatic).",
                    "risk": "manual_review_required", "auto_eligible": False,
                })
            continue

        if s["missing"] > 0:
            if s["kind"] == "numeric":
                plan.append({
                    "column": name, "issue": "missing_values", "action": "fill_missing_median",
                    "detail": f"{s['missing']:,} missing ({100*s['missing']/total:.1f}%) -> fill "
                              f"with median ({s['median']:.6g}).",
                    "risk": "low", "auto_eligible": True, "_fill_value": s["median"],
                })
            elif s["kind"] == "boolean":
                fill_val = bool(s["mean"] >= 0.5)
                plan.append({
                    "column": name, "issue": "missing_values", "action": "fill_missing_mode",
                    "detail": f"{s['missing']:,} missing -> fill with the more common value ({fill_val}).",
                    "risk": "low", "auto_eligible": True, "_fill_value": fill_val,
                })
            elif s["kind"] == "text":
                plan.append({
                    "column": name, "issue": "missing_values", "action": "fill_missing_placeholder",
                    "detail": f"{s['missing']:,} missing ({100*s['missing']/total:.1f}%) -> fill "
                              "with the placeholder 'Missing' (not a guessed value, so it never "
                              "invents false data or inflates a real category's count).",
                    "risk": "low", "auto_eligible": True, "_fill_value": "Missing",
                })

        if s["kind"] == "text" and s["is_categorical"] and s["normalized_unique"] < s["unique"]:
            mapping = None
            try:
                if _is_string_array_backed(col):
                    mapping = _build_category_merge_mapping(_mixed_arg(col))
                elif isinstance(col, list) and col and all(isinstance(v, str) for v in col):
                    mapping = _build_category_merge_mapping(col)
            except Exception:
                mapping = None  # fall through to the pure-Python path below
            if mapping is None:
                mapping = {}
                groups: "dict[str, list]" = {}
                for v in col_list:
                    if v is None or (isinstance(v, str) and v.strip() == ""):
                        continue
                    key = str(v).strip().lower()
                    groups.setdefault(key, []).append(str(v))
                for variants in groups.values():
                    distinct = list(dict.fromkeys(variants))  # de-dup, preserve first-seen order
                    if len(distinct) > 1:
                        vcounts = {v: variants.count(v) for v in distinct}
                        # Tie-break toward the FIRST-SEEN casing, not the
                        # lexicographically largest string. A plain
                        # `max(..., key=lambda v: (count, v))` breaks count
                        # ties by string comparison — and since lowercase
                        # ASCII letters sort after uppercase ones ('it' >
                        # 'IT'), an exact 1-vs-1 tie between "IT" and "it"
                        # would silently pick the lowercase variant as
                        # canonical. Found by testing this exact tie case
                        # (not present in upstream's own test suite, which
                        # only covers a clear-majority-count scenario) —
                        # confirmed as a real, reproducible bug before
                        # fixing it, not assumed. `distinct.index(v)` gives
                        # first-occurrence order for the tie-break, which
                        # is deterministic and matches this whole function's
                        # existing "first occurrence wins" convention
                        # (mirrors drop_duplicate_rows' own `keep="first"`).
                        canonical = max(distinct, key=lambda v: (vcounts[v], -distinct.index(v)))
                        for v in distinct:
                            if v != canonical:
                                mapping[v] = canonical
            if mapping:
                # Build a few human-readable examples FROM the mapping
                # itself (group raw spellings by the canonical value
                # they map to) rather than re-scanning col_list a
                # second time — the mapping already has everything
                # needed for this, and col_list can be a million rows.
                by_canonical: "dict[str, list]" = {}
                for raw, canonical in mapping.items():
                    by_canonical.setdefault(canonical, []).append(raw)
                examples = [
                    "/".join(sorted(raws + [canonical])) + f" -> {canonical}"
                    for canonical, raws in list(by_canonical.items())[:3]
                ]
                plan.append({
                    "column": name, "issue": "near_duplicate_categories", "action": "merge_categories",
                    "detail": f"{len(mapping)} raw value(s) merged into their most common variant, "
                              f"e.g. {'; '.join(examples)}"
                              + (f" (+{len(by_canonical)-3} more)" if len(by_canonical) > 3 else ""),
                    "risk": "low", "auto_eligible": True, "_mapping": mapping,
                })

        if s["mixed_other_count"] > 0:
            plan.append({
                "column": name, "issue": "mixed_types", "action": "review_only",
                "detail": f"{s['mixed_other_count']:,} value(s) of a different type were excluded "
                          f"from this column's '{s['kind']}' statistics (majority type kept, "
                          "minority reported here rather than silently dropped). No automatic "
                          "fix — which type is correct needs a human decision.",
                "risk": "manual_review_required", "auto_eligible": False,
            })

        if s["is_constant"] and non_missing > 0:
            plan.append({
                "column": name, "issue": "constant", "action": "review_only",
                "detail": "Every non-missing value is identical. Might be genuinely constant "
                          "metadata (keep) or a data-entry/export error (drop) — your call. "
                          "Never auto-dropped, regardless of mode.",
                "risk": "manual_review_required", "auto_eligible": False,
            })

        if s["kind"] == "numeric" and s["outlier_count"] > 0:
            iqr = s["q3"] - s["q1"]
            plan.append({
                "column": name, "issue": "outliers", "action": "review_only",
                "detail": f"{s['outlier_count']:,} value(s) outside [Q1-1.5*IQR, Q3+1.5*IQR] = "
                          f"[{s['q1'] - 1.5*iqr:.6g}, {s['q3'] + 1.5*iqr:.6g}]. Could be genuine "
                          "extreme values — no automatic fix applied.",
                "risk": "manual_review_required", "auto_eligible": False,
            })

        if s["kind"] == "text" and non_missing > 0 and s["numeric_looking"] / non_missing >= 0.9:
            sample = [str(v).strip() for v in col_list[:2000] if isinstance(v, str) and v.strip()]
            leading_zero_risk = any(
                len(v) > 1 and v[0] == "0" and v.replace(".", "", 1).replace("-", "", 1).isdigit()
                for v in sample
            )
            plan.append({
                "column": name, "issue": "looks_numeric", "action": "convert_to_numeric",
                "detail": "This text column looks numeric. NOT auto-applied — "
                          + ("some values have leading zeros (e.g. zip/ID codes) that would be "
                             "SILENTLY LOST by numeric conversion; " if leading_zero_risk else "")
                          + "add this action to a plan explicitly if you want it applied.",
                "risk": "high" if leading_zero_risk else "medium", "auto_eligible": False,
            })

        if s["kind"] == "text" and non_missing > 0 and s["date_looking"] / non_missing >= 0.9:
            plan.append({
                "column": name, "issue": "looks_date", "action": "convert_to_date",
                "detail": "This text column looks like dates. NOT auto-applied — date format "
                          "can be genuinely ambiguous (e.g. 03/04/2024 = 3-Apr or 4-Mar "
                          "depending on locale) and a silent misread is worse than no "
                          "conversion; add this action to a plan explicitly once you've "
                          "confirmed the format.",
                "risk": "high", "auto_eligible": False,
            })

        if s["kind"] == "text" and not s["is_categorical"] and non_missing > 0 and s["unique"] / non_missing >= 0.95:
            plan.append({
                "column": name, "issue": "high_cardinality_id_like", "action": "review_only",
                "detail": f"{s['unique']:,} unique values out of {non_missing:,} — looks like an "
                          "identifier column, not categorical data. No action needed.",
                "risk": "none", "auto_eligible": False,
            })

    return plan


def _clean_report_entry(action: dict, applied: bool, result: str) -> dict:
    """Strip internal (`_`-prefixed) execution fields before handing the
    entry back — the report should be readable, not leak implementation
    details like the exact fill value's internal key."""
    public = {k: v for k, v in action.items() if not k.startswith("_")}
    public["applied"] = applied
    public["result"] = result
    return public


def _apply_cleaning_plan(table: Any, plan: list, column_names: Optional[list] = None):
    """Execute exactly the actions in `plan` against a COPY of `table`.
    The original object is never modified. Returns (cleaned_table, report)."""
    applicable = [a for a in plan if a.get("action") != "review_only"]
    # Run duplicate-row removal LAST, regardless of where it appears in
    # the plan: filling missing values and merging near-duplicate
    # categories can turn two previously-distinct rows into exact
    # duplicates — deduping first would miss those.
    applicable.sort(key=lambda a: a.get("action") == "drop_duplicate_rows")
    report: list = []
    mod = type(table).__module__

    if mod.startswith("pandas") and hasattr(table, "columns") and hasattr(table, "iloc"):
        df = table.copy(deep=True)
        for a in applicable:
            col, action = a.get("column"), a["action"]
            try:
                if action == "drop_duplicate_rows":
                    before = len(df)
                    df = df.drop_duplicates(keep="first").reset_index(drop=True)
                    report.append(_clean_report_entry(a, True, f"{before - len(df):,} row(s) dropped"))
                elif action in ("fill_missing_median", "fill_missing_mode", "fill_missing_placeholder"):
                    n = int(df[col].isna().sum())
                    df[col] = df[col].fillna(a["_fill_value"])
                    report.append(_clean_report_entry(a, True, f"{n:,} value(s) filled"))
                elif action == "merge_categories":
                    df[col] = df[col].replace(a["_mapping"])
                    report.append(_clean_report_entry(a, True, f"{len(a['_mapping']):,} value(s) remapped"))
                elif action == "convert_to_numeric":
                    import pandas as pd
                    df[col] = pd.to_numeric(df[col], errors="coerce")
                    report.append(_clean_report_entry(a, True, "converted to numeric (unparsable -> NaN)"))
                elif action == "convert_to_date":
                    import pandas as pd
                    df[col] = pd.to_datetime(df[col], errors="coerce")
                    report.append(_clean_report_entry(a, True, "converted to date (unparsable -> NaT)"))
                else:
                    report.append(_clean_report_entry(a, False, f"unknown action '{action}'"))
            except Exception as e:
                report.append(_clean_report_entry(a, False, f"FAILED: {e}"))
        return df, report

    if mod.startswith("polars") and hasattr(table, "columns") and hasattr(table, "height"):
        import polars as pl
        df = table.clone()
        for a in applicable:
            col, action = a.get("column"), a["action"]
            try:
                if action == "drop_duplicate_rows":
                    before = df.height
                    df = df.unique(keep="first", maintain_order=True)
                    report.append(_clean_report_entry(a, True, f"{before - df.height:,} row(s) dropped"))
                elif action in ("fill_missing_median", "fill_missing_mode", "fill_missing_placeholder"):
                    n = int(df[col].null_count())
                    df = df.with_columns(pl.col(col).fill_null(a["_fill_value"]))
                    report.append(_clean_report_entry(a, True, f"{n:,} value(s) filled"))
                elif action == "merge_categories":
                    df = df.with_columns(pl.col(col).replace(a["_mapping"]))
                    report.append(_clean_report_entry(a, True, f"{len(a['_mapping']):,} value(s) remapped"))
                elif action == "convert_to_numeric":
                    df = df.with_columns(pl.col(col).cast(pl.Float64, strict=False))
                    report.append(_clean_report_entry(a, True, "converted to numeric (unparsable -> null)"))
                elif action == "convert_to_date":
                    df = df.with_columns(pl.col(col).str.to_date(strict=False))
                    report.append(_clean_report_entry(a, True, "converted to date (unparsable -> null)"))
                else:
                    report.append(_clean_report_entry(a, False, f"unknown action '{action}'"))
            except Exception as e:
                report.append(_clean_report_entry(a, False, f"FAILED: {e}"))
        return df, report

    if isinstance(table, list) and table and isinstance(table[0], dict):
        rows = [dict(r) for r in table]
        for a in applicable:
            col, action = a.get("column"), a["action"]
            try:
                if action == "drop_duplicate_rows":
                    seen, kept, dropped = set(), [], 0
                    for r in rows:
                        key = tuple(sorted((k, str(v)) for k, v in r.items()))
                        if key in seen:
                            dropped += 1
                        else:
                            seen.add(key)
                            kept.append(r)
                    rows = kept
                    report.append(_clean_report_entry(a, True, f"{dropped:,} row(s) dropped"))
                elif action in ("fill_missing_median", "fill_missing_mode", "fill_missing_placeholder"):
                    n = 0
                    for r in rows:
                        if r.get(col) is None:
                            r[col] = a["_fill_value"]
                            n += 1
                    report.append(_clean_report_entry(a, True, f"{n:,} value(s) filled"))
                elif action == "merge_categories":
                    mapping, n = a["_mapping"], 0
                    for r in rows:
                        v = r.get(col)
                        if v in mapping:
                            r[col] = mapping[v]
                            n += 1
                    report.append(_clean_report_entry(a, True, f"{n:,} value(s) remapped"))
                elif action == "convert_to_numeric":
                    n = 0
                    for r in rows:
                        try:
                            r[col] = float(r[col])
                            n += 1
                        except (TypeError, ValueError):
                            pass
                    report.append(_clean_report_entry(a, True, f"{n:,} value(s) converted"))
                elif action == "convert_to_date":
                    import datetime as _dt
                    n = 0
                    for r in rows:
                        try:
                            r[col] = _dt.date.fromisoformat(str(r[col]).strip())
                            n += 1
                        except (TypeError, ValueError):
                            pass
                    report.append(_clean_report_entry(a, True, f"{n:,} value(s) converted"))
                else:
                    report.append(_clean_report_entry(a, False, f"unknown action '{action}'"))
            except Exception as e:
                report.append(_clean_report_entry(a, False, f"FAILED: {e}"))
        return rows, report

    # Single column: list/tuple/numpy array/pandas or polars Series
    # that ISN'T a whole table — reuse _wrap_like for output-type parity.
    col_list = list(_to_list(table))
    for a in applicable:
        action = a["action"]
        try:
            if action in ("fill_missing_median", "fill_missing_mode", "fill_missing_placeholder"):
                n = 0
                for i, v in enumerate(col_list):
                    if v is None:
                        col_list[i] = a["_fill_value"]
                        n += 1
                report.append(_clean_report_entry(a, True, f"{n:,} value(s) filled"))
            elif action == "merge_categories":
                mapping, n = a["_mapping"], 0
                for i, v in enumerate(col_list):
                    if v in mapping:
                        col_list[i] = mapping[v]
                        n += 1
                report.append(_clean_report_entry(a, True, f"{n:,} value(s) remapped"))
            elif action == "convert_to_numeric":
                n = 0
                for i, v in enumerate(col_list):
                    try:
                        col_list[i] = float(v)
                        n += 1
                    except (TypeError, ValueError):
                        pass
                report.append(_clean_report_entry(a, True, f"{n:,} value(s) converted"))
            elif action == "convert_to_date":
                import datetime as _dt
                n = 0
                for i, v in enumerate(col_list):
                    try:
                        col_list[i] = _dt.date.fromisoformat(str(v).strip())
                        n += 1
                    except (TypeError, ValueError):
                        pass
                report.append(_clean_report_entry(a, True, f"{n:,} value(s) converted"))
            elif action == "drop_duplicate_rows":
                report.append(_clean_report_entry(a, False, "not applicable to a single column"))
            else:
                report.append(_clean_report_entry(a, False, f"unknown action '{action}'"))
        except Exception as e:
            report.append(_clean_report_entry(a, False, f"FAILED: {e}"))

    index = getattr(table, "index", None)
    return _wrap_like(table, col_list, index=index), report


def CLEAN(
    table: Any,
    plan: Optional[list] = None,
    auto: bool = False,
    top_n: int = 10,
    categorical_max_unique: int = 20,
    categorical_max_ratio: float = 0.05,
    column_names: Optional[list] = None,
):
    """CLEAN(table) — companion to INFO(). Turns INFO's diagnostics into
    concrete, EXPLICIT cleaning actions. Three modes, chosen by you,
    never guessed:

        plan = mx.CLEAN(df)
        # -> a list[dict] of proposed actions. Data is NEVER touched.
        #    Inspect it, delete/edit entries you don't want, then:

        cleaned_df, report = mx.CLEAN(df, plan=plan)
        # -> executes EXACTLY the actions remaining in your plan.

        cleaned_df, report = mx.CLEAN(df, auto=True)
        # -> executes only the provably-safe subset automatically:
        #      - fill missing values (numeric->median, boolean->majority,
        #        text->explicit "Missing" placeholder; dates are NEVER
        #        auto-filled)
        #      - merge case/whitespace near-duplicate categories,
        #        exact-after-normalization only
        #      - drop EXACT duplicate rows
        #    NEVER drops a column, in any mode. NEVER auto-converts a
        #    column's dtype (numeric/date).

    `table`/`plan`/`auto` mirror INFO()'s parameters; the diagnostic
    thresholds are shared with INFO() so a plan always matches what
    INFO would show for the same data.

    Returns:
      plan mode      -> list[dict]
      auto/plan mode -> (cleaned_table, report) — cleaned_table's
                        container type matches the input; report lists
                        exactly what was applied and its measured
                        effect. The original `table` is never modified.
    """
    generated_plan = _build_cleaning_plan(
        table, column_names, top_n, categorical_max_unique, categorical_max_ratio
    )
    if plan is None and not auto:
        return generated_plan
    if auto:
        chosen = [a for a in generated_plan if a.get("auto_eligible")]
        if not any(a.get("action") == "drop_duplicate_rows" for a in chosen):
            # Fill/merge can CREATE duplicate rows that didn't exist in
            # the raw data — auto mode always re-checks for duplicates
            # after normalizing, not just before.
            chosen = chosen + [{
                "column": "__table__", "issue": "duplicate_rows", "action": "drop_duplicate_rows",
                "detail": "Post-cleaning duplicate check (values may have become identical after "
                          "fill/merge even if the raw rows weren't exact duplicates).",
                "risk": "low", "auto_eligible": True,
            }]
    else:
        chosen = plan
    return _apply_cleaning_plan(table, chosen, column_names)


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


# ---------------------------------------------------------------------------
# Case-insensitive formula names: mx.SUM, mx.sum, mx.Sum, mx.sUM all
# resolve to the same function -- Excel itself doesn't care about a
# formula name's case, so neither does magpyxl. Implemented via PEP 562
# module-level __getattr__, which Python only calls when a normal
# attribute lookup already failed -- so exact-case names (mx.SUM) still
# resolve through the fast, ordinary module-attribute path with zero
# added overhead; only a differently-cased name falls through to the
# case-insensitive lookup below. Built directly from `__all__`, so a
# newly added function is automatically covered here too -- nothing to
# keep in sync by hand. `Table`/`read_table` are in `__all__` for
# normal exact-case access but deliberately excluded here: they're a
# class and a loader function, not Excel-style formula names, and
# giving them case-insensitive aliases (`mx.table`, `mx.READ_TABLE`)
# would invite confusion without matching any real Excel convention.
# ---------------------------------------------------------------------------

_CASE_INSENSITIVE_EXCLUDED = {"Table", "read_table"}
_FORMULA_REGISTRY = {
    name: globals()[name]
    for name in __all__
    if name in globals() and name not in _CASE_INSENSITIVE_EXCLUDED
}


def __getattr__(name: str):
    upper = name.upper()
    if upper in _FORMULA_REGISTRY:
        return _FORMULA_REGISTRY[upper]
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
