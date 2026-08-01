# MagpieXL — Project Handoff / State Document

**Read this first.** This file exists so that anyone (human or AI) picking up
this project — with only this file and the `magpiexl/` source folder — can
understand the full history, current state, design decisions, and exactly
what to do next, without re-deriving any of it from scratch.

---

## 1. What MagpieXL Is

Excel-style formulas — `SUM`, `AVERAGE`, `COUNT`, `COUNTIF`, `SUMIF`,
`AVERAGEIF`, `COUNTIFS`, `SUMIFS`, `VLOOKUP`, `XLOOKUP` — implemented with a
compiled **Rust core** (via PyO3) and a thin **Python adapter layer**, callable
with identical syntax whether the data is a plain Python list/tuple, a numpy
array, a pandas DataFrame/Series, a polars DataFrame/Series, or a `.csv`/
`.xlsx` file on disk.

The core idea the user cares about most: **same type in → same type out.**
Pass a pandas Series, get a pandas Series back (index preserved). Pass a
plain list, get a plain list back. This lets `magpiexl` calls sit inline in
existing pandas/polars method chains without ever needing an explicit
Python-level loop written by the user.

## 2. Current Version / Status

**v0.1.0** — 11 functions implemented and manually verified (`SUM`,
`AVERAGE`, `COUNT`, `COUNTIF`, `SUMIF`, `AVERAGEIF`, `COUNTIFS`, `SUMIFS`,
`VLOOKUP`, `XLOOKUP`, `LOOKUPIFS`). `XLOOKUP` additionally supports a
multi-column `return_array` (see §4.5). No committed
automated test suite yet (see §7 "What's explicitly NOT done"). Built and
tested only for **Linux x86_64, CPython 3.12** in this dev sandbox. A
Windows build guide exists (`WINDOWS-BUILD.md`) but was never actually
built/verified on Windows (cross-compiling from this Linux sandbox to
Windows is not possible here — see §6).

## 3. File Layout

```
magpiexl/
├── Cargo.toml              # Rust crate config (pyo3 dependency only, no numpy)
├── pyproject.toml           # maturin build config, mixed python/rust layout
├── src/lib.rs                # ~1020 lines — Rust core, see §4
├── python/magpiexl/
│   └── __init__.py           # ~545 lines — Python adapter layer, see §4
├── README.md                 # user-facing usage guide
└── WINDOWS-BUILD.md          # manual Windows build steps (rustup + maturin)
```

Build command (from inside `magpiexl/`):
```bash
maturin build --release
pip install target/wheels/magpiexl-0.1.0-cp312-cp312-*.whl --break-system-packages
```
(`--break-system-packages` only needed on Debian/Ubuntu-managed Python envs.)

There is currently no `pyproject.toml` runtime dependency on numpy/pandas/
polars — they're all optional, duck-typed at the Python layer (see §4.2).

## 4. Architecture

### 4.1 Rust core (`src/lib.rs`)

**Value model** — `CellValue` enum: `Num(f64)`, `Text(String)`, `Empty`.
`CellValue::from_py()` converts a `Bound<PyAny>` into this using
`.downcast::<PyBool/PyFloat/PyInt/PyString>()` — **not** `.extract::<T>()`
chains. This matters: `.extract()` on a type mismatch raises and discards a
real Python exception internally (expensive, and arguably not "clean" control
flow); `.downcast()` is a cheap type-tag check with no exception machinery.
This was the one performance-adjacent change kept from an earlier, larger
optimization pass that was otherwise reverted (see §5).

**Criteria model** — `Op` enum (`Eq/Ne/Gt/Ge/Lt/Le`), `Criteria { op, value:
CellValue, wildcard: Option<String> }`. `parse_criteria()` parses Excel-style
strings: `">10"`, `"<=5"`, `"<>0"`, `"ab*"`, `"a?c"`, or a bare value (implicit
equality). `matches()` applies a `Criteria` to a `CellValue`. `wildcard_match()`
is a small recursive glob matcher (`*` = any run, `?` = exactly one char),
case-insensitive throughout (matches Excel).

**Lookup key model** — `LookupKey` enum (`Num(f64)` hashed via `.to_bits()`,
`Text(String)` lowercased), used to build `HashMap`s for O(1) lookups.
`cell_to_key()` converts a `CellValue` → `Option<LookupKey>`. `criteria_key()`
returns `Some(key)` **only** if a `Criteria` is a plain equality (no wildcard,
`op == Eq`) — this is the gate used to decide whether a batch of vectorized
criteria can use the HashMap fast path (see §4.3).

**Every public function has two Rust variants where relevant:**
- A **scalar** variant (`*_values`): one range/criteria in, one number/value out.
- A **vectorized** variant (`*_vec_values`): one range, a *list* of criteria in,
  a matching list of results out — computed in ONE Rust call, not by Python
  calling the scalar function N times.

Current full function inventory in `_core` (all inside `#[pymodule] mod _core`):

| Rust function | Signature | Notes |
|---|---|---|
| `sum_values` | `(values: Vec<PyObject>) -> f64` | |
| `average_values` | `(values) -> f64` | raises if no numeric values |
| `count_values` | `(values) -> i64` | Excel COUNT semantics (numeric cells only) |
| `countif_values` | `(range, criteria) -> i64` | scalar criteria |
| `sumif_values` | `(range, criteria, sum_range=None) -> f64` | |
| `averageif_values` | `(range, criteria, average_range=None) -> f64` | raises if no match |
| `countifs_values` | `(pairs: Vec<(Vec<PyObject>, PyObject)>) -> i64` | AND across pairs |
| `sumifs_values` | `(sum_range, pairs) -> f64` | |
| `countif_vec_values` | `(range, criteria_list) -> Vec<i64>` | HashMap fast path for all-equality batches |
| `sumif_vec_values` | `(range, criteria_list, sum_range=None) -> Vec<f64>` | same fast path |
| `averageif_vec_values` | `(range, criteria_list, average_range=None) -> Vec<Option<f64>>` | `None` per-row on no match (not a raised error — see §4.3) |
| `countifs_vec_values` | `(pairs: Vec<(Vec<PyObject>, Vec<PyObject>)>) -> Vec<i64>` | criteria lists must all be same length N; fast path uses a combined `Vec<LookupKey>` as the HashMap key |
| `sumifs_vec_values` | `(sum_range, pairs) -> Vec<f64>` | same |
| `vlookup_values` | `(lookup_value, table: Vec<Vec<PyObject>>, col_index: usize, range_lookup=false) -> PyObject` | raises on miss, no `if_not_found` param here — handled in Python (see §4.4) |
| `xlookup_values` | `(lookup_value, lookup_array, return_array, if_not_found: Option<PyObject>=None) -> PyObject` | `if_not_found` IS a Rust param here |
| `vlookup_many_values` | `(lookup_values, table, col_index, range_lookup=false, if_not_found=None) -> Vec<PyObject>` | HashMap-built-once, O(1) per lookup (O(n+m) total, not O(n·m)) |
| `xlookup_many_values` | `(lookup_values, lookup_array, return_array, if_not_found=None) -> Vec<PyObject>` | same HashMap approach |
| `xlookup_many_indices` | `(lookup_values, lookup_array) -> Vec<Option<i64>>` | same HashMap approach, but returns the matching ROW INDEX instead of the resolved value — lets Python pick MULTIPLE return columns for the same match without duplicating the matching logic per column. Powers `XLOOKUP`'s multi-column `return_array` support. |
| `lookupifs_indices_values` | `(pairs: Vec<(Vec<PyObject>, Vec<PyObject>)>) -> Vec<Vec<i64>>` | Same shape/fast-path as `countifs_vec_values`, but returns the **list** of matching row-indices per output row (not a count) — powers `LOOKUPIFS`. |

### 4.2 Python adapter layer (`python/magpiexl/__init__.py`)

**Universal input normalization:**
- `_to_list(data)` — list/tuple/numpy array/pandas Series/polars Series →
  plain Python list. Already-a-list short-circuits (no copy). Raises
  `TypeError` if given a whole `Table` (must index a column first).
- `Table` class — minimal dependency-free columnar table
  (`dict[str, list]`), with `.from_csv()`/`.from_xlsx()`, `.rows()`,
  `.column_names`, `__getitem__`, `__len__`. Lets magpiexl work with zero
  pandas/polars/numpy installed.
- `read_table(path, sheet=None)` — dispatches `.csv`/`.xlsx` → `Table`.
- `_load_rows_and_columns(data)` — normalizes ANY "table" argument (pandas
  DataFrame, polars DataFrame, `Table`, list-of-dicts, list-of-lists, or a
  csv/xlsx path string) → `(rows: list[list], column_names: list[str] | None)`.
  Used by `VLOOKUP`.
- `_resolve_col_index(col_index, column_names)` — int passthrough, or
  resolves a column-name string to a 1-based index.

**Type-preserving vectorization (the "same type in, same type out" mechanism):**
- `_is_array_like(x)` — True for list/tuple always; for pandas/polars/numpy
  objects only if they have `__len__` (excludes scalars); explicitly False
  for `str`/`bytes` (so a string lookup value is never mistaken for "a column
  of characters").
- `_wrap_like(origin, values, index=None)` — wraps a plain list of results
  back into the same "ecosystem" as `origin`: pandas → `pd.Series` (with
  `index` preserved), polars → `pl.Series`, numpy → `np.array(dtype=object)`,
  tuple → `tuple`, else → plain `list`.

**The `if_not_found` sentinel pattern** (important, reusable pattern for any
future optional-default-that-could-legitimately-be-None parameter):
```python
_UNSET = object()          # "argument not given at all" -> should raise on miss
_NONE_FALLBACK = object()  # internal stand-in for "user explicitly passed None"

def _rust_fallback(if_not_found):
    if if_not_found is _UNSET: return None            # tell Rust: raise on miss
    if if_not_found is None:   return _NONE_FALLBACK   # user wants None back — can't send real None, ambiguous
    return if_not_found

def _unwrap_fallback(value):
    return None if value is _NONE_FALLBACK else value
```
**Why this exists:** Python's `None`, once it crosses into Rust's
`Option<PyObject>` parameter via PyO3's default extraction, is indistinguishable
whether the caller passed `if_not_found=None` explicitly or didn't pass the
argument at all — both collapse to Rust's `Option::None`. Without the
sentinel, `XLOOKUP(x, ..., if_not_found=None)` would incorrectly raise instead
of returning `None`. Both `VLOOKUP` and `XLOOKUP` (and the `.mx` accessor
wrappers) default to `if_not_found=_UNSET`, not `None`.

**Public API functions**, each following the same shape — normalize input →
scalar-or-array auto-detect on the criteria/lookup_value argument → call the
matching Rust function → (for arrays) `_wrap_like()`:

- `SUM`, `AVERAGE`, `COUNT` — no criteria argument, always scalar-in/scalar-out.
- `COUNTIF`, `SUMIF`, `AVERAGEIF` — `criteria` can be scalar or a whole
  column; if array-like, calls the `*_vec_values` Rust function once (NOT a
  Python loop). `AVERAGEIF`'s vector path converts Rust's `None` (no match
  for that row) into `float("nan")` rather than raising, so one "bad" row
  doesn't kill the whole batch (unlike the scalar path, which does raise).
- `_prepare_ifs_vector(ranges, criteria)` — shared helper for `COUNTIFS`/
  `SUMIFS`: detects if ANY criteria arg is array-like, broadcasts scalar
  criteria to the batch length N (`[c] * n`, cheap since N is the small
  "output" size, not the range size), and builds
  `pairs = [(range_list, criteria_list_of_length_N), ...]` ready for the
  vec Rust function. Returns `None` if every criteria is a plain scalar
  (falls back to the old scalar path unchanged).
- `COUNTIFS(*args)`, `SUMIFS(sum_range, *args)` — use `_prepare_ifs_vector`;
  vectorized path = ONE call to `_countifs_vec_values`/`_sumifs_vec_values`.
- `VLOOKUP(lookup_value, table, col_index, range_lookup=False, if_not_found=_UNSET)`
  — `table` accepts pandas/polars DataFrame, `Table`, list-of-dicts,
  list-of-lists, or a csv/xlsx path string (auto-loaded via `read_table`);
  `col_index` accepts a 1-based int OR a column-name string; `lookup_value`
  scalar-or-array auto-detected (array path uses `vlookup_many_values`, O(1)
  per lookup after one O(n) HashMap build — NOT a linear scan per lookup).
- `XLOOKUP(lookup_value, lookup_array, return_array, if_not_found=_UNSET)` —
  same scalar/array pattern; sentinel handling applies to BOTH the scalar
  and array paths here (unlike VLOOKUP's scalar path, which doesn't route
  through Rust's `if_not_found` at all — it just catches the raised
  `ValueError` in Python instead; both approaches are correct, just
  implemented slightly differently for historical reasons).

**`.mx` accessor** — `_make_accessor(object)` builds a class with `SUM`,
`AVERAGE`, `COUNT`, `COUNTIF`, `SUMIF`, `AVERAGEIF`, `COUNTIFS`, `SUMIFS`,
`VLOOKUP`, `XLOOKUP` methods, where string arguments matching a column name
are auto-resolved via `self._col(name)` (else passed through as a literal).
Registered on pandas via `pd.api.extensions.register_dataframe_accessor("mx")`
and on polars via `pl.api.register_dataframe_namespace("mx")`, both wrapped
in `try/except ImportError` (registration is silently skipped if that
library isn't installed — magpiexl has no hard dependency on either).

### 4.3 Why the HashMap fast path exists for vectorized `*IF`/`*IFS`

The motivating use case (from the user): given `table1_ids` (few unique IDs)
and `table2_ids` (many rows), find how many times each `table1` ID appears in
`table2`. Naively, `countif_vec_values` could just loop "for each criteria,
scan the whole range" — correct, but rescans the (potentially large) range
once per criteria value, i.e. O(n·m).

Since **plain equality** is the overwhelmingly common case for this pattern
(as opposed to `">100"`-style comparisons), `criteria_key()` checks: if
**every** criteria in the batch is a plain equality, build ONE frequency map
(`HashMap<LookupKey, count>` or `HashMap<LookupKey, sum>`) from the range in
a single O(n) pass, then answer every criteria with an O(1) map lookup — total
O(n + m) instead of O(n·m). The moment **any** criteria in the batch needs a
comparison or wildcard, the whole batch falls back to the per-criteria scan
(still 100% correct, just without the shortcut) — this fallback is required
for correctness, not optional.

`COUNTIFS`/`SUMIFS` extend this to multi-column AND: the fast-path key is a
`Vec<LookupKey>` (one key per ANDed column) built once per input row; this
works because `Vec<T: Hash + Eq>` is automatically `Hash + Eq` in Rust.

### 4.4 Per-Function Algorithm Reference

This is the detailed algorithmic breakdown for every function currently
implemented — `n` = length of the range/table being scanned, `m` = number of
criteria values / lookup values in a vectorized call (i.e. the output
length), `k` = number of ANDed range/criteria pairs in `*IFS` functions
(small, effectively constant).

---

**Function: SUM**
- Algorithm: Single linear pass, accumulate
- Pseudo:
  1. total = 0
  2. For each value: if numeric, total += value
  3. Return total
- Time: O(n)
- Space: O(1) extra (excluding input storage)
- Data Structure: `Vec<PyObject>` in, scalar accumulator
- Optimizations: `.downcast()`-based type check (no exception-based control
  flow per element — see §4.1)
- Input/Output: list/pandas Series/polars Series/numpy array → `Vec<PyObject>`
  → `f64` (always scalar; no criteria to vectorize)
- Why this algorithm: no criteria/filter involved, so a single pass with a
  running total is already optimal — nothing to cache or precompute.

---

**Function: AVERAGE**
- Algorithm: Single linear pass, dual accumulate (sum + count)
- Pseudo:
  1. total = 0, count = 0
  2. For each value: if numeric, total += value, count += 1
  3. Return total / count (raise if count == 0)
- Time: O(n)
- Space: O(1) extra
- Data Structure: `Vec<PyObject>` in, two scalar accumulators
- Optimizations: computes sum AND count in the SAME pass (not two separate
  passes over the data)
- Input/Output: same as SUM → `f64`
- Why this algorithm: a second pass to count after summing would double the
  work for no benefit; accumulating both together is strictly better.

---

**Function: COUNT**
- Algorithm: Single linear pass, filter + count
- Time: O(n)
- Space: O(1) extra
- Data Structure: `Vec<PyObject>` in, scalar counter
- Optimizations: `.downcast()`-based numeric check only (Excel COUNT
  semantics: numeric cells only, text/blank ignored)
- Input/Output: same as SUM → `i64`
- Why this algorithm: simplest correct approach for "how many of these are
  numbers" — no faster approach exists for an unsorted, unindexed input.

---

**Function: COUNTIF (scalar)** — `countif_values`
- Algorithm: Parse-once, then single linear pass
- Pseudo:
  1. Parse criteria into `Criteria { op, value, wildcard }` ONCE
  2. For each cell in range: if it matches the criteria, count += 1
  3. Return count
- Time: O(n)
- Space: O(1) extra
- Data Structure: `Vec<PyObject>`, one parsed `Criteria` struct
- Optimizations: criteria parsed once before the loop, not re-parsed per
  element
- Input/Output: range (any supported type) + scalar criteria → `i64`
- Why this algorithm: for a single scalar criteria, every cell must be
  inspected at least once regardless of approach — O(n) is optimal here;
  there is nothing to cache across "queries" because there is only one query.

---

**Function: COUNTIF (vectorized) — `countif_vec_values`**
- Algorithm: HashMap + Single Pass (fast path) / Per-criteria scan (fallback)
- Pseudo:
  1. Parse all m criteria once
  2. If every criteria is a plain equality (no wildcard, no `>`/`<`/`<>`):
     a. Build `HashMap<key, count>` from the range in ONE pass — O(n)
     b. For each of the m criteria, look up its key — O(1) each
  3. Else (any criteria needs a comparison or wildcard):
     For each criteria, scan the whole range once (per-criteria scan)
  4. Return list of m counts
- Time: **O(n + m)** fast path (all-equality batch) · O(n·m) fallback
  (any comparison/wildcard criteria present)
- Space: O(n) for the HashMap (fast path) · O(1) extra (fallback)
- Data Structure: `HashMap<LookupKey, i64>`, `Vec<Criteria>`
- Optimizations: HashMap built ONCE and reused for every criteria in the
  batch — the range is scanned exactly once total, not once per criteria
- Input/Output: pandas Series/list/tuple (range) + pandas Series/list/tuple
  (criteria column) → `Vec<PyObject>` × 2 → `Vec<i64>` → same type as the
  criteria argument (pandas → pandas, list → list, tuple → tuple)
- Why this algorithm: the motivating use case (count how many times each ID
  in table1 appears in table2) is pure equality lookups — building one
  frequency map turns an O(n·m) repeated-scan into O(n+m). Falls back
  correctly (not incorrectly-fast) the moment a criteria genuinely needs a
  comparison/wildcard, since those can't be answered by an equality hashmap.

---

**Function: SUMIF (scalar) — `sumif_values`**
- Algorithm: Parse-once, single linear pass, conditional accumulate
- Time: O(n) · Space: O(1) extra
- Data Structure: `Vec<PyObject>` × 2 (range, sum_range), one `Criteria`
- Optimizations: criteria parsed once; `sum_range` defaults to `range` itself
  without an extra copy when not provided
- Input/Output: range + scalar criteria + optional sum_range → `f64`
- Why this algorithm: same reasoning as scalar COUNTIF — single query,
  O(n) is optimal.

---

**Function: SUMIF (vectorized) — `sumif_vec_values`**
- Algorithm: HashMap + Single Pass (fast path) / Per-criteria scan (fallback)
- Pseudo: identical shape to vectorized COUNTIF, but the HashMap accumulates
  a running **sum** per key instead of a count
- Time: O(n+m) fast path · O(n·m) fallback
- Space: O(n) for the HashMap
- Data Structure: `HashMap<LookupKey, f64>`
- Optimizations: same as vectorized COUNTIF — one pass builds the map,
  reused for every criteria in the batch
- Input/Output: same shape as vectorized COUNTIF → `Vec<f64>` → type-preserved
- Why this algorithm: same reasoning as vectorized COUNTIF — the map is
  built once and answers every criteria in O(1), regardless of batch size m.

---

**Function: AVERAGEIF (scalar) — `averageif_values`**
- Algorithm: Parse-once, single pass, dual accumulate (sum + count) per criteria
- Time: O(n) · Space: O(1) extra
- Optimizations: sum and count accumulated together in the same pass
  (same idea as AVERAGE)
- Input/Output: range + scalar criteria + optional average_range → `f64`
  (raises if zero matches)
- Why this algorithm: same as scalar SUMIF/COUNTIF.

---

**Function: AVERAGEIF (vectorized) — `averageif_vec_values`**
- Algorithm: HashMap of `(sum, count)` pairs (fast path) / per-criteria scan
  (fallback)
- Time: O(n+m) fast path · O(n·m) fallback
- Space: O(n)
- Data Structure: `HashMap<LookupKey, (f64, u64)>`
- Optimizations: same one-pass-map idea as SUMIF/COUNTIF vectorized
- Special behavior: a criteria with zero matching numeric values returns
  `None` for **that row only** (converted to `float("nan")` in the Python
  layer) rather than raising — one bad row does not fail the whole batch,
  unlike the scalar version which does raise
- Input/Output: same shape as vectorized SUMIF → `Vec<Option<f64>>` →
  type-preserved (NaN-filled where unmatched)
- Why this algorithm: same map-once reasoning as SUMIF, plus the per-row
  soft-failure behavior specifically because a vectorized batch is likely to
  contain at least one criteria with no matches (e.g. a department that
  exists in the summary table but not yet in the data) — raising would make
  the whole feature unusable in practice.

---

**Function: COUNTIFS (scalar) — `countifs_values`**
- Algorithm: Single pass over n rows, AND across k parsed criteria
- Pseudo:
  1. Parse all k criteria once
  2. For each of the n rows: check all k columns match (AND); if all match, count += 1
  3. Return count
- Time: O(n·k) (k is small/constant in practice) · Space: O(1) extra
- Why this algorithm: k criteria must each be checked per row regardless of
  approach; no precomputation helps for a single query.

---

**Function: COUNTIFS (vectorized) — `countifs_vec_values`**
- Algorithm: HashMap + Single Pass
- Pseudo:
  1. Build HashMap: for each of the n input rows, compute a **combined key**
     (one `LookupKey` per ANDed column, as a `Vec<LookupKey>`) and increment
     `freq[combined_key]`
  2. Iterate once (the build pass above IS the one iteration over n)
  3. For each of the m output rows, build its own combined key from the
     per-row criteria and look it up in the map — O(1) (average case)
  4. Return list of m counts
- Time: **O(n·k + m·k)** ≈ **O(n+m)** for constant k
- Space: O(n) — one HashMap entry per distinct combined key, worst case n
- Data Structure: `HashMap<Vec<LookupKey>, i64>`, `Vec<Criteria>` × k
- Optimizations: HashMap built ONCE across all k ANDed columns simultaneously
  (a combined multi-column key), reused for every one of the m output rows —
  avoids both "one hashmap per column" (wrong semantics for AND) and
  "rescan n rows once per output row" (the O(n·m) naive approach)
- Input/Output: pandas Series(s) → `Vec<PyObject>` per column → `Vec<i64>` →
  type-preserved (matches whichever criteria argument was array-like)
- Why this algorithm: multi-column AND with a shared frequency map is the
  direct generalization of the single-column COUNTIF fast path — falls back
  to the scan approach the moment any of the k criteria needs a
  comparison/wildcard rather than plain equality.

---

**Function: SUMIFS (vectorized) — `sumifs_vec_values`**
- Algorithm: HashMap + Single Pass
- Pseudo:
  1. Build HashMap: combined key (per ANDed column) → running sum, accumulated
     over the n input rows in one pass
  2. Iterate once
  3. Aggregate (sum accumulation happens during the build pass)
  4. For each of the m output rows, look up its combined key and return the
     accumulated sum (0.0 if the key was never seen)
- Time: O(n·k + m·k) ≈ **O(n+m)**
- Space: O(n)
- Data Structure: `HashMap<Vec<LookupKey>, f64>`
- Optimizations:
  - HashMap built once
  - Reused for every criteria row in the vectorized batch
  - No repeated O(n) scans per output row
- Input/Output: Pandas Series → `Vec<PyObject>` → `Vec<f64>` → Pandas Series
  (index preserved) — or list→list / tuple→tuple / polars→polars depending
  on which criteria argument was array-like
- Why this algorithm: avoids repeated O(n×m) scans — this was the exact
  case that motivated the whole vectorized-`*IFS` design (see §4.3): the
  original Python-loop implementation called the scalar Rust function once
  per output row, rescanning the full range every time; this version scans
  the range exactly once regardless of how many output rows are requested.

---

**Function: VLOOKUP (scalar, exact match) — `vlookup_values`**
- Algorithm: Linear scan, first match wins
- Time: O(n) · Space: O(1) extra
- Data Structure: `Vec<Vec<PyObject>>` (rows), no auxiliary structure
- Optimizations: none needed — see "why" below
- Input/Output: single lookup value + table (any supported type) → single
  `PyObject` (raises on miss)
- Why this algorithm: building a HashMap to answer exactly ONE query would
  cost O(n) to build for an O(1) query — no better than just scanning
  directly, and it's extra code for no benefit. The vectorized path (below)
  is where a HashMap actually pays off.
- Note: the `range_lookup=True` (approximate match) branch is also a linear
  scan with early-break, since it assumes the table is pre-sorted ascending
  and stops at the first row exceeding the lookup value — average O(n/2) in
  practice, worst case O(n). This could become O(log n) via binary search
  since the input is assumed sorted, but that optimization has not been
  implemented (see §7/§9 — deliberately deferred, performance phase).

---

**Function: VLOOKUP (vectorized) — `vlookup_many_values`**
- Algorithm: HashMap + Single Pass
- Pseudo:
  1. Build `HashMap<key, row_index>` from the table's first column — O(n)
     (first match wins per key, matching Excel's own VLOOKUP tie-break rule)
  2. Iterate once (the build pass above)
  3. For each of the m lookup values, look up its key — O(1) average
  4. Return the matching column value per lookup (or `if_not_found` fallback)
- Time: **O(n + m)** — versus O(n·m) for a naive "scan the table once per
  lookup value" approach
- Space: O(n) for the HashMap
- Data Structure: `HashMap<LookupKey, usize>`, `Vec<Vec<PyObject>>`
- Optimizations: HashMap built ONCE regardless of how many lookup values are
  requested — this is the single biggest algorithmic win in the whole
  library for the common "join-like" use case (looking up many keys against
  one reference table)
- Input/Output: pandas Series (lookup_value) + any supported table type →
  `Vec<PyObject>` → `Vec<PyObject>` → type-preserved (pandas Series with
  index preserved, polars Series, list, or tuple)
- Why this algorithm: the exact scenario this function exists for — "for
  every row in my main table, look up a value in a reference table" — is by
  definition many lookups against one fixed table, which is precisely what
  a HashMap turns from O(n·m) into O(n+m).

---

**Function: XLOOKUP (scalar) — `xlookup_values`**
- Algorithm: Linear scan, first match wins
- Time: O(n) · Space: O(1) extra
- Input/Output: single lookup value + lookup_array + return_array →
  single `PyObject` (or `if_not_found` fallback, via the `_UNSET` sentinel
  — see §4.2)
- Why this algorithm: same reasoning as scalar VLOOKUP — one query does not
  justify building a HashMap.

---

**Function: XLOOKUP (vectorized) — `xlookup_many_values`**
- Algorithm: HashMap + Single Pass (identical shape to `vlookup_many_values`)
- Time: O(n+m) · Space: O(n)
- Data Structure: `HashMap<LookupKey, usize>`
- Optimizations: same as vectorized VLOOKUP — one HashMap build, O(1) per
  lookup thereafter
- Input/Output: same shape as vectorized VLOOKUP → type-preserved
- Why this algorithm: same reasoning as vectorized VLOOKUP.

---

**Function: XLOOKUP (multi-column `return_array`) — `xlookup_many_indices`**
- Algorithm: HashMap + Single Pass — identical HashMap-build logic to
  `xlookup_many_values`, but returns the matching **row index**
  (`Option<i64>`) instead of directly resolving a value
- Time: O(n+m) · Space: O(n)
- Data Structure: `HashMap<LookupKey, usize>`
- Optimizations: the SAME index resolution is computed once and then reused
  in Python to extract values from however many return columns were
  requested — avoids rebuilding the HashMap (or rescanning) once per
  requested column
- Input/Output: lookup_value(s) + lookup_array → `Vec<Option<i64>>` →
  (Python) column extraction → pandas Series (scalar) or DataFrame (vectorized)
- Why this algorithm: matching logic (which row?) and extraction logic
  (which column value(s) from that row?) are orthogonal concerns — keeping
  them separate means adding "return 5 columns instead of 1" costs zero
  extra Rust work, just 5 cheap Python-side list-index operations instead
  of 1.

---

**Function: LOOKUPIFS (scalar and vectorized) — `lookupifs_indices_values`**
- Algorithm: HashMap + Single Pass (fast path) / Per-row scan (fallback) —
  same shape as `countifs_vec_values`, but each HashMap entry stores a
  **list** of matching row indices instead of a count
- Pseudo:
  1. Parse all criteria (k columns × m output rows) once
  2. If every criteria is a plain equality: build
     `HashMap<Vec<LookupKey>, Vec<usize>>` from the n input rows in one pass
     (combined key across all k ANDed columns → list of row indices that
     produced that key)
  3. Else: for each of the m output rows, scan all n input rows once,
     collecting indices where all k criteria match
  4. For each of the m output rows, look up (or use the collected list from
     step 3) → a `Vec<i64>` of matching row indices
  5. (Python layer) apply `mode` ("first"/"last"/"all") and extract values
     from however many return columns were requested, using those indices
- Time: O(n·k + m·k) ≈ **O(n+m)** fast path · O(n·m·k) fallback
- Space: O(n) — HashMap holds up to n index-list entries total across all
  keys (fast path)
- Data Structure: `HashMap<Vec<LookupKey>, Vec<usize>>`
- Optimizations: identical one-pass-builds-a-map idea as `countifs_vec_values`/
  `sumifs_vec_values`/vectorized VLOOKUP — the scalar (non-vectorized) case
  is just this same function called with a batch of size m=1, no separate
  scalar implementation needed
- Input/Output: pandas Series (or scalars) for range/criteria → `Vec<Vec<i64>>`
  → (Python) `mode` selection + column extraction → scalar / Series / dict /
  pandas Series (multi-column scalar) / DataFrame (multi-column vectorized)
- Why this algorithm: `LOOKUPIFS` is architecturally "COUNTIFS's matching
  logic, but return the matched row(s) instead of counting them" — reusing
  the exact same HashMap-of-combined-keys idea (§4.4, COUNTIFS vectorized)
  means no new matching algorithm had to be invented, just a different
  payload stored per key (indices instead of a running count).

---

### 4.5 Multi-Column Return Support — `LOOKUPIFS` and `XLOOKUP`'s multi-column `return_array`

**The feature:** `XLOOKUP`'s `return_array` and `LOOKUPIFS`'s `return_range`
can each be either a single column OR multiple columns at once (a pandas/
polars DataFrame, or a dict of columns), e.g.:
```python
summary[["Salary", "City", "Dept"]] = mx.XLOOKUP(
    summary["ID"], master["ID"], master[["Salary", "City", "Dept"]]
)
mx.LOOKUPIFS(master[["Salary", "City"]], master["Dept"], "IT", master["City"], "Delhi")
```

**The architectural pattern (reuse this for any future function that needs
multi-column output):** Rust's job is ONLY to resolve **which row(s) match**
— it returns row indices, never the actual column values directly, when
multi-column output is involved. Column extraction (picking values out of
one or many requested columns for those matched rows) happens entirely in
the Python layer via plain indexing. This is why `xlookup_many_indices` and
`lookupifs_indices_values` exist as separate Rust functions from
`xlookup_many_values`/`countifs_vec_values` — same underlying matching logic
(same `HashMap`/`Criteria` machinery), but returning indices instead of
resolved values, so Python can reuse ONE match result across an arbitrary
number of requested return columns without re-running the match per column.

**New Python-layer helpers (in `__init__.py`), reusable for future
multi-column functions:**
- `_resolve_return_columns(data)` → `(is_multi: bool, {col_name_or_None: [values...]})`
  — detects pandas/polars DataFrame or dict-of-columns (multi) vs.
  Series/list/tuple/array (single, keyed under `None`).
- `_pick_by_mode(indices, values, mode)` — applies `"first"`/`"last"`/`"all"`
  to one column's values given a list of matching indices; `"all"` joins
  matches into one comma-separated string. Returns the `_NO_MATCH` sentinel
  if `indices` is empty.
- `_NO_MATCH` — internal marker distinct from `_UNSET`/`_NONE_FALLBACK`
  (§4.2): represents "this particular output row had zero matching indices",
  resolved into either `None` (vectorized, soft-fail) or the given
  `if_not_found` value, or a raised `ValueError` (scalar, no fallback given)
  — resolved ONCE per output row, before any column extraction happens, so
  the raise/fallback decision doesn't have to be repeated per column.
- `_build_scalar_multi_result(row: dict)` → `pd.Series(row)` (falls back to
  a plain `dict` if pandas isn't installed) — used for a scalar query with
  multiple return columns. Chosen deliberately so it's "one row of the
  vectorized-case DataFrame" — see the design discussion below.
- `_build_vector_multi_result(columns: dict, origin, index=None)` →
  `pd.DataFrame`/`pl.DataFrame` (matching `origin`'s ecosystem) or a plain
  `dict` of lists as a dependency-free fallback — used for a vectorized
  query with multiple return columns.

**Design decisions confirmed with the user (do not re-litigate these
without a good reason):**
- Scalar lookup + multiple return columns → **pandas Series** (index =
  column names), not a plain dict or a single-row DataFrame — reasoning:
  this is conceptually "one row of the vectorized result", consistent with
  how pandas itself relates `df.loc[single_index]` (Series) to
  `df.loc[list_of_indices]` (DataFrame).
- `mode="all"` + multiple return columns → **each requested column gets its
  own independent comma-joined string** (e.g. `Salary` column becomes
  `"60000, 62000"`, `City` column becomes `"Mumbai, Delhi"` — NOT one merged
  string combining both columns, and NOT a list-of-rows structure).
- `if_not_found` applies **uniformly across every requested return column**
  for a missed row — there's no per-column fallback control, by design (not
  requested, and it would complicate the API for a case that hasn't come up).
- **Not-found behavior mirrors the AVERAGEIF precedent** (§4.4): a scalar
  call with no match and no `if_not_found` given raises `ValueError`
  immediately; a miss inside a **vectorized** batch is filled with `None`
  instead of raising, so one missing row doesn't take down the whole
  result — this was a deliberate, precedent-following choice, not something
  the user explicitly specified for `LOOKUPIFS`/multi-column `XLOOKUP`
  specifically, so revisit it if it turns out to be wrong in practice.

---

## 5. Design History — What Was Tried, What Got Reverted, and Why

This matters for anyone continuing the project: **do not re-introduce the
things described as reverted below without re-reading this section first.**

1. **Early on**, a full numpy zero-copy fast path was built: `sum_f64`,
   `average_f64`, `count_f64`, `countif_f64`, `sumif_f64`, `averageif_f64`,
   `countifs_f64`, `sumifs_f64`, `parse_numeric_criteria`, all operating on
   `numpy::PyReadonlyArray1<f64>` for true zero-copy numeric access, with a
   `numpy` Rust crate dependency and a `_try_f64_array()` Python-side dispatcher
   that tried the fast path first and fell back to the generic path.

2. **This was later fully removed.** Reasons, in order of discovery:
   - Converting a **plain Python list** to a numpy array first (to try the
     fast path) is pure overhead — no real win, since building that numpy
     array from scratch costs about as much as just processing the list
     directly. The fast path should only ever apply to data that's *already*
     array-backed (an existing pandas Series / numpy array / polars Series),
     never to plain lists/tuples.
   - **Mixed-type `SUMIFS` got *slower*, not faster.** When one criteria
     column was numeric (fast-path eligible) and another was text (not
     eligible), the whole call fell back to the fully-generic path — which
     then *also* boxed the numeric column into a `Vec<PyObject>` (via
     `.tolist()`, which for a 2M-row int64 pandas column costs real,
     measurable time just from Python int-object boxing) purely because a
     sibling column happened to be text. A proper fix would need Rust to
     accept a genuinely heterogeneous per-column representation (numeric
     columns stay as zero-copy arrays, text columns as `Vec<PyObject>`,
     decided once per column, not per element) — this was partially designed
     (a `Column::Numeric`/`Column::Generic` enum) but never finished.
   - At this point the user explicitly redirected priorities: **correctness,
     stability, Excel-compatibility, clean architecture, and maintainability
     come first; performance work is a deliberate later phase.** Given the
     half-finished mixed-column design was adding real complexity and had
     already produced a regression once, the entire numpy fast-path layer was
     removed rather than debugged further, and the codebase was returned to
     ONE simple, correct path per function (using the `.downcast()` CellValue
     optimization from §4.1, which is simple enough to keep — it isn't a
     "fast path" in the dual-implementation sense, it's just the more correct
     way to write the one implementation that exists).

3. **What WAS kept from that work, deliberately:**
   - The `.downcast()`-based `CellValue::from_py()` (§4.1) — simple, no
     architectural cost, arguably more correct PyO3 usage than the original
     `.extract()` chain, not just faster.
   - The HashMap-based `vlookup_many_values`/`xlookup_many_values` — a genuine
     algorithmic improvement (not a "fast path" alternative implementation;
     it's the ONLY implementation of vectorized lookup, and O(1)-per-lookup
     is simply the correct way to implement "look up many values against the
     same table", not a premature optimization bolted on afterward).

4. **Later**, when vectorized `*IF`/`*IFS` were requested, the FIRST
   implementation used a Python-level list comprehension calling the scalar
   Rust function once per criteria value. The user pushed back: a loop
   written once should apply to a whole batch automatically (matching how
   dragging an Excel formula down a column works), it should be implemented
   natively in Rust (not a Python loop), and repeated/redundant computation
   should be cached rather than re-run. This led directly to the
   `*_vec_values` functions and the HashMap-fast-path design in §4.3 — this
   one WAS kept, because (a) it's requested/core functionality (vectorized
   criteria IS the feature, not a performance add-on to a working feature),
   and (b) it doesn't introduce the dual-path fragility that the numpy layer
   had — there's still only ONE Rust implementation per function, it's just
   that the vectorized entry point does a smarter thing internally (hashmap
   when possible) while remaining provably correct via the scan fallback.

**Net effect for anyone continuing this project:** the codebase currently has
exactly one implementation per function (plus its `_vec` sibling where a
criteria argument exists). There is no numpy dependency. If a future
performance phase re-introduces zero-copy numeric paths, **do the
heterogeneous-per-column design properly this time** (decide numeric-vs-text
ONCE per column via a Rust-side enum, not via an all-or-nothing Python-side
dtype probe) — the partial attempt at this is what caused the earlier
regression, not the zero-copy idea itself.

## 6. Windows Wheel — Status

Attempted and confirmed **not possible in this dev sandbox**:
- Rust's `x86_64-pc-windows-gnu` target needs its std library installed via
  `rustup target add ...`; this sandbox has no `rustup` (installed via apt
  instead), and `rustup.rs` / `static.rust-lang.org` are not reachable from
  here (network allowlist doesn't include them).
- Confirmed directly: `rustc --target x86_64-pc-windows-gnu` on a trivial
  "hello world" fails with `error[E0463]: can't find crate for std`.
- `mingw-w64` (the linker) WAS installed successfully via apt, but that alone
  isn't sufficient without the target's std library.

**Resolution:** `WINDOWS-BUILD.md` gives simple manual steps to build
natively on an actual Windows machine (install Rust via rustup, `pip install
maturin`, then `maturin build --release` or just `pip install .` from the
project folder) — this is more reliable than cross-compiling would have
been anyway, since it uses a genuine Windows Python + Windows Rust toolchain.
**This has not been tested on an actual Windows machine** — if picking this
up, that's an open verification step.

## 7. What's Explicitly NOT Done Yet (do not assume these exist)

From the user's own "v1 completion plan", deliberately deferred so far
(performance-work items are paused per §5; the rest just haven't been
reached):

- **`AVERAGEIFS`** — Excel's multi-criteria average (SUMIFS's average
  counterpart). Not implemented at all, scalar or vectorized.
- **`MATCH`, `INDEX`, `XMATCH`** — not implemented.
- **Automated pytest suite** — does not exist as a committed file. All
  verification so far has been ad hoc manual scripts run in the dev sandbox
  (`test_magpiexl.py` covers all 10 functions across standalone/pandas/
  polars/csv/xlsx; further inline checks covered wildcards, error handling,
  the `if_not_found` sentinel, tuple preservation, and the cross-table
  `COUNTIF`/`SUMIFS` vectorization scenarios) — none of this is in a real
  `pytest` file yet. The user's plan asks for 200+ tests covering scalar/
  vector/pandas/numpy/list/tuple/polars/Arrow/missing values/duplicates/
  operators/cross-DataFrame/not-found/empty-data/mismatched-lengths.
- **Performance benchmarks** vs pandas/polars/numpy at 100K/1M/5M rows — not
  done; performance work is paused (§5).
- **Arrow array support** — `_wrap_like` does not handle `pyarrow.Array`;
  only pandas/polars/numpy/list/tuple are covered.
- **README "showcase" example** — the user's plan describes a specific
  HR/Sales/Finance cross-DataFrame report example meant to be the first
  thing a new user sees; the current `README.md` has simpler examples, not
  this specific showcase.
- **Error message rewrite** — some errors are still Excel-terse (e.g.
  `"VLOOKUP: #N/A - value not found"`) rather than the more descriptive
  Python-style wording the user suggested (e.g. `"XLOOKUP: Value 'Legal' not
  found. Use if_not_found=... to specify a default value."`).
- **Fully generic vectorization layer** — `_prepare_ifs_vector` is shared
  between `COUNTIFS`/`SUMIFS`, but `COUNTIF`/`SUMIF`/`AVERAGEIF` each still
  have their own near-identical scalar-vs-array dispatch block rather than
  one single generic wrapper/decorator used by all six `*IF`/`*IFS`
  functions. Works correctly today; a refactor here is optional cleanup, not
  a correctness fix.

## 8. Established Conventions (follow these for anything new)

1. **Ask before starting a new coding direction** (new functions, new
   dependencies, architecture changes) — the user has repeatedly asked for
   this; don't assume permission carries over between separate feature asks.
2. Correctness, Excel-compatibility, and clean architecture outrank raw
   performance **until the user explicitly asks to optimize** — see §5 for
   what happened last time this was ignored.
3. Don't convert a type that doesn't need converting (`_to_list` already
   short-circuits on `isinstance(data, list)` — preserve this pattern in
   anything new).
4. Don't recompute the same thing repeatedly inside one call — cache it
   (this is *why* the HashMap fast path in §4.3 exists; apply the same
   instinct to any new vectorized function).
5. **Same type in, same type out** for anything returning array-shaped data:
   list→list, tuple→tuple, pandas Series→pandas Series (index preserved),
   polars Series→polars Series, numpy array→numpy array. Use `_wrap_like`.
6. Any new optional parameter whose valid value legitimately includes `None`
   (distinct from "not provided") needs the `_UNSET`/sentinel pattern from
   §4.2 — don't default it to `None` directly.
7. New Excel functions should follow the existing shape: a scalar Rust
   `*_values` function; if it takes a criteria argument, also a `*_vec_values`
   Rust function (with the equality-HashMap-fast-path-then-scan-fallback
   pattern if applicable); a Python wrapper doing scalar-vs-array
   auto-detection + `_wrap_like`.

## 9. Suggested Next Steps (priority order, per the user's own plan)

1. Finish "one generic vectorization layer" cleanup (§7, last bullet) —
   optional, low-risk.
2. Implement `AVERAGEIFS` (straightforward — same shape as `SUMIFS` but
   averaging, including a vectorized `_vec_values` sibling).
3. Write the actual pytest suite (§7) — this is listed as the user's
   **highest priority after vectorization**, and vectorization is now done.
4. `MATCH` / `INDEX` / `XMATCH` — new functions, not yet started.
5. Rewrite error messages to be more descriptive (§7).
6. Only after the above: performance benchmarking and any zero-copy
   optimization work — re-read §5 point 6 before starting this.

---

*Last updated: end of the vectorized-`*IF`/`*IFS` + `if_not_found`-sentinel +
Windows-build-investigation work. Current wheel: `magpiexl-0.1.0-cp312-cp312-
manylinux_2_34_x86_64.whl`. Source zip and this file should always be handed
over together.*
