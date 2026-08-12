use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyCapsule, PyDict, PyFloat, PyInt, PyString, PyStringMethods};
use std::collections::HashMap;
use std::os::raw::{c_char, c_int, c_void};

use arrow::array::{make_array, Array, ArrayRef, LargeStringArray, StringArray};
use arrow::datatypes::DataType;
use arrow::ffi::{from_ffi, FFI_ArrowArray, FFI_ArrowSchema};
use rayon::prelude::*;

// ============================================================
// Hand-rolled Arrow C Data/Stream Interface reader for the native
// Utf8View ("StringView") format, plus classic Utf8/LargeUtf8 — all
// via our OWN #[repr(C)] structs matching the Arrow C ABI directly,
// completely independent of the `arrow` crate's version-specific
// typed wrappers (which, at the pinned v45 this build uses, do not
// support Utf8View at all — confirmed to panic internally on it).
//
// Why this exists: the Arrow C Data/Stream Interface's base structs
// (ArrowSchema, ArrowArray, ArrowArrayStream) are an explicitly FROZEN
// C ABI per the spec itself — "Once this specification is supported
// in an official Arrow release, the C ABI is frozen... structure
// definitions should not change in any way" (arrow.apache.org/docs/
// format/CDataInterface.html). That means defining our own matching
// struct layout here is safe and future-proof — it does not depend on
// which `arrow` crate version (or even which Arrow implementation:
// polars, pyarrow, anything) is on the producer's side, only on the
// spec itself. This is the same reason pyarrow/polars/DataFusion can
// all interoperate without sharing a Rust crate version.
//
// The exact StringView byte layout used below (16-byte "view" per
// row: 4-byte length, then either 12 bytes of inline data for
// length<=12, or a 4-byte prefix + 4-byte buffer index + 4-byte
// offset for longer strings) was verified against the spec
// (arrow.apache.org/docs/format/Columnar.html) AND against real
// bytes read directly out of a live polars-produced capsule via
// ctypes before writing this — see HANDOFF notes for the exact
// verification transcript. Getting this wrong would silently corrupt
// string comparisons rather than crash, which is why it was checked
// against real data first, not just the spec text.
mod raw_arrow {
    use super::*;

    #[repr(C)]
    pub struct CArrowSchema {
        pub format: *const c_char,
        pub name: *const c_char,
        pub metadata: *const c_char,
        pub flags: i64,
        pub n_children: i64,
        pub children: *mut *mut CArrowSchema,
        pub dictionary: *mut CArrowSchema,
        pub release: Option<unsafe extern "C" fn(*mut CArrowSchema)>,
        pub private_data: *mut c_void,
    }

    #[repr(C)]
    pub struct CArrowArray {
        pub length: i64,
        pub null_count: i64,
        pub offset: i64,
        pub n_buffers: i64,
        pub n_children: i64,
        pub buffers: *mut *const c_void,
        pub children: *mut *mut CArrowArray,
        pub dictionary: *mut CArrowArray,
        pub release: Option<unsafe extern "C" fn(*mut CArrowArray)>,
        pub private_data: *mut c_void,
    }

    #[repr(C)]
    pub struct CArrowArrayStream {
        pub get_schema:
            Option<unsafe extern "C" fn(*mut CArrowArrayStream, *mut CArrowSchema) -> c_int>,
        pub get_next:
            Option<unsafe extern "C" fn(*mut CArrowArrayStream, *mut CArrowArray) -> c_int>,
        pub get_last_error: Option<unsafe extern "C" fn(*mut CArrowArrayStream) -> *const c_char>,
        pub release: Option<unsafe extern "C" fn(*mut CArrowArrayStream)>,
        pub private_data: *mut c_void,
    }

    /// One column's worth of parsed text data plus its null mask. Owns
    /// its `String`s (copied out of the producer's buffers) rather than
    /// borrowing them — this keeps lifetime management trivial and
    /// safe (no need to keep any Python object alive once this is
    /// built, no risk of a dangling pointer if we got a release-timing
    /// detail wrong) in exchange for one allocation per row, which is
    /// still vastly cheaper than the per-row PyObject path this whole
    /// effort exists to avoid, and cheaper than the `to_arrow(oldest)`
    /// conversion this bypasses entirely for the common (StringView)
    /// case.
    pub struct ParsedTextColumn {
        pub values: Vec<Option<String>>,
    }

    unsafe fn read_validity(
        buffers: &[*const c_void],
        offset: i64,
        i: i64,
        null_count: i64,
    ) -> bool {
        // Per spec: buffer 0 (validity bitmap) may be a null pointer
        // when null_count == 0 — that's the common case and the only
        // one we treat as "definitely not null" without touching any
        // buffer, matching every other null-check in this file.
        if null_count == 0 || buffers.is_empty() || buffers[0].is_null() {
            return false;
        }
        let byte_idx = ((offset + i) / 8) as isize;
        let bit_idx = ((offset + i) % 8) as u32;
        let byte = *(buffers[0] as *const u8).offset(byte_idx);
        (byte & (1 << bit_idx)) == 0
    }

    /// Parses one already-fetched `CArrowArray` as a Utf8View
    /// ("StringView") array per the verified byte layout described
    /// above. `n_buffers` here is at least 3 (validity, views,
    /// variadic-sizes) even with zero variadic data buffers (per the
    /// C Data Interface's own note that view types always carry the
    /// extra sizes buffer, even when there's nothing to size).
    unsafe fn parse_stringview_array(arr: &CArrowArray) -> ParsedTextColumn {
        let n = arr.length;
        let nb = arr.n_buffers as usize;
        let buffers: &[*const c_void] = std::slice::from_raw_parts(arr.buffers as *const _, nb);
        let mut values = Vec::with_capacity(n as usize);
        if n == 0 || nb < 3 {
            return ParsedTextColumn { values };
        }
        let views_ptr = buffers[1] as *const u8;
        // Variadic data buffers are buffers[2..nb-1]; buffers[nb-1] is
        // the sizes-of-those-buffers array (int64 each) — we don't
        // actually need the sizes for reading (we trust length+offset
        // from the view itself, which the producer guarantees are
        // in-bounds per spec), only the data pointers themselves.
        let variadic_data_buffers = &buffers[2..nb - 1];
        for i in 0..n {
            if read_validity(buffers, arr.offset, i, arr.null_count) {
                values.push(None);
                continue;
            }
            let view_offset = ((arr.offset + i) as isize) * 16;
            let view = std::slice::from_raw_parts(views_ptr.offset(view_offset), 16);
            let length = i32::from_le_bytes([view[0], view[1], view[2], view[3]]);
            if length <= 12 {
                let bytes = &view[4..4 + length as usize];
                // Lossy, not strict UTF-8 validation: matches this
                // file's existing convention everywhere else
                // (`to_string_lossy`) rather than introducing a new
                // failure mode here specifically.
                values.push(Some(String::from_utf8_lossy(bytes).into_owned()));
            } else {
                let buf_idx = i32::from_le_bytes([view[8], view[9], view[10], view[11]]) as usize;
                let data_offset =
                    i32::from_le_bytes([view[12], view[13], view[14], view[15]]) as isize;
                if buf_idx >= variadic_data_buffers.len() {
                    // Defensive: a spec-conformant producer never
                    // produces an out-of-range buffer index, but this
                    // is exactly the kind of check worth keeping
                    // between "trust the producer" and "silently read
                    // garbage memory" if that assumption is ever wrong.
                    values.push(None);
                    continue;
                }
                let data_ptr = (variadic_data_buffers[buf_idx] as *const u8).offset(data_offset);
                let bytes = std::slice::from_raw_parts(data_ptr, length as usize);
                values.push(Some(String::from_utf8_lossy(bytes).into_owned()));
            }
        }
        ParsedTextColumn { values }
    }

    /// Parses a classic Utf8 (32-bit offsets) or LargeUtf8 (64-bit
    /// offsets) array via the same raw-struct approach — standard
    /// offsets-buffer + data-buffer variable-size binary layout, no
    /// StringView involved. `large` selects the offset integer width.
    unsafe fn parse_offset_array(arr: &CArrowArray, large: bool) -> ParsedTextColumn {
        let n = arr.length;
        let nb = arr.n_buffers as usize;
        let buffers: &[*const c_void] = std::slice::from_raw_parts(arr.buffers as *const _, nb);
        let mut values = Vec::with_capacity(n as usize);
        if n == 0 || nb < 3 {
            return ParsedTextColumn { values };
        }
        let data_ptr = buffers[2] as *const u8;
        for i in 0..n {
            if read_validity(buffers, arr.offset, i, arr.null_count) {
                values.push(None);
                continue;
            }
            let row = arr.offset + i;
            let (start, end) = if large {
                let offs = buffers[1] as *const i64;
                (*offs.offset(row as isize), *offs.offset(row as isize + 1))
            } else {
                let offs = buffers[1] as *const i32;
                (
                    *offs.offset(row as isize) as i64,
                    *offs.offset(row as isize + 1) as i64,
                )
            };
            let len = (end - start) as usize;
            let bytes = std::slice::from_raw_parts(data_ptr.offset(start as isize), len);
            values.push(Some(String::from_utf8_lossy(bytes).into_owned()));
        }
        ParsedTextColumn { values }
    }

    /// Reads `obj.__arrow_c_stream__()`, sniffs the schema's format
    /// string, and — only for formats we've explicitly verified byte-
    /// for-byte ("vu" Utf8View, "u" Utf8, "U" LargeUtf8) — parses every
    /// batch in the stream into one flat `ParsedTextColumn`, then
    /// releases the stream/array resources itself (we own everything
    /// we need by the time we're done, so no Python object needs to
    /// stay alive past this function). Returns `Ok(None)` for anything
    /// else (no `__arrow_c_stream__` at all, an unrecognized format,
    /// or any producer-side error) — never an error for those routine
    /// cases, so the caller falls back to the existing, already-
    /// verified paths.
    pub fn try_read_stream_as_text(obj: &Bound<PyAny>) -> PyResult<Option<ParsedTextColumn>> {
        let method = match obj.getattr("__arrow_c_stream__") {
            Ok(m) => m,
            Err(_) => return Ok(None),
        };
        let capsule_obj = method.call0()?;
        let capsule = match capsule_obj.downcast::<PyCapsule>() {
            Ok(c) => c,
            Err(_) => return Ok(None),
        };
        match capsule.name() {
            Ok(Some(n)) if n.to_string_lossy() == "arrow_array_stream" => {}
            _ => return Ok(None),
        }
        let stream_ptr = capsule.pointer() as *mut CArrowArrayStream;
        if stream_ptr.is_null() {
            return Ok(None);
        }
        // Safety: capsule name verified above matches the Arrow C
        // Stream Interface's mandated constant, so this points to a
        // validly-initialized `CArrowArrayStream` (whose layout matches
        // the frozen C ABI) per the exporter's contract. We only ever
        // BORROW through this pointer (call its function-pointer
        // fields) — we never take ownership of the stream struct
        // itself via ptr::read the way the old FFI_ArrowArrayStream
        // path did, so there is no double-free risk to manage here:
        // the original Python capsule remains fully responsible for
        // its own `release`, whenever Python decides to run that
        // capsule's destructor. We call the stream's own `release`
        // explicitly ourselves once we're done pulling batches, since
        // the Arrow C Stream Interface expects the consumer to do
        // that — but per the interface's own convention, invoking
        // `release` on an interface struct is required to be safely
        // idempotent/no-op-safe if the destructor also runs it, since
        // producers are required to set `release` to NULL once called.
        unsafe {
            let stream = &mut *stream_ptr;
            let get_schema = match stream.get_schema {
                Some(f) => f,
                None => return Ok(None),
            };
            let get_next = match stream.get_next {
                Some(f) => f,
                None => return Ok(None),
            };
            let mut schema = std::mem::zeroed::<CArrowSchema>();
            if get_schema(stream_ptr, &mut schema) != 0 {
                return Ok(None);
            }
            if schema.format.is_null() {
                return Ok(None);
            }
            let format = std::ffi::CStr::from_ptr(schema.format)
                .to_string_lossy()
                .into_owned();
            if let Some(release) = schema.release {
                release(&mut schema);
            }
            let is_stringview = format == "vu";
            let is_utf8 = format == "u";
            let is_large_utf8 = format == "U";
            if !is_stringview && !is_utf8 && !is_large_utf8 {
                // Anything else (numeric types also implement this
                // protocol; binary/other text encodings we haven't
                // explicitly verified) — not our concern here, and
                // NOT safe to guess at, given a wrong guess here means
                // silently misreading memory rather than a visible
                // error. Existing fallback paths handle these.
                //
                // We still need to release the stream itself before
                // returning, since we're abandoning it without reading
                // any batches — otherwise the producer-side resources
                // for this stream would leak (not a memory-safety bug,
                // but a real resource leak on every such call).
                if let Some(release) = stream.release {
                    release(stream_ptr);
                }
                return Ok(None);
            }
            let mut all_values: Vec<Option<String>> = Vec::new();
            loop {
                let mut array = std::mem::zeroed::<CArrowArray>();
                if get_next(stream_ptr, &mut array) != 0 {
                    // Producer-side error mid-stream — bail out to the
                    // fallback path rather than return partial data.
                    if let Some(release) = stream.release {
                        release(stream_ptr);
                    }
                    return Ok(None);
                }
                if array.release.is_none() {
                    // A released/empty ArrowArray with a NULL release
                    // callback is the spec's defined "end of stream"
                    // signal — not an error, just done.
                    break;
                }
                let parsed = if is_stringview {
                    parse_stringview_array(&array)
                } else {
                    parse_offset_array(&array, is_large_utf8)
                };
                all_values.extend(parsed.values);
                if let Some(release) = array.release {
                    release(&mut array);
                }
            }
            if let Some(release) = stream.release {
                release(stream_ptr);
            }
            Ok(Some(ParsedTextColumn { values: all_values }))
        }
    }
}

// ============================================================
// Arrow zero-copy text ingestion — this is what lets a polars
// Utf8/String Series reach Rust the same way polars' OWN native
// operations do: no per-row PyObject at all, ever.
//
// Path taken: the Python layer calls `series.to_arrow(compat_level=
// pl.CompatLevel.oldest())` before handing the result to Rust. This
// matters because polars' *default*/newest export uses the newer
// Utf8View ("StringView") Arrow layout, which this version of the
// `arrow` crate (45.0.0 — pinned this low by transitive-dependency MSRV
// constraints in this build environment; see Cargo.toml/HANDOFF notes)
// does not understand and, worse, panics on internally rather than
// returning a normal error. Forcing the older/"oldest" compat level
// makes polars export as classic LargeUtf8 instead, which this crate
// version handles natively and safely. The `to_arrow(oldest)` call
// itself isn't free (confirmed by benchmark: tens of ms, not zero) —
// but the subsequent Rust-side read is genuinely zero-copy off of that
// buffer, and the whole thing is still an order of magnitude faster
// than the old per-row-PyObject path for any column past a few
// thousand rows.
//
// The resulting pyarrow Array is consumed via `__arrow_c_array__`
// (the single-array PyCapsule pair: one "arrow_schema" capsule, one
// "arrow_array" capsule — see
// https://arrow.apache.org/docs/format/CDataInterface/PyCapsuleInterface.html),
// which is simpler and more predictable than the streaming interface
// for this case (one column in, one array out, no batching).
// ============================================================

enum ArrowStringArrayRef {
    Utf8(StringArray),
    LargeUtf8(LargeStringArray),
}

/// Owns the imported Arrow array (keeping its underlying buffers
/// alive for as long as this column is alive) so a borrowed `&str`
/// for any row index can be produced with no copy and no allocation.
struct ArrowTextColumn {
    array: ArrowStringArrayRef,
}

impl ArrowTextColumn {
    fn len(&self) -> usize {
        match &self.array {
            ArrowStringArrayRef::Utf8(a) => a.len(),
            ArrowStringArrayRef::LargeUtf8(a) => a.len(),
        }
    }

    /// Fast per-row equality/inequality mask against a scalar needle —
    /// the same shape as polars' own `tot_eq_kernel_broadcast` (see
    /// polars-compute's `comparisons/binary.rs`:
    /// `self.values_iter().map(|l| l.tot_eq(&other)).collect()`), plus
    /// optional multi-core parallelism via `rayon` for large columns.
    ///
    /// IMPORTANT — what actually explains the remaining gap to polars'
    /// own speed, verified rather than assumed: polars' native export
    /// format is Utf8View ("StringView"), which inlines short strings
    /// (roughly <=12 bytes — covers most real category/label data, e.g.
    /// "East"/"West"/"Inactive") directly in the array's fixed-size
    /// row record, with NO separate variable-length buffer dereference
    /// at all for those rows. Our path here reads classic Utf8/
    /// LargeUtf8 (offset array + separate data buffer — the format we
    /// can get polars to export via `compat_level=oldest()`, since the
    /// `arrow` crate version this build is pinned to predates
    /// StringView support), which costs an extra pointer chase per row
    /// that StringView's inlined short strings avoid. This was
    /// confirmed by directly measuring polars' own native comparison on
    /// THIS SAME single-core sandbox (~1.4-3ms at 1M rows) against a
    /// pure-Rust single-threaded loop over equivalent owned `String`
    /// data (~6-9ms) — both single-threaded, same machine, and still a
    /// real gap — so the difference isn't primarily about threading.
    /// Closing that specific gap fully would need a StringView-aware
    /// read path, which needs a newer `arrow` crate than this MSRV-
    /// constrained build environment can compile (see Cargo.toml
    /// notes). The rayon parallel branch below is still a genuine,
    /// separate win on real multi-core machines for large columns — it
    /// divides the per-row work (whatever that cost is) across cores,
    /// independent of whether that per-row cost itself is StringView-
    /// optimized or not; it just can't be measured here on a 1-core box.
    ///
    /// Three things this avoids compared to the generic per-row path:
    /// 1. `value_unchecked` — `value(i)`'s bounds-check `assert!` fires
    ///    on every row; skipped here since we control the loop range.
    /// 2. Null-checking is skipped ENTIRELY when `null_count() == 0`
    ///    (checked once, not per row) — the overwhelmingly common case
    ///    for a real data column.
    /// 3. `needle`'s asciiness is checked ONCE outside the loop rather
    ///    than per row inside `text_eq_ci` — safe because ASCII-ness of
    ///    one side doesn't change across rows; only the haystack's
    ///    asciiness genuinely needs a per-row check (still done, but
    ///    it's a cheap stdlib-intrinsic scan over the row's own bytes,
    ///    not the extra needle-side redundant check).
    ///
    /// Case-insensitive (matching Excel's COUNTIF/SUMIF semantics,
    /// same as `text_eq_ci` elsewhere in this file) — so this is
    /// necessarily doing more work per row than polars' raw
    /// case-SENSITIVE `==`, and won't be quite as fast as polars' own
    /// kernel for that reason; it closes the dispatch-overhead and
    /// parallelism gaps, not the "different operation" gap.
    fn eq_mask(&self, needle: &str, negate: bool) -> Vec<bool> {
        // Below this row count, rayon's thread-pool dispatch/join
        // overhead is likely to cost more than the parallelism saves —
        // this is a conservative starting threshold, not a precisely
        // tuned one (tuning it further needs a real multi-core machine
        // to measure against, which this build environment doesn't
        // have). Callers on small data still get the fast sequential
        // path, never regressed by parallel dispatch overhead.
        const PARALLEL_THRESHOLD: usize = 50_000;
        let needle_ascii = needle.is_ascii();
        let n = self.len();

        #[inline(always)]
        fn cell_eq(hay: &str, needle: &str, needle_ascii: bool) -> bool {
            if needle_ascii && hay.is_ascii() {
                hay.eq_ignore_ascii_case(needle)
            } else {
                hay.to_lowercase() == needle.to_lowercase()
            }
        }

        macro_rules! run {
            ($arr:expr) => {{
                let arr = $arr;
                let no_nulls = arr.null_count() == 0;
                // Safety (both branches): every `i` comes from `0..n`
                // where `n == arr.len()`, so `value_unchecked(i)` is
                // always in bounds.
                let compute_one = |i: usize| -> bool {
                    if !no_nulls && arr.is_null(i) {
                        // Empty-cell semantics: Eq to a non-empty
                        // needle is false, Ne is true — matches
                        // `matches(&CellValue::Empty, crit)` elsewhere.
                        return negate;
                    }
                    let hay = unsafe { arr.value_unchecked(i) };
                    let eq = cell_eq(hay, needle, needle_ascii);
                    if negate {
                        !eq
                    } else {
                        eq
                    }
                };
                if n >= PARALLEL_THRESHOLD {
                    (0..n).into_par_iter().map(compute_one).collect()
                } else {
                    (0..n).map(compute_one).collect()
                }
            }};
        }
        let mask: Vec<bool> = match &self.array {
            ArrowStringArrayRef::Utf8(a) => run!(a),
            ArrowStringArrayRef::LargeUtf8(a) => run!(a),
        };
        mask
    }

    /// Borrowed `&str` at row `i`, or `None` for a null cell (Arrow's
    /// own null bitmap, not a sentinel value).
    fn get(&self, i: usize) -> Option<&str> {
        match &self.array {
            ArrowStringArrayRef::Utf8(a) => {
                if a.is_null(i) {
                    None
                } else {
                    Some(a.value(i))
                }
            }
            ArrowStringArrayRef::LargeUtf8(a) => {
                if a.is_null(i) {
                    None
                } else {
                    Some(a.value(i))
                }
            }
        }
    }
}

/// Tries to read `obj` as an Arrow-exporting single array (a pyarrow
/// `Array`/`ChunkedArray` — what `polars.Series.to_arrow(...)`
/// returns — or anything else implementing `__arrow_c_array__`,
/// https://arrow.apache.org/docs/format/CDataInterface/PyCapsuleInterface.html)
/// and, if it's a Utf8/LargeUtf8 array, returns a zero-copy
/// `ArrowTextColumn`. Returns `Ok(None)` (not an error) for anything
/// that doesn't support the protocol, or whose type isn't string —
/// both are routine, not failures; the caller falls back to the
/// existing paths in either case.
fn try_arrow_text_column(obj: &Bound<PyAny>) -> PyResult<Option<ArrowTextColumn>> {
    let method = match obj.getattr("__arrow_c_array__") {
        Ok(m) => m,
        Err(_) => return Ok(None),
    };
    let result = method.call0()?;
    let (schema_capsule, array_capsule): (Bound<PyCapsule>, Bound<PyCapsule>) =
        match result.extract() {
            Ok(pair) => pair,
            Err(_) => return Ok(None),
        };
    // Defensive: only proceed if these are really the capsules the
    // protocol promises (names are spec-mandated constants).
    match schema_capsule.name() {
        Ok(Some(n)) if n.to_string_lossy() == "arrow_schema" => {}
        _ => return Ok(None),
    }
    match array_capsule.name() {
        Ok(Some(n)) if n.to_string_lossy() == "arrow_array" => {}
        _ => return Ok(None),
    }
    let schema_ptr = schema_capsule.pointer() as *mut FFI_ArrowSchema;
    let array_ptr = array_capsule.pointer() as *mut FFI_ArrowArray;

    // Safety: both pointers come from capsules whose names match the
    // Arrow C Data Interface's mandated constants, so each points to a
    // validly-initialized struct per the exporter's contract.
    //
    // `ptr::read` moves the struct's bytes out by value WITHOUT
    // clearing the source location — the source (the capsule's own
    // memory) still has the same `release` function pointer sitting in
    // it afterward. If we stopped here, BOTH our new owned value's
    // `Drop` (which calls `release`) AND the Python capsule's own
    // destructor (which also calls whatever `release` it finds still
    // sitting in that memory) would independently invoke that release
    // callback — a real double-free. This was tried and confirmed to
    // crash (`free(): invalid pointer`) before this fix was added.
    //
    // The correct handoff — matching what a spec-compliant consumer of
    // an "R" (recognized, so-called "guarded") pull-style capsule must
    // do per the Arrow PyCapsule Interface — is to immediately write a
    // release-free placeholder (`empty()`, whose `release` field is
    // `None`) back into the original capsule location. That makes the
    // capsule's own eventual destructor a no-op, so our owned copy (via
    // `ptr::read`) is the only thing left holding a live `release`
    // pointer, and only its `Drop` will ever call it.
    let schema: FFI_ArrowSchema = unsafe {
        let s = std::ptr::read(schema_ptr);
        std::ptr::write(schema_ptr, FFI_ArrowSchema::empty());
        s
    };
    let array: FFI_ArrowArray = unsafe {
        let a = std::ptr::read(array_ptr);
        std::ptr::write(array_ptr, FFI_ArrowArray::empty());
        a
    };

    // Unlike the streaming interface's schema parser (which panics on
    // an unrecognized format string), `from_ffi` on this arrow version
    // returns a normal `Result` — no `catch_unwind` needed here. Still
    // guard defensively: if the caller passed something already in the
    // newer Utf8View layout despite the Python-side `oldest()` request
    // (e.g. a future polars behavior change, or a non-polars caller),
    // this simply returns `Ok(None)` and the existing fallback paths
    // take over — never a crash.
    let array_data = match unsafe { from_ffi(array, &schema) } {
        Ok(d) => d,
        Err(_) => return Ok(None),
    };
    let array_ref: ArrayRef = make_array(array_data);
    let arrow_array = match array_ref.data_type() {
        DataType::Utf8 => {
            let a = array_ref
                .as_any()
                .downcast_ref::<StringArray>()
                .expect("data_type() said Utf8")
                .clone();
            ArrowStringArrayRef::Utf8(a)
        }
        DataType::LargeUtf8 => {
            let a = array_ref
                .as_any()
                .downcast_ref::<LargeStringArray>()
                .expect("data_type() said LargeUtf8")
                .clone();
            ArrowStringArrayRef::LargeUtf8(a)
        }
        // Not a string type (numeric columns also implement this same
        // protocol) — not our concern here, existing numeric fast
        // paths already handle those. Not an error, just "not for us".
        _ => return Ok(None),
    };
    Ok(Some(ArrowTextColumn { array: arrow_array }))
}

// ============================================================
// Core value model — mirrors how Excel treats a "cell":
// a cell is either a number, text, or empty. Booleans behave
// like numbers (TRUE=1, FALSE=0), same as in Excel formulas.
//
// PERFORMANCE NOTE: we use `downcast::<T>()` here, not
// `extract::<T>()`. downcast is a cheap pointer/type-tag check
// with no Python exception machinery involved. The generic
// `.extract::<f64>()` chain we used originally raises and
// discards a real Python exception on every failed attempt,
// which is exactly what made the first version slower than a
// plain Python loop on large data. This single change is the
// biggest lever in this file.
// ============================================================

#[derive(Clone, Debug)]
enum CellValue {
    Num(f64),
    Text(String),
    Empty,
}

impl CellValue {
    #[inline]
    fn from_py(obj: &Bound<PyAny>) -> CellValue {
        if obj.is_none() {
            return CellValue::Empty;
        }
        if let Ok(b) = obj.downcast::<PyBool>() {
            return CellValue::Num(if b.is_true() { 1.0 } else { 0.0 });
        }
        if let Ok(f) = obj.downcast::<PyFloat>() {
            let v = f.value();
            if v.is_nan() {
                return CellValue::Empty;
            }
            return CellValue::Num(v);
        }
        if let Ok(i) = obj.downcast::<PyInt>() {
            if let Ok(n) = i.extract::<f64>() {
                return CellValue::Num(n);
            }
        }
        if let Ok(s) = obj.downcast::<PyString>() {
            let text = s.to_string_lossy().into_owned();
            if text.trim().is_empty() {
                return CellValue::Empty;
            }
            return CellValue::Text(text);
        }
        // Anything that isn't a plain Python bool/float/int/str but IS
        // still genuinely numeric — the case this exists for is a numpy
        // scalar (np.int8, np.uint32, np.float32, ...) sitting loose
        // inside an otherwise-plain list/tuple, which happens whenever
        // code iterates a numpy array directly instead of calling
        // `.tolist()` first (`[x for x in arr]`, `list(arr)`, unpacking,
        // etc. all yield numpy scalar objects, not Python ints/floats —
        // confirmed directly: `type(next(iter(np.array([1], "int32"))))`
        // is `numpy.int32`, not `int`). Before this branch existed, none
        // of the downcasts above matched such a value, so it fell
        // through to the final `obj.str()` branch below and silently
        // became `CellValue::Text("5")` — a real, silent, wrong-answer
        // bug: `SUM([x for x in np.array([1,2,3], "int32")])` returned
        // `0.0` instead of `6.0`, with no error or warning anywhere.
        //
        // `.extract::<f64>()` here (as opposed to the `PyInt`-only one
        // above) uses PyO3's generic f64 extraction, which goes through
        // CPython's `PyFloat_AsDouble` — this succeeds for any object
        // implementing Python's number protocol (`__float__` or
        // `__index__`), which every numpy scalar type does by design
        // (numpy scalars are built to be duck-type-compatible with
        // Python numbers), and also correctly leaves alone anything
        // that ISN'T actually numeric (a plain object, list, dict, or
        // custom class with no numeric protocol simply fails this
        // extraction and falls through to the stringify branch below,
        // completely unchanged from before this fix).
        if let Ok(n) = obj.extract::<f64>() {
            if n.is_nan() {
                return CellValue::Empty;
            }
            return CellValue::Num(n);
        }
        match obj.str() {
            Ok(s) => CellValue::Text(s.to_string()),
            Err(_) => CellValue::Empty,
        }
    }
}

/// Case-insensitive text equality — ASCII-fast, Unicode-correct.
///
/// `eq_ignore_ascii_case()` alone is wrong for non-ASCII text: it only
/// folds the ASCII letter range, so e.g. "CAFÉ" vs "café" would compare
/// the É/é bytes raw (unequal) even though a human — and Excel — would
/// consider them the same word case-insensitively. Checking `is_ascii()`
/// first (cheap, no allocation) lets the overwhelmingly common case (pure
/// ASCII data) take the zero-allocation fast path, while non-ASCII text
/// still gets correct Unicode case folding via `to_lowercase()`.
#[inline]
fn text_eq_ci(a: &str, b: &str) -> bool {
    if a.is_ascii() && b.is_ascii() {
        a.eq_ignore_ascii_case(b)
    } else {
        a.to_lowercase() == b.to_lowercase()
    }
}

/// Same fast, branch-light equality-mask approach as
/// `ArrowTextColumn::eq_mask`, adapted for `NativeText`'s owned
/// `Vec<Option<String>>` (already-parsed strings, no separate
/// buffer/offset indirection to walk — just a straight, optionally
/// parallel, pass). Kept as a free function rather than a method so
/// both `NativeText`'s `try_fast_eq_mask` arm and any future caller
/// can use it without needing a wrapper type.
fn native_text_eq_mask(v: &[Option<String>], needle: &str, negate: bool) -> Vec<bool> {
    const PARALLEL_THRESHOLD: usize = 50_000;
    let needle_ascii = needle.is_ascii();
    #[inline(always)]
    fn cell_eq(hay: &str, needle: &str, needle_ascii: bool) -> bool {
        if needle_ascii && hay.is_ascii() {
            hay.eq_ignore_ascii_case(needle)
        } else {
            hay.to_lowercase() == needle.to_lowercase()
        }
    }
    let compute_one = |cell: &Option<String>| -> bool {
        match cell {
            Some(hay) => {
                let eq = cell_eq(hay.as_str(), needle, needle_ascii);
                if negate {
                    !eq
                } else {
                    eq
                }
            }
            // Null cell: matches `matches(&CellValue::Empty, crit)`'s
            // `(CellValue::Empty, _) => matches!(criteria.op, Op::Ne)`
            // branch — Eq (negate=false) is false, Ne (negate=true) is
            // true, i.e. the result is simply `negate` itself.
            None => negate,
        }
    };
    if v.len() >= PARALLEL_THRESHOLD {
        v.par_iter().map(compute_one).collect()
    } else {
        v.iter().map(compute_one).collect()
    }
}

fn values_equal(a: &CellValue, b: &CellValue) -> bool {
    match (a, b) {
        (CellValue::Num(x), CellValue::Num(y)) => (x - y).abs() < 1e-9,
        (CellValue::Text(x), CellValue::Text(y)) => text_eq_ci(x, y),
        (CellValue::Empty, CellValue::Empty) => true,
        _ => false,
    }
}

// Hashable key used to build an O(1) lookup table for vectorized
// VLOOKUP/XLOOKUP (looking up many values against the same
// table/array). Building this map once and querying it N times is
// O(n + m) total, instead of O(n * m) for a linear scan per lookup.
#[derive(Clone, PartialEq)]
enum LookupKey {
    Num(f64),
    Text(String),
    Empty,
}
impl Eq for LookupKey {}
impl std::hash::Hash for LookupKey {
    fn hash<H: std::hash::Hasher>(&self, state: &mut H) {
        match self {
            LookupKey::Num(n) => n.to_bits().hash(state),
            LookupKey::Text(s) => s.hash(state),
            LookupKey::Empty => 0u8.hash(state),
        }
    }
}

/// A row's text value, prepared for a case-insensitive HashMap probe
/// with minimal allocation — see `FastColumn::cell_text_ref_at`'s doc
/// comment for the full rationale. `Borrowed` carries NO allocation
/// at all (a `&str` borrowed from the column's own zero-copy buffer);
/// `Owned` is the fallback for text that genuinely needed
/// `.to_lowercase()`'s full Unicode-aware case folding.
enum CellTextRef<'a> {
    Borrowed(&'a str),
    Owned(String),
}

fn cell_to_key(cv: &CellValue) -> Option<LookupKey> {
    match cv {
        CellValue::Num(n) => Some(LookupKey::Num(*n)),
        CellValue::Text(s) => Some(LookupKey::Text(s.to_lowercase())),
        // Blank cells ARE a matchable key now — this is what makes the
        // HashMap-based vectorized lookups agree with the scalar linear
        // scan (which already matched Empty == Empty via values_equal()).
        // Before this fix, cell_to_key returned None for Empty, so blank
        // rows were silently dropped from the map entirely and could
        // never be found by a vectorized VLOOKUP/XLOOKUP/LOOKUPIFS call,
        // while the scalar path found them just fine.
        CellValue::Empty => Some(LookupKey::Empty),
    }
}

/// If a criteria is a plain equality check (no wildcard, no `>`/`<`/`<>`
/// comparison), return its hashable key. Used by the vectorized *IF/*IFS
/// functions: when EVERY criteria in a batch is a plain equality, we can
/// build one frequency map from the range and answer every criteria in
/// O(1), instead of re-scanning the range once per criteria value.
fn criteria_key(c: &Criteria) -> Option<LookupKey> {
    if c.wildcard.is_some() || !matches!(c.op, Op::Eq) {
        return None;
    }
    cell_to_key(&c.value)
}

// ============================================================
// Criteria parsing — handles Excel-style criteria strings such
// as ">10", "<=5", "<>0", "apple*", "*berry", or a plain value.
// ============================================================

#[derive(Clone, Copy, Debug)]
enum Op {
    Eq,
    Ne,
    Gt,
    Ge,
    Lt,
    Le,
}

struct Criteria {
    op: Op,
    value: CellValue,
    wildcard: Option<String>,
}

/// Parses an Excel-style criteria value/string into a `Criteria`. Returns
/// an error for a genuinely invalid combination — a wildcard ('*'/'?')
/// paired with a `>`/`<`/`>=`/`<=` comparison operator (e.g. ">appl*")
/// can never match anything (wildcards only make sense with equality/
/// inequality), so this is rejected up front rather than silently
/// building a criteria that always evaluates to zero matches.
fn parse_criteria(obj: &Bound<PyAny>) -> PyResult<Criteria> {
    let cv = CellValue::from_py(obj);
    match cv {
        CellValue::Num(n) => Ok(Criteria {
            op: Op::Eq,
            value: CellValue::Num(n),
            wildcard: None,
        }),
        CellValue::Text(s) => {
            let ops: [(&str, Op); 6] = [
                (">=", Op::Ge),
                ("<=", Op::Le),
                ("<>", Op::Ne),
                (">", Op::Gt),
                ("<", Op::Lt),
                ("=", Op::Eq),
            ];
            for (prefix, op) in ops.iter() {
                if let Some(rest) = s.strip_prefix(prefix) {
                    let rest = rest.trim();
                    if let Ok(n) = rest.parse::<f64>() {
                        return Ok(Criteria {
                            op: *op,
                            value: CellValue::Num(n),
                            wildcard: None,
                        });
                    }
                    if rest.contains('*') || rest.contains('?') {
                        if !matches!(op, Op::Eq | Op::Ne) {
                            return Err(PyValueError::new_err(format!(
                                "Invalid criteria {:?}: a wildcard ('*' or '?') can't be combined \
                                 with the '{}' comparison operator. Wildcards only work with \
                                 equality ('=') or not-equal ('<>') — e.g. \"={}\" or \"<>{}\".",
                                s, prefix, rest, rest
                            )));
                        }
                        return Ok(Criteria {
                            op: *op,
                            value: CellValue::Text(rest.to_string()),
                            wildcard: Some(rest.to_string()),
                        });
                    }
                    return Ok(Criteria {
                        op: *op,
                        value: CellValue::Text(rest.to_string()),
                        wildcard: None,
                    });
                }
            }
            if s.contains('*') || s.contains('?') {
                return Ok(Criteria {
                    op: Op::Eq,
                    value: CellValue::Text(s.clone()),
                    wildcard: Some(s),
                });
            }
            if let Ok(n) = s.trim().parse::<f64>() {
                return Ok(Criteria {
                    op: Op::Eq,
                    value: CellValue::Num(n),
                    wildcard: None,
                });
            }
            Ok(Criteria {
                op: Op::Eq,
                value: CellValue::Text(s),
                wildcard: None,
            })
        }
        CellValue::Empty => Ok(Criteria {
            op: Op::Eq,
            value: CellValue::Empty,
            wildcard: None,
        }),
    }
}

/// Excel-style wildcard match: '*' = any run of characters, '?' = exactly one.
fn wildcard_match(pattern: &str, text: &str) -> bool {
    let p: Vec<char> = pattern.to_lowercase().chars().collect();
    let t: Vec<char> = text.to_lowercase().chars().collect();

    fn rec(p: &[char], t: &[char]) -> bool {
        if p.is_empty() {
            return t.is_empty();
        }
        match p[0] {
            '*' => {
                for i in 0..=t.len() {
                    if rec(&p[1..], &t[i..]) {
                        return true;
                    }
                }
                false
            }
            '?' => !t.is_empty() && rec(&p[1..], &t[1..]),
            c => !t.is_empty() && t[0] == c && rec(&p[1..], &t[1..]),
        }
    }
    rec(&p, &t)
}

fn matches(cell: &CellValue, criteria: &Criteria) -> bool {
    if let Some(pattern) = &criteria.wildcard {
        return match cell {
            CellValue::Text(t) => {
                let hit = wildcard_match(pattern, t);
                match criteria.op {
                    Op::Eq => hit,
                    Op::Ne => !hit,
                    _ => false,
                }
            }
            _ => false,
        };
    }
    match (&criteria.value, cell) {
        (CellValue::Num(cn), CellValue::Num(vn)) => match criteria.op {
            Op::Eq => (vn - cn).abs() < 1e-9,
            Op::Ne => (vn - cn).abs() >= 1e-9,
            Op::Gt => vn > cn,
            Op::Ge => vn >= cn,
            Op::Lt => vn < cn,
            Op::Le => vn <= cn,
        },
        (CellValue::Text(ct), CellValue::Text(vt)) => match criteria.op {
            // Eq/Ne is overwhelmingly the common case (category/department/
            // status-style criteria) — text_eq_ci avoids allocating two
            // lowercase String copies per row for it.
            Op::Eq => text_eq_ci(vt, ct),
            Op::Ne => !text_eq_ci(vt, ct),
            // Ordering needs an actual lowercase copy to compare correctly.
            Op::Gt | Op::Ge | Op::Lt | Op::Le => {
                let a = vt.to_lowercase();
                let b = ct.to_lowercase();
                match criteria.op {
                    Op::Gt => a > b,
                    Op::Ge => a >= b,
                    Op::Lt => a < b,
                    Op::Le => a <= b,
                    _ => unreachable!(),
                }
            }
        },
        (CellValue::Empty, CellValue::Empty) => matches!(criteria.op, Op::Eq),
        (CellValue::Empty, _) => matches!(criteria.op, Op::Ne),
        _ => matches!(criteria.op, Op::Ne),
    }
}

// ============================================================
// FastColumn — the "mixed" fast path for SUM/AVERAGE/COUNT/*IF/*IFS.
//
// Lesson learned the hard way (see HANDOFF.md §5): converting a plain
// Python list to a numpy array first is pure overhead — building that
// array from scratch costs about as much as just processing the list
// directly, so it's not a real win. And naively falling back to
// Vec<PyObject> for an ENTIRE SUMIFS call just because ONE of several
// criteria columns is text needlessly re-boxes every OTHER column too,
// even ones that were already a clean numeric buffer.
//
// FastColumn fixes both: it's resolved ONCE per column (not once per
// element), and each column independently gets the zero-copy numeric
// path when it's *already* array-backed numeric data (numpy array, or
// a numeric-dtype pandas/polars Series/column) — completely
// independent of what any other column in the same call turns out to
// be. A column that isn't already a clean numeric buffer falls back to
// the existing, fully-generic Vec<PyObject> + CellValue handling,
// which is always available and always correct.
// ============================================================

enum FastColumn<'py> {
    Numeric(numpy::PyReadonlyArray1<'py, f64>),
    // True zero-copy for int64 data — no cast at all, not even a Rust-side
    // one. int64 is the overwhelmingly common numpy/pandas dtype for
    // whole-number columns (IDs, counts, salaries, quantities), and
    // casting an entire array to f64 first (even a cheap vectorized
    // cast) is exactly the kind of "time spent on datatype conversion"
    // this fast path exists to eliminate — so int64 gets its own direct
    // path instead of being funneled through a float64 conversion.
    NumericI64(numpy::PyReadonlyArray1<'py, i64>),
    // Zero-per-row-GIL-call path for a text/string column (a polars Utf8
    // Series, a pandas string/object Series that came through as a numpy
    // object array of Python str, or a plain Python list of str).
    //
    // The key win here isn't "no PyObject involved" (str objects are
    // still PyObjects at the Python boundary) — it's that we pay the
    // Python->Rust string conversion cost ONCE, in a single bulk
    // `extract::<Vec<String>>()` call, instead of once per row inside
    // the hot matching loop. PyO3's bulk sequence extractor is a tight
    // native loop over the underlying buffer; doing the same conversion
    // one element at a time via `.bind(py)` + `downcast::<PyString>()`
    // inside `matches_at` (the old fallback, still used for anything
    // that isn't cleanly string-like) pays per-call overhead N times
    // instead of once. Benchmarked on 2M-row text columns: this took
    // COUNTIF/SUMIF from ~300-400ms down to single-digit ms, matching
    // the numeric fast path's speedup — see HANDOFF.md §7.
    Text(Vec<String>),
    // Native Arrow C Stream reader: parses Utf8View/Utf8/LargeUtf8
    // directly from raw producer buffers via `raw_arrow`, bypassing
    // BOTH the per-row PyObject path (like ArrowText below) AND the
    // `to_arrow(compat_level=oldest())` re-encoding step that ArrowText
    // requires for a native-StringView producer (polars' default) —
    // measured at 5-50ms on its own for a 1-2M row column, on top of
    // whatever the actual comparison then costs. This is the fastest
    // available path for a `__arrow_c_stream__`-exporting object.
    // `Option<String>` per row supports nulls properly (unlike `Text`
    // above, which is only ever fed already-known-clean data).
    NativeText(Vec<Option<String>>),
    // True zero-copy text: no PyObject touched per row at all. Populated
    // via the Arrow C Stream Interface (see `try_arrow_text_column`)
    // directly from a polars Series (or any other `__arrow_c_stream__`
    // exporter) — this is the actual polars-speed path; `Text(Vec<String>)`
    // above remains as the correct, still-much-improved fallback for a
    // plain Python `list[str]`, which has no Arrow buffer to borrow from.
    ArrowText(ArrowTextColumn),
    Generic(Vec<PyObject>),
}

/// Count rows where every mask in `masks` is `true` at that row —
/// specialized for the overwhelmingly common 1- and 2-predicate
/// COUNTIFS cases (a tight zip+AND loop, no inner iterator-of-
/// iterators per row), falling back to the generic `all()` scan for
/// 3+ predicates. Parallelized via rayon above the same size
/// threshold used elsewhere in this file — safe here because every
/// mask is already a plain `Vec<bool>` with no PyObject/GIL
/// involvement at all by this point.
///
/// Added because the original `for i in 0..n { masks.iter().all(...) }`
/// loop measured a real, separate cost on top of the mask-building
/// itself (~6ms of a ~25ms 2-predicate COUNTIFS at 1M rows) — small
/// next to the mask-building cost, but a genuine, avoidable gap, not
/// noise.
///
/// NOTE on summation strategy: a recursive pairwise (divide-and-
/// conquer) summation was tried here for SUM/AVERAGE's numerator, on
/// the theory that it would close the gap to numpy's own pairwise-
/// summation-based `.sum()` (measured: numpy ~0.3-0.5ms vs this
/// file's simple serial `.sum()` ~1.2-1.3ms for 1M f64 elements,
/// STEADY-STATE i.e. after several warmup calls). It was reverted
/// after discovering a worse problem specific to real single-call
/// usage: the recursive version's FIRST call in a fresh process
/// measured ~4.7-6.5ms — slower than the simple serial version it was
/// meant to replace — settling to the hoped-for ~1.6-2ms only after
/// 3-5 repeated calls in the SAME process. A standalone (non-PyO3)
/// Rust binary running the identical recursive function showed NO
/// such first-call penalty (~940µs immediately), while another
/// existing Rust function in this same compiled extension
/// (`countif_mixed`) also showed no such penalty — narrowing the
/// cause to something specific to how THIS function's many small
/// recursive call sites get warmed up the first time they execute
/// inside a loaded PyO3 extension module, not a generic "first call
/// into the shared library" effect. Since this library's real usage
/// (an Excel-formula-style tool) is single-call-dominant, not a tight
/// loop calling the same function thousands of times, the warm-loop
/// benchmark that made the pairwise version look like a win was
/// measuring the wrong thing — it hid the realistic first-call cost
/// a user would actually pay. Left as plain serial summation as a
/// result: slower than numpy in a repeated-call microbenchmark, but
/// faster and more consistent for the realistic single-call case.
/// Cheap, heuristic "does this look like a date string" check —
/// deliberately matching the Python `_looks_like_date` implementation
/// this ports (same length bounds, same `-`/`/` split, same "exactly
/// 3 numeric parts, at least one 4 digits long" rule) so INFO's
/// `date_looking` flag reports the same thing whether a column takes
/// this Rust fast path or the pure-Python fallback. This is a
/// deliberately loose heuristic (advisory, not a real date parser) —
/// it exists to flag "you should probably look at converting this
/// column", not to validate real dates.
fn looks_like_date(s: &str) -> bool {
    if s.len() < 6 || s.len() > 10 {
        return false;
    }
    let parts: Vec<&str> = s.split(|c| c == '-' || c == '/').collect();
    if parts.len() != 3 {
        return false;
    }
    if !parts
        .iter()
        .all(|p| !p.is_empty() && p.len() <= 4 && p.chars().all(|c| c.is_ascii_digit()))
    {
        return false;
    }
    parts.iter().any(|p| p.len() == 4)
}

fn count_all_masks_true(masks: &[Vec<bool>], n: usize) -> i64 {
    const PARALLEL_THRESHOLD: usize = 50_000;
    match masks {
        [] => 0,
        [only] => {
            if n >= PARALLEL_THRESHOLD {
                only.par_iter().filter(|&&b| b).count() as i64
            } else {
                only.iter().filter(|&&b| b).count() as i64
            }
        }
        [a, b] => {
            if n >= PARALLEL_THRESHOLD {
                (0..n).into_par_iter().filter(|&i| a[i] && b[i]).count() as i64
            } else {
                (0..n).filter(|&i| a[i] && b[i]).count() as i64
            }
        }
        _ => {
            if n >= PARALLEL_THRESHOLD {
                (0..n)
                    .into_par_iter()
                    .filter(|&i| masks.iter().all(|m| m[i]))
                    .count() as i64
            } else {
                (0..n).filter(|&i| masks.iter().all(|m| m[i])).count() as i64
            }
        }
    }
}

/// Sum of `values[i]` (skipping NaN, matching `numeric_at`'s existing
/// "NaN is blank" convention) over rows where every mask in `masks` is
/// `true` — the `as_f64_slice`-backed counterpart to
/// `count_all_masks_true` above, used by SUMIFS when `sum_range`
/// resolved to a genuinely zero-copy numeric column. No PyObject/GIL
/// touched anywhere in this function, so it's safe to parallelize the
/// same way.
fn sum_where_all_masks_true(values: &[f64], masks: &[Vec<bool>], n: usize) -> f64 {
    const PARALLEL_THRESHOLD: usize = 50_000;
    #[inline(always)]
    fn add_if_valid(total: f64, v: f64) -> f64 {
        if v.is_nan() {
            total
        } else {
            total + v
        }
    }
    match masks {
        [] => 0.0,
        [only] => {
            if n >= PARALLEL_THRESHOLD {
                (0..n)
                    .into_par_iter()
                    .filter(|&i| only[i])
                    .map(|i| values[i])
                    .filter(|v| !v.is_nan())
                    .sum()
            } else {
                (0..n)
                    .filter(|&i| only[i])
                    .fold(0.0, |t, i| add_if_valid(t, values[i]))
            }
        }
        [a, b] => {
            if n >= PARALLEL_THRESHOLD {
                (0..n)
                    .into_par_iter()
                    .filter(|&i| a[i] && b[i])
                    .map(|i| values[i])
                    .filter(|v| !v.is_nan())
                    .sum()
            } else {
                (0..n)
                    .filter(|&i| a[i] && b[i])
                    .fold(0.0, |t, i| add_if_valid(t, values[i]))
            }
        }
        _ => {
            if n >= PARALLEL_THRESHOLD {
                (0..n)
                    .into_par_iter()
                    .filter(|&i| masks.iter().all(|m| m[i]))
                    .map(|i| values[i])
                    .filter(|v| !v.is_nan())
                    .sum()
            } else {
                (0..n)
                    .filter(|&i| masks.iter().all(|m| m[i]))
                    .fold(0.0, |t, i| add_if_valid(t, values[i]))
            }
        }
    }
}

/// The `(sum, count)` counterpart to `sum_where_all_masks_true` above
/// — used by AVERAGEIFS, which needs both the total and the count of
/// matched-and-numeric rows (SUMIFS only needs the total; COUNTIFS
/// only needs the count of matched rows regardless of numeric-ness —
/// neither of those existing helpers can be reused directly for
/// AVERAGEIFS's specific combination). Same mask-AND, same NaN-is-
/// blank convention (a NaN value is skipped from both the sum and the
/// count, matching `numeric_at`'s existing semantics), same
/// parallel-above-threshold structure as every other mask-driven
/// aggregation in this file.
fn sum_and_count_where_all_masks_true(values: &[f64], masks: &[Vec<bool>], n: usize) -> (f64, u64) {
    const PARALLEL_THRESHOLD: usize = 50_000;
    #[inline(always)]
    fn add_if_valid(acc: (f64, u64), v: f64) -> (f64, u64) {
        if v.is_nan() {
            acc
        } else {
            (acc.0 + v, acc.1 + 1)
        }
    }
    #[inline(always)]
    fn combine(a: (f64, u64), b: (f64, u64)) -> (f64, u64) {
        (a.0 + b.0, a.1 + b.1)
    }
    match masks {
        [] => (0.0, 0),
        [only] => {
            if n >= PARALLEL_THRESHOLD {
                (0..n)
                    .into_par_iter()
                    .filter(|&i| only[i])
                    .map(|i| values[i])
                    .filter(|v| !v.is_nan())
                    .map(|v| (v, 1u64))
                    .reduce(|| (0.0, 0), combine)
            } else {
                (0..n)
                    .filter(|&i| only[i])
                    .fold((0.0, 0), |acc, i| add_if_valid(acc, values[i]))
            }
        }
        [a, b] => {
            if n >= PARALLEL_THRESHOLD {
                (0..n)
                    .into_par_iter()
                    .filter(|&i| a[i] && b[i])
                    .map(|i| values[i])
                    .filter(|v| !v.is_nan())
                    .map(|v| (v, 1u64))
                    .reduce(|| (0.0, 0), combine)
            } else {
                (0..n)
                    .filter(|&i| a[i] && b[i])
                    .fold((0.0, 0), |acc, i| add_if_valid(acc, values[i]))
            }
        }
        _ => {
            if n >= PARALLEL_THRESHOLD {
                (0..n)
                    .into_par_iter()
                    .filter(|&i| masks.iter().all(|m| m[i]))
                    .map(|i| values[i])
                    .filter(|v| !v.is_nan())
                    .map(|v| (v, 1u64))
                    .reduce(|| (0.0, 0), combine)
            } else {
                (0..n)
                    .filter(|&i| masks.iter().all(|m| m[i]))
                    .fold((0.0, 0), |acc, i| add_if_valid(acc, values[i]))
            }
        }
    }
}

impl<'py> FastColumn<'py> {
    fn resolve(obj: &Bound<'py, PyAny>) -> PyResult<FastColumn<'py>> {
        // Only succeeds for data that's ALREADY a numeric numpy-compatible
        // buffer (a real numpy array, or a pandas/polars numeric column,
        // which both expose the same buffer protocol numpy does) — never
        // attempted by building a fresh array from a plain Python list.
        if let Ok(arr) = obj.extract::<numpy::PyReadonlyArray1<f64>>() {
            return Ok(FastColumn::Numeric(arr));
        }
        if let Ok(arr) = obj.extract::<numpy::PyReadonlyArray1<i64>>() {
            return Ok(FastColumn::NumericI64(arr));
        }
        // NOTE: a native Utf8View stream reader (`raw_arrow::
        // try_read_stream_as_text`, further down in this file) exists
        // and is correctness-tested, but is intentionally NOT called
        // here. Direct A/B benchmarking showed it is SLOWER than the
        // `to_arrow(oldest)`-based `ArrowText` path below, not faster
        // — confirmed at 1M rows: ~90ms for the native-stream route vs
        // ~14-16ms for `to_arrow(oldest)` + `ArrowText`'s zero-copy
        // borrow. The reason: `try_read_stream_as_text` materializes
        // an owned `String` per row (a heap allocation per row), which
        // reintroduces exactly the per-row-allocation cost the whole
        // `ArrowText`/`eq_mask` design was built to eliminate — a
        // single bulk `to_arrow()` conversion followed by zero-copy
        // `&str` borrows is cheaper than N individual allocations, even
        // though the conversion step itself isn't free. Left in place
        // (rather than deleted) as a correct, available building block
        // for a future genuinely-zero-copy version of this same reader
        // — one that borrows `&str` directly from the stream's own
        // buffers instead of copying into owned `String`s — but that
        // requires solving the chunk-buffer-lifetime problem (keeping
        // each chunk's buffers alive exactly as long as needed) that
        // was deliberately avoided here in favor of the simpler,
        // memory-safety-first owned-copy design. Do not re-enable this
        // call without re-benchmarking; the measured numbers, not
        // intuition, are what should decide it.
        // True zero-copy text path: try the Arrow C Stream Interface
        // first (polars Series, pyarrow arrays, etc.) — this is the
        // path that actually matches polars' own internal speed, since
        // it never touches a per-row PyObject. Only falls through (not
        // an error) when the object doesn't export Arrow at all, or
        // its column isn't string-typed.
        if let Some(col) = try_arrow_text_column(obj)? {
            return Ok(FastColumn::ArrowText(col));
        }
        // Bulk string extraction. The Python layer only ever hands this
        // a plain `list[str]` for the Text path (a polars `.to_list()`
        // or a clean-dtype pandas `.to_list()` result), so downcasting
        // straight to PyList and walking it with `PyString` downcasts
        // avoids PyO3's generic `FromPyObject for Vec<T>` machinery,
        // which goes through the sequence protocol + a `PyResult` per
        // element rather than a direct concrete-type downcast. This
        // matters at 2M rows: the generic `extract::<Vec<String>>()`
        // cost ~180-200ms end to end (measured), most of it in that
        // extraction, not the matching loop itself (a pure-Rust loop
        // over the same 2M strings is ~16ms). The direct-downcast walk
        // below is the fix for that gap; if the object isn't a PyList
        // (or any element isn't a PyString), this falls through to the
        // fully-generic Vec<PyObject> path, same as today.
        if let Ok(list) = obj.downcast::<pyo3::types::PyList>() {
            let len = list.len();
            let mut v: Vec<String> = Vec::with_capacity(len);
            let mut all_str = true;
            for i in 0..len {
                let item = match list.get_item(i) {
                    Ok(it) => it,
                    Err(_) => {
                        all_str = false;
                        break;
                    }
                };
                match item.downcast::<PyString>() {
                    Ok(s) => v.push(s.to_string_lossy().into_owned()),
                    Err(_) => {
                        all_str = false;
                        break;
                    }
                }
            }
            if all_str {
                return Ok(FastColumn::Text(v));
            }
        }
        let v: Vec<PyObject> = obj.extract()?;
        Ok(FastColumn::Generic(v))
    }

    fn len(&self) -> usize {
        match self {
            FastColumn::Numeric(a) => a.as_slice().map(|s| s.len()).unwrap_or(0),
            FastColumn::NumericI64(a) => a.as_slice().map(|s| s.len()).unwrap_or(0),
            FastColumn::Text(v) => v.len(),
            FastColumn::NativeText(v) => v.len(),
            FastColumn::ArrowText(c) => c.len(),
            FastColumn::Generic(v) => v.len(),
        }
    }

    /// Numeric value at row `i`, or None if that cell isn't numeric
    /// (only possible for the Generic variant — a text/blank cell).
    fn numeric_at(&self, py: Python<'_>, i: usize) -> Option<f64> {
        match self {
            // NaN must be treated as blank (None), same as
            // CellValue::from_py — otherwise a float64 buffer with a
            // NaN (e.g. missing data in a pandas numeric column) would
            // silently poison SUM/AVERAGE, reintroducing the exact bug
            // that was fixed in the generic path.
            FastColumn::Numeric(a) => a.as_slice().ok().and_then(|s| {
                let v = s[i];
                if v.is_nan() {
                    None
                } else {
                    Some(v)
                }
            }),
            FastColumn::NumericI64(a) => a.as_slice().ok().map(|s| s[i] as f64),
            // A text column is never numeric — same semantics as
            // CellValue::Text in the generic path (SUM/AVERAGE ignore it).
            FastColumn::Text(_) => None,
            FastColumn::NativeText(_) => None,
            FastColumn::ArrowText(_) => None,
            FastColumn::Generic(v) => match CellValue::from_py(v[i].bind(py)) {
                CellValue::Num(n) => Some(n),
                _ => None,
            },
        }
    }

    /// The `LookupKey` for row `i` — the single entry point VLOOKUP/
    /// XLOOKUP/LOOKUPIFS's fast path uses to build its HashMap, working
    /// directly against whichever zero-copy `FastColumn` variant the
    /// key column resolved to. This is what lets those functions build
    /// their lookup table from the SAME Arrow/numpy buffers COUNTIF/
    /// SUMIFS already read zero-copy — no per-row PyObject touch for a
    /// `Numeric`/`NumericI64`/`ArrowText`/`NativeText` column, matching
    /// `cell_to_key(&CellValue::from_py(...))`'s semantics exactly for
    /// every variant (case-INsensitive text keys via `.to_lowercase()`,
    /// same as the existing generic path — Excel's own VLOOKUP is
    /// case-insensitive) so a fast-path lookup can never disagree with
    /// the slow-path one on which rows match.
    fn cell_key_at(&self, py: Python<'_>, i: usize) -> Option<LookupKey> {
        match self {
            FastColumn::Numeric(a) => a.as_slice().ok().map(|s| {
                let v = s[i];
                if v.is_nan() {
                    LookupKey::Empty
                } else {
                    LookupKey::Num(v)
                }
            }),
            FastColumn::NumericI64(a) => a.as_slice().ok().map(|s| LookupKey::Num(s[i] as f64)),
            FastColumn::Text(v) => Some(LookupKey::Text(v[i].to_lowercase())),
            FastColumn::ArrowText(c) => Some(match c.get(i) {
                Some(s) => LookupKey::Text(s.to_lowercase()),
                None => LookupKey::Empty,
            }),
            FastColumn::NativeText(v) => Some(match &v[i] {
                Some(s) => LookupKey::Text(s.to_lowercase()),
                None => LookupKey::Empty,
            }),
            FastColumn::Generic(v) => cell_to_key(&CellValue::from_py(v[i].bind(py))),
        }
    }

    /// Row `i`'s text value, prepared for a case-insensitive HashMap
    /// probe with as little allocation as possible — the allocation-
    /// avoiding counterpart to `cell_key_at` used specifically by
    /// `vlookup_many_columnar`'s split text/other map design (see
    /// that function's own comment for the full rationale and the
    /// benchmark numbers behind it). Returns `None` for a null cell or
    /// a non-text column/variant — the caller falls back to
    /// `cell_key_at`'s `LookupKey`-based path for those, unchanged.
    ///
    /// Returns `CellTextRef::Borrowed(&str)` ONLY when the text is
    /// confirmed already fully lowercase ASCII — safe to use directly
    /// as a case-insensitive key with zero allocation, since this
    /// file's case-insensitive matching (`text_eq_ci` and this
    /// method) always normalizes to lowercase, so an already-
    /// lowercase string IS its own normalized form. Anything else
    /// (contains an uppercase ASCII letter, or isn't ASCII at all —
    /// Unicode case-folding needs the full `.to_lowercase()` for
    /// correctness) falls back to `CellTextRef::Owned`, paying exactly
    /// the one allocation the old code always paid — never worse,
    /// only better for the common case.
    fn cell_text_ref_at(&self, py: Python<'_>, i: usize) -> Option<CellTextRef<'_>> {
        #[inline(always)]
        fn prepare(s: &str) -> CellTextRef<'_> {
            if s.bytes().all(|b| !b.is_ascii_uppercase()) && s.is_ascii() {
                CellTextRef::Borrowed(s)
            } else {
                CellTextRef::Owned(s.to_lowercase())
            }
        }
        match self {
            FastColumn::Text(v) => Some(prepare(v[i].as_str())),
            FastColumn::ArrowText(c) => c.get(i).map(prepare),
            FastColumn::NativeText(v) => v[i].as_deref().map(prepare),
            FastColumn::Numeric(_) | FastColumn::NumericI64(_) | FastColumn::Generic(_) => None,
        }
    }

    /// Row `i`'s text value with its ORIGINAL casing preserved — the
    /// case-preserving counterpart to `cell_text_ref_at` above (which
    /// deliberately lowercases everything for VLOOKUP's case-
    /// insensitive key matching). INFO/CLEAN's column summaries need
    /// the real, as-stored values — "IT" and "it" are two distinct
    /// categories to count and report separately, not one merged key
    /// — so this exists as a separate method rather than adding a
    /// case-preserving flag to the VLOOKUP-oriented one and risking a
    /// caller picking the wrong mode by mistake.
    fn cell_text_exact_at(&self, _py: Python<'_>, i: usize) -> Option<&str> {
        match self {
            FastColumn::Text(v) => Some(v[i].as_str()),
            FastColumn::ArrowText(c) => c.get(i),
            FastColumn::NativeText(v) => v[i].as_deref(),
            FastColumn::Numeric(_) | FastColumn::NumericI64(_) | FastColumn::Generic(_) => None,
        }
    }

    /// Row `i`'s value as a Python object — used for the RETURN column
    /// of a fast-path VLOOKUP/XLOOKUP/LOOKUPIFS (as opposed to
    /// `cell_key_at`, used for the LOOKUP/key column). Needed because
    /// the return column can be any dtype (numbers, text, whatever the
    /// caller asked to look up), and the eventual Python-level result
    /// must be a real Python object either way, not a Rust-native type
    /// — for a `Numeric`/`NumericI64`/`ArrowText`/`NativeText` column,
    /// this is the one place a PyObject actually gets constructed per
    /// matched row (there's no way around that: the function's contract
    /// is to hand Python back a value), but it's paid only once per
    /// MATCHED row, never once per row scanned — a real win whenever
    /// most rows aren't matches, which is the common VLOOKUP shape (a
    /// small `lookup_value` list against a much larger table).
    fn to_object_at(&self, py: Python<'_>, i: usize) -> PyObject {
        match self {
            FastColumn::Numeric(a) => match a.as_slice() {
                Ok(s) => {
                    let v = s[i];
                    if v.is_nan() {
                        py.None()
                    } else {
                        v.into_py(py)
                    }
                }
                Err(_) => py.None(),
            },
            FastColumn::NumericI64(a) => match a.as_slice() {
                Ok(s) => s[i].into_py(py),
                Err(_) => py.None(),
            },
            FastColumn::Text(v) => v[i].clone().into_py(py),
            FastColumn::ArrowText(c) => match c.get(i) {
                Some(s) => s.into_py(py),
                None => py.None(),
            },
            FastColumn::NativeText(v) => match &v[i] {
                Some(s) => s.clone().into_py(py),
                None => py.None(),
            },
            FastColumn::Generic(v) => v[i].clone_ref(py),
        }
    }

    /// Does row `i` satisfy `criteria`? A Numeric column with a numeric
    /// criteria compares directly against the raw f64 (no CellValue,
    /// no allocation at all). A Generic column takes a zero-allocation
    /// borrowed-&str fast path for the common case (plain text equality/
    /// Fast-path mask builder: if `self` is a text column (`ArrowText`
    /// or `NativeText`) and `crit` is a simple (no-wildcard) Eq/Ne text
    /// criteria, returns the whole column's match mask via a tight,
    /// per-row-dispatch-free loop (see `ArrowTextColumn::eq_mask`'s doc
    /// comment for why this closes most of the remaining gap to
    /// polars' own comparison kernel speed). Returns `None` for every
    /// other case — the caller falls back to the generic per-row
    /// `matches_at` loop, which remains correct for all of those.
    fn try_fast_eq_mask(&self, crit: &Criteria) -> Option<Vec<bool>> {
        if crit.wildcard.is_some() {
            return None;
        }
        let negate = match crit.op {
            Op::Eq => false,
            Op::Ne => true,
            _ => return None,
        };
        let needle = match &crit.value {
            CellValue::Text(s) => s,
            _ => return None,
        };
        match self {
            FastColumn::ArrowText(c) => Some(c.eq_mask(needle, negate)),
            FastColumn::NativeText(v) => Some(native_text_eq_mask(v, needle, negate)),
            _ => None,
        }
    }

    /// Fast-path mask builder for a NUMERIC column (`Numeric`/
    /// `NumericI64` — the zero-copy `f64`/`i64` numpy buffer variants)
    /// against a numeric criteria, covering ALL SIX operators (Eq, Ne,
    /// Gt, Ge, Lt, Le) — not just Eq/Ne like the text fast path above,
    /// since numeric criteria are overwhelmingly inequality-based
    /// (">5000", "<=100") rather than equality-based in real usage.
    ///
    /// Added because numeric criteria had NO fast-mask path at all
    /// before this — every numeric COUNTIF/SUMIF/COUNTIFS/SUMIFS
    /// criteria check went through `matches_at`'s fully generic route,
    /// which (even for a zero-copy `Numeric`/`NumericI64` column)
    /// still re-derives a fresh `CellValue`/`Criteria` comparison via
    /// the `matches()` function on every single row. Confirmed by
    /// direct benchmark: `COUNTIF(sales, ">5000")` on a 1M-row pandas
    /// float64 column cost ~4x pandas' own `(sales > 5000).sum()` —
    /// this closes that gap the same way `eq_mask` closed it for text.
    ///
    /// `NaN` cells are treated as blank (`Op::Ne` matches, everything
    /// else doesn't) — identical to `numeric_at`'s existing "NaN is
    /// None/blank" convention and to `matches()`'s
    /// `(CellValue::Empty, _) => matches!(criteria.op, Op::Ne)` branch,
    /// so this fast path can never disagree with the generic one on a
    /// NaN cell.
    fn try_fast_numeric_mask(&self, crit: &Criteria) -> Option<Vec<bool>> {
        if crit.wildcard.is_some() {
            return None;
        }
        let needle = match crit.value {
            CellValue::Num(n) => n,
            _ => return None,
        };
        let op = crit.op;

        // Single comparison function, no per-element NaN branch and no
        // pre-scan pass over the data at all — relies entirely on
        // IEEE 754 float semantics, which already do the right thing
        // for a NaN value on FIVE of the six operators for free:
        // `NaN > x`, `NaN >= x`, `NaN < x`, `NaN <= x`, and (since Eq
        // is implemented as `(v-needle).abs() < 1e-9`) `NaN == x` all
        // naturally evaluate to `false` in hardware — which is exactly
        // "a blank cell doesn't satisfy this criteria", the correct
        // Excel-matching answer, with zero extra code.
        //
        // The one exception is `Ne`: naively writing it as
        // `(v-needle).abs() >= 1e-9` would ALSO evaluate to `false` for
        // NaN (since `NaN >= x` is always false), but the correct
        // answer for Ne is `true` (a blank cell IS "not equal" to any
        // specific criteria value — matches `matches()`'s own
        // `(CellValue::Empty, _) => matches!(criteria.op, Op::Ne)`
        // branch elsewhere in this file). Writing Ne as the outright
        // negation of the Eq test — `!((v-needle).abs() < 1e-9)` —
        // gets this right for NaN too: `!false = true`, since Eq's own
        // NaN-false-by-hardware behavior negates correctly. Verified
        // directly against Rust's actual float semantics before
        // relying on it (not assumed) — see HANDOFF notes.
        //
        // This closes the remaining gap to numpy's own raw comparison
        // that the previous two-pass (has-nan scan, then branch)
        // version left on the table: that version did TWO full passes
        // over the data for the common no-NaN case (one to check for
        // NaN, one to compare) — this version does exactly one, same
        // as numpy's own single-pass vectorized comparison.
        #[inline(always)]
        fn cell_matches(v: f64, needle: f64, op: Op) -> bool {
            match op {
                Op::Eq => (v - needle).abs() < 1e-9,
                Op::Ne => !((v - needle).abs() < 1e-9),
                Op::Gt => v > needle,
                Op::Ge => v >= needle,
                Op::Lt => v < needle,
                Op::Le => v <= needle,
            }
        }

        const PARALLEL_THRESHOLD: usize = 50_000;
        match self {
            FastColumn::Numeric(a) => {
                let slice = a.as_slice().ok()?;
                let n = slice.len();
                let mask: Vec<bool> = if n >= PARALLEL_THRESHOLD {
                    slice
                        .par_iter()
                        .map(|&v| cell_matches(v, needle, op))
                        .collect()
                } else {
                    slice.iter().map(|&v| cell_matches(v, needle, op)).collect()
                };
                Some(mask)
            }
            FastColumn::NumericI64(a) => {
                // An i64 value is never NaN, so the same single
                // comparison function is correct here unconditionally
                // — no special-casing needed for this variant either.
                let slice = a.as_slice().ok()?;
                let n = slice.len();
                let mask: Vec<bool> = if n >= PARALLEL_THRESHOLD {
                    slice
                        .par_iter()
                        .map(|&v| cell_matches(v as f64, needle, op))
                        .collect()
                } else {
                    slice
                        .iter()
                        .map(|&v| cell_matches(v as f64, needle, op))
                        .collect()
                };
                Some(mask)
            }
            _ => None,
        }
    }

    /// Builds a match mask for this column against `crit`, using the
    /// fast path from `try_fast_eq_mask` when it applies and falling
    /// back to a per-row `matches_at` loop otherwise. This is the
    /// single entry point every `*IF`/`*IFS` Rust function below goes
    /// through, so the fast path automatically benefits all of them
    /// (COUNTIF, SUMIF, AVERAGEIF, COUNTIFS, SUMIFS) rather than being
    /// wired into each one separately — mirroring how polars itself
    /// builds one boolean mask per predicate and combines/aggregates
    /// from there.
    fn build_mask(&self, py: Python<'_>, crit: &Criteria, n: usize) -> Vec<bool> {
        if let Some(mask) = self.try_fast_eq_mask(crit) {
            return mask;
        }
        if let Some(mask) = self.try_fast_numeric_mask(crit) {
            return mask;
        }
        let mut mask = Vec::with_capacity(n);
        for i in 0..n {
            mask.push(self.matches_at(py, i, crit));
        }
        mask
    }

    /// If `self` is a genuinely zero-copy numeric column (`Numeric` or
    /// `NumericI64`), returns a borrowed `&[f64]` view over it with NO
    /// PyO3/GIL touch per element — `NumericI64`'s `i64` values are
    /// widened to `f64` once, up front, into an owned `Vec<f64>` (a
    /// single bulk pass, not per-row), so the caller always gets a
    /// plain `&[f64]` regardless of the original integer width. `None`
    /// for every other variant (`Text`/`ArrowText`/`NativeText`/
    /// `Generic`), which still need `numeric_at`'s per-row, `py`-bound
    /// path (`Generic` genuinely requires the GIL per element; the text
    /// variants are never numeric at all).
    ///
    /// This exists so a masked-sum loop (SUMIFS' hot path) can run
    /// WITHOUT holding `py` at all for the common case — sum_range is
    /// overwhelmingly a real numeric column in practice — which both
    /// removes real per-row dispatch cost and makes the loop safe to
    /// parallelize with rayon (the GIL-bound `numeric_at` path cannot
    /// be, since `Python<'_>` isn't `Send`).
    fn as_f64_slice(&self) -> Option<std::borrow::Cow<'_, [f64]>> {
        match self {
            FastColumn::Numeric(a) => a.as_slice().ok().map(std::borrow::Cow::Borrowed),
            FastColumn::NumericI64(a) => a
                .as_slice()
                .ok()
                .map(|s| std::borrow::Cow::Owned(s.iter().map(|&x| x as f64).collect())),
            _ => None,
        }
    }

    fn matches_at(&self, py: Python<'_>, i: usize, crit: &Criteria) -> bool {
        match self {
            FastColumn::Numeric(a) => {
                let v = match a.as_slice() {
                    Ok(s) => s[i],
                    Err(_) => return false,
                };
                Self::num_matches(v, crit)
            }
            FastColumn::NumericI64(a) => {
                let v = match a.as_slice() {
                    Ok(s) => s[i] as f64,
                    Err(_) => return false,
                };
                Self::num_matches(v, crit)
            }
            // Zero per-row Python touch at all: `v[i]` is already a plain
            // Rust `String` (extracted once, in bulk, back in `resolve`),
            // so this whole branch runs with no GIL-bound calls, no
            // `.bind(py)`, no `downcast`. Same fast Eq/Ne borrowed-&str
            // comparison as the Generic branch below, but paid zero times
            // instead of N times for the type check.
            FastColumn::Text(v) => {
                let text = v[i].as_str();
                if crit.wildcard.is_none() && matches!(crit.op, Op::Eq | Op::Ne) {
                    if let CellValue::Text(ref ct) = crit.value {
                        let eq = text_eq_ci(text, ct);
                        return if matches!(crit.op, Op::Eq) { eq } else { !eq };
                    }
                }
                matches(&CellValue::Text(text.to_string()), crit)
            }
            // Same shape as `Text` above, but with proper null support
            // (a genuine null-cell distinct from an empty string) since
            // this variant can legitimately hold nulls read straight
            // off the producer's own validity bitmap.
            FastColumn::NativeText(v) => match &v[i] {
                Some(text) => {
                    if crit.wildcard.is_none() && matches!(crit.op, Op::Eq | Op::Ne) {
                        if let CellValue::Text(ref ct) = crit.value {
                            let eq = text_eq_ci(text, ct);
                            return if matches!(crit.op, Op::Eq) { eq } else { !eq };
                        }
                    }
                    matches(&CellValue::Text(text.clone()), crit)
                }
                None => matches(&CellValue::Empty, crit),
            },
            // True zero-copy path: `get(i)` reads directly from the
            // Arrow buffer (no PyObject touched). A null Arrow cell
            // (Arrow's own null bitmap, not a sentinel) maps to
            // CellValue::Empty, matching the same "blank cell" semantics
            // used everywhere else in this file. A read error (should
            // only happen if the batch's column type check in
            // `try_arrow_text_column` somehow let through a mismatched
            // array) is treated as no-match rather than panicking —
            // matching_at has no way to propagate a PyResult here.
            FastColumn::ArrowText(c) => {
                let text = match c.get(i) {
                    Some(t) => t,
                    // Null Arrow cell -> CellValue::Empty, same "blank
                    // cell" semantics as everywhere else in this file
                    // (matches()'s own (CellValue::Empty, _) branch).
                    None => return matches(&CellValue::Empty, crit),
                };
                if crit.wildcard.is_none() && matches!(crit.op, Op::Eq | Op::Ne) {
                    if let CellValue::Text(ref ct) = crit.value {
                        let eq = text_eq_ci(text, ct);
                        return if matches!(crit.op, Op::Eq) { eq } else { !eq };
                    }
                }
                matches(&CellValue::Text(text.to_string()), crit)
            }
            FastColumn::Generic(v) => {
                let obj = v[i].bind(py);
                if crit.wildcard.is_none() && matches!(crit.op, Op::Eq | Op::Ne) {
                    if let CellValue::Text(ref ct) = crit.value {
                        if let Ok(s) = obj.downcast::<PyString>() {
                            // to_string_lossy() returns Cow<str> — Borrowed
                            // (zero allocation) for the common case of a
                            // normal, already-valid string; only allocates
                            // if the string actually needs lossy repair.
                            let text = s.to_string_lossy();
                            let eq = text_eq_ci(&text, ct);
                            return if matches!(crit.op, Op::Eq) { eq } else { !eq };
                        }
                    }
                }
                matches(&CellValue::from_py(obj), crit)
            }
        }
    }

    #[inline]
    fn num_matches(v: f64, crit: &Criteria) -> bool {
        // NaN is treated as blank everywhere else in this codebase
        // (CellValue::from_py maps NaN -> Empty) — mirror that here so
        // a NaN cell behaves identically whether it arrived via the
        // zero-copy numeric path or the generic CellValue path: it
        // matches only a Ne-style "not equal to X" criteria, same as
        // matches()'s (CellValue::Empty, _) branch.
        if v.is_nan() {
            return match crit.value {
                CellValue::Empty => matches!(crit.op, Op::Eq),
                _ => matches!(crit.op, Op::Ne),
            };
        }
        match crit.value {
            CellValue::Num(n) => match crit.op {
                Op::Eq => (v - n).abs() < 1e-9,
                Op::Ne => (v - n).abs() >= 1e-9,
                Op::Gt => v > n,
                Op::Ge => v >= n,
                Op::Lt => v < n,
                Op::Le => v <= n,
            },
            // A numeric column can never satisfy a text/wildcard
            // criteria (mirrors CellValue's existing semantics).
            _ => matches!(crit.op, Op::Ne),
        }
    }
}

// ============================================================
// Python-exposed module
// ============================================================

#[pymodule]
mod _core {
    use super::*;

    #[pyfunction]
    fn sum_values(py: Python<'_>, values: Vec<PyObject>) -> PyResult<f64> {
        let mut total = 0.0;
        for v in &values {
            if let CellValue::Num(n) = CellValue::from_py(v.bind(py)) {
                total += n;
            }
        }
        Ok(total)
    }

    #[pyfunction]
    fn min_values(py: Python<'_>, values: Vec<PyObject>) -> PyResult<f64> {
        let mut result: Option<f64> = None;
        for v in &values {
            if let CellValue::Num(n) = CellValue::from_py(v.bind(py)) {
                result = Some(match result {
                    Some(m) if m <= n => m,
                    _ => n,
                });
            }
        }
        result.ok_or_else(|| PyValueError::new_err("MIN: no numeric values found"))
    }

    #[pyfunction]
    fn max_values(py: Python<'_>, values: Vec<PyObject>) -> PyResult<f64> {
        let mut result: Option<f64> = None;
        for v in &values {
            if let CellValue::Num(n) = CellValue::from_py(v.bind(py)) {
                result = Some(match result {
                    Some(m) if m >= n => m,
                    _ => n,
                });
            }
        }
        result.ok_or_else(|| PyValueError::new_err("MAX: no numeric values found"))
    }

    #[pyfunction]
    fn average_values(py: Python<'_>, values: Vec<PyObject>) -> PyResult<f64> {
        let mut total = 0.0;
        let mut count = 0u64;
        for v in &values {
            if let CellValue::Num(n) = CellValue::from_py(v.bind(py)) {
                total += n;
                count += 1;
            }
        }
        if count == 0 {
            return Err(PyValueError::new_err(
                "AVERAGE: no numeric values found (division by zero)",
            ));
        }
        Ok(total / count as f64)
    }

    #[pyfunction]
    fn count_values(py: Python<'_>, values: Vec<PyObject>) -> PyResult<i64> {
        let mut c = 0i64;
        for v in &values {
            if let CellValue::Num(_) = CellValue::from_py(v.bind(py)) {
                c += 1;
            }
        }
        Ok(c)
    }

    #[pyfunction]
    fn countif_values(py: Python<'_>, range: Vec<PyObject>, criteria: PyObject) -> PyResult<i64> {
        let crit = parse_criteria(criteria.bind(py))?;
        let mut c = 0i64;
        for v in &range {
            if matches(&CellValue::from_py(v.bind(py)), &crit) {
                c += 1;
            }
        }
        Ok(c)
    }

    #[pyfunction]
    #[pyo3(signature = (range, criteria, sum_range=None))]
    fn sumif_values(
        py: Python<'_>,
        range: Vec<PyObject>,
        criteria: PyObject,
        sum_range: Option<Vec<PyObject>>,
    ) -> PyResult<f64> {
        let crit = parse_criteria(criteria.bind(py))?;
        let target = sum_range.unwrap_or_else(|| range.iter().map(|v| v.clone_ref(py)).collect());
        if target.len() != range.len() {
            return Err(PyValueError::new_err(
                "SUMIF: sum_range must be the same length as range",
            ));
        }
        let mut total = 0.0;
        for (i, v) in range.iter().enumerate() {
            if matches(&CellValue::from_py(v.bind(py)), &crit) {
                if let CellValue::Num(n) = CellValue::from_py(target[i].bind(py)) {
                    total += n;
                }
            }
        }
        Ok(total)
    }

    #[pyfunction]
    #[pyo3(signature = (range, criteria, average_range=None))]
    fn averageif_values(
        py: Python<'_>,
        range: Vec<PyObject>,
        criteria: PyObject,
        average_range: Option<Vec<PyObject>>,
    ) -> PyResult<f64> {
        let crit = parse_criteria(criteria.bind(py))?;
        let target =
            average_range.unwrap_or_else(|| range.iter().map(|v| v.clone_ref(py)).collect());
        if target.len() != range.len() {
            return Err(PyValueError::new_err(
                "AVERAGEIF: average_range must be the same length as range",
            ));
        }
        let mut total = 0.0;
        let mut count = 0u64;
        for (i, v) in range.iter().enumerate() {
            if matches(&CellValue::from_py(v.bind(py)), &crit) {
                if let CellValue::Num(n) = CellValue::from_py(target[i].bind(py)) {
                    total += n;
                    count += 1;
                }
            }
        }
        if count == 0 {
            return Err(PyValueError::new_err(
                "AVERAGEIF: no matching numeric values found",
            ));
        }
        Ok(total / count as f64)
    }

    #[pyfunction]
    fn countifs_values(py: Python<'_>, pairs: Vec<(Vec<PyObject>, PyObject)>) -> PyResult<i64> {
        if pairs.is_empty() {
            return Err(PyValueError::new_err(
                "COUNTIFS: at least one range/criteria pair is required",
            ));
        }
        let n = pairs[0].0.len();
        for (range, _) in &pairs {
            if range.len() != n {
                return Err(PyValueError::new_err(
                    "COUNTIFS: all ranges must be the same length",
                ));
            }
        }
        let parsed: Vec<(&Vec<PyObject>, Criteria)> = pairs
            .iter()
            .map(|(r, c)| Ok::<_, PyErr>((r, parse_criteria(c.bind(py))?)))
            .collect::<PyResult<Vec<_>>>()?;
        let mut count = 0i64;
        for i in 0..n {
            let mut ok = true;
            for (range, crit) in &parsed {
                if !matches(&CellValue::from_py(range[i].bind(py)), crit) {
                    ok = false;
                    break;
                }
            }
            if ok {
                count += 1;
            }
        }
        Ok(count)
    }

    #[pyfunction]
    fn sumifs_values(
        py: Python<'_>,
        sum_range: Vec<PyObject>,
        pairs: Vec<(Vec<PyObject>, PyObject)>,
    ) -> PyResult<f64> {
        if pairs.is_empty() {
            return Err(PyValueError::new_err(
                "SUMIFS: at least one range/criteria pair is required",
            ));
        }
        let n = sum_range.len();
        for (range, _) in &pairs {
            if range.len() != n {
                return Err(PyValueError::new_err(
                    "SUMIFS: all ranges must be the same length as sum_range",
                ));
            }
        }
        let parsed: Vec<(&Vec<PyObject>, Criteria)> = pairs
            .iter()
            .map(|(r, c)| Ok::<_, PyErr>((r, parse_criteria(c.bind(py))?)))
            .collect::<PyResult<Vec<_>>>()?;
        let mut total = 0.0;
        for i in 0..n {
            let mut ok = true;
            for (range, crit) in &parsed {
                if !matches(&CellValue::from_py(range[i].bind(py)), crit) {
                    ok = false;
                    break;
                }
            }
            if ok {
                if let CellValue::Num(v) = CellValue::from_py(sum_range[i].bind(py)) {
                    total += v;
                }
            }
        }
        Ok(total)
    }

    // ========================================================
    // VECTORIZED *IF / *IFS — evaluate MANY criteria in one call
    // (e.g. "for every department in table1, how many times does
    // it appear in table2"). Runs entirely in Rust: one call from
    // Python instead of a Python-side loop calling the scalar
    // version N times.
    //
    // Fast path: if every criteria in the batch is a plain equality
    // (no ">", "<", wildcard), we build ONE frequency/sum map from
    // the range and answer every criteria with an O(1) lookup —
    // instead of re-scanning the whole range once per criteria.
    // Falls back to a per-criteria scan (still all in Rust) the
    // moment any criteria needs a comparison or wildcard match.
    // ========================================================

    #[pyfunction]
    fn countif_vec_values(
        py: Python<'_>,
        range: Vec<PyObject>,
        criteria_list: Vec<PyObject>,
    ) -> PyResult<Vec<i64>> {
        let parsed: Vec<Criteria> = criteria_list
            .iter()
            .map(|c| parse_criteria(c.bind(py)))
            .collect::<PyResult<Vec<_>>>()?;
        let keys: Vec<Option<LookupKey>> = parsed.iter().map(criteria_key).collect();

        if keys.iter().all(|k| k.is_some()) {
            let mut freq: HashMap<LookupKey, i64> = HashMap::new();
            for v in &range {
                if let Some(k) = cell_to_key(&CellValue::from_py(v.bind(py))) {
                    *freq.entry(k).or_insert(0) += 1;
                }
            }
            return Ok(keys
                .into_iter()
                .map(|k| *freq.get(&k.unwrap()).unwrap_or(&0))
                .collect());
        }

        Ok(parsed
            .iter()
            .map(|crit| {
                range
                    .iter()
                    .filter(|v| matches(&CellValue::from_py(v.bind(py)), crit))
                    .count() as i64
            })
            .collect())
    }

    #[pyfunction]
    #[pyo3(signature = (range, criteria_list, sum_range=None))]
    fn sumif_vec_values(
        py: Python<'_>,
        range: Vec<PyObject>,
        criteria_list: Vec<PyObject>,
        sum_range: Option<Vec<PyObject>>,
    ) -> PyResult<Vec<f64>> {
        let target = sum_range.unwrap_or_else(|| range.iter().map(|v| v.clone_ref(py)).collect());
        if target.len() != range.len() {
            return Err(PyValueError::new_err(
                "SUMIF: sum_range must be the same length as range",
            ));
        }
        let parsed: Vec<Criteria> = criteria_list
            .iter()
            .map(|c| parse_criteria(c.bind(py)))
            .collect::<PyResult<Vec<_>>>()?;
        let keys: Vec<Option<LookupKey>> = parsed.iter().map(criteria_key).collect();

        if keys.iter().all(|k| k.is_some()) {
            let mut sums: HashMap<LookupKey, f64> = HashMap::new();
            for (v, s) in range.iter().zip(target.iter()) {
                if let Some(k) = cell_to_key(&CellValue::from_py(v.bind(py))) {
                    if let CellValue::Num(n) = CellValue::from_py(s.bind(py)) {
                        *sums.entry(k).or_insert(0.0) += n;
                    }
                }
            }
            return Ok(keys
                .into_iter()
                .map(|k| *sums.get(&k.unwrap()).unwrap_or(&0.0))
                .collect());
        }

        Ok(parsed
            .iter()
            .map(|crit| {
                range
                    .iter()
                    .zip(target.iter())
                    .filter(|(v, _)| matches(&CellValue::from_py(v.bind(py)), crit))
                    .map(|(_, s)| match CellValue::from_py(s.bind(py)) {
                        CellValue::Num(n) => n,
                        _ => 0.0,
                    })
                    .sum()
            })
            .collect())
    }

    #[pyfunction]
    #[pyo3(signature = (range, criteria_list, average_range=None))]
    fn averageif_vec_values(
        py: Python<'_>,
        range: Vec<PyObject>,
        criteria_list: Vec<PyObject>,
        average_range: Option<Vec<PyObject>>,
    ) -> PyResult<Vec<Option<f64>>> {
        let target =
            average_range.unwrap_or_else(|| range.iter().map(|v| v.clone_ref(py)).collect());
        if target.len() != range.len() {
            return Err(PyValueError::new_err(
                "AVERAGEIF: average_range must be the same length as range",
            ));
        }
        let parsed: Vec<Criteria> = criteria_list
            .iter()
            .map(|c| parse_criteria(c.bind(py)))
            .collect::<PyResult<Vec<_>>>()?;
        let keys: Vec<Option<LookupKey>> = parsed.iter().map(criteria_key).collect();

        if keys.iter().all(|k| k.is_some()) {
            let mut sums: HashMap<LookupKey, (f64, u64)> = HashMap::new();
            for (v, s) in range.iter().zip(target.iter()) {
                if let Some(k) = cell_to_key(&CellValue::from_py(v.bind(py))) {
                    if let CellValue::Num(n) = CellValue::from_py(s.bind(py)) {
                        let entry = sums.entry(k).or_insert((0.0, 0));
                        entry.0 += n;
                        entry.1 += 1;
                    }
                }
            }
            return Ok(keys
                .into_iter()
                .map(|k| match sums.get(&k.unwrap()) {
                    Some((total, count)) if *count > 0 => Some(total / *count as f64),
                    _ => None,
                })
                .collect());
        }

        Ok(parsed
            .iter()
            .map(|crit| {
                let mut total = 0.0;
                let mut count = 0u64;
                for (v, s) in range.iter().zip(target.iter()) {
                    if matches(&CellValue::from_py(v.bind(py)), crit) {
                        if let CellValue::Num(n) = CellValue::from_py(s.bind(py)) {
                            total += n;
                            count += 1;
                        }
                    }
                }
                if count == 0 {
                    None
                } else {
                    Some(total / count as f64)
                }
            })
            .collect())
    }

    #[pyfunction]
    fn countifs_vec_values(
        py: Python<'_>,
        pairs: Vec<(Vec<PyObject>, Vec<PyObject>)>,
    ) -> PyResult<Vec<i64>> {
        if pairs.is_empty() {
            return Err(PyValueError::new_err(
                "COUNTIFS: at least one range/criteria pair is required",
            ));
        }
        let range_len = pairs[0].0.len();
        let n = pairs[0].1.len();
        for (range, crit_list) in &pairs {
            if range.len() != range_len {
                return Err(PyValueError::new_err(
                    "COUNTIFS: all ranges must be the same length",
                ));
            }
            if crit_list.len() != n {
                return Err(PyValueError::new_err(
                    "COUNTIFS: all criteria columns must be the same length",
                ));
            }
        }

        let parsed_pairs: Vec<(&Vec<PyObject>, Vec<Criteria>)> = pairs
            .iter()
            .map(|(r, cl)| {
                let crits: PyResult<Vec<Criteria>> =
                    cl.iter().map(|c| parse_criteria(c.bind(py))).collect();
                Ok::<_, PyErr>((r, crits?))
            })
            .collect::<PyResult<Vec<_>>>()?;

        let all_eq = parsed_pairs
            .iter()
            .all(|(_, crits)| crits.iter().all(|c| criteria_key(c).is_some()));

        if all_eq {
            let mut freq: HashMap<Vec<LookupKey>, i64> = HashMap::new();
            for i in 0..range_len {
                let mut key = Vec::with_capacity(parsed_pairs.len());
                let mut ok = true;
                for (range, _) in &parsed_pairs {
                    match cell_to_key(&CellValue::from_py(range[i].bind(py))) {
                        Some(k) => key.push(k),
                        None => {
                            ok = false;
                            break;
                        }
                    }
                }
                if ok {
                    *freq.entry(key).or_insert(0) += 1;
                }
            }
            let mut out = Vec::with_capacity(n);
            for row in 0..n {
                let key: Vec<LookupKey> = parsed_pairs
                    .iter()
                    .map(|(_, crits)| criteria_key(&crits[row]).unwrap())
                    .collect();
                out.push(*freq.get(&key).unwrap_or(&0));
            }
            return Ok(out);
        }

        let mut out = Vec::with_capacity(n);
        for row in 0..n {
            let mut count = 0i64;
            for i in 0..range_len {
                let mut all_match = true;
                for (range, crits) in &parsed_pairs {
                    if !matches(&CellValue::from_py(range[i].bind(py)), &crits[row]) {
                        all_match = false;
                        break;
                    }
                }
                if all_match {
                    count += 1;
                }
            }
            out.push(count);
        }
        Ok(out)
    }

    #[pyfunction]
    fn sumifs_vec_values(
        py: Python<'_>,
        sum_range: Vec<PyObject>,
        pairs: Vec<(Vec<PyObject>, Vec<PyObject>)>,
    ) -> PyResult<Vec<f64>> {
        if pairs.is_empty() {
            return Err(PyValueError::new_err(
                "SUMIFS: at least one range/criteria pair is required",
            ));
        }
        let range_len = sum_range.len();
        let n = pairs[0].1.len();
        for (range, crit_list) in &pairs {
            if range.len() != range_len {
                return Err(PyValueError::new_err(
                    "SUMIFS: all ranges must be the same length as sum_range",
                ));
            }
            if crit_list.len() != n {
                return Err(PyValueError::new_err(
                    "SUMIFS: all criteria columns must be the same length",
                ));
            }
        }

        let parsed_pairs: Vec<(&Vec<PyObject>, Vec<Criteria>)> = pairs
            .iter()
            .map(|(r, cl)| {
                let crits: PyResult<Vec<Criteria>> =
                    cl.iter().map(|c| parse_criteria(c.bind(py))).collect();
                Ok::<_, PyErr>((r, crits?))
            })
            .collect::<PyResult<Vec<_>>>()?;

        let all_eq = parsed_pairs
            .iter()
            .all(|(_, crits)| crits.iter().all(|c| criteria_key(c).is_some()));

        if all_eq {
            let mut sums: HashMap<Vec<LookupKey>, f64> = HashMap::new();
            for i in 0..range_len {
                let mut key = Vec::with_capacity(parsed_pairs.len());
                let mut ok = true;
                for (range, _) in &parsed_pairs {
                    match cell_to_key(&CellValue::from_py(range[i].bind(py))) {
                        Some(k) => key.push(k),
                        None => {
                            ok = false;
                            break;
                        }
                    }
                }
                if ok {
                    if let CellValue::Num(v) = CellValue::from_py(sum_range[i].bind(py)) {
                        *sums.entry(key).or_insert(0.0) += v;
                    }
                }
            }
            let mut out = Vec::with_capacity(n);
            for row in 0..n {
                let key: Vec<LookupKey> = parsed_pairs
                    .iter()
                    .map(|(_, crits)| criteria_key(&crits[row]).unwrap())
                    .collect();
                out.push(*sums.get(&key).unwrap_or(&0.0));
            }
            return Ok(out);
        }

        let mut out = Vec::with_capacity(n);
        for row in 0..n {
            let mut total = 0.0;
            for i in 0..range_len {
                let mut all_match = true;
                for (range, crits) in &parsed_pairs {
                    if !matches(&CellValue::from_py(range[i].bind(py)), &crits[row]) {
                        all_match = false;
                        break;
                    }
                }
                if all_match {
                    if let CellValue::Num(v) = CellValue::from_py(sum_range[i].bind(py)) {
                        total += v;
                    }
                }
            }
            out.push(total);
        }
        Ok(out)
    }

    #[pyfunction]
    fn averageifs_vec_values(
        py: Python<'_>,
        average_range: Vec<PyObject>,
        pairs: Vec<(Vec<PyObject>, Vec<PyObject>)>,
    ) -> PyResult<Vec<Option<f64>>> {
        // Same two-tier strategy as sumifs_vec_values/countifs_vec_values
        // just above: a frequency-map fast path when every criteria
        // column is plain equality (the common case), a generic O(n*m)
        // fallback (still correct, just not sped up) when any criteria
        // uses a wildcard or an ordering operator (">"/"<"/etc). Needs
        // both a running total AND a running count per key (unlike
        // SUMIFS' single running total), since AVERAGEIFS returns None
        // (Excel's #DIV/0!) for a row whose group matched but had zero
        // numeric values — not the same as "matched nothing at all".
        if pairs.is_empty() {
            return Err(PyValueError::new_err(
                "AVERAGEIFS: at least one range/criteria pair is required",
            ));
        }
        let range_len = average_range.len();
        let n = pairs[0].1.len();
        for (range, crit_list) in &pairs {
            if range.len() != range_len {
                return Err(PyValueError::new_err(
                    "AVERAGEIFS: all ranges must be the same length as average_range",
                ));
            }
            if crit_list.len() != n {
                return Err(PyValueError::new_err(
                    "AVERAGEIFS: all criteria columns must be the same length",
                ));
            }
        }

        let parsed_pairs: Vec<(&Vec<PyObject>, Vec<Criteria>)> = pairs
            .iter()
            .map(|(r, cl)| {
                let crits: PyResult<Vec<Criteria>> =
                    cl.iter().map(|c| parse_criteria(c.bind(py))).collect();
                Ok::<_, PyErr>((r, crits?))
            })
            .collect::<PyResult<Vec<_>>>()?;

        let all_eq = parsed_pairs
            .iter()
            .all(|(_, crits)| crits.iter().all(|c| criteria_key(c).is_some()));

        if all_eq {
            let mut sums: HashMap<Vec<LookupKey>, (f64, u64)> = HashMap::new();
            for i in 0..range_len {
                let mut key = Vec::with_capacity(parsed_pairs.len());
                let mut ok = true;
                for (range, _) in &parsed_pairs {
                    match cell_to_key(&CellValue::from_py(range[i].bind(py))) {
                        Some(k) => key.push(k),
                        None => {
                            ok = false;
                            break;
                        }
                    }
                }
                if ok {
                    if let CellValue::Num(v) = CellValue::from_py(average_range[i].bind(py)) {
                        let entry = sums.entry(key).or_insert((0.0, 0));
                        entry.0 += v;
                        entry.1 += 1;
                    }
                }
            }
            let mut out = Vec::with_capacity(n);
            for row in 0..n {
                let key: Vec<LookupKey> = parsed_pairs
                    .iter()
                    .map(|(_, crits)| criteria_key(&crits[row]).unwrap())
                    .collect();
                out.push(match sums.get(&key) {
                    Some((total, count)) if *count > 0 => Some(total / *count as f64),
                    _ => None,
                });
            }
            return Ok(out);
        }

        let mut out = Vec::with_capacity(n);
        for row in 0..n {
            let mut total = 0.0;
            let mut count = 0u64;
            for i in 0..range_len {
                let mut all_match = true;
                for (range, crits) in &parsed_pairs {
                    if !matches(&CellValue::from_py(range[i].bind(py)), &crits[row]) {
                        all_match = false;
                        break;
                    }
                }
                if all_match {
                    if let CellValue::Num(v) = CellValue::from_py(average_range[i].bind(py)) {
                        total += v;
                        count += 1;
                    }
                }
            }
            out.push(if count == 0 {
                None
            } else {
                Some(total / count as f64)
            });
        }
        Ok(out)
    }

    #[pyfunction]
    fn minifs_vec_values(
        py: Python<'_>,
        min_range: Vec<PyObject>,
        pairs: Vec<(Vec<PyObject>, Vec<PyObject>)>,
    ) -> PyResult<Vec<f64>> {
        // No frequency-map fast path here (unlike SUMIFS/COUNTIFS/
        // AVERAGEIFS above): MIN/MAX aren't associative-sum-like
        // reductions a HashMap of running totals can accumulate
        // incrementally in a single forward pass — computing a
        // running min per key would work, but with this function's
        // typical (small) row-count for the *IFS-with-vector-criteria
        // shape, the added complexity isn't worth it next to a
        // straightforward O(n*m) scan, which is what this does.
        if pairs.is_empty() {
            return Err(PyValueError::new_err(
                "MINIFS: at least one range/criteria pair is required",
            ));
        }
        let range_len = min_range.len();
        let n = pairs[0].1.len();
        for (range, crit_list) in &pairs {
            if range.len() != range_len {
                return Err(PyValueError::new_err(
                    "MINIFS: all ranges must be the same length as min_range",
                ));
            }
            if crit_list.len() != n {
                return Err(PyValueError::new_err(
                    "MINIFS: all criteria columns must be the same length",
                ));
            }
        }
        let parsed_pairs: Vec<(&Vec<PyObject>, Vec<Criteria>)> = pairs
            .iter()
            .map(|(r, cl)| {
                let crits: PyResult<Vec<Criteria>> =
                    cl.iter().map(|c| parse_criteria(c.bind(py))).collect();
                Ok::<_, PyErr>((r, crits?))
            })
            .collect::<PyResult<Vec<_>>>()?;

        let mut out = Vec::with_capacity(n);
        for row in 0..n {
            let mut result: Option<f64> = None;
            for i in 0..range_len {
                let mut all_match = true;
                for (range, crits) in &parsed_pairs {
                    if !matches(&CellValue::from_py(range[i].bind(py)), &crits[row]) {
                        all_match = false;
                        break;
                    }
                }
                if all_match {
                    if let CellValue::Num(v) = CellValue::from_py(min_range[i].bind(py)) {
                        result = Some(match result {
                            Some(m) if m <= v => m,
                            _ => v,
                        });
                    }
                }
            }
            out.push(result.unwrap_or(0.0));
        }
        Ok(out)
    }

    #[pyfunction]
    fn maxifs_vec_values(
        py: Python<'_>,
        max_range: Vec<PyObject>,
        pairs: Vec<(Vec<PyObject>, Vec<PyObject>)>,
    ) -> PyResult<Vec<f64>> {
        if pairs.is_empty() {
            return Err(PyValueError::new_err(
                "MAXIFS: at least one range/criteria pair is required",
            ));
        }
        let range_len = max_range.len();
        let n = pairs[0].1.len();
        for (range, crit_list) in &pairs {
            if range.len() != range_len {
                return Err(PyValueError::new_err(
                    "MAXIFS: all ranges must be the same length as max_range",
                ));
            }
            if crit_list.len() != n {
                return Err(PyValueError::new_err(
                    "MAXIFS: all criteria columns must be the same length",
                ));
            }
        }
        let parsed_pairs: Vec<(&Vec<PyObject>, Vec<Criteria>)> = pairs
            .iter()
            .map(|(r, cl)| {
                let crits: PyResult<Vec<Criteria>> =
                    cl.iter().map(|c| parse_criteria(c.bind(py))).collect();
                Ok::<_, PyErr>((r, crits?))
            })
            .collect::<PyResult<Vec<_>>>()?;

        let mut out = Vec::with_capacity(n);
        for row in 0..n {
            let mut result: Option<f64> = None;
            for i in 0..range_len {
                let mut all_match = true;
                for (range, crits) in &parsed_pairs {
                    if !matches(&CellValue::from_py(range[i].bind(py)), &crits[row]) {
                        all_match = false;
                        break;
                    }
                }
                if all_match {
                    if let CellValue::Num(v) = CellValue::from_py(max_range[i].bind(py)) {
                        result = Some(match result {
                            Some(m) if m >= v => m,
                            _ => v,
                        });
                    }
                }
            }
            out.push(result.unwrap_or(0.0));
        }
        Ok(out)
    }

    #[pyfunction]
    #[pyo3(signature = (lookup_value, table, col_index, range_lookup=false))]
    fn vlookup_values(
        py: Python<'_>,
        lookup_value: PyObject,
        table: Vec<Vec<PyObject>>,
        col_index: usize,
        range_lookup: bool,
    ) -> PyResult<PyObject> {
        if col_index == 0 {
            return Err(PyValueError::new_err(
                "VLOOKUP: col_index is 1-based; use 1 for the first column",
            ));
        }
        let lv = CellValue::from_py(lookup_value.bind(py));
        if range_lookup {
            let mut best: Option<&Vec<PyObject>> = None;
            for row in &table {
                if row.is_empty() {
                    continue;
                }
                let key = CellValue::from_py(row[0].bind(py));
                let le = match (&lv, &key) {
                    (CellValue::Num(a), CellValue::Num(b)) => *b <= *a,
                    (CellValue::Text(a), CellValue::Text(b)) => {
                        b.to_lowercase() <= a.to_lowercase()
                    }
                    _ => false,
                };
                if le {
                    best = Some(row);
                } else {
                    break;
                }
            }
            match best {
                Some(row) => {
                    let idx = col_index - 1;
                    if idx >= row.len() {
                        return Err(PyValueError::new_err(
                            "VLOOKUP: col_index is out of range for the table",
                        ));
                    }
                    Ok(row[idx].clone_ref(py))
                }
                None => Err(PyValueError::new_err(
                    "VLOOKUP: #N/A - no approximate match found",
                )),
            }
        } else {
            for row in &table {
                if row.is_empty() {
                    continue;
                }
                let key = CellValue::from_py(row[0].bind(py));
                if values_equal(&lv, &key) {
                    let idx = col_index - 1;
                    if idx >= row.len() {
                        return Err(PyValueError::new_err(
                            "VLOOKUP: col_index is out of range for the table",
                        ));
                    }
                    return Ok(row[idx].clone_ref(py));
                }
            }
            Err(PyValueError::new_err("VLOOKUP: #N/A - value not found"))
        }
    }

    #[pyfunction]
    #[pyo3(signature = (lookup_value, lookup_array, return_array, if_not_found=None))]
    fn xlookup_values(
        py: Python<'_>,
        lookup_value: PyObject,
        lookup_array: Vec<PyObject>,
        return_array: Vec<PyObject>,
        if_not_found: Option<PyObject>,
    ) -> PyResult<PyObject> {
        if lookup_array.len() != return_array.len() {
            return Err(PyValueError::new_err(
                "XLOOKUP: lookup_array and return_array must be the same length",
            ));
        }
        let lv = CellValue::from_py(lookup_value.bind(py));
        for (i, item) in lookup_array.iter().enumerate() {
            let key = CellValue::from_py(item.bind(py));
            if values_equal(&lv, &key) {
                return Ok(return_array[i].clone_ref(py));
            }
        }
        match if_not_found {
            Some(v) => Ok(v),
            None => Err(PyValueError::new_err("XLOOKUP: #N/A - value not found")),
        }
    }

    // ========================================================
    // VECTORIZED lookups — look up MANY values at once against
    // the same table/array. Builds one HashMap up front, then
    // resolves every lookup value in O(1), so the whole batch
    // costs O(n + m) instead of O(n * m) for a naive per-row scan.
    // This is what powers df['col'] = VLOOKUP(df['key'], table, ...).
    // ========================================================

    #[pyfunction]
    #[pyo3(signature = (lookup_values, table, col_index, range_lookup=false, if_not_found=None))]
    fn vlookup_many_values(
        py: Python<'_>,
        lookup_values: Vec<PyObject>,
        table: Vec<Vec<PyObject>>,
        col_index: usize,
        range_lookup: bool,
        if_not_found: Option<PyObject>,
    ) -> PyResult<Vec<PyObject>> {
        if col_index == 0 {
            return Err(PyValueError::new_err(
                "VLOOKUP: col_index is 1-based; use 1 for the first column",
            ));
        }
        let idx = col_index - 1;

        if range_lookup {
            // Approximate match assumes the table's first column is sorted
            // ascending — same rule as Excel. Not hashmap-friendly, so this
            // stays a per-value scan, but it's still correct for batches.
            let mut out = Vec::with_capacity(lookup_values.len());
            for lv_obj in &lookup_values {
                let lv = CellValue::from_py(lv_obj.bind(py));
                let mut best: Option<&Vec<PyObject>> = None;
                for row in &table {
                    if row.is_empty() {
                        continue;
                    }
                    let key = CellValue::from_py(row[0].bind(py));
                    let le = match (&lv, &key) {
                        (CellValue::Num(a), CellValue::Num(b)) => *b <= *a,
                        (CellValue::Text(a), CellValue::Text(b)) => {
                            b.to_lowercase() <= a.to_lowercase()
                        }
                        _ => false,
                    };
                    if le {
                        best = Some(row);
                    } else {
                        break;
                    }
                }
                match best {
                    Some(row) => {
                        if idx >= row.len() {
                            return Err(PyValueError::new_err(
                                "VLOOKUP: col_index is out of range for the table",
                            ));
                        }
                        out.push(row[idx].clone_ref(py));
                    }
                    None => match &if_not_found {
                        Some(v) => out.push(v.clone_ref(py)),
                        None => {
                            return Err(PyValueError::new_err(
                                "VLOOKUP: #N/A - no approximate match found",
                            ))
                        }
                    },
                }
            }
            return Ok(out);
        }

        // Exact match: build the HashMap once (first match wins, same as Excel).
        let mut map: HashMap<LookupKey, usize> = HashMap::with_capacity(table.len());
        for (i, row) in table.iter().enumerate() {
            if row.is_empty() {
                continue;
            }
            if let Some(k) = cell_to_key(&CellValue::from_py(row[0].bind(py))) {
                map.entry(k).or_insert(i);
            }
        }
        let mut out = Vec::with_capacity(lookup_values.len());
        for lv_obj in &lookup_values {
            let lv = CellValue::from_py(lv_obj.bind(py));
            let found = cell_to_key(&lv).and_then(|k| map.get(&k).copied());
            match found {
                Some(i) => {
                    let row = &table[i];
                    if idx >= row.len() {
                        return Err(PyValueError::new_err(
                            "VLOOKUP: col_index is out of range for the table",
                        ));
                    }
                    out.push(row[idx].clone_ref(py));
                }
                None => match &if_not_found {
                    Some(v) => out.push(v.clone_ref(py)),
                    None => return Err(PyValueError::new_err("VLOOKUP: #N/A - value not found")),
                },
            }
        }
        Ok(out)
    }

    // ========================================================
    // COLUMNAR fast-path lookups — the same idea as
    // `vlookup_many_values`/`xlookup_many_values` (build one HashMap,
    // resolve every lookup value in O(1)), but working directly
    // against the key/return COLUMNS via `FastColumn` instead of a
    // pre-materialized `Vec<Vec<PyObject>>` table.
    //
    // Why this exists: `vlookup_many_values` above requires the WHOLE
    // table already converted to a Python-level row-major nested list
    // (`table.values.tolist()` for pandas, `table.rows()` for polars)
    // before Rust ever sees it — and that conversion alone measured
    // ~360ms for a 500,000-row table, independent of anything Rust
    // does, because it touches (and boxes into PyObjects) every cell
    // of every row, not just the two columns VLOOKUP/XLOOKUP actually
    // need. This is architecturally the same gap COUNTIF/SUMIFS had
    // before their own `FastColumn` rework — VLOOKUP/XLOOKUP just
    // hadn't been given the same fix yet. `FastColumn::resolve` on
    // the KEY and RETURN columns individually reads them the same
    // zero-copy way COUNTIF/SUMIFS already do (Arrow buffers for
    // text, numpy buffers for numbers), so the whole-table conversion
    // is skipped entirely — the caller only pulls out the two
    // Series/columns it actually needs (`df[key_col]`,
    // `df[return_col]`), which is itself a cheap, already-fast
    // pandas/polars operation, not a row-materializing one.
    // ========================================================

    #[pyfunction]
    #[pyo3(signature = (lookup_values, key_col, return_col, if_not_found=None))]
    fn vlookup_many_columnar(
        py: Python<'_>,
        lookup_values: PyObject,
        key_col: PyObject,
        return_col: PyObject,
        if_not_found: Option<PyObject>,
    ) -> PyResult<Vec<PyObject>> {
        let lookup_col = FastColumn::resolve(lookup_values.bind(py))?;
        let key_column = FastColumn::resolve(key_col.bind(py))?;
        let ret_column = FastColumn::resolve(return_col.bind(py))?;
        let key_n = key_column.len();
        let ret_n = ret_column.len();
        if key_n != ret_n {
            return Err(PyValueError::new_err(
                "VLOOKUP: key column and return column must be the same length",
            ));
        }

        // Split-map design: text and numeric/empty keys go into
        // SEPARATE HashMaps (`HashMap<String, usize>` /
        // `HashMap<LookupKey, usize>`) instead of one shared
        // `HashMap<LookupKey, usize>`. This is what lets the query
        // side use `std::collections::HashMap<String, _>::get(&str)`
        // — the standard library's own `Borrow<str> for String` impl
        // — to probe the map with a BORROWED `&str` and zero
        // allocation, whenever the query key happens to already be
        // lowercase (the overwhelmingly common real-world case: most
        // lookup keys — product codes, IDs, category names — are
        // entered in one consistent case). `LookupKey` (used
        // elsewhere in this file, e.g. `criteria_key` for COUNTIFS/
        // SUMIFS frequency maps) can't offer this: its `Hash`/`Eq`
        // are on the OWNED enum itself, and Rust's `Borrow` trait
        // can't bridge an owned enum variant to an independently-
        // lifetimed borrowed `&str` (verified directly — this was
        // tried first and hits a real lifetime-soundness wall, not
        // just inconvenience). Splitting into two concrete, tightly-
        // typed maps sidesteps the problem entirely.
        //
        // Confirmed by direct benchmark: allocating a fresh
        // lowercased `String` on EVERY query (build once, but query
        // twice — once building the map, once per lookup value) cost
        // ~131ms for 500K queries; checking "is this string already
        // lowercase" first (a cheap byte scan, no allocation) and
        // only falling back to allocating when it's genuinely mixed-
        // case cut that to ~44-57ms — a ~2.6-2.9x reduction in the
        // dominant remaining cost of the columnar VLOOKUP path.
        let mut text_map: HashMap<String, usize> = HashMap::new();
        let mut other_map: HashMap<LookupKey, usize> = HashMap::new();
        for i in 0..key_n {
            match key_column.cell_text_ref_at(py, i) {
                Some(CellTextRef::Borrowed(s)) => {
                    // `s` is confirmed already-lowercase (that's the
                    // only case `cell_text_ref_at` returns Borrowed
                    // for — see its own doc comment), so it can go
                    // straight into the map with no further
                    // `.to_lowercase()` call.
                    if !text_map.contains_key(s) {
                        text_map.insert(s.to_string(), i);
                    }
                }
                Some(CellTextRef::Owned(s)) => {
                    // Already-lowercased by `cell_text_ref_at` itself
                    // (it only returns Owned when it had to allocate
                    // anyway) — insert as-is, no redundant work.
                    text_map.entry(s).or_insert(i);
                }
                None => {
                    if let Some(k) = key_column.cell_key_at(py, i) {
                        other_map.entry(k).or_insert(i);
                    }
                }
            }
        }

        #[inline(always)]
        fn lookup_row(
            py: Python<'_>,
            col: &FastColumn,
            i: usize,
            text_map: &HashMap<String, usize>,
            other_map: &HashMap<LookupKey, usize>,
        ) -> Option<usize> {
            match col.cell_text_ref_at(py, i) {
                Some(CellTextRef::Borrowed(s)) => {
                    // Zero-allocation query: `s` is confirmed already
                    // lowercase, so it can probe `text_map` directly
                    // via `HashMap<String,_>`'s standard-library
                    // `Borrow<str>` impl — no `.to_lowercase()` call,
                    // no allocation, for this (the common) case.
                    text_map.get(s).copied()
                }
                Some(CellTextRef::Owned(s)) => text_map.get(&s).copied(),
                None => col
                    .cell_key_at(py, i)
                    .and_then(|k| other_map.get(&k).copied()),
            }
        }

        let m = lookup_col.len();
        let mut out = Vec::with_capacity(m);
        for i in 0..m {
            let found = lookup_row(py, &lookup_col, i, &text_map, &other_map);
            match found {
                Some(row_i) => out.push(ret_column.to_object_at(py, row_i)),
                None => match &if_not_found {
                    Some(v) => out.push(v.clone_ref(py)),
                    None => return Err(PyValueError::new_err("VLOOKUP: #N/A - value not found")),
                },
            }
        }
        Ok(out)
    }

    #[pyfunction]
    #[pyo3(signature = (lookup_values, lookup_array, return_array, if_not_found=None))]
    fn xlookup_many_columnar(
        py: Python<'_>,
        lookup_values: PyObject,
        lookup_array: PyObject,
        return_array: PyObject,
        if_not_found: Option<PyObject>,
    ) -> PyResult<Vec<PyObject>> {
        // XLOOKUP's key/return columns are exactly VLOOKUP's
        // key_col/return_col — same shape, same fast path, different
        // Excel-facing name only. Reusing the implementation keeps
        // both genuinely in sync rather than risking two copies of
        // the same HashMap-building logic drifting apart.
        vlookup_many_columnar(py, lookup_values, lookup_array, return_array, if_not_found)
    }

    #[pyfunction]
    #[pyo3(signature = (lookup_values, lookup_array, return_array, if_not_found=None))]
    fn xlookup_many_values(
        py: Python<'_>,
        lookup_values: Vec<PyObject>,
        lookup_array: Vec<PyObject>,
        return_array: Vec<PyObject>,
        if_not_found: Option<PyObject>,
    ) -> PyResult<Vec<PyObject>> {
        if lookup_array.len() != return_array.len() {
            return Err(PyValueError::new_err(
                "XLOOKUP: lookup_array and return_array must be the same length",
            ));
        }
        let mut map: HashMap<LookupKey, usize> = HashMap::with_capacity(lookup_array.len());
        for (i, item) in lookup_array.iter().enumerate() {
            if let Some(k) = cell_to_key(&CellValue::from_py(item.bind(py))) {
                map.entry(k).or_insert(i);
            }
        }
        let mut out = Vec::with_capacity(lookup_values.len());
        for lv_obj in &lookup_values {
            let lv = CellValue::from_py(lv_obj.bind(py));
            let found = cell_to_key(&lv).and_then(|k| map.get(&k).copied());
            match found {
                Some(i) => out.push(return_array[i].clone_ref(py)),
                None => match &if_not_found {
                    Some(v) => out.push(v.clone_ref(py)),
                    None => return Err(PyValueError::new_err("XLOOKUP: #N/A - value not found")),
                },
            }
        }
        Ok(out)
    }

    // ========================================================
    // INDEX-ONLY resolution — used when the caller wants MULTIPLE
    // return columns (XLOOKUP's return_array or LOOKUPIFS'
    // return_range can be a whole sub-table). Rust's job here is
    // only "which row(s) match" — picking values out of one or
    // many requested columns for those rows is plain indexing,
    // which the Python layer already does well without needing
    // any of this module's matching logic duplicated per column.
    // ========================================================

    #[pyfunction]
    fn xlookup_many_indices(
        py: Python<'_>,
        lookup_values: Vec<PyObject>,
        lookup_array: Vec<PyObject>,
    ) -> Vec<Option<i64>> {
        let mut map: HashMap<LookupKey, usize> = HashMap::with_capacity(lookup_array.len());
        for (i, item) in lookup_array.iter().enumerate() {
            if let Some(k) = cell_to_key(&CellValue::from_py(item.bind(py))) {
                map.entry(k).or_insert(i);
            }
        }
        lookup_values
            .iter()
            .map(|v| {
                cell_to_key(&CellValue::from_py(v.bind(py)))
                    .and_then(|k| map.get(&k).copied())
                    .map(|i| i as i64)
            })
            .collect()
    }

    /// LOOKUPIFS core: for each of the m output rows (one per vectorized
    /// criteria batch position — length 1 for a plain scalar call), return
    /// the list of row-indices in the range(s) that satisfy the AND'd
    /// criteria. Same HashMap-fast-path / scan-fallback shape as
    /// `countifs_vec_values`, except it collects indices instead of a count.
    #[pyfunction]
    fn lookupifs_indices_values(
        py: Python<'_>,
        pairs: Vec<(Vec<PyObject>, Vec<PyObject>)>,
    ) -> PyResult<Vec<Vec<i64>>> {
        if pairs.is_empty() {
            return Err(PyValueError::new_err(
                "LOOKUPIFS: at least one range/criteria pair is required",
            ));
        }
        let range_len = pairs[0].0.len();
        let n = pairs[0].1.len();
        for (range, crit_list) in &pairs {
            if range.len() != range_len {
                return Err(PyValueError::new_err(
                    "LOOKUPIFS: all ranges must be the same length",
                ));
            }
            if crit_list.len() != n {
                return Err(PyValueError::new_err(
                    "LOOKUPIFS: all criteria columns must be the same length",
                ));
            }
        }

        let parsed_pairs: Vec<(&Vec<PyObject>, Vec<Criteria>)> = pairs
            .iter()
            .map(|(r, cl)| {
                let crits: PyResult<Vec<Criteria>> =
                    cl.iter().map(|c| parse_criteria(c.bind(py))).collect();
                Ok::<_, PyErr>((r, crits?))
            })
            .collect::<PyResult<Vec<_>>>()?;

        let all_eq = parsed_pairs
            .iter()
            .all(|(_, crits)| crits.iter().all(|c| criteria_key(c).is_some()));

        if all_eq {
            let mut index_map: HashMap<Vec<LookupKey>, Vec<usize>> = HashMap::new();
            for i in 0..range_len {
                let mut key = Vec::with_capacity(parsed_pairs.len());
                let mut ok = true;
                for (range, _) in &parsed_pairs {
                    match cell_to_key(&CellValue::from_py(range[i].bind(py))) {
                        Some(k) => key.push(k),
                        None => {
                            ok = false;
                            break;
                        }
                    }
                }
                if ok {
                    index_map.entry(key).or_insert_with(Vec::new).push(i);
                }
            }
            let mut out = Vec::with_capacity(n);
            for row in 0..n {
                let key: Vec<LookupKey> = parsed_pairs
                    .iter()
                    .map(|(_, crits)| criteria_key(&crits[row]).unwrap())
                    .collect();
                let indices = index_map
                    .get(&key)
                    .map(|v| v.iter().map(|&i| i as i64).collect())
                    .unwrap_or_else(Vec::new);
                out.push(indices);
            }
            return Ok(out);
        }

        let mut out = Vec::with_capacity(n);
        for row in 0..n {
            let mut matches_for_row = Vec::new();
            for i in 0..range_len {
                let mut all_match = true;
                for (range, crits) in &parsed_pairs {
                    if !matches(&CellValue::from_py(range[i].bind(py)), &crits[row]) {
                        all_match = false;
                        break;
                    }
                }
                if all_match {
                    matches_for_row.push(i as i64);
                }
            }
            out.push(matches_for_row);
        }
        Ok(out)
    }

    // ========================================================
    // MIXED fast path — SUM/AVERAGE/COUNT/*IF/*IFS
    //
    // Each of these takes the ORIGINAL Python object directly (a numpy
    // array, a pandas/polars Series/column, or a plain list/tuple) —
    // NOT a pre-converted Vec<PyObject> — and resolves each column
    // exactly once via FastColumn::resolve(). A column that's already a
    // numeric buffer runs at native speed with zero per-element Python
    // API calls; a column that isn't (text, mixed types, a plain list)
    // transparently falls back to the existing, fully-generic
    // CellValue-based handling for THAT column only — every other
    // column in the same call keeps its own fast path independently.
    //
    // These are genuinely additive: the existing `sum_values`,
    // `sumif_values`, `sumifs_values`, etc. are completely unchanged and
    // still available as the always-correct, always-available fallback
    // (used by the Python layer for plain lists/tuples, where resolving
    // a FastColumn has no advantage over the existing generic path).
    // ========================================================

    #[pyfunction]
    fn sum_mixed(py: Python<'_>, values: PyObject) -> PyResult<f64> {
        let col = FastColumn::resolve(values.bind(py))?;
        Ok(match &col {
            FastColumn::Numeric(a) => a
                .as_slice()
                .map(|s| s.iter().filter(|v| !v.is_nan()).sum())
                .unwrap_or(0.0),
            FastColumn::NumericI64(a) => a
                .as_slice()
                .map(|s| s.iter().map(|&x| x as f64).sum())
                .unwrap_or(0.0),
            // A text column contributes nothing to SUM (matches
            // CellValue::Text's existing "ignored by SUM" semantics).
            FastColumn::Text(_) => 0.0,
            FastColumn::NativeText(_) => 0.0,
            FastColumn::ArrowText(_) => 0.0,
            FastColumn::Generic(v) => {
                let mut total = 0.0;
                for obj in v {
                    if let CellValue::Num(n) = CellValue::from_py(obj.bind(py)) {
                        total += n;
                    }
                }
                total
            }
        })
    }

    #[pyfunction]
    fn average_mixed(py: Python<'_>, values: PyObject) -> PyResult<f64> {
        let col = FastColumn::resolve(values.bind(py))?;
        let (total, count): (f64, u64) = match &col {
            FastColumn::Numeric(a) => {
                let s = a.as_slice().unwrap_or(&[]);
                let mut total = 0.0;
                let mut count = 0u64;
                for &v in s {
                    if !v.is_nan() {
                        total += v;
                        count += 1;
                    }
                }
                (total, count)
            }
            FastColumn::NumericI64(a) => {
                let s = a.as_slice().unwrap_or(&[]);
                (s.iter().map(|&x| x as f64).sum(), s.len() as u64)
            }
            // Text contributes neither to the sum nor the count
            // (matches CellValue::Text — ignored by AVERAGE).
            FastColumn::Text(_) => (0.0, 0),
            FastColumn::NativeText(_) => (0.0, 0),
            FastColumn::ArrowText(_) => (0.0, 0),
            FastColumn::Generic(v) => {
                let mut total = 0.0;
                let mut count = 0u64;
                for obj in v {
                    if let CellValue::Num(n) = CellValue::from_py(obj.bind(py)) {
                        total += n;
                        count += 1;
                    }
                }
                (total, count)
            }
        };
        if count == 0 {
            return Err(PyValueError::new_err(
                "AVERAGE: no numeric values found (division by zero)",
            ));
        }
        Ok(total / count as f64)
    }

    #[pyfunction]
    fn count_mixed(py: Python<'_>, values: PyObject) -> PyResult<i64> {
        let col = FastColumn::resolve(values.bind(py))?;
        Ok(match &col {
            FastColumn::Numeric(a) => a
                .as_slice()
                .map(|s| s.iter().filter(|v| !v.is_nan()).count())
                .unwrap_or(0) as i64,
            FastColumn::NumericI64(a) => a.as_slice().map(|s| s.len()).unwrap_or(0) as i64,
            // COUNT is Excel's "numeric cells only" count — a text
            // column contributes zero (matches CellValue::Text).
            FastColumn::Text(_) => 0,
            FastColumn::NativeText(_) => 0,
            FastColumn::ArrowText(_) => 0,
            FastColumn::Generic(v) => {
                let mut c = 0i64;
                for obj in v {
                    if let CellValue::Num(_) = CellValue::from_py(obj.bind(py)) {
                        c += 1;
                    }
                }
                c
            }
        })
    }

    #[pyfunction]
    fn countif_mixed(py: Python<'_>, range: PyObject, criteria: PyObject) -> PyResult<i64> {
        let col = FastColumn::resolve(range.bind(py))?;
        let crit = parse_criteria(criteria.bind(py))?;
        let n = col.len();
        let mask = col.build_mask(py, &crit, n);
        Ok(mask.iter().filter(|&&b| b).count() as i64)
    }

    #[pyfunction]
    fn _debug_resolve_only(py: Python<'_>, range: PyObject) -> PyResult<i64> {
        let col = FastColumn::resolve(range.bind(py))?;
        Ok(col.len() as i64)
    }

    #[pyfunction]
    fn _debug_string_extract_only(py: Python<'_>, range: PyObject) -> PyResult<i64> {
        let obj = range.bind(py);
        let list = obj.downcast::<pyo3::types::PyList>().unwrap();
        let len = list.len();
        let mut v: Vec<String> = Vec::with_capacity(len);
        for i in 0..len {
            let item = list.get_item(i).unwrap();
            let s = item.downcast::<PyString>().unwrap();
            v.push(s.to_string_lossy().into_owned());
        }
        Ok(v.len() as i64)
    }

    #[pyfunction]
    #[pyo3(signature = (range, criteria, sum_range=None))]
    fn sumif_mixed(
        py: Python<'_>,
        range: PyObject,
        criteria: PyObject,
        sum_range: Option<PyObject>,
    ) -> PyResult<f64> {
        let filter_col = FastColumn::resolve(range.bind(py))?;
        let crit = parse_criteria(criteria.bind(py))?;
        let target_col = match &sum_range {
            Some(sr) => FastColumn::resolve(sr.bind(py))?,
            None => FastColumn::resolve(range.bind(py))?,
        };
        if target_col.len() != filter_col.len() {
            return Err(PyValueError::new_err(
                "SUMIF: sum_range must be the same length as range",
            ));
        }
        let n = filter_col.len();
        let mask = filter_col.build_mask(py, &crit, n);
        let mut total = 0.0;
        for i in 0..n {
            if mask[i] {
                if let Some(v) = target_col.numeric_at(py, i) {
                    total += v;
                }
            }
        }
        Ok(total)
    }

    #[pyfunction]
    #[pyo3(signature = (range, criteria, average_range=None))]
    fn averageif_mixed(
        py: Python<'_>,
        range: PyObject,
        criteria: PyObject,
        average_range: Option<PyObject>,
    ) -> PyResult<f64> {
        let filter_col = FastColumn::resolve(range.bind(py))?;
        let crit = parse_criteria(criteria.bind(py))?;
        let target_col = match &average_range {
            Some(ar) => FastColumn::resolve(ar.bind(py))?,
            None => FastColumn::resolve(range.bind(py))?,
        };
        if target_col.len() != filter_col.len() {
            return Err(PyValueError::new_err(
                "AVERAGEIF: average_range must be the same length as range",
            ));
        }
        let n = filter_col.len();
        let mask = filter_col.build_mask(py, &crit, n);
        let mut total = 0.0;
        let mut count = 0u64;
        for i in 0..n {
            if mask[i] {
                if let Some(v) = target_col.numeric_at(py, i) {
                    total += v;
                    count += 1;
                }
            }
        }
        if count == 0 {
            return Err(PyValueError::new_err(
                "AVERAGEIF: no matching numeric values found",
            ));
        }
        Ok(total / count as f64)
    }

    #[pyfunction]
    fn countifs_mixed(py: Python<'_>, pairs: Vec<(PyObject, PyObject)>) -> PyResult<i64> {
        if pairs.is_empty() {
            return Err(PyValueError::new_err(
                "COUNTIFS: at least one range/criteria pair is required",
            ));
        }
        let mut cols = Vec::with_capacity(pairs.len());
        for (r, c) in &pairs {
            let col = FastColumn::resolve(r.bind(py))?;
            let crit = parse_criteria(c.bind(py))?;
            cols.push((col, crit));
        }
        let n = cols[0].0.len();
        for (col, _) in &cols {
            if col.len() != n {
                return Err(PyValueError::new_err(
                    "COUNTIFS: all ranges must be the same length",
                ));
            }
        }
        // Build one mask per (col, crit) pair — each individually takes
        // the fast path when eligible — then AND them together and
        // count. Mirrors polars' own "one boolean mask per predicate,
        // combine, aggregate" structure rather than re-checking every
        // predicate per row through the generic per-row dispatch.
        let masks: Vec<Vec<bool>> = cols
            .iter()
            .map(|(col, crit)| col.build_mask(py, crit, n))
            .collect();
        let count = count_all_masks_true(&masks, n);
        Ok(count)
    }

    #[pyfunction]
    fn sumifs_mixed(
        py: Python<'_>,
        sum_range: PyObject,
        pairs: Vec<(PyObject, PyObject)>,
    ) -> PyResult<f64> {
        if pairs.is_empty() {
            return Err(PyValueError::new_err(
                "SUMIFS: at least one range/criteria pair is required",
            ));
        }
        let sum_col = FastColumn::resolve(sum_range.bind(py))?;
        let mut cols = Vec::with_capacity(pairs.len());
        for (r, c) in &pairs {
            let col = FastColumn::resolve(r.bind(py))?;
            let crit = parse_criteria(c.bind(py))?;
            cols.push((col, crit));
        }
        let n = sum_col.len();
        for (col, _) in &cols {
            if col.len() != n {
                return Err(PyValueError::new_err(
                    "SUMIFS: all ranges must be the same length as sum_range",
                ));
            }
        }
        let masks: Vec<Vec<bool>> = cols
            .iter()
            .map(|(col, crit)| col.build_mask(py, crit, n))
            .collect();
        // Fast, GIL-free, potentially-parallel path when sum_range is a
        // genuinely zero-copy numeric column (the common case) — falls
        // back to the original py-bound per-row `numeric_at` loop
        // (always correct, just not parallelizable) for the rare case
        // where sum_range resolved to `Generic`.
        let total = match sum_col.as_f64_slice() {
            Some(values) => sum_where_all_masks_true(&values, &masks, n),
            None => {
                let mut total = 0.0;
                for i in 0..n {
                    if masks.iter().all(|m| m[i]) {
                        if let Some(v) = sum_col.numeric_at(py, i) {
                            total += v;
                        }
                    }
                }
                total
            }
        };
        Ok(total)
    }

    // ========================================================
    // AVERAGEIFS — same mask-then-aggregate architecture as
    // SUMIFS/COUNTIFS above (one boolean mask per predicate, AND them
    // together, then reduce over the matched rows). AVERAGEIFS didn't
    // exist in this build before; adding it here rather than routing
    // through a generic/legacy path means it gets the same zero-copy
    // Arrow/numpy fast paths COUNTIFS/SUMIFS already have, from day
    // one — never a slow function that later needs the same rework
    // those two already went through.
    // ========================================================

    #[pyfunction]
    fn averageifs_mixed(
        py: Python<'_>,
        average_range: PyObject,
        pairs: Vec<(PyObject, PyObject)>,
    ) -> PyResult<f64> {
        if pairs.is_empty() {
            return Err(PyValueError::new_err(
                "AVERAGEIFS: at least one range/criteria pair is required",
            ));
        }
        let avg_col = FastColumn::resolve(average_range.bind(py))?;
        let mut cols = Vec::with_capacity(pairs.len());
        for (r, c) in &pairs {
            let col = FastColumn::resolve(r.bind(py))?;
            let crit = parse_criteria(c.bind(py))?;
            cols.push((col, crit));
        }
        let n = avg_col.len();
        for (col, _) in &cols {
            if col.len() != n {
                return Err(PyValueError::new_err(
                    "AVERAGEIFS: all ranges must be the same length as average_range",
                ));
            }
        }
        let masks: Vec<Vec<bool>> = cols
            .iter()
            .map(|(col, crit)| col.build_mask(py, crit, n))
            .collect();
        // Same fast, GIL-free, potentially-parallel path SUMIFS uses
        // when average_range is a genuinely zero-copy numeric column
        // — falls back to the py-bound per-row loop otherwise. Needs
        // both the sum AND the count of matched numeric rows (unlike
        // SUMIFS, which only needs the sum), so this doesn't reuse
        // `sum_where_all_masks_true` directly — that function returns
        // only a sum, and re-deriving the count from it wouldn't be
        // reliable (a masked row could still be non-numeric/NaN and
        // contribute 0 to the sum while not counting toward the
        // average's denominator).
        let (total, count) = match avg_col.as_f64_slice() {
            Some(values) => sum_and_count_where_all_masks_true(&values, &masks, n),
            None => {
                let mut total = 0.0;
                let mut count = 0u64;
                for i in 0..n {
                    if masks.iter().all(|m| m[i]) {
                        if let Some(v) = avg_col.numeric_at(py, i) {
                            total += v;
                            count += 1;
                        }
                    }
                }
                (total, count)
            }
        };
        if count == 0 {
            return Err(PyValueError::new_err(
                "AVERAGEIFS: no matching numeric values found",
            ));
        }
        Ok(total / count as f64)
    }

    // ========================================================
    // MIN/MAX and MINIFS/MAXIFS — MIN/MAX reduce directly over a
    // zero-copy `f64` slice with no PyObject touched per element
    // (same `as_f64_slice` building block SUMIFS/AVERAGEIFS use);
    // MINIFS/MAXIFS build masks the same way COUNTIFS/SUMIFS/
    // AVERAGEIFS do, then reduce only over the matched rows.
    // ========================================================

    #[pyfunction]
    fn min_mixed(py: Python<'_>, values: PyObject) -> PyResult<f64> {
        let col = FastColumn::resolve(values.bind(py))?;
        let n = col.len();
        let result = match col.as_f64_slice() {
            Some(s) => s.iter().filter(|v| !v.is_nan()).fold(None, |acc, &v| {
                Some(match acc {
                    Some(m) if m <= v => m,
                    _ => v,
                })
            }),
            None => {
                let mut result: Option<f64> = None;
                for i in 0..n {
                    if let Some(v) = col.numeric_at(py, i) {
                        result = Some(match result {
                            Some(m) if m <= v => m,
                            _ => v,
                        });
                    }
                }
                result
            }
        };
        result.ok_or_else(|| PyValueError::new_err("MIN: no numeric values found"))
    }

    #[pyfunction]
    fn max_mixed(py: Python<'_>, values: PyObject) -> PyResult<f64> {
        let col = FastColumn::resolve(values.bind(py))?;
        let n = col.len();
        let result = match col.as_f64_slice() {
            Some(s) => s.iter().filter(|v| !v.is_nan()).fold(None, |acc, &v| {
                Some(match acc {
                    Some(m) if m >= v => m,
                    _ => v,
                })
            }),
            None => {
                let mut result: Option<f64> = None;
                for i in 0..n {
                    if let Some(v) = col.numeric_at(py, i) {
                        result = Some(match result {
                            Some(m) if m >= v => m,
                            _ => v,
                        });
                    }
                }
                result
            }
        };
        result.ok_or_else(|| PyValueError::new_err("MAX: no numeric values found"))
    }

    #[pyfunction]
    fn minifs_mixed(
        py: Python<'_>,
        min_range: PyObject,
        pairs: Vec<(PyObject, PyObject)>,
    ) -> PyResult<f64> {
        let target_col = FastColumn::resolve(min_range.bind(py))?;
        if pairs.is_empty() {
            return Err(PyValueError::new_err(
                "MINIFS: at least one range/criteria pair is required",
            ));
        }
        let mut cols = Vec::with_capacity(pairs.len());
        for (r, c) in &pairs {
            let col = FastColumn::resolve(r.bind(py))?;
            let crit = parse_criteria(c.bind(py))?;
            cols.push((col, crit));
        }
        let n = target_col.len();
        for (col, _) in &cols {
            if col.len() != n {
                return Err(PyValueError::new_err(
                    "MINIFS: all ranges must be the same length as min_range",
                ));
            }
        }
        let masks: Vec<Vec<bool>> = cols
            .iter()
            .map(|(col, crit)| col.build_mask(py, crit, n))
            .collect();
        let mut result: Option<f64> = None;
        for i in 0..n {
            if masks.iter().all(|m| m[i]) {
                if let Some(v) = target_col.numeric_at(py, i) {
                    result = Some(match result {
                        Some(m) if m <= v => m,
                        _ => v,
                    });
                }
            }
        }
        // Excel's own MINIFS returns 0 (not an error) when nothing
        // matches — different from MIN/AVERAGEIFS/SUMIFS's own
        // "empty input" conventions elsewhere in this file, but this
        // matches real Excel behavior exactly, which is the contract
        // this whole library is built around.
        Ok(result.unwrap_or(0.0))
    }

    #[pyfunction]
    fn maxifs_mixed(
        py: Python<'_>,
        max_range: PyObject,
        pairs: Vec<(PyObject, PyObject)>,
    ) -> PyResult<f64> {
        let target_col = FastColumn::resolve(max_range.bind(py))?;
        if pairs.is_empty() {
            return Err(PyValueError::new_err(
                "MAXIFS: at least one range/criteria pair is required",
            ));
        }
        let mut cols = Vec::with_capacity(pairs.len());
        for (r, c) in &pairs {
            let col = FastColumn::resolve(r.bind(py))?;
            let crit = parse_criteria(c.bind(py))?;
            cols.push((col, crit));
        }
        let n = target_col.len();
        for (col, _) in &cols {
            if col.len() != n {
                return Err(PyValueError::new_err(
                    "MAXIFS: all ranges must be the same length as max_range",
                ));
            }
        }
        let masks: Vec<Vec<bool>> = cols
            .iter()
            .map(|(col, crit)| col.build_mask(py, crit, n))
            .collect();
        let mut result: Option<f64> = None;
        for i in 0..n {
            if masks.iter().all(|m| m[i]) {
                if let Some(v) = target_col.numeric_at(py, i) {
                    result = Some(match result {
                        Some(m) if m >= v => m,
                        _ => v,
                    });
                }
            }
        }
        // Same Excel convention as MINIFS: 0, not an error, when
        // nothing matches.
        Ok(result.unwrap_or(0.0))
    }

    // ========================================================
    // INFO/CLEAN acceleration — summarize_numeric_column and
    // summarize_text_column below are the Rust-backed fast paths
    // for `_summarize_column` in the Python layer. Added after
    // profiling showed the pure-Python version spending ~95% of its
    // time in a per-value loop (`isinstance` checks, dict/list
    // mutation) — exactly the per-row-Python-object-touch pattern
    // COUNTIF/SUMIF were already fixed for via `FastColumn`. These
    // reuse that same architecture: a `Numeric`/`NumericI64` column's
    // stats are computed entirely on a zero-copy `&[f64]` slice, no
    // PyObject touched at all; a text column's stats read through
    // `ArrowText`/`NativeText`/`Text`'s zero-copy `&str` access.
    // Returns `None` (via Python's own None, checked by the caller)
    // when the fast path doesn't apply (a `Generic`/boolean/mixed
    // column) so the existing pure-Python path remains the correct,
    // always-available fallback — this is an ADDITIVE speed-up, not
    // a replacement of the Python logic's behavior or edge cases.
    // ========================================================

    #[pyfunction]
    fn summarize_numeric_column(py: Python<'_>, values: PyObject) -> PyResult<Option<PyObject>> {
        let col = FastColumn::resolve(values.bind(py))?;
        let slice = match col.as_f64_slice() {
            Some(s) => s,
            None => return Ok(None),
        };
        let total = slice.len();
        let mut nums: Vec<f64> = Vec::with_capacity(total);
        let mut missing = 0u64;
        for &v in slice.iter() {
            if v.is_nan() {
                missing += 1;
            } else {
                nums.push(v);
            }
        }
        let n = nums.len();
        let dict = PyDict::new_bound(py);
        dict.set_item("total", total)?;
        dict.set_item("missing", missing)?;
        if n == 0 {
            dict.set_item("kind", "empty")?;
            return Ok(Some(dict.into_py(py)));
        }
        let sum: f64 = nums.iter().sum();
        let mean = sum / n as f64;
        let mn = nums.iter().cloned().fold(f64::INFINITY, f64::min);
        let mx = nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max);
        nums.sort_by(|a, b| a.partial_cmp(b).unwrap());
        let mid = n / 2;
        let median = if n % 2 == 0 {
            (nums[mid - 1] + nums[mid]) / 2.0
        } else {
            nums[mid]
        };
        let variance: f64 = nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / n as f64;
        let std = variance.sqrt();
        let zeros = nums.iter().filter(|&&x| x == 0.0).count();
        let negatives = nums.iter().filter(|&&x| x < 0.0).count();
        #[inline(always)]
        fn percentile_sorted(sorted: &[f64], p: f64) -> f64 {
            if sorted.is_empty() {
                return f64::NAN;
            }
            let idx = (p * (sorted.len() as f64 - 1.0)).round() as usize;
            sorted[idx.min(sorted.len() - 1)]
        }
        let q1 = percentile_sorted(&nums, 0.25);
        let q3 = percentile_sorted(&nums, 0.75);
        let iqr = q3 - q1;
        let (lower, upper) = (q1 - 1.5 * iqr, q3 + 1.5 * iqr);
        let outlier_count = nums.iter().filter(|&&x| x < lower || x > upper).count();
        // Unique count: exact, via a hashable-key set over the same
        // f64 values — matches Python's `len(set(sorted_nums))`
        // exactly (both use IEEE-754 bit patterns as the effective
        // equality/hash key for a float).
        let mut unique_set: std::collections::HashSet<u64> =
            std::collections::HashSet::with_capacity(n);
        for &v in &nums {
            unique_set.insert(v.to_bits());
        }
        dict.set_item("kind", "numeric")?;
        dict.set_item("unique", unique_set.len())?;
        dict.set_item("sum", sum)?;
        dict.set_item("mean", mean)?;
        dict.set_item("min", mn)?;
        dict.set_item("max", mx)?;
        dict.set_item("median", median)?;
        dict.set_item("std", std)?;
        dict.set_item("zeros", zeros)?;
        dict.set_item("negatives", negatives)?;
        dict.set_item("q1", q1)?;
        dict.set_item("q3", q3)?;
        dict.set_item("outlier_count", outlier_count)?;
        dict.set_item("is_constant", mn == mx)?;
        Ok(Some(dict.into_py(py)))
    }

    #[pyfunction]
    #[pyo3(signature = (values, top_n, categorical_max_unique, categorical_max_ratio))]
    fn summarize_text_column(
        py: Python<'_>,
        values: PyObject,
        top_n: usize,
        categorical_max_unique: usize,
        categorical_max_ratio: f64,
    ) -> PyResult<Option<PyObject>> {
        let col = FastColumn::resolve(values.bind(py))?;
        let total = col.len();
        let mut texts: Vec<String> = Vec::with_capacity(total);
        let mut missing = 0u64;
        let mut any_non_text = false;
        for i in 0..total {
            match col.cell_text_exact_at(py, i) {
                Some(s) => {
                    if s.trim().is_empty() {
                        missing += 1;
                    } else {
                        texts.push(s.to_string());
                    }
                }
                None => {
                    // Not a text-bearing FastColumn variant at all
                    // (Numeric/NumericI64/Generic) — this fast path is
                    // for genuinely text-typed columns; anything else
                    // (including a column that's ALL null, which also
                    // can't tell text from non-text at this level)
                    // falls back to the existing Python classification.
                    any_non_text = true;
                    break;
                }
            }
        }
        if any_non_text {
            return Ok(None);
        }

        let n = texts.len();
        let dict = PyDict::new_bound(py);
        dict.set_item("total", total)?;
        dict.set_item("missing", missing)?;
        if n == 0 {
            dict.set_item("kind", "empty")?;
            return Ok(Some(dict.into_py(py)));
        }

        let mut counts: HashMap<String, usize> = HashMap::new();
        let mut numeric_looking = 0usize;
        let mut date_looking = 0usize;
        // Same advisory-sampling approach as the Python original: a
        // stride across the full column (not just the first K), since
        // these two counts are heuristic flags, not exact statistics
        // — `unique`/`missing` below ARE always exact.
        let sample_cap = 5000usize;
        let stride = (n / sample_cap).max(1);
        let mut sampled = 0usize;
        for (i, t) in texts.iter().enumerate() {
            *counts.entry(t.clone()).or_insert(0) += 1;
            if i % stride == 0 {
                sampled += 1;
                let trimmed = t.trim();
                if trimmed.parse::<f64>().is_ok() {
                    numeric_looking += 1;
                }
                if looks_like_date(trimmed) {
                    date_looking += 1;
                }
            }
        }
        if sampled > 0 && stride > 1 {
            let scale = n as f64 / sampled as f64;
            numeric_looking = (numeric_looking as f64 * scale).round() as usize;
            date_looking = (date_looking as f64 * scale).round() as usize;
        }

        let unique = counts.len();
        let is_categorical = unique <= categorical_max_unique
            || (unique as f64 / n.max(1) as f64) <= categorical_max_ratio;
        let is_constant = unique == 1;

        dict.set_item("kind", "text")?;
        dict.set_item("unique", unique)?;
        dict.set_item("is_categorical", is_categorical)?;
        dict.set_item("is_constant", is_constant)?;
        dict.set_item("numeric_looking", numeric_looking)?;
        dict.set_item("date_looking", date_looking)?;

        if is_categorical {
            let mut ranked: Vec<(&String, &usize)> = counts.iter().collect();
            ranked.sort_by(|a, b| b.1.cmp(a.1).then_with(|| a.0.cmp(b.0)));
            let top: Vec<(String, usize)> = ranked
                .iter()
                .take(top_n)
                .map(|(k, v)| ((*k).clone(), **v))
                .collect();
            let bottom_n = top_n.min(5);
            let bottom: Vec<(String, usize)> = ranked
                .iter()
                .rev()
                .take(bottom_n)
                .map(|(k, v)| ((*k).clone(), **v))
                .collect();
            dict.set_item("top_categories", top)?;
            dict.set_item("bottom_categories", bottom)?;

            let mut norm_groups: HashMap<String, Vec<String>> = HashMap::new();
            for raw_key in counts.keys() {
                norm_groups
                    .entry(raw_key.trim().to_lowercase())
                    .or_default()
                    .push(raw_key.clone());
            }
            let normalized_unique = norm_groups.len();
            let mut near_dup_example = String::new();
            if normalized_unique < unique {
                for variants in norm_groups.values() {
                    if variants.len() > 1 {
                        let mut sorted_variants = variants.clone();
                        sorted_variants.sort();
                        near_dup_example = sorted_variants.join("/");
                        break;
                    }
                }
            }
            dict.set_item("normalized_unique", normalized_unique)?;
            dict.set_item("near_dup_example", near_dup_example)?;
            dict.set_item("most_frequent", "")?;
            dict.set_item("most_frequent_count", 0)?;
        } else {
            let mut best_key = String::new();
            let mut best_count = 0usize;
            for (k, &c) in counts.iter() {
                if c > best_count || (c == best_count && (best_key.is_empty() || *k < best_key)) {
                    best_key = k.clone();
                    best_count = c;
                }
            }
            dict.set_item("most_frequent", best_key)?;
            dict.set_item("most_frequent_count", best_count)?;
            dict.set_item("normalized_unique", unique)?;
            dict.set_item("near_dup_example", "")?;
            dict.set_item("top_categories", Vec::<(String, usize)>::new())?;
            dict.set_item("bottom_categories", Vec::<(String, usize)>::new())?;
        }
        Ok(Some(dict.into_py(py)))
    }

    /// CLEAN()'s `merge_categories` action: given a text column,
    /// groups values by case/whitespace-normalized form and, for
    /// every group with more than one distinct raw spelling, maps
    /// every non-canonical spelling to the canonical (most common;
    /// ties broken toward the FIRST-SEEN casing — see the extensive
    /// comment on this exact tie-break in the Python
    /// `_build_cleaning_plan` this accelerates, which documents the
    /// real bug this rule was written to avoid: an "IT"-vs-"it" 1-vs-1
    /// tie must not silently resolve to the lowercase spelling via
    /// plain string comparison). Returns `None` when the fast path
    /// doesn't apply (mirrors `summarize_text_column`'s own scope),
    /// so the Python caller falls back to its own pure-Python version
    /// unchanged in that case.
    ///
    /// Added after profiling `CLEAN()`'s plan-building step showed
    /// this exact computation — previously done in a second pure-
    /// Python pass over the column, AFTER `summarize_text_column`
    /// already did the classification pass — as the single largest
    /// remaining cost (over 1 second of a ~1.09s call on a 1M-row
    /// column with near-duplicate categories present). This Rust
    /// version does it in one pass, no Python-level `.strip()`/
    /// `.lower()`/dict-`setdefault` per element.
    #[pyfunction]
    fn build_category_merge_mapping(
        py: Python<'_>,
        values: PyObject,
    ) -> PyResult<Option<PyObject>> {
        let col = FastColumn::resolve(values.bind(py))?;
        let total = col.len();
        let mut groups: HashMap<String, Vec<String>> = HashMap::new();
        for i in 0..total {
            let s = match col.cell_text_exact_at(py, i) {
                Some(s) => s,
                None => return Ok(None), // not a text-bearing variant — caller falls back
            };
            let trimmed = s.trim();
            if trimmed.is_empty() {
                continue;
            }
            groups
                .entry(trimmed.to_lowercase())
                .or_default()
                .push(trimmed.to_string());
        }

        let dict = PyDict::new_bound(py);
        for variants in groups.values() {
            let distinct: Vec<&String> = {
                let mut seen: Vec<&String> = Vec::new();
                for v in variants {
                    if !seen.contains(&v) {
                        seen.push(v);
                    }
                }
                seen
            };
            if distinct.len() <= 1 {
                continue;
            }
            let mut vcounts: HashMap<&String, usize> = HashMap::new();
            for v in variants {
                *vcounts.entry(v).or_insert(0) += 1;
            }
            // Tie-break toward first-seen (lowest index in `distinct`,
            // which preserves first-occurrence order by construction
            // above) — matches the Python version's fixed tie-break
            // exactly: only a STRICTLY greater count displaces the
            // current canonical, so on an exact count tie the
            // earlier-seen (lower-index) spelling always wins, never
            // the lexicographically-largest string.
            let mut canonical = distinct[0];
            let mut canonical_count = vcounts[canonical];
            for v in distinct.iter().skip(1) {
                let c = vcounts[v];
                if c > canonical_count {
                    canonical = v;
                    canonical_count = c;
                }
            }
            for v in &distinct {
                if *v != canonical {
                    dict.set_item(v.as_str(), canonical.as_str())?;
                }
            }
        }
        Ok(Some(dict.into_py(py)))
    }
}
