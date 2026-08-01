use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyFloat, PyInt, PyString};
use std::collections::HashMap;

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
            return CellValue::Num(f.value());
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
        match obj.str() {
            Ok(s) => CellValue::Text(s.to_string()),
            Err(_) => CellValue::Empty,
        }
    }
}


fn values_equal(a: &CellValue, b: &CellValue) -> bool {
    match (a, b) {
        (CellValue::Num(x), CellValue::Num(y)) => (x - y).abs() < 1e-9,
        (CellValue::Text(x), CellValue::Text(y)) => x.eq_ignore_ascii_case(y),
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
}
impl Eq for LookupKey {}
impl std::hash::Hash for LookupKey {
    fn hash<H: std::hash::Hasher>(&self, state: &mut H) {
        match self {
            LookupKey::Num(n) => n.to_bits().hash(state),
            LookupKey::Text(s) => s.hash(state),
        }
    }
}

fn cell_to_key(cv: &CellValue) -> Option<LookupKey> {
    match cv {
        CellValue::Num(n) => Some(LookupKey::Num(*n)),
        CellValue::Text(s) => Some(LookupKey::Text(s.to_lowercase())),
        CellValue::Empty => None,
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

fn parse_criteria(obj: &Bound<PyAny>) -> Criteria {
    let cv = CellValue::from_py(obj);
    match cv {
        CellValue::Num(n) => Criteria {
            op: Op::Eq,
            value: CellValue::Num(n),
            wildcard: None,
        },
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
                        return Criteria {
                            op: *op,
                            value: CellValue::Num(n),
                            wildcard: None,
                        };
                    }
                    if rest.contains('*') || rest.contains('?') {
                        return Criteria {
                            op: *op,
                            value: CellValue::Text(rest.to_string()),
                            wildcard: Some(rest.to_string()),
                        };
                    }
                    return Criteria {
                        op: *op,
                        value: CellValue::Text(rest.to_string()),
                        wildcard: None,
                    };
                }
            }
            if s.contains('*') || s.contains('?') {
                return Criteria {
                    op: Op::Eq,
                    value: CellValue::Text(s.clone()),
                    wildcard: Some(s),
                };
            }
            if let Ok(n) = s.trim().parse::<f64>() {
                return Criteria {
                    op: Op::Eq,
                    value: CellValue::Num(n),
                    wildcard: None,
                };
            }
            Criteria {
                op: Op::Eq,
                value: CellValue::Text(s),
                wildcard: None,
            }
        }
        CellValue::Empty => Criteria {
            op: Op::Eq,
            value: CellValue::Empty,
            wildcard: None,
        },
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
        (CellValue::Text(ct), CellValue::Text(vt)) => {
            let a = vt.to_lowercase();
            let b = ct.to_lowercase();
            match criteria.op {
                Op::Eq => a == b,
                Op::Ne => a != b,
                Op::Gt => a > b,
                Op::Ge => a >= b,
                Op::Lt => a < b,
                Op::Le => a <= b,
            }
        }
        (CellValue::Empty, CellValue::Empty) => matches!(criteria.op, Op::Eq),
        (CellValue::Empty, _) => matches!(criteria.op, Op::Ne),
        _ => matches!(criteria.op, Op::Ne),
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
        let crit = parse_criteria(criteria.bind(py));
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
        let crit = parse_criteria(criteria.bind(py));
        let target = sum_range
            .unwrap_or_else(|| range.iter().map(|v| v.clone_ref(py)).collect());
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
        let crit = parse_criteria(criteria.bind(py));
        let target = average_range
            .unwrap_or_else(|| range.iter().map(|v| v.clone_ref(py)).collect());
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
            .map(|(r, c)| (r, parse_criteria(c.bind(py))))
            .collect();
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
            .map(|(r, c)| (r, parse_criteria(c.bind(py))))
            .collect();
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
    ) -> Vec<i64> {
        let parsed: Vec<Criteria> = criteria_list.iter().map(|c| parse_criteria(c.bind(py))).collect();
        let keys: Vec<Option<LookupKey>> = parsed.iter().map(criteria_key).collect();

        if keys.iter().all(|k| k.is_some()) {
            let mut freq: HashMap<LookupKey, i64> = HashMap::new();
            for v in &range {
                if let Some(k) = cell_to_key(&CellValue::from_py(v.bind(py))) {
                    *freq.entry(k).or_insert(0) += 1;
                }
            }
            return keys
                .into_iter()
                .map(|k| *freq.get(&k.unwrap()).unwrap_or(&0))
                .collect();
        }

        parsed
            .iter()
            .map(|crit| {
                range
                    .iter()
                    .filter(|v| matches(&CellValue::from_py(v.bind(py)), crit))
                    .count() as i64
            })
            .collect()
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
        let parsed: Vec<Criteria> = criteria_list.iter().map(|c| parse_criteria(c.bind(py))).collect();
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
        let target = average_range.unwrap_or_else(|| range.iter().map(|v| v.clone_ref(py)).collect());
        if target.len() != range.len() {
            return Err(PyValueError::new_err(
                "AVERAGEIF: average_range must be the same length as range",
            ));
        }
        let parsed: Vec<Criteria> = criteria_list.iter().map(|c| parse_criteria(c.bind(py))).collect();
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
            .map(|(r, cl)| (r, cl.iter().map(|c| parse_criteria(c.bind(py))).collect()))
            .collect();

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
            .map(|(r, cl)| (r, cl.iter().map(|c| parse_criteria(c.bind(py))).collect()))
            .collect();

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
            .map(|(r, cl)| (r, cl.iter().map(|c| parse_criteria(c.bind(py))).collect()))
            .collect();

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
}
