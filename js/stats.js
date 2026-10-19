/* Deterministic parsing, type inference and statistics. */
(function (root) {
  const MISSING = new Set(['', 'na', 'n/a', 'nan', 'null', 'none', 'nil', '-', '--', '#n/a', 'undefined']);
  const isMissing = v => v === null || v === undefined || (typeof v === 'number' && !isFinite(v)) ||
    (typeof v === 'string' && MISSING.has(v.trim().toLowerCase()));

  function parseNum(v) {
    if (typeof v === 'number') return isFinite(v) ? v : null;
    if (typeof v !== 'string') return null;
    let s = v.trim();
    if (!s) return null;
    let neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    s = s.replace(/^(?:[A-Z]{3}\s)?[\p{Sc}]*\s?/u, '').replace(/\s?%$/, '').replace(/\s?[\p{Sc}]$/u, '');
    if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) s = s.replace(/,/g, '');
    if (!/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s)) return null;
    const n = parseFloat(s);
    return isFinite(n) ? (neg ? -n : n) : null;
  }

  const MONTHS = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  // Returns {re, order} hint for slash/dash/dot dates; ordering resolved per column.
  function splitDate(s) {
    let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T].*)?$/);
    if (m) return { y: +m[1], mo: +m[2], d: +m[3], iso: true };
    m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2}|\d{4})(?:[ T].*)?$/);
    if (m) return { a: +m[1], b: +m[2], y: m[3].length === 2 ? 2000 + +m[3] : +m[3] };
    return null;
  }
  function mkDate(y, mo, d) {
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    const dt = new Date(Date.UTC(y, mo - 1, d));
    return dt.getUTCMonth() === mo - 1 && dt.getUTCDate() === d ? dt : null;
  }
  function parseNamedDate(s) {
    const m = s.match(/^(?:[A-Za-z]{3,9},?\s+)?(\d{1,2})(?:st|nd|rd|th)?[\s-]+([A-Za-z]{3,9})\.?,?[\s-]+(\d{4})$/) ||
      null;
    if (m) { const mo = MONTHS.indexOf(m[2].slice(0, 3).toLowerCase()) + 1; return mo ? mkDate(+m[3], mo, +m[1]) : null; }
    const m2 = s.match(/^([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})$/);
    if (m2) { const mo = MONTHS.indexOf(m2[1].slice(0, 3).toLowerCase()) + 1; return mo ? mkDate(+m2[3], mo, +m2[2]) : null; }
    return null;
  }

  // Infer column type from raw values; returns {type, values(parsed), dateOrder}
  function inferColumn(name, raw) {
    const present = raw.filter(v => !isMissing(v));
    const n = present.length;
    const res = { name, type: 'empty', values: raw.map(() => null), invalid: 0 };
    if (!n) return res;
    const THRESH = 0.95;

    const dateObjs = present.filter(v => v instanceof Date).length;
    if (dateObjs / n >= THRESH) {
      res.type = 'date';
      res.values = raw.map(v => (v instanceof Date && !isNaN(v)) ? v : null);
      res.invalid = n - res.values.filter(Boolean).length;
      return res;
    }
    const strs = raw.map(v => isMissing(v) ? null : (v instanceof Date ? v.toISOString() : (typeof v === 'string' ? v.trim() : v)));

    const nums = strs.map(v => v === null ? null : parseNum(v));
    const numCount = nums.filter(v => v !== null).length;
    const lowerStr = strs.filter(v => typeof v === 'string').map(v => v.toLowerCase());

    // Dates (strings only, never plain numbers)
    const parts = strs.map(v => typeof v === 'string' ? splitDate(v) : null);
    const slash = parts.filter(p => p && !p.iso);
    let order = 'dmy';
    if (slash.length) {
      if (slash.some(p => p.a > 12)) order = 'dmy';
      else if (slash.some(p => p.b > 12)) order = 'mdy';
    }
    const dates = strs.map((v, i) => {
      if (typeof v !== 'string') return null;
      const p = parts[i];
      if (p) {
        if (p.iso) { const d = mkDate(p.y, p.mo, p.d); if (d && /[T ]\d/.test(v)) { const t = Date.parse(v.replace(' ', 'T')); return isNaN(t) ? d : new Date(t); } return d; }
        return order === 'dmy' ? mkDate(p.y, p.b, p.a) : mkDate(p.y, p.a, p.b);
      }
      return parseNamedDate(v);
    });
    const dateCount = dates.filter(Boolean).length;
    if (dateCount / n >= THRESH && numCount / n < THRESH) {
      res.type = 'date'; res.values = dates; res.invalid = n - dateCount; res.dateOrder = slash.length ? order : null;
      return res;
    }
    // Numeric
    if (numCount / n >= THRESH) {
      res.type = 'numeric'; res.values = nums; res.invalid = n - numCount;
      const vals = nums.filter(v => v !== null);
      const uniq = new Set(vals).size;
      const allInt = vals.every(Number.isInteger);
      if (allInt && uniq === vals.length && vals.length > 20 && /(^|[\s_-])(id|code|no|number|ref|key)$|^id$/i.test(name)) res.type = 'id';
      return res;
    }
    // Boolean
    const bset = new Set(lowerStr);
    if (strs.every(v => v === null || typeof v === 'boolean' || typeof v === 'string') && bset.size <= 2 && [...bset].every(x => ['true','false','yes','no','y','n','0','1'].includes(x))) {
      res.type = 'boolean'; res.values = strs.map(v => v === null ? null : String(v)); return res;
    }
    res.values = strs.map(v => v === null ? null : String(v));
    const uniq = new Set(res.values.filter(v => v !== null)).size;
    res.type = (uniq > 50 && uniq / n > 0.5) ? 'text' : 'categorical';
    if (uniq === n && n > 20) res.type = /id|code|ref|key|number|no$/i.test(name) ? 'id' : 'text';
    return res;
  }

  const sum = a => { let s = 0, c = 0; for (const x of a) { const y = x - c, t = s + y; c = (t - s) - y; s = t; } return s; }; // Kahan
  const mean = a => sum(a) / a.length;
  function quantile(sorted, p) {
    if (!sorted.length) return NaN;
    const h = (sorted.length - 1) * p, lo = Math.floor(h), hi = Math.ceil(h);
    return sorted[lo] + (h - lo) * (sorted[hi] - sorted[lo]);
  }
  function describe(vals) {
    const a = vals.filter(v => v !== null && isFinite(v)), n = a.length;
    if (!n) return null;
    const s = [...a].sort((x, y) => x - y), m = mean(a);
    const varr = n > 1 ? sum(a.map(x => (x - m) ** 2)) / (n - 1) : 0, sd = Math.sqrt(varr);
    const q1 = quantile(s, .25), q3 = quantile(s, .75), iqr = q3 - q1;
    let skew = 0;
    if (n > 2 && sd > 0) { const m3 = sum(a.map(x => ((x - m) / sd) ** 3)); skew = (n / ((n - 1) * (n - 2))) * m3; }
    const lo = q1 - 1.5 * iqr, hi = q3 + 1.5 * iqr;
    return { n, sum: sum(a), mean: m, median: quantile(s, .5), sd, min: s[0], max: s[n - 1], q1, q3, iqr, skew,
      cv: m !== 0 ? sd / Math.abs(m) : NaN, lo, hi, outliers: a.filter(x => x < lo || x > hi), sorted: s };
  }
  function pearson(x, y) {
    const px = [], py = [];
    for (let i = 0; i < x.length; i++) if (x[i] !== null && y[i] !== null) { px.push(x[i]); py.push(y[i]); }
    const n = px.length; if (n < 3) return { r: NaN, n };
    const mx = mean(px), my = mean(py);
    let sxy = 0, sxx = 0, syy = 0;
    for (let i = 0; i < n; i++) { const dx = px[i] - mx, dy = py[i] - my; sxy += dx * dy; sxx += dx * dx; syy += dy * dy; }
    if (sxx === 0 || syy === 0) return { r: NaN, n };
    return { r: Math.max(-1, Math.min(1, sxy / Math.sqrt(sxx * syy))), n };
  }
  function groupBy(cat, num) {
    const g = new Map();
    for (let i = 0; i < cat.length; i++) { if (cat[i] === null || num[i] === null) continue; if (!g.has(cat[i])) g.set(cat[i], []); g.get(cat[i]).push(num[i]); }
    return g;
  }
  // Share of numeric variance explained by a categorical column (eta squared)
  function eta2(cat, num) {
    const g = groupBy(cat, num), all = [].concat(...g.values());
    if (all.length < 3 || g.size < 2) return 0;
    const m = mean(all); let ssb = 0, sst = 0;
    for (const v of g.values()) ssb += v.length * (mean(v) - m) ** 2;
    for (const x of all) sst += (x - m) ** 2;
    return sst ? ssb / sst : 0;
  }
  function countBy(vals) {
    const m = new Map(); for (const v of vals) if (v !== null) m.set(v, (m.get(v) || 0) + 1);
    return [...m.entries()].sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])));
  }
  function histogram(s, d) {
    const n = s.length; if (!n) return { labels: [], counts: [], edges: [] };
    const min = s[0], max = s[n - 1];
    if (min === max) return { labels: [String(min)], counts: [n], edges: [min, max] };
    let w = d.iqr > 0 ? 2 * d.iqr / Math.cbrt(n) : (max - min) / (Math.ceil(Math.log2(n)) + 1);
    let bins = Math.ceil((max - min) / w);
    bins = Math.max(5, Math.min(30, bins));
    w = (max - min) / bins;
    const counts = new Array(bins).fill(0);
    for (const x of s) counts[Math.min(bins - 1, Math.floor((x - min) / w))]++;
    const edges = counts.map((_, i) => min + i * w); edges.push(max);
    return { counts, edges, labels: counts.map((_, i) => fmt(edges[i]) + '–' + fmt(edges[i + 1])) };
  }
  function fmt(x, dp) {
    if (x === null || x === undefined || !isFinite(x)) return '—';
    const ax = Math.abs(x);
    if (ax >= 1e9) return (x / 1e9).toFixed(2) + 'B';
    if (ax >= 1e6) return (x / 1e6).toFixed(2) + 'M';
    if (dp !== undefined) return x.toLocaleString(undefined, { maximumFractionDigits: dp, minimumFractionDigits: dp });
    if (Number.isInteger(x)) return x.toLocaleString();
    return x.toLocaleString(undefined, { maximumFractionDigits: ax < 1 ? 4 : ax < 100 ? 2 : 1 });
  }

  root.Stats = { isMissing, parseNum, inferColumn, sum, mean, quantile, describe, pearson, groupBy, eta2, countBy, histogram, fmt };
  if (typeof module !== 'undefined') module.exports = root.Stats;
})(typeof window !== 'undefined' ? window : globalThis);
