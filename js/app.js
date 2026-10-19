(function () {
  const S = window.Stats, F = S.fmt;
  const $ = id => document.getElementById(id);
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const PAL = ['#3b5bdb', '#2f9e6a', '#e8890c', '#d6414f', '#7048e8', '#0c8599', '#c2255c', '#5c940d', '#f08c00', '#495057'];
  let charts = [];

  /* ---------- Loading ---------- */
  const showError = m => { const e = $('error'); e.textContent = m; e.classList.toggle('hidden', !m); };

  function toTable(rowsObjs, fileName) {
    const cols = [];
    const seen = new Set();
    for (const r of rowsObjs) for (const k of Object.keys(r)) if (!seen.has(k)) { seen.add(k); cols.push(k); }
    return { fileName, cols, rows: rowsObjs };
  }
  function flatten(o, p = '', out = {}) {
    for (const [k, v] of Object.entries(o)) {
      const key = p ? p + '.' + k : k;
      if (v && typeof v === 'object' && !Array.isArray(v) && !(v instanceof Date)) flatten(v, key, out);
      else out[key] = Array.isArray(v) ? JSON.stringify(v) : v;
    }
    return out;
  }
  async function readFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (['csv', 'tsv', 'txt'].includes(ext)) {
      const text = await file.text();
      const r = Papa.parse(text.replace(/^\uFEFF/, ''), { header: true, skipEmptyLines: 'greedy', dynamicTyping: false, transformHeader: h => h.trim() });
      if (!r.data.length) throw new Error('No data rows were found in this file.');
      return toTable(r.data, file.name);
    }
    if (['xlsx', 'xls'].includes(ext)) {
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
      if (aoa.length < 2) throw new Error('The first worksheet has no data rows.');
      const used = new Set();
      const head = aoa[0].map((h, i) => { let n = (h == null || h === '' ? 'Column ' + (i + 1) : String(h).trim()); while (used.has(n)) n += '_'; used.add(n); return n; });
      return toTable(aoa.slice(1).map(r => Object.fromEntries(head.map((h, i) => [h, r[i] ?? null]))), file.name + ' (' + wb.SheetNames[0] + ')');
    }
    if (ext === 'json') {
      let j = JSON.parse(await file.text());
      if (!Array.isArray(j)) { const arr = Object.values(j).find(Array.isArray); if (!arr) throw new Error('JSON must contain an array of records.'); j = arr; }
      return toTable(j.map(r => (r && typeof r === 'object') ? flatten(r) : { value: r }), file.name);
    }
    throw new Error('Unsupported file type ".' + ext + '". Use CSV, TSV, XLSX, XLS or JSON.');
  }

  /* ---------- Analysis model ---------- */
  function analyse(t) {
    const cols = t.cols.map(name => S.inferColumn(name, t.rows.map(r => r[name])));
    const nRows = t.rows.length;
    cols.forEach(c => {
      c.n = c.values.filter(v => v !== null).length;
      c.missing = nRows - c.n;
      c.unique = new Set(c.values.filter(v => v !== null).map(v => v instanceof Date ? +v : v)).size;
      if (c.type === 'numeric') c.d = S.describe(c.values);
    });
    const seen = new Set(); let dups = 0;
    for (const r of t.rows) { const k = JSON.stringify(t.cols.map(c => r[c] instanceof Date ? +r[c] : r[c])); if (seen.has(k)) dups++; else seen.add(k); }
    return { t, cols, nRows, dups,
      num: cols.filter(c => c.type === 'numeric'),
      cat: cols.filter(c => (c.type === 'categorical' || c.type === 'boolean') && c.unique >= 2 && c.unique <= 25),
      date: cols.filter(c => c.type === 'date') };
  }

  /* ---------- Rendering helpers ---------- */
  const card = (n, title, sub, body) => `<section class="card" id="r${n}"><h2><span class="n">${n}</span>${title}</h2><p class="sub">${sub}</p>${body}</section>`;
  const kpi = (v, l) => `<div class="kpi"><b>${v}</b><span>${l}</span></div>`;
  const table = (head, rows, cls = '') => `<div class="tw"><table class="${cls}"><thead><tr>${head.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(r => `<tr>${r.map(c => `<td>${c}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
  const pct = (a, b) => b ? (100 * a / b).toFixed(1) + '%' : '0%';
  let pending = [];
  function chartBox(title, cfg) {
    const id = 'c' + pending.length; pending.push({ id, cfg });
    return `<div class="chart"><h4>${esc(title)}</h4><div class="box"><canvas id="${id}"></canvas></div></div>`;
  }
  const baseOpts = { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } };
  const bar = (labels, data, color = PAL[0], horizontal = false, label = '') => ({ type: 'bar', data: { labels, datasets: [{ label, data, backgroundColor: color, borderRadius: 3 }] }, options: { ...baseOpts, indexAxis: horizontal ? 'y' : 'x' } });
  const heat = r => { if (isNaN(r)) return '#eee'; const a = Math.abs(r); return r >= 0 ? `rgba(59,91,219,${0.1 + a * .85})` : `rgba(214,65,79,${0.1 + a * .85})`; };
  const trunc = (s, n = 22) => { s = String(s); return s.length > n ? s.slice(0, n - 1) + '…' : s; };

  /* ---------- 10 reports ---------- */
  function report1(m) { // Executive summary
    const cells = m.nRows * m.cols.length, miss = m.cols.reduce((a, c) => a + c.missing, 0);
    const comp = cells ? 100 * (1 - miss / cells) : 100;
    const types = {}; m.cols.forEach(c => types[c.type] = (types[c.type] || 0) + 1);
    const cfg = { type: 'doughnut', data: { labels: Object.keys(types), datasets: [{ data: Object.values(types), backgroundColor: PAL }] }, options: { ...baseOpts, plugins: { legend: { display: true, position: 'right' } } } };
    return card(1, 'Executive Summary', 'High-level overview of the uploaded dataset.',
      `<div class="kpis">${kpi(m.nRows.toLocaleString(), 'Records')}${kpi(m.cols.length, 'Columns')}${kpi(comp.toFixed(1) + '%', 'Completeness')}${kpi(m.dups.toLocaleString(), 'Duplicate rows')}${kpi(m.num.length, 'Numeric fields')}${kpi(m.date.length, 'Date fields')}</div>
      <div class="grid">${chartBox('Column types', cfg)}<div><h4>Source</h4><p class="note">${esc(m.t.fileName)}</p><h4>Headline findings</h4><ul class="insights">${headline(m).map(x => `<li>${x}</li>`).join('')}</ul></div></div>`);
  }
  function headline(m) {
    const out = [`The dataset has <b>${m.nRows.toLocaleString()}</b> records across <b>${m.cols.length}</b> columns.`];
    const worst = [...m.cols].sort((a, b) => b.missing - a.missing)[0];
    if (worst && worst.missing) out.push(`Most incomplete column: <b>${esc(worst.name)}</b> (${pct(worst.missing, m.nRows)} missing).`);
    else out.push('No missing values detected.');
    if (m.dups) out.push(`<b>${m.dups}</b> exact duplicate row(s) found (${pct(m.dups, m.nRows)}).`);
    const top = topCorr(m)[0];
    if (top) out.push(`Strongest relationship: <b>${esc(top.a)}</b> vs <b>${esc(top.b)}</b> (r = ${top.r.toFixed(2)}).`);
    return out;
  }
  function report2(m) { // Data quality
    const rows = m.cols.map(c => {
      const flags = [];
      if (c.missing / m.nRows > .2) flags.push('<span class="bad">high missing</span>');
      if (c.unique <= 1 && c.n) flags.push('<span class="warn">constant</span>');
      if (c.invalid) flags.push(`<span class="warn">${c.invalid} unparseable</span>`);
      if (c.type === 'date' && c.dateOrder) flags.push(`<span class="tag">${c.dateOrder.toUpperCase()}</span>`);
      return [esc(c.name), `<span class="tag">${c.type}</span>`, c.n.toLocaleString(), c.missing.toLocaleString(), `<div class="bar"><i style="width:${100 * c.missing / m.nRows}%;background:${c.missing ? 'var(--warn)' : 'var(--good)'}"></i></div>`, pct(c.missing, m.nRows), flags.join(' ') || '<span class="good">ok</span>'];
    });
    const dm = m.cols.filter(c => c.type === 'date' && c.dateOrder === 'dmy' && m.t.rows.length);
    return card(2, 'Data Quality Audit', 'Completeness, validity and duplication per column. Missing markers: blank, NA, N/A, null, NaN, "-".',
      `<div class="kpis">${kpi(m.dups.toLocaleString(), 'Duplicate rows')}${kpi(m.cols.filter(c => c.missing).length, 'Columns with gaps')}${kpi(m.cols.reduce((a, c) => a + c.invalid, 0), 'Unparseable values')}</div>
      ${table(['Column', 'Type', 'Valid', 'Missing', 'Gap', 'Missing %', 'Flags'], rows)}
      ${m.cols.some(c => c.dateOrder) ? '<p class="note">Slash-separated dates are read as day/month/year unless the data proves month/day/year.</p>' : ''}`);
  }
  function report3(m) { // Column profile
    const rows = m.cols.map(c => {
      let ex = '';
      if (c.type === 'numeric') ex = `${F(c.d.min)} → ${F(c.d.max)}`;
      else if (c.type === 'date') { const v = c.values.filter(Boolean).map(Number); ex = v.length ? new Date(Math.min(...v)).toISOString().slice(0, 10) + ' → ' + new Date(Math.max(...v)).toISOString().slice(0, 10) : ''; }
      else { const t = S.countBy(c.values)[0]; ex = t ? 'top: ' + esc(trunc(t[0], 28)) : ''; }
      return [esc(c.name), `<span class="tag">${c.type}</span>`, c.unique.toLocaleString(), pct(c.unique, c.n), ex];
    });
    return card(3, 'Column Profile', 'Detected data type, cardinality and range of every field.', table(['Column', 'Type', 'Unique', 'Unique %', 'Range / top value'], rows));
  }
  function report4(m) { // Descriptive stats
    if (!m.num.length) return card(4, 'Descriptive Statistics', 'Summary statistics for numeric fields.', '<p class="note">No numeric columns detected.</p>');
    const rows = m.num.map(c => { const d = c.d; return [esc(c.name), d.n.toLocaleString(), F(d.sum), F(d.mean), F(d.median), F(d.sd), F(d.min), F(d.q1), F(d.q3), F(d.max), d.skew.toFixed(2)]; });
    return card(4, 'Descriptive Statistics', 'Count, sum, mean, median, sample standard deviation, quartiles (linear interpolation) and skewness.', table(['Field', 'Count', 'Sum', 'Mean', 'Median', 'Std Dev', 'Min', 'Q1', 'Q3', 'Max', 'Skew'], rows));
  }
  function report5(m) { // Distributions
    const cs = m.num.slice(0, 6);
    if (!cs.length) return card(5, 'Numeric Distributions', 'Histograms of numeric fields.', '<p class="note">No numeric columns detected.</p>');
    return card(5, 'Numeric Distributions', 'Histograms using Freedman–Diaconis bin width.', `<div class="grid">${cs.map((c, i) => { const h = S.histogram(c.d.sorted, c.d); return chartBox(c.name, bar(h.labels, h.counts, PAL[i % 10])); }).join('')}</div>`);
  }
  function topCorr(m) {
    const cs = m.num.filter(c => c.d.sd > 0).slice(0, 12), out = [];
    for (let i = 0; i < cs.length; i++) for (let j = i + 1; j < cs.length; j++) { const p = S.pearson(cs[i].values, cs[j].values); if (!isNaN(p.r)) out.push({ a: cs[i].name, b: cs[j].name, r: p.r, n: p.n, ca: cs[i], cb: cs[j] }); }
    return out.sort((x, y) => Math.abs(y.r) - Math.abs(x.r));
  }
  function report6(m) { // Correlations
    const cs = m.num.slice(0, 12);
    if (cs.length < 2) return card(6, 'Correlation Analysis', 'Pearson correlations between numeric fields.', '<p class="note">At least two numeric columns are required.</p>');
    const head = [''].concat(cs.map(c => trunc(c.name, 12)));
    const rows = cs.map(a => [esc(a.name)].concat(cs.map(b => { const r = a === b ? 1 : S.pearson(a.values, b.values).r; return `<span style="display:block;padding:4px;background:${heat(r)};color:${Math.abs(r) > .55 ? '#fff' : '#1c2333'};border-radius:4px">${isNaN(r) ? '—' : r.toFixed(2)}</span>`; })));
    const tc = topCorr(m)[0];
    const sc = tc ? chartBox(`${tc.a} vs ${tc.b} (r = ${tc.r.toFixed(2)}, n = ${tc.n})`, { type: 'scatter', data: { datasets: [{ data: tc.ca.values.map((v, i) => [v, tc.cb.values[i]]).filter(p => p[0] !== null && p[1] !== null).slice(0, 3000).map(p => ({ x: p[0], y: p[1] })), backgroundColor: 'rgba(59,91,219,.5)', pointRadius: 3 }] }, options: { ...baseOpts, scales: { x: { title: { display: true, text: tc.a } }, y: { title: { display: true, text: tc.b } } } } }) : '';
    return card(6, 'Correlation Analysis', 'Pearson r using pairwise-complete observations. Blue = positive, red = negative.', `<div class="grid"><div class="tw">${table(head, rows, 'heat')}</div>${sc}</div>`);
  }
  function report7(m) { // Categorical
    const cs = m.cat.slice(0, 4);
    if (!cs.length) return card(7, 'Categorical Breakdown', 'Frequency of categories.', '<p class="note">No suitable categorical columns (2–25 distinct values) detected.</p>');
    return card(7, 'Categorical Breakdown', 'Top 10 categories per field with share of records.', `<div class="grid">${cs.map((c, i) => { const t = S.countBy(c.values).slice(0, 10); return chartBox(`${c.name} (share of ${c.n.toLocaleString()} records)`, bar(t.map(x => trunc(x[0])), t.map(x => +(100 * x[1] / c.n).toFixed(2)), PAL[(i + 1) % 10], true, '% of records')); }).join('')}</div>`);
  }
  function bucketer(dates) {
    const v = dates.filter(Boolean).map(Number), span = (Math.max(...v) - Math.min(...v)) / 864e5;
    const g = span <= 62 ? 'day' : span <= 900 ? 'month' : 'year';
    return { g, key: d => { const s = d.toISOString(); return g === 'day' ? s.slice(0, 10) : g === 'month' ? s.slice(0, 7) : s.slice(0, 4); } };
  }
  function report8(m) { // Time
    const dc = m.date[0];
    if (!dc || dc.n < 2) return card(8, 'Time-Series Trends', 'Trends over time.', '<p class="note">No date column detected.</p>');
    const b = bucketer(dc.values), recs = new Map(), sums = m.num.slice(0, 2).map(() => new Map());
    dc.values.forEach((d, i) => { if (!d) return; const k = b.key(d); recs.set(k, (recs.get(k) || 0) + 1); m.num.slice(0, 2).forEach((c, j) => { if (c.values[i] !== null) sums[j].set(k, (sums[j].get(k) || 0) + c.values[i]); }); });
    const keys = [...recs.keys()].sort();
    const line = (lbl, mp, col) => ({ type: 'line', data: { labels: keys, datasets: [{ label: lbl, data: keys.map(k => mp.get(k) ?? 0), borderColor: col, backgroundColor: col + '22', fill: true, tension: .25, pointRadius: keys.length > 60 ? 0 : 3 }] }, options: baseOpts });
    const charts = [chartBox(`Records per ${b.g}`, line('Records', recs, PAL[0]))].concat(m.num.slice(0, 2).map((c, j) => chartBox(`Total ${c.name} per ${b.g}`, line(c.name, sums[j], PAL[j + 1]))));
    return card(8, 'Time-Series Trends', `Aggregated by ${b.g} using date column “${esc(dc.name)}”.`, `<div class="grid">${charts.join('')}</div>`);
  }
  function report9(m) { // Category vs numeric
    const pairs = [];
    m.cat.forEach(c => m.num.slice(0, 8).forEach(n => { const e = S.eta2(c.values, n.values); if (e > 0) pairs.push({ c, n, e }); }));
    pairs.sort((a, b) => b.e - a.e);
    if (!pairs.length) return card(9, 'Segment Performance', 'Compares numeric measures across categories.', '<p class="note">Needs at least one categorical and one numeric column.</p>');
    const sel = pairs.slice(0, 4);
    return card(9, 'Segment Performance', 'Average of a numeric field by category, for the combinations where the category explains the most variance (η²).',
      `<div class="grid">${sel.map((p, i) => { const g = [...S.groupBy(p.c.values, p.n.values)].map(([k, v]) => [k, S.mean(v), v.length]).sort((a, b) => b[1] - a[1]).slice(0, 12); return chartBox(`Mean ${p.n.name} by ${p.c.name} (η² = ${p.e.toFixed(2)})`, bar(g.map(x => trunc(x[0])), g.map(x => +x[1].toFixed(4)), PAL[(i + 2) % 10], true)); }).join('')}</div>`);
  }
  function report10(m) { // Outliers & insights
    const rows = m.num.map(c => { const d = c.d, o = d.outliers; return [esc(c.name), `${F(d.lo)} → ${F(d.hi)}`, o.length.toLocaleString(), pct(o.length, d.n), o.length ? F(o.reduce((a, x) => Math.abs(x) > Math.abs(a) ? x : a)) : '—']; });
    const ins = [];
    m.num.forEach(c => { const d = c.d; if (Math.abs(d.skew) > 1) ins.push(`<b>${esc(c.name)}</b> is ${d.skew > 0 ? 'right' : 'left'}-skewed (skew ${d.skew.toFixed(2)}); median (${F(d.median)}) differs from mean (${F(d.mean)}).`); if (d.outliers.length / d.n > .05) ins.push(`<b>${esc(c.name)}</b> has ${pct(d.outliers.length, d.n)} outliers by the 1.5×IQR rule.`); });
    topCorr(m).filter(x => Math.abs(x.r) >= .5).slice(0, 4).forEach(x => ins.push(`<b>${esc(x.a)}</b> and <b>${esc(x.b)}</b> are ${Math.abs(x.r) >= .8 ? 'strongly' : 'moderately'} ${x.r > 0 ? 'positively' : 'negatively'} correlated (r = ${x.r.toFixed(2)}).`));
    m.cat.slice(0, 3).forEach(c => { const t = S.countBy(c.values); if (t.length) ins.push(`<b>${esc(t[0][0])}</b> is the most common ${esc(c.name)} (${pct(t[0][1], c.n)} of records).`); });
    return card(10, 'Outliers & Key Insights', 'Outliers use Tukey fences (Q1 − 1.5·IQR, Q3 + 1.5·IQR). Insights are generated automatically from the statistics above.',
      `<div class="grid"><div>${rows.length ? table(['Field', 'Normal range', 'Outliers', '%', 'Most extreme'], rows) : '<p class="note">No numeric columns.</p>'}</div><div><h4>Insights</h4><ul class="insights">${(ins.length ? ins : ['No notable patterns exceeded the detection thresholds.']).map(x => `<li>${x}</li>`).join('')}</ul></div></div>`);
  }

  const REPORTS = [['Summary', report1], ['Quality', report2], ['Profile', report3], ['Statistics', report4], ['Distributions', report5], ['Correlations', report6], ['Categories', report7], ['Trends', report8], ['Segments', report9], ['Insights', report10]];

  function render(m) {
    charts.forEach(c => c.destroy()); charts = []; pending = [];
    $('report').innerHTML = REPORTS.map(([, f]) => f(m)).join('');
    $('nav').innerHTML = REPORTS.map(([n], i) => `<a href="#r${i + 1}">${i + 1}. ${n}</a>`).join('');
    pending.forEach(p => charts.push(new Chart($(p.id), p.cfg)));
    ['landing'].forEach(i => $(i).classList.add('hidden'));
    ['report', 'nav', 'actions'].forEach(i => $(i).classList.remove('hidden'));
    window.scrollTo(0, 0);
  }

  async function handle(file) {
    showError('');
    try {
      const t = await readFile(file);
      if (!t.rows.length) throw new Error('The file contains no data.');
      render(analyse(t));
    } catch (e) { console.error(e); showError(e.message || String(e)); }
  }

  $('file').addEventListener('change', e => e.target.files[0] && handle(e.target.files[0]));
  const drop = $('drop');
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => e.dataTransfer.files[0] && handle(e.dataTransfer.files[0]));
  $('printBtn').onclick = () => window.print();
  $('resetBtn').onclick = () => { $('report').classList.add('hidden'); $('nav').classList.add('hidden'); $('actions').classList.add('hidden'); $('landing').classList.remove('hidden'); $('file').value = ''; };
  $('sampleBtn').onclick = () => {
    const regions = ['Accra', 'Kumasi', 'Tamale', 'Takoradi'], prods = ['Laptop', 'Phone', 'Tablet', 'Monitor', 'Printer'];
    let seed = 42; const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
    let csv = 'OrderID,Date,Region,Product,Units,UnitPrice,Revenue,Satisfaction\n';
    for (let i = 1; i <= 400; i++) {
      const d = new Date(Date.UTC(2025, 0, 1) + Math.floor(rnd() * 540) * 864e5).toISOString().slice(0, 10);
      const u = 1 + Math.floor(rnd() * 12), p = [4500, 3200, 2100, 1500, 900][Math.floor(rnd() * 5)] + Math.round(rnd() * 200);
      csv += `${1000 + i},${d},${regions[Math.floor(rnd() * 4)]},${prods[Math.floor(rnd() * 5)]},${u},${p},${u * p},${rnd() < .06 ? '' : (2 + rnd() * 3).toFixed(1)}\n`;
    }
    handle(new File([csv], 'sample_sales.csv', { type: 'text/csv' }));
  };
})();
