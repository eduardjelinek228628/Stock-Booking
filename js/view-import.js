/* Probe Viewer – Import of probe test results (paste from Excel or upload .xlsx/.csv), preview, then write. */
(function () {
  'use strict';
  const PV = window.PV;
  const esc = PV.esc;
  const NOT_IMPORTABLE = new Set(['id', 'status', 'bookedFor', 'bookingDate', 'project', 'owner', 'department', 'box', 'mech', 'ele', 'mag', 'state', 'asOf', 'history']);

  function parseDelimited(text) {
    text = text.replace(/\r\n?/g, '\n').replace(/\n+$/, '');
    const first = text.split('\n')[0] || '';
    const delim = first.includes('\t') ? '\t' : (first.split(';').length > first.split(',').length ? ';' : ',');
    const rows = [];
    let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
        else cell += ch;
      } else if (ch === '"' && cell === '') q = true;
      else if (ch === delim) { row.push(cell); cell = ''; }
      else if (ch === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; }
      else cell += ch;
    }
    row.push(cell); rows.push(row);
    return rows;
  }
  function excelCell(v) {
    if (v == null) return '';
    if (v instanceof Date) return PV.isoDate(v);
    if (typeof v === 'object') {
      if (v.result !== undefined) return excelCell(v.result);
      if (v.richText) return v.richText.map((t) => t.text).join('');
      if (v.text !== undefined) return String(v.text);
      if (v.error) return v.error;
      return String(v);
    }
    return v;
  }
  async function parseXlsx(buf) {
    if (!window.ExcelJS) throw new Error('Excel reader not loaded (vendor/exceljs.min.js).');
    const wb = new window.ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    return wb.worksheets.map((ws) => {
      const rows = [];
      ws.eachRow({ includeEmpty: true }, (row, n) => { const vals = []; for (let c = 1; c <= ws.columnCount; c++) vals.push(excelCell(row.getCell(c).value)); rows[n - 1] = vals; });
      for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
      return { name: ws.name, rows };
    });
  }
  /** Pick the header row: first of the first 10 rows that contains an ID header. */
  function toTable(rows, cfg) {
    const idAliases = cfg.columns.id.map(PV.normHeader);
    let h = rows.findIndex((r, i) => i < 10 && r.some((c) => idAliases.includes(PV.normHeader(c))));
    if (h < 0) h = 0;
    const headers = (rows[h] || []).map((x) => PV.clean(x));
    const body = rows.slice(h + 1).filter((r) => r.some((c) => PV.clean(c) !== ''));
    return { headers, rows: body, headerRowNo: h + 1 };
  }

  function targetCols(f) {
    const m = f.stock;
    return m.headers.map((h, c) => ({ c, label: PV.clean(h) || PV.colLetter(m.c1 + c), role: Object.keys(m.roles).find((r) => m.roles[r] === c) }))
      .filter((x) => !NOT_IMPORTABLE.has(x.role) && PV.clean(m.headers[x.c]) !== '');
  }
  function autoMap(imp, f, cfg) {
    const m = f.stock, map = {};
    const tcols = targetCols(f);
    imp.headers.forEach((h, i) => {
      const n = PV.normHeader(h);
      if (!n) return;
      if (cfg.columns.id.map(PV.normHeader).includes(n)) { map[i] = 'id'; return; }
      let t = tcols.find((x) => PV.normHeader(m.headers[x.c]) === n);
      if (!t) { const role = Object.keys(cfg.columns).find((r) => cfg.columns[r].map(PV.normHeader).includes(n)); if (role) t = tcols.find((x) => x.role === role); }
      if (t) map[i] = String(t.c);
    });
    return map;
  }
  function numericTarget(f, c) {
    let n = 0, num = 0;
    for (const row of f.stock.rows) { if (row.kind !== 'data') continue; const v = row.values[c]; if (PV.clean(v) === '') continue; n++; if (PV.toNum(v) !== null) num++; }
    return n > 0 && num / n > 0.8;
  }
  function suggestFamily(store, imp, idCol) {
    if (idCol == null) return null;
    const ids = new Set(imp.rows.map((r) => PV.clean(r[idCol]).toUpperCase()).filter(Boolean));
    let best = null, bestN = 0;
    for (const f of store.families) {
      if (!f.stock) continue;
      const n = f.probes.filter((p) => p.stock && ids.has(PV.clean(p.id).toUpperCase())).length;
      if (n > bestN) { best = f.cfg.key; bestN = n; }
    }
    return best;
  }

  /** Build the preview: per row match + per cell action. */
  function preview(store, st) {
    const f = store.family(st.family);
    const idCol = Object.keys(st.mapping).find((k) => st.mapping[k] === 'id');
    if (idCol == null) return { error: 'Map one column to ID.' };
    const byId = new Map();
    for (const p of f.probes) { if (!p.stock) continue; const k = PV.clean(p.id).toUpperCase(); byId.set(k, byId.has(k) ? 'dup' : p); }
    const mapped = Object.keys(st.mapping).filter((k) => st.mapping[k] !== 'id' && st.mapping[k] !== '').map((k) => ({ i: +k, c: +st.mapping[k] }));
    const numeric = Object.fromEntries(mapped.map((m) => [m.c, numericTarget(f, m.c)]));
    const seenIds = new Map();
    const rows = st.parsed.rows.map((r) => {
      const id = PV.clean(r[idCol]);
      const out = { id, raw: r, cells: [], status: 'ok', note: '' };
      const k = id.toUpperCase();
      seenIds.set(k, (seenIds.get(k) || 0) + 1);
      if (!id) { out.status = 'skip'; out.note = 'no ID'; return out; }
      const p = byId.get(k);
      if (!p) { out.status = 'unknown'; out.note = 'ID not in ' + (f.cfg.label || f.cfg.key); return out; }
      if (p === 'dup') { out.status = 'dup'; out.note = 'ID appears more than once in the sheet'; return out; }
      out.probe = p;
      for (const m of mapped) {
        const ref = p.stock;
        const oldText = PV.displayCell(ref.row.values[m.c], ref.row.text[m.c], ref.row.nf[m.c]);
        const rawNew = r[m.i];
        let nv = typeof rawNew === 'number' ? rawNew : PV.clean(rawNew);
        const cell = { c: m.c, old: oldText, value: nv, action: 'same' };
        const fml = ref.row.formulas[m.c];
        if (nv === '' || nv == null) cell.action = 'empty';
        else if (typeof fml === 'string' && fml.startsWith('=')) cell.action = 'formula';
        else {
          if (numeric[m.c]) {
            const n = PV.toNum(nv);
            if (n === null) { cell.action = 'invalid'; out.cells.push(cell); continue; }
            nv = n; cell.value = n;
          }
          const same = numeric[m.c] ? PV.toNum(ref.row.values[m.c]) === nv : PV.clean(oldText) === PV.clean(nv);
          if (same) cell.action = 'same';
          else if (PV.clean(oldText) === '') cell.action = 'fill';
          else cell.action = st.overwrite ? 'overwrite' : 'blocked';
        }
        out.cells.push(cell);
      }
      if (!out.cells.some((c) => c.action === 'fill' || c.action === 'overwrite')) { out.status = out.cells.some((c) => c.action === 'blocked') ? 'blocked' : 'nochange'; }
      return out;
    });
    rows.forEach((r) => { if (r.status !== 'skip' && seenIds.get(r.id.toUpperCase()) > 1 && r.status !== 'unknown') { r.status = 'dupimport'; r.note = 'ID appears more than once in the import – skipped'; } });
    return { f, mapped, rows };
  }

  PV.views = PV.views || {};
  PV.views.import = function (root, app) {
    const store = app.store;
    const st = app.state.import;
    const fams = store.families.filter((f) => f.stock);
    if (!st.family) st.family = (fams[0] && fams[0].cfg.key) || null;
    const can = store.can('import');
    let html = '<div class="steps">';
    html += '<div class="card"><div class="card-h"><h2>Import test results</h2></div><div class="card-b grid">' +
      (can ? '' : '<div class="banner" style="border-radius:6px">' + (store.writesAllowed ? 'Your role cannot import (operations / importers only). You can still preview.' : 'Writes are switched off – preview only.') + '</div>') +
      '<div class="small muted">Copy the rows from the test Excel (header row included) and paste them below, or upload the .xlsx / .csv. The ID column is matched against the chosen probe type in Stock and Booking. Nothing is written before you confirm.</div>' +
      '<textarea class="input" id="paste" rows="6" placeholder="Paste rows here (Ctrl+V)…">' + esc(st.pasteText || '') + '</textarea>' +
      '<div class="toolbar" style="margin:0"><button class="btn primary" id="parsePaste">Read pasted rows</button><span class="muted">or</span><input type="file" id="file" accept=".xlsx,.csv,.txt"><span class="small muted">' + esc(st.fileName || '') + '</span></div>' +
      (st.sheets && st.sheets.length > 1 ? '<label class="field" style="max-width:300px"><span>Worksheet</span><select class="input" id="sheetSel">' + st.sheets.map((s, i) => '<option value="' + i + '"' + (i === st.sheetIdx ? ' selected' : '') + '>' + esc(s.name) + '</option>').join('') + '</select></label>' : '') +
      '</div></div>';

    if (st.parsed) {
      const f = store.family(st.family);
      const tcols = targetCols(f);
      html += '<div class="card"><div class="card-h"><h3>Columns</h3><span class="small muted">' + st.parsed.rows.length + ' rows read (header on row ' + st.parsed.headerRowNo + ')</span></div><div class="card-b grid">' +
        '<label class="field" style="max-width:320px"><span>Probe type (sheet in Stock and Booking)</span><select class="input" id="famSel">' + fams.map((x) => '<option value="' + esc(x.cfg.key) + '"' + (x.cfg.key === st.family ? ' selected' : '') + '>' + esc(x.cfg.label || x.cfg.key) + '</option>').join('') + '</select>' +
        (st.suggested && st.suggested !== st.family ? '<span class="small" style="color:var(--warning-ink)">IDs fit ' + esc(st.suggested) + ' better.</span>' : '') + '</label>' +
        '<div class="map-grid">' + st.parsed.headers.map((h, i) => '<div class="m"><b>' + esc(h || '(column ' + (i + 1) + ')') + '</b><select class="input" data-map="' + i + '"><option value="">— ignore —</option><option value="id"' + (st.mapping[i] === 'id' ? ' selected' : '') + '>ID (match key)</option>' +
          tcols.map((t) => '<option value="' + t.c + '"' + (st.mapping[i] === String(t.c) ? ' selected' : '') + '>→ ' + esc(t.label) + '</option>').join('') + '</select></div>').join('') + '</div>' +
        '<label class="small" style="display:flex;gap:6px;align-items:center"><input type="checkbox" id="ovw"' + (st.overwrite ? ' checked' : '') + '> Allow overwriting values that are already filled in (otherwise only empty cells are filled)</label>' +
        '</div></div>';

      const pv = preview(store, st);
      if (pv.error) html += '<div class="card"><div class="card-b" style="color:var(--critical-ink)">' + esc(pv.error) + '</div></div>';
      else {
        const sum = { ok: 0, unknown: 0, dup: 0, skip: 0, nochange: 0, blocked: 0, dupimport: 0 };
        let fills = 0, overw = 0, blocked = 0, invalid = 0;
        pv.rows.forEach((r) => { sum[r.status] = (sum[r.status] || 0) + 1; if (!['ok', 'blocked', 'nochange'].includes(r.status)) return; r.cells.forEach((c) => { if (c.action === 'fill') fills++; if (c.action === 'overwrite') overw++; if (c.action === 'blocked') blocked++; if (c.action === 'invalid') invalid++; }); });
        const writable = pv.rows.filter((r) => r.status === 'ok');
        const label = { ok: '<span class="flag ok">will update</span>', unknown: '<span class="flag err">unknown ID</span>', dup: '<span class="flag err">duplicate in sheet</span>', dupimport: '<span class="flag err">duplicate in import</span>', skip: '<span class="flag">no ID</span>', nochange: '<span class="flag ok">no change</span>', blocked: '<span class="flag">would overwrite</span>' };
        const cellCls = { fill: 'chg-new', overwrite: 'chg', blocked: 'chg', invalid: '', formula: '', same: '', empty: '' };
        const cellNote = { blocked: ' (kept – overwrite not allowed)', invalid: ' ⚠ not a number – skipped', formula: ' (formula column – skipped)' };
        html += '<div class="card"><div class="card-h"><h3>Preview</h3><span class="spacer" style="flex:1"></span>' +
          '<span class="small muted">' + writable.length + ' probes to update · ' + fills + ' empty cells filled · ' + overw + ' overwritten · ' + blocked + ' kept · ' + invalid + ' invalid · ' + (sum.unknown + sum.dup + sum.dupimport + sum.skip) + ' rows skipped</span></div>' +
          '<div class="table-wrap" style="max-height:480px"><table class="data"><thead><tr><th><span class="th-btn">ID</span></th><th><span class="th-btn">Result</span></th>' + pv.mapped.map((m) => '<th><span class="th-btn">' + esc(PV.clean(pv.f.stock.headers[m.c])) + '</span></th>').join('') + '</tr></thead><tbody>' +
          pv.rows.map((r) => '<tr class="' + (r.status === 'ok' ? '' : 'row-skip') + '" style="cursor:default"><td><b>' + esc(r.id || '–') + '</b></td><td>' + label[r.status] + (r.note ? ' <span class="small muted">' + esc(r.note) + '</span>' : '') + '</td>' +
            pv.mapped.map((m) => { const c = r.cells.find((x) => x.c === m.c); if (!c) return '<td class="faint">' + esc(PV.clean(r.raw[m.i])) + '</td>'; const show = c.action === 'fill' || c.action === 'overwrite' ? esc(c.old || '∅') + ' → <b>' + esc(c.value) + '</b>' : (c.action === 'same' ? esc(c.old) : esc(c.old || '∅') + (c.value !== '' ? ' / ' + esc(c.value) : '')); return '<td class="' + cellCls[c.action] + '" title="' + esc(c.action + (cellNote[c.action] || '')) + '">' + show + (cellNote[c.action] ? '<span class="small faint">' + cellNote[c.action] + '</span>' : '') + '</td>'; }).join('') + '</tr>').join('') +
          '</tbody></table></div><div class="drawer-f"><button class="btn primary" id="doImport"' + (writable.length && can ? '' : ' disabled') + '>Write ' + (fills + overw) + ' value(s) to ' + esc(pv.f.stock.sheetName) + '</button><span class="small muted">Each probe row is re-read first; rows changed in Excel meanwhile are skipped and reported.</span></div></div>';
        st._preview = pv;
      }
    }
    if (st.result) html += '<div class="card"><div class="card-h"><h3>Result</h3></div><div class="card-b">' + st.result + '</div></div>';
    html += '</div>';
    root.innerHTML = html;

    const parseRows = (rows) => {
      st.parsed = toTable(rows, store.cfg);
      st.result = null;
      const famNow = store.family(st.family);
      st.mapping = autoMap(st.parsed, famNow, store.cfg);
      const idCol = Object.keys(st.mapping).find((k) => st.mapping[k] === 'id');
      st.suggested = suggestFamily(store, st.parsed, idCol);
      if (st.suggested && st.suggested !== st.family) { st.family = st.suggested; st.mapping = autoMap(st.parsed, store.family(st.family), store.cfg); }
      app.render();
    };
    root.querySelector('#parsePaste').onclick = () => {
      const t = root.querySelector('#paste').value;
      st.pasteText = t; st.fileName = ''; st.sheets = null;
      if (!t.trim()) { PV.ui.toast('Paste some rows first.', true); return; }
      parseRows(parseDelimited(t));
    };
    root.querySelector('#file').onchange = async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      st.fileName = file.name; st.pasteText = '';
      try {
        if (/\.xlsx$/i.test(file.name)) { st.sheets = await parseXlsx(await file.arrayBuffer()); st.sheetIdx = 0; parseRows(st.sheets[0].rows); }
        else { st.sheets = null; parseRows(parseDelimited(await file.text())); }
      } catch (err) { PV.ui.toast('Could not read ' + file.name + ': ' + err.message, true); }
    };
    const ss = root.querySelector('#sheetSel'); if (ss) ss.onchange = () => { st.sheetIdx = +ss.value; parseRows(st.sheets[st.sheetIdx].rows); };
    const fs = root.querySelector('#famSel'); if (fs) fs.onchange = () => { st.family = fs.value; st.mapping = autoMap(st.parsed, store.family(st.family), store.cfg); app.render(); };
    root.querySelectorAll('[data-map]').forEach((s) => (s.onchange = () => {
      if (s.value === 'id') Object.keys(st.mapping).forEach((k) => { if (st.mapping[k] === 'id') delete st.mapping[k]; });
      if (s.value) st.mapping[s.dataset.map] = s.value; else delete st.mapping[s.dataset.map];
      app.render();
    }));
    const ov = root.querySelector('#ovw'); if (ov) ov.onchange = () => { st.overwrite = ov.checked; app.render(); };
    const di = root.querySelector('#doImport');
    if (di) di.onclick = async () => {
      const pv = st._preview;
      const items = pv.rows.filter((r) => r.status === 'ok').map((r) => ({ probe: r.probe, cells: r.cells.filter((c) => c.action === 'fill' || c.action === 'overwrite').map((c) => ({ col: c.c, value: c.value })) }));
      di.disabled = true;
      const res = await PV.ui.runAction(() => store.importResults(st.family, items), 'Import finished.');
      if (res) {
        st.result = '<p><b>' + res.written + '</b> value(s) written.</p>' + (res.failed.length ? '<p style="color:var(--critical-ink)">Not written:</p><ul>' + res.failed.map((x) => '<li>' + esc(x.id) + ': ' + esc(x.error) + '</li>').join('') + '</ul>' : '');
        st.parsed = null;
      }
      app.render();
    };
  };
  PV._import = { parseDelimited, toTable, preview, autoMap };
})();
