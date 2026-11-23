/* Probe Viewer – Stock screen: one sub-tab per sheet, Excel-like filters, global ID search, probe drawer. */
(function () {
  'use strict';
  const PV = window.PV;
  const esc = PV.esc;
  const LIFE_ROLES = ['owner', 'department', 'mech', 'ele', 'mag', 'state', 'current', 'asOf', 'history'];

  function famLabel(f) { return f.cfg.label || f.cfg.key; }

  /** Column descriptors for a family table. */
  function columns(f, showLife) {
    const cols = [];
    const base = f.stock || f.life;
    if (!base) return cols;
    base.headers.forEach((h, c) => {
      const role = Object.keys(base.roles).find((r) => base.roles[r] === c);
      let label = PV.clean(h);
      if (!label) label = role ? role.charAt(0).toUpperCase() + role.slice(1) + ' (blank header)' : PV.colLetter(base.c1 + c);
      cols.push({
        key: (f.stock ? 's' : 'l') + c, label, unit: base.units[c] || '', role, life: !f.stock, isStatus: role === 'status',
        isState: ['mech', 'ele', 'mag', 'state', 'current'].includes(role) && !f.stock,
        text: (p) => { const ref = f.stock ? p.stock : p.life; return ref ? PV.displayCell(ref.row.values[c], ref.row.text[c], ref.row.nf[c]) : ''; },
        raw: (p) => { const ref = f.stock ? p.stock : p.life; return ref ? ref.row.values[c] : ''; },
        formula: (p) => { const ref = f.stock ? p.stock : p.life; const x = ref && ref.row.formulas[c]; return typeof x === 'string' && x.startsWith('='); },
      });
    });
    if (f.stock && f.life && showLife) {
      for (const role of LIFE_ROLES) {
        const c = f.life.roles[role];
        if (c == null) continue;
        cols.push({
          key: 'l' + c, label: PV.clean(f.life.headers[c]) || role, unit: '', role, life: true, isState: ['mech', 'ele', 'mag', 'state', 'current'].includes(role),
          text: (p) => (p.life ? PV.displayCell(p.life.row.values[c], p.life.row.text[c], p.life.row.nf[c]) : ''),
          raw: (p) => (p.life ? p.life.row.values[c] : ''), formula: () => false,
        });
      }
    }
    return cols;
  }
  function isNumericCol(col, rows) {
    let n = 0, num = 0;
    for (const p of rows) { const t = col.text(p); if (!t) continue; n++; if (PV.toNum(col.raw(p)) !== null || PV.toNum(t) !== null) num++; }
    return n > 0 && num / n > 0.8 && !col.isStatus && col.role !== 'id';
  }
  function cellHtml(col, p) {
    const t = col.text(p);
    if (col.isStatus) return PV.ui.statusChip(t);
    if (col.isState) return t ? PV.ui.stateChip(col.label, t) : '';
    return esc(t);
  }

  function rowsFor(f) { return f.stock ? f.probes.filter((p) => p.stock) : f.probes; }

  function applyFilters(rows, cols, filt, exceptKey) {
    return rows.filter((p) => cols.every((col) => {
      if (col.key === exceptKey) return true;
      const fl = filt[col.key];
      if (!fl) return true;
      const t = col.text(p);
      if (fl.values && !fl.values.has(t)) return false;
      if (fl.text && !t.toLowerCase().includes(fl.text.toLowerCase())) return false;
      if (fl.min != null || fl.max != null) {
        const n = PV.toNum(col.raw(p)) !== null ? PV.toNum(col.raw(p)) : PV.toNum(t);
        if (n === null) return false;
        if (fl.min != null && n < fl.min) return false;
        if (fl.max != null && n > fl.max) return false;
      }
      return true;
    }));
  }

  PV.views = PV.views || {};
  PV.views.stock = function (root, app) {
    const st = app.state.stock;
    const store = app.store;
    const fams = store.families.filter((f) => f.stock || f.life);
    if (!fams.length) { root.innerHTML = '<div class="empty">No probe sheets could be loaded. See Data checks.</div>'; return; }
    if (!st.family || !fams.find((f) => f.cfg.key === st.family)) st.family = fams[0].cfg.key;
    const f = store.family(st.family);
    const cols = columns(f, st.showLife);
    const allRows = rowsFor(f);
    const filt = (st.filters[f.cfg.key] = st.filters[f.cfg.key] || {});
    let rows = applyFilters(allRows, cols, filt);
    const sort = st.sort[f.cfg.key];
    if (sort) {
      const col = cols.find((c) => c.key === sort.key);
      if (col) {
        const numeric = isNumericCol(col, allRows);
        rows = rows.slice().sort((a, b) => {
          let x, y;
          if (numeric) { x = PV.toNum(col.raw(a)) ?? PV.toNum(col.text(a)); y = PV.toNum(col.raw(b)) ?? PV.toNum(col.text(b)); if (x == null) return 1; if (y == null) return -1; return (x - y) * sort.dir; }
          x = col.text(a); y = col.text(b); if (!x) return 1; if (!y) return -1;
          return x.localeCompare(y, undefined, { numeric: true }) * sort.dir;
        });
      }
    }
    const withStock = fams.filter((x) => x.stock), lifeOnly = fams.filter((x) => !x.stock);
    const tabBtn = (x) => '<button class="subtab" role="tab" aria-selected="' + (x.cfg.key === f.cfg.key) + '" data-fam="' + esc(x.cfg.key) + '">' + esc(famLabel(x)) + '<span class="n">' + rowsFor(x).length + '</span></button>';
    const nFilters = Object.keys(filt).length;
    const source = f.stock ? 'Stock and Booking › ' + f.stock.sheetName + (f.stock.tableName ? ' (Table ' + f.stock.tableName + ')' : '') : 'Probe Lifetime › ' + f.life.sheetName + (f.life.isTable ? '' : ' (plain range – read-only)');

    root.innerHTML =
      '<div class="toolbar"><input class="input" id="q" type="search" placeholder="Search probe ID in all types…" value="' + esc(st.q) + '" style="width:280px" aria-label="Search probe ID">' +
      '<span class="spacer" style="flex:1"></span>' +
      (f.stock && f.life ? '<label class="small muted" style="display:flex;gap:6px;align-items:center"><input type="checkbox" id="showLife"' + (st.showLife ? ' checked' : '') + '> Show Lifetime columns</label>' : '') +
      '<button class="btn" id="clearF"' + (nFilters || sort ? '' : ' disabled') + '>Clear filters' + (nFilters ? ' (' + nFilters + ')' : '') + '</button></div>' +
      '<div id="searchRes"></div>' +
      '<div class="subtabs" role="tablist"><span class="subtab-group">Stock and Booking</span>' + withStock.map(tabBtn).join('') +
      (lifeOnly.length ? '<span class="subtab-group">Lifetime only</span>' + lifeOnly.map(tabBtn).join('') : '') + '</div>' +
      '<div class="small muted" style="margin:-4px 0 8px">' + esc(source) + ' · showing ' + rows.length + ' of ' + allRows.length + '</div>' +
      '<div class="table-wrap"><table class="data"><thead><tr>' +
      cols.map((c) => '<th class="' + (c.life && f.stock ? 'grp-life ' : '') + (filt[c.key] || (sort && sort.key === c.key) ? 'filtered' : '') + '"><button class="th-btn" data-col="' + c.key + '">' + esc(c.label) + (c.unit ? ' <span class="unit">' + esc(c.unit) + '</span>' : '') +
        '<span class="ficon">' + (sort && sort.key === c.key ? (sort.dir > 0 ? '▲' : '▼') : '') + (filt[c.key] ? '⏷' : '▾') + '</span></button></th>').join('') +
      '</tr></thead><tbody>' +
      rows.map((p) => '<tr data-uid="' + esc(p.uid) + '">' + cols.map((c) => {
        const numeric = typeof c.raw(p) === 'number' && !c.isStatus && c.role !== 'id';
        return '<td class="' + (numeric ? 'num ' : '') + (c.life && f.stock ? 'life ' : '') + (c.formula(p) ? 'formula' : '') + '" title="' + esc(c.text(p)) + '">' + cellHtml(c, p) + '</td>';
      }).join('') + '</tr>').join('') +
      '</tbody></table>' + (rows.length ? '' : '<div class="empty">No probes match the filters.</div>') + '</div>';

    root.querySelectorAll('[data-fam]').forEach((b) => (b.onclick = () => { st.family = b.dataset.fam; app.render(); }));
    root.querySelector('#clearF').onclick = () => { st.filters[f.cfg.key] = {}; delete st.sort[f.cfg.key]; app.render(); };
    const sl = root.querySelector('#showLife'); if (sl) sl.onchange = () => { st.showLife = sl.checked; app.render(); };
    root.querySelector('tbody').onclick = (e) => { const tr = e.target.closest('tr[data-uid]'); if (tr) PV.openProbe(store.byUid.get(tr.dataset.uid), app); };
    root.querySelectorAll('.th-btn').forEach((b) => (b.onclick = () => openFilter(b, cols.find((c) => c.key === b.dataset.col), allRows, cols, filt, app, f)));
    const q = root.querySelector('#q');
    const doSearch = () => { st.q = q.value; renderSearch(root.querySelector('#searchRes'), store, st.q, app); };
    q.oninput = doSearch; doSearch();
  };

  function renderSearch(host, store, q, app) {
    q = PV.clean(q).toUpperCase();
    if (!q) { host.innerHTML = ''; return; }
    const hits = store.probes.filter((p) => PV.clean(p.id).toUpperCase().includes(q)).slice(0, 50);
    host.innerHTML = '<div class="card" style="margin-bottom:12px"><div class="card-h"><h3>' + hits.length + (hits.length === 50 ? '+' : '') + ' probe(s) matching "' + esc(q) + '"</h3></div>' +
      (hits.length ? '<table class="plain"><tr><th>ID</th><th>Type</th><th>Status</th><th>Booked for / Owner</th><th>In</th></tr>' + hits.map((p) =>
        '<tr data-uid="' + esc(p.uid) + '" style="cursor:pointer"><td><b>' + esc(PV.logId(p)) + '</b></td><td>' + esc(p.family.label || p.family.key) + '</td><td>' + PV.ui.statusChip(PV.ptext(p, 'status')) + '</td><td>' + esc(PV.ptext(p, 'bookedFor') || PV.ptext(p, 'owner')) + '</td><td class="small muted">' + (p.stock ? 'Stock' : '') + (p.stock && p.life ? ' + ' : '') + (p.life ? 'Lifetime' : '') + '</td></tr>').join('') + '</table>' : '<div class="empty">No probe with that ID.</div>') + '</div>';
    host.querySelectorAll('tr[data-uid]').forEach((tr) => (tr.onclick = () => PV.openProbe(store.byUid.get(tr.dataset.uid), app)));
  }

  function openFilter(btn, col, allRows, cols, filt, app, f) {
    document.querySelectorAll('.filter-pop').forEach((n) => n.remove());
    const st = app.state.stock;
    const base = applyFilters(allRows, cols, filt, col.key);
    const counts = new Map();
    base.forEach((p) => { const t = col.text(p); counts.set(t, (counts.get(t) || 0) + 1); });
    const numeric = isNumericCol(col, allRows);
    const vals = [...counts.keys()].sort((a, b) => (a === '' ? 1 : b === '' ? -1 : numeric ? (PV.toNum(a) ?? 0) - (PV.toNum(b) ?? 0) : a.localeCompare(b, undefined, { numeric: true })));
    const cur = filt[col.key] || {};
    const pop = document.createElement('div');
    pop.className = 'filter-pop';
    pop.innerHTML =
      '<div class="row"><button class="btn sm" data-sort="1">Sort ' + (numeric ? '1→9' : 'A→Z') + '</button><button class="btn sm" data-sort="-1">Sort ' + (numeric ? '9→1' : 'Z→A') + '</button></div>' +
      '<input class="input" data-text placeholder="Contains…" value="' + esc(cur.text || '') + '">' +
      (numeric ? '<div class="row"><input class="input" data-min placeholder="min" style="width:50%" value="' + (cur.min ?? '') + '"><input class="input" data-max placeholder="max" style="width:50%" value="' + (cur.max ?? '') + '"></div>' : '') +
      '<div class="vals"><label><input type="checkbox" data-all' + (!cur.values ? ' checked' : '') + '> <b>(Select all)</b></label>' +
      vals.map((v, i) => '<label><input type="checkbox" data-v="' + i + '"' + (!cur.values || cur.values.has(v) ? ' checked' : '') + '> ' + (v === '' ? '<i>(Blanks)</i>' : esc(v)) + '<span class="c">' + counts.get(v) + '</span></label>').join('') + '</div>' +
      '<div class="row" style="justify-content:flex-end"><button class="btn sm" data-clear>Clear</button><button class="btn sm primary" data-ok>Apply</button></div>';
    document.body.appendChild(pop);
    const r = btn.getBoundingClientRect();
    pop.style.left = Math.min(r.left, window.innerWidth - 270) + 'px';
    pop.style.top = Math.min(r.bottom + 4, window.innerHeight - pop.offsetHeight - 8) + 'px';
    const boxes = [...pop.querySelectorAll('[data-v]')];
    const all = pop.querySelector('[data-all]');
    all.onchange = () => boxes.forEach((b) => (b.checked = all.checked));
    boxes.forEach((b) => (b.onchange = () => { all.checked = boxes.every((x) => x.checked); }));
    pop.querySelectorAll('[data-sort]').forEach((b) => (b.onclick = () => { st.sort[f.cfg.key] = { key: col.key, dir: +b.dataset.sort }; pop.remove(); app.render(); }));
    pop.querySelector('[data-clear]').onclick = () => { delete filt[col.key]; if (st.sort[f.cfg.key] && st.sort[f.cfg.key].key === col.key) delete st.sort[f.cfg.key]; pop.remove(); app.render(); };
    pop.querySelector('[data-ok]').onclick = () => {
      const fl = {};
      const sel = boxes.filter((b) => b.checked).map((b) => vals[+b.dataset.v]);
      if (sel.length !== vals.length) fl.values = new Set(sel);
      const t = pop.querySelector('[data-text]').value.trim(); if (t) fl.text = t;
      if (numeric) {
        const mn = PV.toNum(pop.querySelector('[data-min]').value), mx = PV.toNum(pop.querySelector('[data-max]').value);
        if (mn !== null) fl.min = mn; if (mx !== null) fl.max = mx;
      }
      if (Object.keys(fl).length) filt[col.key] = fl; else delete filt[col.key];
      pop.remove(); app.render();
    };
    pop.querySelector('[data-text]').onkeydown = (e) => { if (e.key === 'Enter') pop.querySelector('[data-ok]').click(); };
    setTimeout(() => {
      const off = (e) => { if (!pop.contains(e.target)) { pop.remove(); document.removeEventListener('mousedown', off); } };
      document.addEventListener('mousedown', off);
    });
  }

  // ---------------- probe drawer (desktop) ----------------
  function kv(ref, title) {
    if (!ref) return '';
    const m = ref.model;
    const items = m.headers.map((h, c) => {
      const label = PV.clean(h) || (m.roles.status === c ? 'Status' : PV.colLetter(m.c1 + c));
      const t = PV.displayCell(ref.row.values[c], ref.row.text[c], ref.row.nf[c]);
      return '<dt>' + esc(label) + (m.units[c] ? ' <span class="faint small">' + esc(m.units[c]) + '</span>' : '') + '</dt><dd>' + (m.roles.status === c ? PV.ui.statusChip(t) : esc(t) || '<span class="faint">–</span>') + '</dd>';
    }).join('');
    return '<section><div class="section-title" style="margin-top:0">' + esc(title) + '</div><dl class="kv">' + items + '</dl><div class="small faint" style="margin-top:6px">' + esc(m.sheetName) + ' · row ' + ref.row.sheetRow + '</div></section>';
  }

  PV.openProbe = function (p, app) {
    if (!p) return;
    const store = app.store;
    const s = PV.statusKey(p);
    const flags = p.flags.map((fl) => '<span class="flag">' + esc({ 'not-in-stock': 'only in Probe Lifetime', 'duplicate-id': 'duplicate ID', 'lifetime-ambiguous': 'several Lifetime rows' }[fl] || fl) + '</span>').join(' ');
    const hist = PV.ptext(p, 'history');
    const histHtml = hist ? '<section><div class="section-title" style="margin-top:0">History</div><div class="history">' + hist.split(store.cfg.historySeparator).map((x) => '<div>' + esc(x) + '</div>').join('') + '</div></section>' : '';
    const fam = store.family(p.family.key);
    const body = '<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">' + PV.ui.statusChip(PV.ptext(p, 'status')) + PV.stateFields(p).map((x) => PV.ui.stateChip(x.label, x.value, x.label)).join(' ') + ' ' + flags + '</div>' +
      kv(p.stock, 'Stock and Booking') + (fam.cfg.stockSheet && fam.cfg.lifetimeSheet && !p.life ? '<div class="muted small">No row in Probe Lifetime yet – booking creates one.</div>' : '') + kv(p.life, 'Probe Lifetime') + histHtml;
    const canBook = p.stock && s === 'in stock' && store.can('book');
    const foot = (canBook ? '<button class="btn primary" data-act="book">Book…</button>' : '') +
      (fam.life && fam.life.isTable && store.can('state') ? '<button class="btn" data-act="state">Update state…</button>' : '') +
      (store.can('status') ? '<button class="btn" data-act="status">Change status…</button>' : '') +
      (!store.writesAllowed ? '<span class="small muted">Read-only (phase 1). Writes are switched off in config.</span>' : '');
    PV.ui.drawer(esc(p.family.label || p.family.key) + ' · ' + esc(PV.logId(p)), body, foot, (d, close) => {
      d.querySelectorAll('[data-act]').forEach((b) => (b.onclick = () => {
        if (b.dataset.act === 'book') PV.bookDialog(p, app, close);
        if (b.dataset.act === 'state') PV.stateDialog(p, app, close);
        if (b.dataset.act === 'status') PV.statusDialog(p, app, close);
      }));
    });
  };

  PV.bookDialog = function (p, app, after) {
    const store = app.store;
    const projects = store.list('project');
    const body = '<div class="grid"><div>Book <b>' + esc(PV.logId(p)) + '</b> (' + esc(p.family.label || p.family.key) + ') for <b>' + esc(store.user.name) + ' (' + esc(store.user.initials) + ')</b>.</div>' +
      '<label class="field"><span>Project (required)</span>' +
      (projects.length ? '<select class="input" data-proj><option value="">Choose…</option>' + projects.map((x) => '<option>' + esc(x) + '</option>').join('') + '<option value="__other">Other…</option></select>' : '') +
      '<input class="input' + (projects.length ? ' hidden' : '') + '" data-proj-text placeholder="Project name"></label>' +
      '<div class="small muted">Writes Status = booked, Booked for, Date of booking' + (p.stock.model.roles.project != null ? ', Project' : '') + ' in Stock and Booking and creates/updates the Probe Lifetime row in the same step.</div></div>';
    PV.ui.modal('Book probe', body, [
      { label: 'Cancel' },
      { label: 'Book', primary: true, onClick: async (m) => {
        const sel = m.querySelector('[data-proj]'), txt = m.querySelector('[data-proj-text]');
        const proj = sel && sel.value && sel.value !== '__other' ? sel.value : txt.value.trim();
        if (!proj) { PV.ui.toast('Choose a project first.', true); return false; }
        const r = await PV.ui.runAction(() => store.book(p, { project: proj }), 'Booked ' + PV.logId(p) + ' for ' + proj + '.');
        if (r) { if (after) after(); app.render(); }
      } },
    ]);
    const sel = document.querySelector('.modal [data-proj]');
    if (sel) sel.onchange = () => document.querySelector('.modal [data-proj-text]').classList.toggle('hidden', sel.value !== '__other');
  };

  PV.statusDialog = function (p, app, after) {
    const store = app.store;
    const cur = PV.ptext(p, 'status');
    const opts = store.list('status', p.family);
    PV.ui.modal('Change status', '<p class="small muted">Semafor rule: probe status changes only by agreement between applications, development and operations.</p><label class="field"><span>Status of ' + esc(PV.logId(p)) + '</span><select class="input" data-s>' + opts.map((o) => '<option' + (PV.key(o) === PV.key(cur) ? ' selected' : '') + '>' + esc(o) + '</option>').join('') + '</select></label>', [
      { label: 'Cancel' },
      { label: 'Save', primary: true, onClick: async (m) => {
        const v = m.querySelector('[data-s]').value;
        if (PV.key(v) === PV.key(cur)) return true;
        const r = await PV.ui.runAction(() => store.setStatus(p, v), 'Status of ' + PV.logId(p) + ' set to ' + v + '.');
        if (r) { if (after) after(); app.render(); }
      } },
    ]);
  };

  PV.stateDialog = function (p, app, after) {
    const store = app.store;
    const fam = store.family(p.family.key);
    const fields = ['mech', 'ele', 'mag', 'state', 'current'].filter((r) => fam.life.roles[r] != null);
    const body = '<div class="grid">' + fields.map((r) => {
      const cur = PV.ptext(p, r);
      const opts = store.list(r, p.family);
      return '<label class="field"><span>' + esc(fam.life.labels[fam.life.roles[r]]) + '</span><select class="input" data-r="' + r + '"><option value="">' + (cur ? esc(cur) + ' (current)' : '–') + '</option>' + opts.filter((o) => PV.key(o) !== PV.key(cur)).map((o) => '<option>' + esc(o.trim()) + '</option>').join('') + '</select></label>';
    }).join('') + '<label class="field"><span>Note for History</span><textarea class="input" rows="3" data-note placeholder="e.g. tip crashed during approach on sample S12"></textarea></label></div>';
    PV.ui.modal('Update state · ' + esc(PV.logId(p)), body, [
      { label: 'Cancel' },
      { label: 'Save', primary: true, onClick: async (m) => {
        const ch = {};
        m.querySelectorAll('[data-r]').forEach((s) => { if (s.value) ch[s.dataset.r] = s.value; });
        const note = m.querySelector('[data-note]').value;
        const r = await PV.ui.runAction(() => store.updateState(p, ch, note), 'Saved to Probe Lifetime.');
        if (r) { if (after) after(); app.render(); } else return false;
      } },
    ]);
  };
})();
