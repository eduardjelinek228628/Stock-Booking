/* Probe Viewer – Stock screen: one sub-tab per sheet, two column groups, Excel-like filters, ID search, probe drawer. */
(function () {
  'use strict';
  const PV = window.PV;
  const esc = PV.esc;
  const STATE_ROLES = ['mech', 'ele', 'mag', 'state'];

  function famLabel(f) { return f.cfg.label || f.cfg.key; }
  function roleOf(m, c) { return Object.keys(m.roles).find((r) => m.roles[r] === c); }

  /** Column descriptors: first the Stock and Booking columns in sheet order, then the usage group in config order. */
  function columns(f, cfg) {
    const m = f.stock;
    if (!m) return [];
    const usageCols = cfg.usageRoles.map((r) => m.roles[r]).filter((c) => c != null);
    const mk = (c, usage) => {
      const role = roleOf(m, c);
      let label = PV.clean(m.headers[c]);
      if (!label) label = role ? role.charAt(0).toUpperCase() + role.slice(1) : PV.colLetter(m.c1 + c);
      return {
        key: 'c' + c, c, label, unit: m.units[c] || '', role, usage, isStatus: role === 'status', isState: STATE_ROLES.includes(role),
        text: (p) => PV.displayCell(p.stock.row.values[c], p.stock.row.text[c], p.stock.row.nf[c]),
        raw: (p) => p.stock.row.values[c],
        formula: (p) => { const x = p.stock.row.formulas[c]; return typeof x === 'string' && x.startsWith('='); },
      };
    };
    const base = m.headers.map((_, c) => c).filter((c) => !usageCols.includes(c)).map((c) => mk(c, false));
    const usage = usageCols.map((c) => mk(c, true));
    if (usage.length) usage[0].first = true;
    return base.concat(usage);
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
    if (col.role === 'history' && t) return esc(t.split(window.PV_CONFIG.historySeparator)[0]);
    return esc(t);
  }

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
    const fams = store.families.filter((f) => f.stock);
    if (!fams.length) { root.innerHTML = '<div class="empty">No probe sheets could be loaded. See Data checks.</div>'; return; }
    if (!st.family || !fams.find((f) => f.cfg.key === st.family)) st.family = fams[0].cfg.key;
    const f = store.family(st.family);
    const cols = columns(f, store.cfg);
    const allRows = f.probes;
    const filt = (st.filters[f.cfg.key] = st.filters[f.cfg.key] || {});
    let rows = applyFilters(allRows, cols, filt);
    const sort = st.sort[f.cfg.key];
    if (sort) {
      const col = cols.find((c) => c.key === sort.key);
      if (col) {
        const numeric = isNumericCol(col, allRows);
        rows = rows.slice().sort((a, b) => {
          if (numeric) { const x = PV.toNum(col.raw(a)) ?? PV.toNum(col.text(a)), y = PV.toNum(col.raw(b)) ?? PV.toNum(col.text(b)); if (x == null) return 1; if (y == null) return -1; return (x - y) * sort.dir; }
          const x = col.text(a), y = col.text(b); if (!x) return 1; if (!y) return -1;
          return x.localeCompare(y, undefined, { numeric: true }) * sort.dir;
        });
      }
    }
    const nBase = cols.filter((c) => !c.usage).length, nUsage = cols.length - nBase;
    const nFilters = Object.keys(filt).length;
    const missing = f.stock.missing.filter((h) => ['Owner', 'Department', 'Box', 'History'].includes(h));

    root.innerHTML =
      '<div class="toolbar"><input class="input" id="q" type="search" placeholder="Search probe ID in all types…" value="' + esc(st.q) + '" style="width:300px" aria-label="Search probe ID">' +
      '<span class="spacer" style="flex:1"></span><button class="btn" id="clearF"' + (nFilters || sort ? '' : ' disabled') + '>Clear filters' + (nFilters ? ' (' + nFilters + ')' : '') + '</button></div>' +
      '<div id="searchRes"></div>' +
      '<div class="subtabs" role="tablist">' + fams.map((x) => '<button class="subtab" role="tab" aria-selected="' + (x.cfg.key === f.cfg.key) + '" data-fam="' + esc(x.cfg.key) + '">' + esc(famLabel(x)) + '<span class="n">' + x.probes.length + '</span></button>').join('') + '</div>' +
      '<div class="small muted" style="margin:-2px 0 8px">' + esc(f.stock.sheetName) + ' (Table ' + esc(f.stock.tableName || '–') + ') · showing ' + rows.length + ' of ' + allRows.length +
      (missing.length ? ' · <span style="color:var(--warning-ink)">usage columns not added yet – see Data checks</span>' : '') + '</div>' +
      '<div class="table-wrap"><table class="data"><thead>' +
      '<tr class="groups"><th colspan="' + nBase + '">Stock and Booking</th>' + (nUsage ? '<th colspan="' + nUsage + '" class="g-usage">Usage</th>' : '') + '</tr>' +
      '<tr class="cols">' + cols.map((c) => '<th class="' + (c.usage ? 'usage ' : '') + (c.first ? 'first-usage ' : '') + (filt[c.key] || (sort && sort.key === c.key) ? 'filtered' : '') + '"><button class="th-btn" data-col="' + c.key + '">' + esc(c.label) + (c.unit ? ' <span class="unit">' + esc(c.unit) + '</span>' : '') +
        '<span class="ficon">' + (sort && sort.key === c.key ? (sort.dir > 0 ? '▲' : '▼') : '') + (filt[c.key] ? '⏷' : '▾') + '</span></button></th>').join('') + '</tr>' +
      '</thead><tbody>' +
      rows.map((p) => '<tr data-uid="' + esc(p.uid) + '">' + cols.map((c) => {
        const numeric = typeof c.raw(p) === 'number' && !c.isStatus && c.role !== 'id';
        return '<td class="' + (numeric ? 'num ' : '') + (c.usage ? 'usage ' : '') + (c.first ? 'first-usage ' : '') + (c.formula(p) ? 'formula' : '') + '" title="' + esc(c.text(p)) + '">' + cellHtml(c, p) + '</td>';
      }).join('') + '</tr>').join('') +
      '</tbody></table>' + (rows.length ? '' : '<div class="empty">No probes match the filters.</div>') + '</div>';

    root.querySelectorAll('[data-fam]').forEach((b) => (b.onclick = () => { st.family = b.dataset.fam; app.render(); }));
    root.querySelector('#clearF').onclick = () => { st.filters[f.cfg.key] = {}; delete st.sort[f.cfg.key]; app.render(); };
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
    host.innerHTML = '<div class="card" style="margin-bottom:14px"><div class="card-h"><h3>' + hits.length + (hits.length === 50 ? '+' : '') + ' probe(s) matching "' + esc(q) + '"</h3></div>' +
      (hits.length ? '<table class="plain"><tr><th>ID</th><th>Type</th><th>Status</th><th>Booked for</th><th>Owner</th></tr>' + hits.map((p) =>
        '<tr data-uid="' + esc(p.uid) + '" style="cursor:pointer"><td><b>' + esc(PV.logId(p)) + '</b></td><td>' + esc(p.family.label || p.family.key) + '</td><td>' + PV.ui.statusChip(PV.ptext(p, 'status')) + '</td><td>' + esc(PV.ptext(p, 'bookedFor')) + '</td><td>' + esc(PV.ptext(p, 'owner')) + '</td></tr>').join('') + '</table>' : '<div class="empty">No probe with that ID.</div>') + '</div>';
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
      vals.map((v, i) => '<label><input type="checkbox" data-v="' + i + '"' + (!cur.values || cur.values.has(v) ? ' checked' : '') + '> ' + (v === '' ? '<i>(Blanks)</i>' : esc(v.length > 60 ? v.slice(0, 60) + '…' : v)) + '<span class="c">' + counts.get(v) + '</span></label>').join('') + '</div>' +
      '<div class="row" style="justify-content:flex-end"><button class="btn sm" data-clear>Clear</button><button class="btn sm primary" data-ok>Apply</button></div>';
    document.body.appendChild(pop);
    const r = btn.getBoundingClientRect();
    pop.style.left = Math.max(8, Math.min(r.left, window.innerWidth - 280)) + 'px';
    pop.style.top = Math.max(8, Math.min(r.bottom + 4, window.innerHeight - pop.offsetHeight - 8)) + 'px';
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
  function kv(p, usage, cfg) {
    const m = p.stock.model;
    const usageCols = cfg.usageRoles.map((r) => m.roles[r]).filter((c) => c != null);
    const cs = usage ? usageCols : m.headers.map((_, c) => c).filter((c) => !usageCols.includes(c));
    if (!cs.length) return '';
    const items = cs.map((c) => {
      const role = roleOf(m, c);
      if (role === 'history') return '';
      const label = PV.clean(m.headers[c]) || (role === 'status' ? 'Status' : PV.colLetter(m.c1 + c));
      const t = PV.displayCell(p.stock.row.values[c], p.stock.row.text[c], p.stock.row.nf[c]);
      const v = role === 'status' ? PV.ui.statusChip(t) : STATE_ROLES.includes(role) && t ? PV.ui.stateChip(label, t) : (esc(t) || '<span class="faint">–</span>');
      return '<dt>' + esc(label) + (m.units[c] ? ' <span class="faint small">' + esc(m.units[c]) + '</span>' : '') + '</dt><dd>' + v + '</dd>';
    }).join('');
    return '<section><div class="section-title" style="margin-top:0">' + (usage ? 'Usage' : 'Stock and Booking') + '</div><dl class="kv">' + items + '</dl></section>';
  }

  PV.openProbe = function (p, app) {
    if (!p) return;
    const store = app.store, cfg = store.cfg;
    const s = PV.statusKey(p);
    const flags = p.flags.map((fl) => '<span class="flag">' + esc({ 'duplicate-id': 'duplicate ID' }[fl] || fl) + '</span>').join(' ');
    const hist = PV.ptext(p, 'history');
    const histHtml = hist ? '<section><div class="section-title" style="margin-top:0">History</div><div class="history">' + hist.split(cfg.historySeparator).map((x) => '<div>' + esc(x) + '</div>').join('') + '</div></section>' : '';
    const m = p.stock.model;
    const body = '<div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap">' + PV.ui.statusChip(PV.ptext(p, 'status')) + PV.stateFields(p).map((x) => PV.ui.stateChip(x.label, x.value, x.label)).join(' ') + ' ' + flags + '</div>' +
      kv(p, false, cfg) + kv(p, true, cfg) + histHtml + '<div class="small faint">' + esc(m.sheetName) + ' · row ' + p.stock.row.sheetRow + '</div>';
    const hasState = PV.stateFields(p).length || m.roles.history != null;
    const foot = (s === 'in stock' && store.can('book') ? '<button class="btn primary" data-act="book">Book…</button>' : '') +
      (hasState && store.can('state') ? '<button class="btn" data-act="state">Update state…</button>' : '') +
      (store.can('status') ? '<button class="btn" data-act="status">Change status…</button>' : '') +
      (!store.writesAllowed ? '<span class="small muted">Read-only: writes are switched off in config.</span>' : '');
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
    const m = p.stock.model;
    const writes = ['Status = booked', 'Booked for', m.roles.bookingDate != null ? 'Date of booking' : null, m.roles.project != null ? 'Project' : null, m.roles.owner != null ? 'Owner' : null, m.roles.history != null ? 'History' : null].filter(Boolean);
    const body = '<div class="grid"><div>Book <b>' + esc(PV.logId(p)) + '</b> (' + esc(p.family.label || p.family.key) + ') for <b>' + esc(store.user.name) + ' (' + esc(store.user.initials) + ')</b>.</div>' +
      '<label class="field"><span>Project (required)</span>' +
      (projects.length ? '<select class="input" data-proj><option value="">Choose…</option>' + projects.map((x) => '<option>' + esc(x) + '</option>').join('') + '<option value="__other">Other…</option></select>' : '') +
      '<input class="input' + (projects.length ? ' hidden' : '') + '" data-proj-text placeholder="Project name"></label>' +
      '<div class="small muted">Writes ' + esc(writes.join(', ')) + ' in Stock and Booking.</div></div>';
    PV.ui.modal('Book probe', body, [
      { label: 'Cancel' },
      { label: 'Book', primary: true, onClick: async (mo) => {
        const sel = mo.querySelector('[data-proj]'), txt = mo.querySelector('[data-proj-text]');
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
    PV.ui.modal('Change status', '<p class="small muted">Semafor rule: probe status changes only by agreement between applications, development and operations.</p><label class="field"><span>Status of ' + esc(PV.logId(p)) + '</span><select class="input" data-s>' + store.list('status').map((o) => '<option' + (PV.key(o) === PV.key(cur) ? ' selected' : '') + '>' + esc(o) + '</option>').join('') + '</select></label>', [
      { label: 'Cancel' },
      { label: 'Save', primary: true, onClick: async (mo) => {
        const v = mo.querySelector('[data-s]').value;
        if (PV.key(v) === PV.key(cur)) return true;
        const r = await PV.ui.runAction(() => store.setStatus(p, v), 'Status of ' + PV.logId(p) + ' set to ' + v + '.');
        if (r) { if (after) after(); app.render(); }
      } },
    ]);
  };

  PV.stateDialog = function (p, app, after) {
    const store = app.store;
    const fields = PV.stateFields(p);
    const body = '<div class="grid">' + fields.map((fl) => {
      const opts = store.list(fl.role, p.family);
      return '<label class="field"><span>' + esc(fl.label) + '</span><select class="input" data-r="' + fl.role + '"><option value="">' + (fl.value ? esc(fl.value) + ' (current)' : '–') + '</option>' + opts.filter((o) => PV.key(o) !== PV.key(fl.value)).map((o) => '<option>' + esc(o) + '</option>').join('') + '</select></label>';
    }).join('') + '<label class="field"><span>Note for History</span><textarea class="input" rows="3" data-note placeholder="e.g. tip crashed during approach on sample S12"></textarea></label></div>';
    PV.ui.modal('Update state · ' + esc(PV.logId(p)), body, [
      { label: 'Cancel' },
      { label: 'Save', primary: true, onClick: async (mo) => {
        const ch = {};
        mo.querySelectorAll('[data-r]').forEach((s) => { if (s.value) ch[s.dataset.r] = s.value; });
        const r = await PV.ui.runAction(() => store.updateState(p, ch, mo.querySelector('[data-note]').value), 'Saved.');
        if (r) { if (after) after(); app.render(); } else return false;
      } },
    ]);
  };
})();
