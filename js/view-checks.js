/* Probe Viewer – Data checks: what the viewer found in Stock and Booking, and the one-time column setup. */
(function () {
  'use strict';
  const PV = window.PV;
  const esc = PV.esc;

  function planHtml(store, plan) {
    const rows = plan.sheets.map((s) => {
      const after = s.family.stock.headers.map((h) => esc(PV.clean(h) || '·'));
      s.ops.forEach((o) => after.splice(o.index, 0, '<span class="newcol">' + esc(o.header) + '</span>'));
      return '<div style="margin-bottom:8px"><b>' + esc(s.family.stock.sheetName) + '</b><div class="plan">' + after.join(' | ') + '</div>' +
        (s.manual.length ? '<div class="small" style="color:var(--warning-ink)">Add by hand (cannot insert after the last column): ' + esc(s.manual.join(', ')) + '</div>' : '') + '</div>';
    }).join('');
    return rows + (plan.log ? '<div><b>New sheet "' + esc(store.cfg.log.sheet) + '"</b> with Table "' + esc(store.cfg.log.table) + '": <span class="plan">' + store.cfg.logColumns.map(esc).join(' | ') + '</span></div>' : '');
  }

  PV.views = PV.views || {};
  PV.views.checks = function (root, app) {
    const store = app.store, cfg = store.cfg;
    const all = store.warnings.slice();
    for (const f of store.families) if (f.stock) all.push(...f.stock.warnings);
    const order = { error: 0, warn: 1, info: 2 };
    all.sort((a, b) => order[a.level] - order[b.level]);
    const plan = store.setupPlan();
    const loaded = store.families.filter((f) => f.stock);
    const missingSheets = loaded.filter((f) => f.stock.missing.length);
    const depts = new Map();
    store.probes.forEach((p) => { const d = PV.clean(PV.ptext(p, 'department')); if (d) depts.set(d, (depts.get(d) || 0) + 1); });
    const priceFilled = store.probes.filter((p) => PV.toNum(PV.ptext(p, 'price')) !== null).length;
    const todo = [
      ['New columns (Price, Project, Owner, Department, Box, state, As of, History) in every sheet' + (missingSheets.length ? ' – missing in ' + missingSheets.map((f) => f.cfg.label || f.cfg.key).join(', ') : ''), !missingSheets.length],
      ['Log sheet with Table "' + cfg.log.table + '"', !!store.log],
      ['Prices filled in (' + priceFilled + ' of ' + store.probes.length + ' probes, or pricePerFamily in config)', priceFilled > 0 || Object.keys(cfg.pricePerFamily).length > 0],
      ['Status header in NP_Pt-Sharp named "Status"', !store.families.some((f) => f.stock && f.stock.roles.status != null && PV.clean(f.stock.headers[f.stock.roles.status]) === '')],
      ['Department values unified (' + depts.size + ' spelling' + (depts.size === 1 ? '' : 's') + ' in use)', [...new Set([...depts.keys()].map((d) => PV.normDept(d)))].length === depts.size],
    ];
    const canSetup = store.can('setup');
    const setupNeeded = plan.sheets.length || plan.log;

    root.innerHTML =
      '<div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(400px,1fr))">' +
      '<div class="card"><div class="card-h"><h3>Workbook setup for the test period</h3></div><div class="card-b">' +
      todo.map(([t, ok]) => '<div class="check" style="grid-template-columns:64px 1fr"><span class="lvl ' + (ok ? 'done' : 'warn') + '">' + (ok ? '✓ done' : 'to do') + '</span><span>' + esc(t) + '</span></div>').join('') +
      '<div class="toolbar" style="margin:14px 0 0">' +
      (setupNeeded ? '<button class="btn primary" id="setup"' + (canSetup ? '' : ' disabled') + '>Add missing columns…</button>' + (canSetup ? '' : '<span class="small muted">Only operations can change the workbook structure' + (store.writesAllowed ? ' (config: roles.operations).' : ' – writes are switched off.') + '</span>') : '<span class="small" style="color:var(--good-ink)">All new columns and the Log are in place.</span>') +
      (store.mode === 'demo' ? '<span class="spacer" style="flex:1"></span><button class="btn subtle sm" id="demoLegacy" title="Rebuild the demo as the workbook looks today, so you can try the setup">Demo: start without new columns</button>' : '') +
      '</div></div></div>' +
      '<div class="card"><div class="card-h"><h3>Sheets</h3></div><div class="card-b"><table class="plain"><tr><th>Type</th><th>Sheet / Table</th><th class="num">Probes</th><th class="num">In stock</th><th class="num">Duplicate IDs</th><th class="num">New columns</th></tr>' +
      cfg.families.map((fc) => {
        const f = store.family(fc.key);
        if (!f || !f.stock) return '<tr><td>' + esc(fc.label || fc.key) + '</td><td colspan="5" style="color:var(--critical-ink)">not loaded – see findings</td></tr>';
        const want = 1 + PV.wantedColumns(cfg, fc).after.length;
        return '<tr><td>' + esc(fc.label || fc.key) + '</td><td class="small">' + esc(f.stock.sheetName) + ' / ' + esc(f.stock.tableName) + '</td><td class="num">' + f.probes.length + '</td><td class="num">' + f.probes.filter((p) => PV.statusKey(p) === 'in stock').length + '</td><td class="num">' + f.stock.stats.dupIds.length + '</td><td class="num">' + (want - f.stock.missing.length) + ' / ' + want + '</td></tr>';
      }).join('') + '</table></div></div>' +
      '</div>' +
      (store.semafor.length ? '<div class="card" style="margin-top:14px"><div class="card-h"><h3>Semafor</h3></div><div class="card-b"><table class="plain"><tr><th>Probe</th><th>Status</th><th>Note</th></tr>' + store.semafor.map((r) => '<tr><td>' + esc(r[0]) + '</td><td>' + esc(r[1]) + '</td><td>' + esc(r[2]) + '</td></tr>').join('') + '</table></div></div>' : '') +
      '<div class="card" style="margin-top:14px"><div class="card-h"><h3>Findings</h3><span class="small muted">' + all.length + '</span></div><div class="card-b">' +
      (all.length ? all.map((w) => '<div class="check"><span class="lvl ' + w.level + '">' + w.level + '</span><span><b>' + esc(w.where || '') + '</b> – ' + esc(w.msg) + '</span></div>').join('') : '<div class="empty">No findings.</div>') +
      '</div></div>';

    const sb = root.querySelector('#setup');
    if (sb) sb.onclick = () => PV.ui.modal('Add missing columns',
      '<p>The viewer inserts these columns with Excel\'s own "insert table column" – existing columns, values and formulas move along, nothing is deleted or overwritten. New columns are highlighted.</p>' +
      (store.mode === 'live' ? '<p style="color:var(--warning-ink)"><b>Try it on a copy of the file first.</b> Excel has no undo for changes made through the API.</p>' : '') +
      '<div style="max-height:46vh;overflow:auto;border:1px solid var(--line);border-radius:8px;padding:10px">' + planHtml(store, plan) + '</div>',
      [{ label: 'Cancel' }, { label: 'Add columns', primary: true, onClick: async () => {
        const r = await PV.ui.runAction(() => store.setupColumns(), 'Columns added.');
        if (r && r.manual && r.manual.length) PV.ui.toast('Add by hand: ' + r.manual.join(', '), true, 12000);
        app.render();
      } }]);
    const dl = root.querySelector('#demoLegacy');
    if (dl) dl.onclick = async () => {
      store.source.data = PV.buildDemoBooks(new Date(), { legacy: true });
      store.source.persist();
      await app.refresh();
      app.go('checks');
    };
  };
})();
