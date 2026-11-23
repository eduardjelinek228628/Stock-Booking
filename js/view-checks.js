/* Probe Viewer – Data checks: what the viewer found in the workbooks, as input for the talk with operations. */
(function () {
  'use strict';
  const PV = window.PV;
  const esc = PV.esc;

  PV.views = PV.views || {};
  PV.views.checks = function (root, app) {
    const store = app.store;
    const all = store.warnings.slice();
    const famRows = [];
    for (const f of store.families) {
      for (const m of [f.stock, f.life]) if (m) all.push(...m.warnings);
      const joined = f.probes.filter((p) => p.stock && p.life);
      const lifeOnly = f.stock ? f.probes.filter((p) => !p.stock) : [];
      const mismatch = joined.filter((p) => PV.key(PV.clean(p.stock.row.values[p.stock.model.roles.status])) !== PV.key(PV.clean(p.life.row.values[p.life.model.roles.status])));
      famRows.push({ f, stockN: f.stock ? f.probes.filter((p) => p.stock).length : null, lifeN: f.life ? f.probes.filter((p) => p.life).length : null, joined: joined.length, lifeOnly, mismatch });
      if (lifeOnly.length) all.push({ level: 'warn', where: 'Probe Lifetime › ' + f.cfg.lifetimeSheet, msg: lifeOnly.length + ' probe(s) have a Lifetime row but no row in Stock and Booking › ' + f.cfg.stockSheet + ' (e.g. ' + lifeOnly.slice(0, 5).map((p) => p.id).join(', ') + ').' });
      if (mismatch.length) all.push({ level: 'info', where: f.cfg.key, msg: mismatch.length + ' probe(s) have a different Status in Stock and Booking and in Probe Lifetime (e.g. ' + mismatch.slice(0, 3).map((p) => p.id + ': ' + PV.ptext(p, 'status') + ' / ' + PV.clean(p.life.row.values[p.life.model.roles.status])).join(', ') + '). The viewer shows the Stock value.' });
      if (f.cfg.stockSheet && !f.cfg.lifetimeSheet) all.push({ level: 'info', where: f.cfg.key, msg: 'No Probe Lifetime sheet mapped (config lifetimeSheet: null) – booking only writes Stock and Booking.' });
    }
    const hasProject = store.families.some((f) => f.stock && f.stock.roles.project != null);
    const hasPrice = store.families.some((f) => f.stock && f.stock.roles.price != null);
    const depts = new Map();
    store.probes.forEach((p) => { const d = PV.clean(PV.ptext(p, 'department')); if (d) depts.set(d, (depts.get(d) || 0) + 1); });
    const order = { error: 0, warn: 1, info: 2 };
    all.sort((a, b) => order[a.level] - order[b.level]);

    const todo = [
      ['Project column in Stock and Booking (+ "projects" list in the data sheet)', hasProject],
      ['Prices (Price column, or price per type in config)', hasPrice || Object.keys(store.cfg.pricePerFamily).length > 0],
      ['Log sheet with Table "' + store.cfg.log.table + '"', !!store.log],
      ['Adama and CDT-FMR converted to Excel Tables', !store.families.some((f) => (f.cfg.key === 'Adama' || f.cfg.key === 'CDT-FMR') && f.life && !f.life.isTable)],
      ['Status header in NP_Pt-Sharp named "Status"', !store.families.some((f) => f.stock && f.stock.roles.status != null && PV.clean(f.stock.headers[f.stock.roles.status]) === '')],
      ['Department values unified (' + [...depts.keys()].length + ' spellings now)', [...new Set([...depts.keys()].map((d) => PV.normDept(d)))].length === [...depts.keys()].length],
    ];

    root.innerHTML =
      (store.mode === 'demo' ? '<div class="banner info" style="border-radius:6px;margin-bottom:12px">Demo mode: these checks run on the synthetic sample workbooks, which already include the proposed Project, Price and Log additions. In live mode they run on the real files.</div>' : '') +
      '<div class="grid" style="grid-template-columns:repeat(auto-fit,minmax(380px,1fr))">' +
      '<div class="card"><div class="card-h"><h3>Excel changes to agree with operations</h3></div><div class="card-b">' +
      todo.map(([t, ok]) => '<div class="check" style="grid-template-columns:60px 1fr"><span class="lvl ' + (ok ? 'info' : 'warn') + '">' + (ok ? '✓ done' : 'to do') + '</span><span>' + esc(t) + '</span></div>').join('') + '</div></div>' +
      '<div class="card"><div class="card-h"><h3>Sheets and matching</h3></div><div class="card-b"><table class="plain"><tr><th>Type</th><th class="num">Stock rows</th><th class="num">Lifetime rows</th><th class="num">Matched</th><th class="num">Only in Lifetime</th></tr>' +
      famRows.map((r) => '<tr><td>' + esc(r.f.cfg.label || r.f.cfg.key) + '</td><td class="num">' + (r.stockN ?? '–') + '</td><td class="num">' + (r.lifeN ?? '–') + '</td><td class="num">' + (r.stockN != null && r.lifeN != null ? r.joined : '–') + '</td><td class="num">' + (r.stockN != null && r.lifeN != null ? r.lifeOnly.length : '–') + '</td></tr>').join('') +
      '</table><div class="small muted" style="margin-top:8px">Rows are matched by ID (Akiyama: Box ID + position/Number). Duplicate IDs are not matched.</div></div></div>' +
      '</div>' +
      (store.semafor.length ? '<div class="card" style="margin-top:12px"><div class="card-h"><h3>Semafor</h3></div><div class="card-b"><table class="plain"><tr><th>Probe</th><th>Status</th><th>Note</th></tr>' + store.semafor.map((r) => '<tr><td>' + esc(r[0]) + '</td><td>' + esc(r[1]) + '</td><td>' + esc(r[2]) + '</td></tr>').join('') + '</table></div></div>' : '') +
      '<div class="card" style="margin-top:12px"><div class="card-h"><h3>Findings</h3><span class="small muted">' + all.length + '</span></div><div class="card-b">' +
      (all.length ? all.map((w) => '<div class="check"><span class="lvl ' + w.level + '">' + w.level + '</span><span><b>' + esc(w.where || '') + '</b> – ' + esc(w.msg) + '</span></div>').join('') : '<div class="empty">No findings.</div>') +
      '</div></div>';
  };
})();
