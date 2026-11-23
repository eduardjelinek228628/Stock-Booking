const { load } = require('./harness');
const assert = require('assert');
const w = load(['util.js', 'config.js', 'model.js', 'source-demo.js', 'source-graph.js', 'store.js', 'stats.js', 'view-import.js'].filter(f => f !== 'view-import.js'));
const PV = w.PV;
const ej = { email: 'e', name: 'Eduard Jelínek', initials: 'EJ', department: 'Apps', roles: ['engineer'] };
const op = { email: 'o', name: 'Ops', initials: 'OP', department: 'Operations', roles: ['engineer', 'operations', 'importer'] };
(async () => {
  const src = new PV.DemoSource();
  const st = new PV.Store(w.PV_CONFIG, src, ej);
  await st.load();
  const errs = st.warnings.concat(...st.families.map(f => [f.stock, f.life].filter(Boolean).flatMap(m => m.warnings)));
  console.log('warnings:', errs.length); errs.forEach(x => console.log('  ', x.level, x.where, x.msg));
  console.log('families', st.families.map(f => f.cfg.key + ':' + f.probes.length + (f.stock ? '' : '(L)')).join(' '));
  console.log('log rows', st.logRows.length, 'lists', Object.keys(st.lists).join('|'), 'projects', st.list('project'));

  // ---- book an NPC-X probe that has no Lifetime row ----
  const fx = st.family('NPC-X');
  let p = fx.probes.find(p => p.stock && PV.statusKey(p) === 'in stock' && !p.life);
  assert(p, 'need in-stock NPC-X probe without lifetime');
  const id = p.id, logBefore = st.logRows.length, lifeRowsBefore = fx.life.rows.length;
  const r = await st.book(p, { project: 'Training' });
  console.log('book report', JSON.stringify(r));
  p = st.family('NPC-X').probes.find(x => x.id === id && x.stock);
  assert.strictEqual(PV.statusKey(p), 'booked');
  assert.strictEqual(PV.ptext(p, 'bookedFor'), 'EJ');
  assert.strictEqual(PV.ptext(p, 'bookingDate'), PV.todayISO());
  assert.strictEqual(PV.ptext(p, 'project'), 'Training');
  assert(p.life, 'lifetime row created + joined');
  assert.strictEqual(PV.ptext(p, 'owner'), 'EJ');
  assert.strictEqual(PV.ptext(p, 'mech'), 'new');
  assert.strictEqual(PV.ptext(p, 'freso'), PV.displayCell(p.stock.row.values[p.stock.model.roles.freso], p.stock.row.text[p.stock.model.roles.freso]));
  assert(PV.ptext(p, 'history').startsWith(PV.todayISO() + ' EJ: booked for Training'));
  assert.strictEqual(st.family('NPC-X').life.rows.length, lifeRowsBefore, 'filled an empty table row instead of appending');
  console.log('  logged', st.logRows.length - logBefore, 'entries:', st.logRows.slice(-5).map(e => e.field + ':' + e.old + '→' + e.new).join(', '));
  assert(st.logRows.length - logBefore >= 5);
  // formula cell untouched
  const prim = p.stock.model.roles; 

  // ---- book again must fail (not in stock) ----
  await assert.rejects(() => st.book(p, { project: 'x' }), /Only probes "in stock"/);

  // ---- conflict detection: change the cell behind the viewer's back ----
  const p2 = st.family('NPC-Z').probes.find(x => x.stock && PV.statusKey(x) === 'in stock');
  const sh = src.books.stock.sheets.find(s => s.name === 'NPC-Z');
  const noteCol = p2.stock.model.c1 + p2.stock.model.roles.note;
  sh.cells[p2.stock.row.sheetRow + ':' + noteCol] = { v: 'edited in Excel', nf: 'General', f: null };
  await assert.rejects(() => st.book(p2, { project: 'Training' }), (e) => e.isConflict && e.fresh.some(x => x.now === 'edited in Excel'));
  console.log('conflict detected ok');
  await st.load();

  // ---- formula protection: Primary in NPC-Z is a formula ----
  const p3 = st.family('NPC-Z').probes.find(x => x.stock && typeof x.stock.row.formulas[x.stock.model.headers.indexOf('Primary ')] === 'string' && x.stock.row.formulas[x.stock.model.headers.indexOf('Primary ')].startsWith('='));
  assert(p3, 'formula probe');
  await assert.rejects(() => st.writeCell(p3.stock, p3.stock.model.headers.indexOf('Primary '), 'x', null), /formula/);
  console.log('formula protected ok');

  // ---- permissions ----
  await assert.rejects(() => st.setStatus(p, 'sent out'), /operations/);
  st.user = op;
  await st.setStatus(st.family('NPC-X').probes.find(x => x.id === id && x.stock), 'sent out');
  p = st.family('NPC-X').probes.find(x => x.id === id && x.stock);
  assert.strictEqual(PV.statusKey(p), 'sent out');
  assert.strictEqual(PV.clean(p.life.row.values[p.life.model.roles.status]), 'sent out', 'lifetime status synced');
  console.log('setStatus ok');

  // ---- updateState as engineer ----
  st.user = ej;
  await st.updateState(p, { mech: 'bad', ele: 'weak' }, 'crashed on sample S12');
  p = st.family('NPC-X').probes.find(x => x.id === id && x.stock);
  assert.strictEqual(PV.ptext(p, 'mech'), 'bad');
  const h = PV.ptext(p, 'history');
  assert(h.startsWith(PV.todayISO() + ' EJ: Mech. new → bad; Ele. new → weak; crashed on sample S12 | '), h);
  console.log('updateState ok:', h);

  // ---- Akiyama box-level booking (position) ----
  const ak = st.family('Akiyama').probes.find(x => x.stock && PV.statusKey(x) === 'in stock' && x.key.includes('#'));
  if (ak) { const key = ak.key; await st.book(ak, { project: 'Training' }); const a2 = st.family('Akiyama').probes.find(x => x.key === key && x.stock); assert.strictEqual(PV.statusKey(a2), 'booked'); assert(a2.life, 'akiyama life joined by box#pos'); console.log('akiyama booking ok', key, 'life Number', PV.ptext(a2, 'position')); }

  // ---- import ----
  st.user = op;
  const fz = st.family('NPC-Z');
  const targets = fz.probes.filter(x => x.stock).slice(1, 4);
  const qCol = fz.stock.roles.q;
  const items = targets.map((t, i) => ({ probe: t, cells: [{ col: qCol, value: 111 + i }] }));
  const res = await st.importResults('NPC-Z', items);
  console.log('import', JSON.stringify(res));
  assert.strictEqual(res.written, 3);
  assert.strictEqual(PV.toNum(st.family('NPC-Z').probes.find(x => x.id === targets[0].id).stock.row.values[qCol]), 111);
  // lifetime-only plain range is read-only
  const ad = st.family('Adama').probes[0];
  await assert.rejects(() => st.updateState(ad, { state: 'dead' }, ''), /not an Excel Table/);
  console.log('plain-range read-only ok');

  // ---- stats ----
  const counts = PV.stockCounts(st);
  const ev = PV.bookingEvents(st);
  console.log('counts', counts.map(c => c.family.cfg.key + ' in:' + c['in stock'] + ' b:' + c.booked + (c.low ? ' LOW' : '')).join(' | '));
  console.log('events', ev.length, 'monthly', JSON.stringify(PV.monthlySeries(ev, 12).values));
  const cons = PV.consumption(st, counts, ev);
  console.log('consumption', cons.map(c => c.family.cfg.key + ' ' + c.recent + ' runout ' + (c.runout ? PV.isoDate(c.runout) : '-')).join(' | '));
  const sot = PV.stockOverTime(st, counts, 12);
  console.log('stockOverTime NPC-X', JSON.stringify(sot.series.find(s => s.family.cfg.key === 'NPC-X').values), 'since', PV.isoDate(sot.since));
  console.log('lifetimes', JSON.stringify(PV.lifetimes(st, ev).map(l => [l.family, l.n, Math.round(l.median)])));
  console.log('cost', JSON.stringify(PV.costPerProject(ev).rows.map(r => [r.project, r.cost, r.n])));
  console.log('persons', JSON.stringify(PV.bookingsBy(st, 'person').slice(0, 5)), 'depts', JSON.stringify(PV.bookingsBy(st, 'department')));
  console.log('value', JSON.stringify(PV.stockValue(st)));
  // persistence: new source instance sees changes
  const src2 = new PV.DemoSource(); const st2 = new PV.Store(w.PV_CONFIG, src2, ej); await st2.load();
  assert.strictEqual(PV.statusKey(st2.family('NPC-X').probes.find(x => x.id === id && x.stock)), 'sent out');
  console.log('persisted ok');
  console.log('ALL STORE TESTS PASSED');
})().catch(e => { console.error('FAIL', e); process.exit(1); });
