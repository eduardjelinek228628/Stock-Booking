// Store actions on demo data: booking, conflicts, formula protection, roles, state updates, import, statistics, column setup.
const { load } = require('./harness');
const assert = require('assert');
const w = load(['util.js', 'config.js', 'model.js', 'source-demo.js', 'store.js', 'stats.js']);
const PV = w.PV;
const ej = { email: 'e', name: 'Eduard Jelínek', initials: 'EJ', department: 'Apps', roles: ['engineer'] };
const op = { email: 'o', name: 'Ops', initials: 'OP', department: 'Operations', roles: ['engineer', 'operations', 'importer'] };
const find = (st, fam, key) => st.family(fam).probes.find((x) => x.key === key);

(async () => {
  const src = new PV.DemoSource({ persistent: false });
  const st = new PV.Store(w.PV_CONFIG, src, ej);
  await st.load();
  const warns = st.warnings.concat(...st.families.map((f) => (f.stock ? f.stock.warnings : [])));
  console.log('warnings:', warns.length); warns.forEach((x) => console.log('  ', x.level, x.where, x.msg));
  assert(!warns.some((x) => x.level === 'error'), 'no errors on demo data');
  console.log('families', st.families.map((f) => f.cfg.key + ':' + f.probes.length).join(' '));
  assert(st.families.every((f) => f.stock && f.stock.missing.length === 0), 'demo has all new columns');
  console.log('log rows', st.logRows.length, 'projects', st.list('project').join(', '));

  // ---- book: one row, all fields ----
  let p = st.family('NPC-X').probes.find((x) => PV.statusKey(x) === 'in stock');
  const key = p.key, logBefore = st.logRows.length;
  const r = await st.book(p, { project: 'Training' });
  console.log('book report', JSON.stringify(r));
  p = find(st, 'NPC-X', key);
  assert.strictEqual(PV.statusKey(p), 'booked');
  assert.strictEqual(PV.ptext(p, 'bookedFor'), 'EJ');
  assert.strictEqual(PV.ptext(p, 'owner'), 'EJ');
  assert.strictEqual(PV.ptext(p, 'department'), 'Apps');
  assert.strictEqual(PV.ptext(p, 'bookingDate'), PV.todayISO());
  assert.strictEqual(PV.ptext(p, 'asOf'), PV.todayISO());
  assert.strictEqual(PV.ptext(p, 'project'), 'Training');
  assert.strictEqual(PV.ptext(p, 'mech'), 'new');
  assert.strictEqual(PV.ptext(p, 'ele'), 'new');
  assert(PV.ptext(p, 'history').startsWith(PV.todayISO() + ' EJ: booked for Training'));
  console.log('  logged', st.logRows.length - logBefore, ':', st.logRows.slice(-(st.logRows.length - logBefore)).map((e) => e.field).join(', '));
  assert(st.logRows.length - logBefore >= 8);
  await assert.rejects(() => st.book(p, { project: 'x' }), /Only probes "in stock"/);

  // ---- conflict detection ----
  const p2 = st.family('NPC-Z').probes.find((x) => PV.statusKey(x) === 'in stock');
  const sh = src.book.sheets.find((s) => s.name === 'NPC-Z');
  sh.cells[p2.stock.row.sheetRow + ':' + (p2.stock.model.c1 + p2.stock.model.roles.note)] = { v: 'edited in Excel', nf: 'General', f: null };
  await assert.rejects(() => st.book(p2, { project: 'Training' }), (e) => e.isConflict && e.fresh.some((x) => x.now === 'edited in Excel'));
  console.log('conflict detected ok');
  await st.load();

  // ---- formula protection ----
  const pc = st.family('NPC-Z').probes.find((x) => { const f = x.stock.row.formulas[x.stock.model.headers.indexOf('Primary ')]; return typeof f === 'string' && f.startsWith('='); });
  await assert.rejects(() => st.writeCell(pc.stock, pc.stock.model.headers.indexOf('Primary '), 'x', null), /formula/);
  console.log('formula protected ok');

  // ---- roles ----
  await assert.rejects(() => st.setStatus(p, 'sent out'), /operations/);
  await assert.rejects(() => st.setupColumns(), /operations/);
  st.user = op;
  await st.setStatus(find(st, 'NPC-X', key), 'sent out');
  assert.strictEqual(PV.statusKey(find(st, 'NPC-X', key)), 'sent out');
  console.log('setStatus ok');

  // ---- state update ----
  st.user = ej;
  await st.updateState(find(st, 'NPC-X', key), { mech: 'bad', ele: 'weak' }, 'crashed on sample S12');
  p = find(st, 'NPC-X', key);
  assert.strictEqual(PV.ptext(p, 'mech'), 'bad');
  const h = PV.ptext(p, 'history');
  assert(h.startsWith(PV.todayISO() + ' EJ: Mech. new → bad; Ele. new → weak; crashed on sample S12 | '), h);
  console.log('updateState ok:', h);

  // ---- boxes: move probes, book straight into a box ----
  const mineEJ = st.probes.filter((x) => PV.currentOwner(PV.ptext(x, 'owner')) === 'EJ' && x.stock.model.roles.box != null);
  const b1 = mineEJ.filter((x) => PV.boxOf(x) === 'Box 1').slice(0, 2);
  assert(b1.length === 2, 'EJ has probes in Box 1');
  const mv = await st.moveToBox(b1, 'box 9');
  assert.strictEqual(mv.moved, 2); assert.strictEqual(mv.failed.length, 0);
  for (const x of b1) {
    const n = find(st, x.family.key, x.key);
    assert.strictEqual(PV.ptext(n, 'box'), 'Box 9');
    assert(PV.ptext(n, 'history').startsWith(PV.todayISO() + ' EJ: moved to Box 9 (from Box 1)'), PV.ptext(n, 'history'));
  }
  assert(st.logRows.slice(-4).some((e) => e.field === 'Box' && e.old === 'Box 1' && e.new === 'Box 9'));
  await assert.rejects(() => st.moveToBox([find(st, b1[0].family.key, b1[0].key)], 'Box 9'), /Already in Box 9/);
  const pb = st.family('NPC-W').probes.find((x) => PV.statusKey(x) === 'in stock');
  await st.book(pb, { project: 'Training', box: 'Box 9' });
  assert.strictEqual(PV.ptext(find(st, 'NPC-W', pb.key), 'box'), 'Box 9');
  assert(PV.ptext(find(st, 'NPC-W', pb.key), 'history').includes('booked for Training, put in Box 9'));
  const hs = {}; mineEJ.forEach((x) => { const h = PV.health(x); hs[h] = (hs[h] || 0) + 1; });
  console.log('boxes ok, health of EJ probes', JSON.stringify(hs));

  // ---- Akiyama probe row with position ----
  const ak = st.family('Akiyama').probes.find((x) => PV.statusKey(x) === 'in stock' && x.key.includes('#'));
  if (ak) { const k = ak.key; await st.book(ak, { project: 'Training' }); assert.strictEqual(PV.ptext(find(st, 'Akiyama', k), 'state'), 'new'); console.log('akiyama booking ok', k); }

  // ---- import ----
  st.user = op;
  const fz = st.family('NPC-Z');
  const targets = fz.probes.slice(1, 4);
  const qCol = fz.stock.roles.q;
  const res = await st.importResults('NPC-Z', targets.map((t, i) => ({ probe: t, cells: [{ col: qCol, value: 111 + i }] })));
  assert.strictEqual(res.written, 3);
  assert.strictEqual(PV.toNum(find(st, 'NPC-Z', targets[0].key).stock.row.values[qCol]), 111);
  console.log('import ok');

  // ---- statistics ----
  const counts = PV.stockCounts(st);
  const ev = PV.bookingEvents(st);
  console.log('counts', counts.map((c) => c.family.cfg.key + ' in:' + c['in stock'] + ' b:' + c.booked + (c.low ? ' LOW' : '')).join(' | '));
  for (const per of [1, 3, 12, 'all']) { const s = PV.bookingSeries(ev, per); console.log('  bookings', per, s.weekly ? 'weekly' : 'monthly', s.labels.length, 'bins, sum', s.values.reduce((a, b) => a + b, 0)); }
  const sot = PV.stockOverTime(st, counts, 12, new Date(), ['NPC-X', 'NPC-W']);
  console.log('  stock over time X+W', JSON.stringify(sot.values));
  assert(sot.values.some((v) => v != null));
  console.log('  lifetimes', JSON.stringify(PV.lifetimes(st, ev).map((l) => [l.family, l.n, Math.round(l.median)])));
  assert(PV.lifetimes(st, ev).length > 3);
  console.log('  cost', JSON.stringify(PV.costPerProject(ev).rows.slice(0, 3).map((x) => [x.project, x.cost])));
  console.log('  depts', JSON.stringify(PV.bookingsBy(st, 'department')));

  // ---- column setup on today's workbook (no new columns, no Log) ----
  const legacy = new PV.DemoSource({ persistent: false, legacy: true });
  const st2 = new PV.Store(w.PV_CONFIG, legacy, op);
  await st2.load();
  assert(!st2.log);
  const plan = st2.setupPlan();
  for (const s of plan.sheets) console.log('  plan', s.family.cfg.key, s.ops.map((o) => o.index + ':' + o.header).join(' '), s.manual.length ? 'MANUAL ' + s.manual : '');
  const before = {}; st2.families.forEach((f) => { before[f.cfg.key] = { headers: f.stock.headers.slice(), first: f.probes[3] && f.probes[3].stock.row.values.slice() }; });
  const out = await st2.setupColumns();
  console.log('  setup done', out.done.length, 'manual', out.manual.length);
  assert.strictEqual(out.manual.length, 0);
  assert(st2.log, 'log table created');
  for (const f of st2.families) {
    assert.strictEqual(f.stock.missing.length, 0, f.cfg.key + ' still missing ' + f.stock.missing);
    // existing columns keep their order and values
    const old = before[f.cfg.key];
    const kept = f.stock.headers.filter((hh) => old.headers.includes(hh));
    assert.strictEqual(JSON.stringify(kept), JSON.stringify(old.headers), f.cfg.key + ' order changed');
    const p3 = f.probes[3];
    old.headers.forEach((hh, i) => assert(String(p3.stock.row.values[f.stock.headers.indexOf(hh)]) === String(old.first[i]), f.cfg.key + ' value moved in ' + hh));
    console.log('  ', f.cfg.key, f.stock.headers.map((x) => PV.clean(x) || '·').join(' | '));
  }
  // and the viewer works on it right away
  const lp = st2.family('NPC-W').probes.find((x) => PV.statusKey(x) === 'in stock');
  st2.user = ej; await st2.book(lp, { project: 'Training' });
  assert.strictEqual(PV.ptext(find(st2, 'NPC-W', lp.key), 'owner'), 'EJ');
  assert(st2.logRows.length >= 8, 'log written after setup');
  const again = st2.setupPlan(); assert(again.sheets.length === 0 && again.log === false, 'nothing left to set up');
  console.log('setup ok');
  console.log('ALL STORE TESTS PASSED');
})().catch((e) => { console.error('FAIL', e); process.exit(1); });
