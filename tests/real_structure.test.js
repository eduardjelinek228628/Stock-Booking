// Runs the parser/store over the REAL Stock and Booking (exported locally with tools/export_json.py) and prints what the viewer would see,
// including where "Add missing columns" would insert the new columns.
// Usage: python3 tools/export_json.py "Stock and Booking.xlsx" data/stock.json && node tests/real_structure.test.js
// (data/ is git-ignored – never commit real data)
const { load, JsonSource, FILES } = require('./harness');
const fs = require('fs'), path = require('path');
const w = load(FILES);
const PV = w.PV;
const book = JSON.parse(fs.readFileSync(process.argv[2] || path.join(__dirname, '..', 'data', 'stock.json')));
(async () => {
  const st = new PV.Store(w.PV_CONFIG, new JsonSource(book), { initials: 'EJ', name: 'Eduard', roles: ['operations'] });
  await st.load();
  console.log('store warnings:'); st.warnings.forEach((x) => console.log('  ', x.level, x.where, '-', x.msg));
  for (const f of st.families) {
    const cnt = {};
    f.probes.forEach((p) => { const k = PV.statusKey(p) || '(blank)'; cnt[k] = (cnt[k] || 0) + 1; });
    console.log(`\n## ${f.cfg.key}: probes=${f.probes.length} status=${JSON.stringify(cnt)}`);
    if (!f.stock) continue;
    console.log('  roles', JSON.stringify(f.stock.roles));
    f.stock.warnings.forEach((x) => console.log('   ', x.level, '-', x.msg));
  }
  console.log('\nColumn setup plan:');
  const plan = st.setupPlan();
  for (const s of plan.sheets) {
    const after = s.family.stock.headers.map((h) => PV.clean(h) || '·');
    s.ops.forEach((o) => after.splice(o.index, 0, '[' + o.header + ']'));
    console.log(' ', s.family.cfg.key.padEnd(12), after.join(' | '), s.manual.length ? '  MANUAL: ' + s.manual.join(', ') : '');
  }
  console.log('  Log sheet to create:', plan.log);
  const counts = PV.stockCounts(st), ev = PV.bookingEvents(st);
  console.log('\nin stock:', counts.map((c) => c.family.cfg.key + ' ' + c['in stock']).join(', '), '| booking events with a date:', ev.length);
})().catch((e) => { console.error(e); process.exit(1); });
