// Runs the parser/store over the REAL workbooks (exported locally with tools/export_json.py) and prints what the viewer would see.
// Usage: python3 tools/export_json.py StockBooking.xlsx data/stock.json && python3 tools/export_json.py ProbeLifetime.xlsx data/lifetime.json
//        node tests/real_structure.test.js     (data/ is git-ignored – never commit real data)
const { load, JsonSource, FILES } = require('./harness');
const fs = require('fs');
const w = load(FILES);
const books = { stock: JSON.parse(fs.readFileSync(require('path').join(__dirname,'..','data','stock.json'))), lifetime: JSON.parse(fs.readFileSync(require('path').join(__dirname,'..','data','lifetime.json'))) };
(async () => {
  const st = new w.PV.Store(w.PV_CONFIG, new JsonSource(books), { initials: 'EJ', name: 'Eduard', roles: [] });
  await st.load();
  console.log('store warnings:'); st.warnings.forEach((x) => console.log('  ', x.level, x.where, '-', x.msg));
  for (const f of st.families) {
    const cnt = {};
    f.probes.forEach((p) => { const k = w.PV.statusKey(p) || '(blank)'; cnt[k] = (cnt[k] || 0) + 1; });
    const joined = f.probes.filter((p) => p.stock && p.life).length;
    const sm = f.stock, lm = f.life;
    console.log(`\n## ${f.cfg.key}: probes=${f.probes.length} joined=${joined} status=${JSON.stringify(cnt)}`);
    if (sm) console.log('  stock roles', JSON.stringify(sm.roles), 'units', JSON.stringify(sm.units.filter(Boolean)));
    if (lm) console.log('  life roles ', JSON.stringify(lm.roles));
    for (const m of [sm, lm]) if (m) m.warnings.forEach((x) => console.log('   ', x.level, x.where, '-', x.msg));
  }
  console.log('\nlists:', Object.keys(st.lists).join(' | '));
  console.log('status list', st.list('status'), 'mech', st.list('mech'), 'akiyama state', st.list('state', w.PV_CONFIG.families.find(f=>f.key==='Akiyama')));
  const ej = st.probes.filter((p) => w.PV.currentOwner(w.PV.ptext(p, 'owner')).toUpperCase() === 'EJ');
  console.log('EJ owned probes:', ej.length, ej.slice(0, 5).map((p) => p.family.key + ':' + p.id + ' ' + w.PV.ptext(p,'status') + ' mech=' + w.PV.ptext(p,'mech')));
  const sample = st.family('NPC-X').probes.find((p) => p.stock && p.life);
  if (sample) console.log('sample join', sample.id, 'asOf', w.PV.ptext(sample, 'asOf'), 'booking', w.PV.ptext(sample, 'bookingDate'), 'freso', w.PV.ptext(sample,'freso'));
})().catch((e) => { console.error(e); process.exit(1); });
