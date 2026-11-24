// GraphSource + Store end-to-end against a fake Graph Excel API backed by the demo workbook.
const { load } = require('./harness');
const assert = require('assert');
const w = load(['util.js', 'config.js', 'model.js', 'source-demo.js', 'source-graph.js', 'store.js', 'stats.js']);
const PV = w.PV;
const log = [];
let demo, throttleOnce = true, sessions = 0;
const ITEM = '/drives/drive-1/items/item-1/workbook';
function res(status, body, headers) {
  return { status, ok: status >= 200 && status < 300, headers: { get: (k) => (headers || {})[k] || null }, text: async () => (body == null ? '' : JSON.stringify(body)) };
}
async function fakeFetch(url, init) {
  const u = new URL(url);
  const path = decodeURIComponent(u.pathname.replace('/v1.0', ''));
  const method = init.method;
  const body = init.body ? JSON.parse(init.body) : undefined;
  const session = init.headers['workbook-session-id'] || null;
  log.push({ method, path, session, body });
  assert.strictEqual(init.headers.Authorization, 'Bearer TOKEN');
  if (body !== undefined) assert.strictEqual(init.headers['Content-Type'], 'application/json');
  if (!path.startsWith(ITEM)) return res(404, { error: { code: 'itemNotFound', message: 'no ' + path } });
  const rest = path.slice(ITEM.length);
  const needSession = () => assert(session, 'writes use a session: ' + method + ' ' + rest);
  let m;
  if (rest === '/worksheets' && method === 'GET') { if (throttleOnce) { throttleOnce = false; return res(429, { error: { code: 'TooManyRequests' } }, { 'Retry-After': '0' }); } return res(200, { value: await demo.listWorksheets() }); }
  if ((m = rest.match(/^\/worksheets\/([^/]+)\/tables$/))) return res(200, { value: await demo.listTables(m[1]) });
  if ((m = rest.match(/^\/tables\/([^/]+)\/range$/))) return res(200, Object.assign({ '@odata.context': 'x', cellCount: 1 }, await demo.getTableRange(m[1])));
  if ((m = rest.match(/^\/worksheets\/([^/]+)\/usedRange$/))) return res(200, await demo.getUsedRange(m[1]));
  if ((m = rest.match(/^\/worksheets\/([^/]+)\/range\(address='([A-Z]+\d+(?::[A-Z]+\d+)?)'\)$/))) {
    if (method === 'GET') return res(200, await demo.getRange(m[1], m[2]));
    if (method === 'PATCH') { needSession(); await demo.writeRange(m[1], m[2], body.values); return res(200, {}); }
  }
  if ((m = rest.match(/^\/tables\/([^/]+)\/rows\/add$/)) && method === 'POST') { assert(Array.isArray(body.values[0])); await demo.addTableRows(m[1], body.values); return res(201, {}); }
  if ((m = rest.match(/^\/tables\/([^/]+)\/columns\/add$/)) && method === 'POST') {
    needSession(); assert(Number.isInteger(body.index));
    try { await demo.addTableColumn(m[1], body.index, undefined); } catch (e) { return res(400, { error: { code: 'InvalidArgument', message: e.message } }); }
    return res(200, { index: body.index, name: 'Column1' });
  }
  if (rest === '/worksheets/add' && method === 'POST') { needSession(); return res(201, await demo.addWorksheet(body.name)); }
  if (rest === '/tables/add' && method === 'POST') {
    needSession(); assert.strictEqual(body.hasHeaders, true);
    const a = body.address.split('!'); return res(201, await demo.addTable(a[0].replace(/^'|'$/g, ''), a[1]));
  }
  if ((m = rest.match(/^\/tables\/([^/]+)$/)) && method === 'PATCH') { needSession(); await demo.renameTable(m[1], body.name); return res(200, { name: body.name }); }
  if (rest === '/createSession' && method === 'POST') { assert.deepStrictEqual(body, { persistChanges: true }); return res(201, { id: 'sess-' + (++sessions) }); }
  if (rest === '/closeSession' && method === 'POST') { assert(session); return res(204); }
  return res(400, { error: { code: 'invalidRequest', message: method + ' ' + rest } });
}
function liveStore(user) {
  const cfg = JSON.parse(JSON.stringify(w.PV_CONFIG));
  cfg.workbook.driveId = 'drive-1'; cfg.workbook.itemId = 'item-1';
  const src = new PV.GraphSource({ getToken: async () => 'TOKEN', workbook: cfg.workbook, fetchImpl: fakeFetch, sleep: async () => {} });
  return new PV.Store(cfg, src, user);
}
const ej = { initials: 'EJ', name: 'Eduard', department: 'Apps', roles: ['engineer'] };
(async () => {
  // ---- 1. today's workbook: operations runs the column setup through Graph ----
  demo = new PV.DemoSource({ persistent: false, legacy: true });
  const st = liveStore({ initials: 'OP', name: 'Ops', roles: ['engineer', 'operations'] });
  await st.load();
  assert.strictEqual(st.mode, 'live');
  assert.strictEqual(log.filter((x) => x.path.endsWith('/worksheets')).length, 2, 'throttled call retried');
  assert(log.some((x) => x.path.includes('/worksheets/{demo-')), 'worksheets addressed by id');
  let before = log.length;
  const out = await st.setupColumns();
  const setupReqs = log.slice(before);
  console.log('setup:', out.done.length, 'changes;', setupReqs.filter((x) => x.path.endsWith('/columns/add')).length, 'columns/add,',
    setupReqs.filter((x) => x.method === 'PATCH' && x.path.includes('range(')).length, 'header writes,', setupReqs.filter((x) => x.path.endsWith('/worksheets/add')).length, 'worksheets/add,',
    setupReqs.filter((x) => x.path.endsWith('/tables/add')).length, 'tables/add');
  assert.strictEqual(out.manual.length, 0);
  assert(st.log && st.log.table === 'ViewerLog', 'log table created and renamed');
  assert(st.families.every((f) => f.stock.missing.length === 0), 'all columns present');
  const ca = setupReqs.find((x) => x.path.endsWith('/columns/add'));
  console.log('  e.g.', ca.method, ca.path.replace(ITEM, '…'), JSON.stringify(ca.body));

  // ---- 2. booking through Graph: one session, re-read first, single-cell writes ----
  st.user = ej;
  const p = st.family('NPC-W').probes.find((x) => PV.statusKey(x) === 'in stock');
  before = log.length;
  const rep = await st.book(p, { project: 'Training' });
  const b = log.slice(before);
  console.log('booking requests:'); b.forEach((x) => console.log('  ', x.method, x.path.replace(ITEM, '…'), x.session ? '[' + x.session + ']' : '', x.body ? JSON.stringify(x.body).slice(0, 90) : ''));
  assert(rep.logged, JSON.stringify(rep));
  const patches = b.filter((x) => x.method === 'PATCH');
  assert(patches.length >= 8 && patches.every((x) => x.session && !/:[A-Z]/.test(x.path.split("address='")[1])), 'single-cell writes in a session');
  assert.strictEqual(b.filter((x) => x.path.endsWith('/rows/add')).length, 1, 'log written in one request');
  assert.strictEqual(b.filter((x) => x.path.endsWith('/createSession')).length, 1);
  assert.strictEqual(b.filter((x) => x.path.endsWith('/closeSession')).length, 1);
  const firstWrite = b.findIndex((x) => x.method === 'PATCH');
  assert(b.slice(0, firstWrite).some((x) => x.method === 'GET' && x.path.includes("range(address=")), 're-read before the first write');
  const fresh = st.family('NPC-W').probes.find((x) => x.key === p.key);
  assert.strictEqual(PV.statusKey(fresh), 'booked'); assert.strictEqual(PV.ptext(fresh, 'owner'), 'EJ');

  // ---- 3. Excel refuses columns/add at the end of a table -> planner never asks for it ----
  await assert.rejects(() => st.source.addTableColumn('NP_Con_Si', 999, 'X', 'sess-x'), /InvalidArgument|≤/);
  // ---- 4. sharing-link resolution still works (Files.ReadWrite.All setups) ----
  const enc = PV.encodeSharingUrl('https://nenovision.sharepoint.com/:x:/s/Ops/EzZz+/=?e=2');
  assert(enc.startsWith('u!') && !/[=+/]/.test(enc.slice(2)), 'unpadded base64url');
  console.log('ALL GRAPH TESTS PASSED');
})().catch((e) => { console.error('FAIL', e); process.exit(1); });
