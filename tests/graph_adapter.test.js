// GraphSource + Store end-to-end against a fake Graph Excel API backed by the demo workbooks.
const { load } = require('./harness');
const assert = require('assert');
const w = load(['util.js', 'config.js', 'model.js', 'source-demo.js', 'source-graph.js', 'store.js', 'stats.js']);
const PV = w.PV;
const demo = new PV.DemoSource();               // backing "Excel"
const BASE = 'https://graph.microsoft.com/v1.0';
const SHARE = { stock: 'https://nenovision.sharepoint.com/:x:/s/Ops/EaBcD?e=1', lifetime: 'https://nenovision.sharepoint.com/:x:/s/Ops/EzZz+/=?e=2' };
const log = [];
let throttleOnce = true, sessions = 0;
function res(status, body, headers) {
  return { status, ok: status >= 200 && status < 300, headers: { get: (k) => (headers || {})[k] || null }, text: async () => (body == null ? '' : JSON.stringify(body)) };
}
async function fakeFetch(url, init) {
  const u = new URL(url);
  const path = decodeURIComponent(u.pathname.replace('/v1.0', ''));
  const method = init.method;
  const body = init.body ? JSON.parse(init.body) : undefined;
  log.push({ method, path, query: u.search, session: init.headers['workbook-session-id'] || null, auth: init.headers.Authorization, body });
  assert.strictEqual(init.headers.Authorization, 'Bearer TOKEN');
  if (body !== undefined) assert.strictEqual(init.headers['Content-Type'], 'application/json');
  let m;
  if ((m = path.match(/^\/shares\/(u![A-Za-z0-9_-]+)\/driveItem$/))) {
    const wb = Object.keys(SHARE).find((k) => PV.encodeSharingUrl(SHARE[k]) === m[1]);
    assert(wb, 'share token must decode to a configured URL');
    assert(!/[=+/]/.test(m[1].slice(2)), 'unpadded base64url');
    return res(200, { id: 'item-' + wb, name: wb + '.xlsx', parentReference: { driveId: 'drive-1' } });
  }
  m = path.match(/^\/drives\/drive-1\/items\/item-(stock|lifetime)\/workbook(.*)$/);
  if (!m) return res(404, { error: { code: 'itemNotFound', message: 'no ' + path } });
  const wb = m[1], rest = m[2];
  if (rest === '/worksheets' && method === 'GET') { if (throttleOnce) { throttleOnce = false; return res(429, { error: { code: 'TooManyRequests' } }, { 'Retry-After': '0' }); } return res(200, { value: await demo.listWorksheets(wb) }); }
  if ((m = rest.match(/^\/worksheets\/([^/]+)\/tables$/))) return res(200, { value: await demo.listTables(wb, m[1]) });
  if ((m = rest.match(/^\/tables\/([^/]+)\/range$/))) return res(200, Object.assign({ '@odata.context': 'x', cellCount: 1 }, await demo.getTableRange(wb, m[1])));
  if ((m = rest.match(/^\/worksheets\/([^/]+)\/usedRange$/))) return res(200, await demo.getUsedRange(wb, m[1]));
  if ((m = rest.match(/^\/worksheets\/([^/]+)\/range\(address='([A-Z]+\d+(?::[A-Z]+\d+)?)'\)$/))) {
    if (method === 'GET') return res(200, await demo.getRange(wb, m[1], m[2]));
    if (method === 'PATCH') { assert(init.headers['workbook-session-id'], 'writes use a session'); await demo.writeRange(wb, m[1], m[2], body.values); return res(200, {}); }
  }
  if ((m = rest.match(/^\/tables\/([^/]+)\/rows\/add$/)) && method === 'POST') { assert(Array.isArray(body.values) && Array.isArray(body.values[0])); await demo.addTableRow(wb, m[1], body.values[0]); return res(201, { index: 1, values: body.values }); }
  if (rest === '/createSession' && method === 'POST') { assert.deepStrictEqual(body, { persistChanges: true }); return res(201, { id: 'sess-' + (++sessions), persistChanges: true }); }
  if (rest === '/closeSession' && method === 'POST') { assert(init.headers['workbook-session-id']); return res(204); }
  return res(400, { error: { code: 'invalidRequest', message: method + ' ' + rest } });
}
(async () => {
  const cfg = JSON.parse(JSON.stringify(w.PV_CONFIG));
  cfg.workbooks.stock.sharingUrl = SHARE.stock; cfg.workbooks.lifetime.sharingUrl = SHARE.lifetime;
  cfg.liveWritesEnabled = true;
  const src = new PV.GraphSource({ getToken: async () => 'TOKEN', workbooks: cfg.workbooks, fetchImpl: fakeFetch, sleep: async () => {} });
  const st = new PV.Store(cfg, src, { initials: 'EJ', name: 'Eduard', department: 'Apps', roles: ['engineer'] });
  await st.load();
  assert.strictEqual(st.mode, 'live');
  const loadReqs = log.length;
  console.log('load requests:', loadReqs, '(one range request per table; 429 retried once)');
  assert(log.filter((x) => x.path.endsWith('/worksheets')).length === 3, 'retried throttled worksheets call');
  // sheet ids with braces/quotes must be URL-encoded
  const npm = log.find((x) => x.path.includes('/tables') && x.path.includes('{demo-'));
  assert(npm, 'worksheet addressed by id');
  assert.strictEqual(st.family('NPC-X').probes.length, 120);
  const p = st.family('NPC-W').probes.find((x) => x.stock && PV.statusKey(x) === 'in stock' && x.life == null);
  const before = log.length;
  const rep = await st.book(p, { project: 'Training' });
  assert(rep.stock && rep.lifetime, JSON.stringify(rep));
  const w2 = log.slice(before);
  console.log('booking requests:'); w2.forEach((x) => console.log('  ', x.method, x.path.replace(/\/drives\/drive-1\/items\/item-/, '…'), x.session ? '[session ' + x.session + ']' : '', x.body ? JSON.stringify(x.body).slice(0, 120) : ''));
  const patches = w2.filter((x) => x.method === 'PATCH');
  assert(patches.every((x) => x.session), 'all patches in a session');
  assert(patches.every((x) => /range\(address='[A-Z]+\d+(:[A-Z]+\d+)?'\)$/.test(x.path)));
  assert(w2.some((x) => x.path.endsWith('/createSession')) && w2.filter((x) => x.path.endsWith('/closeSession')).length === 2, 'sessions opened and closed for both workbooks');
  assert(w2.filter((x) => x.method === 'GET' && x.path.includes("range(address=")).length >= 2, 're-read before write');
  const fresh = st.family('NPC-W').probes.find((x) => x.id === p.id && x.stock);
  assert.strictEqual(PV.statusKey(fresh), 'booked'); assert(fresh.life);
  // a stock PATCH must never cover more than one cell (formulas next door stay intact)
  patches.filter((x) => x.path.includes('item-stock')).forEach((x) => assert(!x.path.includes(':'), 'single-cell writes in Stock'));
  // error mapping
  const bad = new PV.GraphSource({ getToken: async () => 'TOKEN', workbooks: { stock: { label: 'S', sharingUrl: 'https://x/notconfigured' } }, fetchImpl: fakeFetch, sleep: async () => {} });
  await assert.rejects(() => bad.listWorksheets('stock'), /share token|AssertionError|decode/);
  console.log('ALL GRAPH TESTS PASSED');
})().catch((e) => { console.error('FAIL', e); process.exit(1); });
