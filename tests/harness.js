// Loads the viewer's classic scripts into a Node vm context with a minimal window.
const fs = require('fs'), vm = require('vm'), path = require('path');
const APP = path.join(__dirname, '..', 'js');
function load(files, extra) {
  const store = {};
  const win = { localStorage: { getItem: (k) => (k in store ? store[k] : null), setItem: (k, v) => { store[k] = String(v); }, removeItem: (k) => { delete store[k]; } },
    TextEncoder, btoa: (s) => Buffer.from(s, 'binary').toString('base64'), console, setTimeout, Date, Math, JSON, Promise, Map, Set };
  Object.assign(win, extra || {});
  win.window = win; win.globalThis = win;
  const ctx = vm.createContext(win);
  for (const f of files) vm.runInContext(fs.readFileSync(path.join(APP, f), 'utf8'), ctx, { filename: f });
  return win;
}
/** Read-only source over an exported workbook (tools/export_json.py) – same interface as GraphSource. */
class JsonSource {
  constructor(book) { this.kind = 'demo'; this.book = book; }
  sh(id) { return this.book.sheets.find((s) => s.id === id); }
  async listWorksheets() { return this.book.sheets.map((s) => ({ id: s.id, name: s.name })); }
  async listTables(id) { return this.sh(id).tables.map((t) => ({ name: t.name })); }
  async getTableRange(name) { for (const s of this.book.sheets) { const t = s.tables.find((x) => x.name === name); if (t) return JSON.parse(JSON.stringify(t.range)); } throw new Error('no table ' + name); }
  async getUsedRange(id) { return JSON.parse(JSON.stringify(this.sh(id).used)); }
  async getRange() { throw new Error('read-only fixture'); }
  async beginSession() { return null; } async endSession() {}
  async writeRange() { throw new Error('read-only fixture'); } async addTableRows() { throw new Error('read-only fixture'); }
}
module.exports = { load, JsonSource, FILES: ['util.js', 'config.js', 'model.js', 'store.js', 'stats.js'] };
