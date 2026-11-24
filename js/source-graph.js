/* Probe Viewer – live data source: Microsoft Graph Excel API (delegated, signed-in user), one workbook. */
(function () {
  'use strict';
  const PV = window.PV || (window.PV = {});

  function encodeSharingUrl(url) {
    // base64 -> unpadded base64url -> prefix "u!" (Graph "shares" API)
    const bytes = new TextEncoder().encode(url.trim());
    let bin = '';
    bytes.forEach((b) => { bin += String.fromCharCode(b); });
    return 'u!' + btoa(bin).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  }

  class GraphError extends Error {
    constructor(status, code, message, path) { super(message); this.status = status; this.code = code; this.path = path; }
  }

  /** opts: { getToken: async () => accessToken, workbook: config.workbook, fetchImpl?, baseUrl?, sleep? } */
  class GraphSource {
    constructor(opts) {
      this.kind = 'live';
      this.getToken = opts.getToken;
      this.workbook = opts.workbook;
      this.fetch = opts.fetchImpl || window.fetch.bind(window);
      this.base = opts.baseUrl || 'https://graph.microsoft.com/v1.0';
      this.sleep = opts.sleep || ((ms) => new Promise((r) => setTimeout(r, ms)));
      this.item = null; // '/drives/{d}/items/{i}'
    }

    async request(method, path, body, headers) {
      const url = path.startsWith('http') ? path : this.base + path;
      for (let attempt = 0; ; attempt++) {
        const token = await this.getToken();
        const h = Object.assign({ Authorization: 'Bearer ' + token }, headers || {});
        if (body !== undefined) h['Content-Type'] = 'application/json';
        const res = await this.fetch(url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body) });
        if (res.status === 429 || res.status === 503 || res.status === 504) {
          if (attempt >= 4) throw new GraphError(res.status, 'throttled', 'Microsoft Graph is busy (HTTP ' + res.status + '). Try again in a minute.', path);
          const ra = parseFloat(res.headers.get('Retry-After'));
          await this.sleep((isFinite(ra) ? ra : Math.pow(2, attempt)) * 1000);
          continue;
        }
        if (res.status === 204) return null;
        const text = await res.text();
        let json = null;
        try { json = text ? JSON.parse(text) : null; } catch (e) { json = null; }
        if (!res.ok) {
          const err = (json && json.error) || {};
          let msg = err.message || ('HTTP ' + res.status);
          if (res.status === 403 || res.status === 401) msg += ' – check that the app has access to this file (admin grant) and that you can open it in Excel.';
          throw new GraphError(res.status, err.code || 'http' + res.status, msg, path);
        }
        return json;
      }
    }

    async itemPath() {
      if (this.item) return this.item;
      const wb = this.workbook;
      let driveId = wb.driveId, itemId = wb.itemId;
      if (!(driveId && itemId)) {
        if (!wb.sharingUrl) throw new Error('Set workbook.driveId + workbook.itemId (or sharingUrl) in js/config.js.');
        const item = await this.request('GET', '/shares/' + encodeSharingUrl(wb.sharingUrl) + '/driveItem?$select=id,name,parentReference');
        driveId = item.parentReference && item.parentReference.driveId;
        itemId = item.id;
        if (!driveId || !itemId) throw new Error('Could not resolve the file behind the sharing link.');
      }
      this.item = '/drives/' + encodeURIComponent(driveId) + '/items/' + encodeURIComponent(itemId) + '/workbook';
      return this.item;
    }
    sh(session) { return session ? { 'workbook-session-id': session } : undefined; }
    static pickRange(r) { return { address: r.address, values: r.values, text: r.text, numberFormat: r.numberFormat, formulas: r.formulas }; }
    ws(sheetId) { return '/worksheets/' + encodeURIComponent(sheetId); }

    async listWorksheets() {
      const p = await this.itemPath();
      const r = await this.request('GET', p + '/worksheets?$select=id,name,visibility');
      return r.value.map((w) => ({ id: w.id, name: w.name, visibility: w.visibility }));
    }
    async listTables(sheetId) {
      const p = await this.itemPath();
      const r = await this.request('GET', p + this.ws(sheetId) + '/tables?$select=id,name');
      return r.value.map((t) => ({ name: t.name, id: t.id }));
    }
    // One request returns header + body with values, text, numberFormat and formulas.
    async getTableRange(tableName) { const p = await this.itemPath(); return GraphSource.pickRange(await this.request('GET', p + '/tables/' + encodeURIComponent(tableName) + '/range')); }
    async getUsedRange(sheetId) { const p = await this.itemPath(); return GraphSource.pickRange(await this.request('GET', p + this.ws(sheetId) + '/usedRange')); }
    async getRange(sheetId, address, session) {
      const p = await this.itemPath();
      return GraphSource.pickRange(await this.request('GET', p + this.ws(sheetId) + "/range(address='" + address + "')", undefined, this.sh(session)));
    }
    async beginSession() { const p = await this.itemPath(); return (await this.request('POST', p + '/createSession', { persistChanges: true })).id; }
    async endSession(session) {
      if (!session) return;
      const p = await this.itemPath();
      try { await this.request('POST', p + '/closeSession', {}, this.sh(session)); } catch (e) { /* session expires anyway */ }
    }
    async writeRange(sheetId, address, values, session) {
      const p = await this.itemPath();
      return this.request('PATCH', p + this.ws(sheetId) + "/range(address='" + address + "')", { values }, this.sh(session));
    }
    async addTableRows(tableName, rows, session) {
      const p = await this.itemPath();
      return this.request('POST', p + '/tables/' + encodeURIComponent(tableName) + '/rows/add', { index: null, values: rows }, this.sh(session));
    }
    // ---- structure changes (one-time column setup) ----
    /** Excel's "insert table column": index = position of the new column (0-based); existing columns from there shift right. */
    async addTableColumn(tableName, index, name, session) {
      const p = await this.itemPath();
      return this.request('POST', p + '/tables/' + encodeURIComponent(tableName) + '/columns/add', { index, name }, this.sh(session));
    }
    async addWorksheet(name, session) {
      const p = await this.itemPath();
      const r = await this.request('POST', p + '/worksheets/add', { name }, this.sh(session));
      return { id: r.id, name: r.name };
    }
    async addTable(sheetName, address, session) {
      const p = await this.itemPath();
      const q = /^[A-Za-z0-9_]+$/.test(sheetName) ? sheetName : "'" + sheetName.replace(/'/g, "''") + "'";
      const r = await this.request('POST', p + '/tables/add', { address: q + '!' + address, hasHeaders: true }, this.sh(session));
      return { name: r.name, id: r.id };
    }
    async renameTable(oldName, newName, session) {
      const p = await this.itemPath();
      return this.request('PATCH', p + '/tables/' + encodeURIComponent(oldName), { name: newName }, this.sh(session));
    }
  }

  PV.GraphSource = GraphSource;
  PV.GraphError = GraphError;
  PV.encodeSharingUrl = encodeSharingUrl;
})();
