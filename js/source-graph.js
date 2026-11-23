/* Probe Viewer – live data source: Microsoft Graph Excel API (delegated, signed-in user). */
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

  /**
   * opts: { getToken: async () => accessToken, workbooks: config.workbooks, fetchImpl?, baseUrl?, sleep? }
   */
  class GraphSource {
    constructor(opts) {
      this.kind = 'live';
      this.getToken = opts.getToken;
      this.workbooks = opts.workbooks;
      this.fetch = opts.fetchImpl || window.fetch.bind(window);
      this.base = opts.baseUrl || 'https://graph.microsoft.com/v1.0';
      this.sleep = opts.sleep || ((ms) => new Promise((r) => setTimeout(r, ms)));
      this.items = {}; // wbKey -> '/drives/{d}/items/{i}'
      this.sheets = {}; // wbKey -> [{id,name}]
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
          throw new GraphError(res.status, err.code || 'http' + res.status, err.message || ('HTTP ' + res.status), path);
        }
        return json;
      }
    }

    async itemPath(wbKey) {
      if (this.items[wbKey]) return this.items[wbKey];
      const wb = this.workbooks[wbKey];
      if (!wb) throw new Error('Workbook "' + wbKey + '" is not configured.');
      let driveId = wb.driveId, itemId = wb.itemId;
      if (!(driveId && itemId)) {
        if (!wb.sharingUrl) throw new Error('Set sharingUrl (or driveId + itemId) for "' + wb.label + '" in js/config.js.');
        const item = await this.request('GET', '/shares/' + encodeSharingUrl(wb.sharingUrl) + '/driveItem?$select=id,name,parentReference');
        driveId = item.parentReference && item.parentReference.driveId;
        itemId = item.id;
        if (!driveId || !itemId) throw new Error('Could not resolve the file behind the sharing link of "' + wb.label + '".');
      }
      this.items[wbKey] = '/drives/' + encodeURIComponent(driveId) + '/items/' + encodeURIComponent(itemId);
      return this.items[wbKey];
    }

    sessionHeaders(session) { return session ? { 'workbook-session-id': session } : undefined; }

    async listWorksheets(wbKey) {
      const p = await this.itemPath(wbKey);
      const r = await this.request('GET', p + '/workbook/worksheets?$select=id,name,visibility');
      this.sheets[wbKey] = r.value.map((w) => ({ id: w.id, name: w.name, visibility: w.visibility }));
      return this.sheets[wbKey];
    }
    async listTables(wbKey, sheetId) {
      const p = await this.itemPath(wbKey);
      const r = await this.request('GET', p + '/workbook/worksheets/' + encodeURIComponent(sheetId) + '/tables?$select=id,name');
      return r.value.map((t) => ({ name: t.name, id: t.id }));
    }
    static pickRange(r) {
      return { address: r.address, values: r.values, text: r.text, numberFormat: r.numberFormat, formulas: r.formulas };
    }
    async getTableRange(wbKey, tableName) {
      const p = await this.itemPath(wbKey);
      // one request returns header + body with values, text, numberFormat and formulas
      return GraphSource.pickRange(await this.request('GET', p + '/workbook/tables/' + encodeURIComponent(tableName) + '/range'));
    }
    async getUsedRange(wbKey, sheetId) {
      const p = await this.itemPath(wbKey);
      return GraphSource.pickRange(await this.request('GET', p + '/workbook/worksheets/' + encodeURIComponent(sheetId) + '/usedRange'));
    }
    async getRange(wbKey, sheetId, address, session) {
      const p = await this.itemPath(wbKey);
      return GraphSource.pickRange(await this.request('GET', p + '/workbook/worksheets/' + encodeURIComponent(sheetId) + "/range(address='" + address + "')", undefined, this.sessionHeaders(session)));
    }
    async beginSession(wbKey) {
      const p = await this.itemPath(wbKey);
      const r = await this.request('POST', p + '/workbook/createSession', { persistChanges: true });
      return r.id;
    }
    async endSession(wbKey, session) {
      if (!session) return;
      const p = await this.itemPath(wbKey);
      try { await this.request('POST', p + '/workbook/closeSession', {}, this.sessionHeaders(session)); } catch (e) { /* session expires anyway */ }
    }
    async writeRange(wbKey, sheetId, address, values, session) {
      const p = await this.itemPath(wbKey);
      return this.request('PATCH', p + '/workbook/worksheets/' + encodeURIComponent(sheetId) + "/range(address='" + address + "')", { values }, this.sessionHeaders(session));
    }
    async addTableRow(wbKey, tableName, values, session) {
      const p = await this.itemPath(wbKey);
      return this.request('POST', p + '/workbook/tables/' + encodeURIComponent(tableName) + '/rows/add', { index: null, values: [values] }, this.sessionHeaders(session));
    }
  }

  PV.GraphSource = GraphSource;
  PV.GraphError = GraphError;
  PV.encodeSharingUrl = encodeSharingUrl;
})();
