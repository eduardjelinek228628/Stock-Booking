/* Probe Viewer – loads Stock and Booking through a data source and performs all writes (same code in demo and live). */
(function () {
  'use strict';
  const PV = window.PV;

  class ConflictError extends Error {
    constructor(probe, fresh) { super('Probe ' + probe.id + ' was changed in Excel meanwhile.'); this.probe = probe; this.fresh = fresh; this.isConflict = true; }
  }

  function sameValue(a, b) {
    const na = PV.toNum(a), nb = PV.toNum(b);
    if (typeof a === 'number' || typeof b === 'number') { if (na !== null && nb !== null) return Math.abs(na - nb) < 1e-9; }
    return PV.clean(a) === PV.clean(b);
  }

  class Store {
    constructor(cfg, source, user) {
      this.cfg = cfg; this.source = source; this.user = user;
      this.families = []; this.probes = []; this.lists = {}; this.log = null; this.logRows = [];
      this.warnings = []; this.semafor = []; this.loadedAt = null; this.sheets = null;
    }

    get mode() { return this.source.kind; }
    get writesAllowed() { return this.mode === 'demo' || !!this.cfg.liveWritesEnabled; }
    can(action) {
      const r = this.user.roles || [];
      if (!this.writesAllowed) return false;
      if (action === 'status' || action === 'setup') return r.includes('operations');
      if (action === 'import') return r.includes('operations') || r.includes('importer');
      return true; // book, state
    }

    // ---------------- loading ----------------
    async load() {
      const cfg = this.cfg;
      this.warnings = [];
      try { this.sheets = await this.source.listWorksheets(); }
      catch (e) { this.sheets = null; this.warnings.push({ level: 'error', where: cfg.workbook.label, msg: 'Cannot open the workbook: ' + e.message }); }
      const models = await PV.pool(cfg.families, 4, (fam) => this.loadSheet(fam).catch((e) => {
        this.warnings.push({ level: 'error', where: cfg.workbook.label + ' › ' + fam.sheet, msg: e.message });
        return null;
      }));
      this.families = cfg.families.map((fam, i) => ({ cfg: fam, stock: models[i], probes: [] }));
      await Promise.all([this.loadLists(), this.loadLog(), this.loadSemafor()]);
      this.rebuild();
      this.loadedAt = new Date();
    }

    findSheet(name) {
      if (!this.sheets) return null;
      return this.sheets.find((x) => x.name === name) || this.sheets.find((x) => PV.normHeader(x.name) === PV.normHeader(name)) || null;
    }

    async loadSheet(fam) {
      if (!this.sheets) return null;
      const sh = this.findSheet(fam.sheet);
      if (!sh) { this.warnings.push({ level: 'error', where: this.cfg.workbook.label, msg: 'Expected sheet "' + fam.sheet + '" is missing.' }); return null; }
      const tables = await this.source.listTables(sh.id);
      if (!tables.length) { this.warnings.push({ level: 'error', where: this.cfg.workbook.label + ' › ' + sh.name, msg: 'No Excel Table on this sheet – convert the probe list into a Table (Ctrl+T).' }); return null; }
      if (tables.length > 1) this.warnings.push({ level: 'info', where: this.cfg.workbook.label + ' › ' + sh.name, msg: tables.length + ' tables on this sheet – using "' + tables[0].name + '".' });
      const raw = await this.source.getTableRange(tables[0].name);
      raw.tableName = tables[0].name; raw.sheetName = sh.name; raw.sheetId = sh.id;
      return PV.buildSheetModel(raw, { cfg: this.cfg, family: fam });
    }

    async loadLists() {
      this.lists = {};
      const sh = this.findSheet(this.cfg.dataSheet);
      if (!sh) return;
      try { Object.assign(this.lists, PV.parseNamedLists((await this.source.getUsedRange(sh.id)).values)); }
      catch (e) { this.warnings.push({ level: 'warn', where: this.cfg.workbook.label + ' › ' + this.cfg.dataSheet, msg: e.message }); }
    }
    /** Allowed values for a field: data sheet list, then family/config fallback. Projects also include values already used. */
    list(field, famCfg) {
      const name = this.cfg.valueLists[field];
      const fromSheet = (name && this.lists[name.toLowerCase()]) || [];
      let out = fromSheet.map((v) => v.trim());
      if (!out.length) out = ((famCfg && famCfg.stateLists && famCfg.stateLists[field]) || this.cfg.fallbackLists[field] || []).slice();
      if (field === 'project' || field === 'department') {
        const seen = new Set(out.map((x) => x.toLowerCase()));
        for (const p of this.probes) { const v = PV.clean(PV.ptext(p, field)); if (v && !seen.has(v.toLowerCase())) { seen.add(v.toLowerCase()); out.push(v); } }
        if (field === 'project') out.sort((a, b) => a.localeCompare(b));
      }
      return out;
    }

    async loadLog() {
      this.log = null; this.logRows = [];
      const L = this.cfg.log;
      const sh = this.findSheet(L.sheet);
      if (!sh) { this.warnings.push({ level: 'warn', where: this.cfg.workbook.label, msg: 'No "' + L.sheet + '" sheet – changes made in the viewer are not logged yet. Data checks → "Add missing columns" creates it.' }); return; }
      const tables = await this.source.listTables(sh.id);
      const t = tables.find((x) => x.name === L.table) || tables[0];
      if (!t) { this.warnings.push({ level: 'warn', where: this.cfg.workbook.label + ' › ' + L.sheet, msg: 'No Table on the Log sheet.' }); return; }
      const raw = await this.source.getTableRange(t.name);
      const headers = raw.values[0].map(PV.normHeader);
      const idx = {};
      this.cfg.logColumns.forEach((h) => { idx[h] = headers.indexOf(PV.normHeader(h)); });
      const missing = this.cfg.logColumns.filter((h) => idx[h] < 0);
      if (missing.length) this.warnings.push({ level: 'warn', where: 'Log', msg: 'Log table lacks columns: ' + missing.join(', ') });
      this.log = { sheetId: sh.id, table: t.name, headers: raw.values[0], idx };
      const get = (row, h) => (idx[h] >= 0 ? row[idx[h]] : '');
      for (let r = 1; r < raw.values.length; r++) {
        const row = raw.values[r];
        if (row.every((v) => PV.clean(v) === '')) continue;
        const tsRaw = get(row, 'Timestamp');
        let ts = null;
        if (typeof tsRaw === 'number') ts = PV.serialToDate(tsRaw);
        else { const m = PV.clean(tsRaw).match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2}))?/); if (m) ts = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0))); }
        if (!ts) continue;
        this.logRows.push({ ts, user: PV.clean(get(row, 'User')), sheet: PV.clean(get(row, 'Sheet')), id: PV.clean(get(row, 'Probe ID')),
          field: PV.clean(get(row, 'Field')), old: PV.clean(get(row, 'Old value')), new: PV.clean(get(row, 'New value')), project: PV.clean(get(row, 'Project')) });
      }
      this.logRows.sort((a, b) => a.ts - b.ts);
    }

    async loadSemafor() {
      this.semafor = [];
      const sh = this.findSheet(this.cfg.semaforSheet);
      if (!sh) return;
      try {
        const tables = await this.source.listTables(sh.id);
        if (!tables.length) return;
        const r = await this.source.getTableRange(tables[0].name);
        this.semafor = r.values.slice(1).filter((x) => PV.clean(x[0]));
      } catch (e) { /* optional */ }
    }

    rebuild() {
      this.probes = [];
      for (const f of this.families) { f.probes = PV.buildProbes(f.cfg, f.stock); this.probes.push(...f.probes); }
      this.byUid = new Map(this.probes.map((p) => [p.uid, p]));
    }

    /** Re-read only the given sheets (by family key) and the Log. */
    async reload(famKeys) {
      for (const k of famKeys) { const f = this.family(k); if (f) f.stock = await this.loadSheet(f.cfg); }
      await this.loadLog();
      this.rebuild();
      this.loadedAt = new Date();
    }

    family(key) { return this.families.find((f) => f.cfg.key === key); }

    // ---------------- writing primitives ----------------
    async tx(fn) {
      const session = await this.source.beginSession();
      try { return await fn(session); }
      finally { await this.source.endSession(session); }
    }
    rowAddress(ref) { return PV.makeAddress(ref.model.c1, ref.row.sheetRow, ref.model.c2, ref.row.sheetRow); }
    async verify(ref, probe, session) {
      const fresh = await this.source.getRange(ref.model.sheetId, this.rowAddress(ref), session);
      const vals = fresh.values[0];
      if (ref.row.values.some((v, i) => !sameValue(v, vals[i]))) {
        throw new ConflictError(probe, ref.model.headers.map((h, i) => ({ header: PV.clean(h), was: PV.clean(ref.row.text[i] || ref.row.values[i]), now: PV.clean(fresh.text[0][i] || vals[i]) })).filter((x) => x.was !== x.now));
      }
    }
    assertWritable(ref) {
      if (!ref.model.isTable) throw new Error('"' + ref.model.sheetName + '" is not an Excel Table – convert it first (Ctrl+T).');
    }
    async writeCell(ref, col, value, session) {
      const f = ref.row.formulas[col];
      if (typeof f === 'string' && f.startsWith('=')) throw new Error('Column "' + ref.model.labels[col] + '" contains a formula – the viewer does not overwrite formulas.');
      await this.source.writeRange(ref.model.sheetId, PV.makeAddress(ref.model.c1 + col, ref.row.sheetRow), [[value]], session);
    }
    logEntry(sheet, id, field, oldV, newV, project) {
      const d = new Date();
      const ts = PV.todayISO() + ' ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0') + ':' + String(d.getSeconds()).padStart(2, '0');
      return { Timestamp: ts, User: this.user.initials, Sheet: sheet, 'Probe ID': id, Field: field, 'Old value': oldV, 'New value': newV, Project: project || '' };
    }
    async appendLog(entries, session) {
      if (!this.log || !entries.length) return;
      // Log cells are always text: a leading apostrophe stops Excel turning "2026-11-23" or "36.6" into a date/number.
      const asText = (v) => { const s = v == null ? '' : String(v); return /^[=+\-@]/.test(s) || /^[\d\s.,:\/-]+$/.test(s) ? "'" + s : s; };
      const rows = entries.map((e) => this.log.headers.map((h) => { const k = this.cfg.logColumns.find((c) => PV.normHeader(c) === PV.normHeader(h)); return k ? asText(e[k]) : null; }));
      await this.source.addTableRows(this.log.table, rows, session); // one request for all entries
    }
    historyPrefix(text) { return PV.todayISO() + ' ' + this.user.initials + ': ' + text; }
    prependHistory(old, text) {
      const o = PV.clean(old);
      return this.historyPrefix(text) + (o ? this.cfg.historySeparator + o : '');
    }
    personValue() { return this.cfg.bookedForFormat === 'name' ? this.user.name : this.user.initials; }

    /** Re-read the row, then write the changed cells one by one. changes: {role: value}. Returns log entries. */
    async applyRow(probe, changes, session, project) {
      const ref = probe.stock;
      this.assertWritable(ref);
      await this.verify(ref, probe, session);
      const out = [];
      for (const role of Object.keys(changes)) {
        const col = ref.model.roles[role];
        const nv = changes[role];
        if (col == null || nv === undefined || nv === null) continue;
        const old = PV.displayCell(ref.row.values[col], ref.row.text[col], ref.row.nf[col]);
        if (sameValue(old, nv)) continue;
        await this.writeCell(ref, col, typeof nv === 'string' ? PV.safeText(nv) : nv, session);
        out.push(this.logEntry(ref.model.sheetName, PV.logId(probe), ref.model.labels[col], old, role === 'history' ? '(note added)' : String(nv), project));
      }
      return out;
    }
    missingNote(probe, roles) {
      const m = probe.stock.model;
      const miss = roles.filter((r) => m.roles[r] == null);
      return miss.length ? 'Not written (no column yet in ' + m.sheetName + '): ' + miss.join(', ') + '.' : '';
    }

    // ---------------- actions ----------------
    /** Book a probe: one row in Stock and Booking (status, booking, project, owner, state, history) plus the log. */
    async book(probe, opts) {
      if (!this.can('book')) throw new Error('Writing is disabled.');
      if (PV.statusKey(probe) !== 'in stock') throw new Error('Only probes "in stock" can be booked.');
      const project = PV.clean(opts.project);
      const today = PV.todayISO();
      const m = probe.stock.model;
      const ch = { status: 'booked', bookedFor: this.personValue(), bookingDate: today, project: project || undefined, owner: this.personValue(), department: this.user.department || undefined, asOf: today,
        history: this.prependHistory(PV.ptext(probe, 'history'), 'booked' + (project ? ' for ' + project : '')) };
      for (const r of ['mech', 'ele', 'mag', 'state']) if (m.roles[r] != null && !PV.ptext(probe, r)) ch[r] = 'new';
      const report = { logged: !!this.log, notes: [] };
      const note = this.missingNote(probe, ['project', 'owner', 'history']);
      if (note) report.notes.push(note);
      return this.tx(async (session) => {
        const entries = await this.applyRow(probe, ch, session, project);
        try { await this.appendLog(entries, session); } catch (e) { report.logged = false; report.notes.push('Log not written: ' + e.message); }
        return report;
      }).finally(() => this.reload([probe.family.key]).catch(() => {}));
    }

    /** Operations only: change Status (Semafor rule). */
    async setStatus(probe, status) {
      if (!this.can('status')) throw new Error('Only operations may change Status (Semafor rule).');
      return this.tx(async (session) => {
        await this.appendLog(await this.applyRow(probe, { status, asOf: PV.todayISO() }, session), session);
        return { ok: true };
      }).finally(() => this.reload([probe.family.key]).catch(() => {}));
    }

    /** Engineer update (phone or desktop): state columns + a dated note in History. */
    async updateState(probe, changes, note) {
      if (!this.can('state')) throw new Error('Writing is disabled.');
      const m = probe.stock.model;
      if (m.roles.history == null && !Object.keys(changes).length) throw new Error('No History column in ' + m.sheetName + ' yet – Data checks → "Add missing columns".');
      for (const r of Object.keys(changes)) if (m.roles[r] == null) throw new Error('No ' + r + ' column in ' + m.sheetName + ' yet – Data checks → "Add missing columns".');
      const parts = [];
      for (const r of Object.keys(changes)) {
        const old = PV.ptext(probe, r);
        if (PV.clean(old) !== PV.clean(changes[r])) parts.push(m.labels[m.roles[r]] + ' ' + (old || '–') + ' → ' + changes[r]);
      }
      const n = PV.clean(note);
      if (!parts.length && !n) throw new Error('Nothing to save.');
      const histText = parts.concat(n ? [n] : []).join('; ');
      return this.tx(async (session) => {
        const ch = Object.assign({}, changes, { asOf: PV.todayISO(), history: this.prependHistory(PV.ptext(probe, 'history'), histText) });
        await this.appendLog(await this.applyRow(probe, ch, session), session);
        return { ok: true };
      }).finally(() => this.reload([probe.family.key]).catch(() => {}));
    }

    /** Test-results import: items = [{probe, cells: [{col, value}]}] for one family's sheet. */
    async importResults(famKey, items) {
      if (!this.can('import')) throw new Error('Import is limited to operations / importers.');
      const results = { written: 0, failed: [] };
      await this.tx(async (session) => {
        const entries = [];
        for (const it of items) {
          try {
            const ref = it.probe.stock;
            this.assertWritable(ref);
            await this.verify(ref, it.probe, session);
            for (const c of it.cells) {
              const old = PV.displayCell(ref.row.values[c.col], ref.row.text[c.col], ref.row.nf[c.col]);
              await this.writeCell(ref, c.col, typeof c.value === 'string' ? PV.safeText(c.value) : c.value, session);
              entries.push(this.logEntry(ref.model.sheetName, PV.logId(it.probe), ref.model.labels[c.col], old, String(c.value), 'test results import'));
              results.written++;
            }
          } catch (e) { results.failed.push({ id: it.probe.id, error: e.message }); }
        }
        try { await this.appendLog(entries, session); } catch (e) { results.failed.push({ id: '(log)', error: e.message }); }
      });
      await this.reload([famKey]);
      return results;
    }

    // ---------------- one-time setup (operations) ----------------
    /** What "Add missing columns" would do: per sheet the column inserts, plus the Log sheet. */
    setupPlan() {
      const sheets = [];
      for (const f of this.families) {
        if (!f.stock || !f.stock.isTable) continue;
        const p = PV.planColumns(this.cfg, f.cfg, f.stock);
        if (p.ops.length || p.manual.length) sheets.push({ family: f, ops: p.ops, manual: p.manual });
      }
      return { sheets, log: !this.log };
    }
    /** Insert the missing columns with Excel's own "insert table column" and create the Log table. */
    async setupColumns() {
      if (!this.can('setup')) throw new Error('Only operations may change the workbook structure.');
      const plan = this.setupPlan();
      const done = [];
      await this.tx(async (session) => {
        for (const s of plan.sheets) {
          const m = s.family.stock;
          for (const op of s.ops) {
            await this.source.addTableColumn(m.tableName, op.index, op.header, session);
            // set the header text explicitly (Excel names new columns "Column1" …)
            await this.source.writeRange(m.sheetId, PV.makeAddress(m.c1 + op.index, m.headerRow), [[op.header]], session);
            done.push(m.sheetName + ': ' + op.header);
          }
        }
        if (plan.log) {
          const L = this.cfg.log;
          let sh = this.findSheet(L.sheet);
          if (!sh) sh = await this.source.addWorksheet(L.sheet, session);
          const last = PV.colLetter(this.cfg.logColumns.length);
          await this.source.writeRange(sh.id, 'A1:' + last + '1', [this.cfg.logColumns], session);
          const t = await this.source.addTable(sh.name, 'A1:' + last + '1', session);
          if (t.name !== L.table) await this.source.renameTable(t.name, L.table, session);
          done.push('Log sheet with table ' + L.table);
        }
      });
      await this.load();
      return { done, manual: plan.sheets.flatMap((s) => s.manual.map((h) => s.family.stock.sheetName + ': ' + h)) };
    }
  }

  PV.Store = Store;
  PV.ConflictError = ConflictError;
})();
