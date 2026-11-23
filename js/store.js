/* Probe Viewer – loads both workbooks through a data source and performs all writes (same code in demo and live). */
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
      this.warnings = []; this.semafor = []; this.loadedAt = null; this.sheetIndex = {};
    }

    get mode() { return this.source.kind; }
    get writesAllowed() { return this.mode === 'demo' || !!this.cfg.liveWritesEnabled; }
    can(action) {
      const r = this.user.roles || [];
      if (!this.writesAllowed) return false;
      if (action === 'status') return r.includes('operations');
      if (action === 'import') return r.includes('operations') || r.includes('importer');
      return true; // book, state
    }

    // ---------------- loading ----------------
    async load() {
      const cfg = this.cfg;
      this.warnings = [];
      const wbKeys = Object.keys(cfg.workbooks);
      for (const k of wbKeys) {
        try { this.sheetIndex[k] = await this.source.listWorksheets(k); }
        catch (e) { this.sheetIndex[k] = null; this.warnings.push({ level: 'error', where: cfg.workbooks[k].label, msg: 'Cannot open workbook: ' + e.message }); }
      }
      const jobs = [];
      for (const fam of cfg.families) {
        if (fam.stockSheet) jobs.push({ fam, side: 'stock', wb: 'stock', sheet: fam.stockSheet });
        if (fam.lifetimeSheet) jobs.push({ fam, side: 'lifetime', wb: 'lifetime', sheet: fam.lifetimeSheet });
      }
      const models = await PV.pool(jobs, 4, (j) => this.loadSheet(j.wb, j.sheet, j.fam, j.side).catch((e) => {
        this.warnings.push({ level: 'error', where: cfg.workbooks[j.wb].label + ' › ' + j.sheet, msg: e.message });
        return null;
      }));
      this.families = cfg.families.map((fam) => ({ cfg: fam, stock: null, life: null, probes: [] }));
      jobs.forEach((j, i) => {
        const f = this.families.find((x) => x.cfg === j.fam);
        if (j.side === 'stock') f.stock = models[i]; else f.life = models[i];
      });
      await Promise.all([this.loadLists(), this.loadLog(), this.loadSemafor()]);
      this.rebuild();
      this.loadedAt = new Date();
    }

    findSheet(wb, name) {
      const list = this.sheetIndex[wb];
      if (!list) return null;
      let s = list.find((x) => x.name === name);
      if (!s) s = list.find((x) => PV.normHeader(x.name) === PV.normHeader(name));
      return s || null;
    }

    async loadSheet(wb, sheetName, fam, side) {
      if (!this.sheetIndex[wb]) return null;
      const sh = this.findSheet(wb, sheetName);
      if (!sh) { this.warnings.push({ level: 'error', where: this.cfg.workbooks[wb].label, msg: 'Expected sheet "' + sheetName + '" is missing.' }); return null; }
      const tables = await this.source.listTables(wb, sh.id);
      let raw;
      if (tables.length) {
        if (tables.length > 1) this.warnings.push({ level: 'info', where: this.cfg.workbooks[wb].label + ' › ' + sh.name, msg: tables.length + ' tables on this sheet – using "' + tables[0].name + '".' });
        raw = await this.source.getTableRange(wb, tables[0].name);
        raw.tableName = tables[0].name;
      } else {
        raw = await this.source.getUsedRange(wb, sh.id);
        raw.tableName = null;
      }
      raw.sheetName = sh.name; raw.sheetId = sh.id;
      return PV.buildSheetModel(raw, { cfg: this.cfg, family: fam, side, wb });
    }

    async loadLists() {
      this.lists = {};
      for (const wb of ['lifetime', 'stock']) { // stock wins on conflicts (loaded last)
        const sh = this.sheetIndex[wb] && this.findSheet(wb, this.cfg.dataSheet);
        if (!sh) { if (this.sheetIndex[wb]) this.warnings.push({ level: 'warn', where: this.cfg.workbooks[wb].label, msg: 'No "' + this.cfg.dataSheet + '" sheet – using built-in value lists.' }); continue; }
        try {
          const r = await this.source.getUsedRange(wb, sh.id);
          Object.assign(this.lists, PV.parseNamedLists(r.values));
        } catch (e) { this.warnings.push({ level: 'warn', where: this.cfg.workbooks[wb].label + ' › ' + this.cfg.dataSheet, msg: e.message }); }
      }
    }
    /** Allowed values for a field (from the data sheet; falls back to config). */
    list(field, famCfg) {
      const name = (famCfg && famCfg.valueLists && famCfg.valueLists[field]) || this.cfg.valueLists[field];
      const l = name && this.lists[name.toLowerCase()];
      if (l && l.length) return l.map((v) => v.trim());
      const fb = this.cfg.fallbackLists[field] || (['mech', 'ele', 'mag', 'current'].includes(field) ? this.cfg.fallbackLists.state : []);
      return fb.slice();
    }

    async loadLog() {
      this.log = null; this.logRows = [];
      const L = this.cfg.log;
      const sh = this.sheetIndex[L.workbook] && this.findSheet(L.workbook, L.sheet);
      if (!sh) { this.warnings.push({ level: 'warn', where: this.cfg.workbooks[L.workbook].label, msg: 'No "' + L.sheet + '" sheet with table "' + L.table + '" – changes made in the viewer are not logged and the statistics fall back to "Date of booking" only.' }); return; }
      const tables = await this.source.listTables(L.workbook, sh.id);
      const t = tables.find((x) => x.name === L.table);
      if (!t) { this.warnings.push({ level: 'warn', where: this.cfg.workbooks[L.workbook].label + ' › ' + L.sheet, msg: 'Table "' + L.table + '" not found on the Log sheet.' }); return; }
      const raw = await this.source.getTableRange(L.workbook, t.name);
      const headers = raw.values[0].map(PV.normHeader);
      const idx = {};
      this.cfg.logColumns.forEach((h) => { idx[h] = headers.indexOf(PV.normHeader(h)); });
      const missing = this.cfg.logColumns.filter((h) => idx[h] < 0);
      if (missing.length) this.warnings.push({ level: 'warn', where: 'Log', msg: 'Log table lacks columns: ' + missing.join(', ') });
      this.log = { wb: L.workbook, sheetId: sh.id, table: t.name, headers: raw.values[0], idx };
      const get = (row, h) => (idx[h] >= 0 ? row[idx[h]] : '');
      for (let r = 1; r < raw.values.length; r++) {
        const row = raw.values[r];
        if (row.every((v) => PV.clean(v) === '')) continue;
        const tsRaw = get(row, 'Timestamp');
        let ts = null;
        if (typeof tsRaw === 'number') ts = PV.serialToDate(tsRaw);
        else { const m = PV.clean(tsRaw).match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2}))?/); if (m) ts = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0))); }
        if (!ts) continue;
        this.logRows.push({ ts, user: PV.clean(get(row, 'User')), wb: PV.clean(get(row, 'Workbook')), sheet: PV.clean(get(row, 'Sheet')), id: PV.clean(get(row, 'Probe ID')),
          field: PV.clean(get(row, 'Field')), old: PV.clean(get(row, 'Old value')), new: PV.clean(get(row, 'New value')), project: PV.clean(get(row, 'Project')) });
      }
      this.logRows.sort((a, b) => a.ts - b.ts);
    }

    async loadSemafor() {
      this.semafor = [];
      const sh = this.sheetIndex.stock && this.findSheet('stock', this.cfg.semaforSheet);
      if (!sh) return;
      try {
        const tables = await this.source.listTables('stock', sh.id);
        if (!tables.length) return;
        const r = await this.source.getTableRange('stock', tables[0].name);
        this.semafor = r.values.slice(1).filter((x) => PV.clean(x[0]));
      } catch (e) { /* optional */ }
    }

    rebuild() {
      this.probes = [];
      for (const f of this.families) {
        f.probes = PV.buildProbes(f.cfg, f.stock, f.life);
        this.probes.push(...f.probes);
      }
      this.byUid = new Map(this.probes.map((p) => [p.uid, p]));
    }

    async reload(parts) {
      // parts: [{wb, sheet}] – re-read only what changed
      for (const part of parts) {
        const f = this.families.find((x) => (part.wb === 'stock' ? x.cfg.stockSheet : x.cfg.lifetimeSheet) === part.sheet);
        if (!f) continue;
        const m = await this.loadSheet(part.wb, part.sheet, f.cfg, part.wb === 'stock' ? 'stock' : 'lifetime');
        if (part.wb === 'stock') f.stock = m; else f.life = m;
      }
      await this.loadLog();
      // drop warnings that came from the re-read sheets' builders; rebuild from models
      this.rebuild();
      this.loadedAt = new Date();
    }

    family(key) { return this.families.find((f) => f.cfg.key === key); }

    // ---------------- writing primitives ----------------
    async tx(wbKeys, fn) {
      const sessions = {};
      for (const k of wbKeys) sessions[k] = await this.source.beginSession(k);
      try { return await fn(sessions); }
      finally { for (const k of wbKeys) await this.source.endSession(k, sessions[k]); }
    }
    rowAddress(ref) { return PV.makeAddress(ref.model.c1, ref.row.sheetRow, ref.model.c2, ref.row.sheetRow); }
    async verify(ref, probe, session) {
      const fresh = await this.source.getRange(ref.model.wb, ref.model.sheetId, this.rowAddress(ref), session);
      const vals = fresh.values[0];
      const differs = ref.row.values.some((v, i) => !sameValue(v, vals[i]));
      if (differs) throw new ConflictError(probe, ref.model.headers.map((h, i) => ({ header: PV.clean(h), was: PV.clean(ref.row.text[i] || ref.row.values[i]), now: PV.clean(fresh.text[0][i] || vals[i]) })).filter((x) => x.was !== x.now));
    }
    assertWritable(ref) {
      if (!ref.model.isTable) throw new Error('"' + ref.model.sheetName + '" is not an Excel Table – convert it first (Ctrl+T).');
    }
    async writeCell(ref, col, value, session) {
      const f = ref.row.formulas[col];
      if (typeof f === 'string' && f.startsWith('=')) throw new Error('Column "' + ref.model.labels[col] + '" contains a formula – the viewer does not overwrite formulas.');
      const addr = PV.makeAddress(ref.model.c1 + col, ref.row.sheetRow);
      await this.source.writeRange(ref.model.wb, ref.model.sheetId, addr, [[value]], session);
    }
    logEntry(wb, sheet, id, field, oldV, newV, project) {
      const d = new Date();
      const ts = PV.todayISO() + ' ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0') + ':' + String(d.getSeconds()).padStart(2, '0');
      return { Timestamp: ts, User: this.user.initials, Workbook: wb, Sheet: sheet, 'Probe ID': id, Field: field, 'Old value': oldV, 'New value': newV, Project: project || '' };
    }
    async appendLog(entries, sessions) {
      if (!this.log || !entries.length) return;
      const s = sessions[this.log.wb] || null;
      for (const e of entries) {
        // Log cells are always text: a leading apostrophe stops Excel turning "2026-11-23" or "36.6" into a date/number.
        const asText = (v) => { const s = v == null ? '' : String(v); return /^[=+\-@]/.test(s) || /^[\d\s.,:\/-]+$/.test(s) ? "'" + s : s; };
        const values = this.log.headers.map((h) => { const k = this.cfg.logColumns.find((c) => PV.normHeader(c) === PV.normHeader(h)); return k ? asText(e[k]) : null; });
        await this.source.addTableRow(this.log.wb, this.log.table, values, s);
      }
    }
    historyPrefix(text) { return PV.todayISO() + ' ' + this.user.initials + ': ' + text; }
    prependHistory(old, text) {
      const o = PV.clean(old);
      return this.historyPrefix(text) + (o ? this.cfg.historySeparator + o : '');
    }
    personValue() { return this.cfg.bookedForFormat === 'name' ? this.user.name : this.user.initials; }

    /** Write changes into one row; changes: {role: value}. Returns log entries. */
    async applyRow(ref, probe, changes, session, project) {
      this.assertWritable(ref);
      await this.verify(ref, probe, session);
      const out = [];
      for (const role of Object.keys(changes)) {
        const col = ref.model.roles[role];
        if (col == null) continue;
        const nv = changes[role];
        if (nv === undefined || nv === null) continue;
        const old = PV.displayCell(ref.row.values[col], ref.row.text[col], ref.row.nf[col]);
        if (sameValue(old, nv)) continue;
        await this.writeCell(ref, col, typeof nv === 'string' ? PV.safeText(nv) : nv, session);
        out.push(this.logEntry(ref.model.side, ref.model.sheetName, PV.logId(probe), ref.model.labels[col], old, role === 'history' ? '(note added)' : String(nv), project));
      }
      return out;
    }

    /** New Lifetime row for a probe that has none: copy matching datasheet columns from Stock. */
    lifetimeValuesFor(probe, lifeModel, extra) {
      const stockRef = probe.stock;
      return lifeModel.headers.map((h, c) => {
        const role = Object.keys(lifeModel.roles).find((r) => lifeModel.roles[r] === c);
        if (role && extra[role] !== undefined) return typeof extra[role] === 'string' ? PV.safeText(extra[role]) : extra[role];
        if (!stockRef) return null;
        const n = PV.normHeader(h);
        let sc = stockRef.model.headers.findIndex((x) => PV.normHeader(x) === n);
        if (sc < 0 && role && stockRef.model.roles[role] != null && !['status', 'owner', 'department', 'history', 'asOf', 'note'].includes(role)) sc = stockRef.model.roles[role];
        if (sc < 0 || n === '') return null;
        const v = stockRef.row.values[sc];
        if (v === '' || v == null) return null;
        return typeof v === 'string' ? PV.safeText(v) : v;
      });
    }
    async ensureLifetime(probe, fam, extra, sessions, project) {
      const lm = fam.life;
      if (!lm) throw new Error('No Probe Lifetime sheet configured for ' + fam.cfg.key + '.');
      this.assertWritable({ model: lm });
      const values = this.lifetimeValuesFor(probe, lm, extra);
      const s = sessions.lifetime;
      const empty = lm.rows.find((r) => r.kind === 'empty' && r.formulas.every((f) => !(typeof f === 'string' && f.startsWith('='))));
      if (empty) {
        // re-check the row is still empty, then fill it
        const fresh = await this.source.getRange('lifetime', lm.sheetId, PV.makeAddress(lm.c1, empty.sheetRow, lm.c2, empty.sheetRow), s);
        if (fresh.values[0].some((v) => PV.clean(v) !== '')) throw new ConflictError(probe, [{ header: 'Lifetime row ' + empty.sheetRow, was: '(empty)', now: 'filled by someone else' }]);
        await this.source.writeRange('lifetime', lm.sheetId, PV.makeAddress(lm.c1, empty.sheetRow, lm.c2, empty.sheetRow), [values], s);
      } else {
        await this.source.addTableRow('lifetime', lm.tableName, values, s);
      }
      return [this.logEntry('lifetime', lm.sheetName, PV.logId(probe), 'row', '', 'created', project)];
    }

    // ---------------- actions ----------------
    /** Book a probe: Stock row + Lifetime row in one action, plus log. */
    async book(probe, opts) {
      if (!this.can('book')) throw new Error('Writing is disabled.');
      if (!probe.stock) throw new Error('This probe type is not in Stock and Booking.');
      if (PV.statusKey(probe) !== 'in stock') throw new Error('Only probes "in stock" can be booked.');
      const fam = this.family(probe.family.key);
      const project = PV.clean(opts.project);
      const today = PV.todayISO();
      const report = { stock: false, lifetime: null, logged: !!this.log, notes: [] };
      const sm = probe.stock.model;
      if (sm.roles.bookedFor == null) report.notes.push('No "Booked for" column in ' + sm.sheetName + '.');
      if (project && sm.roles.project == null) report.notes.push('No "Project" column in ' + sm.sheetName + ' – the project is only kept in the Log.');
      const wbs = ['stock'].concat(fam.life && fam.life.isTable ? ['lifetime'] : []);
      return this.tx(wbs, async (sessions) => {
        let entries = await this.applyRow(probe.stock, probe, { status: 'booked', bookedFor: this.personValue(), bookingDate: today, project: project || undefined }, sessions.stock, project);
        report.stock = true;
        if (fam.life && fam.life.isTable) {
          try {
            const stateInit = {};
            for (const r of ['mech', 'ele', 'mag', 'state']) if (fam.life.roles[r] != null) stateInit[r] = 'new';
            const hist = 'booked' + (project ? ' for ' + project : '');
            if (probe.life) {
              const ch = { status: 'booked', owner: this.user.initials, department: this.user.department || undefined, asOf: today, history: this.prependHistory(PV.ptext(probe, 'history'), hist) };
              for (const r of Object.keys(stateInit)) if (!PV.clean(PV.cell(probe.life, r) && PV.cell(probe.life, r).v)) ch[r] = 'new';
              entries = entries.concat(await this.applyRow(probe.life, probe, ch, sessions.lifetime, project));
            } else {
              entries = entries.concat(await this.ensureLifetime(probe, fam, Object.assign({ status: 'booked', owner: this.user.initials, department: this.user.department || '', asOf: today, history: this.historyPrefix(hist) }, stateInit), sessions, project));
            }
            report.lifetime = true;
          } catch (e) { report.lifetime = false; report.lifetimeError = e.message; }
        } else if (fam.cfg.lifetimeSheet) report.notes.push('Lifetime sheet "' + fam.cfg.lifetimeSheet + '" is not writable (missing or not a Table).');
        try { await this.appendLog(entries, sessions); } catch (e) { report.logged = false; report.notes.push('Log not written: ' + e.message); }
        return report;
      }).finally(() => this.reload([{ wb: 'stock', sheet: fam.cfg.stockSheet }].concat(fam.cfg.lifetimeSheet ? [{ wb: 'lifetime', sheet: fam.cfg.lifetimeSheet }] : [])).catch(() => {}));
    }

    /** Operations only: change Status in Stock (and keep Lifetime Status in sync). */
    async setStatus(probe, status) {
      if (!this.can('status')) throw new Error('Only operations may change Status (Semafor rule).');
      const fam = this.family(probe.family.key);
      const ref = probe.stock || probe.life;
      const wbs = [ref.model.wb].concat(probe.stock && probe.life ? ['lifetime'] : []);
      return this.tx(wbs, async (sessions) => {
        let entries = await this.applyRow(ref, probe, { status }, sessions[ref.model.wb]);
        if (probe.stock && probe.life && probe.life.model.isTable) entries = entries.concat(await this.applyRow(probe.life, probe, { status, asOf: PV.todayISO() }, sessions.lifetime));
        await this.appendLog(entries, sessions);
        return { ok: true };
      }).finally(() => this.reload([{ wb: ref.model.wb, sheet: ref.model.sheetName }].concat(probe.stock && probe.life ? [{ wb: 'lifetime', sheet: fam.cfg.lifetimeSheet }] : [])).catch(() => {}));
    }

    /** Engineer update from the phone: state fields + note -> Lifetime (As of + History). */
    async updateState(probe, changes, note) {
      if (!this.can('state')) throw new Error('Writing is disabled.');
      const fam = this.family(probe.family.key);
      if (!fam.life) throw new Error('No Probe Lifetime sheet for ' + fam.cfg.key + '.');
      this.assertWritable({ model: fam.life });
      const parts = [];
      for (const r of Object.keys(changes)) {
        const old = PV.ptext(probe, r);
        if (PV.clean(old) !== PV.clean(changes[r])) parts.push(fam.life.labels[fam.life.roles[r]] + ' ' + (old || '–') + ' → ' + changes[r]);
      }
      const n = PV.clean(note);
      if (!parts.length && !n) throw new Error('Nothing to save.');
      const histText = parts.concat(n ? [n] : []).join('; ');
      return this.tx(['lifetime'], async (sessions) => {
        let entries;
        if (probe.life) {
          const ch = Object.assign({}, changes, { asOf: PV.todayISO(), history: this.prependHistory(PV.ptext(probe, 'history'), histText) });
          entries = await this.applyRow(probe.life, probe, ch, sessions.lifetime);
        } else {
          entries = await this.ensureLifetime(probe, fam, Object.assign({ status: PV.ptext(probe, 'status'), owner: this.user.initials, department: this.user.department || '', asOf: PV.todayISO(), history: this.historyPrefix(histText) }, changes), sessions);
        }
        const all = Object.assign({}, sessions);
        if (this.log && !all[this.log.wb]) all[this.log.wb] = null;
        await this.appendLog(entries, all);
        return { ok: true };
      }).finally(() => this.reload([{ wb: 'lifetime', sheet: fam.cfg.lifetimeSheet }]).catch(() => {}));
    }

    /** Test-results import: items = [{probe, cells: [{col, value}]}] for one family's Stock sheet. */
    async importResults(famKey, items) {
      if (!this.can('import')) throw new Error('Import is limited to operations / importers.');
      const fam = this.family(famKey);
      const results = { written: 0, failed: [] };
      await this.tx(['stock'], async (sessions) => {
        let entries = [];
        for (const it of items) {
          try {
            const ref = it.probe.stock;
            this.assertWritable(ref);
            await this.verify(ref, it.probe, sessions.stock);
            for (const c of it.cells) {
              const old = PV.displayCell(ref.row.values[c.col], ref.row.text[c.col], ref.row.nf[c.col]);
              await this.writeCell(ref, c.col, typeof c.value === 'string' ? PV.safeText(c.value) : c.value, sessions.stock);
              entries.push(this.logEntry('stock', ref.model.sheetName, PV.logId(it.probe), ref.model.labels[c.col], old, String(c.value), 'test results import'));
              results.written++;
            }
          } catch (e) { results.failed.push({ id: it.probe.id, error: e.message }); }
        }
        try { await this.appendLog(entries, sessions); } catch (e) { results.failed.push({ id: '(log)', error: e.message }); }
      });
      await this.reload([{ wb: 'stock', sheet: fam.cfg.stockSheet }]);
      return results;
    }
  }

  PV.Store = Store;
  PV.ConflictError = ConflictError;
})();
