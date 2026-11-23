/* Probe Viewer – turns raw Excel ranges into sheet models and probes. Pure functions, shared by demo and live mode. */
(function () {
  'use strict';
  const PV = window.PV;

  const UNIT_TOKENS = new Set(['technique', 'class', 'compatibility']);

  function isUnitsRow(values) {
    let hits = 0;
    for (const v of values) {
      if (typeof v !== 'string') continue;
      const s = PV.clean(v).toLowerCase();
      if (/^\(.+\)$/.test(s) || UNIT_TOKENS.has(s)) hits++;
    }
    return hits >= 2;
  }
  function isEmptyRow(values) {
    return values.every((v) => v == null || PV.clean(v) === '');
  }

  /** Resolve role -> column index using header aliases. */
  function resolveRoles(headers, aliasMap) {
    const norm = headers.map(PV.normHeader);
    const roles = {};
    for (const role of Object.keys(aliasMap)) {
      const aliases = aliasMap[role].map(PV.normHeader);
      for (const a of aliases) {
        const i = norm.indexOf(a);
        if (i >= 0) { roles[role] = i; break; }
      }
    }
    return roles;
  }

  function aliasesFor(cfg, family, side) {
    const map = Object.assign({}, cfg.columns);
    const over = family && family[side === 'stock' ? 'stockColumns' : 'lifetimeColumns'];
    if (over) for (const k of Object.keys(over)) map[k] = over[k];
    return map;
  }

  /**
   * raw = {sheetName, sheetId, tableName, address, values, text, numberFormat, formulas}
   * Header is the first row of a table range; for plain ranges it is family.headerRow (sheet row number).
   */
  PV.buildSheetModel = function (raw, ctx) {
    const { cfg, family, side, wb } = ctx;
    const addr = PV.parseAddress(raw.address);
    let hIdx = 0;
    if (!raw.tableName) {
      const hr = (family && family.headerRow) || addr.r1;
      hIdx = Math.max(0, hr - addr.r1);
    }
    const get = (arr, r) => (arr && arr[r]) || [];
    const headers = (raw.values[hIdx] || []).map((h) => (h == null ? '' : String(h)));
    const model = {
      wb, side, family: family ? family.key : null, sheetName: raw.sheetName, sheetId: raw.sheetId, tableName: raw.tableName || null,
      isTable: !!raw.tableName, c1: addr.c1, c2: addr.c2, headerRow: addr.r1 + hIdx,
      headers, labels: headers.map((h) => PV.clean(h)), units: headers.map(() => ''),
      roles: resolveRoles(headers, aliasesFor(cfg, family, side)),
      rows: [], warnings: [], stats: { unitsRows: 0, emptyRows: 0, numberText: {}, noId: 0, dupIds: [] },
    };
    for (let r = hIdx + 1; r < raw.values.length; r++) {
      const values = raw.values[r];
      const row = {
        idx: r - hIdx - 1, sheetRow: addr.r1 + r, values,
        text: get(raw.text, r), nf: get(raw.numberFormat, r), formulas: get(raw.formulas, r), kind: 'data',
      };
      if (isEmptyRow(values)) { row.kind = 'empty'; model.stats.emptyRows++; }
      else if (isUnitsRow(values)) {
        row.kind = 'units'; model.stats.unitsRows++;
        values.forEach((v, c) => { const s = PV.clean(v); if (/^\(.+\)$/.test(s)) model.units[c] = s; });
      }
      model.rows.push(row);
    }
    // ---- structural checks ----
    const where = (side === 'stock' ? 'Stock and Booking' : 'Probe Lifetime') + ' › ' + raw.sheetName;
    if (!model.isTable) model.warnings.push({ level: 'warn', where, msg: 'Sheet is a plain range, not an Excel Table – read-only in the viewer. Convert it with Ctrl+T.' });
    if (model.roles.id == null) model.warnings.push({ level: 'error', where, msg: 'No ID column found (expected "ID" or "Box ID").' });
    if (model.roles.status == null) model.warnings.push({ level: 'error', where, msg: 'No Status column found.' });
    headers.forEach((h, c) => {
      if (PV.clean(h) === '') {
        const isStatus = model.roles.status === c;
        model.warnings.push({ level: 'warn', where, msg: 'Column ' + PV.colLetter(model.c1 + c) + ' has an empty header' + (isStatus ? ' (used as Status by config – rename it to "Status").' : '.') });
      }
    });
    if (headers.some((h) => /boing/i.test(h))) model.warnings.push({ level: 'info', where, msg: 'Header typo "Boing info".' });
    // numbers stored as text + duplicate IDs
    const seen = new Map();
    for (const row of model.rows) {
      if (row.kind !== 'data') continue;
      row.values.forEach((v, c) => {
        if (PV.isNumberLikeText(v) && c !== model.roles.id) model.stats.numberText[c] = (model.stats.numberText[c] || 0) + 1;
      });
      const id = model.roles.id != null ? PV.clean(row.values[model.roles.id]) : '';
      if (!id) { model.stats.noId++; continue; }
      const k = PV.rowKey(model, row);
      seen.set(k, (seen.get(k) || 0) + 1);
    }
    for (const [k, n] of seen) if (n > 1) model.stats.dupIds.push(k + ' ×' + n);
    if (model.stats.dupIds.length) model.warnings.push({ level: 'warn', where, msg: 'Duplicate IDs: ' + model.stats.dupIds.slice(0, 8).join(', ') + (model.stats.dupIds.length > 8 ? ' …' : '') + ' – import and joins skip them.' });
    if (model.stats.noId) model.warnings.push({ level: 'info', where, msg: model.stats.noId + ' row(s) with data but no ID are hidden.' });
    const nt = Object.entries(model.stats.numberText);
    if (nt.length) model.warnings.push({ level: 'info', where, msg: 'Numbers stored as text: ' + nt.map(([c, n]) => '"' + model.labels[c] + '" ×' + n).join(', ') + ' (shown as numbers, cells unchanged).' });
    if (model.stats.unitsRows) model.warnings.push({ level: 'info', where, msg: model.stats.unitsRows + ' units row(s) inside the table are used as column units and hidden.' });
    return model;
  };

  /** Row identity used to match Stock <-> Lifetime and imports. */
  PV.rowKey = function (model, row) {
    const id = model.roles.id != null ? PV.clean(row.values[model.roles.id]).toUpperCase() : '';
    const fam = window.PV_CONFIG.families.find((f) => f.key === model.family);
    if (fam && fam.boxLevel && model.roles.position != null) {
      const pos = PV.clean(row.values[model.roles.position]);
      if (pos && PV.toNum(pos) !== null) return id + '#' + PV.toNum(pos);
    }
    return id;
  };

  /**
   * Named lists from the data sheet: each column holds blocks separated by blank cells;
   * the first cell of a block is the list name, the rest are its values.
   */
  PV.parseNamedLists = function (values) {
    const lists = {};
    if (!values || !values.length) return lists;
    const cols = Math.max(...values.map((r) => r.length));
    for (let c = 0; c < cols; c++) {
      let name = null;
      for (let r = 0; r < values.length; r++) {
        const s = PV.clean(values[r][c]);
        if (!s) { name = null; continue; }
        if (name == null) { name = s.toLowerCase(); if (!lists[name]) lists[name] = []; continue; }
        if (!lists[name].includes(s)) lists[name].push(s);
      }
    }
    return lists;
  };

  // ---- probes ----
  const LIFETIME_FIRST = new Set(['owner', 'department', 'mech', 'ele', 'mag', 'state', 'current', 'asOf', 'history']);

  PV.cell = function (ref, role) {
    if (!ref) return null;
    const c = ref.model.roles[role];
    if (c == null) return null;
    return { v: ref.row.values[c], t: ref.row.text[c], nf: ref.row.nf[c], f: ref.row.formulas[c], col: c };
  };
  /** Value of a role on a probe: lifetime-owned fields come from Lifetime first, everything else from Stock first. */
  PV.pget = function (p, role) {
    const order = LIFETIME_FIRST.has(role) ? [p.life, p.stock] : [p.stock, p.life];
    for (const ref of order) { const c = PV.cell(ref, role); if (c && PV.clean(c.t !== '' && c.t != null ? c.t : c.v) !== '') return c; }
    for (const ref of order) { const c = PV.cell(ref, role); if (c) return c; }
    return null;
  };
  PV.ptext = function (p, role) {
    const c = PV.pget(p, role);
    if (!c) return '';
    return PV.displayCell(c.v, c.t, c.nf);
  };
  PV.displayCell = function (v, t, nf) {
    const d = PV.cellDate(v, nf, null);
    if (d && typeof v === 'number') return PV.isoDate(d);
    if (v instanceof Date) return PV.isoDate(v);
    const s = PV.clean(t != null && t !== '' ? t : v);
    return s;
  };

  /** Build probes for one family from its Stock and Lifetime sheet models. */
  PV.buildProbes = function (family, stockModel, lifeModel) {
    const probes = [];
    const lifeIndex = new Map();
    const lifeUsed = new Set();
    if (lifeModel) {
      const counts = new Map();
      for (const row of lifeModel.rows) {
        if (row.kind !== 'data') continue;
        const k = PV.rowKey(lifeModel, row);
        if (!k) continue;
        counts.set(k, (counts.get(k) || 0) + 1);
        lifeIndex.set(k, row);
      }
      for (const [k, n] of counts) if (n > 1) lifeIndex.set(k, null); // ambiguous -> no join
    }
    const mk = (stockRef, lifeRef) => {
      const ref = stockRef || lifeRef;
      const id = PV.clean(ref.row.values[ref.model.roles.id]);
      const p = {
        uid: family.key + '|' + (stockRef ? 'S' + stockRef.row.sheetRow : 'L' + lifeRef.row.sheetRow),
        family, stock: stockRef, life: lifeRef, id, key: PV.rowKey(ref.model, ref.row), flags: [],
      };
      return p;
    };
    if (stockModel) {
      for (const row of stockModel.rows) {
        if (row.kind !== 'data') continue;
        if (stockModel.roles.id == null || !PV.clean(row.values[stockModel.roles.id])) continue;
        const k = PV.rowKey(stockModel, row);
        let lifeRef = null;
        if (lifeModel && lifeIndex.has(k)) {
          const lr = lifeIndex.get(k);
          if (lr) { lifeRef = { model: lifeModel, row: lr }; lifeUsed.add(lr); }
        }
        const p = mk({ model: stockModel, row }, lifeRef);
        if (lifeModel && lifeIndex.has(k) && !lifeIndex.get(k)) p.flags.push('lifetime-ambiguous');
        probes.push(p);
      }
    }
    if (lifeModel) {
      for (const row of lifeModel.rows) {
        if (row.kind !== 'data' || lifeUsed.has(row)) continue;
        if (lifeModel.roles.id == null || !PV.clean(row.values[lifeModel.roles.id])) continue;
        const p = mk(null, { model: lifeModel, row });
        if (stockModel) p.flags.push('not-in-stock');
        probes.push(p);
      }
    }
    // mark duplicate keys within this family
    const kc = new Map();
    probes.forEach((p) => kc.set(p.key, (kc.get(p.key) || 0) + 1));
    probes.forEach((p) => { if (kc.get(p.key) > 1) p.flags.push('duplicate-id'); });
    return probes;
  };

  /** ID used in the Log (box-level rows carry their position: BOX#3). */
  PV.logId = function (p) { return p.key.includes('#') ? p.id + '#' + p.key.split('#')[1] : p.id; };

  /** Normalised status key: 'to come' | 'in stock' | 'booked' | 'sent out' | 'missing' | '' | other. */
  PV.statusKey = function (p) {
    const k = PV.key(PV.ptext(p, 'status'));
    return k;
  };
  PV.STATUS_ORDER = ['in stock', 'booked', 'sent out', 'missing', 'to come'];

  /** State fields present for this probe: [{role, label, value}] e.g. Mech./Ele. for NPC, Mag. for NPM, State otherwise. */
  PV.stateFields = function (p) {
    const out = [];
    const model = p.life ? p.life.model : null;
    const roles = ['mech', 'ele', 'mag', 'state', 'current'];
    for (const r of roles) {
      if (model && model.roles[r] != null) out.push({ role: r, label: model.labels[model.roles[r]], value: PV.ptext(p, r) });
    }
    return out;
  };

  PV.normDept = function (v) {
    const s = PV.clean(v);
    if (!s) return '';
    const a = window.PV_CONFIG.departmentAliases || {};
    return a[s.toLowerCase()] || s;
  };
})();
