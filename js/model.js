/* Probe Viewer – turns raw Excel ranges into sheet models and probes. Pure functions, shared by demo and live mode. */
(function () {
  'use strict';
  const PV = window.PV;

  const UNIT_TOKENS = new Set(['technique', 'class', 'compatibility']);
  const STATE_ROLES = ['mech', 'ele', 'mag', 'state'];

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

  function aliasesFor(cfg, family) {
    const map = Object.assign({}, cfg.columns);
    if (family && family.stockColumns) for (const k of Object.keys(family.stockColumns)) map[k] = family.stockColumns[k];
    return map;
  }
  /** Resolve role -> column index using header aliases. */
  function resolveRoles(headers, aliasMap) {
    const norm = headers.map(PV.normHeader);
    const roles = {};
    for (const role of Object.keys(aliasMap)) {
      for (const a of aliasMap[role].map(PV.normHeader)) {
        const i = norm.indexOf(a);
        if (i >= 0) { roles[role] = i; break; }
      }
    }
    return roles;
  }
  /** Role of a header name (e.g. 'Mech.' -> 'mech'), or null. */
  PV.roleOfHeader = function (cfg, family, header) {
    const n = PV.normHeader(header);
    const map = aliasesFor(cfg, family);
    for (const role of Object.keys(map)) if (map[role].map(PV.normHeader).includes(n)) return role;
    return null;
  };

  /** The new columns this type should have (Price, Project and the usage group), as header names. */
  PV.wantedColumns = function (cfg, family) {
    return { price: 'Price', after: ['Project', 'Owner', 'Department', 'Box'].concat(family.stateColumns || []).concat(['As of', 'History']) };
  };

  /**
   * raw = {sheetName, sheetId, tableName, address, values, text, numberFormat, formulas}
   * The header is the first row of the table range.
   */
  PV.buildSheetModel = function (raw, ctx) {
    const { cfg, family } = ctx;
    const addr = PV.parseAddress(raw.address);
    const get = (arr, r) => (arr && arr[r]) || [];
    const headers = (raw.values[0] || []).map((h) => (h == null ? '' : String(h)));
    const model = {
      family: family ? family.key : null, sheetName: raw.sheetName, sheetId: raw.sheetId, tableName: raw.tableName || null,
      isTable: !!raw.tableName, c1: addr.c1, c2: addr.c2, headerRow: addr.r1,
      headers, labels: headers.map((h) => PV.clean(h)), units: headers.map(() => ''),
      roles: resolveRoles(headers, aliasesFor(cfg, family)),
      rows: [], warnings: [], missing: [], stats: { unitsRows: 0, emptyRows: 0, numberText: {}, noId: 0, dupIds: [] },
    };
    for (let r = 1; r < raw.values.length; r++) {
      const values = raw.values[r];
      const row = { idx: r - 1, sheetRow: addr.r1 + r, values, text: get(raw.text, r), nf: get(raw.numberFormat, r), formulas: get(raw.formulas, r), kind: 'data' };
      if (isEmptyRow(values)) { row.kind = 'empty'; model.stats.emptyRows++; }
      else if (isUnitsRow(values)) {
        row.kind = 'units'; model.stats.unitsRows++;
        values.forEach((v, c) => { const s = PV.clean(v); if (/^\(.+\)$/.test(s)) model.units[c] = s; });
      }
      model.rows.push(row);
    }
    // ---- structural checks ----
    const where = 'Stock and Booking › ' + raw.sheetName;
    if (!model.isTable) model.warnings.push({ level: 'error', where, msg: 'Sheet is not an Excel Table – read-only in the viewer. Convert it with Ctrl+T.' });
    if (model.roles.id == null) model.warnings.push({ level: 'error', where, msg: 'No ID column found (expected "ID" or "Box ID").' });
    if (model.roles.status == null) model.warnings.push({ level: 'error', where, msg: 'No Status column found.' });
    headers.forEach((h, c) => {
      if (PV.clean(h) === '') {
        const isStatus = model.roles.status === c;
        model.warnings.push({ level: 'warn', where, msg: 'Column ' + PV.colLetter(model.c1 + c) + ' has an empty header' + (isStatus ? ' (used as Status by config – rename it to "Status").' : '.') });
      }
    });
    if (family) {
      const want = PV.wantedColumns(cfg, family);
      [want.price].concat(want.after).forEach((h) => {
        const role = PV.roleOfHeader(cfg, family, h);
        if (role == null || model.roles[role] == null) model.missing.push(h);
      });
      if (model.missing.length) model.warnings.push({ level: 'info', where, msg: 'New columns not added yet: ' + model.missing.join(', ') + '. Data checks → "Add missing columns" adds them.' });
    }
    // numbers stored as text + duplicate IDs
    const seen = new Map();
    for (const row of model.rows) {
      if (row.kind !== 'data') continue;
      row.values.forEach((v, c) => { if (PV.isNumberLikeText(v) && c !== model.roles.id) model.stats.numberText[c] = (model.stats.numberText[c] || 0) + 1; });
      const id = model.roles.id != null ? PV.clean(row.values[model.roles.id]) : '';
      if (!id) { model.stats.noId++; continue; }
      const k = PV.rowKey(model, row);
      seen.set(k, (seen.get(k) || 0) + 1);
    }
    for (const [k, n] of seen) if (n > 1) model.stats.dupIds.push(k + ' ×' + n);
    if (model.stats.dupIds.length) model.warnings.push({ level: 'warn', where, msg: 'Duplicate IDs: ' + model.stats.dupIds.slice(0, 8).join(', ') + (model.stats.dupIds.length > 8 ? ' …' : '') + ' – the import skips them.' });
    if (model.stats.noId) model.warnings.push({ level: 'info', where, msg: model.stats.noId + ' row(s) with data but no ID are hidden.' });
    const nt = Object.entries(model.stats.numberText);
    if (nt.length) model.warnings.push({ level: 'info', where, msg: 'Numbers stored as text: ' + nt.map(([c, n]) => '"' + model.labels[c] + '" ×' + n).join(', ') + ' (shown as numbers, cells unchanged).' });
    if (model.stats.unitsRows) model.warnings.push({ level: 'info', where, msg: model.stats.unitsRows + ' units row(s) inside the table are used as column units and hidden.' });
    return model;
  };

  /**
   * Where to insert the missing columns, as Graph "columns/add" operations applied in order.
   * Price goes in front of "Booked for" (end of the datasheet part); Project and the usage group go right after the
   * booking columns (Booked for, Date of booking), i.e. in front of Note. Excel cannot insert after the last column,
   * so a column that would land at the very end is reported as manual.
   */
  PV.planColumns = function (cfg, family, model) {
    const cur = model.headers.slice();
    const roleAt = (name) => { const role = PV.roleOfHeader(cfg, family, name); if (role == null) return -1; return cur.findIndex((h) => PV.roleOfHeader(cfg, family, h) === role); };
    const ops = [], manual = [];
    const insert = (index, header) => {
      if (index > cur.length - 1) { manual.push(header); return false; }
      ops.push({ index, header }); cur.splice(index, 0, header); return true;
    };
    const want = PV.wantedColumns(cfg, family);
    if (roleAt(want.price) < 0) {
      let i = roleAt('Booked for');
      if (i < 0) i = roleAt('Note');
      if (i < 0) i = cur.length - 1;
      insert(i, want.price);
    }
    let anchor = Math.max(roleAt('Booked for'), roleAt('Date of booking')) + 1;
    if (anchor <= 0) { const n = roleAt('Note'); anchor = n >= 0 ? n : cur.length; }
    for (const h of want.after) {
      const i = roleAt(h);
      if (i >= 0) { anchor = Math.max(anchor, i + 1); continue; }
      if (insert(anchor, h)) anchor++;
    }
    return { ops, manual };
  };

  /** Row identity used for imports and the log (box-level rows: Box ID + position). */
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
  PV.cell = function (ref, role) {
    if (!ref) return null;
    const c = ref.model.roles[role];
    if (c == null) return null;
    return { v: ref.row.values[c], t: ref.row.text[c], nf: ref.row.nf[c], f: ref.row.formulas[c], col: c };
  };
  PV.pget = function (p, role) { return PV.cell(p.stock, role); };
  PV.ptext = function (p, role) {
    const c = PV.pget(p, role);
    return c ? PV.displayCell(c.v, c.t, c.nf) : '';
  };
  PV.displayCell = function (v, t, nf) {
    const d = PV.cellDate(v, nf, null);
    if (d && typeof v === 'number') return PV.isoDate(d);
    if (v instanceof Date) return PV.isoDate(v);
    return PV.clean(t != null && t !== '' ? t : v);
  };

  /** Probes of one sheet: every data row with an ID. */
  PV.buildProbes = function (family, model) {
    const probes = [];
    if (!model) return probes;
    for (const row of model.rows) {
      if (row.kind !== 'data') continue;
      if (model.roles.id == null || !PV.clean(row.values[model.roles.id])) continue;
      probes.push({ uid: family.key + '|' + row.sheetRow, family, stock: { model, row }, id: PV.clean(row.values[model.roles.id]), key: PV.rowKey(model, row), flags: [] });
    }
    const kc = new Map();
    probes.forEach((p) => kc.set(p.key, (kc.get(p.key) || 0) + 1));
    probes.forEach((p) => { if (kc.get(p.key) > 1) p.flags.push('duplicate-id'); });
    return probes;
  };

  /** ID used in the Log (box-level rows carry their position: BOX#3). */
  PV.logId = function (p) { return p.key.includes('#') ? p.id + '#' + p.key.split('#')[1] : p.id; };

  /** Normalised status key: 'to come' | 'in stock' | 'booked' | 'sent out' | 'missing' | '' | other. */
  PV.statusKey = function (p) { return PV.key(PV.ptext(p, 'status')); };

  /** State columns present for this probe: [{role, label, value}] e.g. Mech./Ele. for NPC, Mag. for NPM, State otherwise. */
  PV.stateFields = function (p) {
    const m = p.stock.model;
    return STATE_ROLES.filter((r) => m.roles[r] != null).map((r) => ({ role: r, label: m.labels[m.roles[r]], value: PV.ptext(p, r) }));
  };

  PV.normDept = function (v) {
    const s = PV.clean(v);
    if (!s) return '';
    const a = window.PV_CONFIG.departmentAliases || {};
    return a[s.toLowerCase()] || s;
  };
})();
