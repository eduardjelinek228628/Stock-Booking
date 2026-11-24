/*
 * Probe Viewer – demo data source. A tiny in-memory "Excel" with the same sheets, Tables and headers
 * as Stock and Booking, filled with SYNTHETIC data (no real probes, customers or people).
 * It implements the same interface as GraphSource, so demo mode runs exactly the same read/write code.
 * The demo workbook already has the new columns (Price, Project, Owner … History) and the Log sheet;
 * build(today, { legacy: true }) gives the workbook as it is today, without them (used by the tests of the column setup).
 */
(function () {
  'use strict';
  const PV = window.PV;
  const STORE_KEY = 'probeViewer.demo.v2';
  const DATE_NF = 'd.m.yyyy';

  // ---------- mini workbook ----------
  function addSheet(book, name) {
    const s = { id: '{demo-' + book.sheets.length + '}', name, cells: {}, tables: [] };
    book.sheets.push(s);
    return s;
  }
  function setCell(s, r, c, v, nf, f) {
    if (v === undefined) return;
    const k = r + ':' + c;
    if ((v === null || v === '') && !f) { delete s.cells[k]; return; }
    s.cells[k] = { v, nf: nf || 'General', f: f || null };
  }
  function getCell(s, r, c) { return s.cells[r + ':' + c] || null; }
  function cellText(cell) {
    if (!cell || cell.v == null) return '';
    if (typeof cell.v === 'number' && PV.isDateFormat(cell.nf)) {
      const d = PV.serialToDate(cell.v);
      return d.getUTCDate() + '.' + (d.getUTCMonth() + 1) + '.' + d.getUTCFullYear();
    }
    return String(cell.v);
  }
  function quoted(name) { return /^[A-Za-z0-9_]+$/.test(name) ? name : "'" + name.replace(/'/g, "''") + "'"; }
  function readBlock(s, c1, r1, c2, r2) {
    const values = [], text = [], numberFormat = [], formulas = [];
    for (let r = r1; r <= r2; r++) {
      const vr = [], tr = [], nr = [], fr = [];
      for (let c = c1; c <= c2; c++) {
        const cell = getCell(s, r, c);
        vr.push(cell && cell.v != null ? cell.v : '');
        tr.push(cellText(cell));
        nr.push(cell ? cell.nf : 'General');
        fr.push(cell && cell.f ? cell.f : (cell && cell.v != null ? cell.v : ''));
      }
      values.push(vr); text.push(tr); numberFormat.push(nr); formulas.push(fr);
    }
    return { address: quoted(s.name) + '!' + PV.makeAddress(c1, r1, c2, r2), values, text, numberFormat, formulas };
  }
  /** Header row + rows (arrays) at row r0, registered as an Excel Table when tableName is given. */
  function writeTable(s, tableName, r0, headers, rows, emptyTail) {
    headers.forEach((h, c) => setCell(s, r0, c + 1, h));
    rows.forEach((row, i) => row.forEach((cell, c) => {
      if (cell == null) return;
      if (typeof cell === 'object') setCell(s, r0 + 1 + i, c + 1, cell.v, cell.nf, cell.f);
      else setCell(s, r0 + 1 + i, c + 1, cell);
    }));
    if (tableName) s.tables.push({ name: tableName, c1: 1, r1: r0, c2: headers.length, r2: r0 + rows.length + (emptyTail || 0) });
  }
  const D = (serial) => ({ v: serial, nf: DATE_NF });

  // ---------- synthetic content ----------
  const PEOPLE = ['EJ', 'MP', 'JK', 'PK', 'ON', 'KK', 'RS'];
  const DEPT_OF = { EJ: 'Apps', MP: 'Apps', JK: 'Devel', PK: 'Apps', ON: 'Devel', KK: 'Devel', RS: 'Apps' };
  const GROUPS = ['RD', 'Apps', 'Development'];
  const CUSTOMERS = ['Customer Alpha', 'University Beta', 'Institute Gamma', 'Partner Delta', 'Lab Epsilon', 'Demo tour', 'Workshop'];
  const PROJECTS = ['SiC SSRM study', 'Customer demos', 'FEBID tips R&D', 'Training', 'Installations', 'Internal R&D'];
  const PRICE = { 'NPC-A': 55, 'NPC-Z': 60, 'NPC-Y': 70, 'NPC-X': 75, 'NPC-W': 80, 'NP_Pt-Sharp': 140, 'NP-Mag': 120, Akiyama: 95, PRSA: 160 };
  const NEW_COLS = new Set(['Price', 'Project', 'Owner', 'Department', 'Box', 'Mech.', 'Ele.', 'Mag.', 'State', 'As of', 'History']);

  /** Insert the new columns where they belong: Price before "Booked for", the rest after the booking columns. */
  function layout(base, fam, legacy) {
    if (legacy) return base;
    const usage = ['Project', 'Owner', 'Department', 'Box'].concat(fam.stateColumns).concat(['As of', 'History']);
    const out = [];
    for (const h of base) {
      if (h === 'Booked for') out.push('Price');
      out.push(h);
    }
    const after = Math.max(out.indexOf('Booked for'), out.indexOf('Date of booking')) + 1;
    out.splice(after, 0, ...usage);
    return out;
  }

  function build(today, opts) {
    opts = opts || {};
    const legacy = !!opts.legacy;
    const R = PV.rng(20261123);
    const T = PV.dateToSerial(today);
    const book = { sheets: [] };
    const log = [];
    const logEvent = (serial, user, sheet, id, field, oldV, newV, project) => {
      const d = PV.serialToDate(serial + R.next() * 0.4 + 0.33);
      log.push([PV.isoDate(d) + ' ' + String(d.getUTCHours()).padStart(2, '0') + ':' + String(d.getUTCMinutes()).padStart(2, '0'), user, sheet, id, field, oldV, newV, project || '']);
    };
    const famBy = (k) => window.PV_CONFIG.families.find((f) => f.key === k);
    const pad = (n, w) => String(n).padStart(w, '0');

    /** Life story of one probe: arrival in batches, wait, booking, maybe sent out / dead. Returns the cells it sets. */
    function story(fam, id, o) {
      const batch = Math.min(4, Math.floor((o.frac || 0) * 5));
      const offset = [...String(fam.key)].reduce((a, ch) => a + ch.charCodeAt(0), 0) % 70; // types are delivered on different days
      const arrival = T - 470 + offset + batch * 100 + R.int(0, 6);
      const res = { Status: 'in stock' };
      if (arrival > T - 3) { res.Status = 'to come'; return res; }
      logEvent(arrival, 'OP', fam.sheet, id, 'Status', 'to come', 'in stock', '');
      const wait = Math.round(-Math.log(1 - R.next() * 0.999) * (o.meanWait || 75)) + 2;
      const bd = arrival + wait;
      if (bd >= T) return res; // still in stock
      const outAfter = R.int(3, 45);
      res.Status = bd + outAfter < T && R.next() < 0.72 ? (R.next() < 0.04 ? 'missing!' : 'sent out') : 'booked';
      const person = R.next() < 0.55 ? R.pick(PEOPLE) : (R.next() < 0.5 ? R.pick(GROUPS) : R.pick(CUSTOMERS));
      const project = R.next() < 0.85 ? R.pick(PROJECTS) : '';
      res['Booked for'] = person; res['Date of booking'] = D(bd); res.Project = project || null;
      logEvent(bd, PEOPLE.includes(person) ? person : 'OP', fam.sheet, id, 'Status', 'in stock', 'booked', project);
      if (res.Status !== 'booked') logEvent(bd + outAfter, 'OP', fam.sheet, id, 'Status', 'booked', res.Status, project);
      // usage columns for engineers' probes
      if (PEOPLE.includes(person) || R.next() < 0.3) {
        const owner = PEOPLE.includes(person) ? person : (R.next() < 0.15 ? 'PB -> EJ' : 'RD');
        const who = owner.slice(-2);
        const endDay = bd + R.int(10, 220);
        const dead = endDay < T && R.next() < 0.45;
        const stateField = fam.stateColumns[0];
        res.Owner = owner;
        res.Department = DEPT_OF[who] || R.pick(['Apps', 'Devel', 'Dev', 'Development']);
        if (who === 'EJ') res.Box = 'Box ' + R.int(1, 3);
        else if (R.next() < 0.5) res.Box = 'Box ' + R.int(1, 6);
        res[stateField] = dead ? 'dead' : R.pick(['new', 'good', 'good', stateField === 'State' ? 'blunt' : 'bad', 'unknown']);
        if (fam.stateColumns[1]) res[fam.stateColumns[1]] = dead ? 'none' : R.pick(['good', 'good', 'weak', 'none', 'unknown']);
        res['As of'] = D(dead ? endDay : bd);
        const hist = [PV.isoDate(PV.serialToDate(bd)) + ' ' + who + ': booked' + (project ? ' for ' + project : '')];
        if (dead) hist.unshift(PV.isoDate(PV.serialToDate(endDay)) + ' ' + who + ': ' + stateField + ' good → dead; tip broken during approach');
        res.History = hist.join(' | ');
        if (dead) logEvent(endDay, who, fam.sheet, id, stateField, 'good', 'dead', project);
      }
      return res;
    }
    function rowFrom(headers, obj) { return headers.map((h) => (obj[h] === undefined ? null : obj[h])); }

    // ---- NPC sheets ----
    const npcSpecs = [
      { key: 'NPC-A', table: 'NP_Con_Si', prefix: 'A', n: 32, lengths: [500, 600, 850], tipHdr: 'Tip  ', variant: 'A', units: true },
      { key: 'NPC-Z', table: 'NP_Con_PolySi', prefix: 'Z', n: 80, lengths: [500, 600, 850], tipHdr: 'Tip ', variant: 'A', units: true },
      { key: 'NPC-Y', table: 'NP_Con_PolySi2', prefix: 'Y', n: 60, lengths: [270, 365, 600], tipHdr: 'Tip ', variant: 'Y', units: false },
      { key: 'NPC-X', table: 'NP_Con_PolySi27', prefix: 'X', n: 120, lengths: [310, 420, 600], tipHdr: 'Tip ', variant: 'X', units: true },
      { key: 'NPC-W', table: 'NP_Con_PolySi2710', prefix: 'W', n: 100, lengths: [310, 400, 630], tipHdr: 'Tip ', variant: 'X', units: true },
      { key: 'NP_Pt-Sharp', table: 'NP_Con_PolySi278', prefix: 'S', n: 36, lengths: [310, 400, 420, 600], tipHdr: 'Tip ', variant: 'P', units: true },
    ];
    for (const sp of npcSpecs) {
      const fam = famBy(sp.key);
      const sh = addSheet(book, sp.key);
      let base;
      if (sp.variant === 'A') base = ['ID', 'Status', 'Length', 'freso', 'Q factor', 'Stiffness', 'Primary ', 'Secondary', sp.tipHdr, 'Booked for', 'Note'];
      else if (sp.variant === 'Y') base = ['ID', 'Status', 'Length', 'freso', 'Q factor', 'Stiffness', 'Primary ', 'Secondary', 'Tip current', 'Tip ', 'Booked for', 'Note'];
      else if (sp.variant === 'X') base = ['ID', 'Status', 'Length', 'freso', 'Q factor', 'Stiffness', 'Primary ', 'Secondary', 'El. properties', 'Tip ', 'Notes', 'Booked for', 'Date of booking', 'Note'];
      else base = ['ID', ' ', 'Length', 'freso', 'Q factor', 'Stiffness', 'Primary ', 'Secondary', 'Tip current', 'Tip ', 'Booked for', 'Date of booking', 'Note'];
      const headers = layout(base, fam, legacy);
      const statusHdr = base[1];
      const rows = [];
      if (sp.units) rows.push(rowFrom(headers, { ID: sp.variant === 'A' ? null : 'A', Length: '(µm)', freso: '(kHz)', Stiffness: '(N/m)', 'Primary ': 'technique', Secondary: 'technique', [sp.tipHdr]: 'class' }));
      for (let i = 1; i <= sp.n; i++) {
        const id = sp.prefix + pad(i * (sp.key === 'NPC-X' ? 3 : 1) + (sp.key === 'NPC-W' ? 10 : 0), 3);
        const L = R.pick(sp.lengths);
        const fr = R.round(Math.max(8, 12000 / L - 5), 12000 / L + 25, 1);
        const k = R.round(0.6, L < 350 ? 140 : 40, 2);
        const tip = R.weighted([[1, 6], [2, 3], [3, sp.variant === 'X' ? 1 : 0]]);
        const o = { ID: id, Length: L, freso: fr, 'Q factor': R.int(120, 600), Stiffness: k, [sp.tipHdr]: tip, Price: PRICE[sp.key] + R.pick([0, 0, 5, -5]) };
        const prim = k <= 10 ? (tip === 2 ? 'C-AFM' : 'KPFM') : 'PFM, KPFM';
        const sRow = 2 + rows.length + 1;
        o['Primary '] = sp.key === 'NPC-Z' || sp.key === 'NPC-Y' ? { v: prim, f: '=IF(AND(F' + sRow + '<=10,' + PV.colLetter(headers.indexOf(sp.tipHdr) + 1) + sRow + '=2),"C-AFM","KPFM")' } : prim;
        if (R.next() < 0.3) o.Secondary = R.pick(['C-AFM', 'PFM', 'KPFM']);
        if (base.includes('Tip current')) o['Tip current'] = R.next() < 0.9 ? 'yes' : 'no';
        if (base.includes('El. properties')) o['El. properties'] = R.weighted([['Ok', 6], ['Ok via sample', 3], ['Not ok', 1]]);
        const st = story(fam, id, { frac: (i - 1) / sp.n });
        Object.assign(o, st); o[statusHdr] = st.Status;
        if (!base.includes('Date of booking')) delete o['Date of booking'];
        if (R.next() < 0.12) o.Note = R.pick(['demo', 'training - used', 'to fair and back, unused', 'low q-factor']);
        rows.push(rowFrom(headers, o));
      }
      // realistic quirks
      const col = (h) => headers.indexOf(h);
      if (sp.key === 'NPC-X') { rows[5][col('freso')] = String(rows[5][col('freso')]); rows[9][col('Q factor')] = ' -'; }
      if (sp.key === 'NPC-W') {
        if (rows[7][1] === 'sent out') rows[7][1] = 'Sent out';
        const dup = rows[20].slice(); dup[col('Note')] = 'duplicate row'; rows.push(dup);
        rows[30][1] = 'missing';
      }
      writeTable(sh, sp.table, 2, headers, rows, sp.key === 'NPC-Y' ? 6 : 2);
      setCell(sh, 1, 1, 'Probe'); setCell(sh, 1, 4, 'Datasheet'); setCell(sh, 1, 7, 'Selection info'); setCell(sh, 1, col('Booked for') + 1, 'Booking state');
    }

    // ---- NP-Mag ----
    {
      const fam = famBy('NP-Mag');
      const sh = addSheet(book, 'NP-Mag');
      const headers = layout(['ID', 'Status', 'Type', ' freso (Hz)      ', 'Q factor', 'Original Akiyama box    ', 'Tip', 'Booked for', 'Note'], fam, legacy);
      const rows = [];
      for (let i = 1; i <= 80; i++) {
        const id = i <= 6 ? 12000 + i * 3 : 'MAG-' + pad(i, 4);
        const o = { ID: id, Type: R.next() < 0.97 ? 'Zaragoza' : 'NanoSens.', ' freso (Hz)      ': R.int(39000, 43000), 'Q factor': R.next() < 0.85 ? R.int(1100, 1600) : null,
          'Original Akiyama box    ': 'BOX-M' + pad(Math.ceil(i / 10), 2) + '-' + pad(((i - 1) % 10) + 1, 2), Tip: R.weighted([[1, 5], [2, 1], ['?', 3]]), Price: PRICE['NP-Mag'] };
        Object.assign(o, story(fam, String(id), { frac: (i - 1) / 80 }));
        if (R.next() < 0.1) o.Note = 'weak MFM';
        rows.push(rowFrom(headers, o));
      }
      writeTable(sh, 'NP_Mag', 2, headers, rows);
    }

    // ---- Akiyama (box-level: 10-pack rows + single probes with position) ----
    {
      const fam = famBy('Akiyama');
      const sh = addSheet(book, 'Akiyama');
      const headers = layout(['Box ID', 'Status  ', 'Type', 'position', 'Booked for', 'Note', 'Sloupec1'], fam, legacy);
      const rows = [];
      const posCol = PV.colLetter(headers.indexOf('Type') + 1);
      for (let b = 1; b <= 16; b++) {
        const box = 'AK-' + pad(70000 + b * 37, 5) + '-' + pad(b, 2);
        const kind = R.weighted([['10-pack', 5], ['probe', 6], ['buffer', 1]]);
        if (kind === 'probe') {
          for (let pos = 1; pos <= 10; pos++) {
            const o = Object.assign({ 'Box ID': box, Type: 'probe', position: pos, Price: PRICE.Akiyama }, story(fam, box + '#' + pos, { frac: (b - 1) / 16, meanWait: 50 }));
            o['Status  '] = o.Status; delete o['Date of booking'];
            rows.push(rowFrom(headers, o));
          }
        } else {
          const sRow = 2 + rows.length + 1;
          const o = Object.assign({ 'Box ID': box, Type: kind, position: { v: kind === 'buffer' ? '-' : ' ', f: '=IF(' + posCol + sRow + ' = "buffer","-"," ")' }, Price: kind === '10-pack' ? PRICE.Akiyama * 10 : PRICE.Akiyama },
            story(fam, box, { frac: (b - 1) / 16, meanWait: 50 }));
          o['Status  '] = o.Status; delete o['Date of booking'];
          rows.push(rowFrom(headers, o));
        }
      }
      writeTable(sh, 'Akiyama', 2, headers, rows);
    }

    // ---- PRSA ----
    {
      const fam = famBy('PRSA');
      const sh = addSheet(book, 'PRSA');
      const headers = layout(['ID', 'Status', 'Type', ' freso (kHz)', 'Original box ID', 'Booked for', 'Note'], fam, legacy);
      const types = ['PRSA-L300-F50-Si-PCB', 'PRSA-L100-F500-Si-PCB', 'PRS-L70-F900-Si-PCB', 'PRSA-L400-F30-SCD-PCB'];
      const rows = [];
      for (let i = 1; i <= 26; i++) {
        const id = 'PRS-' + pad(i, 4);
        const o = Object.assign({ ID: id, Type: R.pick(types), ' freso (kHz)': R.next() < 0.6 ? R.round(30, 65, 2) : R.pick(['30 - 65', '15 - 40', '200-800']),
          'Original box ID': '2025-0' + (1 + (i % 6)) + '-12 Box ' + (1 + Math.floor(i / 9)), Price: PRICE.PRSA }, story(fam, id, { frac: (i - 1) / 26, meanWait: 120 }));
        delete o['Date of booking'];
        rows.push(rowFrom(headers, o));
      }
      writeTable(sh, 'PRSA', 2, headers, rows);
    }

    // ---- helper sheets ----
    {
      const sem = addSheet(book, 'Semafor');
      writeTable(sem, 'Tabulka10', 3, ['Probe', 'Status', 'Note'], [['Akiyama', 'All pass', null], ['PRSA', 'All pass', null], ['NPC', 'All tested', 'SEM imaging, frequency sweep, conductivity test'], ['NPM', 'All pass', null]]);
      setCell(sem, 9, 2, 'Probe status can only be changed by agreement between applications, development and operations');
      const data = addSheet(book, 'data');
      const lists = [['status', 'to come', 'in stock', 'booked', 'sent out', 'missing!'], ['methods', 'YES', 'NO']];
      if (!legacy) lists.push(['projects'].concat(PROJECTS));
      lists.forEach((list, c) => list.forEach((v, r) => setCell(data, r + 1, c + 1, v)));
      if (!legacy) {
        const lg = addSheet(book, 'Log');
        log.sort((a, b) => (a[0] < b[0] ? -1 : 1));
        writeTable(lg, 'ViewerLog', 1, window.PV_CONFIG.logColumns, log);
      }
    }
    return { book, builtFor: PV.isoDate(today), legacy };
  }

  // ---------- source ----------
  class DemoSource {
    constructor(opts) {
      opts = opts || {};
      this.kind = 'demo';
      this.persistent = opts.persistent !== false;
      const saved = this.persistent ? PV.storage.get(STORE_KEY) : null;
      this.data = saved && saved.book ? saved : build(new Date(), opts);
      this.sessions = 0;
      this.tableSeq = 1;
    }
    get book() { return this.data.book; }
    reset() { PV.storage.del(STORE_KEY); this.data = build(new Date()); }
    persist() { if (this.persistent) PV.storage.set(STORE_KEY, this.data); }
    sheet(id) { const s = this.book.sheets.find((x) => x.id === id); if (!s) throw new Error('Unknown sheet'); return s; }
    findTable(name) {
      for (const s of this.book.sheets) { const t = s.tables.find((x) => x.name === name); if (t) return { s, t }; }
      throw new Error('Table ' + name + ' not found');
    }
    async listWorksheets() { return this.book.sheets.map((s) => ({ id: s.id, name: s.name, visibility: 'Visible' })); }
    async listTables(sheetId) { return this.sheet(sheetId).tables.map((t) => ({ name: t.name, id: t.name })); }
    async getTableRange(name) { const { s, t } = this.findTable(name); return readBlock(s, t.c1, t.r1, t.c2, t.r2); }
    async getUsedRange(sheetId) {
      const s = this.sheet(sheetId);
      let r1 = Infinity, c1 = Infinity, r2 = 0, c2 = 0;
      for (const key of Object.keys(s.cells)) { const [r, c] = key.split(':').map(Number); r1 = Math.min(r1, r); c1 = Math.min(c1, c); r2 = Math.max(r2, r); c2 = Math.max(c2, c); }
      if (!r2) return { address: quoted(s.name) + '!A1', values: [['']], text: [['']], numberFormat: [['General']], formulas: [['']] };
      return readBlock(s, c1, r1, c2, r2);
    }
    async getRange(sheetId, address) { const a = PV.parseAddress(address); return readBlock(this.sheet(sheetId), a.c1, a.r1, a.c2, a.r2); }
    async beginSession() { return 'demo-session-' + (++this.sessions); }
    async endSession() { this.persist(); }
    /** Like typing into Excel: ISO dates become dates, plain numbers become numbers, a leading ' keeps text; null = unchanged. */
    async writeRange(sheetId, address, values) {
      const s = this.sheet(sheetId), a = PV.parseAddress(address);
      values.forEach((row, i) => row.forEach((v, j) => {
        if (v === null) return;
        let val = v, nf = 'General';
        const old = getCell(s, a.r1 + i, a.c1 + j);
        if (old) nf = old.nf;
        if (typeof v === 'string') {
          if (v.startsWith("'")) val = v.slice(1);
          else if (/^\d{4}-\d{2}-\d{2}$/.test(v)) { val = PV.dateToSerial(new Date(v + 'T00:00:00')); nf = DATE_NF; }
          else if (/^-?\d+(\.\d+)?$/.test(v.trim())) val = +v;
        }
        setCell(s, a.r1 + i, a.c1 + j, val === '' ? null : val, nf);
      }));
      return {};
    }
    async addTableRows(name, rows) {
      const { s, t } = this.findTable(name);
      for (const values of rows) {
        t.r2 += 1;
        await this.writeRange(s.id, PV.makeAddress(t.c1, t.r2, t.c2, t.r2), [values.map((v) => (v === null ? '' : v))]);
      }
      return {};
    }
    /** Excel's "insert table column": cells of the table rows from that column on shift one column right. */
    async addTableColumn(name, index, header) {
      const { s, t } = this.findTable(name);
      if (index > t.c2 - t.c1) throw new Error('Index must be ≤ the last column index (Excel cannot append with columns/add).');
      const at = t.c1 + index;
      for (let r = t.r1; r <= t.r2; r++) {
        for (let c = t.c2; c >= at; c--) {
          const cell = getCell(s, r, c);
          if (cell) s.cells[r + ':' + (c + 1)] = cell; else delete s.cells[r + ':' + (c + 1)];
        }
        delete s.cells[r + ':' + at];
      }
      t.c2 += 1;
      setCell(s, t.r1, at, header || 'Column' + (this.tableSeq++));
      return { index, name: header };
    }
    async addWorksheet(name) { const s = addSheet(this.book, name); return { id: s.id, name: s.name }; }
    async addTable(sheetName, address) {
      const s = this.book.sheets.find((x) => x.name === sheetName);
      const a = PV.parseAddress(address);
      const t = { name: 'Table' + (this.tableSeq++), c1: a.c1, r1: a.r1, c2: a.c2, r2: Math.max(a.r2, a.r1 + 1) }; // Excel adds one empty data row
      s.tables.push(t);
      return { name: t.name, id: t.name };
    }
    async renameTable(oldName, newName) { this.findTable(oldName).t.name = newName; return {}; }
  }

  PV.DemoSource = DemoSource;
  PV.buildDemoBooks = build;
})();
