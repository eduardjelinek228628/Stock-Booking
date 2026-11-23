/*
 * Probe Viewer – demo data source. A tiny in-memory "Excel" with the same sheets, Tables and headers
 * as the real workbooks, filled with SYNTHETIC data (no real probes, customers or people).
 * It implements the same interface as GraphSource, so demo mode runs exactly the same read/write code.
 */
(function () {
  'use strict';
  const PV = window.PV;
  const STORE_KEY = 'probeViewer.demo.v1';
  const DATE_NF = 'd.m.yyyy';

  // ---------- mini workbook ----------
  function newBook() { return { sheets: [] }; }
  function addSheet(book, name) {
    const s = { id: '{demo-' + book.sheets.length + '}', name, cells: {}, tables: [] };
    book.sheets.push(s);
    return s;
  }
  function setCell(s, r, c, v, nf, f) {
    if (v === undefined) return;
    const k = r + ':' + c;
    if (v === null && !f) { delete s.cells[k]; return; }
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
    const q = /^[A-Za-z0-9_]+$/.test(s.name) ? s.name : "'" + s.name.replace(/'/g, "''") + "'";
    return { address: q + '!' + PV.makeAddress(c1, r1, c2, r2), values, text, numberFormat, formulas };
  }
  /** Write a header row + data rows starting at (r0,c0) and register an Excel Table over it (if tableName). */
  function writeTable(s, tableName, r0, headers, rows, opts) {
    opts = opts || {};
    headers.forEach((h, c) => setCell(s, r0, c + 1, h));
    rows.forEach((row, i) => row.forEach((cell, c) => {
      if (cell == null) return;
      if (typeof cell === 'object' && !(cell instanceof Date)) setCell(s, r0 + 1 + i, c + 1, cell.v, cell.nf, cell.f);
      else setCell(s, r0 + 1 + i, c + 1, cell);
    }));
    const r2 = r0 + rows.length + (opts.emptyTail || 0);
    if (tableName) s.tables.push({ name: tableName, c1: 1, r1: r0, c2: headers.length, r2 });
  }
  const D = (serial) => ({ v: serial, nf: DATE_NF });

  // ---------- synthetic content ----------
  const PEOPLE = ['EJ', 'MP', 'JK', 'PK', 'ON', 'KK', 'RS'];
  const GROUPS = ['RD', 'Apps', 'Development'];
  const CUSTOMERS = ['Customer Alpha', 'University Beta', 'Institute Gamma', 'Partner Delta', 'Lab Epsilon', 'Demo tour', 'Workshop'];
  const PROJECTS = ['SiC SSRM study', 'Customer demos', 'FEBID tips R&D', 'Training', 'Installations', 'Internal R&D'];
  const PRICE = { 'NPC-A': 55, 'NPC-Z': 60, 'NPC-Y': 70, 'NPC-X': 75, 'NPC-W': 80, 'NP_Pt-Sharp': 140, 'NP-Mag': 120, Akiyama: 95, PRSA: 160 };

  function build(today) {
    const R = PV.rng(20261123);
    const T = PV.dateToSerial(today);
    const stock = newBook(), life = newBook();
    const log = [];
    const lifeRows = {}; // lifetime sheet -> rows
    const logEvent = (serial, user, sheet, id, field, oldV, newV, project, wb) => {
      const d = PV.serialToDate(serial + R.next() * 0.4 + 0.33);
      log.push([PV.isoDate(d) + ' ' + String(d.getUTCHours()).padStart(2, '0') + ':' + String(d.getUTCMinutes()).padStart(2, '0'), user, wb || 'stock', sheet, id, field, oldV, newV, project || '']);
    };

    /** Common life story for one probe; returns stock cells for status/booking and pushes lifetime + log rows. */
    function story(fam, id, opts) {
      // Probes arrive in 5 batches over ~14 months; each waits a random time before it is booked.
      const frac = opts.frac || 0;
      const batch = Math.min(4, Math.floor(frac * 5));
      const offset = [...String(fam.key)].reduce((a, ch) => a + ch.charCodeAt(0), 0) % 70; // types are delivered on different days
      const arrival = T - 470 + offset + batch * 100 + R.int(0, 6);
      const res = { status: 'in stock', bookedFor: '', bookingDate: null, project: '', price: '' };
      if (arrival > T - 3) { res.status = 'to come'; return res; }
      logEvent(arrival, 'OP', fam.stockSheet, id, 'Status', 'to come', 'in stock', '', 'stock');
      const wait = Math.round(-Math.log(1 - R.next() * 0.999) * (opts.meanWait || 75)) + 2;
      const bd = arrival + wait;
      if (bd >= T) return res; // still in stock
      const outAfter = R.int(3, 45);
      const status = bd + outAfter < T && R.next() < 0.72 ? (R.next() < 0.04 ? 'missing!' : 'sent out') : 'booked';
      res.status = status;
      const person = R.next() < 0.55 ? R.pick(PEOPLE) : (R.next() < 0.5 ? R.pick(GROUPS) : R.pick(CUSTOMERS));
      const project = R.next() < 0.85 ? R.pick(PROJECTS) : '';
      res.bookedFor = person; res.bookingDate = bd; res.project = project;
      logEvent(bd, PEOPLE.includes(person) ? person : 'OP', fam.stockSheet, id, 'Status', 'in stock', 'booked', project, 'stock');
      if (status !== 'booked') logEvent(bd + outAfter, 'OP', fam.stockSheet, id, 'Status', 'booked', status, project, 'stock');
      // lifetime record for most bookings
      if (fam.lifetimeSheet && opts.life && R.next() < 0.75) {
        const endDay = bd + R.int(10, 220);
        const dead = endDay < T && R.next() < 0.45;
        const owner = PEOPLE.includes(person) ? person : (R.next() < 0.1 ? 'PB -> EJ' : 'RD');
        const dept = R.pick(['Apps', 'Apps', 'Devel', 'Dev', 'Development']);
        const mech = dead ? 'dead' : R.pick(['new', 'good', 'good', 'bad', 'unknown']);
        const ele = dead ? 'none' : R.pick(['good', 'good', 'weak ', 'none', 'unknown']);
        const who = owner.slice(-2);
        const hist = dead ? PV.isoDate(PV.serialToDate(endDay)) + ' ' + who + ': tip broken during approach' : (R.next() < 0.5 ? 'used for ' + (project || 'tests') : '');
        opts.life({ status, owner, dept, mech, ele, asOf: dead ? endDay : bd, hist, project });
        logEvent(bd, who, fam.lifetimeSheet, id, 'Owner', '', owner, project, 'lifetime');
        if (dead) logEvent(endDay, who, fam.lifetimeSheet, id, opts.stateField || 'Mech.', 'good', 'dead', project, 'lifetime');
      }
      return res;
    }

    const famBy = (k) => window.PV_CONFIG.families.find((f) => f.key === k);
    const pad = (n, w) => String(n).padStart(w, '0');
    const primaryFor = (k, tip) => (k <= 10 ? (tip === 2 ? 'C-AFM' : 'KPFM') : 'PFM, KPFM');

    // ---- NPC sheets (Stock + Lifetime) ----
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
      const sh = addSheet(stock, sp.key);
      let headers;
      if (sp.variant === 'A') headers = ['ID', 'Status', 'Length', 'freso', 'Q factor', 'Stiffness', 'Primary ', 'Secondary', sp.tipHdr, 'Booked for', 'Note'];
      else if (sp.variant === 'Y') headers = ['ID', 'Status', 'Length', 'freso', 'Q factor', 'Stiffness', 'Primary ', 'Secondary', 'Tip current', 'Tip ', 'Booked for', 'Note'];
      else if (sp.variant === 'X') headers = ['ID', 'Status', 'Length', 'freso', 'Q factor', 'Stiffness', 'Primary ', 'Secondary', 'El. properties', 'Tip ', 'Notes', 'Booked for', 'Date of booking', 'Note'];
      else headers = ['ID', ' ', 'Length', 'freso', 'Q factor', 'Stiffness', 'Primary ', 'Secondary', 'Tip current', 'Tip ', 'Booked for', 'Date of booking', 'Note'];
      headers = headers.concat(['Project', 'Price']);
      const col = (h) => headers.indexOf(h);
      const rows = [];
      if (sp.units) {
        const u = headers.map(() => null);
        u[0] = sp.variant === 'A' ? null : 'A'; u[col('Length')] = '(µm)'; u[col('freso')] = '(kHz)'; u[col('Stiffness')] = '(N/m)';
        u[col('Primary ')] = 'technique'; u[col('Secondary')] = 'technique'; u[col(sp.tipHdr)] = 'class';
        rows.push(u);
      }
      sh.overview = true;
      const lifeSheet = fam.lifetimeSheet;
      for (let i = 1; i <= sp.n; i++) {
        const id = sp.prefix + pad(i * (sp.key === 'NPC-X' ? 3 : 1) + (sp.key === 'NPC-W' ? 10 : 0), 3);
        const L = R.pick(sp.lengths);
        const fr = R.round(Math.max(8, 12000 / L - 5), 12000 / L + 25, 1);
        const k = R.round(0.6, L < 350 ? 140 : 40, 2);
        const tip = R.weighted([[1, 6], [2, 3], [3, sp.variant === 'X' ? 1 : 0]]);
        const r = headers.map(() => null);
        r[0] = id; r[col('Length')] = L; r[col('freso')] = fr; r[col('Q factor')] = R.int(120, 600); r[col('Stiffness')] = k;
        r[col('Primary ')] = sp.key === 'NPC-Z' || sp.key === 'NPC-Y' ? { v: primaryFor(k, tip), f: '=IF(AND(F' + (rows.length + 3) + '<=10,' + PV.colLetter(col(sp.tipHdr) + 1) + (rows.length + 3) + '=2),"C-AFM","KPFM")' } : primaryFor(k, tip);
        if (R.next() < 0.3) r[col('Secondary')] = R.pick(['C-AFM', 'PFM', 'KPFM']);
        r[col(sp.tipHdr)] = tip;
        if (col('Tip current') >= 0) r[col('Tip current')] = R.next() < 0.9 ? 'yes' : 'no';
        if (col('El. properties') >= 0) r[col('El. properties')] = R.pick(['Ok', 'Ok', 'Ok via sample']);
        const st = story(fam, id, {
          frac: (i - 1) / sp.n,
          life: lifeSheet ? (x) => (lifeRows[lifeSheet] = lifeRows[lifeSheet] || []).push([id, x.status, L, fr, r[col('Q factor')], k, R.round(20, 90, 1), x.dept, x.owner, x.mech, x.ele, D(x.asOf), x.hist]) : null,
        });
        r[col(headers[1])] = st.status;
        r[col('Booked for')] = st.bookedFor || null;
        if (col('Date of booking') >= 0 && st.bookingDate) r[col('Date of booking')] = D(st.bookingDate);
        if (R.next() < 0.15) r[col('Note')] = R.pick(['demo', 'training - used', 'to fair and back, unused', 'low q-factor', 'Box 1', 'Box 2']);
        r[col('Project')] = st.project || null;
        r[col('Price')] = PRICE[sp.key] + R.pick([0, 0, 5, -5]);
        rows.push(r);
      }
      // realistic quirks
      if (sp.key === 'NPC-X') { rows[5][col('freso')] = String(rows[5][col('freso')]); rows[9][col('Q factor')] = ' -'; }
      if (sp.key === 'NPC-W') {
        rows[7][1] = rows[7][1] === 'sent out' ? 'Sent out' : rows[7][1];
        const dup = rows[20].slice(); dup[col('Note')] = 'duplicate row'; rows.push(dup);
        rows[30][1] = 'missing';
      }
      writeTable(sh, sp.table, 2, headers, rows, { emptyTail: sp.key === 'NPC-Y' ? 6 : 2 });
      setCell(sh, 1, 1, 'Probe'); setCell(sh, 1, 4, 'Datasheet'); setCell(sh, 1, 7, 'Selection info'); setCell(sh, 1, col('Booked for') + 1, 'Booking state');
    }

    // ---- NP-Mag ----
    {
      const fam = famBy('NP-Mag');
      const sh = addSheet(stock, 'NP-Mag');
      const headers = ['ID', 'Status', 'Type', ' freso (Hz)      ', 'Q factor', 'Original Akiyama box    ', 'Tip', 'Booked for', 'Note', 'Project', 'Price'];
      const rows = [];
      for (let i = 1; i <= 80; i++) {
        const id = i <= 6 ? 12000 + i * 3 : 'MAG-' + pad(i, 4);
        const fr = R.int(39000, 43000), q = R.next() < 0.85 ? R.int(1100, 1600) : null;
        const box = 'BOX-M' + pad(Math.ceil(i / 10), 2) + '-' + pad(((i - 1) % 10) + 1, 2);
        const tip = R.weighted([[1, 5], [2, 1], ['?', 3]]);
        const st = story(fam, String(id), {
          frac: (i - 1) / 80, stateField: 'Mech.',
          life: (x) => (lifeRows[fam.lifetimeSheet] = lifeRows[fam.lifetimeSheet] || []).push([id, x.status, 'Zaragoza', fr, q, box, tip, x.dept, x.owner, x.mech, x.mech === 'dead' ? 'none' : R.pick(['good', 'weak ', 'unknown']), D(x.asOf), x.hist]),
        });
        rows.push([id, st.status, R.next() < 0.97 ? 'Zaragoza' : 'NanoSens.', fr, q, box, tip, st.bookedFor || null, R.next() < 0.1 ? 'weak MFM' : null, st.project || null, PRICE['NP-Mag']]);
      }
      writeTable(sh, 'NP_Mag', 2, headers, rows);
    }

    // ---- Akiyama (box-level: 10-pack rows + single probes with position) ----
    {
      const fam = famBy('Akiyama');
      const sh = addSheet(stock, 'Akiyama');
      const headers = ['Box ID', 'Status  ', 'Type', 'position', 'Booked for', 'Note', 'Sloupec1', 'Project', 'Price'];
      const rows = [];
      for (let b = 1; b <= 16; b++) {
        const box = 'AK-' + pad(70000 + b * 37, 5) + '-' + pad(b, 2);
        const kind = R.weighted([['10-pack', 5], ['probe', 6], ['buffer', 1]]);
        const pushLife = (pos) => (x) => (lifeRows.Akiyama = lifeRows.Akiyama || []).push([box, x.status, kind, pos, x.dept, x.owner, x.mech === 'bad' ? 'blunt' : x.mech, D(x.asOf), x.hist]);
        if (kind === 'probe') {
          for (let pos = 1; pos <= 10; pos++) {
            const st = story(fam, box + '#' + pos, { frac: (b - 1) / 16, meanWait: 50, stateField: 'State', life: pushLife(pos) });
            rows.push([box, st.status, 'probe', pos, st.bookedFor || null, null, null, st.project || null, PRICE.Akiyama]);
          }
        } else {
          const st = story(fam, box, { frac: (b - 1) / 16, meanWait: 50, stateField: 'State', life: pushLife(null) });
          rows.push([box, st.status, kind, { v: kind === 'buffer' ? '-' : ' ', f: '=IF(C' + (rows.length + 3) + ' = "buffer","-"," ")' }, st.bookedFor || null, null, null, st.project || null, kind === '10-pack' ? PRICE.Akiyama * 10 : PRICE.Akiyama]);
        }
      }
      writeTable(sh, 'Akiyama', 2, headers, rows);
    }

    // ---- PRSA ----
    {
      const fam = famBy('PRSA');
      const sh = addSheet(stock, 'PRSA');
      const headers = ['ID', 'Status', 'Type', ' freso (kHz)', 'Original box ID', 'Booked for', 'Note', 'Project', 'Price'];
      const types = ['PRSA-L300-F50-Si-PCB', 'PRSA-L100-F500-Si-PCB', 'PRS-L70-F900-Si-PCB', 'PRSA-L400-F30-SCD-PCB'];
      const rows = [];
      for (let i = 1; i <= 26; i++) {
        const id = 'PRS-' + pad(i, 4);
        const t = R.pick(types);
        const fr = R.next() < 0.6 ? R.round(30, 65, 2) : R.pick(['30 - 65', '15 - 40', '200-800']);
        const box = '2025-0' + (1 + (i % 6)) + '-12 Box ' + (1 + Math.floor(i / 9));
        const st = story(fam, id, { frac: (i - 1) / 26, meanWait: 120, stateField: 'State', life: (x) => (lifeRows.PRSA = lifeRows.PRSA || []).push([id, x.status, t, fr, box, x.dept, x.owner, x.mech, D(x.asOf), x.hist]) });
        rows.push([id, st.status, t, fr, box, st.bookedFor || null, null, st.project || null, PRICE.PRSA]);
      }
      writeTable(sh, 'PRSA', 2, headers, rows);
    }

    // ---- helper sheets ----
    {
      const sem = addSheet(stock, 'Semafor');
      writeTable(sem, 'Tabulka10', 3, ['Probe', 'Status', 'Note'], [['Akiyama', 'All pass', null], ['PRSA', 'All pass', null], ['NPC', 'All tested', 'SEM imaging, frequency sweep, conductivity test'], ['NPM', 'All pass', null]]);
      setCell(sem, 9, 2, 'Probe status can only be changed by agreement between applications, development and operations');
      const data = addSheet(stock, 'data');
      const cols = [['status', 'to come', 'in stock', 'booked', 'sent out', 'missing!'], ['methods', 'YES', 'NO'], ['projects'].concat(PROJECTS)];
      cols.forEach((list, c) => list.forEach((v, r) => setCell(data, r + 1, c + 1, v)));
      const lg = addSheet(stock, 'Log');
      log.sort((a, b) => (a[0] < b[0] ? -1 : 1));
      writeTable(lg, 'ViewerLog', 1, window.PV_CONFIG.logColumns, log);
    }

    // ---- Probe Lifetime ----
    const lifeNpcHdr = ['ID', 'Status', 'Length', 'freso', 'Q factor', 'Stiffness', 'Sensitivity', 'Department', 'Owner', 'Mech.', 'Ele.', 'As of:', 'History'];
    const lifeUnits = ['A', null, '(µm)', '(kHz)', null, '(N/m)', '(mV/nm)', null, null, null, null, null, null];
    const lifeTables = { 'NPC-A': 'NP_Con_Si', 'NPC-Z': 'NP_Con_PolySi', 'NPC-Y': 'NP_Con_PolySi2', 'NPC-X': 'NP_Con_PolySi27', 'NPC-W': 'NP_Con_PolySi2713' };
    for (const k of Object.keys(lifeTables)) {
      const sh = addSheet(life, k);
      const hdr = lifeNpcHdr.slice(); if (k === 'NPC-A') hdr[11] = 'Last update:';
      writeTable(sh, lifeTables[k], 2, hdr, [lifeUnits].concat(lifeRows[k] || []), { emptyTail: 3 });
      setCell(sh, 1, 8, k === 'NPC-A' ? 'Booking info' : 'Boing info');
    }
    {
      const mk = (name, table, hdr, rows) => writeTable(addSheet(life, name), table, 2, hdr, rows);
      mk('FEBID-NPC', 'NP_Con_PolySi278', lifeNpcHdr, [lifeUnits].concat(Array.from({ length: 12 }, (_, i) => ['F' + pad(i + 1, 3), R.pick(['in stock', 'in stock', 'booked']), 420, R.round(60, 140, 1), R.int(150, 500), R.round(10, 60, 2), null, 'Apps', i < 3 ? 'EJ' : 'RD', R.pick(['new', 'good', 'unknown']), R.pick(['good', 'weak ', 'unknown']), D(T - R.int(5, 200)), null])));
      mk('NPM ("Maki")', 'NP_Mag', ['ID', 'Status', 'Type', 'freso', 'Q factor', 'Original Akiyama box    ', 'Tip  ', 'Department     ', 'Owner', 'Mech.', 'Mag.', 'As of:', 'History'], lifeRows['NPM ("Maki")'] || []);
      mk('PRSA', 'PRSA', ['ID', 'Status', 'Type', ' freso (kHz)', 'Original box ID', 'Department', 'Owner', 'State', 'As of:', 'Note'], lifeRows.PRSA || []);
      mk('Akiyama', 'Akiyama', ['Box ID', 'Status  ', 'Type', 'Number', 'Department', 'Owner', 'State', 'As of:', 'History'], lifeRows.Akiyama || []);
      mk('NPC-BDD', 'NP_Con_PolySi2710', ['ID', 'Status', 'Length', 'freso', 'Q factor', 'Leakage res.', 'Contact res.', 'Department', 'Owner', 'Mech.', 'Ele.', 'As of:', 'History'],
        [['A', null, '(µm)', '(kHz)', null, '(Ohm)', '(Ohm)', null, null, null, null, null, null]].concat(Array.from({ length: 10 }, (_, i) => ['B' + pad(i + 1, 3), 'in stock', 420, R.round(40, 90, 1), R.int(150, 400), R.int(1e6, 9e6), R.int(2e3, 9e4), null, i < 4 ? 'EJ' : 'ON', R.pick(['new', 'unknown', 'dead']), R.pick(['none', 'weak ', 'good']), D(T - R.int(5, 120)), null])));
      mk('Trilayer', 'NP_Con_PolySi271012', ['ID', 'Status', 'Length', 'Tip material', 'freso', 'Q factor', 'Leakage res.', 'Contact res.', 'Department', 'Owner', 'Mech.', 'Ele.', 'As of:', 'History'],
        Array.from({ length: 14 }, (_, i) => ['T' + pad(i + 1, 3), R.pick(['in stock', 'sent out', 'booked']), 310, R.pick(['Pt', 'W']), R.round(80, 150, 1), R.int(200, 500), R.int(1e6, 9e6), R.int(2e3, 9e4), R.pick(['Development', 'Development ', 'Application ']), R.pick(['ON', 'JK', 'RS']), R.pick(['new', 'dead', 'good']), R.pick(['new', 'none', 'good']), D(T - R.int(5, 300)), null]));
      mk('NP-200', 'NP_Con_PolySi271014', ['ID', 'Status', 'Length', 'freso', 'Q factor', 'Stiffness ', 'Contact res.', 'Department', 'Owner', 'Mech.', 'Ele.', 'As of:', 'History'],
        Array.from({ length: 6 }, (_, i) => ['N' + pad(i + 1, 3), null, 200, R.round(250, 400, 1), R.int(300, 700), R.round(40, 120, 2), null, 'Dev', 'ON', R.pick(['new', 'dead']), 'none', D(T - R.int(5, 300)), null]));
      // Adama and CDT-FMR are plain ranges (no Table) in the real file – kept that way here.
      const ad = addSheet(life, 'Adama');
      writeTable(ad, null, 2, ['Box ID', 'Status  ', 'Type', 'Number', 'k [N/m]', 'fr [kHz]', 'Department', 'Owner', 'Tip state', 'Current', 'As of:', 'History'],
        Array.from({ length: 8 }, (_, i) => ['AD' + pad(100 + i, 5), R.pick(['in stock', 'sent out', 'missing!']), R.pick(['Apex', 'Cone']), 1, R.int(20, 60), R.int(200, 350), 'Apps', i < 2 ? 'EJ' : 'ON', R.pick(['new', 'sharp', 'blunt', 'dead']), R.pick(['new', 'good', 'weak ']), D(T - R.int(10, 500)), i === 0 ? 'used for SiC MOSFET cross-section' : null]));
      setCell(ad, 1, 1, 'Probe'); setCell(ad, 1, 7, 'Booking info');
      const cd = addSheet(life, 'CDT-FMR');
      writeTable(cd, null, 2, ['Box ID', 'Status  ', 'Type', 'Number', 'k [N/m]', 'fr [kHz]', 'Department', 'Owner', 'State', 'As of:', 'History'],
        Array.from({ length: 6 }, (_, i) => ['CD' + pad(4000 + i, 6), R.pick(['in stock', 'sent out']), 'probe', 1, null, null, 'Dev', i < 2 ? 'EJ' : 'ON', R.pick(['new', 'good', 'dead']), D(T - R.int(10, 500)), null]));
      setCell(cd, 1, 1, 'Probe');
      const sem = addSheet(life, 'Semafor');
      writeTable(sem, 'Tabulka10', 3, ['Probe', 'Status', 'Note'], [['Akiyama', 'All pass', null]]);
      const data = addSheet(life, 'data');
      const lists = [
        ['status', 'to come', 'in stock', 'booked', 'sent out', 'missing!'],
        ['state mech.', 'new', 'good', 'bad', 'dead', 'unknown', null, 'state el/mag', 'good', 'weak ', 'none', 'unknown', null, 'state akiyama', 'new', 'good', 'blunt', 'spikes', 'dead', 'unknown'],
        ['department', 'Apps', 'Devel'],
      ];
      lists.forEach((list, c) => list.forEach((v, r) => { if (v != null) setCell(data, r + 1, c + 1, v); }));
      ['Adama type', 'Apex', 'Cone', null, 'Adama current', 'new', 'good', 'weak ', 'none', 'unknown', null, 'Adama tip state', 'new', 'sharp', 'blunt', 'dead', 'unknown'].forEach((v, r) => { if (v != null) setCell(data, r + 1, 10, v); });
    }
    return { stock, lifetime: life, builtFor: PV.isoDate(today) };
  }

  // ---------- source ----------
  class DemoSource {
    constructor() {
      this.kind = 'demo';
      const saved = PV.storage.get(STORE_KEY);
      this.books = saved && saved.stock && saved.lifetime ? saved : build(new Date());
      this.sessions = 0;
    }
    reset() { PV.storage.del(STORE_KEY); this.books = build(new Date()); }
    persist() { PV.storage.set(STORE_KEY, this.books); }
    book(k) { const b = this.books[k]; if (!b) throw new Error('Unknown workbook ' + k); return b; }
    sheet(k, id) { const s = this.book(k).sheets.find((x) => x.id === id); if (!s) throw new Error('Unknown sheet'); return s; }
    async listWorksheets(k) { return this.book(k).sheets.map((s) => ({ id: s.id, name: s.name, visibility: 'Visible' })); }
    async listTables(k, sheetId) { return this.sheet(k, sheetId).tables.map((t) => ({ name: t.name, id: t.name })); }
    findTable(k, name) {
      for (const s of this.book(k).sheets) { const t = s.tables.find((x) => x.name === name); if (t) return { s, t }; }
      throw new Error('Table ' + name + ' not found');
    }
    async getTableRange(k, name) { const { s, t } = this.findTable(k, name); return readBlock(s, t.c1, t.r1, t.c2, t.r2); }
    async getUsedRange(k, sheetId) {
      const s = this.sheet(k, sheetId);
      let r1 = Infinity, c1 = Infinity, r2 = 0, c2 = 0;
      for (const key of Object.keys(s.cells)) { const [r, c] = key.split(':').map(Number); r1 = Math.min(r1, r); c1 = Math.min(c1, c); r2 = Math.max(r2, r); c2 = Math.max(c2, c); }
      if (!r2) return { address: s.name + '!A1', values: [['']], text: [['']], numberFormat: [['General']], formulas: [['']] };
      return readBlock(s, c1, r1, c2, r2);
    }
    async getRange(k, sheetId, address) { const a = PV.parseAddress(address); return readBlock(this.sheet(k, sheetId), a.c1, a.r1, a.c2, a.r2); }
    async beginSession() { return 'demo-session-' + (++this.sessions); }
    async endSession() { this.persist(); }
    async writeRange(k, sheetId, address, values) {
      const s = this.sheet(k, sheetId), a = PV.parseAddress(address);
      values.forEach((row, i) => row.forEach((v, j) => {
        if (v === null) return; // null = leave unchanged (same as Graph)
        let val = v, nf = 'General';
        const old = getCell(s, a.r1 + i, a.c1 + j);
        if (old) nf = old.nf;
        if (typeof v === 'string') {
          if (v.startsWith("'")) val = v.slice(1);
          else if (/^\d{4}-\d{2}-\d{2}$/.test(v)) { val = PV.dateToSerial(new Date(v + 'T00:00:00')); nf = DATE_NF; }
          else if (PV.toNum(v) !== null && /^-?\d+(\.\d+)?$/.test(v.trim())) val = +v;
        }
        setCell(s, a.r1 + i, a.c1 + j, val === '' ? null : val, nf);
      }));
      return {};
    }
    async addTableRow(k, name, values) {
      const { s, t } = this.findTable(k, name);
      t.r2 += 1;
      await this.writeRange(k, s.id, PV.makeAddress(t.c1, t.r2, t.c2, t.r2), [values.map((v) => (v === null ? '' : v))]);
      return { index: t.r2 - t.r1 - 1 };
    }
  }

  PV.DemoSource = DemoSource;
  PV.buildDemoBooks = build;
})();
