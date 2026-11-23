/* Probe Viewer – statistics for the Overview. Pure functions over the store. */
(function () {
  'use strict';
  const PV = window.PV;
  const DAY = 86400000;

  function priceOf(store, p) {
    const c = PV.pget(p, 'price');
    const n = c ? PV.toNum(c.v) : null;
    if (n !== null) return n;
    const fb = store.cfg.pricePerFamily[p.family.key];
    return typeof fb === 'number' ? fb : null;
  }
  function usable(store, p) {
    const fields = PV.stateFields(p).filter((f) => ['mech', 'state'].includes(f.role));
    if (!fields.length) return null;
    const v = PV.key(fields[0].value);
    if (!v) return null;
    return store.cfg.usableStates.includes(v);
  }

  /** Counts per family. Families with a Stock sheet count Stock rows only; Lifetime-only families count their Lifetime rows. */
  PV.stockCounts = function (store) {
    return store.families.filter((f) => f.stock || f.life).map((f) => {
      const src = f.stock ? f.probes.filter((p) => p.stock) : f.probes;
      const c = { family: f, total: src.length, 'in stock': 0, booked: 0, 'sent out': 0, missing: 0, 'to come': 0, other: 0, usableBooked: 0, bookedWithState: 0, tips: {} };
      for (const p of src) {
        const s = PV.statusKey(p);
        if (s in c && s !== 'total') c[s]++; else c.other++;
        if (s === 'in stock' && p.stock && p.stock.model.roles.tip != null) {
          const t = PV.clean(p.stock.row.values[p.stock.model.roles.tip]) || '?';
          c.tips[t] = (c.tips[t] || 0) + 1;
        }
        if (s === 'booked') { const u = usable(store, p); if (u !== null) { c.bookedWithState++; if (u) c.usableBooked++; } }
      }
      const thr = f.cfg.lowStock != null ? f.cfg.lowStock : store.cfg.lowStockDefault;
      c.threshold = f.stock ? thr : null;
      c.low = f.stock ? c['in stock'] < thr : false;
      return c;
    });
  };

  /** One booking event per probe: Date of booking cell if present, else the first logged Status -> booked. */
  PV.bookingEvents = function (store) {
    const logFirst = new Map();
    for (const e of store.logRows) {
      if (PV.normHeader(e.field) === 'status' && PV.key(e.new) === 'booked') {
        const k = PV.normHeader(e.sheet) + '|' + PV.clean(e.id).toUpperCase();
        if (!logFirst.has(k)) logFirst.set(k, e);
      }
    }
    const out = [];
    for (const p of store.probes) {
      if (!p.stock) continue;
      const s = PV.statusKey(p);
      const c = PV.cell(p.stock, 'bookingDate');
      let date = c ? PV.cellDate(c.v, c.nf, c.t) : null;
      const le = logFirst.get(PV.normHeader(p.stock.model.sheetName) + '|' + PV.clean(PV.logId(p)).toUpperCase());
      if (!date && le) date = le.ts;
      if (!date) continue;
      if (!['booked', 'sent out', 'missing'].includes(s) && !le) continue;
      const proj = PV.ptext(p, 'project') || (le && le.project) || '';
      out.push({ probe: p, date, family: p.family.key, project: proj, person: PV.ptext(p, 'bookedFor'), price: priceOf(store, p) });
    }
    return out.sort((a, b) => a.date - b.date);
  };

  PV.monthlySeries = function (events, months, now) {
    now = now || new Date();
    const start = PV.addMonths(new Date(Date.UTC(now.getFullYear(), now.getMonth(), 1)), -(months - 1));
    const keys = [];
    for (let i = 0; i < months; i++) keys.push(PV.monthKey(PV.addMonths(start, i)));
    const counts = Object.fromEntries(keys.map((k) => [k, 0]));
    for (const e of events) { const k = PV.monthKey(e.date); if (k in counts) counts[k]++; }
    return { keys, values: keys.map((k) => counts[k]) };
  };

  /** Consumption (last 90 days) and reorder estimate per family. */
  PV.consumption = function (store, counts, events, now) {
    now = now || new Date();
    const from = now.getTime() - 90 * DAY;
    const lead = (store.cfg.reorderLeadTimeWeeks || 0) * 7 * DAY;
    return counts.filter((c) => c.family.stock).map((c) => {
      const ev = events.filter((e) => e.family === c.family.cfg.key);
      const recent = ev.filter((e) => e.date.getTime() >= from).length;
      const perMonth = recent / 3;
      const first = ev.length ? ev[0].date : null;
      let runout = null, orderBy = null;
      if (perMonth > 0) {
        runout = new Date(now.getTime() + (c['in stock'] / perMonth) * 30.44 * DAY);
        orderBy = new Date(runout.getTime() - lead);
      }
      return { family: c.family, inStock: c['in stock'], recent, perMonth, runout, orderBy, dataSince: first, events: ev.length };
    });
  };

  /** In-stock level per month, reconstructed backwards from the Log (only as far back as the Log goes). */
  PV.stockOverTime = function (store, counts, months, now) {
    now = now || new Date();
    if (!store.logRows.length) return null;
    const logStart = store.logRows[0].ts;
    const series = [];
    const ends = [];
    const startMonth = PV.addMonths(new Date(Date.UTC(now.getFullYear(), now.getMonth(), 1)), -(months - 1));
    for (let i = 0; i < months; i++) ends.push(new Date(Math.min(PV.addMonths(startMonth, i + 1).getTime() - 1, now.getTime())));
    const sheetToFam = {};
    store.families.forEach((f) => { if (f.cfg.stockSheet) sheetToFam[PV.normHeader(f.cfg.stockSheet)] = f.cfg.key; });
    const statusEv = store.logRows.filter((e) => PV.normHeader(e.field) === 'status' && PV.normHeader(e.wb || 'stock') !== 'lifetime' && sheetToFam[PV.normHeader(e.sheet)]);
    for (const c of counts.filter((x) => x.family.stock)) {
      const fk = c.family.cfg.key;
      const evs = statusEv.filter((e) => sheetToFam[PV.normHeader(e.sheet)] === fk).sort((a, b) => b.ts - a.ts);
      let level = c['in stock'];
      let j = 0;
      const vals = new Array(months).fill(null);
      for (let i = months - 1; i >= 0; i--) {
        const t = ends[i];
        while (j < evs.length && evs[j].ts > t) {
          const e = evs[j];
          if (PV.key(e.new) === 'in stock') level--;
          if (PV.key(e.old) === 'in stock') level++;
          j++;
        }
        vals[i] = t >= logStart ? Math.max(0, level) : null;
      }
      series.push({ family: c.family, values: vals });
    }
    return { labels: ends.map((d) => PV.monthKey(d)), series, since: logStart };
  };

  /** Days from booking to the first logged dead/bad state, per family. */
  PV.lifetimes = function (store, events) {
    const dead = new Set(['dead', 'bad', 'blunt']);
    const bookedAt = new Map(events.map((e) => [e.family + '|' + PV.clean(PV.logId(e.probe)).toUpperCase(), e.date]));
    const lifeSheetToFam = {};
    store.families.forEach((f) => { if (f.cfg.lifetimeSheet) lifeSheetToFam[PV.normHeader(f.cfg.lifetimeSheet)] = f.cfg.key; });
    const per = {};
    const seen = new Set();
    for (const e of store.logRows) {
      if (!dead.has(PV.key(e.new))) continue;
      const fk = lifeSheetToFam[PV.normHeader(e.sheet)];
      if (!fk) continue;
      const id = PV.clean(e.id).toUpperCase();
      const k = fk + '|' + id;
      if (seen.has(k)) continue;
      const b = bookedAt.get(k);
      if (!b || e.ts < b) continue;
      seen.add(k);
      (per[fk] = per[fk] || []).push((e.ts - b) / DAY);
    }
    return Object.entries(per).map(([fk, arr]) => {
      arr.sort((a, b) => a - b);
      const med = arr.length % 2 ? arr[(arr.length - 1) / 2] : (arr[arr.length / 2 - 1] + arr[arr.length / 2]) / 2;
      return { family: fk, n: arr.length, median: med, mean: arr.reduce((a, b) => a + b, 0) / arr.length };
    }).sort((a, b) => a.family.localeCompare(b.family));
  };

  PV.costPerProject = function (events) {
    const m = new Map();
    let unpriced = 0, noProject = 0;
    for (const e of events) {
      if (!e.project) { noProject++; continue; }
      if (e.price == null) { unpriced++; continue; }
      const r = m.get(e.project) || { project: e.project, cost: 0, n: 0 };
      r.cost += e.price; r.n++;
      m.set(e.project, r);
    }
    return { rows: [...m.values()].sort((a, b) => b.cost - a.cost), unpriced, noProject };
  };

  PV.bookingsBy = function (store, which) {
    const m = new Map();
    for (const p of store.probes) {
      const s = PV.statusKey(p);
      if (!['booked', 'sent out', 'missing'].includes(s)) continue;
      let k;
      if (which === 'person') { if (!p.stock) continue; k = PV.clean(PV.ptext(p, 'bookedFor')); }
      else { if (!p.life) continue; k = PV.normDept(PV.ptext(p, 'department')); }
      if (!k) continue;
      m.set(k, (m.get(k) || 0) + 1);
    }
    return [...m.entries()].map(([k, n]) => ({ k, n })).sort((a, b) => b.n - a.n);
  };

  PV.stockValue = function (store) {
    let v = 0, n = 0, missing = 0;
    for (const p of store.probes) {
      if (!p.stock || PV.statusKey(p) !== 'in stock') continue;
      const pr = priceOf(store, p);
      if (pr == null) { missing++; continue; }
      v += pr; n++;
    }
    return { value: v, n, missing };
  };
  PV.priceOf = priceOf;
})();
