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
    const f = PV.stateFields(p).find((x) => x.role === 'mech' || x.role === 'state');
    if (!f) return null;
    const v = PV.key(f.value);
    if (!v) return null;
    return store.cfg.usableStates.includes(v);
  }
  function sheetToFam(store) {
    const m = {};
    store.families.forEach((f) => { m[PV.normHeader(f.cfg.sheet)] = f.cfg.key; });
    return m;
  }

  /** Counts per probe type. */
  PV.stockCounts = function (store) {
    return store.families.filter((f) => f.stock).map((f) => {
      const c = { family: f, total: f.probes.length, 'in stock': 0, booked: 0, 'sent out': 0, missing: 0, 'to come': 0, other: 0, usableBooked: 0, bookedWithState: 0, tips: {}, worn: 0 };
      for (const p of f.probes) {
        const s = PV.statusKey(p);
        if (s in c && s !== 'total') c[s]++; else c.other++;
        if (s === 'in stock' && f.stock.roles.tip != null) {
          const t = PV.clean(p.stock.row.values[f.stock.roles.tip]) || '?';
          c.tips[t] = (c.tips[t] || 0) + 1;
        }
        if (s === 'booked') { const u = usable(store, p); if (u !== null) { c.bookedWithState++; if (u) c.usableBooked++; else c.worn++; } }
      }
      c.threshold = f.cfg.lowStock != null ? f.cfg.lowStock : store.cfg.lowStockDefault;
      c.low = c['in stock'] < c.threshold;
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

  /**
   * Time bins for a period: months = 1, 2, 3, 6, 12, 24 or 'all'.
   * Up to 3 months the bins are weeks (Monday–Sunday), otherwise calendar months.
   * 'all' starts at the first event (or the first log entry).
   */
  PV.bins = function (period, now, firstDate) {
    now = now || new Date();
    const end = now.getTime();
    let start;
    if (period === 'all') start = firstDate ? firstDate.getTime() : end - 365 * DAY;
    else start = new Date(now.getFullYear(), now.getMonth() - period, now.getDate()).getTime();
    const weekly = period !== 'all' && period <= 3;
    const bins = [];
    if (weekly) {
      const d0 = new Date(start); const dow = (d0.getUTCDay() + 6) % 7;
      let t = Date.UTC(d0.getUTCFullYear(), d0.getUTCMonth(), d0.getUTCDate() - dow);
      while (t <= end) { const e = t + 7 * DAY; bins.push({ from: t, to: e, label: PV.isoDate(new Date(t)).slice(5) }); t = e; }
    } else {
      const d0 = new Date(start);
      let m = Date.UTC(d0.getUTCFullYear(), d0.getUTCMonth(), 1);
      while (m <= end) { const n = PV.addMonths(new Date(m), 1).getTime(); bins.push({ from: m, to: n, label: PV.monthKey(new Date(m)) }); m = n; }
    }
    return { bins, weekly, start };
  };

  PV.bookingSeries = function (events, period, now) {
    const b = PV.bins(period, now, events.length ? events[0].date : null);
    const values = b.bins.map((x) => events.filter((e) => e.date.getTime() >= x.from && e.date.getTime() < x.to).length);
    return { labels: b.bins.map((x) => x.label), values, weekly: b.weekly };
  };
  // kept for compatibility (monthly bins over the last n months)
  PV.monthlySeries = function (events, months, now) { const s = PV.bookingSeries(events, months, now); return { keys: s.labels, values: s.values }; };

  /** Consumption (last 90 days) and reorder estimate per family. */
  PV.consumption = function (store, counts, events, now) {
    now = now || new Date();
    const from = now.getTime() - 90 * DAY;
    const lead = (store.cfg.reorderLeadTimeWeeks || 0) * 7 * DAY;
    return counts.map((c) => {
      const ev = events.filter((e) => e.family === c.family.cfg.key);
      const recent = ev.filter((e) => e.date.getTime() >= from).length;
      const perMonth = recent / 3;
      let runout = null, orderBy = null;
      if (perMonth > 0) {
        runout = new Date(now.getTime() + (c['in stock'] / perMonth) * 30.44 * DAY);
        orderBy = new Date(runout.getTime() - lead);
      }
      return { family: c.family, inStock: c['in stock'], recent, perMonth, runout, orderBy, dataSince: ev.length ? ev[0].date : null, events: ev.length };
    });
  };

  /** In-stock level at the end of each bin, reconstructed backwards from the Log (only as far back as the Log goes). */
  PV.stockOverTime = function (store, counts, period, now, famKeys) {
    now = now || new Date();
    if (!store.logRows.length) return null;
    const logStart = store.logRows[0].ts;
    const b = PV.bins(period, now, logStart);
    const ends = b.bins.map((x) => new Date(Math.min(x.to - 1, now.getTime())));
    const s2f = sheetToFam(store);
    const statusEv = store.logRows.filter((e) => PV.normHeader(e.field) === 'status' && s2f[PV.normHeader(e.sheet)]);
    const keys = famKeys && famKeys.length ? famKeys : counts.map((c) => c.family.cfg.key);
    const total = ends.map(() => null);
    for (const c of counts.filter((x) => keys.includes(x.family.cfg.key))) {
      const fk = c.family.cfg.key;
      const evs = statusEv.filter((e) => s2f[PV.normHeader(e.sheet)] === fk).sort((a, z) => z.ts - a.ts);
      let level = c['in stock'], j = 0;
      for (let i = ends.length - 1; i >= 0; i--) {
        while (j < evs.length && evs[j].ts > ends[i]) {
          if (PV.key(evs[j].new) === 'in stock') level--;
          if (PV.key(evs[j].old) === 'in stock') level++;
          j++;
        }
        if (ends[i] >= logStart) total[i] = (total[i] || 0) + Math.max(0, level);
      }
    }
    return { labels: b.bins.map((x) => x.label), values: total, since: logStart, weekly: b.weekly };
  };

  /** Days from booking to the first logged dead/bad/blunt state, per family. */
  PV.lifetimes = function (store, events) {
    const dead = new Set(['dead', 'bad', 'blunt']);
    const bookedAt = new Map(events.map((e) => [e.family + '|' + PV.clean(PV.logId(e.probe)).toUpperCase(), e.date]));
    const s2f = sheetToFam(store);
    const per = {}, seen = new Set();
    for (const e of store.logRows) {
      if (!dead.has(PV.key(e.new))) continue;
      const fk = s2f[PV.normHeader(e.sheet)];
      if (!fk) continue;
      const k = fk + '|' + PV.clean(e.id).toUpperCase();
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
      if (!['booked', 'sent out', 'missing'].includes(PV.statusKey(p))) continue;
      const k = which === 'person' ? PV.clean(PV.ptext(p, 'bookedFor')) : PV.normDept(PV.ptext(p, 'department'));
      if (!k) continue;
      m.set(k, (m.get(k) || 0) + 1);
    }
    return [...m.entries()].map(([k, n]) => ({ k, n })).sort((a, b) => b.n - a.n);
  };

  PV.stockValue = function (store) {
    let v = 0, n = 0, missing = 0;
    for (const p of store.probes) {
      if (PV.statusKey(p) !== 'in stock') continue;
      const pr = priceOf(store, p);
      if (pr == null) { missing++; continue; }
      v += pr; n++;
    }
    return { value: v, n, missing };
  };
  PV.priceOf = priceOf;
})();
