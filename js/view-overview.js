/* Probe Viewer – Overview: compact stock cards first, then statistics. */
(function () {
  'use strict';
  const PV = window.PV;
  const esc = PV.esc;
  const PERIODS = [[1, 'Last month'], [2, '2 months'], [3, '3 months'], [6, '6 months'], [12, '12 months'], [24, '24 months'], ['all', 'All time']];

  function css(name) { return getComputedStyle(document.documentElement).getPropertyValue(name).trim(); }
  function baseOptions(horizontal) {
    const ink2 = css('--ink-2'), grid = css('--grid'), axis = css('--axis');
    const scale = (isValue) => ({ grid: { color: isValue ? grid : 'transparent', drawTicks: false }, border: { color: axis }, ticks: { color: ink2, padding: 6, precision: 0, font: { size: 11 }, maxRotation: 0, autoSkipPadding: 8 } });
    return {
      responsive: true, maintainAspectRatio: false, animation: false,
      indexAxis: horizontal ? 'y' : 'x',
      plugins: { legend: { display: false }, tooltip: { mode: 'index', intersect: false, padding: 10, cornerRadius: 8, backgroundColor: '#111827' } },
      interaction: { mode: 'index', intersect: false },
      scales: { x: scale(horizontal), y: Object.assign(scale(!horizontal), { beginAtZero: true }) },
    };
  }
  function chart(app, canvas, cfg) {
    if (!window.Chart) { canvas.parentElement.innerHTML = '<div class="empty">Chart library not loaded.</div>'; return; }
    app.charts.push(new window.Chart(canvas, cfg));
  }
  function bar(app, el, labels, data, horizontal, label) {
    chart(app, el, { type: 'bar', data: { labels, datasets: [{ label, data, backgroundColor: css('--series-1'), borderRadius: 4, borderSkipped: 'start', maxBarThickness: 28, categoryPercentage: 0.8, barPercentage: 0.9 }] }, options: baseOptions(horizontal) });
  }
  function dataTable(head, rows) {
    return '<details class="datatable"><summary>Show as table</summary><table class="plain"><tr>' + head.map((h, i) => '<th' + (i ? ' class="num"' : '') + '>' + esc(h) + '</th>').join('') + '</tr>' +
      rows.map((r) => '<tr>' + r.map((c, i) => '<td' + (i ? ' class="num"' : '') + '>' + esc(c) + '</td>').join('') + '</tr>').join('') + '</table></details>';
  }
  const famName = (f) => f.cfg.label || f.cfg.key;

  PV.views = PV.views || {};
  PV.views.overview = function (root, app) {
    const store = app.store, cfg = store.cfg;
    const st = app.state.overview;
    if (!st.open) st.open = {};
    if (!st.types) st.types = [];
    const counts = PV.stockCounts(store);
    const events = PV.bookingEvents(store);
    const now = new Date();
    const totalIn = counts.reduce((a, c) => a + c['in stock'], 0);
    const totalBooked = counts.reduce((a, c) => a + c.booked, 0);
    const last30 = events.filter((e) => now - e.date < 30 * 86400000).length;
    const value = PV.stockValue(store);
    const low = counts.filter((c) => c.low);

    const card = (c) => {
      const k = c.family.cfg.key, open = !!st.open[k];
      const tot = Math.max(1, c.total);
      const seg = (s, color) => (c[s] ? '<i style="width:' + (100 * c[s] / tot) + '%;background:' + color + '" title="' + s + ': ' + c[s] + '"></i>' : '');
      const fevents = events.filter((e) => e.family === k);
      const lastBooking = fevents.length ? fevents[fevents.length - 1].date : null;
      const details = open ? '<div class="details">' +
        '<div>Probes in the sheet: <b>' + c.total + '</b> · low-stock threshold: <b>' + c.threshold + '</b></div>' +
        (c.bookedWithState ? '<div>Booked and usable (state new/good): <b>' + c.usableBooked + '</b> of ' + c.bookedWithState + (c.worn ? ' · worn: <b>' + c.worn + '</b>' : '') + '</div>' : '<div>No state recorded for booked probes yet.</div>') +
        (Object.keys(c.tips).length ? '<div>In stock by tip class: ' + Object.entries(c.tips).sort().map(([t, n]) => '<span class="nowrap">class ' + esc(t) + ': <b>' + n + '</b></span>').join(' · ') + '</div>' : '') +
        (lastBooking ? '<div>Last booking: <b>' + PV.isoDate(lastBooking) + '</b></div>' : '') +
        '</div>' : '';
      return '<div class="card scard' + (c.low ? ' low' : '') + (open ? ' open' : '') + '">' +
        '<div class="name"><span>' + esc(famName(c.family)) + '</span>' + (c.low ? '<span class="lowtag" title="Fewer than ' + c.threshold + ' in stock">⚠ low</span>' : '') + '</div>' +
        '<div class="big">' + c['in stock'] + ' <small>in stock</small></div>' +
        '<div class="meter" aria-hidden="true">' + seg('in stock', css('--good')) + seg('booked', css('--booked')) + seg('sent out', css('--line-2')) + seg('missing', css('--critical')) + seg('to come', css('--warning')) + '</div>' +
        '<div class="line"><span>booked <b>' + c.booked + '</b></span><span>sent out <b>' + c['sent out'] + '</b></span>' + (c.missing ? '<span>missing <b>' + c.missing + '</b></span>' : '') + (c['to come'] ? '<span>to come <b>' + c['to come'] + '</b></span>' : '') + '</div>' +
        details + '<button class="more" data-more="' + esc(k) + '" aria-expanded="' + open + '">' + (open ? 'Less ▴' : 'More info ▾') + '</button></div>';
    };

    // ---- the two time charts and their filters ----
    const types = st.types.filter((k) => counts.some((c) => c.family.cfg.key === k));
    const tEvents = types.length ? events.filter((e) => types.includes(e.family)) : events;
    const bSeries = PV.bookingSeries(tEvents, st.period, now);
    const sot = PV.stockOverTime(store, counts, st.period, now, types);
    const typeLabel = types.length ? types.map((k) => famName(store.family(k))).join(', ') : 'all types';
    const periodLabel = PERIODS.find((p) => p[0] === st.period)[1].toLowerCase();

    const cons = PV.consumption(store, counts, events, now);
    const cost = PV.costPerProject(events);
    const persons = PV.bookingsBy(store, 'person').slice(0, 12);
    const depts = PV.bookingsBy(store, 'department');
    const lifes = PV.lifetimes(store, events);
    const hasPrice = store.families.some((f) => f.stock && f.stock.roles.price != null) || Object.keys(cfg.pricePerFamily).length > 0;
    const hasProject = store.families.some((f) => f.stock && f.stock.roles.project != null);
    const logSince = store.logRows.length ? store.logRows[0].ts : null;
    const collecting = (what, need) => '<div class="collecting">' + esc(what) + '<br><span class="small">' + esc(need) + '</span></div>';

    root.innerHTML =
      '<div class="section-title" style="margin-top:0">Stock overview</div>' +
      '<div class="grid kpis" style="margin-bottom:14px">' +
      '<div class="card kpi"><div class="v">' + totalIn + '</div><div class="l">in stock, all types</div></div>' +
      '<div class="card kpi"><div class="v">' + totalBooked + '</div><div class="l">booked right now</div></div>' +
      '<div class="card kpi"><div class="v">' + last30 + '</div><div class="l">bookings in the last 30 days</div></div>' +
      '<div class="card kpi"><div class="v">' + (value.n ? PV.ui.fmtMoney(value.value, cfg.currency) : '–') + '</div><div class="l">stock value' + (value.missing ? ' (' + value.missing + ' without price)' : '') + '</div></div>' +
      '<div class="card kpi"><div class="v" style="color:' + (low.length ? 'var(--warning-ink)' : 'var(--good-ink)') + '">' + (low.length ? '⚠ ' + low.length : '✓ 0') + '</div><div class="l">' + (low.length ? 'low stock: ' + esc(low.map((c) => famName(c.family)).join(', ')) : 'types below their low-stock threshold') + '</div></div>' +
      '</div>' +
      '<div class="grid stock-cards">' + counts.map(card).join('') + '</div>' +

      '<div class="section-title">Statistics</div>' +
      '<div class="card filters-card" style="margin-bottom:14px">' +
      '<div class="frow"><span class="lbl">Types</span><div class="pick" data-pick="types"><button data-v="" aria-pressed="' + (!types.length) + '">All</button>' +
      counts.map((c) => '<button data-v="' + esc(c.family.cfg.key) + '" aria-pressed="' + types.includes(c.family.cfg.key) + '">' + esc(famName(c.family)) + '</button>').join('') + '</div></div>' +
      '<div class="frow"><span class="lbl">Period</span><div class="pick" data-pick="period">' + PERIODS.map(([v, l]) => '<button data-v="' + v + '" aria-pressed="' + (st.period === v) + '">' + l + '</button>').join('') + '</div></div>' +
      '<div class="small faint">These filters apply to the two charts below. ' + (logSince ? 'Viewer log since ' + PV.isoDate(logSince) + '.' : 'No viewer log yet.') + '</div></div>' +
      '<div class="grid pair" style="margin-bottom:14px">' +
      '<div class="card"><div class="card-h"><h3>Bookings per ' + (bSeries.weekly ? 'week' : 'month') + '</h3><span class="small muted">' + esc(typeLabel) + ' · ' + esc(periodLabel) + '</span></div><div class="card-b">' +
      (tEvents.length ? '<div class="chart-box"><canvas id="chBookings" aria-label="Bookings over time"></canvas></div>' + dataTable([bSeries.weekly ? 'Week from' : 'Month', 'Bookings'], bSeries.labels.map((k, i) => [k, bSeries.values[i]])) : collecting('No booking dates yet for this selection.', 'Filled from "Date of booking" and from bookings made in the viewer.')) + '</div></div>' +
      '<div class="card"><div class="card-h"><h3>In stock over time</h3><span class="small muted">' + esc(typeLabel) + ' · ' + esc(periodLabel) + '</span></div><div class="card-b">' +
      (sot && sot.values.some((v) => v != null) ? '<div class="chart-box"><canvas id="chStock" aria-label="In stock over time"></canvas></div>' + dataTable([sot.weekly ? 'Week from' : 'Month', 'In stock'], sot.labels.map((k, i) => [k, sot.values[i] == null ? '–' : sot.values[i]])) : collecting('Collecting data since the viewer log starts.', 'Needs the Log table – every status change made in the viewer is recorded there.')) + '</div></div>' +
      '</div>' +

      '<div class="grid stats-grid">' +
      '<div class="card"><div class="card-h"><h3>Consumption and reorder estimate</h3></div><div class="card-b"><table class="plain"><tr><th>Type</th><th class="num">In stock</th><th class="num">Booked last 90 d</th><th class="num">Per month</th><th>Runs out ≈</th><th>Order by ≈</th></tr>' +
      cons.map((c) => '<tr><td>' + esc(famName(c.family)) + '</td><td class="num">' + c.inStock + '</td><td class="num">' + c.recent + '</td><td class="num">' + PV.ui.fmtNum(c.perMonth) + '</td><td class="date">' + (c.runout ? PV.isoDate(c.runout) : (c.events ? '<span class="faint">no recent use</span>' : '<span class="faint">collecting data</span>')) + '</td><td class="date">' + (c.orderBy ? (c.orderBy < now ? '<b style="color:var(--critical-ink)">now</b>' : PV.isoDate(c.orderBy)) : '') + '</td></tr>').join('') +
      '</table><div class="small muted" style="margin-top:8px">Rate = bookings in the last 90 days ÷ 3. Order-by date subtracts a ' + cfg.reorderLeadTimeWeeks + '-week lead time (config).</div></div></div>' +

      '<div class="card"><div class="card-h"><h3>Cost per project</h3></div><div class="card-b">' +
      (cost.rows.length ? '<div class="chart-box"><canvas id="chCost" aria-label="Cost per project"></canvas></div>' + dataTable(['Project', 'Cost (' + cfg.currency + ')', 'Probes'], cost.rows.map((r) => [r.project, PV.ui.fmtNum(r.cost, 0), r.n])) +
        '<div class="small muted">' + (cost.noProject ? cost.noProject + ' booking(s) without project. ' : '') + (cost.unpriced ? cost.unpriced + ' without price.' : '') + '</div>'
        : collecting(hasProject ? 'Collecting data – no bookings with a project yet.' : 'Needs the Project column (Data checks → Add missing columns).', (hasPrice ? '' : 'Also needs prices (Price column or pricePerFamily in config). ') + 'Each booking in the viewer asks for a project.')) + '</div></div>' +

      '<div class="card"><div class="card-h"><h3>Booked / sent out per person or customer</h3></div><div class="card-b">' +
      (persons.length ? '<div class="chart-box" style="height:' + Math.max(160, persons.length * 22 + 40) + 'px"><canvas id="chPerson"></canvas></div>' + dataTable(['Booked for', 'Probes'], persons.map((x) => [x.k, x.n])) + '<div class="small muted">Top 12 values of "Booked for" (free text, shown as written).</div>' : collecting('No bookings found.', '')) + '</div></div>' +

      '<div class="card"><div class="card-h"><h3>Probes in use per department</h3></div><div class="card-b">' +
      (depts.length ? '<div class="chart-box" style="height:' + Math.max(140, depts.length * 26 + 40) + 'px"><canvas id="chDept"></canvas></div>' + dataTable(['Department', 'Probes'], depts.map((x) => [x.k, x.n])) + '<div class="small muted">From the Department column; spelling variants (Dev / Devel / Development) are grouped for display only.</div>' : collecting('No departments filled in yet.', 'Booking in the viewer fills Department from your profile.')) + '</div></div>' +

      '<div class="card"><div class="card-h"><h3>Average probe lifetime</h3></div><div class="card-b">' +
      (lifes.length ? '<table class="plain"><tr><th>Type</th><th class="num">Probes</th><th class="num">Median days</th><th class="num">Mean days</th></tr>' + lifes.map((l) => '<tr><td>' + esc(famName(store.family(l.family))) + '</td><td class="num">' + l.n + '</td><td class="num">' + PV.ui.fmtNum(l.median, 0) + '</td><td class="num">' + PV.ui.fmtNum(l.mean, 0) + '</td></tr>').join('') + '</table><div class="small muted" style="margin-top:8px">Days from booking to the first logged dead / bad / blunt state.</div>'
        : collecting('Collecting data' + (logSince ? ' since ' + PV.isoDate(logSince) : '') + '.', 'Needs the Log: state changes to dead/bad made in the viewer are timed against the booking date.')) + '</div></div>' +
      '</div>';

    root.querySelectorAll('[data-more]').forEach((b) => (b.onclick = () => { st.open[b.dataset.more] = !st.open[b.dataset.more]; app.render(); }));
    root.querySelectorAll('[data-pick="types"] button').forEach((b) => (b.onclick = () => {
      const v = b.dataset.v;
      if (!v) st.types = [];
      else st.types = st.types.includes(v) ? st.types.filter((x) => x !== v) : st.types.concat([v]);
      app.render();
    }));
    root.querySelectorAll('[data-pick="period"] button').forEach((b) => (b.onclick = () => { st.period = b.dataset.v === 'all' ? 'all' : +b.dataset.v; app.render(); }));

    const q = (id) => root.querySelector('#' + id);
    if (q('chBookings')) bar(app, q('chBookings'), bSeries.labels, bSeries.values, false, 'Bookings');
    if (q('chStock')) {
      chart(app, q('chStock'), { type: 'line', data: { labels: sot.labels, datasets: [{ label: 'In stock', data: sot.values, borderColor: css('--series-1'), backgroundColor: css('--series-1'), borderWidth: 2, pointRadius: sot.labels.length > 30 ? 0 : 3, pointHoverRadius: 5, spanGaps: false, tension: 0 }] }, options: baseOptions(false) });
    }
    if (q('chCost')) {
      const o = baseOptions(true);
      o.plugins.tooltip.callbacks = { label: (ctx) => PV.ui.fmtMoney(ctx.parsed.x, cfg.currency) + ' · ' + cost.rows[ctx.dataIndex].n + ' probes' };
      chart(app, q('chCost'), { type: 'bar', data: { labels: cost.rows.map((r) => r.project), datasets: [{ label: 'Cost', data: cost.rows.map((r) => Math.round(r.cost)), backgroundColor: css('--series-1'), borderRadius: 4, borderSkipped: 'start', maxBarThickness: 24 }] }, options: o });
    }
    if (q('chPerson')) bar(app, q('chPerson'), persons.map((x) => x.k), persons.map((x) => x.n), true, 'Probes');
    if (q('chDept')) bar(app, q('chDept'), depts.map((x) => x.k), depts.map((x) => x.n), true, 'Probes');
  };
})();
