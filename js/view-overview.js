/* Probe Viewer – Overview: stock first, then statistics. */
(function () {
  'use strict';
  const PV = window.PV;
  const esc = PV.esc;

  function css(name) { return getComputedStyle(document.documentElement).getPropertyValue(name).trim(); }
  function baseOptions(horizontal) {
    const ink2 = css('--ink-2'), grid = css('--grid'), axis = css('--axis');
    const scale = (isValue) => ({ grid: { color: isValue ? grid : 'transparent', drawTicks: false }, border: { color: axis }, ticks: { color: ink2, padding: 6, precision: 0, font: { size: 11 } } });
    return {
      responsive: true, maintainAspectRatio: false, animation: false,
      indexAxis: horizontal ? 'y' : 'x',
      plugins: { legend: { display: false }, tooltip: { mode: 'index', intersect: false, padding: 10, cornerRadius: 6 } },
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

  PV.views = PV.views || {};
  PV.views.overview = function (root, app) {
    const store = app.store, cfg = store.cfg;
    const st = app.state.overview;
    const counts = PV.stockCounts(store);
    const events = PV.bookingEvents(store);
    const now = new Date();
    const stockFams = counts.filter((c) => c.family.stock);
    const lifeFams = counts.filter((c) => !c.family.stock);
    const totalIn = stockFams.reduce((a, c) => a + c['in stock'], 0);
    const totalBooked = stockFams.reduce((a, c) => a + c.booked, 0);
    const last30 = events.filter((e) => now - e.date < 30 * 86400000).length;
    const value = PV.stockValue(store);
    const low = stockFams.filter((c) => c.low);

    const card = (c) => {
      const tot = Math.max(1, c.total);
      const seg = (k, color) => (c[k] ? '<i style="width:' + (100 * c[k] / tot) + '%;background:' + color + '" title="' + k + ': ' + c[k] + '"></i>' : '');
      const tips = Object.keys(c.tips).length > 1 ? '<div class="line"><span>In stock by tip class: ' + Object.entries(c.tips).sort().map(([t, n]) => '<span class="nowrap">class ' + esc(t) + ': <b>' + n + '</b></span>').join(' · ') + '</span></div>' : '';
      return '<div class="card scard' + (c.low ? ' low' : '') + '" data-fam="' + esc(c.family.cfg.key) + '" style="cursor:pointer">' +
        '<div class="name">' + esc(c.family.cfg.label || c.family.cfg.key) + (c.low ? '<span class="lowtag">⚠ Low stock (&lt; ' + c.threshold + ')</span>' : '') + '</div>' +
        '<div class="big">' + c['in stock'] + ' <small>in stock</small></div>' +
        '<div class="meter" aria-hidden="true">' + seg('in stock', css('--good')) + seg('booked', css('--booked')) + seg('sent out', css('--line-2')) + seg('missing', css('--critical')) + seg('to come', css('--warning')) + '</div>' +
        '<div class="line"><span>booked <b>' + c.booked + '</b></span><span>sent out <b>' + c['sent out'] + '</b></span>' + (c.missing ? '<span>missing <b>' + c.missing + '</b></span>' : '') + (c['to come'] ? '<span>to come <b>' + c['to come'] + '</b></span>' : '') + '</div>' +
        (c.bookedWithState ? '<div class="line"><span>usable among booked (state new/good): <b>' + c.usableBooked + '</b> of ' + c.bookedWithState + '</span></div>' : '') + tips + '</div>';
    };

    const famOptions = '<option value="">All probe types</option>' + stockFams.map((c) => '<option value="' + esc(c.family.cfg.key) + '"' + (st.family === c.family.cfg.key ? ' selected' : '') + '>' + esc(c.family.cfg.label || c.family.cfg.key) + '</option>').join('');
    const fEvents = st.family ? events.filter((e) => e.family === st.family) : events;
    const monthly = PV.monthlySeries(fEvents, st.months, now);
    const cons = PV.consumption(store, counts, events, now);
    const sot = PV.stockOverTime(store, counts, st.months, now);
    const cost = PV.costPerProject(st.family ? fEvents : events);
    const persons = PV.bookingsBy(store, 'person').slice(0, 12);
    const depts = PV.bookingsBy(store, 'department');
    const lifes = PV.lifetimes(store, events);
    const hasPrice = store.families.some((f) => f.stock && f.stock.roles.price != null) || Object.keys(cfg.pricePerFamily).length > 0;
    const hasProject = store.families.some((f) => f.stock && f.stock.roles.project != null);
    const logSince = store.logRows.length ? store.logRows[0].ts : null;
    const collecting = (what, need) => '<div class="collecting">' + esc(what) + '<br><span class="small">' + esc(need) + '</span></div>';

    root.innerHTML =
      '<div class="section-title" style="margin-top:0">Stock overview</div>' +
      '<div class="grid kpis" style="margin-bottom:12px">' +
      '<div class="card kpi"><div class="v">' + totalIn + '</div><div class="l">probes in stock (all Stock and Booking types)</div></div>' +
      '<div class="card kpi"><div class="v">' + totalBooked + '</div><div class="l">booked right now</div></div>' +
      '<div class="card kpi"><div class="v">' + last30 + '</div><div class="l">bookings in the last 30 days</div></div>' +
      '<div class="card kpi"><div class="v">' + (value.n ? PV.ui.fmtMoney(value.value, cfg.currency) : '–') + '</div><div class="l">stock value' + (value.missing ? ' (' + value.missing + ' probes without price)' : '') + (!hasPrice ? ' – needs prices' : '') + '</div></div>' +
      '<div class="card kpi"><div class="v" style="color:' + (low.length ? 'var(--warning-ink)' : 'var(--good-ink)') + '">' + (low.length ? '⚠ ' + low.length : '✓ 0') + '</div><div class="l">' + (low.length ? 'low-stock types: ' + esc(low.map((c) => c.family.cfg.label || c.family.cfg.key).join(', ')) : 'types below their low-stock threshold') + '</div></div>' +
      '</div>' +
      '<div class="grid stock-cards">' + stockFams.map(card).join('') + '</div>' +
      (lifeFams.length ? '<div class="section-title">Only in Probe Lifetime</div><div class="grid stock-cards">' + lifeFams.map(card).join('') + '</div>' : '') +

      '<div class="section-title">Statistics</div>' +
      '<div class="toolbar"><select class="input" id="ovFam" aria-label="Probe type">' + famOptions + '</select>' +
      '<select class="input" id="ovMonths" aria-label="Period">' + [6, 12, 24].map((m) => '<option value="' + m + '"' + (st.months === m ? ' selected' : '') + '>Last ' + m + ' months</option>').join('') + '</select>' +
      '<span class="small muted">' + (logSince ? 'Viewer log since ' + PV.isoDate(logSince) + ' · ' : 'No viewer log yet · ') + 'bookings from "Date of booking" where the sheet has it, otherwise from the log.</span></div>' +
      '<div class="grid stats-grid">' +
      '<div class="card"><div class="card-h"><h3>Bookings per month' + (st.family ? ' · ' + esc(st.family) : '') + '</h3></div><div class="card-b">' +
      (fEvents.length ? '<div class="chart-box"><canvas id="chMonthly" aria-label="Bookings per month"></canvas></div>' + dataTable(['Month', 'Bookings'], monthly.keys.map((k, i) => [k, monthly.values[i]])) : collecting('No booking dates yet for this selection.', 'Filled from "Date of booking" and from bookings made in the viewer.')) + '</div></div>' +

      '<div class="card"><div class="card-h"><h3>Consumption and reorder estimate</h3></div><div class="card-b"><table class="plain"><tr><th>Type</th><th class="num">In stock</th><th class="num">Booked last 90 d</th><th class="num">Per month</th><th>Runs out ≈</th><th>Order by ≈</th></tr>' +
      cons.map((c) => '<tr><td>' + esc(c.family.cfg.label || c.family.cfg.key) + '</td><td class="num">' + c.inStock + '</td><td class="num">' + c.recent + '</td><td class="num">' + PV.ui.fmtNum(c.perMonth) + '</td><td>' + (c.runout ? PV.isoDate(c.runout) : (c.events ? '<span class="faint">no recent use</span>' : '<span class="faint">collecting data</span>')) + '</td><td>' + (c.orderBy ? (c.orderBy < now ? '<b style="color:var(--critical-ink)">now</b>' : PV.isoDate(c.orderBy)) : '') + '</td></tr>').join('') +
      '</table><div class="small muted" style="margin-top:8px">Rate = bookings in the last 90 days ÷ 3. Order-by date subtracts a ' + cfg.reorderLeadTimeWeeks + '-week lead time (config).</div></div></div>' +

      '<div class="card"><div class="card-h"><h3>In stock over time' + (st.family ? ' · ' + esc(st.family) : ' · all types') + '</h3></div><div class="card-b">' +
      (sot ? '<div class="chart-box"><canvas id="chStock" aria-label="In stock over time"></canvas></div>' : collecting('Collecting data since the viewer log starts.', 'Needs the Log table – every status change made in the viewer is recorded there.')) + '</div></div>' +

      '<div class="card"><div class="card-h"><h3>Cost per project</h3></div><div class="card-b">' +
      (cost.rows.length ? '<div class="chart-box"><canvas id="chCost" aria-label="Cost per project"></canvas></div>' + dataTable(['Project', 'Cost (' + cfg.currency + ')', 'Probes'], cost.rows.map((r) => [r.project, PV.ui.fmtNum(r.cost, 0), r.n])) +
        '<div class="small muted">' + (cost.noProject ? cost.noProject + ' booking(s) without project. ' : '') + (cost.unpriced ? cost.unpriced + ' without price.' : '') + '</div>'
        : collecting(hasProject ? 'Collecting data – no bookings with a project yet.' : 'Needs a Project column in Stock and Booking.', (hasPrice ? '' : 'Also needs prices (Price column or pricePerFamily in config). ') + 'Each booking in the viewer asks for a project.')) + '</div></div>' +

      '<div class="card"><div class="card-h"><h3>Booked / sent out per person or customer</h3></div><div class="card-b">' +
      (persons.length ? '<div class="chart-box" style="height:' + Math.max(160, persons.length * 22 + 40) + 'px"><canvas id="chPerson"></canvas></div>' + dataTable(['Booked for', 'Probes'], persons.map((x) => [x.k, x.n])) + '<div class="small muted">Top 12 values of "Booked for" (free text, shown as written).</div>' : collecting('No bookings found.', '')) + '</div></div>' +

      '<div class="card"><div class="card-h"><h3>Probes in use per department</h3></div><div class="card-b">' +
      (depts.length ? '<div class="chart-box" style="height:' + Math.max(140, depts.length * 26 + 40) + 'px"><canvas id="chDept"></canvas></div>' + dataTable(['Department', 'Probes'], depts.map((x) => [x.k, x.n])) + '<div class="small muted">From Probe Lifetime "Department"; spelling variants (Dev / Devel / Development) are grouped for display only.</div>' : collecting('No departments in Probe Lifetime.', '')) + '</div></div>' +

      '<div class="card"><div class="card-h"><h3>Average probe lifetime</h3></div><div class="card-b">' +
      (lifes.length ? '<table class="plain"><tr><th>Type</th><th class="num">Probes</th><th class="num">Median days</th><th class="num">Mean days</th></tr>' + lifes.map((l) => '<tr><td>' + esc(l.family) + '</td><td class="num">' + l.n + '</td><td class="num">' + PV.ui.fmtNum(l.median, 0) + '</td><td class="num">' + PV.ui.fmtNum(l.mean, 0) + '</td></tr>').join('') + '</table><div class="small muted" style="margin-top:8px">Days from booking to the first logged dead / bad / blunt state.</div>'
        : collecting('Collecting data' + (logSince ? ' since ' + PV.isoDate(logSince) : '') + '.', 'Needs the Log: state changes to dead/bad made in the viewer are timed against the booking date.')) + '</div></div>' +
      '</div>';

    root.querySelectorAll('.scard[data-fam]').forEach((c) => (c.onclick = () => { app.state.stock.family = c.dataset.fam; app.go('stock'); }));
    root.querySelector('#ovFam').onchange = (e) => { st.family = e.target.value; app.render(); };
    root.querySelector('#ovMonths').onchange = (e) => { st.months = +e.target.value; app.render(); };

    const q = (id) => root.querySelector('#' + id);
    if (q('chMonthly')) bar(app, q('chMonthly'), monthly.keys, monthly.values, false, 'Bookings');
    if (q('chStock') && sot) {
      const series = st.family ? sot.series.filter((s) => s.family.cfg.key === st.family) : sot.series;
      const total = sot.labels.map((_, i) => series.reduce((a, s) => (s.values[i] == null ? a : (a == null ? 0 : a) + s.values[i]), null));
      const o = baseOptions(false);
      chart(app, q('chStock'), { type: 'line', data: { labels: sot.labels, datasets: [{ label: 'In stock', data: total, borderColor: css('--series-1'), backgroundColor: css('--series-1'), borderWidth: 2, pointRadius: 3, pointHoverRadius: 5, spanGaps: false, tension: 0 }] }, options: o });
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
