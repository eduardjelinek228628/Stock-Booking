/* Probe Viewer – phone view: my probes by box, tap to update state/History, quick booking, search. */
(function () {
  'use strict';
  const PV = window.PV;
  const esc = PV.esc;

  function isMine(p, user) {
    const me = [user.initials, user.name].filter(Boolean).map((x) => x.toLowerCase());
    const owner = PV.currentOwner(PV.ptext(p, 'owner')).toLowerCase();
    if (owner && me.includes(owner)) return true;
    if (p.stock && PV.statusKey(p) === 'booked') {
      const bf = PV.clean(PV.ptext(p, 'bookedFor')).toLowerCase();
      if (bf && me.includes(bf)) return true;
    }
    return false;
  }
  /** Storage box: Box ID for box-level types (Akiyama, Adama), otherwise a "Box 2" / batch "SA1165" mention in Note. */
  function boxOf(p) {
    if (p.family.boxLevel) { const b = PV.ptext(p, 'box') || p.id; if (b) return b; }
    const n = PV.ptext(p, 'note');
    const m = n.match(/\b(box\s*\d+|SA\d{4})\b/i);
    return m ? m[1] : '';
  }
  function wornOut(p) {
    const f = PV.stateFields(p).find((x) => x.role === 'mech' || x.role === 'state');
    return f ? ['dead', 'bad', 'blunt'].includes(PV.key(f.value)) : false;
  }
  function meta(p) {
    const parts = [p.family.label || p.family.key];
    const fr = PV.pget(p, 'freso');
    if (fr) { const unit = (fr && (p.stock || p.life).model.units[fr.col]) || ''; const t = PV.displayCell(fr.v, fr.t, fr.nf); if (t) parts.push('f ' + t + (unit ? ' ' + unit.replace(/[()]/g, '') : '')); }
    const k = PV.ptext(p, 'stiffness'); if (k) parts.push('k ' + k);
    const tip = PV.ptext(p, 'tip'); if (tip) parts.push('tip ' + tip);
    return parts.join(' · ');
  }
  const SHORT = { mech: 'M', ele: 'E', mag: 'Mag', state: '', current: 'I' };
  function chips(p) {
    const st = PV.stateFields(p);
    return PV.ui.statusChip(PV.ptext(p, 'status')) + st.map((x) => PV.ui.stateChip(x.label, x.value, SHORT[x.role])).join('');
  }
  function item(p) {
    return '<button class="pitem" data-uid="' + esc(p.uid) + '"><span class="pmain"><span class="pid">' + esc(PV.logId(p)) + '</span><span class="pmeta">' + esc(meta(p)) + '</span></span><span class="pst">' + chips(p) + '</span></button>';
  }

  function detail(p, app) {
    const store = app.store;
    const fam = store.family(p.family.key);
    const fields = [];
    for (const ref of [p.stock, p.life]) {
      if (!ref) continue;
      ref.model.headers.forEach((h, c) => {
        const t = PV.displayCell(ref.row.values[c], ref.row.text[c], ref.row.nf[c]);
        const role = Object.keys(ref.model.roles).find((r) => ref.model.roles[r] === c);
        if (!t || role === 'history' || role === 'id') return;
        if (ref === p.life && p.stock && ['status', 'length', 'freso', 'q', 'stiffness', 'type', 'box', 'tip'].includes(role)) return;
        fields.push([(PV.clean(h) || role || '') + (ref.model.units[c] ? ' ' + ref.model.units[c] : ''), role === 'status' ? PV.ui.statusChip(t) : esc(t)]);
      });
    }
    const hist = PV.ptext(p, 'history');
    const canState = fam.life && fam.life.isTable && store.can('state');
    const stateRoles = fam.life ? ['mech', 'ele', 'mag', 'state', 'current'].filter((r) => fam.life.roles[r] != null) : [];
    const form = canState ? '<div class="pbox"><div class="pbox-h">Update state</div><div style="padding:12px;display:grid;gap:12px">' +
      stateRoles.map((r) => { const cur = PV.ptext(p, r); return '<div><div class="small muted" style="margin-bottom:6px">' + esc(fam.life.labels[fam.life.roles[r]]) + '</div><div class="seg" data-role="' + r + '">' + store.list(r, p.family).map((o) => '<button type="button" aria-pressed="' + (PV.key(o) === PV.key(cur)) + '" data-v="' + esc(o.trim()) + '">' + esc(o.trim()) + '</button>').join('') + '</div></div>'; }).join('') +
      '<textarea class="input" rows="2" data-note placeholder="Note for History (optional)"></textarea><button class="btn primary" data-save style="height:40px">Save to Probe Lifetime</button></div></div>'
      : (fam.life && !fam.life.isTable ? '<div class="small muted">This sheet is not an Excel Table yet – read-only.</div>' : (!store.writesAllowed ? '<div class="small muted">Read-only (writes switched off).</div>' : ''));
    const book = p.stock && PV.statusKey(p) === 'in stock' && store.can('book') ? '<button class="btn primary" data-book style="height:40px">Book this probe</button>' : '';
    return '<div class="sheet"><div class="phone-top"><div class="phone-title"><button class="btn subtle" data-back aria-label="Back">←</button>' + esc(PV.logId(p)) + '<span class="small muted" style="font-weight:400">' + esc(p.family.label || p.family.key) + '</span></div><div style="display:flex;gap:4px;flex-wrap:wrap">' + chips(p) + '</div></div>' +
      '<div class="phone-body">' + book + form +
      '<div class="pbox"><div class="pbox-h">Parameters</div><dl class="kv" style="padding:10px 12px;grid-template-columns:120px 1fr">' + fields.map(([k, v]) => '<dt>' + esc(k) + '</dt><dd>' + v + '</dd>').join('') + '</dl></div>' +
      (hist ? '<div class="pbox"><div class="pbox-h">History</div><div class="history" style="padding:10px 12px">' + hist.split(store.cfg.historySeparator).map((x) => '<div>' + esc(x) + '</div>').join('') + '</div></div>' : '') +
      '</div></div>';
  }

  PV.views = PV.views || {};
  PV.views.phone = function (root, app, framed) {
    const store = app.store;
    const st = app.state.phone;
    const user = store.user;
    let body = '', top = '';
    if (st.tab === 'mine') {
      let mine = store.probes.filter((p) => isMine(p, user));
      const nAll = mine.length, nWorn = mine.filter(wornOut).length;
      if (st.filter === 'usable') mine = mine.filter((p) => !wornOut(p) && PV.statusKey(p) !== 'missing');
      if (st.filter === 'worn') mine = mine.filter(wornOut);
      const groups = new Map();
      for (const p of mine) {
        const b = boxOf(p);
        const g = b ? (/^box\b/i.test(b) ? b.replace(/^box\s*/i, 'Box ') : 'Box ' + b) + ' · ' + (p.family.label || p.family.key) : (p.family.label || p.family.key);
        if (!groups.has(g)) groups.set(g, []);
        groups.get(g).push(p);
      }
      const sorted = [...groups.entries()].sort((a, b) => a[0].localeCompare(b[0], undefined, { numeric: true }));
      top = '<div class="seg" data-filter><button aria-pressed="' + (st.filter === 'all') + '" data-f="all">All ' + nAll + '</button><button aria-pressed="' + (st.filter === 'usable') + '" data-f="usable">Usable ' + (nAll - nWorn) + '</button><button aria-pressed="' + (st.filter === 'worn') + '" data-f="worn">Worn out ' + nWorn + '</button></div>';
      body = sorted.length ? sorted.map(([g, ps]) => '<div class="pbox"><div class="pbox-h"><span>' + esc(g) + '</span><span>' + ps.length + '</span></div>' +
        ps.sort((a, b) => (wornOut(a) - wornOut(b)) || PV.logId(a).localeCompare(PV.logId(b), undefined, { numeric: true })).map(item).join('') + '</div>').join('')
        : '<div class="empty">No probes with Owner "' + esc(user.initials) + '".<br><span class="small">Probes count as yours when Probe Lifetime Owner or Stock "Booked for" is ' + esc(user.initials) + '.</span></div>';
    } else if (st.tab === 'book') {
      const fams = store.families.filter((f) => f.stock);
      if (!st.bookFam) st.bookFam = fams[0] && fams[0].cfg.key;
      const f = store.family(st.bookFam);
      const avail = f ? f.probes.filter((p) => p.stock && PV.statusKey(p) === 'in stock' && !p.flags.includes('duplicate-id')) : [];
      top = '<select class="input" data-bookfam style="height:40px">' + fams.map((x) => '<option value="' + esc(x.cfg.key) + '"' + (x.cfg.key === st.bookFam ? ' selected' : '') + '>' + esc(x.cfg.label || x.cfg.key) + ' – ' + x.probes.filter((p) => p.stock && PV.statusKey(p) === 'in stock').length + ' in stock</option>').join('') + '</select>';
      body = !store.can('book') ? '<div class="empty">Booking is switched off (read-only phase).</div>' :
        (avail.length ? '<div class="pbox"><div class="pbox-h"><span>Available</span><span>' + avail.length + '</span></div>' + avail.map(item).join('') + '</div>' : '<div class="empty">Nothing in stock for this type.</div>');
    } else {
      top = '<input class="input" data-q type="search" placeholder="Probe ID…" value="' + esc(st.q) + '" style="height:40px" aria-label="Search probe ID">';
      const q = PV.clean(st.q).toUpperCase();
      const hits = q ? store.probes.filter((p) => PV.clean(p.id).toUpperCase().includes(q)).slice(0, 40) : [];
      body = q ? (hits.length ? '<div class="pbox">' + hits.map(item).join('') + '</div>' : '<div class="empty">No probe with that ID.</div>') : '<div class="empty small">Type a probe ID, e.g. from the box label.</div>';
    }
    const title = { mine: 'My probes', book: 'Quick booking', search: 'Search' }[st.tab];
    const p = st.open ? store.byUid.get(st.open) : null;
    const html = '<div class="phone' + (framed ? '' : ' full') + '"><div class="phone-top"><div class="phone-title">' + title + '<span class="spacer" style="flex:1"></span><span class="avatar" title="' + esc(user.name) + '">' + esc(user.initials) + '</span></div>' + top + '</div>' +
      '<div class="phone-body">' + body + '</div>' +
      '<nav class="phone-nav">' + [['mine', '▦', 'My probes'], ['book', '＋', 'Book'], ['search', '⌕', 'Search']].map(([k, ic, l]) => '<button data-ptab="' + k + '" aria-selected="' + (st.tab === k) + '"><span class="ic">' + ic + '</span>' + l + '</button>').join('') + '</nav>' +
      (p ? detail(p, app) : '') + (st.confirm ? confirmSheet(store.byUid.get(st.confirm), store) : '') + '</div>';
    root.innerHTML = framed ? '<div class="phone-stage"><div class="phone-frame">' + html + '</div></div><p class="small muted" style="text-align:center">Same app on a phone: open the page on the phone (it switches to this layout below 760 px width).</p>' : html;

    const q = (s) => root.querySelector(s);
    root.querySelectorAll('[data-ptab]').forEach((b) => (b.onclick = () => { st.tab = b.dataset.ptab; st.open = null; st.confirm = null; app.render(); }));
    root.querySelectorAll('[data-f]').forEach((b) => (b.onclick = () => { st.filter = b.dataset.f; app.render(); }));
    root.querySelectorAll('.phone > .phone-body .pitem').forEach((b) => (b.onclick = () => {
      if (st.tab === 'book') st.confirm = b.dataset.uid; else st.open = b.dataset.uid;
      app.render();
    }));
    const bf = q('[data-bookfam]'); if (bf) bf.onchange = () => { st.bookFam = bf.value; app.render(); };
    const qi = q('[data-q]'); if (qi) { qi.oninput = () => { st.q = qi.value; const pos = qi.selectionStart; app.render(); const n = root.querySelector('[data-q]'); n.focus(); n.setSelectionRange(pos, pos); }; }
    const back = q('[data-back]'); if (back) back.onclick = () => { st.open = null; app.render(); };
    // state form
    root.querySelectorAll('.sheet .seg[data-role] button').forEach((b) => (b.onclick = () => { b.parentElement.querySelectorAll('button').forEach((x) => x.setAttribute('aria-pressed', String(x === b))); }));
    const save = q('[data-save]');
    if (save && p) save.onclick = async () => {
      const ch = {};
      root.querySelectorAll('.sheet .seg[data-role]').forEach((s) => {
        const sel = s.querySelector('[aria-pressed="true"]');
        if (sel && PV.key(sel.dataset.v) !== PV.key(PV.ptext(p, s.dataset.role))) ch[s.dataset.role] = sel.dataset.v;
      });
      const note = q('[data-note]').value;
      save.disabled = true;
      const r = await PV.ui.runAction(() => store.updateState(p, ch, note), 'Saved to Probe Lifetime.');
      save.disabled = false;
      if (r) { const np = store.probes.find((x) => x.family.key === p.family.key && x.key === p.key); st.open = np ? np.uid : null; app.render(); }
    };
    const bk = q('[data-book]'); if (bk && p) bk.onclick = () => { st.confirm = p.uid; app.render(); };
    // confirm booking
    const cb = q('[data-confirm-book]');
    if (cb) {
      const cp = store.byUid.get(st.confirm);
      q('[data-cancel-book]').onclick = () => { st.confirm = null; app.render(); };
      root.querySelectorAll('.sheet .seg[data-proj] button').forEach((b) => (b.onclick = () => { b.parentElement.querySelectorAll('button').forEach((x) => x.setAttribute('aria-pressed', String(x === b))); }));
      cb.onclick = async () => {
        const sel = root.querySelector('.sheet .seg[data-proj] [aria-pressed="true"]');
        const proj = (sel && sel.dataset.v) || PV.clean(root.querySelector('[data-proj-text]').value);
        if (!proj) { PV.ui.toast('Choose a project first.', true); return; }
        cb.disabled = true;
        const r = await PV.ui.runAction(() => store.book(cp, { project: proj }), 'Booked ' + PV.logId(cp) + ' for ' + proj + '.');
        cb.disabled = false;
        if (r) { st.confirm = null; st.open = null; st.tab = 'mine'; app.render(); }
      };
    }
  };

  function confirmSheet(p, store) {
    if (!p) return '';
    const projects = store.list('project');
    return '<div class="sheet"><div class="phone-top"><div class="phone-title"><button class="btn subtle" data-cancel-book aria-label="Back">←</button>Book ' + esc(PV.logId(p)) + '</div></div><div class="phone-body">' +
      '<div class="pbox"><div class="pbox-h">Probe</div>' + item(p).replace('<button class="pitem"', '<button class="pitem" disabled') + '</div>' +
      '<div class="pbox"><div class="pbox-h">Project (required)</div><div style="padding:12px;display:grid;gap:10px">' +
      (projects.length ? '<div class="seg" data-proj>' + projects.map((x) => '<button type="button" aria-pressed="false" data-v="' + esc(x) + '">' + esc(x) + '</button>').join('') + '</div>' : '') +
      '<input class="input" data-proj-text placeholder="' + (projects.length ? 'or type another project' : 'Project name') + '" style="height:40px"></div></div>' +
      '<div class="small muted">Booked for ' + esc(store.user.initials) + ' · today. Writes Stock and Booking and Probe Lifetime together.</div>' +
      '<button class="btn primary" data-confirm-book style="height:44px">Confirm booking</button></div></div>';
  }
  PV._phone = { isMine };
})();
