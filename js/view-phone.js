/* Probe Viewer – phone view: my probes by box, tap to update state/History, quick booking, search. */
(function () {
  'use strict';
  const PV = window.PV;
  const esc = PV.esc;
  const SHORT = { mech: 'M', ele: 'E', mag: 'Mag', state: '' };

  function isMine(p, user) {
    const me = [user.initials, user.name].filter(Boolean).map((x) => x.toLowerCase());
    const owner = PV.currentOwner(PV.ptext(p, 'owner')).toLowerCase();
    if (owner) return me.includes(owner);
    if (PV.statusKey(p) === 'booked') { const bf = PV.clean(PV.ptext(p, 'bookedFor')).toLowerCase(); return !!bf && me.includes(bf); }
    return false;
  }
  /** Storage box: the Box column, else Box ID for box-level types (Akiyama), else a "Box 2" / "SA1165" mention in Note. */
  function boxOf(p) {
    const b = PV.ptext(p, 'box');
    if (b) return b;
    if (p.family.boxLevel) return p.id;
    const m = PV.ptext(p, 'note').match(/\b(box\s*\d+|SA\d{4})\b/i);
    return m ? m[1] : '';
  }
  function wornOut(p) {
    const f = PV.stateFields(p).find((x) => x.role === 'mech' || x.role === 'state');
    return f ? ['dead', 'bad', 'blunt'].includes(PV.key(f.value)) : false;
  }
  function numWithUnit(p, role, short) {
    const c = PV.pget(p, role);
    if (!c) return '';
    const t = PV.displayCell(c.v, c.t, c.nf);
    if (!t) return '';
    const unit = (p.stock.model.units[c.col] || '').replace(/[()]/g, '');
    return short + ' ' + t + (unit ? ' ' + unit : '');
  }
  /** Mechanical datasheet values: length, resonance, stiffness, Q. */
  function mechLine(p) {
    return [numWithUnit(p, 'length', 'L'), numWithUnit(p, 'freso', 'f'), numWithUnit(p, 'stiffness', 'k'), numWithUnit(p, 'q', 'Q')].filter(Boolean).join(' · ');
  }
  /** Electrical / tip info: El. properties, tip current, tip class. */
  function elecChips(p) {
    const out = [];
    const el = PV.ptext(p, 'elProps');
    if (el) { const k = PV.key(el); const cls = k === 'ok' ? 'sv-good' : /via/.test(k) ? 'sv-weak' : 'sv-bad'; out.push('<span class="chip ' + cls + '" title="El. properties">El. ' + esc(el) + '</span>'); }
    const tc = PV.ptext(p, 'tipCurrent');
    if (tc) out.push('<span class="chip ' + (PV.key(tc) === 'yes' ? 'sv-good' : 'sv-bad') + '" title="Tip current">Tip current ' + esc(tc) + '</span>');
    const tip = PV.ptext(p, 'tip');
    if (tip) out.push('<span class="chip" title="Tip class">tip ' + esc(tip) + '</span>');
    return out.join('');
  }
  function stateChips(p) { return PV.stateFields(p).map((x) => PV.ui.stateChip(x.label, x.value, SHORT[x.role])).join(''); }

  /** Row in "My probes": no status (a probe here is simply mine), only its state. */
  function mineItem(p) {
    return '<button class="pitem" data-uid="' + esc(p.uid) + '"><span class="pmain"><span class="pid">' + esc(PV.logId(p)) + '</span><span class="pmeta">' + esc([p.family.label || p.family.key, mechLine(p)].filter(Boolean).join(' · ')) + '</span></span><span class="pst">' + stateChips(p) + '</span></button>';
  }
  /** Row in the booking list: everything needed to choose a probe. */
  function bookItem(p) {
    const mech = mechLine(p), el = elecChips(p);
    return '<button class="pitem" data-uid="' + esc(p.uid) + '"><span class="pmain"><span class="pid">' + esc(PV.logId(p)) + '</span>' +
      (mech ? '<span class="pmeta wrap">' + esc(mech) + '</span>' : '') + (el ? '<span class="pst" style="justify-content:flex-start;max-width:none;margin-top:3px">' + el + '</span>' : '') + '</span></button>';
  }
  /** Row in search results: status matters here. */
  function searchItem(p) {
    return '<button class="pitem" data-uid="' + esc(p.uid) + '"><span class="pmain"><span class="pid">' + esc(PV.logId(p)) + '</span><span class="pmeta">' + esc([p.family.label || p.family.key, PV.ptext(p, 'owner') || PV.ptext(p, 'bookedFor')].filter(Boolean).join(' · ')) + '</span></span><span class="pst">' + PV.ui.statusChip(PV.ptext(p, 'status')) + '</span></button>';
  }

  function detail(p, app) {
    const store = app.store;
    const m = p.stock.model;
    const fields = [];
    m.headers.forEach((h, c) => {
      const t = PV.displayCell(p.stock.row.values[c], p.stock.row.text[c], p.stock.row.nf[c]);
      const role = Object.keys(m.roles).find((r) => m.roles[r] === c);
      if (!t || role === 'history' || role === 'id') return;
      fields.push([(PV.clean(h) || role || '') + (m.units[c] ? ' ' + m.units[c] : ''), role === 'status' ? PV.ui.statusChip(t) : esc(t)]);
    });
    const hist = PV.ptext(p, 'history');
    const states = PV.stateFields(p);
    const canState = (states.length || m.roles.history != null) && m.isTable && store.can('state');
    const form = canState ? '<div class="pbox"><div class="pbox-h">Update state</div><div style="padding:12px 14px;display:grid;gap:12px">' +
      states.map((fl) => '<div><div class="small muted" style="margin-bottom:6px">' + esc(fl.label) + '</div><div class="seg" data-role="' + fl.role + '">' + store.list(fl.role, p.family).map((o) => '<button type="button" aria-pressed="' + (PV.key(o) === PV.key(fl.value)) + '" data-v="' + esc(o) + '">' + esc(o) + '</button>').join('') + '</div></div>').join('') +
      '<textarea class="input" rows="2" data-note placeholder="Note for History (optional)"></textarea><button class="btn primary" data-save style="height:42px">Save</button></div></div>'
      : (!store.writesAllowed ? '<div class="small muted">Read-only (writes switched off).</div>' : (!states.length ? '<div class="small muted">No state columns in ' + esc(m.sheetName) + ' yet – Data checks → Add missing columns.</div>' : ''));
    const book = PV.statusKey(p) === 'in stock' && store.can('book') ? '<button class="btn primary" data-book style="height:42px">Book this probe</button>' : '';
    return '<div class="sheet"><div class="phone-top"><div class="phone-title"><button class="btn subtle" data-back aria-label="Back">←</button>' + esc(PV.logId(p)) + '<span class="small muted" style="font-weight:400">' + esc(p.family.label || p.family.key) + '</span></div><div style="display:flex;gap:4px;flex-wrap:wrap">' + stateChips(p) + elecChips(p) + '</div></div>' +
      '<div class="phone-body">' + book + form +
      '<div class="pbox"><div class="pbox-h">Parameters</div><dl class="kv" style="padding:10px 14px;grid-template-columns:120px 1fr">' + fields.map(([k, v]) => '<dt>' + esc(k) + '</dt><dd>' + v + '</dd>').join('') + '</dl></div>' +
      (hist ? '<div class="pbox"><div class="pbox-h">History</div><div class="history" style="padding:10px 14px">' + hist.split(store.cfg.historySeparator).map((x) => '<div>' + esc(x) + '</div>').join('') + '</div></div>' : '') +
      '</div></div>';
  }

  function confirmSheet(p, store) {
    if (!p) return '';
    const projects = store.list('project');
    return '<div class="sheet"><div class="phone-top"><div class="phone-title"><button class="btn subtle" data-cancel-book aria-label="Back">←</button>Book ' + esc(PV.logId(p)) + '</div></div><div class="phone-body">' +
      '<div class="pbox"><div class="pbox-h">Probe · ' + esc(p.family.label || p.family.key) + '</div>' + bookItem(p).replace('<button class="pitem"', '<button class="pitem" disabled') + '</div>' +
      '<div class="pbox"><div class="pbox-h">Project (required)</div><div style="padding:12px 14px;display:grid;gap:10px">' +
      (projects.length ? '<div class="seg" data-proj>' + projects.map((x) => '<button type="button" aria-pressed="false" data-v="' + esc(x) + '">' + esc(x) + '</button>').join('') + '</div>' : '') +
      '<input class="input" data-proj-text placeholder="' + (projects.length ? 'or type another project' : 'Project name') + '" style="height:42px"></div></div>' +
      '<div class="small muted">Booked for ' + esc(store.user.initials) + ' · today. Status, Owner and History are written to Stock and Booking.</div>' +
      '<button class="btn primary" data-confirm-book style="height:46px">Confirm booking</button></div></div>';
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
        const g = b ? (/^box\b/i.test(b) ? b.replace(/^box\s*/i, 'Box ') : (p.family.boxLevel && !PV.ptext(p, 'box') ? b : 'Box ' + b)) : (p.family.label || p.family.key);
        if (!groups.has(g)) groups.set(g, []);
        groups.get(g).push(p);
      }
      const sorted = [...groups.entries()].sort((a, b) => a[0].localeCompare(b[0], undefined, { numeric: true }));
      top = '<div class="seg"><button aria-pressed="' + (st.filter === 'all') + '" data-f="all">All ' + nAll + '</button><button aria-pressed="' + (st.filter === 'usable') + '" data-f="usable">Usable ' + (nAll - nWorn) + '</button><button aria-pressed="' + (st.filter === 'worn') + '" data-f="worn">Worn out ' + nWorn + '</button></div>';
      body = sorted.length ? sorted.map(([g, ps]) => '<div class="pbox"><div class="pbox-h"><span>' + esc(g) + '</span><span>' + ps.length + '</span></div>' +
        ps.sort((a, b) => (wornOut(a) - wornOut(b)) || PV.logId(a).localeCompare(PV.logId(b), undefined, { numeric: true })).map(mineItem).join('') + '</div>').join('')
        : '<div class="empty">No probes with Owner "' + esc(user.initials) + '".<br><span class="small">A probe is yours when its Owner (or, for booked probes without Owner, Booked for) is ' + esc(user.initials) + '.</span></div>';
    } else if (st.tab === 'book') {
      const fams = store.families.filter((f) => f.stock);
      if (!st.bookFam || !fams.find((f) => f.cfg.key === st.bookFam)) st.bookFam = fams[0] && fams[0].cfg.key;
      const f = store.family(st.bookFam);
      const avail = f ? f.probes.filter((p) => PV.statusKey(p) === 'in stock' && !p.flags.includes('duplicate-id')) : [];
      top = '<select class="input" data-bookfam style="height:42px">' + fams.map((x) => '<option value="' + esc(x.cfg.key) + '"' + (x.cfg.key === st.bookFam ? ' selected' : '') + '>' + esc(x.cfg.label || x.cfg.key) + ' – ' + x.probes.filter((p) => PV.statusKey(p) === 'in stock').length + ' in stock</option>').join('') + '</select>';
      body = !store.can('book') ? '<div class="empty">Booking is switched off (read-only).</div>' :
        (avail.length ? '<div class="pbox"><div class="pbox-h"><span>In stock</span><span>' + avail.length + '</span></div>' + avail.map(bookItem).join('') + '</div>' : '<div class="empty">Nothing in stock for this type.</div>');
    } else {
      top = '<input class="input" data-q type="search" placeholder="Probe ID…" value="' + esc(st.q) + '" style="height:42px" aria-label="Search probe ID">';
      const q = PV.clean(st.q).toUpperCase();
      const hits = q ? store.probes.filter((p) => PV.clean(p.id).toUpperCase().includes(q)).slice(0, 40) : [];
      body = q ? (hits.length ? '<div class="pbox">' + hits.map(searchItem).join('') + '</div>' : '<div class="empty">No probe with that ID.</div>') : '<div class="empty small">Type a probe ID, e.g. from the box label.</div>';
    }
    const title = { mine: 'My probes', book: 'Book a probe', search: 'Search' }[st.tab];
    const p = st.open ? store.byUid.get(st.open) : null;
    const html = '<div class="phone' + (framed ? '' : ' full') + '"><div class="phone-top"><div class="phone-title">' + title + '<span class="spacer" style="flex:1"></span><span class="avatar" title="' + esc(user.name) + '">' + esc(user.initials) + '</span></div>' + top + '</div>' +
      '<div class="phone-body">' + body + '</div>' +
      '<nav class="phone-nav">' + [['mine', '▦', 'My probes'], ['book', '＋', 'Book'], ['search', '⌕', 'Search']].map(([k, ic, l]) => '<button data-ptab="' + k + '" aria-selected="' + (st.tab === k) + '"><span class="ic">' + ic + '</span>' + l + '</button>').join('') + '</nav>' +
      (p ? detail(p, app) : '') + (st.confirm ? confirmSheet(store.byUid.get(st.confirm), store) : '') + '</div>';
    root.innerHTML = framed ? '<div class="phone-stage"><div class="phone-frame">' + html + '</div></div><p class="small muted" style="text-align:center">On a phone the page switches to this layout automatically.</p>' : html;

    const q = (s) => root.querySelector(s);
    root.querySelectorAll('[data-ptab]').forEach((b) => (b.onclick = () => { st.tab = b.dataset.ptab; st.open = null; st.confirm = null; app.render(); }));
    root.querySelectorAll('[data-f]').forEach((b) => (b.onclick = () => { st.filter = b.dataset.f; app.render(); }));
    root.querySelectorAll('.phone > .phone-body .pitem').forEach((b) => (b.onclick = () => { if (st.tab === 'book') st.confirm = b.dataset.uid; else st.open = b.dataset.uid; app.render(); }));
    const bf = q('[data-bookfam]'); if (bf) bf.onchange = () => { st.bookFam = bf.value; app.render(); };
    const qi = q('[data-q]'); if (qi) qi.oninput = () => { st.q = qi.value; const pos = qi.selectionStart; app.render(); const n = root.querySelector('[data-q]'); n.focus(); n.setSelectionRange(pos, pos); };
    const back = q('[data-back]'); if (back) back.onclick = () => { st.open = null; app.render(); };
    root.querySelectorAll('.sheet .seg button').forEach((b) => (b.onclick = () => { b.parentElement.querySelectorAll('button').forEach((x) => x.setAttribute('aria-pressed', String(x === b))); }));
    const save = q('[data-save]');
    if (save && p) save.onclick = async () => {
      const ch = {};
      root.querySelectorAll('.sheet .seg[data-role]').forEach((s) => {
        const sel = s.querySelector('[aria-pressed="true"]');
        if (sel && PV.key(sel.dataset.v) !== PV.key(PV.ptext(p, s.dataset.role))) ch[s.dataset.role] = sel.dataset.v;
      });
      save.disabled = true;
      const r = await PV.ui.runAction(() => store.updateState(p, ch, q('[data-note]').value), 'Saved.');
      save.disabled = false;
      if (r) { const np = store.family(p.family.key).probes.find((x) => x.key === p.key); st.open = np ? np.uid : null; app.render(); }
    };
    const bk = q('[data-book]'); if (bk && p) bk.onclick = () => { st.confirm = p.uid; app.render(); };
    const cb = q('[data-confirm-book]');
    if (cb) {
      const cp = store.byUid.get(st.confirm);
      q('[data-cancel-book]').onclick = () => { st.confirm = null; app.render(); };
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
  PV._phone = { isMine };
})();
