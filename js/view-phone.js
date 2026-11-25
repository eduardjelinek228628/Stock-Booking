/* Probe Viewer – phone view: my boxes (overview → box → probe), moving probes between boxes, new boxes, quick booking, search. */
(function () {
  'use strict';
  const PV = window.PV;
  const esc = PV.esc;
  const SHORT = { mech: 'M', ele: 'E', mag: 'Mag', state: '' };
  const LOCAL_BOXES = 'probeViewer.boxes.v1'; // new boxes that hold no probe yet live on this device only
  const HEALTH = [['good', 'good', '--good'], ['weak', 'weak', '--warning'], ['bad', 'bad', '--critical'], ['unknown', 'not checked', '--line-2']];
  const HEALTH_ORDER = { bad: 0, weak: 1, unknown: 2, good: 3 };

  function isMine(p, user) {
    const me = [user.initials, user.name].filter(Boolean).map((x) => x.toLowerCase());
    const owner = PV.currentOwner(PV.ptext(p, 'owner')).toLowerCase();
    if (owner) return me.includes(owner);
    if (PV.statusKey(p) === 'booked') { const bf = PV.clean(PV.ptext(p, 'bookedFor')).toLowerCase(); return !!bf && me.includes(bf); }
    return false;
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
  const famName = (p) => p.family.label || p.family.key;
  const byId = (a, b) => PV.logId(a).localeCompare(PV.logId(b), undefined, { numeric: true });

  // ---------------- boxes ----------------
  function localBoxes(user) { const all = PV.storage.get(LOCAL_BOXES) || {}; return (all[user.initials] || []).slice(); }
  function setLocalBoxes(user, names) { const all = PV.storage.get(LOCAL_BOXES) || {}; all[user.initials] = names; PV.storage.set(LOCAL_BOXES, all); }

  /** My probes grouped by box: [{name, probes, counts, local}], named boxes first (natural order), "No box" last. */
  function myBoxes(store, user) {
    const map = new Map();
    for (const p of store.probes) {
      if (!isMine(p, user)) continue;
      const name = PV.boxOf(p);
      const k = PV.key(name);
      if (!map.has(k)) map.set(k, { name, probes: [], local: false });
      map.get(k).probes.push(p);
    }
    const loc = localBoxes(user);
    const keep = loc.filter((n) => !map.has(PV.key(n)));
    if (keep.length !== loc.length) setLocalBoxes(user, keep); // a new box that now holds probes exists in Excel
    for (const n of keep) map.set(PV.key(n), { name: n, probes: [], local: true });
    const boxes = [...map.values()];
    for (const b of boxes) {
      b.counts = { good: 0, weak: 0, bad: 0, unknown: 0 };
      b.probes.forEach((p) => { b.counts[PV.health(p)]++; });
      b.probes.sort((x, y) => (HEALTH_ORDER[PV.health(x)] - HEALTH_ORDER[PV.health(y)]) || byId(x, y));
    }
    return boxes.sort((a, b) => (!a.name) - (!b.name) || a.name.localeCompare(b.name, undefined, { numeric: true }));
  }
  const boxTitle = (name) => name || 'No box';
  /** Next free "Box N" across the whole workbook (box names are physical labels, so avoid other people's numbers too). */
  function nextBoxName(store, user) {
    let max = 0;
    for (const n of store.probes.map(PV.boxOf).concat(localBoxes(user))) { const m = /^Box (\d+)$/i.exec(n); if (m) max = Math.max(max, +m[1]); }
    return 'Box ' + (max + 1);
  }
  function meter(counts) {
    const tot = Math.max(1, counts.good + counts.weak + counts.bad + counts.unknown);
    return '<div class="meter" aria-hidden="true">' + HEALTH.map(([k, , c]) => (counts[k] ? '<i style="width:' + (100 * counts[k] / tot) + '%;background:var(' + c + ')"></i>' : '')).join('') + '</div>';
  }
  function countLine(counts) {
    return '<div class="hcounts">' + HEALTH.filter(([k]) => counts[k] || k === 'good' || k === 'bad').map(([k, l]) => '<span class="hc hc-' + k + '"><b>' + counts[k] + '</b> ' + l + '</span>').join('') + '</div>';
  }
  function boxCard(b, i) {
    const types = new Map();
    b.probes.forEach((p) => types.set(famName(p), (types.get(famName(p)) || 0) + 1));
    const n = b.probes.length;
    return '<button class="boxcard' + (b.name ? '' : ' nobox') + '" data-boxi="' + i + '">' +
      '<span class="bc-h"><span class="bc-name">' + esc(boxTitle(b.name)) + '</span><span class="bc-n">' + n + (n === 1 ? ' probe' : ' probes') + ' ›</span></span>' +
      (n ? meter(b.counts) + countLine(b.counts) + '<span class="bc-types">' + esc([...types.entries()].map(([t, c]) => t + ' ' + c).join(' · ')) + '</span>'
        : '<span class="bc-types">Empty – new box, kept on this device until you move a probe in.</span>') + '</button>';
  }

  // ---------------- list rows ----------------
  /** Row in a box: no status (a probe here is simply mine), its state and datasheet values. */
  function mineItem(p) {
    return '<button class="pitem" data-uid="' + esc(p.uid) + '"><span class="hdot hd-' + PV.health(p) + '" aria-hidden="true"></span><span class="pmain"><span class="pid">' + esc(PV.logId(p)) + '</span><span class="pmeta">' + esc([famName(p), mechLine(p)].filter(Boolean).join(' · ')) + '</span></span><span class="pst">' + stateChips(p) + '</span></button>';
  }
  /** Row in the booking list: everything needed to choose a probe. */
  function bookItem(p) {
    const mech = mechLine(p), el = elecChips(p);
    return '<button class="pitem" data-uid="' + esc(p.uid) + '"><span class="pmain"><span class="pid">' + esc(PV.logId(p)) + '</span>' +
      (mech ? '<span class="pmeta wrap">' + esc(mech) + '</span>' : '') + (el ? '<span class="pst" style="justify-content:flex-start;max-width:none;margin-top:3px">' + el + '</span>' : '') + '</span></button>';
  }
  /** Row in search results: status matters here. */
  function searchItem(p) {
    return '<button class="pitem" data-uid="' + esc(p.uid) + '"><span class="pmain"><span class="pid">' + esc(PV.logId(p)) + '</span><span class="pmeta">' + esc([famName(p), PV.ptext(p, 'owner') || PV.ptext(p, 'bookedFor'), PV.boxOf(p)].filter(Boolean).join(' · ')) + '</span></span><span class="pst">' + PV.ui.statusChip(PV.ptext(p, 'status')) + '</span></button>';
  }
  const sheetTop = (back, title, sub) => '<div class="phone-top"><div class="phone-title"><button class="btn subtle" ' + back + ' aria-label="Back">←</button><span class="ellip">' + esc(title) + '</span>' + (sub ? '<span class="small muted" style="font-weight:400">' + esc(sub) + '</span>' : '') + '</div></div>';

  // ---------------- sheets ----------------
  function detail(p, app) {
    const store = app.store;
    const m = p.stock.model;
    const fields = [];
    m.headers.forEach((h, c) => {
      const t = PV.displayCell(p.stock.row.values[c], p.stock.row.text[c], p.stock.row.nf[c]);
      const role = Object.keys(m.roles).find((r) => m.roles[r] === c);
      if (!t || role === 'history' || role === 'id' || role === 'box') return;
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
    const box = PV.boxOf(p);
    const canMove = isMine(p, store.user) && store.can('state') && m.roles.box != null && m.isTable;
    const boxPanel = isMine(p, store.user) || box ? '<div class="pbox"><div class="pbox-h">Box</div><div class="boxrow"><span><b>' + esc(boxTitle(box)) + '</b></span>' +
      (canMove ? '<button class="btn" data-move>Move to another box</button>' : (m.roles.box == null && isMine(p, store.user) ? '<span class="small muted">No Box column in ' + esc(m.sheetName) + ' yet</span>' : '')) + '</div></div>' : '';
    return '<div class="sheet" data-sheet="detail"><div class="phone-top"><div class="phone-title"><button class="btn subtle" data-back aria-label="Back">←</button>' + esc(PV.logId(p)) + '<span class="small muted" style="font-weight:400">' + esc(famName(p)) + '</span></div><div style="display:flex;gap:4px;flex-wrap:wrap">' + stateChips(p) + elecChips(p) + '</div></div>' +
      '<div class="phone-body">' + book + boxPanel + form +
      '<div class="pbox"><div class="pbox-h">Parameters</div><dl class="kv" style="padding:10px 14px;grid-template-columns:120px 1fr">' + fields.map(([k, v]) => '<dt>' + esc(k) + '</dt><dd>' + v + '</dd>').join('') + '</dl></div>' +
      (hist ? '<div class="pbox"><div class="pbox-h">History</div><div class="history" style="padding:10px 14px">' + hist.split(store.cfg.historySeparator).map((x) => '<div>' + esc(x) + '</div>').join('') + '</div></div>' : '') +
      '</div></div>';
  }

  /** Choose the target box for one probe: my other boxes, or a new one. */
  function moveSheet(p, store) {
    const cur = PV.key(PV.boxOf(p));
    const boxes = myBoxes(store, store.user).filter((b) => b.name && PV.key(b.name) !== cur);
    return '<div class="sheet" data-sheet="move">' + sheetTop('data-move-cancel', 'Move ' + PV.logId(p), 'from ' + boxTitle(PV.boxOf(p))) + '<div class="phone-body">' +
      (boxes.length ? '<div class="pbox"><div class="pbox-h">To one of my boxes</div>' + boxes.map((b) => '<button class="pitem" data-move-to="' + esc(b.name) + '"><span class="pmain"><span class="pid">' + esc(b.name) + '</span><span class="pmeta">' + b.probes.length + ' probes · ' + b.counts.good + ' good · ' + b.counts.bad + ' bad</span></span><span class="muted">›</span></button>').join('') + '</div>' : '') +
      '<div class="pbox"><div class="pbox-h">To a new box</div><div style="padding:12px 14px;display:grid;gap:10px"><input class="input" data-newbox-name value="' + esc(nextBoxName(store, store.user)) + '" style="height:42px" aria-label="New box name"><button class="btn primary" data-move-new style="height:42px">Create box and move</button></div></div>' +
      '</div></div>';
  }

  /** Create a new (empty) box. */
  function newBoxSheet(store) {
    return '<div class="sheet" data-sheet="newbox">' + sheetTop('data-newbox-cancel', 'New box') + '<div class="phone-body">' +
      '<div class="pbox"><div class="pbox-h">Name on the box label</div><div style="padding:12px 14px;display:grid;gap:10px"><input class="input" data-newbox-name value="' + esc(nextBoxName(store, store.user)) + '" style="height:42px" aria-label="Box name">' +
      '<button class="btn primary" data-newbox-ok style="height:42px">Create and add probes</button></div></div>' +
      '<div class="small muted">The box is written to the Box column of Stock and Booking when you move the first probe into it. Until then it is kept on this device.</div></div></div>';
  }

  /** Pick probes from my other boxes to move into box `target`. */
  function addSheet(target, store) {
    const boxes = myBoxes(store, store.user).filter((b) => PV.key(b.name) !== PV.key(target) && b.probes.length);
    const writable = (p) => p.stock.model.roles.box != null && p.stock.model.isTable;
    const body = boxes.length ? boxes.map((b) => '<div class="pbox"><div class="pbox-h"><span>' + esc(boxTitle(b.name)) + '</span><span>' + b.probes.length + '</span></div>' +
      b.probes.map((p) => '<label class="pitem pick-item' + (writable(p) ? '' : ' disabled') + '"><input type="checkbox" data-pick-uid="' + esc(p.uid) + '"' + (writable(p) ? '' : ' disabled') + '><span class="hdot hd-' + PV.health(p) + '" aria-hidden="true"></span><span class="pmain"><span class="pid">' + esc(PV.logId(p)) + '</span><span class="pmeta">' + esc(famName(p)) + (writable(p) ? '' : ' · no Box column yet') + '</span></span><span class="pst">' + stateChips(p) + '</span></label>').join('') + '</div>').join('')
      : '<div class="empty">You have no probes in other boxes.<br><span class="small">Book a probe and choose this box, or move one here later.</span></div>';
    return '<div class="sheet" data-sheet="add">' + sheetTop('data-add-cancel', 'Add probes to ' + target) + '<div class="phone-body">' + body + '</div>' +
      (boxes.length ? '<div class="sheet-foot"><button class="btn primary" data-add-ok disabled style="height:46px;width:100%">Choose probes to move</button></div>' : '') + '</div>';
  }

  function confirmSheet(p, store) {
    if (!p) return '';
    const projects = store.list('project');
    const boxes = myBoxes(store, store.user).filter((b) => b.name);
    const hasBox = p.stock.model.roles.box != null;
    return '<div class="sheet" data-sheet="book">' + sheetTop('data-cancel-book', 'Book ' + PV.logId(p)) + '<div class="phone-body">' +
      '<div class="pbox"><div class="pbox-h">Probe · ' + esc(famName(p)) + '</div>' + bookItem(p).replace('<button class="pitem"', '<button class="pitem" disabled') + '</div>' +
      '<div class="pbox"><div class="pbox-h">Project (required)</div><div style="padding:12px 14px;display:grid;gap:10px">' +
      (projects.length ? '<div class="seg" data-proj>' + projects.map((x) => '<button type="button" aria-pressed="false" data-v="' + esc(x) + '">' + esc(x) + '</button>').join('') + '</div>' : '') +
      '<input class="input" data-proj-text placeholder="' + (projects.length ? 'or type another project' : 'Project name') + '" style="height:42px"></div></div>' +
      (hasBox ? '<div class="pbox"><div class="pbox-h">Put in box (optional)</div><div style="padding:12px 14px"><div class="seg" data-bookbox>' +
        '<button type="button" aria-pressed="true" data-v="">No box</button>' + boxes.map((b) => '<button type="button" aria-pressed="false" data-v="' + esc(b.name) + '">' + esc(b.name) + '</button>').join('') + '</div></div></div>' : '') +
      '<div class="small muted">Booked for ' + esc(store.user.initials) + ' · today. Status, Owner, Box and History are written to Stock and Booking.</div>' +
      '<button class="btn primary" data-confirm-book style="height:46px">Confirm booking</button></div></div>';
  }

  PV.views = PV.views || {};
  PV.views.phone = function (root, app, framed) {
    const store = app.store;
    const st = app.state.phone;
    const user = store.user;
    let body = '', top = '', title = '';
    let boxes = [];
    if (st.tab === 'mine') {
      boxes = myBoxes(store, user);
      const cur = st.box == null ? null : boxes.find((b) => PV.key(b.name) === PV.key(st.box));
      if (st.box != null && !cur) st.box = null;
      if (!cur) {
        // ---- box overview ----
        const nProbes = boxes.reduce((a, b) => a + b.probes.length, 0);
        const tot = { good: 0, weak: 0, bad: 0, unknown: 0 };
        boxes.forEach((b) => Object.keys(tot).forEach((k) => { tot[k] += b.counts[k]; }));
        title = 'My boxes';
        top = '<div class="boxsum"><span>' + nProbes + ' probes · ' + boxes.filter((b) => b.name).length + ' boxes</span>' + (store.can('state') ? '<button class="btn" data-newbox>＋ New box</button>' : '') + '</div>' + (nProbes ? countLine(tot) : '');
        body = boxes.length ? boxes.map(boxCard).join('')
          : '<div class="empty">No probes with Owner "' + esc(user.initials) + '".<br><span class="small">A probe is yours when its Owner (or, for booked probes without Owner, Booked for) is ' + esc(user.initials) + '.</span></div>';
      } else {
        // ---- one box ----
        title = '<button class="btn subtle" data-boxback aria-label="All boxes">←</button>' + esc(boxTitle(cur.name));
        top = cur.probes.length ? meter(cur.counts) + countLine(cur.counts) : '';
        const actions = cur.name && store.can('state') ? '<div class="boxactions"><button class="btn" data-addhere>＋ Add probes here</button>' + (cur.local && !cur.probes.length ? '<button class="btn subtle" data-rmbox>Remove empty box</button>' : '') + '</div>' : '';
        body = actions + (cur.probes.length ? '<div class="pbox">' + cur.probes.map(mineItem).join('') + '</div>'
          : '<div class="empty">This box is empty.<br><span class="small">Add probes from your other boxes, or choose it when booking.</span></div>');
      }
    } else if (st.tab === 'book') {
      title = 'Book a probe';
      const fams = store.families.filter((f) => f.stock);
      if (!st.bookFam || !fams.find((f) => f.cfg.key === st.bookFam)) st.bookFam = fams[0] && fams[0].cfg.key;
      const f = store.family(st.bookFam);
      const avail = f ? f.probes.filter((p) => PV.statusKey(p) === 'in stock' && !p.flags.includes('duplicate-id')) : [];
      top = '<select class="input" data-bookfam style="height:42px">' + fams.map((x) => '<option value="' + esc(x.cfg.key) + '"' + (x.cfg.key === st.bookFam ? ' selected' : '') + '>' + esc(x.cfg.label || x.cfg.key) + ' – ' + x.probes.filter((p) => PV.statusKey(p) === 'in stock').length + ' in stock</option>').join('') + '</select>';
      body = !store.can('book') ? '<div class="empty">Booking is switched off (read-only).</div>' :
        (avail.length ? '<div class="pbox"><div class="pbox-h"><span>In stock</span><span>' + avail.length + '</span></div>' + avail.map(bookItem).join('') + '</div>' : '<div class="empty">Nothing in stock for this type.</div>');
    } else {
      title = 'Search';
      top = '<input class="input" data-q type="search" placeholder="Probe ID…" value="' + esc(st.q) + '" style="height:42px" aria-label="Search probe ID">';
      const q = PV.clean(st.q).toUpperCase();
      const hits = q ? store.probes.filter((p) => PV.clean(p.id).toUpperCase().includes(q)).slice(0, 40) : [];
      body = q ? (hits.length ? '<div class="pbox">' + hits.map(searchItem).join('') + '</div>' : '<div class="empty">No probe with that ID.</div>') : '<div class="empty small">Type a probe ID, e.g. from the box label.</div>';
    }
    const p = st.open ? store.byUid.get(st.open) : null;
    if (st.open && !p) st.open = null;
    const mp = st.move ? store.byUid.get(st.move) : null;
    if (st.move && !mp) st.move = null;
    const html = '<div class="phone' + (framed ? '' : ' full') + '"><div class="phone-top"><div class="phone-title">' + title + '<span class="spacer" style="flex:1"></span><span class="avatar" title="' + esc(user.name) + '">' + esc(user.initials) + '</span></div>' + top + '</div>' +
      '<div class="phone-body">' + body + '</div>' +
      '<nav class="phone-nav">' + [['mine', '▦', 'My boxes'], ['book', '＋', 'Book'], ['search', '⌕', 'Search']].map(([k, ic, l]) => '<button data-ptab="' + k + '" aria-selected="' + (st.tab === k) + '"><span class="ic">' + ic + '</span>' + l + '</button>').join('') + '</nav>' +
      (p ? detail(p, app) : '') + (mp ? moveSheet(mp, store) : '') + (st.newBox ? newBoxSheet(store) : '') + (st.add != null ? addSheet(st.add, store) : '') +
      (st.confirm ? confirmSheet(store.byUid.get(st.confirm), store) : '') + '</div>';
    root.innerHTML = framed ? '<div class="phone-stage"><div class="phone-frame">' + html + '</div></div><p class="small muted" style="text-align:center">On a phone the page switches to this layout automatically.</p>' : html;

    const q = (s) => root.querySelector(s);
    const on = (s, fn) => { const n = q(s); if (n) n.onclick = fn; return n; };
    const reopen = (probe) => { const np = probe && store.family(probe.family.key).probes.find((x) => x.key === probe.key); return np ? np.uid : null; };
    const nameOf = (sel) => PV.normBox(q(sel) ? q(sel).value : '');

    root.querySelectorAll('[data-ptab]').forEach((b) => (b.onclick = () => { st.tab = b.dataset.ptab; Object.assign(st, { open: null, confirm: null, move: null, add: null, newBox: false, box: null }); app.render(); }));
    root.querySelectorAll('[data-boxi]').forEach((b) => (b.onclick = () => { st.box = boxes[+b.dataset.boxi].name; app.render(); }));
    on('[data-boxback]', () => { st.box = null; app.render(); });
    root.querySelectorAll('.phone > .phone-body .pitem').forEach((b) => (b.onclick = () => { if (st.tab === 'book') st.confirm = b.dataset.uid; else st.open = b.dataset.uid; app.render(); }));
    const bf = q('[data-bookfam]'); if (bf) bf.onchange = () => { st.bookFam = bf.value; app.render(); };
    const qi = q('[data-q]'); if (qi) qi.oninput = () => { st.q = qi.value; const pos = qi.selectionStart; app.render(); const n = root.querySelector('[data-q]'); n.focus(); n.setSelectionRange(pos, pos); };
    on('[data-back]', () => { st.open = null; app.render(); });
    root.querySelectorAll('.sheet .seg button').forEach((b) => (b.onclick = () => { b.parentElement.querySelectorAll('button').forEach((x) => x.setAttribute('aria-pressed', String(x === b))); }));

    // ---- state update ----
    const save = q('[data-save]');
    if (save && p) save.onclick = async () => {
      const ch = {};
      root.querySelectorAll('[data-sheet="detail"] .seg[data-role]').forEach((s) => {
        const sel = s.querySelector('[aria-pressed="true"]');
        if (sel && PV.key(sel.dataset.v) !== PV.key(PV.ptext(p, s.dataset.role))) ch[s.dataset.role] = sel.dataset.v;
      });
      save.disabled = true;
      const r = await PV.ui.runAction(() => store.updateState(p, ch, q('[data-note]').value), 'Saved.');
      save.disabled = false;
      if (r) { st.open = reopen(p); app.render(); }
    };

    // ---- move one probe ----
    on('[data-move]', () => { st.move = p.uid; app.render(); });
    on('[data-move-cancel]', () => { st.move = null; app.render(); });
    const doMove = async (probes, target, btn) => {
      if (!target) { PV.ui.toast('Type a box name.', true); return null; }
      if (btn) btn.disabled = true;
      const r = await PV.ui.runAction(() => store.moveToBox(probes, target), null);
      if (btn) btn.disabled = false;
      if (r) PV.ui.toast(r.failed.length ? 'Moved ' + r.moved + ', not moved: ' + r.failed.map((x) => x.id + ' (' + x.error + ')').join('; ') : 'Moved ' + (r.moved === 1 ? PV.logId(probes[0]) : r.moved + ' probes') + ' to ' + PV.normBox(target) + '.', !!r.failed.length);
      return r;
    };
    root.querySelectorAll('[data-move-to]').forEach((b) => (b.onclick = async () => {
      if (await doMove([mp], b.dataset.moveTo, b)) { st.move = null; st.open = reopen(mp); app.render(); }
    }));
    on('[data-move-new]', async function () {
      if (await doMove([mp], nameOf('[data-sheet="move"] [data-newbox-name]'), this)) { st.move = null; st.open = reopen(mp); app.render(); }
    });

    // ---- new box ----
    on('[data-newbox]', () => { st.newBox = true; app.render(); });
    on('[data-newbox-cancel]', () => { st.newBox = false; app.render(); });
    on('[data-newbox-ok]', () => {
      const name = nameOf('[data-sheet="newbox"] [data-newbox-name]');
      if (!name) { PV.ui.toast('Type a box name.', true); return; }
      const exists = myBoxes(store, user).find((b) => PV.key(b.name) === PV.key(name));
      if (!exists) setLocalBoxes(user, localBoxes(user).concat([name]));
      Object.assign(st, { newBox: false, box: exists ? exists.name : name, add: exists ? null : name });
      app.render();
    });
    on('[data-rmbox]', () => { setLocalBoxes(user, localBoxes(user).filter((n) => PV.key(n) !== PV.key(st.box))); st.box = null; app.render(); });

    // ---- add probes to the open box ----
    on('[data-addhere]', () => { st.add = st.box; app.render(); });
    on('[data-add-cancel]', () => { st.add = null; app.render(); });
    const addOk = q('[data-add-ok]');
    if (addOk) {
      const picked = () => [...root.querySelectorAll('[data-pick-uid]:checked')].map((x) => store.byUid.get(x.dataset.pickUid)).filter(Boolean);
      root.querySelectorAll('[data-pick-uid]').forEach((c) => (c.onchange = () => { const n = picked().length; addOk.disabled = !n; addOk.textContent = n ? 'Move ' + n + (n === 1 ? ' probe' : ' probes') + ' to ' + st.add : 'Choose probes to move'; }));
      addOk.onclick = async () => {
        const target = st.add;
        if (await doMove(picked(), target, addOk)) { st.add = null; st.box = target; app.render(); }
      };
    }

    // ---- booking ----
    on('[data-book]', () => { st.confirm = p.uid; app.render(); });
    const cb = q('[data-confirm-book]');
    if (cb) {
      const cp = store.byUid.get(st.confirm);
      on('[data-cancel-book]', () => { st.confirm = null; app.render(); });
      cb.onclick = async () => {
        const sel = root.querySelector('.sheet .seg[data-proj] [aria-pressed="true"]');
        const proj = (sel && sel.dataset.v) || PV.clean(root.querySelector('[data-proj-text]').value);
        if (!proj) { PV.ui.toast('Choose a project first.', true); return; }
        const bsel = root.querySelector('.sheet .seg[data-bookbox] [aria-pressed="true"]');
        const box = bsel ? bsel.dataset.v : '';
        cb.disabled = true;
        const r = await PV.ui.runAction(() => store.book(cp, { project: proj, box }), 'Booked ' + PV.logId(cp) + ' for ' + proj + (box ? ', in ' + box : '') + '.');
        cb.disabled = false;
        if (r) { Object.assign(st, { confirm: null, open: null, tab: 'mine', box: box || '' }); app.render(); }
      };
    }
  };
  PV._phone = { isMine, myBoxes };
})();
