/* Probe Viewer – small UI helpers (drawer, modal, toast, chips). */
(function () {
  'use strict';
  const PV = window.PV;
  const esc = PV.esc;

  PV.ui = {
    statusChip(text) {
      const k = PV.key(text);
      if (!k) return '<span class="faint">–</span>';
      return '<span class="chip st-' + k.replace(/\s+/g, '-') + '">' + esc(PV.clean(text)) + '</span>';
    },
    stateChip(label, value, short) {
      const k = PV.key(value);
      const cls = k ? 'sv-' + k.replace(/\s+/g, '-') : '';
      return '<span class="chip ' + cls + '" title="' + esc(label + ': ' + (value || '–')) + '">' + (short ? esc(short) + ' ' : '') + esc(PV.clean(value) || '–') + '</span>';
    },
    toast(msg, isErr, ms) {
      let host = document.querySelector('.toast-host');
      if (!host) { host = document.createElement('div'); host.className = 'toast-host'; host.setAttribute('role', 'status'); document.body.appendChild(host); }
      const t = document.createElement('div');
      t.className = 'toast' + (isErr ? ' err' : '');
      t.textContent = msg;
      host.appendChild(t);
      setTimeout(() => t.remove(), ms || (isErr ? 7000 : 3500));
    },
    closeOverlays() { document.querySelectorAll('.scrim, .drawer, .modal, .filter-pop').forEach((n) => n.remove()); },
    drawer(title, bodyHtml, footHtml, onMount) {
      PV.ui.closeOverlays();
      const scrim = document.createElement('div'); scrim.className = 'scrim';
      const d = document.createElement('aside'); d.className = 'drawer'; d.setAttribute('role', 'dialog'); d.setAttribute('aria-label', title);
      d.innerHTML = '<div class="drawer-h"><h2>' + title + '</h2><span class="spacer" style="flex:1"></span><button class="btn subtle" data-close aria-label="Close">✕</button></div>' +
        '<div class="drawer-b">' + bodyHtml + '</div>' + (footHtml ? '<div class="drawer-f">' + footHtml + '</div>' : '');
      document.body.append(scrim, d);
      const close = () => { scrim.remove(); d.remove(); };
      scrim.onclick = close; d.querySelector('[data-close]').onclick = close;
      if (onMount) onMount(d, close);
      return { el: d, close };
    },
    modal(title, bodyHtml, buttons) {
      // buttons: [{label, primary, onClick(el, close) -> false to keep open}]
      const scrim = document.createElement('div'); scrim.className = 'scrim'; scrim.style.zIndex = 80;
      const m = document.createElement('div'); m.className = 'modal'; m.style.zIndex = 81; m.setAttribute('role', 'dialog');
      m.innerHTML = '<div class="card-h"><h3>' + title + '</h3></div><div class="card-b">' + bodyHtml + '</div><div class="drawer-f" style="justify-content:flex-end"></div>';
      const f = m.querySelector('.drawer-f');
      const close = () => { scrim.remove(); m.remove(); };
      (buttons || [{ label: 'OK', primary: true }]).forEach((b) => {
        const el = document.createElement('button'); el.className = 'btn' + (b.primary ? ' primary' : ''); el.textContent = b.label;
        el.onclick = async () => {
          if (!b.onClick) return close();
          el.disabled = true;
          try { const r = await b.onClick(m, close); if (r !== false) close(); } finally { el.disabled = false; }
        };
        f.appendChild(el);
      });
      scrim.onclick = close;
      document.body.append(scrim, m);
      const first = m.querySelector('input, select, textarea'); if (first) first.focus();
      return { el: m, close };
    },
    /** Runs a write action and reports conflicts / partial failures in plain words. */
    async runAction(fn, okMsg) {
      try {
        const r = await fn();
        if (r && r.lifetime === false) PV.ui.toast('Stock and Booking updated, but Probe Lifetime failed: ' + r.lifetimeError + ' – please add it by hand.', true, 12000);
        else PV.ui.toast(okMsg + (r && r.notes && r.notes.length ? ' ' + r.notes.join(' ') : '') + (r && r.logged === false ? ' (not logged)' : ''));
        return r || true;
      } catch (e) {
        if (e.isConflict) {
          const rows = (e.fresh || []).map((x) => '<tr><td>' + esc(x.header) + '</td><td>' + esc(x.was) + '</td><td><b>' + esc(x.now) + '</b></td></tr>').join('');
          PV.ui.modal('Changed in Excel meanwhile', '<p>Someone edited probe <b>' + esc(e.probe.id) + '</b> after you opened it. Nothing was written.</p><table class="plain"><tr><th>Column</th><th>You saw</th><th>Now</th></tr>' + rows + '</table><p class="muted small">Reload, check the new values and try again.</p>',
            [{ label: 'Reload data', primary: true, onClick: async () => { await PV.app.refresh(); } }]);
        } else {
          PV.ui.toast(e.message || String(e), true);
          console.error(e);
        }
        return null;
      }
    },
    fmtNum(n, dp) { if (n == null || !isFinite(n)) return '–'; return n.toLocaleString('en-GB', { maximumFractionDigits: dp == null ? 1 : dp, minimumFractionDigits: 0 }); },
    fmtMoney(n, cur) { if (n == null) return '–'; return n.toLocaleString('en-GB', { maximumFractionDigits: 0 }) + ' ' + (cur || ''); },
    fmtDate(d) { return d ? PV.isoDate(d) : '–'; },
  };
})();
