/* Probe Viewer – app shell: mode selection, sign-in, tabs, rendering. */
(function () {
  'use strict';
  const PV = window.PV;
  const esc = PV.esc;
  const cfg = window.PV_CONFIG;

  const DEMO_PERSONAS = [
    { email: 'demo.engineer@example.com', name: 'Eduard Jelínek', initials: 'EJ', department: 'Apps', roles: ['engineer'] },
    { email: 'demo.operations@example.com', name: 'Operations (demo)', initials: 'OP', department: 'Operations', roles: ['engineer', 'operations', 'importer'] },
  ];
  const TABS = [['stock', 'Stock'], ['overview', 'Overview'], ['import', 'Import'], ['phone', 'Phone'], ['checks', 'Data checks']];

  const app = (PV.app = {
    cfg, store: null, auth: null, charts: [],
    narrow: window.matchMedia('(max-width: 760px)'),
    state: {
      tab: 'stock',
      stock: { family: null, filters: {}, sort: {}, showLife: true, q: '' },
      overview: { types: [], period: 12, open: {} },
      import: { mapping: {}, overwrite: false },
      phone: { tab: 'mine', filter: 'all', q: '', open: null, confirm: null },
    },

    mode() {
      const m = new URLSearchParams(location.search).get('mode');
      return m === 'live' || m === 'demo' ? m : cfg.mode;
    },

    async start() {
      document.title = cfg.appName;
      const h = location.hash.replace('#', '');
      if (TABS.some((t) => t[0] === h)) this.state.tab = h;
      if (this.narrow.matches) this.state.tab = 'phone';
      this.narrow.addEventListener('change', () => this.render());
      this.shell('<div class="loading"><div class="spinner"></div><div>Loading Stock and Booking…</div></div>');
      try {
        if (this.mode() === 'live') {
          this.auth = new PV.Auth(cfg);
          const account = await this.auth.init();
          if (!account) return this.signInScreen();
          const user = PV.makeUser(cfg, account.username, account.name);
          const source = new PV.GraphSource({ getToken: () => this.auth.getToken(), workbook: cfg.workbook });
          this.store = new PV.Store(cfg, source, user);
        } else {
          const idx = Math.min(DEMO_PERSONAS.length - 1, +(PV.storage.get('probeViewer.demo.persona') || 0));
          this.store = new PV.Store(cfg, new PV.DemoSource(), Object.assign({}, DEMO_PERSONAS[idx]));
        }
        await this.store.load();
        this.render();
      } catch (e) {
        console.error(e);
        this.shell('<div class="card" style="max-width:640px;margin:40px auto"><div class="card-h"><h2>Could not start</h2></div><div class="card-b"><p>' + esc(e.message || e) + '</p>' +
          (this.mode() === 'live' ? '<p class="small muted">Check js/config.js (tenant, client ID, file links) and the README. <a href="?mode=demo">Open the demo instead</a>.</p>' : '') + '</div></div>');
      }
    },

    signInScreen() {
      this.shell('<div class="card" style="max-width:460px;margin:60px auto;text-align:center"><div class="card-b" style="display:grid;gap:14px;padding:32px">' +
        '<div class="brand" style="justify-content:center;font-size:20px"><span class="brand-mark">PV</span>' + esc(cfg.appName) + '</div>' +
        '<p class="muted">Sign in with your NenoVision Microsoft 365 account. The viewer works directly in ' + esc(cfg.workbook.label) + ' on SharePoint; nothing is stored on this page.</p>' +
        '<button class="btn primary" id="signin" style="justify-self:center;height:40px">Sign in with Microsoft</button><a class="small" href="?mode=demo">Open the demo instead</a></div></div>');
      document.getElementById('signin').onclick = () => this.auth.signIn();
    },

    shell(mainHtml) {
      document.body.innerHTML = '<div id="top"></div><div id="banner"></div><main id="main">' + mainHtml + '</main>';
    },

    go(tab) { this.state.tab = tab; history.replaceState(null, '', '#' + tab); this.render(); },

    async refresh() {
      PV.ui.closeOverlays();
      try { await this.store.load(); this.render(); PV.ui.toast('Reloaded from ' + (this.store.mode === 'demo' ? 'the demo workbook' : 'SharePoint') + '.'); }
      catch (e) { PV.ui.toast('Reload failed: ' + e.message, true); }
    },

    render() {
      if (!this.store) return;
      const store = this.store, narrow = this.narrow.matches;
      if (!document.getElementById('main')) this.shell('');
      this.charts.forEach((c) => c.destroy()); this.charts = [];
      document.querySelectorAll('.filter-pop').forEach((n) => n.remove());
      const tab = narrow && !['phone'].includes(this.state.tab) && this.state.tab !== 'overview' && this.state.tab !== 'checks' ? 'phone' : this.state.tab;
      const u = store.user;
      const persona = store.mode === 'demo'
        ? '<select class="input" id="persona" title="Demo: switch user" aria-label="Demo user">' + DEMO_PERSONAS.map((p, i) => '<option value="' + i + '"' + (p.initials === u.initials ? ' selected' : '') + '>' + esc(p.name) + (p.roles.includes('operations') ? ' · operations' : ' · engineer') + '</option>').join('') + '</select><button class="btn subtle" id="resetDemo" title="Reset demo data">Reset demo</button>'
        : '<span class="uname small muted">' + esc(u.name) + (u.roles.includes('operations') ? ' · operations' : '') + '</span><button class="btn subtle" id="signout">Sign out</button>';
      document.getElementById('top').innerHTML = '<header class="topbar"><div class="brand"><span class="brand-mark">PV</span><span class="t">' + esc(cfg.appName) + '</span></div>' +
        '<span class="mode-badge ' + (store.mode === 'live' ? 'live' : '') + '">' + (store.mode === 'live' ? 'LIVE' : 'DEMO') + '</span>' +
        '<nav class="tabs" role="tablist">' + TABS.filter((t) => !narrow || ['phone', 'overview', 'checks'].includes(t[0])).map(([k, l]) => '<button class="tab" role="tab" data-tab="' + k + '" aria-selected="' + (tab === k) + '">' + (narrow && k === 'checks' ? 'Checks' : l) + '</button>').join('') + '</nav>' +
        '<span class="spacer"></span><div class="userbox">' + (narrow ? '' : persona) +
        '<button class="btn subtle" id="reload" title="Reload from ' + (store.mode === 'demo' ? 'demo data' : 'SharePoint') + (store.loadedAt ? ' (loaded ' + store.loadedAt.toLocaleTimeString() + ')' : '') + '">⟳</button><span class="avatar" title="' + esc(u.name) + '">' + esc(u.initials) + '</span></div></header>';
      const errors = store.warnings.filter((w) => w.level === 'error').length + store.families.reduce((a, f) => a + (f.stock ? f.stock.warnings.filter((w) => w.level === 'error').length : 0), 0);
      document.getElementById('banner').innerHTML =
        (store.mode === 'live' && !store.writesAllowed ? '<div class="banner info">Read-only: live data from SharePoint. Booking and updates are switched off until the Excel changes are agreed (config: liveWritesEnabled).</div>' : '') +
        (store.mode === 'demo' && !narrow ? '<div class="banner info">Demo mode – synthetic sample data in the same structure as Stock and Booking. Changes stay in this browser; “Reset demo” restores the sample.</div>' : '') +
        (errors ? '<div class="banner">' + errors + ' problem(s) reading Stock and Booking – see <a href="#checks" data-go="checks">Data checks</a>.</div>' : '');
      const main = document.getElementById('main');
      main.className = narrow && tab === 'phone' ? 'phone-only' : '';
      if (tab === 'phone') PV.views.phone(main, this, !narrow);
      else PV.views[tab](main, this);

      document.querySelectorAll('.topbar [data-tab]').forEach((b) => (b.onclick = () => this.go(b.dataset.tab)));
      document.querySelectorAll('[data-go]').forEach((a) => (a.onclick = (e) => { e.preventDefault(); this.go(a.dataset.go); }));
      document.getElementById('reload').onclick = () => this.refresh();
      const ps = document.getElementById('persona');
      if (ps) ps.onchange = () => { PV.storage.set('probeViewer.demo.persona', +ps.value); store.user = Object.assign({}, DEMO_PERSONAS[+ps.value]); this.render(); };
      const rd = document.getElementById('resetDemo');
      if (rd) rd.onclick = () => PV.ui.modal('Reset demo data?', '<p>All bookings and changes made in the demo are discarded and the sample data is rebuilt.</p>', [{ label: 'Cancel' }, { label: 'Reset', primary: true, onClick: async () => { store.source.reset(); this.state.phone.open = null; await this.refresh(); } }]);
      const so = document.getElementById('signout'); if (so) so.onclick = () => this.auth.signOut();
    },
  });

  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') PV.ui.closeOverlays(); });
  window.addEventListener('DOMContentLoaded', () => app.start());
})();
