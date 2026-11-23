/* Probe Viewer – Microsoft 365 sign-in (MSAL.js, redirect flow so it also works on phones). */
(function () {
  'use strict';
  const PV = window.PV;

  PV.Auth = class Auth {
    constructor(cfg) {
      this.cfg = cfg;
      if (!window.msal) throw new Error('MSAL library not loaded (vendor/msal-browser.min.js).');
      if (!cfg.auth.clientId || !cfg.auth.tenantId) throw new Error('Live mode needs auth.tenantId and auth.clientId in js/config.js (see README).');
      if (!window.isSecureContext) throw new Error('Live mode must run over https:// (or http://localhost). Opening index.html from disk only supports demo mode.');
      this.scopes = cfg.auth.scopes;
      this.app = new window.msal.PublicClientApplication({
        auth: {
          clientId: cfg.auth.clientId,
          authority: 'https://login.microsoftonline.com/' + cfg.auth.tenantId,
          redirectUri: cfg.auth.redirectUri || (location.origin + location.pathname),
        },
        cache: { cacheLocation: 'localStorage' },
      });
      this.account = null;
    }
    async init() {
      await this.app.initialize();
      const res = await this.app.handleRedirectPromise();
      if (res && res.account) this.app.setActiveAccount(res.account);
      this.account = this.app.getActiveAccount() || this.app.getAllAccounts()[0] || null;
      if (this.account) this.app.setActiveAccount(this.account);
      return this.account;
    }
    signIn() { return this.app.loginRedirect({ scopes: this.scopes, prompt: 'select_account' }); }
    signOut() { return this.app.logoutRedirect({ account: this.account }); }
    async getToken() {
      try {
        const r = await this.app.acquireTokenSilent({ scopes: this.scopes, account: this.account });
        return r.accessToken;
      } catch (e) {
        if (e instanceof window.msal.InteractionRequiredAuthError) {
          await this.app.acquireTokenRedirect({ scopes: this.scopes, account: this.account });
          return new Promise(() => {}); // page navigates away
        }
        throw e;
      }
    }
  };

  /** Map a signed-in account (or demo persona) to the viewer's user object. */
  PV.makeUser = function (cfg, email, name) {
    const e = String(email || '').toLowerCase();
    const known = cfg.people[e] || {};
    const initials = known.initials || String(name || e).split(/[\s.@_-]+/).filter(Boolean).slice(0, 2).map((x) => x[0].toUpperCase()).join('');
    const roles = ['engineer'];
    if ((cfg.roles.operations || []).map((x) => x.toLowerCase()).includes(e)) roles.push('operations');
    if ((cfg.roles.importers || []).map((x) => x.toLowerCase()).includes(e)) roles.push('importer');
    return { email: e, name: known.name || name || e, initials, department: known.department || '', roles };
  };
})();
