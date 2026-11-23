/*
 * Probe Viewer – configuration. This is the only file you normally edit.
 * Sheet names and headers below were taken from "Stock and Booking.xlsx" and "Probe Lifetime.xlsx" (Nov 2026).
 * Columns are always found by header name (the aliases below), never by position.
 */
window.PV_CONFIG = {
  appName: 'Probe Viewer',

  // 'demo' = synthetic sample data, no login. 'live' = Microsoft 365 sign-in + Graph.
  // You can also switch with ?mode=demo / ?mode=live in the URL.
  mode: 'demo',

  // ---- Live mode: Entra ID app registration (single-page application) ----
  auth: {
    tenantId: '',            // NenoVision tenant ID (GUID) or domain, e.g. 'nenovision.onmicrosoft.com'
    clientId: '',            // Application (client) ID of the app registration
    redirectUri: '',         // empty = the page's own URL without query (must be registered as SPA redirect URI)
    scopes: ['Files.ReadWrite.All'], // delegated; use ['Files.Read.All'] for a read-only pilot
  },

  // Live writes stay off until operations agree with the Excel changes (Project column, Log table ...).
  // Demo mode always allows writes (on its own sample data).
  liveWritesEnabled: false,

  workbooks: {
    // Paste the SharePoint "Copy link" URL of each file (people in NenoVision link is fine),
    // or fill driveId + itemId directly.
    stock: { label: 'Stock and Booking', sharingUrl: '', driveId: '', itemId: '' },
    lifetime: { label: 'Probe Lifetime', sharingUrl: '', driveId: '', itemId: '' },
  },

  dataSheet: 'data',      // dropdown lists (status, states, departments, projects) are read from here
  semaforSheet: 'Semafor',

  // Where the viewer appends its change log (one Excel Table). Statistics are built from it.
  log: { workbook: 'stock', sheet: 'Log', table: 'ViewerLog' },

  /*
   * One entry per probe type. stockSheet / lifetimeSheet = worksheet names (null = not in that file).
   * Matching rows between the files uses the ID column (plus position/Number for box-level Akiyama rows).
   * stockColumns / lifetimeColumns override the global aliases for that sheet only.
   * valueLists overrides which dropdown list in the data sheet is used for a field.
   */
  families: [
    { key: 'NPC-A', group: 'NPC', stockSheet: 'NPC-A', lifetimeSheet: 'NPC-A', lowStock: 10 },
    { key: 'NPC-Z', group: 'NPC', stockSheet: 'NPC-Z', lifetimeSheet: 'NPC-Z', lowStock: 10 },
    { key: 'NPC-Y', group: 'NPC', stockSheet: 'NPC-Y', lifetimeSheet: 'NPC-Y', lowStock: 10 },
    { key: 'NPC-X', group: 'NPC', stockSheet: 'NPC-X', lifetimeSheet: 'NPC-X', lowStock: 15 },
    { key: 'NPC-W', group: 'NPC', stockSheet: 'NPC-W', lifetimeSheet: 'NPC-W', lowStock: 15 },
    // Status header is blank (" ") in this sheet. Lifetime counterpart unknown -> null (see README, open questions).
    { key: 'NP_Pt-Sharp', label: 'Pt-Sharp', group: 'NPC', stockSheet: 'NP_Pt-Sharp', lifetimeSheet: null, lowStock: 5,
      stockColumns: { status: ['Status', ''] } },
    { key: 'NP-Mag', label: 'NPM (Mag)', group: 'Magnetic', stockSheet: 'NP-Mag', lifetimeSheet: 'NPM ("Maki")', lowStock: 10 },
    { key: 'Akiyama', group: 'Akiyama', stockSheet: 'Akiyama', lifetimeSheet: 'Akiyama', lowStock: 10, boxLevel: true,
      valueLists: { state: 'state akiyama' } },
    { key: 'PRSA', group: 'PRSA', stockSheet: 'PRSA', lifetimeSheet: 'PRSA', lowStock: 3,
      lifetimeColumns: { history: ['Note'] } },
    // Only in Probe Lifetime:
    { key: 'FEBID-NPC', group: 'NPC', stockSheet: null, lifetimeSheet: 'FEBID-NPC' },
    { key: 'NPC-BDD', group: 'NPC', stockSheet: null, lifetimeSheet: 'NPC-BDD' },
    { key: 'Trilayer', group: 'NPC', stockSheet: null, lifetimeSheet: 'Trilayer' },
    { key: 'NP-200', group: 'NPC', stockSheet: null, lifetimeSheet: 'NP-200' },
    { key: 'Adama', group: 'Adama', stockSheet: null, lifetimeSheet: 'Adama', boxLevel: true, headerRow: 2,
      valueLists: { state: 'adama tip state', current: 'adama current' } },
    { key: 'CDT-FMR', group: 'Adama', stockSheet: null, lifetimeSheet: 'CDT-FMR', boxLevel: true, headerRow: 2 },
  ],

  // Global header aliases per role (matched after trimming, collapsing spaces, dropping a trailing ":" and ignoring case).
  columns: {
    id: ['ID', 'Box ID'],
    position: ['position', 'Number'],
    status: ['Status'],
    type: ['Type'],
    length: ['Length'],
    freso: ['freso', 'freso (Hz)', 'freso (kHz)', 'fr [kHz]'],
    q: ['Q factor'],
    stiffness: ['Stiffness', 'k [N/m]'],
    tip: ['Tip'],
    box: ['Box ID', 'Original box ID', 'Original Akiyama box'],
    bookedFor: ['Booked for'],
    bookingDate: ['Date of booking'],
    note: ['Note'],
    project: ['Project'],
    price: ['Price', 'Price (EUR)', 'Price (CZK)'],
    department: ['Department'],
    owner: ['Owner'],
    mech: ['Mech.'],
    ele: ['Ele.'],
    mag: ['Mag.'],
    state: ['State', 'Tip state'],
    current: ['Current'],
    asOf: ['As of', 'Last update'],
    history: ['History'],
  },

  // Headers of the Log table (create it with exactly these headers, see README).
  logColumns: ['Timestamp', 'User', 'Workbook', 'Sheet', 'Probe ID', 'Field', 'Old value', 'New value', 'Project'],

  // Which named list in the data sheet feeds which field. Lists are column blocks headed by these names.
  valueLists: {
    status: 'status', mech: 'state mech.', ele: 'state el/mag', mag: 'state el/mag', state: 'state mech.',
    current: 'adama current', department: 'department', project: 'projects',
  },
  // Used only when a list is not found in the data sheet.
  fallbackLists: {
    status: ['to come', 'in stock', 'booked', 'sent out', 'missing!'],
    state: ['new', 'good', 'weak', 'bad', 'blunt', 'spikes', 'dead', 'none', 'unknown'],
    department: ['Apps', 'Devel'],
    project: [],
  },
  usableStates: ['new', 'good', 'ok', 'sharp'],

  // Display-only grouping of spelling variants (cells are never rewritten).
  departmentAliases: { dev: 'Devel', devel: 'Devel', development: 'Devel', apps: 'Apps', application: 'Apps', applications: 'Apps' },

  // People: e-mail (lower case) -> initials used in Owner / Booked for. Unknown users fall back to initials from their name.
  people: {
    // 'eduard.jelinek@nenovision.com': { initials: 'EJ', name: 'Eduard Jelínek', department: 'Apps' },
  },
  // Roles. Operations may change Status (Semafor rule) and import test results; engineers book and update state/History.
  roles: {
    operations: [],          // e-mails
    importers: [],           // e-mails allowed to import test results besides operations
  },
  bookedForFormat: 'initials', // what booking writes into "Booked for": 'initials' or 'name'

  // Prices (until a Price column exists). Per probe type, in the currency below.
  currency: 'EUR',
  pricePerFamily: {},
  reorderLeadTimeWeeks: 6,
  lowStockDefault: 5,
  historySeparator: ' | ',
};
