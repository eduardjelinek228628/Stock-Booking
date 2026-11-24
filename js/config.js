/*
 * Probe Viewer – configuration. This is the only file you normally edit.
 * Iteration 2: the viewer works with ONE workbook, "Stock and Booking". Probe Lifetime is no longer used.
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
    // Must match the delegated permission the admin granted:
    //   'Files.SelectedOperations.Selected' – access to the one file the admin assigned (recommended)
    //   'Files.ReadWrite.All'               – access to every file the signed-in user can open
    scopes: ['Files.SelectedOperations.Selected'],
  },

  // Booking, state updates, status changes and the column setup write into the workbook.
  // Set to false for a read-only trial. Demo mode always allows writes (on its own sample data).
  liveWritesEnabled: true,

  // The Stock and Booking file. With Files.SelectedOperations.Selected use driveId + itemId
  // (the admin sees both when granting access; see README). sharingUrl works with Files.ReadWrite.All.
  workbook: { label: 'Stock and Booking', driveId: '', itemId: '', sharingUrl: '' },

  dataSheet: 'data',      // dropdown lists (status, projects, states …) are read from here when present
  semaforSheet: 'Semafor',

  // Where the viewer appends its change log (one Excel Table). Statistics are built from it.
  // "Data checks → Set up columns" creates it if it is missing.
  log: { sheet: 'Log', table: 'ViewerLog' },
  logColumns: ['Timestamp', 'User', 'Sheet', 'Probe ID', 'Field', 'Old value', 'New value', 'Project'],

  /*
   * One entry per Stock and Booking sheet.
   * stateColumns = the state columns of the usage group for this type (Mech./Ele. for electrical probes …).
   * stockColumns overrides the global aliases for that sheet only.
   */
  families: [
    { key: 'NPC-A', group: 'NPC', sheet: 'NPC-A', lowStock: 10, stateColumns: ['Mech.', 'Ele.'] },
    { key: 'NPC-Z', group: 'NPC', sheet: 'NPC-Z', lowStock: 10, stateColumns: ['Mech.', 'Ele.'] },
    { key: 'NPC-Y', group: 'NPC', sheet: 'NPC-Y', lowStock: 10, stateColumns: ['Mech.', 'Ele.'] },
    { key: 'NPC-X', group: 'NPC', sheet: 'NPC-X', lowStock: 15, stateColumns: ['Mech.', 'Ele.'] },
    { key: 'NPC-W', group: 'NPC', sheet: 'NPC-W', lowStock: 15, stateColumns: ['Mech.', 'Ele.'] },
    // Status header is blank (" ") in this sheet.
    { key: 'NP_Pt-Sharp', label: 'Pt-Sharp', group: 'NPC', sheet: 'NP_Pt-Sharp', lowStock: 5, stateColumns: ['Mech.', 'Ele.'],
      stockColumns: { status: ['Status', ''] } },
    { key: 'NP-Mag', label: 'NPM (Mag)', group: 'Magnetic', sheet: 'NP-Mag', lowStock: 10, stateColumns: ['Mech.', 'Mag.'] },
    { key: 'Akiyama', group: 'Akiyama', sheet: 'Akiyama', lowStock: 10, boxLevel: true, stateColumns: ['State'],
      stateLists: { state: ['new', 'good', 'blunt', 'spikes', 'dead', 'unknown'] } },
    { key: 'PRSA', group: 'PRSA', sheet: 'PRSA', lowStock: 3, stateColumns: ['State'] },
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
    tipCurrent: ['Tip current'],
    elProps: ['El. properties'],
    bookedFor: ['Booked for'],
    bookingDate: ['Date of booking'],
    note: ['Note'],
    project: ['Project'],
    price: ['Price', 'Price (EUR)', 'Price (CZK)'],
    // usage group (new columns for the test period)
    owner: ['Owner'],
    department: ['Department'],
    box: ['Box'],
    mech: ['Mech.'],
    ele: ['Ele.'],
    mag: ['Mag.'],
    state: ['State', 'Tip state'],
    asOf: ['As of', 'Last update'],
    history: ['History'],
  },
  // Roles shown in the second ("Usage") column group of the Stock table, in this order.
  usageRoles: ['owner', 'department', 'box', 'mech', 'ele', 'mag', 'state', 'asOf', 'history'],

  // Which named list in the data sheet feeds which field (lists are column blocks headed by these names).
  valueLists: { status: 'status', project: 'projects', department: 'department', mech: 'state mech.', ele: 'state el/mag', mag: 'state el/mag', state: 'state' },
  // Used when a list is not in the data sheet (values from the former Probe Lifetime dropdowns).
  fallbackLists: {
    status: ['to come', 'in stock', 'booked', 'sent out', 'missing!'],
    mech: ['new', 'good', 'bad', 'dead', 'unknown'],
    ele: ['good', 'weak', 'none', 'unknown'],
    mag: ['good', 'weak', 'none', 'unknown'],
    state: ['new', 'good', 'bad', 'dead', 'unknown'],
    department: ['Apps', 'Devel', 'Operations'],
    project: [],
  },
  usableStates: ['new', 'good', 'ok', 'sharp'],

  // Display-only grouping of spelling variants (cells are never rewritten).
  departmentAliases: { dev: 'Devel', devel: 'Devel', development: 'Devel', apps: 'Apps', application: 'Apps', applications: 'Apps' },

  // People: e-mail (lower case) -> initials used in Owner / Booked for. Unknown users fall back to initials from their name.
  people: {
    // 'eduard.jelinek@nenovision.com': { initials: 'EJ', name: 'Eduard Jelínek', department: 'Apps' },
  },
  // Roles. Operations may change Status (Semafor rule), set up columns and import test results;
  // engineers book and update state / History.
  roles: {
    operations: [],          // e-mails
    importers: [],           // e-mails allowed to import test results besides operations
  },
  bookedForFormat: 'initials', // what booking writes into "Booked for" and "Owner": 'initials' or 'name'

  // Prices (until a Price column is filled). Per probe type, in the currency below.
  currency: 'EUR',
  pricePerFamily: {},
  reorderLeadTimeWeeks: 6,
  lowStockDefault: 5,
  historySeparator: ' | ',
};
