# Probe Viewer

A web viewer over NenoVision's probe workbooks **Stock and Booking.xlsx** and **Probe Lifetime.xlsx** on SharePoint.
The Excel files stay the single source of truth and operations keep working in them as today. The viewer adds:

- **Stock**: one tab per probe type (mirrors the sheets), Excel-like filter and sort on every column, search by probe ID across all types, Lifetime columns (Owner, Department, Mech./Ele., As of) shown next to the Stock row.
- **Overview**: stock per type first (in stock, booked, sent out, missing, usable, tip class, low-stock warnings), then statistics: bookings per month, consumption and reorder estimate, stock over time, cost per project, bookings per person and department, average lifetime, stock value.
- **Import**: test results pasted from Excel or uploaded as .xlsx/.csv, matched by probe ID, previewed (unknown IDs, duplicates, invalid numbers, overwrites) and written only after confirmation.
- **Phone**: "My probes" grouped by box with status colours (two states for electrical probes), tap to update state and add a History note, quick booking, search by ID. The same page switches to this layout on a phone.
- **Data checks**: what the viewer found in the workbooks (missing sheets/columns, duplicates, rows only in Lifetime, status mismatches) and the checklist of Excel changes to agree with operations.

**Booking** writes Status, Booked for, Date of booking and Project into Stock and Booking and creates or updates the probe's Probe Lifetime row in the same action, which removes the double entry.

## Try the demo (no login)

Unzip, open `index.html` by double-clicking. Demo mode uses **synthetic** data with the same sheets, Tables, headers and quirks as the real workbooks. No real probes, customers or people are included. Changes stay in your browser; **Reset demo** restores the sample. Use the user switcher (top right) to see the engineer and operations roles.

## Files

```
index.html            entry page (classic scripts, no build step)
css/app.css           styles (light + dark)
js/config.js          ← the only file you normally edit
js/util.js            helpers (A1 addresses, Excel dates, number parsing)
js/model.js           sheet → model (header roles, units rows, warnings) and Stock↔Lifetime join
js/source-demo.js     demo data source (in-memory "Excel", synthetic data)
js/source-graph.js    live data source (Microsoft Graph Excel API)
js/auth.js            Microsoft 365 sign-in (MSAL.js, redirect flow)
js/store.js           loading + all writes (booking, state, status, import, log)
js/stats.js           Overview statistics
js/view-*.js, ui.js   screens
vendor/               Chart.js 4.6, MSAL Browser 3.30, ExcelJS 4.4 (bundled for offline use; MIT licences included)
tests/                Node tests (demo store, Graph adapter) and a browser click-through
tools/export_json.py  exports a real workbook for the local real-structure test
```

## Going live

### 1. App registration (NenoVision Entra ID admin)

1. Entra admin center → App registrations → **New registration**. Name "Probe Viewer", single tenant.
2. **Authentication** → Add a platform → **Single-page application** → Redirect URI = the exact URL where the viewer is hosted, e.g. `https://nenovision.github.io/probe-viewer/` (and `http://localhost:8080/` for testing). No client secret is needed.
3. **API permissions** → Microsoft Graph → Delegated → **Files.ReadWrite.All** (or Files.Read.All for a read-only pilot). Grant admin consent if the tenant requires it.
4. Copy the **Directory (tenant) ID** and **Application (client) ID**.

The viewer only acts as the signed-in user: a person sees and changes exactly what they could in Excel. The page itself holds no data.

### 2. config.js

```js
mode: 'live',
auth: { tenantId: '<tenant id>', clientId: '<client id>', redirectUri: '', scopes: ['Files.ReadWrite.All'] },
workbooks: {
  stock:    { label: 'Stock and Booking', sharingUrl: '<SharePoint "Copy link" of the file>' },
  lifetime: { label: 'Probe Lifetime',    sharingUrl: '<SharePoint "Copy link" of the file>' },
},
people: { 'eduard.jelinek@nenovision.com': { initials: 'EJ', name: 'Eduard Jelínek', department: 'Apps' } },
roles:  { operations: ['someone@nenovision.com'], importers: ['eduard.jelinek@nenovision.com'] },
```

Also set the low-stock threshold per type (`lowStock`), prices (`pricePerFamily`) until a Price column exists, and the reorder lead time. You can switch modes without editing the file: `?mode=demo` / `?mode=live`.

`liveWritesEnabled: false` keeps live mode read-only (phase 1). Set it to `true` once operations agree with the Excel changes below.

### 3. Hosting

The app is static files. Host them on **GitHub Pages** (a private repository with Pages needs a paid GitHub plan) or **Azure Static Web Apps** (free tier), over https. Live mode does not work from a file opened from disk, because sign-in needs an https origin.

A SharePoint document library is a good place to **store** this package, but it will not **run** it: SharePoint Online downloads .html files instead of running their scripts.

## Excel changes to agree with operations

The viewer runs without these changes (read-only, statistics limited to what exists). They unlock writes and the full statistics:

1. **Log**: new sheet `Log` in Stock and Booking with one Excel Table named `ViewerLog` and exactly these headers:
   `Timestamp | User | Workbook | Sheet | Probe ID | Field | Old value | New value | Project`.
   Every change made in the viewer appends one row per changed cell. Edits made directly in Excel are not logged.
2. **Project** column in every Stock sheet's Table, plus a `projects` list in the `data` sheet (a column block with "projects" in the first cell, as the other lists there).
3. **Prices**: a `Price` column per probe, or prices per type in `config.js`.
4. Convert **Adama** and **CDT-FMR** (Probe Lifetime) into Excel Tables (Ctrl+T). Until then they are read-only.
5. Rename the blank Status header in **NP_Pt-Sharp** to `Status`. It is already handled via config, so this is cosmetic.
6. Unify **Department** spellings (Dev / Devel / Development, Apps / Application). The viewer groups them for display only.
7. Decide where **NP_Pt-Sharp** probes live in Probe Lifetime (`lifetimeSheet` is `null` now, so booking writes only Stock and Booking).

## How it treats the data

- Columns are found **by header name**, never by position (aliases in `config.columns`; trailing spaces, colons and case are ignored).
- Units rows inside Tables (`(µm)`, `(kHz)`, `technique`, `class`) become column units and are hidden. Empty rows and rows without ID are hidden.
- Values are normalised **for display only**: numbers stored as text, "Sent out" vs "sent out", "missing" vs "missing!", zero-width characters. Existing cells are never rewritten silently.
- Allowed values (status, states, departments, projects) are read from the `data` sheets.
- **Writes** go cell by cell (never whole rows), so neighbouring cells and the formula columns (Primary/Secondary technique, Akiyama position) are untouched. The viewer refuses to overwrite a formula cell. Text that starts with `=`, `+`, `-` or `@` is written as text.
- **Before every write** the row is re-read and compared with what the user saw. If someone changed it in Excel meanwhile, nothing is written and the user sees what changed.
- Writes use Graph workbook sessions. Stock and Lifetime are two files, so a booking is not atomic across them. If the Lifetime part fails, the user gets a clear message to fix it by hand.
- Roles: only operations may change Status (Semafor rule); engineers book probes and update state and History; operations and importers import test results.
- New Lifetime rows fill the first empty row of the Table (the sheets have empty rows prepared), otherwise a row is appended. Datasheet values are copied from Stock by matching header names.
- Stock ↔ Lifetime matching is by ID (Akiyama and Adama: Box ID + position/Number). Duplicate IDs are not matched and are listed in Data checks.

## What the real workbooks look like (Nov 2026)

Checked with `tests/real_structure.test.js` on the files in the project:

- **Stock and Booking** sheets: NPC-A, NPC-Z, NPC-Y, NPC-X, NPC-W, NP_Pt-Sharp, NP-Mag, Akiyama, PRSA (all Excel Tables) + Selector, Semafor, data, readme. It has **no** Owner/Department/State/History columns; those exist only in Probe Lifetime. "Date of booking" exists only in NPC-X, NPC-W and NP_Pt-Sharp, so consumption rates work there today and elsewhere fill up from the Log.
- Many probes are in Probe Lifetime without a row in Stock and Booking (NPC-Z 181, NPC-Y 119, NPC-X 107, Akiyama 66, PRSA 15), and some matched probes have a different Status in the two files (NP-Mag 39, NPC-Y 15, NPC-X 13 …). This is the double-entry problem in numbers.
- Duplicate IDs: W289, W294, W296 (Stock NPC-W), MAG-0291 ×3, Akiyama box SA1161 ×6; several in Lifetime.
- Probes with Owner "EJ": 43.

## Tests

```
node tests/store_actions.test.js     # booking, conflicts, formula protection, roles, import, statistics (demo data)
node tests/graph_adapter.test.js     # every Graph request against a fake Graph API: share-link resolution, sessions, retries on 429, single-cell writes
python3 tests/ui_demo_test.py        # clicks through all screens in Chromium (needs playwright), saves screenshots/
node tests/real_structure.test.js    # parser over the real workbooks (export first with tools/export_json.py; data/ is git-ignored)
```

## Deviations from the brief

- **ExcelJS** instead of SheetJS for reading uploaded .xlsx. The current SheetJS builds are only distributed from its own CDN, and the npm version (0.18.5) has known security advisories.
- The brief assumed Stock and Booking has Owner/Department/State/History. It doesn't, so these come from Probe Lifetime via the ID join, and booking writes "Booked for" in Stock and Owner/Department in Lifetime.
- The demo workbook already contains the proposed Project, Price and Log additions, so the pitch can show the full statistics. Live mode shows "collecting data" or "needs …" until they exist.

## Open questions

- Exact SharePoint locations of both files (for the sharing links).
- Who gets the operations role, and who may import test results?
- Is price tracked per probe or per batch?
- Should Probe Lifetime stay a separate file, or should its columns move into Stock and Booking? The viewer works either way; merging would make bookings atomic.
- Which Lifetime sheet do NP_Pt-Sharp probes belong to?
