# Probe Viewer

A web viewer over NenoVision's **Stock and Booking** workbook on SharePoint. The Excel file stays the single source of truth and operations keep working in it as today. Iteration 2 is the first version meant for real use during a test period.

- **Stock**: one tab per probe type (mirrors the sheets), Excel-like filter and sort on every column, search by probe ID across all types. Columns come in two groups: **Stock and Booking** (datasheet, status, booking) and **Usage** (Owner, Department, Box, Mech./Ele. or State, As of, History).
- **Overview**: compact cards per type (in-stock count, colour bar, status numbers; "More info" shows usable/worn, stock by tip class, last booking). Below that are the statistics. "Bookings per month/week" and "In stock over time" sit side by side with their own filters: pick several types by clicking, and a period from the last month to all time. Then come consumption and reorder estimates, cost per project, bookings per person and department, and average lifetime.
- **Phone**: "My probes" grouped by box, showing only each probe's state (no booked/in-stock chip: a probe there is simply yours). Tap a probe to update its state and add a History note. The booking list shows the mechanical values (L, f, k, Q) and the electrical ones (El. properties, tip current, tip class) for each probe. There is also search by ID. On a phone the page switches to this layout by itself.
- **Data checks**: findings in the workbook, the setup checklist and **Add missing columns**.
- **Import** (low priority): test results pasted from Excel or uploaded, previewed before writing.

**Booking** writes Status, Booked for, Date of booking, Project, Owner, Department, As of, History and the initial state (`new`) into the probe's row. Everything is in one row of one file, so there's no double entry.

## Iteration 2 changes

- **Probe Lifetime is no longer used.** Everything lives in Stock and Booking. The Lifetime fields become new columns there for the test period.
- **New columns** per sheet, placed where they make sense rather than appended at the end: `Price` in front of *Booked for*; then after the booking columns `Project, Owner, Department, Box, <state>, As of, History`, in front of *Note*. `<state>` is `Mech., Ele.` for NPC types and Pt-Sharp, `Mech., Mag.` for NP-Mag, and `State` for Akiyama and PRSA. Plus a `Log` sheet with Table `ViewerLog`.
- **Add missing columns** (Data checks, operations role) inserts them with Excel's own "insert table column" through Microsoft Graph. Existing columns, values and formulas move along, and nothing is overwritten. It shows a preview first, and it was checked against the real file's structure: every sheet gets all columns automatically, none needs adding by hand. **Run it on a copy of the file first.** You can also add the columns by hand; the viewer finds them by header name.
- Bright, light design (it no longer follows the dark system setting).
- Compact overview cards with "More info"; the two time charts side by side with multi-type and period filters (1, 2, 3, 6, 12, 24 months, all time; weekly bins up to 3 months).
- Phone: no status chip in My probes; electrical and mechanical properties in the booking list.
- The Log is written in one request per action.

## Try the demo (no login)

Open `index.html` (double-click), or the GitHub Pages address. The demo uses **synthetic** data with the same sheets, Tables, headers and quirks as Stock and Booking, already including the new columns and the Log. To try the setup itself: Data checks → **Demo: start without new columns**, switch the user (top right) to *Operations*, then **Add missing columns…**. **Reset demo** restores the sample.

## Files

```
index.html            entry page (classic scripts, no build step)
css/app.css           styles
js/config.js          ← the only file you normally edit
js/util.js            helpers (A1 addresses, Excel dates, number parsing)
js/model.js           sheet → model (header roles, units rows, warnings), column setup plan
js/source-demo.js     demo data source (in-memory "Excel", synthetic data)
js/source-graph.js    live data source (Microsoft Graph Excel API)
js/auth.js            Microsoft 365 sign-in (MSAL.js, redirect flow)
js/store.js           loading + all writes (booking, state, status, import, log, column setup)
js/stats.js           Overview statistics
js/view-*.js, ui.js   screens
vendor/               Chart.js 4.6.1, MSAL Browser 3.30.0, ExcelJS 4.4.0 (MIT)
tests/                Node tests (store, Graph adapter), browser click-through, real-structure check
tools/export_json.py  exports a real workbook for the local real-structure check
```

## Going live

### 1. App registration (NenoVision Entra ID admin)

1. Entra admin center → App registrations → **New registration**. Name "Probe Viewer", single tenant.
2. **Authentication** → Add a platform → **Single-page application** → Redirect URI = the exact hosting URL, e.g. `https://eduardjelinek228628.github.io/Stock-Booking/`. No client secret.
3. **API permissions** → Microsoft Graph → Delegated → **Files.SelectedOperations.Selected**, with admin consent. Then grant the app read/write on the Stock and Booking file only, through the file's `/permissions` in Graph; this step needs a SharePoint admin. The app then sees no other file. (Alternative: Files.ReadWrite.All, which covers every file the signed-in user can open.)
4. Send back the tenant ID, client ID, and the file's **drive ID** and **item ID** (the admin sees both while granting access).

### 2. config.js

```js
mode: 'live',
auth: { tenantId: '<tenant id>', clientId: '<client id>', redirectUri: '', scopes: ['Files.SelectedOperations.Selected'] },
workbook: { label: 'Stock and Booking', driveId: '<drive id>', itemId: '<item id>', sharingUrl: '' },
people: { 'eduard.jelinek@nenovision.com': { initials: 'EJ', name: 'Eduard Jelínek', department: 'Apps' } },
roles:  { operations: ['someone@nenovision.com'], importers: [] },
```

With Files.ReadWrite.All you can use `sharingUrl` (the SharePoint "Copy link") instead of the two IDs. Also set `lowStock` per type, `pricePerFamily` if prices aren't in the sheet, and `reorderLeadTimeWeeks`. `liveWritesEnabled: false` makes live mode read-only. You can switch modes without editing the file: `?mode=demo` / `?mode=live`.

### 3. Hosting

Static files on **GitHub Pages** (this repo: Settings → Pages → branch `main`, folder `/`) or Azure Static Web Apps. Live mode needs https. SharePoint can store the package but won't run it.

## How it treats the data

- Columns are found **by header name**, never by position (aliases in `config.columns`; trailing spaces, colons and case are ignored).
- Units rows inside Tables (`(µm)`, `(kHz)`, `technique`, `class`) become column units and are hidden. Empty rows and rows without an ID are hidden.
- Values are normalised **for display only** (numbers stored as text, "Sent out"/"sent out", "missing"/"missing!", zero-width characters). Existing cells are never rewritten silently.
- **Writes** go cell by cell, never whole rows, so formula columns (Primary/Secondary technique, Akiyama position) stay intact. The viewer refuses to overwrite a formula cell. Text starting with `=`, `+`, `-` or `@` is written as text.
- **Before every write** the row is re-read and compared with what the user saw. If someone changed it in Excel meanwhile, nothing is written and the user sees what changed.
- Roles: only operations may change Status (Semafor rule) and the workbook structure; engineers book probes and update state and History; operations and importers import test results.
- "My probes" = Owner is me (the last name in "PB -> EJ"), or, for booked probes without Owner, Booked for is me.
- Edits made directly in Excel are not logged; the statistics use the Log plus "Date of booking" where the sheet has it.

## Tests

```
npm test                              # store + Graph adapter (Node, no dependencies)
python3 tests/ui_demo_test.py         # clicks through every screen in Chromium (needs playwright)
node tests/real_structure.test.js     # parser + column plan over the real file (export first with tools/export_json.py; data/ is git-ignored)
```

## Open questions

- Who gets the operations role, and who may import test results?
- Is price tracked per probe or per batch? The viewer supports a Price column (per probe) and a per-type fallback in config.
