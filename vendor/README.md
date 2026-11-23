# vendor/

The app loads three third-party libraries from this folder so the demo also works offline:

| File | Library | Version |
|---|---|---|
| chart.umd.min.js | Chart.js (`chart.js/dist/chart.umd.min.js`) | 4.6.1 |
| msal-browser.min.js | Microsoft MSAL Browser (`@azure/msal-browser/lib/msal-browser.min.js`) | 3.30.0 |
| exceljs.min.js | ExcelJS (`exceljs/dist/exceljs.min.js`) | 4.4.0 |

They are in probe-viewer.zip. If you only have the source files, restore them with:

```
npm install chart.js@4.6.1 @azure/msal-browser@3.30.0 exceljs@4.4.0
cp node_modules/chart.js/dist/chart.umd.min.js vendor/
cp node_modules/@azure/msal-browser/lib/msal-browser.min.js vendor/
cp node_modules/exceljs/dist/exceljs.min.js vendor/
```

All three are MIT licensed (licence files in this folder).
