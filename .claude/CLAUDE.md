# intrepid-coe-report-gas — Project Guide

## What This Is
Google Apps Script (GAS) Web App that reads the **TH AR Monitoring spreadsheet** and renders
a dashboard showing Working Capital & AR status for Intrepid SEA Thailand.

Deployed via `npx clasp push` from this repo. Bound script (runs inside the spreadsheet).

---

## Architecture

```
Code.gs          ← server-side GAS logic (data read, cache, summary)
dashboard.html   ← single-file SPA served by doGet()
appsscript.json  ← scopes, timezone, web app config
.clasp.json      ← scriptId binding to GAS project
```

`doGet()` → `HtmlService.createTemplateFromFile('dashboard')` → Web App URL

---

## Spreadsheet Layout (TH sheet, 0-based col index)

| Col | Index | Field         | Notes |
|-----|-------|---------------|-------|
| A   | 0     | Type          | Business Class (Retail / Consignment / Store Management / FFM / Live Commerce / Marketing / MBFB / Market place) |
| B   | 1     | Period        | Date object → parsed to "YYYY MMM" by _parsePeriod() |
| D   | 3     | Brand         | Brand name |
| I   | 8     | Issued date   | Invoice issue date. **Blank = unbilled/accrued** |
| J   | 9     | Collected date| Payment received date. **Blank = still outstanding** |
| K   | 10    | AmtAccrue     | Accrual amount ($000 units) |
| L   | 11    | AmtInvoice    | Invoice amount |
| P   | 15    | AmtLatest     | Latest/adjusted amount |
| AA  | 26    | dFin          | Lead time: FIN prepared (days) |
| AB  | 27    | dKAM          | Lead time: KAM/MKT reviewed |
| AC  | 28    | dBrand        | Lead time: Brand reviewed |
| AD  | 29    | dInvoice      | Lead time: Invoice issued |
| AE  | 30    | dCollect      | Lead time: Payment collected |
| AG  | 32    | Status        | AG Group Status string |

### ⚠️ Known Data Issue
Columns K, L, P in TH monitoring sheet are **mostly 0** for pending rows.
The real AR amounts reside in individual brand spreadsheets (not read here due to timeout risk).
This causes totalAR to show near-zero / small negative values.
**Fix needed**: confirm with team which column actually holds the current AR amount.

---

## AR Logic

### Row Classification
```
if collected (col J) is a date  → skip (already collected)
else if issued (col I) is blank → UNBILLED (accrued, not yet invoiced)
else if status == "Pending collection under credit term" → UNDER CREDIT TERM
else if status == "Pending collection over due"         → OVERDUE
else if status == "Pending confirmation"                → PENDING CFM
```

### Amount Selection
```
hasInvoice (col I filled) → amtL || amtK || amtP
no invoice yet            → amtK || amtP
```

### Scope
No period filter applied — AR is cumulative. All rows where collected is blank are
outstanding regardless of accrual period.

### _isDate(v) — treats a value as "has a date" if:
- JS Date object (not NaN)
- Non-empty string that is NOT '0', 'false', '#DIV/0!'

---

## Performance

| Layer | Mechanism |
|-------|-----------|
| Sheet read | Sheets API v4 batchGet — only 6 column ranges (not all 43) |
| Fallback | sheet.getRange(3,1,n,33).getValues() if batchGet fails |
| Cache | CacheService.getScriptCache() 10-min TTL |
| Large data | Chunked 90KB keys (_chunkWrite / _chunkRead) |
| Invalidation | Run clearCache() in GAS editor |

batchGet ranges: A3:B, D3:D, I3:L, P3:P, AG3:AG, AA3:AE

---

## Lead Time Chart

- Rows where type ∈ {Retail, Consignment, Store Management}
- Average days per stage per brand, sorted by total days DESC
- Uses chartjs-plugin-datalabels for inline labels
- Stages: FIN → KAM/MKT → Brand → Invoice → Collected

---

## GAS Functions (public)

| Function | Called by | Purpose |
|----------|-----------|---------|
| doGet() | GAS runtime | Serves the web app HTML |
| getReportData(params) | google.script.run | Summary cards + lead time |
| getDetailRows(params) | google.script.run | Unbilled detail table |
| clearCache() | GAS editor (manual) | Wipe script cache |
| debugData() | GAS editor (manual) | Log full data breakdown |

---

## Deployment

```bash
npx clasp push
```

scriptId: 1QmOgDx9wbt5TQhFrCps22BGjYyosxm5WDnba-V8PwBsJGAcIaslmZXkz

appsscript.json scopes:
- spreadsheets
- drive.readonly (declared, currently unused)
- script.external_request (for batchGet via UrlFetchApp)

---

## Dashboard UI

- Summary tab: 5 cards + lead time chart
- Unbilled Detail tab: table of unbilled rows
- Business Class chips: dynamic from data
- Click "Unbilled (Accrued) ↗" → jumps to Detail tab

---

## Debug Checklist (run debugData() in GAS Editor)

1. total rows — expect ~1500+
2. Periods — expect "2026 May", "2026 Apr" etc.
3. Collected / Unbilled / Invoiced counts
4. Status breakdown — verify strings match STATUS constants in Code.gs
5. Non-zero amount — if low, amounts are in brand sheets (see warning above)
6. Sample unbilled rows — check issued/collected truly blank
