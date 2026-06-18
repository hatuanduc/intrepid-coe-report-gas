# Intrepid COE – Working Capital & AR Report (GAS)

Google Apps Script project generating the AR dashboard.

## Setup

### 1. Install clasp
```bash
npm install
npx clasp login
```

### 2. Create or link GAS project
```bash
# New project (bound to the master spreadsheet):
npx clasp create --type sheets --title "Intrepid COE AR Report" --parentId 1wQFumTUIQCKs2fR4Q4q3xVR4a8QIGYPjcJ1LLcs6bNw

# OR link to existing script – paste scriptId into .clasp.json
```

### 3. Push code
```bash
npx clasp push
```

### 4. Deploy as Web App
```bash
npx clasp deploy --description "v1"
# Or: npx clasp open → Deploy → New deployment → Web App
```

## Column mapping (TH sheet)

| Col | Index | Field |
|-----|-------|-------|
| I   | 8     | 4. Issued invoice (date) |
| J   | 9     | 5. Collected (date) |
| K   | 10    | Amount Accrue USD |
| L   | 11    | Amount Invoice USD |
| P   | 15    | Latest amount USD |
| AG  | 32    | Group Status |

**Group Status values:**
- `Pending collection under credit term`
- `Pending collection over due`
- `Pending confirmation`
- `Collected`

## AR Summary Logic

```
Total AR   = Σ P  where J is blank
Unbilled   = Σ P  where I blank AND J blank
Under term = Σ P  where I filled AND J blank AND AG = "Pending collection under credit term"
Over due   = Σ P  where I filled AND J blank AND AG = "Pending collection over due"
```

## Project structure

```
├── Code.gs          # Backend: data loading, summary, aging, lead-time
├── dashboard.html   # Frontend: summary cards + 2 charts (Chart.js)
├── appsscript.json  # GAS manifest
├── .clasp.json      # clasp config (add your scriptId)
└── package.json     # npm / clasp dev dependency
```

## Git workflow

```bash
git add .
git commit -m "feat: ..."
git push origin main
npx clasp push          # sync to GAS
```
