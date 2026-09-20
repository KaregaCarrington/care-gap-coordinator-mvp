# Care Gap Coordinator MVP

Local-first Electron desktop prototype for automated care-gap review using a payer list and a generic EMR export.

## Important
The included measures are **synthetic test rules only**. Do not use them to make real clinical/HEDIS decisions. Replace them with verified payer/HEDIS specifications before production use.

## Architecture
- Electron + React + TypeScript + Vite
- No backend server
- No database
- Rules persist locally in Electron's user-data directory (`rules.json`)
- Excel/CSV parsing is local
- Export is only performed when the user clicks **Export Excel**
- `exceljs` is used for styled Notes cells

## Run locally
Requirements: Node.js 18+.

```bash
npm install
npm run dev
```

For a production build:

```bash
npm run dist
```

Installers are written to `release/`.

## Test data
Synthetic samples are included in `samples/` and contain no real PHI.

## Workflow
1. Upload payer file.
2. Upload EMR export.
3. Confirm/edit generic EMR column mapping and select measures.
4. Run automated review.
5. Filter/group results.
6. Manually handle only rows determined to be closed/green for external supplemental-data submission.
7. Click Export Excel.

The export keeps the payer columns and row order and adds one `Notes` column. Notes cells are colored green for Closed, yellow for Open, and gray for Not Found.
