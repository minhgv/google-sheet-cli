---
name: google-sheet
description: Read, write, and manage Google Sheets AND local Excel (.xlsx) files from the terminal using the google-sheet-cli oclif CLI. Supports OAuth 2.0 user auth and Service Account JWT for cloud, plus fully-offline XLSX inspect/read/write, declarative report automation (finance cash-flow, manpower estimation) from templates, cell formatting (borders, bold, colors, number formats, merge), and cell search returning A1 coordinates. Use when the user asks to read/update a Google Sheet, batch-read ranges, append table rows, create spreadsheets, work with a docs.google.com/spreadsheets URL, inspect or edit a local .xlsx, generate a report from CSV/JSON via a template, find a cell/row by value, or format a sheet for a pretty report. Trigger on "đọc sheet", "google sheet", "spreadsheet", "xlsx", "excel", "report", "báo cáo", "data:get", "data:find", "format:cells", "gsheet", or a Google Sheets URL.
---

# google-sheet-cli

CLI + npm library wrapping Google Sheets v4 API + local XLSX engine + declarative report runner.

## Invocation

- Published package: `npx google-sheet-cli <command>`. Commands `workbook:*` and `report:*` exist from 3.x — if a version lacks them, build from source.
- From source: `git clone <repo> && npm ci && npm run build`, then `node bin/run.js <command>` (requires Node >= 22).
- Generated command reference: `docs/*.md` in the repository, or `<command> --help`.

## Authentication (cloud commands only)

| Mode | When | Setup |
|------|------|-------|
| **OAuth 2.0** (preferred) | User's own sheets, no sharing needed | `auth:login` once → browser login → token cached at `~/.config/google-sheet-cli/token.json` |
| **Service Account** | CI, shared sheets | `-f <credentials.json>` or env `GSHEET_CLIENT_EMAIL` + `GSHEET_PRIVATE_KEY`; sheet MUST be shared with the SA email |

Auth resolution: OAuth is used automatically when `--useOauth` is set OR no service-account flags/env are present. Service account stays explicit via `-c/-p/-f`.

- `auth:status` — check token validity. `auth:logout` — delete cached tokens.
- OAuth client secret expected at `~/.config/google-sheet-cli/client_secret.json`; override with `--clientSecretFile`.
- Advanced Protection Program accounts CANNOT consent to unverified OAuth apps — use a secondary account or service account.
- **Local-only commands (`workbook:*`, `report:run` with only local input/output) need NO auth and make NO network calls.** Auth initializes only when `--sourceSpreadsheet` or `--spreadsheetId` is used.

## Commands

```
# Cloud (needs auth)
data:get          -s <id> -t <title> [--minRow --maxRow --range --hasHeaderRow --rawOutput]
                  [--valueRenderOption FORMATTED_VALUE|UNFORMATTED_VALUE|FORMULA]
                  [--dateTimeRenderOption SERIAL_NUMBER|FORMATTED_STRING]
data:append       -s <id> -t <title> -d '<json rows>' [--rawDataInput]
data:update       -s <id> -t <title> -d '<json rows>' [--minCol --minRow]
data:append-table -s <id> -t <title> -i <file|-> [--inputFormat json|csv] [--range Sheet1!A1]
                  [--valueInputOption] [--insertDataOption OVERWRITE|INSERT_ROWS]
data:batch-get    -s <id> --ranges '["Sheet1!A1:B10","Sheet2!C1:D5"]' [--chunkSize] [render options]
data:batch-update -s <id> -i <file|-> [--valueInputOption] [--dryRun] [--overwriteFormulas]
                  [--chunkByteSize --maxRowsPerChunk]     payload: [{"range":"A1","values":[[...]]}]
data:find         -s <id> -t <title> (--equals|--contains|--regex <v>) [--range A1:Z100]
                  [--column B | --header "Status"] [--ignoreCase|--no-ignoreCase]
                  [--limit N|--first] [--byRow] [render options] → A1 coordinates JSON
format:cells      -s <id> -t <title> --range A1:J1 [--bold --italic --underline --strikethrough]
                  [--fontSize --fontFamily --textColor #RRGGBB --backgroundColor #RRGGBB]
                  [--horizontalAlignment LEFT|CENTER|RIGHT] [--verticalAlignment TOP|MIDDLE|BOTTOM]
                  [--wrapStrategy OVERFLOW_CELL|CLIP|WRAP] [--numberFormat "#,##0.00" [--numberFormatType]]
                  [--borders top,bottom,all,inner --borderStyle SOLID_THICK --borderColor #RRGGBB]
                  [--clear] [-i spec.json|-] [--dryRun]   format-only; never touches values
format:merge      -s <id> -t <title> --range A1:J1 [--type MERGE_ALL|MERGE_COLUMNS|MERGE_ROWS|--unmerge] [--dryRun]
grid:insert       -s <id> -t <title> --dimension ROWS|COLUMNS --start N [--count N] [--inheritFromBefore] [--dryRun]
grid:delete       -s <id> -t <title> --dimension ROWS|COLUMNS --start N [--count N] [--dryRun]
                  (dryRun previews the values about to be removed)
grid:hide         -s <id> -t <title> --dimension ROWS|COLUMNS --start N [--count N] [--unhide] [--dryRun]
grid:resize       -s <id> -t <title> --dimension ROWS|COLUMNS --start N [--count N] (--pixels N|--auto) [--dryRun]
grid:freeze       -s <id> -t <title> [--rows N] [--columns N] [--dryRun]   (0 unfreezes that axis)
spreadsheet:share    -s <id> (--email <addr>...|--domain <d>|--anyone) [--role reader|commenter|writer] [--notify --message <text>]
spreadsheet:permissions -s <id> [--rawOutput]
spreadsheet:unshare -s <id> (--permissionId <id>|--email <addr>)
spreadsheet:add   --spreadsheetTitle <name>
spreadsheet:get   -s <id> [--rawOutput]
worksheet:add|get|remove -s <id> -t <title>
worksheet:rename  -s <id> -t <old> -n <new>

# Local XLSX (offline, no auth, no network)
workbook:inspect  -f <file.xlsx>
workbook:read     -f <file.xlsx> --range 'Sheet1!A1:E20' [--mode unformatted|formatted|formula]
workbook:write    [-f template.xlsx] [-o out.xlsx] -i <data.csv|json|-> [-t Sheet1] [--startCell A1]
                  [--inPlace] [--dryRun] [--overwrite] [--overwriteFormulas] [-r]
                  (or pass data as positional JSON 2D array / ReportDocument)

# Reports (template-driven)
report:run --template <template.json> \
  --input <data.csv|json|-> [--inputFormat]   # exactly ONE source: this, or
  --sourceWorkbook <f.xlsx> --ranges '[...]' [--allowCachedFormulaValues]   # or
  --sourceSpreadsheet <id> --ranges '[...]'
  --output <out.xlsx> | --spreadsheetId <id> [--workbookTemplate <f.xlsx>]
  [--dryRun] [--overwrite] [--overwriteFormulas] [-c/-p/-f auth flags]
```

- `-s` = spreadsheet ID (from URL `docs.google.com/spreadsheets/d/<ID>/edit`), `-t` = worksheet tab title (exact, case-sensitive — quote it).
- All data commands accept: positional JSON array-of-arrays, `-i <file>` (format inferred from extension), or `-i -` (stdin; pipe CSV → add `--inputFormat csv`).
- Table output supports oclif ux flags: `--csv`, `--columns`, `--sort`, `--filter`. `--rawOutput`/`-r` prints JSON.
- Writes are formula-safe by default: existing formula cells are NOT overwritten without explicit `--overwriteFormulas`; batch/update writes support `--dryRun` preview with zero side effects.
- `spreadsheet:share|permissions|unshare` use the Drive API (`drive.file` scope). OAuth tokens issued before that scope was added must re-run `auth:login`. Under `drive.file` only files this app created or has opened are shareable — a pre-existing sheet may need one `spreadsheet:get` through this app first. `--notify` defaults OFF.
- `data:find` returns `{matchCount, truncated, matches:[{a1,row,column,columnLetter,value,rowValues?}]}`. Use it to locate a row/cell before a targeted `data:update`. `--header "Name"` resolves a column by its header text (first scanned row); `--byRow` returns full row values. Zero matches → `matchCount:0`, exit 0.
- `format:cells`/`format:merge` only touch `userEnteredFormat` — values and formulas are never overwritten. `--dryRun` prints the exact batchUpdate request bodies. For a pretty report: `format:cells --range A1:J1 --bold --backgroundColor "#1a73e8" --textColor "#ffffff"` then `format:cells --range A2:J50 --numberFormat "#,##0.00" --borders all`.
- `grid:*` commands mutate structure, not values: `insert`/`delete` shift cells (delete `--dryRun` shows the values about to be lost), `hide`/`resize`/`freeze` are non-destructive. All are 1-based `--start`/`--count`; `--dryRun` prints the exact batchUpdate request.

## Workflow for "read a sheet URL"

1. Extract ID from URL.
2. `spreadsheet:get -s <ID> --rawOutput` → worksheet titles (`| jq '.sheets[].properties.title'`).
3. `data:get -s <ID> -t "<title>" --minRow 1 --maxRow N` → read rows. Raw numbers → `--valueRenderOption UNFORMATTED_VALUE`.

## Workflow for reports

Ready-made templates + data live in `examples/reports/` in the repository (finance-template.json, manpower-template.json, sales-report-template.json + CSVs):

```sh
# Finance cash-flow: FX conversion, refunds, opening/closing reconciliation (cash, NOT profit)
report:run --template examples/reports/finance-template.json \
  --input examples/reports/finance-transactions.csv --output out.xlsx

# Manpower: quantity x unit-effort x complexity, rates, contingency (effort in person-days, NOT duration)
report:run --template examples/reports/manpower-template.json \
  --input examples/reports/manpower-tasks.csv --output out.xlsx
```

Templates declare: schema fields (`string|number|integer|decimal|boolean|date` + `required/min/max/enum/dateFormat`), policy (`strictTypes`, `treatMissingAsZero`, `duplicateKeyPolicy`, `maxInputRows`), transforms (`computed`, `filter`, `sort`, `aggregate`, `pivot`), outputs (sheet name, columns, freeze, widths, formats). Templates never eval JavaScript. Financial math is exact decimal (no float drift).

Safety invariants: reports to Google Sheets/XLSX mark managed ranges and re-runs replace them cleanly (idempotent); string data starting with `=`,`+`,`-`,`@` is escaped (formula-injection guard); formula cache values from source workbooks are only consumed with explicit `--allowCachedFormulaValues` (nothing recalculates formulas).

## Errors

- `403 The caller does not have permission` → sheet not shared with the authenticated identity. Share it or login with the owning account.
- `Sheet "X" not found` → wrong tab title; list via `spreadsheet:get --rawOutput`.
- `No OAuth tokens found` → run `auth:login`.
- `Client secret file not found` → place `client_secret.json` in `~/.config/google-sheet-cli/` or pass `--clientSecretFile`.
- `Cannot overwrite unowned populated cell` / formula-collision error → target range has data/formulas outside report ownership; add `--overwrite` (and `--overwriteFormulas` for formulas) only intentionally.
