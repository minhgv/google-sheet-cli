---
name: google-sheet
description: Read, write, and manage Google Sheets AND local Excel (.xlsx) files from the terminal using the google-sheet-cli oclif CLI. Supports OAuth 2.0 user auth and Service Account JWT for cloud, plus fully-offline XLSX inspect/read/write, declarative report automation (finance cash-flow, manpower estimation) from templates, cell formatting (borders, bold, colors, number formats, merge), and cell search returning A1 coordinates. Use when the user asks to read/update a Google Sheet, batch-read ranges, append table rows, create spreadsheets, work with a docs.google.com/spreadsheets URL, inspect or edit a local .xlsx, generate a report from CSV/JSON via a template, find a cell/row by value, or format a sheet for a pretty report, upsert rows keyed by a column, discover or validate a sheet's schema, duplicate a spreadsheet or worksheet, or export a spreadsheet to PDF/XLSX. Trigger on "đọc sheet", "google sheet", "spreadsheet", "xlsx", "excel", "report", "báo cáo", "data:get", "data:find", "data:upsert", "data:validate", "format:cells", "gsheet", "spreadsheet:export", "data:export-csv", "spreadsheet:list", "capabilities", "export csv", "list spreadsheets", or a Google Sheets URL.
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
- Config dir override: `GSHEET_CONFIG_DIR` relocates the whole `~/.config/google-sheet-cli` directory (token.json + client_secret.json) — for tests and sandboxed runs; never point it at a shared path. Tokens are written atomically at mode 0600, dir 0700.
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
data:schema       -s <id> -t <title> [--minRow 1 --minCol 1 --maxRow 100 --maxCol 26]
                  [table flags]   read-only schema discovery; first sampled row is the header
data:validate     -s <id> -t <title> (--schema '<TableSchema JSON>'|--schemaFile <path|->)
                  [--minRow --minCol --maxRow --maxCol]   read-only; exit 1 on violations
data:clear        -s <id> -t <title> --range 'Sheet1!A2:D20' [--dryRun] [--overwriteFormulas]
                  values-only clear of a bounded range; formulas refuse without --overwriteFormulas
data:upsert       -s <id> -t <title> --key <column> (-d '<json rows>'|-i <file|->) [--inputFormat json|csv]
                  [--range 'Sheet1!A1:F100'] [--valueInputOption] [--dryRun] [--overwriteFormulas]
data:export-csv   (-s <id> -t <title> [--range A1:B10] | --workbook <f.xlsx> --range 'Sheet1!A1:D10')
                  [--mode raw|formatted|formula] [--injection safe|preserve] [-o out.csv] [--overwrite]
                  RFC 4180 CSV from either backend; -o = atomic write + receipt, omit = bytes on stdout
spreadsheet:copy  -s <id> [--title <name>]   Drive copy; source untouched; sharing NOT cloned
spreadsheet:export -s <id> --format pdf|xlsx -o <file> [--overwrite]   atomic write; Drive 10 MB cap
worksheet:copy    -s <id> -t <title> --destinationSpreadsheetId <destId>   server assigns title on collision
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
spreadsheet:list  [--name <fragment>] [--exact] [--pageSize N] [--pageToken <t>|--all] [--rawOutput]
                  drive.file scope: app-created/opened only; empty ≠ absent; never auto-selects an id
worksheet:add|get|remove -s <id> -t <title>
worksheet:rename  -s <id> -t <old> -n <new>

# Local XLSX (offline, no auth, no network)
workbook:inspect  -f <file.xlsx> [--includeFormulaCells]
                  per-sheet merged ranges, formula counts, frozen panes, validations, cond-formatting
workbook:read     -f <file.xlsx> --range 'Sheet1!A1:E20' [--mode unformatted|formatted|formula]
workbook:find     -f <file.xlsx> [-t <sheet>] (--equals|--contains|--regex <v>) [--range A1:Z100]
                  [--column B | --header "Status"] [--ignoreCase|--no-ignoreCase]
                  [--limit N|--first] [--byRow] → same match JSON shape as data:find
workbook:write    [-f template.xlsx] [-o out.xlsx] -i <data.csv|json|-> [-t Sheet1] [--startCell A1]
                  [--inPlace] [--dryRun] [--overwrite] [--overwriteFormulas] [--discardUnsupported] [-r]
                  (or pass data as positional JSON 2D array / ReportDocument)
                  --cells '{"B5":"x","J8":"GD_WEB1"}' sparse per-cell write; every other cell untouched

# Dual-backend mutations: --workbook <f.xlsx> instead of -s <id>; -t selects the sheet inside
# the file; result lands on --output <path> or --inPlace (.bak backup), never implicitly on the
# source; --dryRun previews without saving; --discardUnsupported consents to dropping
# charts/pivots/macros the engine cannot preserve
grid:insert       --workbook <f.xlsx> -t <sheet> --dimension ROWS|COLUMNS --start N [--count N]
                  [--inheritFromBefore] [--force] (--inPlace|-o <out.xlsx>|--dryRun)
grid:delete       --workbook <f.xlsx> -t <sheet> --dimension ROWS|COLUMNS --start N [--count N]
                  [--force] (--inPlace|-o <out.xlsx>|--dryRun)   dryRun lists removed values
format:merge      --workbook <f.xlsx> -t <sheet> --range A1:J1
                  [--type MERGE_ALL|MERGE_COLUMNS|MERGE_ROWS|--unmerge] (--inPlace|-o|--dryRun)
data:clear        --workbook <f.xlsx> --range 'Sheet1!A2:D20' [--overwriteFormulas]
                  (--inPlace|-o <out.xlsx>|--dryRun)

# Reports (template-driven)
report:run --template <template.json> \
  --input <data.csv|json|-> [--inputFormat]   # exactly ONE source: this, or
  --sourceWorkbook <f.xlsx> --ranges '[...]' [--allowCachedFormulaValues]   # or
  --sourceSpreadsheet <id> --ranges '[...]'
  --output <out.xlsx> | --spreadsheetId <id> [--workbookTemplate <f.xlsx>]
  [--dryRun] [--overwrite] [--overwriteFormulas] [-c/-p/-f auth flags]

# Meta (no auth, no flags)
capabilities      print the machine-readable capability document: per-backend operations,
                  I/O forms, formula write vs recalculation, guards, mutation limits,
                  fidelity exclusions, drive.file visibility; listed ≠ authorized
```
- Local structural edits: `grid:insert`/`grid:delete` on `--workbook` shift cells and manage merges explicitly — a merge intersecting the splice boundary refuses the operation unless `--force` (insert extends it, delete shrinks or drops it). Formula references are NEVER rewritten: the receipt reports `formulasAtRisk` so callers verify them; data validations and conditional formatting ranges are not adjusted. `format:merge` on `--workbook` decomposes MERGE_COLUMNS/MERGE_ROWS into per-line merges. `workbook:write --cells` writes only the listed addresses — the primitive for "fill columns A,B,C,D,J, skip formula columns" template work.

- `-s` = spreadsheet ID (from URL `docs.google.com/spreadsheets/d/<ID>/edit`), `-t` = worksheet tab title (exact, case-sensitive — quote it).
- All data commands accept: positional JSON array-of-arrays, `-i <file>` (format inferred from extension), or `-i -` (stdin; pipe CSV → add `--inputFormat csv`).
- Table output supports oclif ux flags: `--csv`, `--columns`, `--sort`, `--filter`. `--rawOutput`/`-r` prints JSON.
- Writes are formula-safe by default: existing formula cells are NOT overwritten without explicit `--overwriteFormulas`; batch/update writes support `--dryRun` preview with zero side effects.
- `spreadsheet:share|permissions|unshare` use the Drive API (`drive.file` scope). OAuth tokens issued before that scope was added must re-run `auth:login`. Under `drive.file` only files this app created or has opened are shareable — a pre-existing sheet may need one `spreadsheet:get` through this app first. `--notify` defaults OFF.
- `data:find` returns `{matchCount, truncated, matches:[{a1,row,column,columnLetter,value,rowValues?}]}`. Use it to locate a row/cell before a targeted `data:update`. `--header "Name"` resolves a column by its header text (first scanned row); `--byRow` returns full row values. Zero matches → `matchCount:0`, exit 0.
- Structured errors for agents: any command with `--json`/`-j` prints ONE failure envelope on stderr and exits 1 — `{"error":{"code":"NOT_FOUND","message":"...","retryable":false,"retryAfterMs"?,"issues"?}}` — success output unchanged. `--rawOutput`/`-r` gives JSON success AND failure. Codes: `USAGE` `AUTH_REQUIRED` `UNAUTHORIZED` `FORBIDDEN` `NOT_FOUND` `CONFLICT` `RATE_LIMITED` `UPSTREAM` `NETWORK` (the retryable set) `REQUEST_INVALID` `VALIDATION` `SCHEMA_INVALID` `DATA_INVALID` `INTERNAL`. `issues` carry `{row, column, a1, field, code, message, value}` for `data:validate` violations and upsert rejections. Retry only `retryable:true` codes; never replay an ambiguous write blindly.
- Multi-request writes (`data:batch-update`, `data:upsert`) may carry `error.mutation` in the failure envelope: per-request outcomes `acknowledged`/`rejected`/`unknown` (dispatch may have succeeded) /`not-attempted` with logical A1 ranges, `gridGrowth` tracked separately, the failing `phase`, and conservative `retryGuidance` (`safe-replay`/`verify-then-replay`/`never-blind-replay`). Honor the guidance — `unknown` means the write may already be applied.
- `--redacted` (env `GSHEET_REDACTED`) strips cell contents, formulas, incoming values and credentials from error envelopes and dry-run diagnostics before serialization — coordinates, counts, statuses and mutation states are kept. Use it when piping failures into logs or transcripts.
- `data:export-csv` emits deterministic RFC 4180 CSV: `--injection safe` (default) prefixes dangerous leading characters (`= + - @` tab CR) with `'`, `--injection preserve` is byte-faithful and says so in the receipt. Typed negative numbers are never prefixed. With `-o` the write is atomic (0600) and refuses to clobber without `--overwrite`; without `-o` stdout carries only CSV bytes.
- `data:schema` samples rows 1-100 x columns A-Z by default (`--minRow/--minCol/--maxRow/--maxCol`; first sampled row is the header) and reports header A1 cells, inferred types with counts, formula-cell counts, data-validation rules and named ranges. Types are sample evidence — Google stores dates as numbers — not an authoritative schema.
- `data:validate` takes the report TableSchema shape (`{"fields":[{"name":"Name","type":"string","required":true}]}`; types `string|number|integer|decimal|boolean|date` plus `required/min/max/enum/unique`). Violations arrive with exact row/A1 coordinates and error code `DATA_INVALID`; a malformed schema fails with `SCHEMA_INVALID`. Both commands are strictly read-only.
- `data:upsert` is single-writer, keyed by one `--key` column with typed matching (string `"001"` never matches number `1`; leading zeros kept). Header-first input; omitted columns and formulas are preserved; writes are `RAW` by default; it refuses when a bounded `--range` hides trailing table data; `--dryRun` previews added/updated/unchanged rows plus planned ranges. No transaction, no automatic retry, no concurrent writers.
- `spreadsheet:copy` / `worksheet:copy` never touch the source; sharing grants are NOT duplicated (the clone is private to this app); on a title collision `worksheet:copy` reports the server-assigned title and the new sheet id.
- `spreadsheet:export` writes PDF/XLSX atomically (temp file + rename), refuses to clobber an existing file without `--overwrite`, and is capped at 10 MB by Drive; authorization stays in the `drive.file` scope (files this app created or opened), never full Drive.
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

Machine-readable mode: run any command with `--json`/`-j` (or `--rawOutput`/`-r`) → one JSON envelope on stderr, exit 1. React by `code`: `AUTH_REQUIRED`/`UNAUTHORIZED` → re-auth, `FORBIDDEN` → share the sheet, `RATE_LIMITED` → wait `retryAfterMs`, `USAGE`/`VALIDATION`/`SCHEMA_INVALID` → fix the invocation, `DATA_INVALID` → inspect `issues`, `NOT_FOUND` → wrong id/title.

- `403 The caller does not have permission` → sheet not shared with the authenticated identity. Share it or login with the owning account.
- `Sheet "X" not found` → wrong tab title; list via `spreadsheet:get --rawOutput`.
- `No OAuth tokens found` → run `auth:login`.
- `Client secret file not found` → place `client_secret.json` in `~/.config/google-sheet-cli/` or pass `--clientSecretFile`.
- `Cannot overwrite unowned populated cell` / formula-collision error → target range has data/formulas outside report ownership; add `--overwrite` (and `--overwriteFormulas` for formulas) only intentionally.
