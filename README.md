# google-sheet-cli

A simple helper cli to interact with google sheets.

## Features

The fork provides **40 commands**. See [CHANGELOG.md](CHANGELOG.md) for the latest additions; the 2.x migration notes below describe the upstream platform release, not the fork's later feature set.

| Group | Commands | What it does |
|---|---|---|
| 🔐 Auth | `auth:login` `auth:logout` `auth:status` | OAuth 2.0 Desktop flow; token at `~/.config/google-sheet-cli/token.json` |
| 📊 Data | `data:get` `data:append` `data:append-table` `data:update` `data:batch-get` `data:batch-update` `data:export-csv` `data:find` `data:schema` `data:validate` `data:clear` `data:upsert` | Read/write ranges; batch APIs; `--dryRun`, `--overwriteFormulas` guard, `--valueRenderOption`, file/stdin input; `data:find` locates cells by condition; `data:schema` discovers column schema (read-only); `data:validate` checks data against a TableSchema; `data:clear` clears bounded values; `data:upsert` updates-or-appends rows by key column; `data:export-csv` exports RFC 4180 CSV from a Sheets worksheet or a local `.xlsx` range (`--mode raw\|formatted\|formula`, `--injection safe\|preserve`) |
| 📄 Spreadsheet | `spreadsheet:add` `spreadsheet:get` `spreadsheet:list` `spreadsheet:share` `spreadsheet:permissions` `spreadsheet:unshare` `spreadsheet:copy` `spreadsheet:export` | Create/inspect spreadsheets; Drive-API sharing (`--email`/`--domain`/`--anyone`, `--role`, `--notify` off by default); `spreadsheet:list` discovers Drive-visible spreadsheets (`drive.file` scope — an empty result never proves a file is absent); `spreadsheet:copy` duplicates a spreadsheet; `spreadsheet:export` writes PDF/XLSX to a local file |
| 📑 Worksheet | `worksheet:add` `worksheet:get` `worksheet:remove` `worksheet:rename` `worksheet:copy` | Manage worksheets inside a spreadsheet; `worksheet:copy` copies a sheet into another spreadsheet |
| 🎨 Format | `format:cells` `format:merge` | `userEnteredFormat` only — bold, colors, borders, number formats, merge; `--dryRun` |
| 📐 Grid | `grid:insert` `grid:delete` `grid:hide` `grid:resize` `grid:freeze` | Structural row/column mutations; `grid:delete --dryRun` previews values about to be lost |
| 📁 Workbook | `workbook:inspect` `workbook:read` `workbook:write` | Offline local .xlsx — no credentials, no network; atomic write + SHA-256 |
| 📋 Report | `report:run` | Render a JSON report spec onto a sheet (finance/manpower/sales) |
| 🧭 Meta | `capabilities` | Print the machine-readable capability document — per-backend operations, I/O forms, formula write-vs-recalculation semantics, destructive guards, mutation limits, fidelity exclusions — with no credentials and no flags |

## Migrating from 2.x

3.0.0 is a platform release. Every command, every flag and everything the commands print is the same as on 2.3.0, and the sequence of Sheets API calls each library method makes is unchanged — same requests, same order, same bodies. What moved is the Node floor, the module a library consumer imports, the type of the errors that are thrown, and one hostname the auth stack talks to.

Coming from 2.2.x rather than 2.3.0? 3.0.0 carries everything in [Changes in 2.3.0](#changes-in-230) as well, so read both. Two of those are outright behaviour changes and are repeated [at the end of this section](#two-22x-calls-that-changed-in-230).

Using the `jroehl/gsheet.action` GitHub Action rather than this package directly? It bundles a pinned copy of `google-sheet-cli` into its committed `dist/`, so none of this reaches a workflow until that action bumps the dependency and publishes.

### Node 22 or newer

`engines.node` is now `>=22`, so `google-sheet-cli@latest` needs Node 22 or newer. Node 14 through 20 are no longer supported.

The 2.x line stays on npm and keeps its `>=14` floor. Ask for it by range, which always resolves to the newest 2.x:

```sh-session
$ npm install -g google-sheet-cli@^2
$ npx google-sheet-cli@^2 spreadsheet:get -s <spreadsheetId>
```

The `2.x` branch carries that line and gets fixes, not features.

### The bin scripts are `bin/run.js` and `bin/dev.js`

oclif 5 expects its executables to carry a file extension, so `bin/run` became `bin/run.js` and `bin/dev` became `bin/dev.js`. This only matters where something names the file by path — a checkout, a container image, a script calling `node_modules/google-sheet-cli/bin/run`. The installed `google-sheet` binary and `npx google-sheet-cli` are unaffected.

### A side-effect-free `google-sheet-cli/sheet` subpath

Using the library without the cli used to mean reaching into the build output:

```ts
import GoogleSheet from 'google-sheet-cli/lib/lib/google-sheet';
```

That path still works and still carries its types. 3.0.0 adds a declared one:

```ts
import GoogleSheet from 'google-sheet-cli/sheet';
```

The package root (`google-sheet-cli`) also exports oclif's `run`, so importing it loads the whole cli; `google-sheet-cli/sheet` loads the Sheets client and its auth chain and nothing else — 128 modules against 289 for the root, with no `@oclif/core` anywhere in the graph.

The subpath is declared through the package's `exports` map, so a TypeScript consumer has to be on `moduleResolution` `node16`, `nodenext` or `bundler`. Under the older `node10` resolution `google-sheet-cli/sheet` does not resolve at all (`TS2307: Cannot find module`), and only the deep path does. If moving that setting is not an option, keep the deep import.

### The `exports` map seals paths that used to be reachable

Declaring `exports` at all closes off everything it does not name. `google-sheet-cli`, `google-sheet-cli/sheet`, `google-sheet-cli/lib/*` (with or without `.js`) and `google-sheet-cli/package.json` resolve; anything else now fails with `ERR_PACKAGE_PATH_NOT_EXPORTED`. The one path known to be closed this way is `google-sheet-cli/oclif.manifest.json`. Nothing in this project or its GitHub action reads it, but a consumer that did will have to stop.

Separately, `google-sheet-cli/lib/lib/types` is gone (`MODULE_NOT_FOUND`). It was a second, stale copy of the `GoogleSheetCli` type namespace that nothing imported — the live one has always been declared in `google-sheet.ts` and is re-exported from the package root and from `google-sheet-cli/sheet`.

### Errors are `Error` instances

Eight places used to `throw` a bare string. They now throw an `Error` carrying byte-identical text:

- `Spreadsheet "<id>" not found`
- `Sheet "<title>" not found in "<spreadsheet>"`
- `Option property "worksheetTitle" is required`
- `No header row exists`
- `Specify worksheetTitle`
- `Check "data" property - has to be supplied as nested array ([["1", "2"], ["3", "4"]])`
- `col has to be greater than 1`
- `Label has to be uppercase alphabet letter but is "<label>"`

The cli prints exactly what it printed before, and so does the GitHub action, which already read `err.message || err`. What breaks is a library consumer that compares the caught value to a string:

```ts
catch (err) {
  if (err === 'Specify worksheetTitle') { … }                     // 2.x, no longer matches
  if ((err as Error).message === 'Specify worksheetTitle') { … }  // 3.x
}
```

### The OAuth token endpoint and the scope moved

3.0.0 builds its Sheets client from `@googleapis/sheets` instead of the whole `googleapis` bundle. That brings a newer auth stack with it, and two things on the wire follow from it.

The token request goes to a different host, and the assertion asks for a different scope:

```
2.x   POST https://www.googleapis.com/oauth2/v4/token   scope https://spreadsheets.google.com/feeds/
3.x   POST https://oauth2.googleapis.com/token          scope https://www.googleapis.com/auth/spreadsheets
```

**If you run this behind an egress allowlist that names `www.googleapis.com`, add `oauth2.googleapis.com`.** `sheets.googleapis.com`, where every actual Sheets call goes, is unchanged, and so is every one of those calls.

The scope change needs no action for the normal setup: a service account that was shared onto a spreadsheet the way [Step 2](#step-2-sharing-the-spreadsheet) describes keeps working, because the sharing is what grants the access and the account authorises itself for the scope. **The exception is Google Workspace domain-wide delegation.** If the service account is used through delegation, an administrator has to add `https://www.googleapis.com/auth/spreadsheets` to that client's allowed scopes in the Admin console, beside the old feeds scope. Until they do, every call comes back 401.

### Request headers changed

Same upgrade, same cause — the HTTP client underneath went from gaxios 5 to gaxios 7. Google ignores all of this; it is listed because it is visible to a proxy, a request log or a strict middlebox.

On every Sheets call:

```
2.x   Accept: application/json
3.x   Accept: */*
```

That is the header getting looser, not stricter, so nothing starts refusing to answer. What it can break is something in the middle that was matching on `application/json` — a proxy rule, a request filter, a recorded-cassette test fixture.

On the token request:

```
2.x   Content-Type: application/x-www-form-urlencoded
      Accept-Encoding: gzip,deflate
3.x   Content-Type: application/x-www-form-urlencoded;charset=UTF-8
      Accept-Encoding: gzip, deflate, br
```

The `User-Agent` and `x-goog-api-client` version strings move with the client version, as they do on any upgrade.

### `help`'s argument is rendered differently

`@oclif/plugin-help` 7 declares the argument as a variadic, so the usage line and `docs/help.md` move from

```
$ google-sheet help [COMMANDS] [-n]
```

to

```
$ google-sheet help [COMMAND...] [-n]
```

Nothing about invoking it changes: `google-sheet help data:get` works exactly as before. The strings are named here because they are in the generated docs and in anyone's screenshots.

The same plugin also lays every command's `--help` out slightly differently: a flag that reads an environment variable is annotated `[env: NAME]`, a flag with no short character is indented under the ones that have one, and the dim styling on descriptions is gone. Every flag itself — long name, short character, default, `=<value>` shape — is unchanged. If something parses `--help` output, this is the release that will break it.

### `js-yaml` is pinned to 3.x on purpose

`data:get --output=yaml` is rendered by a copy of `@oclif/core@2.8.11`'s table, carried in `src/lib/table.ts` because core 5 has no `ux.table` and no successor that keeps the eight table flags. That code calls `safeDump`, which js-yaml 4 renamed to `dump`. The pin exists so the yaml output stays byte-for-byte what 2.2.x emitted, which is the property the whole vendored table was verified against. js-yaml 3.14.1 is end of life; this is a deliberate compatibility pin, not neglect, and moving it means re-running that output comparison, not just changing the version.

### `getData` without `minCol` returns instead of throwing

On 2.x, `getData({ worksheetTitle: 'Sheet1' })` threw `col has to be greater than 1`, and so did a whole-worksheet quoted range, `getData({ range: "'Sheet1'!" })`. The read itself was correct — it started at A1 — but the code that names unlabelled columns counted from column 0, and `colToA` refuses anything below 1. The cli never reached it, because `data:get` defaults `--minCol` to 1; a library caller and the GitHub action did.

Those calls now return what the same call with an explicit `minCol: 1` returns, field for field: the identical request was already being sent, so the identical result is the only defensible answer.

**Calls that already returned something return exactly what they returned before, wrong labels included.** With `hasHeaderRow` and a non-empty first heading, 2.x never reached `colToA(0)` at all — it just numbered the blank headings one column short, calling column C `(B)` and column D `(C)`, and emitting one heading too many when the read came back with no rows. Those labels are the keys of every `formatted` row, and the GitHub action serialises them into its `results`, so a workflow may be reading them today. 3.0.0 keeps them. Correcting them is a change to output that currently works and is left for a release that announces it.

Measured against published 2.2.0 through the library's own HTTP fake, over 125 shapes — six worksheet fixtures (regular, ragged header, header-only, blank first heading, empty, header wider than its data) crossed with `hasHeaderRow`, `minCol` absent/0/1/2 and `minRow` absent/2, plus `maxCol`/`maxRow` variants, the four range forms, and six chained-call scenarios on one shared instance: **83 identical, 0 different, 42 previously throwing**.

### Three 2.2.x calls that changed in 2.3.0

These arrived in 2.3.0, not in 3.0.0, so they are new only to someone upgrading from 2.2.x. All three were found by running the published 2.2.0 build and this one side by side through the library's HTTP fake; they are the only differences.

**A range that names no worksheet, together with a `worksheetTitle` that does not exist, now fails.** `updateData(data, { worksheetTitle: 'Ghost', range: 'A1:B1' })` handed `A1:B1` to the API on 2.2.0, which resolved it against the first sheet and wrote there. Now the grid-sizing step looks the worksheet up first, does not find it, and throws. It needs both halves — a range carrying no worksheet, and a title naming a sheet that is gone. The cli cannot reach it (no write command exposes a `range` flag); only the library and the GitHub action's `range` option can.

**`appendData` no longer writes the range's worksheet back onto the options object you passed.** With a quoted range and an explicit title, 2.2.0 left your `worksheetTitle` mutated to the range's worksheet; it is now left as you passed it. The data lands in the same cells either way. It is visible through the GitHub action, which serialises `command.kwargs` into its `results` output, so a workflow reading `kwargs[1].worksheetTitle` after such a call sees a different value.

**An anchor range with no end is rejected before the request rather than by the API.** `getData({ range: "'Sheet1'!A2:" })` fails on 2.2.x and it fails now; what changed is where. 2.2.0 sent the request and surfaced the API's `Unable to parse range:`, and the call now fails locally with `Invalid range "'Sheet1'!A2:"` and sends nothing. Only the message and the request count differ, so this matters if something is matching on the old text.

### What has not changed

- Every command, flag, short character, default and argument. `data:get`'s eight table flags (`--columns`, `--sort`, `--filter`, `--csv`, `--output`, `-x/--extended`, `--no-truncate`, `--no-header`) all survive, rendering the same table.
- The Sheets requests each library method makes, down to the query string and the body.
- Authentication itself: service accounts, the same three ways of handing over credentials, the same `GSHEET_*` environment variables.
- A `range` that names a different worksheet than `worksheetTitle` still warns on stderr and writes to the worksheet the range names, exactly as 2.2.x and 2.3.0 do. 3.0.0 deliberately does not turn that warning into a refusal.

## Changes in 2.3.0

- The `engines.node` floor is now `>=14`, which was always the real minimum.
- Unusable service account credentials are rejected before the first request, with a message that names the fix instead of an OpenSSL parser error.
- The new `--credentialsFile` flag reads the credentials straight out of the service account JSON file. See [Credentials](#credentials).
- Writing past the last row or column of a worksheet now grows the grid first instead of failing with "exceeds grid limits", so appending to a sheet that is already full works again.
- A `range` naming a different worksheet than `worksheetTitle` still writes to the worksheet the range names, exactly as 2.2.x did, but now says so on stderr instead of resolving the contradiction silently. It is still a warning on 3.0.0 and there is no plan to make it a refusal in that release: refusing would break a call every 2.x version completed, which is a different kind of change from the platform moves 3.0.0 is made of.
- `updateData` now accepts a `range` carrying a quoted worksheet title with no `worksheetTitle` beside it, taking the worksheet from the range instead of insisting on a title the range already named. As in 2.2.x, an unquoted title inside a range does not name the worksheet.
- A write now re-points the worksheet a `GoogleSheet` instance remembers. `updateData` resolves and fetches its target worksheet before writing, and that worksheet becomes the one a later command uses when it omits `worksheetTitle`. Before 2.3.0 only a read moved it. This is only visible when several commands share one instance, which is what the GitHub action does.
- A write that sizes a grid costs one extra API read. Growing the grid means knowing how big it is, so `updateData` fetches the spreadsheet before the update, plus one more request when the grid actually has to grow. The one write that reads nothing extra is an unquoted `range` naming a worksheet other than the one the call resolves to, where nothing is sized because nothing there is written to.

[![oclif](https://img.shields.io/badge/cli-oclif-brightgreen.svg)](https://oclif.io)
[![Version](https://img.shields.io/npm/v/google-sheet-cli.svg)](https://npmjs.org/package/google-sheet-cli)
[![Downloads/week](https://img.shields.io/npm/dw/google-sheet-cli.svg)](https://npmjs.org/package/google-sheet-cli)
[![License](https://img.shields.io/npm/l/google-sheet-cli.svg)](https://github.com/jroehl/google-sheet-cli/blob/master/package.json)

[![CodeQL](https://github.com/jroehl/google-sheet-cli/actions/workflows/codeql-analysis.yml/badge.svg)](https://github.com/jroehl/google-sheet-cli/actions/workflows/codeql-analysis.yml)
[![Test version](https://github.com/jroehl/google-sheet-cli/actions/workflows/test.yml/badge.svg)](https://github.com/jroehl/google-sheet-cli/actions/workflows/test.yml)

- [google-sheet-cli](#google-sheet-cli)
  - [Migrating from 2.x](#migrating-from-2x)
  - [Changes in 2.3.0](#changes-in-230)
  - [Usage as CLI](#usage-as-cli)
  - [Usage as library](#usage-as-library)
    - [Google Sheets Client (`google-sheet-cli/sheet`)](#google-sheets-client-google-sheet-clisheet)
    - [Local Excel Workbook Engine (`google-sheet-cli/xlsx`)](#local-excel-workbook-engine-google-sheet-clixlsx)
    - [Report Core Engine & Transforms (`google-sheet-cli/report`)](#report-core-engine--transforms-google-sheet-clireport)
  - [Data Operations & Automation Guide](#data-operations--automation-guide)
    - [Data Input: Positional JSON, CSV/JSON Files, and Stdin](#data-input-positional-json-csvjson-files-and-stdin)
    - [Read Render Options](#read-render-options)
    - [Batch Operations & Native Append Table vs Legacy Semantics](#batch-operations--native-append-table-vs-legacy-semantics)
    - [Local Excel (XLSX) Workbooks: Inspect, Read, and Safe Write](#local-excel-xlsx-workbooks-inspect-read-and-safe-write)
    - [Report Automation (`report:run`)](#report-automation-reportrun)
    - [Authentication Behavior: Cloud vs Local](#authentication-behavior-cloud-vs-local)
    - [Report Templates: Schemas, Policies, and Transforms](#report-templates-schemas-policies-and-transforms)
    - [Technical Invariants, Safety Guarantees, and Domain Policies](#technical-invariants-safety-guarantees-and-domain-policies)
  - [Using with coding agents](#using-with-coding-agents)
- [Command Topics](#command-topics)
- [Info](#info)
  - [How to configure](#how-to-configure)
    - [Step 1: Setting Up Google Service Account](#step-1-setting-up-google-service-account)
    - [Step 2: Sharing the Spreadsheet](#step-2-sharing-the-spreadsheet)
  - [Credentials](#credentials)
  - [Build with](#build-with)
  - [Contributing](#contributing)
  - [Versioning](#versioning)
  - [License](#license)
  - [TODO](#todo)
## Usage as CLI
<!-- usage -->
```sh-session
$ npm install -g google-sheet-cli
$ google-sheet COMMAND
running command...
$ google-sheet (--version)
google-sheet-cli/0.0.0 darwin-arm64 node-v25.9.0
$ google-sheet --help [COMMAND]
USAGE
  $ google-sheet COMMAND
...
```
<!-- usagestop -->

## Usage as library

The package exposes focused subpaths through its `exports` map, allowing you to import exactly what you need without pulling unnecessary runtime dependencies into your bundle.

### Google Sheets Client (`google-sheet-cli/sheet`)

A side-effect-free client for interacting directly with the Google Sheets v4 API. It contains zero CLI framework (`@oclif/core`) or Excel workbook (`exceljs`) dependencies:

```ts
import GoogleSheet from 'google-sheet-cli/sheet';

const sheet = new GoogleSheet({
  spreadsheetId: '<spreadsheetId>',
  clientEmail: process.env.GSHEET_CLIENT_EMAIL,
  privateKey: process.env.GSHEET_PRIVATE_KEY,
});

// Read rows with structured rendering
const { rawData, formatted, header, range } = await sheet.getData({
  worksheetTitle: 'Sheet1',
  minRow: 1,
  minCol: 1,
  valueRenderOption: 'UNFORMATTED_VALUE',
});

// Batch read multiple disjoint ranges in one API call
const batchResults = await sheet.getDataBatch(
  ['Sheet1!A1:B10', 'Sheet2!C1:D5'],
  { valueRenderOption: 'FORMATTED_VALUE' }
);

// Append data using native Google Sheets table append
await sheet.appendTableData(
  [['2026-03-15', 'Hosting', 120.00]],
  { worksheetTitle: 'Transactions' }
);
```

### Local Excel Workbook Engine (`google-sheet-cli/xlsx`)

An ExcelJS-backed adapter providing local `.xlsx` workbook manipulation, structural preflight inspection, formula freshness tracking, range reads/writes, diff generation, and safe atomic persistence:

```ts
import { XlsxWorkbook } from 'google-sheet-cli/xlsx';
// 1. Load existing workbook or create a new empty one
const workbook = await XlsxWorkbook.load('./templates/financial-model.xlsx');

// 2. Inspect workbook structure, sheets, defined names, and capabilities
const preflight = workbook.inspect();
console.log('Sheets found:', preflight.sheets.map(s => s.name));
console.log('Can evaluate formulas:', preflight.capabilities.canRecalculateFormulas);

// 3. Read specific range with formula awareness
const rangeData = workbook.read('Summary!A1:D10', {
  mode: 'unformatted',
});
console.log('Cell values:', rangeData.values);
console.log('Formula freshness:', rangeData.freshnessSummary);

// 4. Preview diff before writing (or apply directly)
const diff = workbook.preview(reportDoc, { overwriteFormulas: false });
console.log(`Planned cell changes: ${diff.changes.length}, conflicts: ${diff.conflicts.length}`);

const result = workbook.apply(reportDoc, { overwriteFormulas: false });
console.log(`Applied changes: ${result.appliedChanges}`);

// 5. Save changes safely with atomic write and automatic backup
await workbook.save('./output/financial-model-q1.xlsx', {
  overwrite: true,
  backup: true,
});
```

### Report Core Engine & Transforms (`google-sheet-cli/report`)

A pure, synchronous report transformation engine that validates schemas, executes declarative pipelines, and generates deterministic report models:

```ts
import { buildReport } from 'google-sheet-cli/report';

const template = {
  id: 'sales-summary-v1',
  version: 1,
  kind: 'table' as const,
  title: 'Quarterly Sales Summary',
  schema: {
    fields: [
      { name: 'region', type: 'string' as const, required: true },
      { name: 'revenue', type: 'decimal' as const, required: true, min: 0 },
    ],
  },
  transforms: [
    {
      kind: 'aggregate' as const,
      groupBy: ['region'],
      aggregations: [{ field: 'revenue', op: 'sum' as const, as: 'total_revenue' }],
    },
  ],
  outputs: [
    {
      name: 'Regional Summary',
      columns: [
        { header: 'Region', field: 'region' },
        { header: 'Total Revenue', field: 'total_revenue' },
      ],
    },
  ],
};

const rawData = [
  { region: 'North America', revenue: '12500.50' },
  { region: 'North America', revenue: '8200.00' },
  { region: 'Europe', revenue: '14300.25' },
];

// Generate report document with deterministic SHA-256 source hash
const reportDoc = buildReport(template, rawData);
console.log('Source Hash:', reportDoc.provenance.sourceHash);
console.log('Worksheet Rows:', reportDoc.sheets[0].rows);
```

---

## Data Operations & Automation Guide

### Data Input: Positional JSON, CSV/JSON Files, and Stdin

Commands accepting tabular datasets (`data:append`, `data:update`, `data:append-table`, `data:batch-update`, `data:upsert`, `workbook:write`, `report:run`) support three flexible input mechanisms:

1. **Positional JSON argument (inline arrays):**
   ```sh-session
   $ google-sheet data:append '[["2026-03-01", "Hardware", "1450.00"], ["2026-03-02", "Software", "299.00"]]' -t Expenses -s <spreadsheetId>
   ```
2. **Input file (`--input` / `-i`):**
   Pass a path to a `.csv` or `.json` file. The format is automatically inferred from the file extension:
   ```sh-session
   $ google-sheet data:append-table -i ./data/march-sales.csv -t Sales -s <spreadsheetId>
   ```
3. **Standard Input (`-i -`):**
   Stream data directly via stdin. When piping CSV data, pass `--inputFormat csv`:
   ```sh-session
   $ cat ./transactions.csv | google-sheet data:update -i - --inputFormat csv --range 'Sheet1!A2:C100' -s <spreadsheetId>
   ```

### Read Render Options

When reading data from Google Sheets (`data:get`, `data:batch-get`) or local Excel workbooks (`workbook:read`), you can control how cell values, dates, and formula outputs are extracted:

* **`--valueRenderOption` (Google Sheets API):**
  * `FORMATTED_VALUE` *(default)*: Returns cell values formatted according to spreadsheet rules (e.g. `"$1,234.50"`, `"25.0%"`).
  * `UNFORMATTED_VALUE`: Returns raw scalar numbers, booleans, and strings without display formatting (e.g. `1234.5`, `0.25`).
  * `FORMULA`: Returns the formula text itself (e.g. `"=SUM(A1:A10)"`) rather than the evaluated result.
* **`--dateTimeRenderOption` (Google Sheets API):**
  * `SERIAL_NUMBER`: Returns dates as standard spreadsheet serial day numbers (e.g. `46098.5`).
  * `FORMATTED_STRING`: Returns dates as localized text strings matching cell formatting.
* **`--mode` (Local Excel `workbook:read`):**
  * `unformatted` *(default)*: Extracts underlying raw scalar values.
  * `formatted`: Extracts text representations as rendered by Excel format masks.
  * `formula`: Extracts formula expressions.

### Batch Operations & Native Append Table vs Legacy Semantics

#### Multi-Range Batch Reading and Updating

To minimize API round-trips and avoid quota throttling, use batch operations:

* **`data:batch-get`:** Fetches multiple disjoint ranges across one or more worksheets in a single API request:
  ```sh-session
  $ google-sheet data:batch-get --ranges '["Summary!A1:B5", "Transactions!A1:D50"]' -s <spreadsheetId>
  ```
* **`data:batch-update`:** Applies multiple range updates atomically in a single batch request:
  ```sh-session
  $ google-sheet data:batch-update -i ./batch-payload.json -s <spreadsheetId>
  ```
  *(Payload format: `[{"range": "Summary!B2", "values": [[100]]}, {"range": "Audit!A1:B2", "values": [["Status", "OK"]]}]`)*

#### `data:append-table` vs `data:append`

* **`data:append-table` (Recommended for structured tables):** Calls Google Sheets API's native `POST /values/{range}:append`. It scans for existing table boundaries, locates the true end of the tabular data, and appends rows directly to the table without scanning the entire worksheet grid. This preserves empty buffer rows and allows multiple tables to coexist on a single sheet.
* **`data:append` (Legacy bounding box append):** Reads the worksheet grid (`getData`), computes the overall bounding box (`maxRow`), and writes to row `maxRow + 1`.

### Schema Discovery, Validation, and Safe Data Mutation

#### `data:schema`: see what a worksheet actually holds (read-only)

```sh-session
$ google-sheet data:schema -s <spreadsheetId> -t "Q3 Sales"
```

Samples a bounded range (default rows 1–100, columns A–Z; move it with `--minRow/--minCol/--maxRow/--maxCol` — the first sampled row is the header row) and prints one line per column: the header text with its absolute A1 cell, inferred types with sample counts (`number (2), string`), the formula-cell count, and any data-validation rules (`ONE_OF_LIST (3)`). Named ranges are listed below the table. Inferred types describe the sample and nothing more — Google stores dates as numbers and formatting decides what they look like, so treat the output as evidence, not an authoritative schema. Empty headers, duplicate headers and mixed types come back as explicit warnings.

#### `data:validate`: check sheet data against a TableSchema (read-only)

```sh-session
$ google-sheet data:validate -s <spreadsheetId> -t "Q3 Sales" \
    --schema='{"fields":[{"name":"Name","type":"string","required":true},{"name":"Amount","type":"decimal","min":0},{"name":"Code","type":"string","unique":true}]}'
```

`--schema` takes the same TableSchema JSON that `report:run` templates use (fields typed `string|number|integer|decimal|boolean|date` with `required`/`min`/`max`/`enum`/`unique`); pass `--schemaFile <path>` (or `-` for stdin) instead of inline JSON. Every violation is reported with its sheet row, A1 coordinate, field, stable code and message: `MISSING_HEADER`/`DUPLICATE_HEADER` for layout problems, report codes such as `MISSING_REQUIRED_FIELD`, `ENUM_VIOLATION`, `MIN_VALUE_VIOLATION` or `NOT_AN_INTEGER` for data, and `DUPLICATE_VALUE` for `unique` fields. Invalid data exits 1 with error code `DATA_INVALID`; a malformed schema fails with `SCHEMA_INVALID` before any request leaves the process. Nothing is ever written.

Validation uses the same sample bounds as `data:schema` (default rows 1–100, columns A–Z), so a passing result does not certify rows outside that range. Uniqueness violations are collected even when other rows have type or required-field errors.

#### `data:clear`: values-only clear of a bounded range

```sh-session
$ google-sheet data:clear -s <spreadsheetId> -t "Scratch" --range "Scratch!A2:D20" --dryRun
$ google-sheet data:clear -s <spreadsheetId> -t "Scratch" --range "Scratch!A2:D20"
```

The range must be bounded on both axes — open-ended ranges like `A2:D` are refused. Clearing is values-only: number formats and every other cell property survive (one native values-clear call after the guard). Formula cells refuse the clear unless `--overwriteFormulas` is passed. `--dryRun` reports how many cells would be cleared — and which formula cells would refuse — with zero writes.

#### `data:upsert`: update matching rows, append new ones, keyed by one column

```sh-session
$ google-sheet data:upsert -s <spreadsheetId> -t "Expenses" --key=id '[["id", "category", "amount"], ["001", "Hardware", 1450], ["002", "Software", 299]]'
$ cat new-rows.csv | google-sheet data:upsert -s <spreadsheetId> -t "Expenses" --key=id -i - --inputFormat csv
$ google-sheet data:upsert -s <spreadsheetId> -t "Expenses" --key=id --range "Expenses!A1:F100" --dryRun -i rows.json
```

Input is header-first: the first input row names the columns and must match sheet headers. Existing keys update only the cells the input supplies — columns absent from the input keep their values and formulas untouched — and new keys append below the table. The semantics an agent should rely on:

- `--key` names exactly one key column; keys are typed, so the string `"001"` keeps its leading zero and never matches the number `1`. Empty and duplicate keys are rejected.
- Writes go out as `RAW` by default; `--valueInputOption USER_ENTERED` opts into Google's parsing.
- If `--range` bounds the table but populated cells continue beyond it, the upsert refuses before writing instead of silently missing rows.
- Formula cells this run would change are protected; `--overwriteFormulas` overrides.
- `--dryRun` reports added/updated/unchanged rows and the planned write ranges with zero mutations.
- Single-writer semantics only: re-running the same input does not duplicate rows, but the read-modify-write is not a transaction, there is no automatic retry after an ambiguous failure, and concurrent upserts against one sheet are unsupported.

#### `spreadsheet:copy` / `worksheet:copy`: duplicate a spreadsheet or a sheet

```sh-session
$ google-sheet spreadsheet:copy -s <spreadsheetId> --title "Archive 2026-09"
$ google-sheet worksheet:copy -s <spreadsheetId> -t "Template" --destinationSpreadsheetId <destId>
```

`spreadsheet:copy` leaves the source untouched and returns the new spreadsheet's server-assigned id and title. The CLI does not replay source sharing grants; inspect the copy's effective permissions before sharing sensitive data. `worksheet:copy` copies one worksheet into an explicit destination spreadsheet and returns the server-assigned title and new sheet id; an existing destination title does not require removing or renaming that sheet first.

#### `spreadsheet:export`: PDF/XLSX to a local file

```sh-session
$ google-sheet spreadsheet:export -s <spreadsheetId> --format pdf --output ./report.pdf
$ google-sheet spreadsheet:export -s <spreadsheetId> --format xlsx -o ./report.xlsx --overwrite
```

Binary-safe and atomic: the bytes are written to a temporary file in the output directory and renamed into place, so a failure leaves no partial file and preserves any existing output. An existing file is only replaced with explicit `--overwrite` (refused before the network round-trip, re-checked after the bytes arrive). Two Drive-side limits apply: exports are capped at 10 MB — larger spreadsheets fail with Google's own error — and access stays within the `drive.file` scope, meaning only files this app created or has opened are reachable, never the whole Drive.

#### `spreadsheet:list`: discover spreadsheets this app can see

```sh-session
$ google-sheet spreadsheet:list --name "Report 2026" --rawOutput
{"files":[{"id":"1AbC...","name":"Report 2026","mimeType":"application/vnd.google-apps.spreadsheet"}],"nextPageToken":"...","visibilityNote":"Listing uses the drive.file OAuth scope: ..."}
```

Discovery runs through the Drive API under the `drive.file` scope: only spreadsheets this application created or has opened are listed, and every response carries a `visibilityNote` saying so — **an empty result does not prove a spreadsheet is absent**; address such files by ID. `--name` filters by substring (`--exact` for a whole-title match), `--pageSize` bounds a page (1–100, default 50), `--pageToken` continues a listing, and `--all` aggregates pages up to a 1,000-file safety bound, returning the outstanding token when it truncates. Duplicate titles stay separate rows and nothing is ever selected automatically — pick an `id` yourself. OAuth tokens issued before `drive.file` was added must re-run `auth:login`.

#### `data:export-csv`: deterministic CSV from either backend

```sh-session
$ google-sheet data:export-csv --spreadsheetId <id> -t Sheet1 --output ./data.csv
Exported 3 row(s) x 2 column(s) from <id>/Sheet1 (Sheet1!A1:B3) to ./data.csv (48 bytes, mode: raw, injection: safe, transformed cells: 0)
$ google-sheet data:export-csv --workbook ./report.xlsx --range 'Sheet1!A1:D10' --mode formatted --output ./report.csv --overwrite
```

One source, exactly: `--spreadsheetId`/`--worksheetTitle` (optionally bounded by `--range`) for Google Sheets, or `--workbook <file.xlsx>` (with a required `--range`) for a local file. `--mode` selects the cell rendering — `raw` machine values (default), `formatted` display strings, `formula` formula text where the backend stores it — and `--injection` selects the formula-injection policy: `safe` (default) prefixes dangerous leading characters (`=`, `+`, `-`, `@`, tab, CR) with an apostrophe; `preserve` emits them byte-faithfully and says so in the receipt. Output is RFC 4180: with `--output` the file is written atomically (temp file + rename) and an existing file is replaced only with `--overwrite`, and the receipt (rows, columns, range, mode, policy, transformed cells, warnings) is printed; without `--output`, stdout carries only the CSV bytes and warnings go to stderr. Local XLSX formula results are the cached values from the last save by a real spreadsheet engine — never recalculated — and the warnings report how much of the export came from cache.

#### `capabilities`: the machine-readable capability document

```sh-session
$ google-sheet capabilities | jq '.backends["local-xlsx"].formulaSemantics'
```

Prints a versioned JSON document — no credentials, no auth, no flags — describing what this build can do: per backend (`google-sheets`, `local-xlsx`) the supported operations, input and output forms, formula write-vs-recalculation semantics, destructive guards, mutation limits and local fidelity exclusions, plus the `drive.file` visibility boundary. Static support is not authorization: an operation listed there still needs valid credentials and permissions when it runs.

### Structured JSON Errors for Agents

Any command can report failures as exactly one machine-readable envelope on stderr with exit code 1:

```sh-session
$ google-sheet data:get -s <spreadsheetId> --json
{"error":{"code":"USAGE","message":"...","retryable":false}}
```

- `--json`/`-j` turns on failure envelopes only — success output is unchanged. `--rawOutput`/`-r` prints success output as JSON too.
- `code` is stable: `USAGE` (bad argv), `AUTH_REQUIRED`/`UNAUTHORIZED`/`FORBIDDEN` (auth), `NOT_FOUND`, `CONFLICT`, `REQUEST_INVALID` (request rejected), `RATE_LIMITED`, `UPSTREAM`, `NETWORK` (retryable — `RATE_LIMITED` carries `retryAfterMs` when Google sends a hint), `VALIDATION`, `SCHEMA_INVALID`, `DATA_INVALID` (data contracts), `INTERNAL`.
- Coordinate-linked `issues` (`row`/`column`/`a1`/`field`/`code`/`message`/`value`) ride along on the envelope for `data:validate` violations and upsert rejections.
- Multi-request writes (`data:batch-update`, `data:upsert`) may attach a `mutation` block on failure: per-request `outcomes` (`acknowledged`, `rejected`, `unknown`, `not-attempted`) with the logical A1 ranges, grid-growth effects tracked separately, the `phase` the failure happened in, and a conservative `retryGuidance` (`safe-replay`, `verify-then-replay`, `never-blind-replay`). States, ranges and counts only — never cell contents. `unknown` means the write may have landed; verify state before replaying anything.
- `--redacted` (global flag) strips cell contents, formulas, incoming values and credentials from error envelopes and dry-run diagnostics before they are written — coordinates, counts, statuses and mutation outcome states survive. Enable it for unattended/CI runs; normal data reads and exports are unaffected.
- Envelopes never contain credentials, request/response internals or arbitrary error objects; retry only what reports `retryable: true`, and never replay an ambiguous write without checking state first.
- Errors raised before a command class loads, such as an unknown command or topic, still use oclif's human-readable output.

### Local Excel (XLSX) Workbooks: Inspect, Read, and Safe Write

Manage local `.xlsx` files without Google Cloud credentials or network calls:

#### 1. Preflight Inspection (`workbook:inspect`)

Audit workbook health, sheet names, defined names, comments, and check for unsupported binary features before running automation:

```sh-session
$ google-sheet workbook:inspect -f ./templates/budget-template.xlsx
```

#### 2. Range Extraction (`workbook:read`)

Extract structured cell data and inspect formula freshness:

```sh-session
$ google-sheet workbook:read -f ./models/q1-report.xlsx --range 'Summary!A1:E20' --mode unformatted
```

#### 3. Safe In-Place Writing (`workbook:write`)

Inject tabular data into an existing template or create a new workbook with built-in safety controls:

```sh-session
$ google-sheet workbook:write \
    -f ./templates/invoice-template.xlsx \
    -o ./dist/invoice-1042.xlsx \
    -i ./data/line-items.csv \
    -t "Line Items" \
    --startCell A5
```

**Safety flags:**
* `--inPlace`: Modifies the input file directly. Automatically creates a timestamped `.bak` backup file prior to write.
* `--dryRun`: Calculates diffs and outputs the planned cell modifications without touching disk.
* `--overwrite`: Explicitly confirms replacing an existing target file.
* `--overwriteFormulas`: By default, writing static data over a cell containing a formula is **rejected** to prevent accidental destruction of workbook logic. Pass `--overwriteFormulas` only when formula replacement is intentional.

#### 4. Structured Workbook Authoring: the ReportDocument (`workbook:write`)

Beyond flat rows, `workbook:write` accepts a full ReportDocument JSON object (the same shape `report:run` renders): `sheets` are created when missing, each cell may be a scalar or a `{formula, result}` object, and presentation rides on the sheet object — per-column `numberFormats`, `freezeRows`, and `columnWidths`. Formulas are stored as formula cells together with the cached result you supply; they are **never recalculated** — there is no calculation engine — so consumers read what a real spreadsheet application last computed. Features the local engine cannot represent (VBA macros, charts, pivot tables, external links, …) are detected at load time, and a save that would drop them **refuses** (`--allowUnsupportedFeatures` to override) — nothing is silently discarded.

### Report Automation (`report:run`)

The `report:run` command is a declarative report generation runner. It takes a template definition (`--template`), ingests input data from local files, stdin, local workbooks, or Google Sheets, executes validation and transformation pipelines, and renders formatted outputs into a local `.xlsx` file or a live Google Spreadsheet.

Runnable examples using templates and datasets from `examples/reports/`:

#### Example 1: Corporate Cash Flow & Multi-Currency FX Reconciliation (Finance)

Runs multi-account cash reconciliation with FX currency conversions, refund adjustments, and opening/closing balance verification:

```sh-session
$ google-sheet report:run \
    --template examples/reports/finance-template.json \
    --input examples/reports/finance-transactions.csv \
    --output dist/finance-reconciliation-q1.xlsx
```

#### Example 2: Platform Modernization & Manpower Estimation (Manpower)

Computes engineering effort, role rate multipliers, scenario contingency buffers (Optimistic, Base, Pessimistic), and budget allocations:

```sh-session
$ google-sheet report:run \
    --template examples/reports/manpower-template.json \
    --input examples/reports/manpower-tasks.csv \
    --output dist/manpower-plan.xlsx
```

#### Example 3: Sales Performance Report Targeting Google Sheets

Transforms raw order records into summarized regional sales tables with calculated discount margins, publishing directly to Google Sheets:

```sh-session
$ google-sheet report:run \
    --template examples/reports/sales-report-template.json \
    --input examples/reports/sales-data.csv \
    --spreadsheetId <spreadsheetId> \
    -f ./service-account.json
```

#### Example 4: Piped Standard Input with Formatting

```sh-session
$ cat examples/reports/finance-transactions.csv | google-sheet report:run \
    --template examples/reports/finance-template.json \
    --input - \
    --inputFormat csv \
    --output dist/finance-piped.xlsx
```

### Authentication Behavior: Cloud vs Local

* **Purely Local Execution:** When both input sources (`--input`, `--sourceWorkbook`) and targets (`--output`) are local files, `google-sheet` operates completely offline. No Google Cloud credentials, OAuth tokens, or network connectivity are required.
* **Cloud Execution:** When accessing Google Sheets (`--sourceSpreadsheet` or `--spreadsheetId`), the CLI initializes authentication via:
  1. `--credentialsFile` / `GSHEET_CREDENTIALS_FILE` (Service account JSON)
  2. `--clientEmail` / `--privateKey` (Service account keys)
  3. OAuth 2.0 user tokens established via `google-sheet auth:login`

### Report Templates: Schemas, Policies, and Transforms

Templates define the contract, validation rules, and transformation pipeline for reporting datasets:

* **Schema Field Definitions:** Fields declare data types (`string`, `number`, `integer`, `decimal`, `boolean`, `date`) and validation constraints (`required`, `min`, `max`, `enum`, `dateFormat`).
* **Validation Policies (`policy`):**
  * `strictTypes`: Rejects type coercion when set to `true`.
  * `treatMissingAsZero`: Substitutes numerical nulls/blanks with `0`.
  * `duplicateKeyPolicy`: Specifies deduplication behavior (`reject`, `keepFirst`, `keepLast`).
  * `maxInputRows`: Bounds maximum processed input rows (e.g. `50000`) to prevent memory exhaustion.
* **Transformation Pipeline (`transforms`):**
  * `computed`: Computes new fields via mathematical expressions (`add`, `sub`, `mul`, `div`, `round`, `abs`, `neg`), text manipulation, or conditionals.
  * `filter`: Evaluates boolean predicate conditions (`eq`, `neq`, `gt`, `gte`, `lt`, `lte`, `in`, `contains`).
  * `sort`: Orders datasets by one or more fields ascending or descending.
  * `aggregate`: Groups data by keys and calculates `sum`, `avg`, `min`, `max`, `count`, and `countDistinct`.
  * `pivot`: Pivots row values into columnar aggregates.

### Technical Invariants, Safety Guarantees, and Domain Policies

To guarantee reliability and prevent data corruption in automated pipelines, the engine enforces strict invariants:

1. **Formula Cache Is Not Recalculated:** Neither Google Sheets batch APIs nor ExcelJS embed a full formula recalculation engine. A formula cell contains the formula expression plus whatever static cached value was stored by the authoring spreadsheet software. When reading workbook ranges as source inputs for downstream reports, formula cells are tracked for freshness (`fresh` vs `stale_or_missing`). To consume cached formula values without live recalculation, you must explicitly provide `--allowCachedFormulaValues`.
2. **No Claim of Arbitrary Excel Compatibility:** The workbook engine safely handles standard OpenXML spreadsheets (worksheets, styles, numbers, dates, formulas, and data tables). Complex binary or proprietary Excel features (VBA macros `.xlsm`, embedded OLE objects, external SQL data connections, encrypted password-protected workbooks, and pivot cache slices) are detected during preflight inspection and **safely rejected** rather than silently corrupted.
3. **In-Place Backup & Concurrency Protection:** When `workbook:write --inPlace` is executed, the engine writes an atomic timestamped `.bak` file before applying changes. If the target file's SHA-256 hash changes between read and write, the write is aborted to avoid race conditions unless `--overwrite` is specified.
4. **Zero Side-Effect Dry Runs:** Passing `--dryRun` guarantees that no files are written to disk and no mutating API requests are sent to Google Cloud. A structured preview of planned mutations (cells modified, formulas preserved) is returned.
5. **Idempotence & Managed Named Ranges:** Reports written to Google Sheets or Excel use designated managed ranges (`_GS_MANAGED_<output_name>`) and deterministic SHA-256 content hashes. Re-running a report replaces the existing managed range cleanly without duplicating rows or leaving orphan headers.
6. **Chunk Partial Failure & Quota Protection:** Cloud batch operations are chunked according to Google Sheets API quotas (100 requests per 100 seconds per user) and retry transient errors with exponential backoff.
7. **Domain Semantics:**
   * **Cash-flow, not profit:** The finance report engine reconciles multi-account cash liquidity, opening/closing cash balances, FX conversions, and cash movements. It does **not** calculate accrual accounting profit or P&L statements.
   * **Effort, not duration:** The manpower report engine calculates engineering effort (person-days, person-hours), role rate multipliers, scenario budgets, and contingency buffers. It does **not** compute calendar project scheduling durations or Gantt timelines.
8. **Formula Injection Boundary:** Raw string values starting with `=`, `+`, `-`, `@`, or tab characters are sanitized (escaped with a leading single quote `'`) when written as data cells to prevent CSV/DDE formula injection vulnerabilities. Formulas are only evaluated when explicitly declared as `formulaTemplate` in report output column configurations.
9. **Money Rounding Policy:** All financial and currency arithmetic uses exact decimal math (`decimal.js` with half-up rounding, default 2 decimal places, and tolerance `0.01`), eliminating floating-point IEEE-754 precision drift (`0.1 + 0.2 !== 0.3`).

## Using with coding agents

google-sheet-cli is built to be driven by coding agents: `workbook:*` and local `report:run` run with zero credentials and zero network calls, every command emits machine-readable output (`--rawOutput`/`-r`, `--csv`), and writes fail closed — formulas are never overwritten without an explicit `--overwriteFormulas`, and `--dryRun` previews any mutation with zero side effects. Failures are programmatically consumable too: `--json`/`-j` reports any failure as one structured envelope on stderr (stable `code`, safe `message`, `retryable`, coordinate-linked `issues`, typed `mutation` outcomes for multi-request writes, exit code 1) while leaving success output unchanged — pair it with `--rawOutput`/`-r` when the agent wants success output as JSON as well. `capabilities` prints a credential-free, versioned JSON document of everything this build can do — read it before scripting; `spreadsheet:list` discovers spreadsheets under the `drive.file` scope (an empty listing never proves a file is absent — address those by ID); `data:export-csv` produces deterministic RFC 4180 interchange from either backend; and `--redacted` strips data values from diagnostics for unattended runs. See [Structured JSON Errors for Agents](#structured-json-errors-for-agents).

The repository ships an agent skill and an integration guide:

- [`skill/SKILL.md`](skill/SKILL.md) — agent skill covering every command, flag and workflow (install: copy into `~/.omp/skills/google-sheet/` or `~/.claude/skills/google-sheet/`).
- [docs/agents.md](docs/agents.md) — rules for agents: credentials via env/file never inline, the read → `--dryRun` → write safety ladder, offline-first workflow, exact-decimal report semantics.

<!-- commands -->
# Command Topics

* [`google-sheet auth`](docs/auth.md) - Authenticate with your Google account via OAuth 2.0
* [`google-sheet capabilities`](docs/capabilities.md) - Print the machine-readable capability document: per backend (google-sheets, local-xlsx) the supported operations, input and output forms, formula write-vs-recalculation semantics, destructive guards, mutation limits and local fidelity exclusions. Credential-free.
* [`google-sheet data`](docs/data.md) - Manage data in worksheet
* [`google-sheet format`](docs/format.md) - Apply cell formatting (text style, colors, alignment, wrap, number format, borders) or clear formatting. Only formatting is touched - cell values and formulas are never overwritten. With --workbook the format runs on a local XLSX file instead of Google Sheets, using the subset of flags ExcelJS models (wrapStrategy only as WRAP; --numberFormatType has no local equivalent).
* [`google-sheet grid`](docs/grid.md) - Delete rows or columns from a worksheet. Data after the deleted range shifts up/left. --dryRun previews the values about to be removed. With --workbook the delete runs on a local XLSX file instead of Google Sheets.
* [`google-sheet help`](docs/help.md) - Display help for google-sheet.
* [`google-sheet report`](docs/report.md) - Report automation runner and template generation
* [`google-sheet spreadsheet`](docs/spreadsheet.md) - Manage spreadsheets
* [`google-sheet workbook`](docs/workbook.md) - Manage local Excel (XLSX) workbooks
* [`google-sheet worksheet`](docs/worksheet.md) - Manage worksheets

<!-- commandsstop -->

# Info

## How to configure

### Step 1: Setting Up Google Service Account

1. Login to Google API Console: Visit the Google Cloud Console website (https://console.cloud.google.com/) and log in using your Google account credentials.
2. Enable Google Sheets API: In the Google Cloud Console, navigate to the "Library" section. Here, search for "Google Sheets API" and enable it.
3. Create a Service Account: Next, go to the "Credentials" section. Here, click on the "Create Credentials" dropdown button and select "Service Account". There's no need to assign any special role to this service account. Simply follow the prompts to create the account.
4. Download Credentials: Once the service account is created, a JSON file containing the credentials of the service account will be automatically generated. Download this file and keep it safe. You will need the `client_email` and `private_key` from this file to setup the Google Sheets Action.

### Step 2: Sharing the Spreadsheet

1. Share Spreadsheet: Go to the Google Spreadsheet that you want to use with this action. Click on the "Share" button (usually at the top right corner) and in the sharing settings, add the `client_email` (that you got from the downloaded JSON file) with read permissions.
2. Get Document ID: The document ID is the string of random characters in the URL of your Google Spreadsheet, found between '/d/' and '/edit'. Keep this document ID handy.

## Credentials

Every command authenticates as the service account from Step 1, using the `client_email` and the `private_key` of the downloaded JSON file. There are three ways to hand them over:

1. **The JSON file** - `--credentialsFile=<path>` (`-f`), or the `GSHEET_CREDENTIALS_FILE` env variable.
2. **The two values** - `--clientEmail=<value>` (`-c`) and `--privateKey=<value>` (`-p`), or the `GSHEET_CLIENT_EMAIL` and `GSHEET_PRIVATE_KEY` env variables.
3. **Neither** - the cli prompts for what is missing.

They can be mixed, because precedence is field by field: `--clientEmail` and `--privateKey` win over the same field in the credentials file, the file fills in whatever they left out, and only what is still missing is prompted for.

```sh-session
$ google-sheet spreadsheet:get -s <spreadsheetId> -f ./service-account.json
```

The `private_key` is the whole PEM block from the JSON file, `BEGIN` and `END` lines included:

```
-----BEGIN PRIVATE KEY-----
MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQ...
-----END PRIVATE KEY-----
```

Escaped newlines (`\n`) and surrounding quotes are taken care of, so the value can be pasted straight out of the JSON file or out of a CI secret. A key converted to the older `BEGIN RSA PRIVATE KEY` (PKCS#1) form is accepted as well. Note that the `private_key_id` is a different field and will not work. A key that is truncated or damaged is rejected before any request is made; run with `DEBUG=gsheet:credentials` to see the parser error behind the rejection.

For anything scripted, hand the key over with `--credentialsFile` or the environment variables rather than piping it into the prompt. The prompt strips backslashes when its input is not a terminal, which turns the escaped newlines into stray `n` characters and makes a perfectly good key look broken.

## Build with

- [@googleapis/sheets](https://github.com/googleapis/google-api-nodejs-client/tree/main/src/apis/sheets) - The node module used for manipulating the google sheet. 2.x used the whole `googleapis` bundle; see [Migrating from 2.x](#migrating-from-2x)
- [oclif](https://oclif.io) - The node module used to create the cli
- [semantic-release](https://github.com/semantic-release/semantic-release) - for releasing new versions
- [typescript](https://www.typescriptlang.org)

## Contributing

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on our code of conduct, and the process for submitting pull requests to us.

## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/your/project/tags).

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details

## TODO

- [x] documentation
- [ ] more tests
- [ ] add prettier
