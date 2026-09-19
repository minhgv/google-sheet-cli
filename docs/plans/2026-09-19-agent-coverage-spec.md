# Agent coverage roadmap — find, formatting, grid structure, sharing

Work ID: 2026-09-19-agent-coverage-spec
Status: All three phases implemented, verified, live-tested, committed `ac78853` and pushed to `origin/master` (2026-09-19). semantic-release will cut a minor version on the next release run.
Next safe action: none — record closed. Monitor CI for the release tag.

## Context

Assessment (2026-09-19, same-day conversation) found the CLI covers ~85–90% of agent↔Sheets workflows: rawOutput/csv everywhere, formula guard, `--dryRun`, batch APIs, offline workbook commands, skill file. Remaining gaps, ranked by agent value:

1. **Cell formatting** — borders, bold headers, money/number formats, colors, merge. User's stated need: "kẻ bảng, bôi đậm header, format số tiền để xuất báo cáo đẹp". Internal primitives already exist (`repeatCell`/`updateDimensionProperties`/`updateSheetProperties` built in `google-sheet.ts:1196-1255,1373-1426` for the report pipeline) but no user-facing command exposes them.
2. **`data:find`** — locate cell/row coordinates by condition so an agent can target writes without downloading whole sheets.
3. **Grid structure** — insert/delete/hide rows/columns, freeze panes, column resize. `appendDimension`/`updateSheetProperties` already used internally; `insertDimension`/`deleteDimension`/`updateDimensionProperties`/`autoResizeDimensions` are new request types.
4. **`spreadsheet:share`** — Drive permissions so an agent can hand a generated report to a human. Requires a new OAuth/JWT scope (`drive.file`) — a real design decision, see Phase 3.
5. **Charts / pivot tables** — deferred; large spec surface, no concrete consumer yet.

Non-goals: reimplementing the `financial` repo's audit logic inside the CLI; a general Excel-style formatting engine; conditional formatting and data validation (deferred, noted); charts/pivots (deferred).

## Approach

- One new `format:` topic and one new `grid:` topic; `spreadsheet:` gains share commands. Naming mirrors Sheets API vocabulary (`repeatCell`, `DimensionRange`, `permissions`).
- All new write commands follow the existing safety ladder: `--dryRun` previews the exact `batchUpdate` request bodies (zero side effects); real write requires omitting it. Formatting never touches `userEnteredValue` — field masks are restricted to `userEnteredFormat.*` so a format command cannot clobber data by construction.
- Reuse, don't rebuild: `parseStrictA1Range`/`formatBoundedA1Range`/`escapeWorksheetTitle`/`toGoogleNumberFormat`/`toGoogleExtendedValue` from `src/lib/sheet-batch.ts`; `resolveDataMatrix`/`resolveBatchUpdates` input pattern from `src/lib/cli-input.ts`; `spreadsheetId`/`worksheetTitle`/`rawOutput` flags from `base-class.ts`; `factory.createGoogleSheet()` seam for offline tests.
- Library-first: every command is a thin oclif wrapper over a `GoogleSheet` method, matching the existing architecture (`gsheet.action` consumes the library half — new methods must be additive, never signature changes).
- Fake-first testing: extend `test/fake-sheets.ts` `batchUpdate` handler for the new request types *behaviorally* where state matters (insert/delete dimension must actually shift cells so a follow-up `values.get` proves it); request-shape-only types (repeatCell, updateBorders, mergeCells) are asserted via `recordedRequests` bodies.
- Conventional Commits per command (`feat:` → minor bumps under semantic-release). Docs regenerate via `npm run version`; never hand-edit `docs/*.md` or README usage blocks. `skill/SKILL.md` updated per phase and re-copied to `~/.omp/skills/google-sheet/SKILL.md`.

## Critical files and ownership

- `src/lib/google-sheet.ts` — new methods: `findData`, `formatCells`, `mergeCells`, `mutateDimension` (insert/delete/hide/resize shared builder), `setFrozen`, `shareSpreadsheet`, `listPermissions`, `revokePermission`. Sole owner: whichever worker implements the phase; methods are additive, no overlap between phases.
- `src/lib/sheet-format.ts` (new) — pure builders: CLI style spec → `sheets_v4.Schema$Request[]` + field-mask computation. No I/O; unit-testable without the fake.
- `src/lib/cli-input.ts` — add `resolveStyleSpec` (JSON from positional arg / `-i file` / `-i -`), reusing the existing three-source resolution contract.
- `src/commands/format/cells.ts`, `format/merge.ts` (new); `src/commands/grid/insert.ts`, `grid/delete.ts`, `grid/hide.ts`, `grid/resize.ts`, `grid/freeze.ts` (new); `src/commands/data/find.ts` (new); `src/commands/spreadsheet/share.ts`, `spreadsheet/permissions.ts`, `spreadsheet/unshare.ts` (new).
- `test/fake-sheets.ts` — extend `batchUpdate` for `insertDimension`, `deleteDimension`, `mergeCells`, `unmergeCells`, `autoResizeDimensions`, `updateBorders`, `repeatCell` (record-only), `updateDimensionProperties` (record + hiddenByUser/pixelSize state); add `drive.googleapis.com` host handling for Phase 3.
- `test/commands/offline.test.ts` — command-level cases per phase; `test/` new unit file for `sheet-format.ts` builders.
- `skill/SKILL.md`, `docs/agents.md` — per-phase updates; `README.md`/`docs/*.md` via `npm run version` only.

## Spec — Phase 1: `format:cells`, `format:merge`, `data:find`

### `data:find` — locate cells/rows by condition

Purpose: answer "where is X" without pulling a whole sheet into context. Pure read; no new API surface (one `values.get`).

Flags:
- `-s/--spreadsheetId` (required), `-t/--worksheetTitle` (required).
- `--range <A1>` — bound the scan inside the worksheet; default = whole used grid.
- Match mode, exactly one required: `--equals <v>` | `--contains <v>` | `--regex <pattern>`.
- `--column <letter>` — restrict to one column by A1 letter. `--header <name>` — restrict to the column whose first scanned row cell equals `<name>` (mutually exclusive with `--column`).
- `--ignoreCase` (default true; `--no-ignoreCase` for exact case).
- `--valueRenderOption FORMATTED_VALUE|UNFORMATTED_VALUE|FORMULA` (default `FORMATTED_VALUE` — matches what a human sees, e.g. `26,500`).
- `--limit <n>` (default 100), `--first` (alias for `--limit 1`), `--byRow` (collapse matches to unique rows, include full row values).
- `--rawOutput` standard.

Matching semantics: empty cells skipped. `--equals` compares stringified cell text; when both sides parse as finite numbers, numeric equality wins (so `26500` matches `26,500` under UNFORMATTED but not FORMATTED — documented). `--regex` is JS `RegExp` over the stringified cell; invalid pattern → `Error` exit 1 before any network call.

Output (`--rawOutput`):
```json
{
  "range": "T1!A3:K64",
  "matchCount": 3,
  "truncated": false,
  "matches": [
    { "a1": "T1!B12", "row": 12, "column": 2, "columnLetter": "B", "value": "NV0123" }
  ]
}
```
`--byRow` adds `rowValues` (full scanned row array) and dedups by `row`. Human output: `ux.table` of `a1 | row | column | value`.

Library: `findData(options: FindOptions, spreadsheetId?): Promise<FindResult>` — resolves worksheet, one `values.get` with the chosen render option, client-side scan. No formula guard, no writes.

Acceptance: `data:find -t T1 --equals "NV0123" --column B --first --rawOutput` returns the exact `a1` coordinate; `--byRow` returns one entry per row; zero matches → `matchCount: 0`, `matches: []`, exit 0 (not an error — agents branch on count).

### `format:cells` — styling without touching values

Purpose: borders, bold headers, money formats, colors, alignment, wrap — the "pretty report" surface.

Flags:
- `-s`, `-t`, `--range <A1>` (required; single bounded range per call — multi-range via repeated calls or `--input` batch spec).
- Convenience style flags (any combination): `--bold`, `--italic`, `--underline`, `--strikethrough`, `--fontSize <pt>`, `--fontFamily <name>`, `--textColor <#RRGGBB>`, `--backgroundColor <#RRGGBB>`, `--horizontalAlignment LEFT|CENTER|RIGHT`, `--verticalAlignment TOP|MIDDLE|BOTTOM`, `--wrapStrategy OVERFLOW_CELL|CLIP|WRAP`, `--numberFormat <pattern>` (e.g. `#,##0.00`, `0.0%`, `YYYY-MM-DD`), `--numberFormatType TEXT|NUMBER|PERCENT|CURRENCY|DATE|TIME|DATE_TIME|SCIENTIFIC` (optional; else inferred by existing `toGoogleNumberFormat`).
- Borders: `--borders <sides>` where sides = comma list of `top,bottom,left,right,innerHorizontal,innerVertical,all`; `--borderStyle DOTTED|DASHED|SOLID|SOLID_MEDIUM|SOLID_THICK|DOUBLE|NONE` (default `SOLID`); `--borderColor <#RRGGBB>` (default black).
- `--style <json>` / `-i <file>` / `-i -` — full spec escape hatch: `{ "cellFormat": {...CellFormat subset...}, "borders": {"top": {"style": "SOLID", "color": "#000"}} , "ranges": ["A1:C1","A2:C9"] }`. `ranges` inside the spec makes one call format many ranges atomically; `--range` and spec `ranges` are mutually exclusive.
- `--clear` — reset all formatting on the range (`repeatCell` with empty `userEnteredFormat`, fields `userEnteredFormat`). Mutually exclusive with every style flag.
- `--dryRun` — print the computed `batchUpdate` request bodies + field masks; zero network mutation.
- `--rawOutput` standard.

Semantics: builds ONE `spreadsheets.batchUpdate` containing `repeatCell` (format fields) and/or `updateBorders` — atomic. Field mask is computed from exactly the flags supplied (never `*`): `--bold` → `userEnteredFormat.textFormat.bold`, etc. Values are never in the mask — formatting cannot overwrite data by construction. Hex colors → `colorStyle.rgbColor` floats (0–1). `--clear` on a range with formulas is safe (format-only).

Output: `{ range(s), requests: <count>, fields: [...], dryRun }`; human output summarizes applied properties.

Library: `formatCells(spec: CellFormatSpec, spreadsheetId?): Promise<FormatReceipt>`; spec→requests lives in pure `sheet-format.ts` (`buildFormatRequests(spec, sheetId)`), unit-tested without network.

Acceptance: `format:cells -t Report --range A1:J1 --bold --backgroundColor "#1a73e8" --textColor "#ffffff"` then `format:cells --range A2:J50 --numberFormat "#,##0.00" --borders all` — verified live by `data:get --valueRenderOption FORMATTED_VALUE` showing formatted numbers, and by fake-recorded request bodies offline. `--dryRun` emits identical request JSON with no mutation.

### `format:merge` — merge/unmerge cells

Flags: `-s`, `-t`, `--range <A1>` (required), `--type MERGE_ALL|MERGE_COLUMNS|MERGE_ROWS` (default `MERGE_ALL`), `--unmerge` (switch to `unmergeCells`), `--dryRun`, `--rawOutput`.

One `mergeCells`/`unmergeCells` request. Merging a range with data in >1 cell keeps top-left (Google semantics — documented, not guarded).

## Spec — Phase 2: `grid:` structural commands


Shared flags: `-s`, `-t`, `--dimension ROWS|COLUMNS` (required), `--start <n>` (1-based first index), `--count <n>` (default 1), `--dryRun`, `--rawOutput`. Internally all resolve worksheet → `sheetId`, convert to 0-based `startIndex`/`endIndex` (exclusive), build a `DimensionRange`.

- **`grid:insert`** → `insertDimension` + `--inheritFromBefore` (default false = inherit from row/col after, matching Sheets UI "insert above/left" behavior).
- **`grid:delete`** → `deleteDimension`. `--dryRun` additionally fetches and prints the values about to be removed (bounded to the dimension range) — the only preview that matters for a destructive op.
- **`grid:hide`** → `updateDimensionProperties` `{hiddenByUser: true}`; `--unhide` flag flips to `false` (one command, not two).
- **`grid:resize`** → `--pixels <n>` (`updateDimensionProperties` `pixelSize`) XOR `--auto` (`autoResizeDimensions`).
- **`grid:freeze`** → `updateSheetProperties` `gridProperties.frozenRowCount`/`frozenColumnCount`; flags `--rows <n>`/`--columns <n>` (0 unfreezes), no `--dimension`/`--start`.

Fake-sheets work (the real cost of this phase): `insertDimension`/`deleteDimension` must shift the `cells` map and adjust `rowCount`/`columnCount` so a post-op `data:get` proves the move; `updateDimensionProperties` stores `hiddenByUser`/`pixelSize` on the worksheet; `autoResizeDimensions`/`mergeCells`/`unmergeCells`/`updateBorders`/`repeatCell` record-and-ack (assert via `recordedRequests`).

Acceptance: insert 2 rows before row 5 → prior row-5 values readable at row 7; delete → gone and grid shrunk; hide → `spreadsheets.get` shows `hiddenByUser`; freeze → `frozenRowCount` in metadata.

## Spec — Phase 3: `spreadsheet:share` (+ `permissions`, `unshare`)

Design decision (needs user sign-off): sharing is **Drive API**, not Sheets. Scope choice:

- **`drive.file`** (recommended): covers files the CLI itself created (`spreadsheet:add` under OAuth) — the agent report-handoff flow. Cannot share pre-existing files the app never opened.
- **`drive`** (full): shares anything the user can share; sensitive scope → heavier Google verification for the public-client plan.

Implementation: raw `authClient.request({url: 'https://www.googleapis.com/drive/v3/files/...'})` — no `@googleapis/drive` dependency (repo deliberately carries only `@googleapis/sheets`). Scope added to both OAuth `SCOPES` and the service-account JWT scopes → **existing `token.json` must re-login** (documented in command description + docs). Service accounts can share files they own/created.

- **`spreadsheet:share`** — `--email <addr>` (`multiple: true`) XOR `--domain <d>` XOR `--anyone`; `--type user|group|domain|anyone` (inferred from which flag, overridable); `--role reader|commenter|writer` (default `reader`); `--notify` (default **false** — agents shouldn't spam; Google default is true); `--message <text>`; `--rawOutput` → `{permissionId, role, type, emailAddress}`.
- **`spreadsheet:permissions`** — list: `{id, type, role, emailAddress, displayName}`[].
- **`spreadsheet:unshare`** — `--permissionId <id>` XOR `--email <addr>` (resolves id via list).

Fake-sheets: add `www.googleapis.com/drive/v3` host branch — `permissions.create`/`list`/`delete` against an in-memory permission map keyed by fileId.

Acceptance: create spreadsheet via `spreadsheet:add` → `share --email x@y --role writer` → `permissions` lists it → `unshare` removes it — all offline against the fake; live path verified once on the test spreadsheet.

## Deferred (explicit non-goals this roadmap)

- **Charts / pivot tables** (`addChart`, `addPivotTable`, `updateChartSpec`): large spec surface, no consumer. Revisit when a report template needs one.
- **Conditional formatting** (`addConditionalFormatRule`): natural `format:` extension later.
- **Data validation** (`setDataValidation`): dropdowns for agent-built input sheets; small, but no current need.
- **`copyPaste`/`cutPaste`/`findReplace`**: `findReplace` is a plausible `data:replace` later; `data:find` covers the locate half now.

## Verification

- AC-01 (`data:find`): offline fake — equals/contains/regex, `--column`/`--header`, `--byRow`, `--limit`, zero-match exit 0; live probe on test spreadsheet.
- AC-02 (`format:cells`/`merge`): recorded request bodies carry exact field masks and never `userEnteredValue`; `--dryRun` mutates nothing; live formatted read-back.
- AC-03 (`grid:*`): fake cell-shift verified via post-op `values.get`; freeze/hide visible in `spreadsheets.get` metadata.
- AC-04 (`spreadsheet:share`): offline Drive fake round-trip; scope documented; re-login requirement stated.
- AC-05 (regression): `npm run test:unit` green incl. `test/regression.test.ts`/`grid-growth.test.ts` untouched; `npm run build`; `./bin/run.js <cmd> --help` smoke per command.
- AC-06 (docs): `npm run version` regenerates README/docs; `skill/SKILL.md` + `docs/agents.md` updated; installed skill copy refreshed.

## Execution checklist (per phase, on approval)

- [x] T-01: `sheet-format.ts` builders + unit tests (AC-02).
- [x] T-02: `format:cells` + `format:merge` commands + offline tests (AC-02).
- [x] T-03: `data:find` + `findData` + tests (AC-01).
- [x] T-04: fake-sheets dimension mutations + `grid:*` commands + tests (AC-03).
- [x] T-05: `drive.file` scope → `spreadsheet:share|permissions|unshare` + Drive fake + tests (AC-04). Live path pending Drive API enablement.
- [x] T-06: docs regeneration, SKILL.md, agents.md, full suite via test-runner (AC-05, AC-06) — Phase 1 scope.

## Assumptions and contingencies

- No new npm dependencies for any phase (Drive via raw `authClient.request`; formatting via existing `@googleapis/sheets` types).
- `drive.file` scope is the default recommendation; user may choose full `drive` — one-line scope change, noted verification cost.
- Existing OAuth tokens lack the Drive scope → Phase 3 ships with a documented re-login requirement, not silent scope creep.
- Each phase is independently shippable (`feat:` commits); Phase 1 alone already closes the user's stated "pretty report" gap.

## Phase 1 execution evidence (2026-09-19)

- [x] T-01: `src/lib/sheet-format.ts` — `hexToRgbColor`, `a1ToGridRange`, `buildCellFormat`, `buildFormatRequests`, `buildMergeRequest`; field masks confined to `userEnteredFormat.*`.
- [x] T-02: `src/commands/format/cells.ts` + `format/merge.ts`; `GoogleSheet.formatCells`/`setMerge`; `test/fake-sheets.ts` extended with `mergeCells`/`unmergeCells`/`updateBorders` handlers and `merges` on `FakeWorksheet`/`renderSheetProperties`.
- [x] T-03: `src/commands/data/find.ts` + `GoogleSheet.findData`; equals/contains/regex, `--column`/`--header`, `--byRow`, `--limit`, zero-match exit 0.
- [x] T-06 (Phase 1 scope): `npm run version` regenerated `README.md`, `docs/format.md`, `docs/data.md`; `skill/SKILL.md` + `docs/agents.md` updated; installed skill copy refreshed at `~/.omp/skills/google-sheet/SKILL.md`.
- Verification: `npm run test:unit` → 353 passing, 0 failing (incl. `test/format-find.test.ts` 26 cases, `test/commands/offline.test.ts` updated with 3 command entries + stub methods). `npm run build` clean.
- Live verification on `1WD2go7gcZRSus83RBPJUStHQM04JOFESVRZsZQSkFU0`/`feature-check`: `data:find --equals` → `A2`/`A4`; `--header "Status" --byRow` → column B + full row values; `format:cells --dryRun` → exact request JSON, no mutation; real `format:cells` → `26,500.00` formatted read-back with values intact; `format:merge` → merge visible in `spreadsheet:get` metadata, `--unmerge` removed it.
- Deviations from spec: `--clear`/`--input` exclusivity enforced in `run()` rather than oclif `exclusive` (clearer error text); `borderStyle`/`borderColor` carry no `dependsOn` because oclif 5 treats a defaulted flag as provided and would demand `--borders` on every call.
- Not committed: awaiting explicit user authorization per repo rules.

## Edge-case hardening (2026-09-19, post-Phase-1)

Edge-case tests written first (red), then fixes. 28 new cases: 19 in `test/format-find.test.ts`, 9 command-level in `test/commands/offline.test.ts`. Suite: 353 → 381 passing.

Bugs found and fixed:
- **`a1ToGridRange` single-cell range** (`src/lib/sheet-format.ts`): a bare `"B2"` produced a GridRange with no end indices, which the API reads as "to the grid edge" — one cell would have formatted the whole row/column/grid. End indices now pin to the start when only the start is given; column-only (`B:D`) and row-only (`2:4`) ranges keep the other axis unbounded on purpose.
- **`findData` empty needle** (`src/lib/google-sheet.ts`): `contains: ''` and `regex: ''` matched every cell in the range — a match-all trap. Both now reject with a `non-empty` error. `equals: ''` is meaningful (locates empty cells) and now matches them; Google trims trailing empty rows so only in-grid empty cells are reachable.
- **`findData` `limit: 0`**: an explicit 0 fell back to the default 100. Now honored as "count only, return nothing" (`matchCount` still reports the true total, `matches: []`).

Behaviors pinned by test (not bugs):
- `--header` resolves against the first scanned row, not worksheet row 1 — a sub-range starting below the real header throws `Header "X" not found`.
- A range carrying another sheet's quoted title scans that sheet; `a1` coordinates name the sheet the data came from.
- `equals` uses numeric equality for numeric cells (`"26500"` matches `26500`, `"26,500"` does not).
- Command-level: `--equals`+`--contains`, `--column`+`--header`, `--clear`+style, `--input`+style, `--type`+`--unmerge` all rejected before any Sheets mutation.

## Phase 2 execution evidence (2026-09-19)

- [x] T-04: `GoogleSheet.mutateDimension` (insert/delete/hide/resize) + `setFrozen`; `DimensionOptions`/`DimensionReceipt`/`FreezeOptions` types; 5 commands `grid:insert|delete|hide|resize|freeze` sharing `dimensionFlags` from `grid/insert.ts`.
- fake-sheets: `insertDimension`/`deleteDimension` shift the cells map and adjust `rowCount`/`columnCount` (post-op `data:get` proves the move); `updateDimensionProperties` stores `hiddenByUser`/`pixelSize` on `FakeWorksheet.dimensionProps`; `autoResizeDimensions` ack.
- `grid:delete --dryRun` returns `affectedValues` — the values about to be removed.
- Verification: `npm run test:unit` → 409 passing, 0 failing (23 new in `test/grid.test.ts`, 5 command entries + 2 stubs in `offline.test.ts`). `npm run build` clean.
- Live verification on `feature-check`: `grid:insert ROWS start=3 count=2` → rows shifted, `26,500.00` format preserved on moved cells; `grid:delete` restored; `grid:freeze --rows=1` → `frozenRowCount` in metadata; `grid:hide`/`--unhide` on column D round-tripped.
- Deviations: none beyond spec. Human output polished post-live ("Hid 1 column(s)", "Froze 1 row(s)").
- Docs: `npm run version` regenerated `docs/grid.md` + README; `skill/SKILL.md` + `docs/agents.md` updated; installed skill copy refreshed.
- Not committed: awaiting explicit user authorization per repo rules.

## Phase 3 execution evidence (2026-09-19)

- [x] T-05: scope decision = `drive.file` (user-approved). `shareSpreadsheet`/`listPermissions`/`unshareSpreadsheet` via raw `authClient.request` — no `@googleapis/drive` dependency. `authClient` field stored on both `authorize` (JWT, scopes `[sheets, drive.file]`) and `authorizeOAuth` paths; `oauth.ts` `generateAuthUrl` gains `drive.file`.
- Commands: `spreadsheet:share` (`--email`×N XOR `--domain` XOR `--anyone`, `--type` override, `--role` default reader, `--notify` default OFF + `--message` dependsOn notify), `spreadsheet:permissions`, `spreadsheet:unshare` (`--permissionId` XOR `--email`, resolves via list).
- fake-sheets: `www.googleapis.com/drive/v3` host branch — `permissions.create`/`list`/`delete` against in-memory map keyed by fileId; 404 for unknown fileId.
- Error translation in `driveRequest`: 403 → "re-run auth:login" hint; 404 → "drive.file only sees files this app created/opened" hint.
- Live path: token re-granted with `drive.file` (user re-logged in); Drive API enabled in Cloud project 701850506685. Round-trip verified on `1WD2go7gcZRSus83RBPJUStHQM04JOFESVRZsZQSkFU0`: `permissions` listed owner → `share --email giapminh79@gmail.com --role reader` granted `perm-13392153234400355928` → `unshare --email` resolved + removed → final list back to owner only.
- Docs: `npm run version` regenerated `docs/spreadsheet.md` + README; `skill/SKILL.md` + `docs/agents.md` updated; installed skill copy refreshed.
- Not committed: awaiting explicit user authorization per repo rules.
