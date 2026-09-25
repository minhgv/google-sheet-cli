# Local XLSX Multi-Tool Roadmap — Spec & Feasibility

Status: ALL WAVES DELIVERED (2026-09-24) — 809 offline tests green; reviewer pass done, 5 findings repaired.
Date: 2026-09-23

## Context

`google-sheet-cli` today is asymmetric: the Google Sheets backend has a rich
command surface (`data:*`, `grid:*`, `format:*`, `worksheet:*`, `spreadsheet:*`)
while the local XLSX backend has exactly three commands
(`workbook:inspect|read|write`) plus `report:run` output. The engine
(`src/lib/xlsx.ts`, `XlsxWorkbook` over **ExcelJS 4.4.0**) already does atomic
save, sha256 tracking, preflight feature detection, formula-conflict guards and
managed extents — but exposes none of ExcelJS's structural APIs.

Goal: make the CLI a **general-purpose** spreadsheet tool — one command surface
that works on both backends wherever the underlying engine allows — not a
filler for any specific template. The openpyxl gap analysis (what a
template-surgery script like `fill_excel.py` can do that we cannot) is used as
the benchmark for "general purpose".

Non-goals:

- Formula recalculation (neither ExcelJS nor openpyxl computes; documented
  exclusion, freshness tracking already exists).
- Chart / pivot-table / VBA preservation (ExcelJS hard limit — see
  "Hard limits").
- `.xls` / `.xlsb` / `.xlsm` formats.
- A second engine (e.g. raw OOXML surgery) — out of scope; ExcelJS limits are
  documented, not worked around.

## Feasibility

ExcelJS 4.4.0 capability inventory vs. what `XlsxWorkbook` exposes today:

| Capability | ExcelJS API | Exposed? | Effort to expose |
|---|---|---|---|
| Insert/delete rows & cols | `ws.spliceRows`, `ws.spliceColumns` | No | S |
| Merge/unmerge | `ws.mergeCells`, `ws.unMergeCells` | No (cloud-only `format:merge`) | S |
| Read merged ranges | `ws.model.merges` | No | S |
| Sparse per-cell write | `ws.getCell(r,c).value` | No (`apply()` writes whole ranges) | S |
| Cell/row styles | `cell.style`, `row.style` clone | Partial (numberFormats only) | M |
| Freeze / col width / row height | `ws.views`, `col.width`, `row.height` | Only on report outputs | S |
| Sheet add/remove/rename | `addWorksheet`, `removeWorksheet`, `ws.name` | No (cloud-only `worksheet:*`) | S |
| Data validation | `ws.dataValidations` (partial types) | No | M |
| Conditional formatting | `ws.addConditionalFormatting` (partial rules) | No | M |
| Defined names | `workbook.definedNames` | Read-only via `inspect()` | S |
| Cell comments | `cell.note` | No | S |
| **Formula ref rewrite on structural edit** | **none** | No | **L — new code** |
| Charts / pivots / VBA round-trip | **none — dropped on save** | Preflight warns | impossible |

Key finding: ~80% of the gap is **exposing what ExcelJS already does**. The
only genuinely new logic is formula-reference rewriting after row/column
insertion — and openpyxl does not have that either (its users patch refs by
hand). Shipping it would make the CLI strictly stronger than openpyxl at the
exact operation template surgery needs most.

## Design decision: one surface, two backends

`data:export-csv` already proves the pattern: `-s <id> -t <title>` **or**
`--workbook <file.xlsx>` selects the backend. Recommendation: extend the
existing `grid:*`, `format:*`, `worksheet:*`, `data:find`, `data:clear`
commands with a `--workbook` flag instead of growing a parallel `workbook:*`
namespace.

- Pro: single mental model, flags identical across backends, `capabilities`
  already models per-backend operation lists.
- Con: commands gain a second code path; mitigated by keeping all XLSX logic
  inside `XlsxWorkbook` and having commands dispatch on which flag group is
  present (exactly what `data:export-csv` does).
- Alternative (rejected): `workbook:insert-rows`, `workbook:merge`, … — doubles
  the surface, drifts from cloud flag names, harder to document.

Invariants carried over to every new local mutation:

- `--dryRun` previews with zero side effects (diff or planned ops, never a
  saved file).
- Atomic save only (existing `saveBufferAtomic`, mode 0600); `--inPlace`
  requires the file to be loaded from that path; refuse clobber without
  `--overwrite`.
- Formula cells are never overwritten/shifted-away silently:
  `--overwriteFormulas` consent gate, same vocabulary as today.
- `--json`/`-j` failure envelope + `--redacted` apply unchanged.
- Preflight `unsupportedFeatures` gate: saving a workbook that would drop
  charts/pivots requires explicit `--discardUnsupported` (new consent flag —
  today it is only a warning; see AC-09).

## Feature spec

### F1 — `workbook:find` (local `data:find`)

Mirror of `data:find` flags: `--equals|--contains|--regex`, `--range`,
`--column|--header`, `--ignoreCase`, `--limit|--first`, `--byRow`, render
options. Returns the same `{matchCount, truncated, matches:[{a1,row,column,
columnLetter,value,rowValues?}]}` shape so agent code is backend-agnostic.
Implementation: iterate `ws.eachRow` inside the parsed range, reuse the
matcher semantics of `data:find`. Effort: **S**.

### F2 — `workbook:inspect` extension

Add per sheet: `mergedRanges: string[]` (A1 ranges from `ws.model.merges`),
`formulaCells: {count, addresses?}` (addresses behind `--includeFormulaCells`
to bound output), `dataValidations: {range,type}[]`, `conditionalFormatting:
{range,ruleCount}[]`, `frozen: {rows,columns}`. Top level: `definedNames`
(already present), `unsupportedFeatures` (already present). Effort: **S**.

### F3 — `grid:insert` / `grid:delete` on `--workbook`

Flags unchanged: `--dimension ROWS|COLUMNS --start N [--count N]
[--inheritFromBefore] [--dryRun]`. Local semantics via `spliceRows`/
`spliceColumns`; `--inheritFromBefore` clones the style of row `start-1` onto
inserted rows (ExcelJS inserts bare rows otherwise). `delete --dryRun` prints
the values about to be removed — same contract as cloud. New flag
`--updateRefs` (F7) optional. Effort: **S** without refs, rides on F7 with.

### F4 — `format:merge` on `--workbook`

`--range A1:J1 [--type MERGE_ALL|MERGE_COLUMNS|MERGE_ROWS|--unmerge]
[--dryRun]`. ExcelJS merges whole ranges only; `MERGE_COLUMNS`/`MERGE_ROWS`
decompose into per-column/per-row merges in the adapter. Unmerge restores
individual cells (values in non-top-left cells are already lost — same as
Sheets semantics, documented). Effort: **S**.

### F5 — `workbook:write` sparse mode

New input form: `--cells '<json>'` accepting `{a1: value}` map or
`[{a1, value}]` list; each entry writes exactly one cell, everything else
untouched. Formula-target cells refuse without `--overwriteFormulas`.
`--dryRun` diffs per cell. This is the primitive that makes "fill columns
A,B,C,D,J only, skip formula columns" expressible without a script. Effort:
**S**.

### F6 — `data:clear` on `--workbook`

Bounded-range values clear; formula cells refuse without
`--overwriteFormulas`; `--dryRun` lists affected cells. Effort: **S**.

### F7 — `--updateRefs` formula-reference rewriting (the differentiator)

On `grid:insert`/`grid:delete` with `--workbook`: parse every formula in the
workbook, shift A1 references that point into/through the mutated region.

- Phase 1 (this spec): same-sheet cell refs and ranges (`A5`, `$B$2`,
  `A5:A20`), plus defined names. Absolute/relative markers preserved; refs
  entirely inside a deleted range become `#REF!` (Excel semantics).
- Phase 2 (follow-up): cross-sheet refs (`'Sheet 2'!A1`), structured/table
  refs — deferred, flagged in capabilities.
- Implementation: a small A1-ref tokenizer over formula strings (regex-class
  problem, not a full formula parser — we only rewrite ref tokens, never
  evaluate). New module `src/lib/xlsx-refs.ts`, pure functions, unit-testable
  in isolation.
- Without `--updateRefs`, behavior matches openpyxl/ExcelJS today: refs are
  left stale and the receipt warns `N formula(s) reference the mutated
  region`. Effort: **L** (the only L in the roadmap).

### F8 — `format:cells` on `--workbook`

Subset of cloud flags that ExcelJS supports: `--bold --italic --underline
--strikethrough --fontSize --fontFamily --textColor --backgroundColor
--horizontalAlignment --verticalAlignment --wrapText --numberFormat
--borders --borderStyle --borderColor --clear`. Unsupported flags
(`--wrapStrategy` values without an ExcelJS equivalent, `--numberFormatType`)
error `USAGE` on the local backend rather than silently no-op. Effort: **M**.

### F9 — `grid:freeze` / `grid:resize` / `grid:hide` on `--workbook`

Freeze → `ws.views = [{state:'frozen', xSplit, ySplit}]`; resize →
`col.width`/`row.height` (`--auto` computes from max cell text length, capped);
hide → `col.hidden`/`row.hidden`. All non-destructive, `--dryRun` prints the
planned model change. Effort: **S**.

### F10 — `worksheet:add|remove|rename` on `--workbook`

`add` creates an empty sheet (refuse duplicate title), `remove` refuses the
last visible sheet, `rename` updates `ws.name` and — phase 2 of F7 — offers
`--updateRefs` to rewrite `'Old'!` refs. `worksheet:copy` local = duplicate
within the same file (ExcelJS has no sheet copy; implement as cell+style
clone — **M**, or defer). Effort: **S** (copy deferred).

### F11 — `workbook:names` (defined-name management)

`list` (already in inspect), `add --name X --refersTo 'Sheet!$A$1:$A$9'`,
`remove`. Needed by template workflows that anchor formulas to named ranges.
Effort: **S**.

### F12 — Feature-loss consent gate

`save()` currently warns when preflight found unsupported features. Add
`--discardUnsupported` required to save such a workbook through any mutation
command; without it the command fails `CONFLICT` listing the features at
risk. Read-only commands unaffected. This converts a silent-fidelity-loss
footgun into an explicit decision — required before the CLI can claim
general-purpose safety. Effort: **S**.

## Critical files and ownership

| File | Change |
|---|---|
| `src/lib/xlsx.ts` | New adapter methods: `find`, `splice`, `merge`, `clearRange`, `setCells`, `setStyle`, `sheetOps`, `names`; inspect extension |
| `src/lib/xlsx-refs.ts` | **New** — A1-ref tokenizer + shift engine (F7), pure |
| `src/lib/xlsx-types.ts` | Types for all new options/results |
| `src/lib/xlsx-file.ts` | `--discardUnsupported` gate in save path |
| `src/commands/grid/*.ts`, `format/*.ts`, `worksheet/*.ts`, `data/{find,clear}.ts` | `--workbook` flag + backend dispatch (export-csv pattern) |
| `src/commands/workbook/{find,names}.ts` | New commands |
| `src/lib/capabilities.ts` | Local-backend operation list + fidelity exclusions |
| `test/` | New offline suites per feature; fixture workbooks under `test/fixtures/` |
| `skill/SKILL.md`, `docs/agents.md` | Sync after each wave |

Wave ownership: waves are sequential; within a wave, lib work and command
work can split across two workers (lib first — commands consume its API).

## Verification

- AC-01 `workbook:find` output shape is byte-compatible with `data:find`
  receipts (same keys, same coordinate semantics).
- AC-02 `grid:insert --workbook --dryRun` writes nothing; file sha256
  unchanged.
- AC-03 Sparse `--cells` write touches only listed cells; formula cells in
  the map refuse without `--overwriteFormulas`.
- AC-04 `grid:insert --updateRefs` on a fixture with `=SUM(A5:A20)` below the
  insertion point produces `=SUM(A5:A23)`; refs above unchanged; refs inside
  a deleted range become `#REF!`; defined names updated.
- AC-05 `format:merge --workbook` round-trips: merge → save → reload →
  `inspect` reports the range; unmerge restores.
- AC-06 Mutating a workbook with charts fails `CONFLICT` without
  `--discardUnsupported`; succeeds with it; receipt lists dropped features.
- AC-07 Every new mutation honors `--dryRun`, atomic save, `--redacted`, and
  the JSON error envelope — verified by reusing the existing offline command
  harness patterns.
- AC-08 `capabilities` output lists new local operations and the phase-2
  exclusions (cross-sheet refs, charts, VBA).
- AC-09 Full `npm run test:unit` green; new fixtures committed; no network in
  any new test.

## Execution checklist

Wave 1 — read/discovery (unblocks everything, all S):
- [x] T-01 F2 inspect extension (merges, formulas, validations, freeze) → AC-05(prep), AC-08
- [x] T-02 F1 `workbook:find` → AC-01

Wave 2 — structural mutation:
- [x] T-03 F3 `grid:insert/delete --workbook` (no refs yet) → AC-02
- [x] T-04 F4 `format:merge --workbook` → AC-05
- [x] T-05 F5 sparse `--cells` write → AC-03
- [x] T-06 F6 `data:clear --workbook` → AC-07
- [x] T-07 F12 `--discardUnsupported` gate → AC-06 (engine gate `allowUnsupportedFeatures` pre-existed; flag exposed on workbook:write, report:run, and every --workbook mutation)

Wave 3 — the differentiator:
- [x] T-08 F7 `xlsx-refs.ts` tokenizer + shift engine, unit-tested standalone → AC-04
- [x] T-09 Wire `--updateRefs` into grid:insert/delete → AC-04

Wave 4 — presentation & sheet management:
- [x] T-10 F8 `format:cells --workbook` subset → AC-07
- [x] T-11 F9 freeze/resize/hide → AC-07
- [x] T-12 F10 worksheet add/remove/rename → AC-07
- [x] T-13 F11 `workbook:names` → AC-08

Wave 5 — docs & closeout:
- [x] T-14 capabilities.ts + SKILL.md + agents.md sync → AC-08 (oclif docs regenerated via `npx oclif readme --multi`)
- [x] T-15 Full suite via test-runner, consolidated reviewer pass → AC-09 (809 passing; reviewer found 1 blocker + 4 minors, all repaired in one consolidated pass, suite re-verified green)

Deferred (phase 2, not in this plan): cross-sheet ref rewriting,
`worksheet:copy` local, data-validation/conditional-formatting write support,
cell comments, streaming mode for >50 MB workbooks.

## Evidence and handoff

- Engine audit: `src/lib/xlsx.ts` exposes only
  `load/create/inspect/read/preview/apply/save`; no merge/splice/sparse paths
  (grep: zero hits for `splice|mergeCells` under `src/lib` outside cloud
  `sheet-format.ts`).
- `data:export-csv` (`src/commands/data/export-csv.ts`) is the working
  dual-backend precedent: `-s` or `--workbook`, one flag set per backend.
- Preflight (`src/lib/xlsx-zip.ts`) already detects chart/pivot/macro/etc.
  into `XlsxUnsupportedFeatureType` — F12 only needs to gate on it.
- ExcelJS pinned at 4.4.0 (`package.json`), `uuid` override in place.
- Wave 1+2 implementation evidence (2026-09-23): `npm run test:unit` → 740
  passing, 0 failing; new suite `test/xlsx-ops.test.ts` (27 cases) covers
  find/inspect/splice/merge/setCells/clearRange incl. merge-conflict gate,
  forced merge adjust, dryRun no-mutation, formula-overwrite refusal, and
  merge persistence through save→reload. `test/commands/offline.test.ts`
  flag-surface table updated for `workbook:find` and the new `--workbook` /
  `--cells` / `--discardUnsupported` / `--force` flags. CLI smoke-tested
  end-to-end on a fixture workbook (insert, merge, sparse write, clear,
  forced delete through a merge).
- New files: `src/lib/xlsx-target.ts` (shared `--workbook` target resolution +
  save plumbing), `src/commands/workbook/find.ts`. New `XlsxWorkbook` methods:
  `find`, `splice`, `mergeCellsRange`, `setCells`, `clearRange`; `inspect()`
  gained `XlsxInspectOptions`.
- Ref semantics as shipped: without `--updateRefs` refs stay stale and every
  splice receipt reports `formulasAtRisk` + warning; with it, same-sheet refs
  + defined names are rewritten (`refsRewritten`/`refsBroken` in receipt).
  Data validations and conditional-formatting ranges are never adjusted.
- Wave 3+4 evidence (2026-09-24): `src/lib/xlsx-refs.ts` (pure tokenizer +
  shift engine + defined-name diff); `splice(updateRefs)` rewrites same-sheet
  refs + defined names, refs inside deleted span → #REF!, cross-sheet
  untouched; `--updateRefs` on grid:insert/delete (USAGE error with -s).
  Wave 4: `formatCells`, `freezePanes`, `resizeGrid`, `setGridHidden`,
  `addSheet`/`removeSheet`/`renameSheet`, `listDefinedNames`/`addDefinedName`/
  `removeDefinedName`; new command `workbook:names`; `--workbook` on
  format:cells, grid:freeze/resize/hide, worksheet:add|remove|rename.
- Reviewer pass (Tier-3) on `git diff e25b126`: 5 findings, all repaired —
  (1) BLOCKER shared-formula slave double-shift → pre-splice
  `_snapshotFormulaTexts` keyed by pre-splice address, rewrite re-resolves
  post-splice; (2) freezePanes merged into prior views[0] preserving
  unspecified axis + view attrs; (3) removeSheet drops defined names scoped
  to the removed sheet; (4) _autoColumnWidth iterates allocated rows via
  eachRow (sparse-row fix); (5) shiftSpan insert overflow → #REF!.
- Test totals: 809 passing, 0 failing (`npm run test:unit`, test-runner
  receipt). New suites: xlsx-refs.test.ts (33+), xlsx-format.test.ts (27+).
- ExcelJS caveats discovered: spliceRows/spliceColumns DO shift defined
  names internally (diff-based counting avoids double-shift) and rebuild
  moved rows as fresh cell objects (snapshot must key by address, not cell
  ref); definedNames.model getter returns fresh array (removeDefinedName
  edits internal matrixMap — verified persists); definedNames.add(loc,name)
  arg order; BorderStyle has no 'none'; freeze auto-adds topLeftCell.

## Assumptions and contingencies

- ExcelJS `spliceRows` shifts cell values and styles but not formula strings,
  merges below the splice point, or defined names — F7 must handle refs;
  merges intersecting the splice boundary need an explicit policy (spec:
  merges fully below shift down; merges intersecting the boundary are
  rejected with `CONFLICT` unless `--force`).
- If a target workbook proves to rely on charts/pivots, F12 makes the
  limitation explicit instead of silent — acceptable; the alternative engine
  swap is a separate decision.
- `--updateRefs` phase 1 deliberately excludes cross-sheet refs; shipping it
  without them is still strictly better than openpyxl's status quo.
- Risk: ExcelJS style-clone semantics for `--inheritFromBefore` are shallow
  in places (shared style objects) — verify with a fixture before promising
  style inheritance in the receipt.
