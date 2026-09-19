# Spreadsheet capability assessment and roadmap

Work ID: 2026-09-19-spreadsheet-roadmap
Status: Assessment complete; user subsequently approved implementation of the recommended priorities.
Next safe action: Follow the active implementation and verification record at [2026-09-19-report-automation.md](2026-09-19-report-automation.md). The assessment evidence and proposed scope below remain historical; this record does not certify implementation acceptance.

## Context
User requests an assessment of current google-sheet-cli capabilities, a roadmap for fast and comprehensive financial and manpower-estimation reporting, and a recommendation on supporting both Google Sheets and local Excel. Prior OAuth implementation is historical context; its old work record is stale and is not being resumed.
Non-goals: modifying application code, installing dependencies, accessing credentials or private spreadsheets, committing/pushing, building a complete accounting system.

## Approach
Separate current CLI/library behavior from underlying Google API capabilities and proposed additions. Evaluate bulk I/O, typed values, formulas/calculation, report templates, safe writes and local XLSX fidelity. Prefer a small shared report model with explicit backend capabilities, not a universal spreadsheet/calculation engine. Main synthesizes findings; one read-only scout supplies repository evidence. Consult primary vendor documentation for external constraints.

## Critical files and ownership
- Main sole owner: this work record and user-facing recommendation.
- SheetCapabilities scout (read-only): src/commands, src/lib/google-sheet.ts, base/auth flags, package.json and relevant existing tests/docs.
- No concurrent source writes; no application implementation wave.

## Verification
- AC-01: Current capability statements anchored to actual code/docs; no inferred support presented as implemented.
- AC-02: Financial and manpower examples distinguish data ingestion, calculations, validation, presentation and export.
- AC-03: Google Sheets vs local Excel recommendation covers formula evaluation, fidelity, safety and dependency tradeoffs using primary sources.
- AC-04: Roadmap is ordered by dependency/business value with observable acceptance examples, and explicitly proposed rather than delivered.
- AC-05: At least one read-only, credential-free real CLI probe verifies discovered surface. No test suite or build needed for this investigation.

## Execution checklist
- [x] T-01: Verify current capabilities and bounded CLI probes (AC-01, AC-05).
- [x] T-02: Analyze financial and effort-estimation workflows (AC-02).
- [x] T-03: Evaluate local Excel support and calculation boundaries (AC-03).
- [x] T-04: Publish prioritized roadmap and evidence (AC-04).

## Evidence and handoff
- Existing docs/plans/2026-09-18-oauth2-upgrade.md was read before exploration. Its unchecked implementation status is historical/stale; this assessment does not treat it as current behavior.
- Runtime evidence: `./bin/run.js --help && ./bin/run.js data:get --help && ./bin/run.js data:update --help` exited 0. Verified real CLI topics, read bounds/A1 range/header options, table CSV/JSON/YAML output, positional DATA nested-array argument, and `-v RAW|USER_ENTERED` defaulting to RAW. No authenticated calls, dependency installation, build, or test suite was run.
- Scout receipt: `agent://SheetCapabilities` inventories 12 domain commands, library methods, dependencies and absent exposed capabilities. Some scout line numbers remained stale; permanent code anchors below come from Main's current-disk range checks, not those stale numbers.
- Current disk, `src/lib/google-sheet.ts:193-234`: getData resolves worksheet metadata then calls values.get without a valueRenderOption. `--rawOutput` is JSON serialization, not UNFORMATTED_VALUE or FORMULA.
- Current disk, `src/lib/google-sheet.ts:309-335`: repeated per-row array spreads are a potential local allocation hotspot; appendData reads rows, sets minRow, then calls updateData. Concurrent read-then-write callers can target the same row; this is a structural risk, not a reproduced data-loss incident.
- Current disk, `src/lib/google-sheet.ts:374-403`: nested-array validation, empty-write short circuit, grid sizing and values.update with RAW default. Existing behavior must remain stable unless an explicitly versioned contract changes it.
- Current disk, `src/lib/google-sheet.ts:59-97`: existing quota-aware backoff is 3/12/48 seconds; do not propose adding retries as if none exist. POST replay after ambiguous network failures needs distinct safety semantics.
- Google primary sources: [usage limits](https://developers.google.com/workspace/sheets/api/limits), [batch requests](https://developers.google.com/workspace/sheets/api/guides/batch), [render options](https://developers.google.com/workspace/sheets/api/reference/rest/v4/ValueRenderOption). Standard read/write quota tables list 300/minute/project and 60/minute/user/project; effective project quota can differ. Recommended request payload is 2 MB, not a hard universal size limit. One batchUpdate is atomic; multiple chunks/files are not one transaction.
- ExcelJS primary README, [formula values](https://github.com/exceljs/exceljs#formula-value) and [streaming I/O](https://github.com/exceljs/exceljs#streaming-io): formulas can be stored but not calculated by ExcelJS; supplied/cached results are not freshness proof. Streaming reduces memory but committed rows cannot be revisited and unmerge is unsupported.
- Optional future engines: [LibreOffice headless/CLI](https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html); [HyperFormula license](https://hyperformula.handsontable.com/docs/guide/license-key.html) is GPLv3 or proprietary. No compatibility or performance benchmark was run for either.
- Current disk corroboration: metadata requests at `src/lib/google-sheet.ts:158-182`, optional additional header read at `244-253`, installed dependencies at `package.json:9-20`, supported exports at `41-59`, MIT package license at `73`.
- Acceptance: AC-01/05 satisfied by current-source evidence plus the credential-free CLI probes; AC-02/03/04 satisfied by the workflow examples, backend boundaries, primary sources and acceptance-oriented roadmap below. This verifies the assessment, not implementation or performance of proposed features.
- Residual limits: no current financial workbook inspection, data-loss reproduction, formula-engine evaluation or end-to-end report generation was attempted. Financial rules and template fidelity require representative approved fixtures before implementation.

## Assumptions and contingencies
- User requests advice, not immediate implementation or architecture approval by a separate advisor.
- Local Excel means primarily .xlsx; legacy .xls/.xlsb, VBA and perfect arbitrary-workbook round-trips require separate scope decisions.
- Performance targets must be measured with fixtures; no runtime speed claims without benchmarks.
- No private financial data is needed for the capability/roadmap assessment.

## Proposed roadmap — not implemented

### Product boundary
Evolve into a spreadsheet/report automation tool, not an Excel clone or accounting ERP. Keep the existing package, CLI and library entry points stable. Introduce a minimal typed report model and explicit Google Sheets/XLSX capabilities only as the second backend needs them. Separate deterministic business calculations from workbook formula execution and from rendering.

### Priority 0 — trustworthy data and writes
- Add explicit formatted/unformatted/formula read modes; preserve the current default for compatibility.
- Support JSON/CSV file and stdin input rather than forcing large inline JSON arguments.
- Specify dates, timezone, locale, money precision/rounding, empty/null/error behavior and identifiers with leading zeroes.
- Distinguish trusted formula fields from imported text; avoid formula injection.
- Add write preview/diff and deliberate overwrite behavior. Scope backups/checks to target ranges/files, not unrelated data.
- Acceptance: numeric values, display text and formulas are distinguishable; unknown or invalid financial inputs fail visibly; a literal leading `=` remains text unless explicitly authorized as a formula.

### Priority 1 — bulk operations and efficient execution
- Add multi-range values.batchGet/values.batchUpdate, bounded payload chunking, per-job metadata reuse with invalidation, minimal requested ranges/fields and bounded concurrency.
- Remove repeated copying in large read materialization and avoid generating human-facing representations when not requested.
- Define append semantics before changing them: native values.append is a candidate for contiguous table append, not an automatic drop-in replacement for the historical column/range behavior. Preserve public compatibility or release a deliberate breaking contract.
- Separate safe fixed-range retries from append with unknown commit state; no generic exactly-once claim or blind append replay.
- Acceptance: 12 bounded ranges fitting one batch need one batch values read plus at most one metadata fetch; controlled 10,000 x 20-cell fixtures record time, peak memory and API calls. Concurrent append acceptance must check no overwrite and document duplicate/uncertain-outcome handling.

### Priority 2 — local XLSX adapter
- Start with .xlsx: inspect tabs/ranges, read values/formulas/cache provenance, write data and controlled templates, preserve supported formatting, export a new file.
- Evaluate ExcelJS against representative customer fixtures before committing to it. Keep its dependency out of the existing Sheets-only import path.
- New-file output by default; in-place mode only with explicit opt-in, temp-file replacement, conflict check and backup policy. Local-only mode must never upload workbooks.
- Reject or warn clearly for unsupported workbook features; never silently strip macros, charts, external links or other features.
- Acceptance: offline fixture can be read/filled/exported without Google credentials or network access; supported formatting survives; unknown/stale formula results are reported, not labeled recalculated.

### Priority 3 — reusable report templates and business calculations
- Versioned declarative templates with input schema, column/named-range mapping, computed fields, validation, output formatting and source provenance.
- Shared operations: filter/map, group/sum, lookup/join, aggregate/pivot-shaped tables, duplicate detection and reconciliation. No arbitrary script eval in template config.
- Financial preset: ingest monthly transaction tabs, normalize accounts/currency/dates, reject duplicates/errors, compute receipts/payments and period balances, reconcile differences, render Google Sheets or XLSX. Preserve original amount, FX rate/date/source and converted amount.
- Do not equate receipts minus payments with accounting profit. Statutory statements require opening balances, complete journal/adjustments, account mapping and an approved accounting policy.
- Manpower preset: quantity x unit effort x approved adjustment, role rates, separately identified QA/PM/BA work, contingency, scenarios and module/role totals. Do not count included overhead twice.
- Example: 5 work items x 2 person-days x 1.2 complexity = 12 person-days; plus 20% contingency = 14.4 person-days; 2.5 million VND/person-day = 36 million VND. These are illustrative policy inputs, not inferred rates. Effort is not elapsed duration.
- Acceptance: identical normalized inputs and policy versions produce equal business totals for both outputs within declared rounding; template source hash/version and validation failures are retained; re-running the same report does not duplicate rows.

### Priority 4 — advanced workbook presentation and recalculation
- Add formatting, named ranges, freeze panes, validation and protected formulas as needed by report templates; chart/native pivot support remains backend-specific and capability-gated.
- Optional local recalculation via a controlled LibreOffice process; isolate profile and temporary files, disable macros/external links and constrain network/file access. Benchmark real template fidelity first. Excel-specific formulas may require Excel itself.
- Consider HyperFormula only if deterministic embedded recalculation is a demonstrated requirement and licensing/compatibility are acceptable; never implement a general Excel engine from scratch.
- Acceptance: known formula fixtures recompute; unsupported formulas return explicit errors; cached values are never silently treated as fresh; unrelated template content remains intact.

### Cross-cutting acceptance
- Golden report fixtures cover decimal money, FX, negative/refund entries, localized dates, leading-zero identifiers, formulas/errors, empty sheets, duplicate IDs and concurrent writes.
- Performance gates measure same fixture/environment before and after; no promised wall-clock speed-up without benchmarks.
- Google batch atomicity is limited to an individual request. Multi-batch writes require planned partial-failure recovery; optimistic checks do not eliminate simultaneous edit races.
- Accessibility: output tables have meaningful headers; totals/errors are labeled in text rather than color alone; formatted templates retain readable contrast.
- Keep Sheets-only consumers lightweight and existing exported behavior intentional. No automatic package rename, OAuth scope expansion, upload of private financial data, commit or push.
