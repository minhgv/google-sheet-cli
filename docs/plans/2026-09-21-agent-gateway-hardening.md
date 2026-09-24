# Agent gateway hardening — approved blueprint

Status: DELIVERED — all waves complete, consolidated review repaired, suite green.
Next safe action: none; commit only on explicit user authorization.

## Context

Problem: the 37-command gateway has substantial reporting/mutation capability, but agents lack explicit recovery semantics, resource discovery and backend capability contracts. Diagnostic and credential persistence paths carry avoidable confidentiality risks. Local workbook capability is broader than command count suggests and needs accurate documentation, not reimplementation.

Objective: make existing operations safer to automate; add scope-honest spreadsheet discovery and deterministic CSV interchange. Preserve compatibility for the published library and the gsheet.action consumer (~80 workflows).

Evidence (advisor-verified anchors):
- `updateDataBatch` grows grids and runs sequential values requests, then reports completed ranges inside Error text (`src/lib/google-sheet.ts:958-1166`).
- Workbook ReportDocument input already supports sheet creation, formula objects, number formats, freeze rows, column widths (`src/commands/workbook/write.ts:233-280`; `src/lib/xlsx.ts:537-686`).
- Credential persistence lacks restrictive modes and atomic replacement (`src/lib/oauth.ts:40-44,73-77` — `writeFileSync(TOKEN_PATH)` direct write).
- Local persistence uses a 0644 temp file (`src/lib/xlsx-file.ts:128`); `spreadsheet:export` already reuses `saveBufferAtomic` (`src/commands/spreadsheet/export.ts:5,63`).
- Shared error envelope/serializer lives in `src/lib/cli-errors.ts` (`GSheetError`, `classifyError`, `toErrorEnvelope`, `sanitizeMessage`).

Non-goals: transactions, exactly-once writes, concurrent-writer support, general Excel formatting parity, a formula calculation engine, charts/pivots, native XLSX→Sheets conversion, cross-spreadsheet orchestration, broader OAuth scopes, deployment, commit or push. No new worksheet lifecycle engine for XLSX beyond documenting existing bounded editing.

## Approach

### Architectural boundaries
- Keep `GoogleSheet` as the remote boundary, `XlsxWorkbook` as the local adapter, `ReportDocument` as the reporting contract. Shared pure contracts/serializers only where semantics are genuinely common. No fake backend-parity interface.

### Mutation outcomes (additive, compatible)
- Typed mutation detail rides on existing `Error`/`GSheetError` instances; messages and causes unchanged for legacy consumers. Structured properties added, never replacing Error semantics.
- Model outcomes at physical request/chunk level, retaining logical input ranges: `acknowledged` / `rejected` (only when response proves non-application) / `unknown` (dispatch may have succeeded) / `not-attempted`. Grid-growth effects tracked separately from cell updates. Include operation phase, cause, conservative retry guidance. No rollback promise; never auto-label the failing request uncommitted.
- Retry metadata describes the operation (read / deterministic overwrite / non-idempotent append-create-copy); unknown outcome never becomes an unconditional replay instruction. Existing request retry scheduling unchanged.

### Versioned capability contract
- Deterministic machine-readable capability document, own schema version independent of package version. Per backend: supported operations, I/O forms, formula-write vs recalculation, destructive guards, mutation/retry limits, local fidelity exclusions. Static support separated from runtime authorization. Accessible credential-free; reused for public docs.

### Scope-honest discovery
- `spreadsheet:list` (name search) over Drive `files.list` with spreadsheet MIME filter, bounded page size, opaque continuation token, explicit exact-vs-partial matching, controlled query builder (no arbitrary interpolation). Duplicate names stay separate candidates; never auto-select a write target. Responses disclose `drive.file` visibility boundary; empty result ≠ file absent. Old tokens lacking `drive.file` → actionable scope error, no interactive reauth.

### CSV interchange + bounded local editing
- Shared deterministic CSV serializer fed by existing remote/local read adapters; explicit worksheet or bounded range only (no implicit whole-workbook). Covers BOTH backends (Google Sheets + local XLSX) — scope confirmed.
- Render modes: raw/unformatted, formatted, formula text; default machine-consumable. XLSX cached formula results stay cached, never recalculated. Unsupported mode → refuse, never silently substitute.
- Deterministic delimiter/encoding/newline/quoting/blank/rectangular/header rules; no invented headers, no identifier coercion, no silent rounding. Data bytes and JSON receipts stay distinct channels.
- Formula-injection policy explicit: `safe` (neutralize dangerous string prefixes, keep typed negatives numeric) vs `preserve` (byte fidelity + warning). Receipt records selected policy and any transformed cells.
- Keep existing unsupported-feature guards, hash conflict checks, overwrite consent; never advertise as locking/lossless editing.

### Confidentiality boundary
- Token persistence: same-directory atomic replacement, restrictive dir/file modes, symlink handling defined, existing permissive artifacts handled; token format/location unchanged.
- Restrictive temp-file mode for generated exports (XLSX writer, new CSV writer, PDF/XLSX export path via `saveBufferAtomic`); final rename must not widen access. Atomic rename ≠ fsync durability ≠ locking — document honestly.
- Opt-in redacted-diagnostics mode: coordinates/counts/status/conflict-type retained; cell contents, formulas, incoming values, credentials stripped before serialization/logging, including nested `cause`. Normal data reads/exports stay data-bearing.

## Critical files and ownership

| Boundary | Files | Owner rule |
|---|---|---|
| Shared contracts (mutation outcome types, capability schema, CSV serializer, redaction policy) | `src/lib/cli-errors.ts`, new `src/lib/mutation-outcome.ts`, `src/lib/capabilities.ts`, `src/lib/csv.ts` | ONE owner across all waves |
| Remote ops | `src/lib/google-sheet.ts` (updateDataBatch, Drive auth, export) | Mutation + discovery serial, same boundary |
| Credentials | `src/lib/oauth.ts` | Independent of google-sheet.ts |
| Local persistence | `src/lib/xlsx-file.ts` (saveBufferAtomic), `src/lib/xlsx.ts` (read/apply — do not duplicate existing features) | One owner |
| New commands | `src/commands/spreadsheet/list.ts`, `src/commands/data/export-csv.ts` (names finalized at implementation) | Existing factory/output conventions |
| Test seams | `test/fake-sheets.ts`, `test/commands/offline.test.ts` | Shared — serialize edits, single owner per wave |
| Docs | generated `docs/*.md` + README blocks via `npm run version` only; `skill/SKILL.md` + `~/.omp` sync | Never hand-edit generated blocks |
| This record | `docs/plans/2026-09-21-agent-gateway-hardening.md` | Main only |

## Verification

- **AC-01 additive compatibility**: public imports resolvable; legacy success shapes/Error behavior unchanged; existing offline suite green; new structured properties observable without message parsing.
- **AC-02 precise partial outcomes**: offline faults cover reject-before-write, failure after acknowledged chunk, grid-growth-then-values-failure, lost response after possible commit; receipts distinguish acknowledged/rejected/unknown/not-attempted incl. chunked/overlapping ranges; no rollback or exactly-once claim.
- **AC-03 conservative retry guidance**: reads vs non-idempotent mutations classified distinctly; unknown append/create/copy never recommends blind replay; existing retry dispatch unchanged.
- **AC-04 truthful capability doc**: credential-free static response validates against versioned schema; both backends identified; XLSX formula-write vs recalculation distinguished; authorization-dependent visibility explicit.
- **AC-05 bounded discovery**: offline tests cover pagination, exact/partial names, escaped chars, duplicate titles, empty results, Drive errors, insufficient scope; IDs retained; `drive.file` boundary disclosed; no scope expansion.
- **AC-06 deterministic CSV**: both backends export selected worksheet/range with documented rendering/escaping; fixtures cover quotes, delimiters, multiline Unicode, blanks, identifiers, numerics, formula/cache; unsatisfiable mode refuses.
- **AC-07 explicit injection policy**: safe vs preserve observably different for dangerous strings, negative numbers, formula-text export; receipt names policy; never presented as lossless.
- **AC-08 atomic confidential artifacts**: isolated fs tests prove restrictive POSIX modes on new token/export files, atomic replacement on failure, compatible token loading, no leaked temp files; existing files not made more permissive; platform limits reported.
- **AC-09 redaction**: canary cell values/formulas/incoming data/token strings absent from stdout/stderr/envelope/dry-run when enabled; coordinates/counts/outcomes intact; data commands unaffected.
- **AC-10 accurate XLSX contract**: docs demonstrate existing ReportDocument sheet-add/formula/presentation without claiming new work; unsupported-feature refusal and cached-formula limits visible.
- **AC-11 integration evidence**: final offline suite via test-runner, built-CLI exercise of new surfaces, one consolidated review; live-service limits recorded, never inferred from FakeSheets.

## Execution checklist

### Wave A — shared contracts + confidentiality (parallel, ≤3)
- [x] T-01 Mutation outcome vocabulary + retry metadata types in `src/lib/mutation-outcome.ts`; extend `GSheetError`/`cli-errors.ts` additively (AC-01, AC-02, AC-03)
- [x] T-02 Capability schema `src/lib/capabilities.ts` + credential-free `capabilities` surface (AC-04)
- [x] T-03 Confidential artifacts: atomic+restrictive token write in `oauth.ts`; restrictive temp mode in `xlsx-file.ts` (AC-08)
- [x] T-04 Redacted-diagnostics mode in `cli-errors.ts` + flag plumbing in `base-class.ts` (AC-09)

### Wave B — remote behavior + discovery (serial on google-sheet.ts)
- [x] T-05 Typed partial outcomes in `updateDataBatch` + grid-growth reporting (AC-02, AC-03)
- [x] T-06 `spreadsheet:list` Drive discovery: query builder, pagination, visibility disclosure (AC-05)

### Wave C — CSV + docs (parallel where seams free)
- [x] T-07 CSV serializer `src/lib/csv.ts` + injection policy + render modes (AC-06, AC-07)
- [x] T-08 `data:export-csv` command for Sheets + XLSX adapters (AC-06, AC-07)
- [x] T-09 Capability doc published; XLSX ReportDocument features documented accurately (AC-04, AC-10)

### Wave D — integrated acceptance
- [x] T-10 Full offline suite + built-CLI smoke; consolidated review + repair pass (AC-01..AC-11)

## Evidence and handoff

- Advisor blueprint: `agent://AdvisorConsult` (consultation mode; T-XX checklist added by Main per role boundary).
- Prior accepted work: `docs/plans/2026-09-20-agent-data-workflows.md` (37-command baseline, commit `ae0b443` on origin/master).
- Tests: 712 passing offline (was 429 at plan start). New modules: `mutation-outcome.ts`, `csv.ts`, `capabilities.ts`; new commands: `capabilities`, `spreadsheet:list`, `data:export-csv`; new test files: mutation-outcome (28), mutation-outcome-batch (11), capabilities (19), confidential-artifacts (9), csv (46), spreadsheet-list (17), export-csv (19).
- Consolidated review (reviewer, Tier-3): 8 findings, all repaired — confidential-artifacts isolation fixed via `GSHEET_CONFIG_DIR` env override (documented, docs/agents.md rule 3); `redactedMessage` on GSheetError + ValidationError wired into both formula-conflict throws; `redactDiagnostic` plumbed into `logRaw` so dry-run receipts redact; `setRedactionEnabled` per-command authoritative; `MANAGED_ENV` extended; `--all` 1000-bound truncation test added; upsert failures labeled `operation: 'upsert'`.
- CLI smoke: `capabilities` valid JSON credential-free; `spreadsheet:list --help`, `data:export-csv --help`, `--redacted` visible on `data:get --help`.
- Residual risks (accepted): remote edits non-transactional; `unknown` outcomes irreducible; TOCTOU on local hash checks; FakeSheets cannot certify real OAuth/Drive visibility; `drive.file` discovery intentionally incomplete.

## Assumptions and contingencies

- Node 22+, oclif 5, current GoogleSheet public contract, single-writer semantics fixed.
- FakeSheets is the PR proof boundary; extend it for discovery/failure behavior but never claim real Google auth semantics from it.
- CSV scope confirmed as BOTH backends (Sheets + XLSX).
- New command filenames finalized at implementation after targeted discovery; no second conventions.
- `drive.file` retained; unrestricted Drive discovery is a separate authorization decision.
- Some transport failures are irreducibly `unknown` — conservative uncertainty is correct, not a bug.
- POSIX restrictive modes have platform equivalents/limits; report, don't overclaim.
- Redaction is opt-in for compatibility; unattended deployments must enable it.
- No commit/push/deploy granted by this plan.
