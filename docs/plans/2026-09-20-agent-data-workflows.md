# Agent data workflows — approved implementation

Work ID: 2026-09-20-agent-data-workflows
Status: ACCEPTED — approved scope implemented, repaired and verified offline.
Next safe action: Publish the accepted implementation to origin/master under the user's subsequent explicit commit/push authorization; report the actual commit and remote receipt in the publication handoff. Live Google API verification remains outside scope.

## Context
The user approved the prioritized implementation chain, not the entire conditional roadmap: structured CLI errors; data:schema and data:validate; data:clear and key-based data:upsert; spreadsheet/worksheet copy; spreadsheet export to PDF/XLSX. Existing 30-command behavior and published library entry points must remain compatible. Historical roadmap records have been read; this is new work, not reopening completed milestones.
Non-goals: multi-profile/OAuth changes, new OAuth scopes, full Drive discovery, charts/pivots, protected ranges, general plan/apply, bidirectional sync, general Excel calculation, production spreadsheet mutations, commits or pushes.

## Approach
Reuse command factory, shared flags, A1 helpers, input parsing, report field validation, formula guards and existing authenticated API client. Keep output and typed operations consistent with existing command patterns. Do not introduce a second validation engine or dependency-heavy Drive SDK.

### Behavior contracts
- JSON error mode: opt-in machine-readable errors for raw-output/JSON invocations, stable error code/message/retry information and nonzero exit; retain existing human output outside that mode. Never serialize credentials, request/response internals or arbitrary error objects. Parser/auth/API/network failures must use the same safe boundary.
- Schema: read-only discovery of headers, actual column coordinates, inferred types and formula presence, relevant validation rules/named ranges. Report sample bounds/coverage, empty/duplicate headers and mixed types honestly. Dates/numbers inferred from samples are not authoritative schemas.
- Validation: reuse report schema/rules where applicable (required, type, enum, uniqueness); read-only, structured coordinate-linked issues and nonzero status for invalid data. No accounting-specific rules. Malformed schema is distinct from data violating a valid schema.
- Clear: explicit bounded range, values-only clearing preserves formatting; formula guard by default, explicit overwrite override, dry-run makes no mutation.
- Upsert: header-mapped record input using established JSON/CSV/stdin parsing; one explicit key column; reject empty/duplicate keys and ambiguous headers; preserve unspecified columns and formulas; RAW default. Type-aware keys distinguish string identifiers from numbers without stripping leading zeros. Dry-run reports added/updated/unchanged rows and planned ranges without writes. Re-running same input under single-writer conditions must not append duplicates. No blind replay after ambiguous mutation failures. No multi-writer/exactly-once/transaction claims. Pre-write stale-data checks, if used, only detect conflicts, not lock data.
- Copy: authenticated Drive files.copy for a spreadsheet; Sheets sheets.copyTo for one worksheet to an explicit destination spreadsheet; source data unchanged; return identifiers and metadata of actual copy. Do not duplicate sharing grants automatically or expand scopes.
- Export: authenticated Drive files.export using PDF or XLSX MIME type; binary-safe local output, no clobber without explicit opt-in, temporary-file publication/cleanup preserving old output on failure, machine-readable receipt without binary noise. Document Drive export's 10 MB limit and per-file drive.file authorization boundary.
- Atomicity: one supported Google batch request can be atomic; workflows spanning multiple requests/services/files are not a transaction.

## Critical files and ownership
Main owns this record exclusively and acceptance/integration decisions.
Locked slices and shared-core serialization:
- AgentContracts owns shared CLI error handling (base/auth/local command boundaries as needed), new schema/validation helpers and commands, and dedicated contract tests. Reuse normalizeTable(data, TableSchema, tableName) and ReportSchemaError from report/schema; no change to existing report semantics. It may add a minimal bounded schema metadata method only during its shared-core turn.
- SafeData owns new clear/upsert commands, their helper/types and dedicated tests. It has the first exclusive edit turn for src/lib/google-sheet.ts.
- CopyExport owns new spreadsheet/worksheet copy and export commands, output-file helper if needed and dedicated tests. It has the final exclusive edit turn for src/lib/google-sheet.ts.
- Shared-core handoff order: SafeData → AgentContracts → CopyExport, communicated explicitly through hub. Each owner re-grounds the latest file before edits and releases it immediately when done.
- CopyExport is sole writer to test/fake-sheets.ts; adds clear, copyTo, Drive copy/export and any bounded metadata support requested by peers. Other workers supply required fake contracts rather than editing it.
- Existing public exports remain compatible. New methods live on GoogleSheet, preserving its three entry points; no new broad internal export surface. New command implementations use the shared error boundary, not custom serialization.
- Main alone updates this record. Generated docs/changelog/skill changes occur after smoke proof, centrally. All workers skip validation/build/lint/tests while concurrent changes are in flight.

## Verification
- AC-01: JSON failures are parseable/safe/nonzero across parse/auth/API/network paths; ordinary output remains compatible.
- AC-02: Schema identifies offset A1 columns, duplicate/empty headers, mixed types, formulas and metadata with honest sampling bounds; no writes.
- AC-03: Standalone validation reuses report semantics, returns correct coordinates for violations, exits nonzero on invalid data and never writes.
- AC-04: Clear preserves format and refuses formulas by default; override works; dry-run produces zero mutations.
- AC-05: Upsert updates matching keys, adds new keys, preserves unspecified cells, refuses bad/duplicate keys and protected formulas, handles empty inputs, is stable on repeated single-writer invocation and does not mutate on preview/rejection.
- AC-06: Both copy commands call correct endpoints, return created metadata, preserve source and fail visibly on authorization/API errors without unsafe replay.
- AC-07: PDF/XLSX exports preserve exact bytes, use correct MIME types, protect existing files, leave no partial published output on failures, and report errors safely.
- AC-08: Offline suite/build and actual CLI smoke checks pass; generated command docs, existing agent skill and changelog describe new behavior and limits. Consolidated post-implementation review findings are fixed or explicitly blocked. No live remote mutation, commit or push.

## Execution checklist
- [x] T-01: Map architecture and lock ownership/contracts (AC-01–08).
- [x] T-02: Implement structured CLI errors (AC-01 verified).
- [x] T-03: Implement schema discovery (AC-02 verified).
- [x] T-04: Implement standalone validation (AC-03 verified).
- [x] T-05: Implement safe range clear (AC-04 verified).
- [x] T-06: Implement key-based upsert (AC-05 verified).
- [x] T-07: Implement spreadsheet and worksheet copy (AC-06 verified).
- [x] T-08: Implement PDF/XLSX export (AC-07 verified).
- [x] T-09: Integrate and verify offline suite/build/CLI (AC-01–08 verified).
- [x] T-10: Consolidated review, repairs and documentation handoff (AC-08 verified; one review, no re-review).
- [x] T-11: Post-smoke cleanup of task-owned throwaway artifacts and final evidence handoff (AC-08; added only after runtime smoke passed).

## Evidence and handoff
The entries below are chronological; the final acceptance subsection supersedes earlier pending/failed intermediate states.

- User selected “Chuỗi ưu tiên đã chốt” after explicit scope clarification.
- Existing docs/plans/2026-09-19-spreadsheet-roadmap.md read; old feature assessment does not certify this implementation.
- SafeData implementation receipt: agent://SafeData. Added clear/upsert library methods, src/lib/upsert-plan.ts, data:clear/upsert commands and test/clear-upsert.test.ts (26 authored behavioral cases, not yet executed). Bounded clear uses RAW empty-string writes via existing batch guard; upsert uses type-aware keys, preflight planning, changed-cell ranges and preserves omitted columns. Multi-chunk failures are not transactional; no concurrent-writer guarantee.
- Official contracts verified: https://developers.google.com/workspace/drive/api/reference/rest/v3/files/copy ; https://developers.google.com/workspace/drive/api/reference/rest/v3/files/export ; https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets.sheets/copyTo . Export returns bytes and is limited to 10 MB; copyTo returns SheetProperties.
- CopyExport implementation receipt: agent://CopyExport. Added three public library methods and CLI commands, extended fake API binary/copy routes and authored test/copy-export.test.ts. Binary publication reuses saveBufferAtomic. Worker ran an unauthorized mid-flight scoped check/build; that observation is NOT final acceptance (known compile errors in the other slice and stale lib output were reported). All subsequent verification is reserved for test-runner after source work settles.
- Implementation accepted as delivered code only; integrated verification remains pending.
- AgentContracts implementation receipt: agent://AgentContracts. Added src/lib/cli-errors.ts, src/lib/data-schema.ts, data:schema/validate, scoped getWorksheetMetadata, base error boundary and test/agent-contracts.test.ts (20 authored cases). Existing report normalizeTable reused; data errors map to absolute A1 coordinates. Seven-command inventory updated. Known limits pending acceptance: --json currently controls failures only, unknown commands bypass boundary, untyped local input Errors still classify INTERNAL.
- CopyExport corrected fake copyTo collision behavior and added same-spreadsheet copy coverage. Core copy response already uses server-assigned title.
- Source graph reconcile: xd://sot_reconcile paths src/bin returned 70 unchanged, 0 failed. Build/suite/CLI verification dispatched to WorkflowVerification; no accepted execution evidence yet.
- WorkflowVerification receipt (agent://WorkflowVerification): `npm run build` exit 1, nullable API metadata at src/lib/data-schema.ts:199–201,214. `npm run test:unit` exit 1; test/agent-contracts.test.ts:124,233 has invalid Error→Record casts. No passing full-gate claim.
- Bounded checks: clear/upsert 34 passed; copy/export 24 passed / 1 failed (`spreadsheet:copy` command EEXIT1 at test/copy-export.test.ts:295); existing output tests 10 passed. Offline inventory/schema/validate tests blocked by source type errors. CLI help discovered all seven new commands, but output came from a failed/partial build and is not final acceptance.
- Actual CLI parser-error smoke: --json, -r and --rawOutput produced parseable JSON stderr with exit 1; -j incorrectly produced human text. Treat this as a failure despite receipt's aggregate status field. Sole consolidated reviewer dispatched; source is frozen during review.
- Sole reviewer receipt: agent://WorkflowReview, overall incorrect pending repairs. Blocking/important findings: nullable schema and test casts; unquoted spaced title in oclif test; short -j; Drive wrapper status loss; bounded upsert can overwrite or duplicate unseen rows below read bound. Additional findings: plannedRanges ordering, stale copy help, OAuth/input error classification, large clear allocation, missing fake grid metadata coverage, unused import, failure-only --json documentation.
- One consolidated repair wave dispatched: CoreRepair owns google-sheet/upsert planning and clear/upsert tests; ContractRepair owns JSON/schema boundary and contract tests; TransportRepair owns fake/copy tests/copy help and real-client metadata tests. No shared concurrent writes. Clear will use native values.clear after formula preflight rather than materializing blank matrices. Explicit incomplete upsert ranges must refuse unseen trailing table data before any write.
- Reviewer established copy EEXIT1 was a test argv quoting error (--title with spaces), not an API copy defect. No live account was used; review diagnostic ran against the fake. No second review will be invoked; repairs will be checked through concrete regression/CLI evidence.
- CoreRepair receipt: agent://CoreRepair. Bounded upsert now scans trailing table columns through grid end with FORMULA and refuses any unseen populated cell before writes (cell address only, never cell content); plannedRanges follows execution order. Clear now uses one native values.clear after formula guard. Drive wrappers retain status/non-enumerable cause. New dependency-free ValidationError and cli-input metadata distinguish local input failures without importing oclif into ./sheet.
- ContractRepair receipt: agent://ContractRepair plus hub closure. Null normalization/test casts repaired; -j scanner accounts for option-value slots and -- passthrough; local OAuth failures distinguish auth from transport. refreshTokens now retains status/code/non-enumerable cause, preserving legacy message/behavior. JSON failures remain opt-in; --rawOutput is required for JSON success output.
- TransportRepair receipt: agent://TransportRepair. Spaced-title test correctly quoted; fake includes native clear and scoped grid metadata/named ranges; test/schema-transport.test.ts exercises real GoogleSheet client and command flow. Drive status envelope cases and native clear/upsert safety regressions authored. All repair workers skipped build/suite execution; final verification dispatched to RepairVerification.
- RepairVerification: `npm run build` exit 0. Full offline output artifact://468 confirms **559 passing, 1 failing** (initial compact receipt incorrectly said 189; corrected by runner and checked against raw summary). Remaining assertion at test/agent-contracts.test.ts:600 expects DUPLICATE_VALUE; ValidationClosure assigned to fix the actual contract/test defect, not weaken uniqueness validation.
- Actual ./bin/run.js smoke passed: help for all seven commands; --json/-j/-r/--rawOutput parser failures have empty stdout, one parseable stderr envelope, nonzero exit; ordinary human output preserved. Mock-factory subprocess exercised clear/upsert/spreadsheet copy with spaced title/worksheet copy/export with exact byte match. ./sheet import loaded zero @oclif modules. Live Google endpoints were not invoked.
- Smoke proof enables final documentation cleanup. WorkflowDocs owns existing README manual sections/CHANGELOG/skill plus ~/.omp skill sync; generated command docs remain reserved for generation without git staging. No new roadmap scope introduced.
- ValidationClosure root cause: uniqueness checks ran only after successful normalization, so any row-level error hid duplicates elsewhere. Extracted normalizeTableWithErrors result without duplicating validation; original normalizeTable API still throws identically for existing report callers. data:validate now aggregates row issues and uniqueness with source-row alignment. The previously failing assertion remains unchanged; no test weakening.
- WorkflowDocs handoff: README manual feature table/examples now cover 37 commands; CHANGELOG Unreleased section added without fabricated release version/date; repository/global google-sheet skill synced byte-identically. Generated README blocks/docs were untouched by the worker. No source/manual-doc edits remain in flight.
- FinalAcceptance dispatched for build, `./node_modules/.bin/oclif readme --multi && sh ./bin/clean.sh` (without git staging), full offline suite, actual CLI/mocked subprocess proof and throwaway cleanup.

### Final acceptance
- `npm run build` — exit 0.
- `./node_modules/.bin/oclif readme --multi && sh ./bin/clean.sh` — exit 0; README generated blocks and command documentation regenerated without git staging. New command documentation is in docs/data.md, docs/spreadsheet.md and docs/worksheet.md.
- `npm run test:unit` — exit 0, **560 passing, 0 failing, 0 skipped** (30s). Raw output: artifact://478; final summary independently read by Main. Receipt: agent://FinalAcceptance.
- AC-01: central error tests and actual CLI parser/local-validation smoke verify stable safe envelopes, nonzero exits, empty success stdout on failures, -j/--json/-r/--rawOutput behavior, and preserved human output. Drive/OAuth transport status propagation and local VALIDATION classification are covered.
- AC-02/03: test/agent-contracts.test.ts and test/schema-transport.test.ts cover sampled schema/offset coordinates/formulas/named ranges/validation metadata; real client with fake transport; malformed schemas and simultaneous required/type/unique failures. Original report normalization behavior remains covered by report-core tests.
- AC-04/05: test/clear-upsert.test.ts covers native values.clear, preservation/formula guard/preview, typed keys, omitted-column preservation, repeated single-writer upsert, planned order, and rejection of unseen trailing data before mutation. Mocked external CLI success probes passed in RepairVerification.
- AC-06/07: test/copy-export.test.ts covers spreadsheet/worksheet copy, source preservation, title collisions and spaces, Drive error codes/no copy replay, binary PDF/XLSX bytes, no-clobber/overwrite and failure protection. Mocked external CLI copy/export probes passed; no live service calls.
- AC-08: all gates above pass; ./sheet module-cache probe loads zero @oclif modules. README manual examples/37-command feature table, CHANGELOG Unreleased and repository/global skill are updated; skill files compare identical. Source graph final reconciliation: 100 unchanged, 0 failed.
- Review closure: all enumerated reviewer defects addressed in the consolidated repair and subsequent validation defect closure. No second review invoked. Evidence is regression/build/runtime results, not an assertion of independent re-approval.
- Cleanup: WorkflowVerification, RepairVerification, WorkflowReview and FinalAcceptance confirmed no task-created probes/preloaders/test exports remain. Existing user files and npm cache preserved. No commit, push or git staging performed.
- Residual boundaries: verification is offline/fake transport, not live Google API/OAuth certification. Upsert is single-writer only; formula preflight is not a lock and multi-request operations are not transactional. Schema/validate inspect declared sample bounds (defaults 100 rows × 26 columns). --json/-j controls failures only; --rawOutput enables JSON success too. Drive export remains subject to server 10 MB limit and per-file drive.file access. Unknown command/topic errors before a command class loads still use the oclif human boundary.

### Authorized publication follow-up
- User explicitly requested updating the changelog/README, committing and pushing the completed implementation.
- Destination verified: master tracks origin/master; origin is https://github.com/minhgv/google-sheet-cli.
- README/CHANGELOG publication pass clarifies 37-command scope, sampled validation/combined uniqueness errors, server-assigned copy metadata and the pre-command JSON-error boundary. Existing generated command documentation remains generator-owned.
- The earlier no-commit/no-push statements describe implementation acceptance before this new authorization. No live credentials or Google spreadsheet mutations are authorized.

## Assumptions and contingencies
- Existing Service Account and OAuth clients remain the authentication seam; no credentials outside workspace/~/.omp may be read.
- No live Google account is required for offline proof. Report lack of live verification explicitly.
- Single-writer upsert only; Google Sheets is not a database with unique constraints or compare-and-swap cell writes.
- Prefer focused permanent tests for actual boundary/safety risks; use actual CLI smoke for surface proof.
- Generated docs must use the repository oclif generation path without inadvertent git staging; never hand-edit generated blocks.
- Architecture receipt: agent://ImplementationMap. GoogleSheet Sheets/auth/Drive transport are private; public getData/getDataBatch/updateDataBatch/getSpreadsheet and batchUpdateSpreadsheet are reusable. Metadata for cell-level validation must explicitly request grid data if the current getSpreadsheet response omits it; do not assume absence of data means absence of rules.
- Reuse seams: src/lib/factory.ts createGoogleSheet; src/lib/cli-input.ts resolveDataMatrix; src/lib/report/input.ts parseInput; src/lib/report/schema.ts normalizeTable and ReportSchemaError; src/lib/report/types.ts TableSchema/FieldSchema/ValidationError.
- If graph/LSP is incomplete, use bounded AST/range evidence rather than infer absent features from a single no-match.
