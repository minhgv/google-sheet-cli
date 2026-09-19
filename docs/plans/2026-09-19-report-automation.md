# Kế hoạch triển khai spreadsheet/report automation

Work ID: 2026-09-19-report-automation
Status: ACCEPTED — P0–P3 và phần định dạng template cần thiết hoàn tất; official gates và smoke/benchmark đạt trong phạm vi offline/synthetic.
Next safe action: Dùng examples local hoặc mở task riêng cho template thực tế/live Google/P4 nâng cao. Không commit/push hoặc mở rộng quyền truy cập nếu chưa được yêu cầu rõ ràng.
Approved source: docs/plans/2026-09-19-spreadsheet-roadmap.md. Giữ nguyên hồ sơ đánh giá đã hoàn thành; đây là hồ sơ thực thi mới.

## Context
Người dùng yêu cầu viết roadmap thành kế hoạch Markdown và thực hiện đề xuất. Phạm vi đã đề xuất: P0 dữ liệu/an toàn, P1 batch/hiệu năng, P2 XLSX local, P3 template và hai báo cáo nghiệp vụ; P4 chỉ đưa vào các khả năng trình bày cần cho template thực tế, không mặc định xây engine công thức tổng quát.
Giữ nguyên CLI, library entry points và hành vi tương thích 2.x/3.x đã được kiểm thử. Không thay bộ số liệu thật hoặc credentials, không commit/push khi chưa được yêu cầu cho lần thay đổi này.
Non-goals: statutory accounting certification, Excel clone, VBA/Power Query, .xls/.xlsb, bảo toàn mọi workbook tùy ý, upload file local để tính, engine công thức tổng quát/LibreOffice runtime chưa được yêu cầu và chưa có template cần nó. Không coi kết quả formula cache là vừa được tính lại.

## Approach
- Giữ GoogleSheet là API Google hiện tại; thêm read render modes, batch API và native table append riêng thay vì đổi ngầm appendData cũ.
- Tạo report core độc lập nhà cung cấp: typed schema, input JSON/CSV/stdin, decimal arithmetic, safe declarative transforms, financial/manpower report generation và provenance.
- Tạo XLSX adapter riêng, không import ExcelJS qua đường google-sheet-cli/sheet. Tạo file mới mặc định; ghi đè rõ ràng, kiểm tra xung đột và backup.
- CLI hiện có nhận thêm input/render/preview; commands workbook local và report dùng plain Command để không khởi động OAuth ngoài ý muốn. Google report output lấy client qua factory.
- Templates JSON version 1 khai báo schema, mapping/steps và policy. Không eval JavaScript, không tự suy đoán định mức nhân lực hay chính sách kế toán.
- Formula là kiểu riêng, không suy đoán từ chuỗi bắt đầu bằng '='. Tất cả text không tin cậy giữ literal; formula rõ ràng có kết quả đã tính theo report rule khi có thể, workbook cache khác được đánh dấu unknown freshness.
- Finance: quản trị thu/chi, FX, số dư, đối soát; không gọi dòng tiền là lợi nhuận. Manpower: quantity x unit effort x complexity, role rates, contingency/scenarios; effort không phải duration.
- Validation, formatting, readable headers và formula protection cần cho hai template thuộc phạm vi tích hợp; native charts/pivots và general recalculation giữ P4 có điều kiện.

## Critical files and ownership
Main duy nhất ghi hồ sơ này. Agents trả receipt, không sửa docs/plans.

### Wave A — backend/core độc lập
- GoogleData owner: src/lib/google-sheet.ts; existing types declarations; test/google-sheet-bulk.test.ts. Không sửa CLI/base-class/fake-sheets/package.
- ReportCore owner: src/lib/report/types.ts, input.ts, schema.ts, transform.ts, finance.ts, manpower.ts, index.ts; test/report-core.test.ts; examples/reports/*.json, *.csv. Không sửa runner/CLI/package/README.
- XlsxBackend owner: src/lib/xlsx.ts và module helper riêng nếu cần; test/xlsx.test.ts. Không sửa shared report types/package/CLI.
- Main cài dependencies và giữ quyền package.json/package-lock.json trong wave A.

### Wave B — tích hợp sau wave A
- Integration owner: src/commands/data, src/commands/workbook, src/commands/report; src/lib/base-class.ts; src/lib/report/runner.ts. Chỉ gọi API contract wave A.
- Verification fixtures owner: test/fake-sheets.ts, test/commands/offline.test.ts, các integration tests mới và smoke/benchmark script trong tmp/; sửa isolation OAuth tests nếu cần. Không chạy suite trong khi workers đang sửa.
- Documentation/package owner: README phần thủ công ngoài usage/commands, package exports/scripts, src entrypoints riêng nếu cần. Không tự sửa docs/*.md generated. Có thể bổ sung hướng dẫn và examples theo request.

### Wave C — verify/review/repair
Test-runner duy nhất chạy build/unit/full offline-capable suite, actual CLI scenarios và controlled benchmark. Một Reviewer pass sau implement. Nếu có lỗi, một consolidated repair pass, rồi verify phần ảnh hưởng; không re-review.

## Shared integration contracts
Các cấu trúc dưới đây là đặc tả giao tiếp, không phải implementation. Worker phải giữ tên/shape hoặc hỏi Main trước khi đổi.
- ReportScalar: string, number, boolean hoặc null. ReportFormula: object formula (không có dấu '=' ở đầu), result tùy chọn là ReportScalar. ReportCell là hợp của hai kiểu.
- ReportSheet: name, rows (ReportCell[][] gồm header nếu có), freezeRows tùy chọn, columnWidths tùy chọn, numberFormats tùy chọn (mảng column 1-based và format). ReportDocument: sheets và provenance chứa templateId, templateVersion, sourceHash.
- ReportSheet optional mapping bổ sung đã khóa giữa workers: startCell, targetNamedRange, clearManagedRange, metadata. Provenance có thể bổ sung generatedAt/summary; sourceHash phải ổn định theo input, timestamp không được làm thay đổi kết quả tính.
- readInput(options): Promise<unknown>; options gồm data? (inline JSON), file? ('-' đọc stdin), format? ('json'/'csv'). Chỉ một nguồn; JSON bảo toàn kiểu; CSV không tự biến leading-zero IDs thành số. Export parseInput(content, format).
- buildReport(template: unknown, input: unknown): ReportDocument, đồng bộ. Runtime validation đầy đủ, lỗi cụ thể theo dòng/cột. Input report là records hoặc matrix theo schema có header. Template id, version:1, kind:table/finance/manpower; config rõ ràng, policy không ẩn.
- GoogleSheet.getData giữ mặc định cũ; QueryOptions bổ sung valueRenderOption và dateTimeRenderOption. Tối ưu mảng không đổi output.
- GoogleSheet.getDataBatch(ranges:string[], options?, spreadsheetId?): trả mảng theo thứ tự input {range, values}; render defaults giữ tương thích. Chunk có giới hạn, sequential concurrency mặc định; không trả một phần như thành công khi lỗi.
- GoogleSheet.updateDataBatch(updates:{range,values}[], options?, spreadsheetId?): receipt lô/range đã áp dụng; options valueInputOption?, dryRun?, overwriteFormulas?; inspect-before-write bảo vệ formula khi không opt-in. Validate toàn input trước ghi; batch lỗi giữa chừng báo partial outcome rõ ràng. Giới hạn payload ~2MB, xử lý vùng lớn bằng chia dòng không thay semantics.
- GoogleSheet.appendTableData(data, options, spreadsheetId?): native values.append cho contiguous table; options worksheetTitle/range và valueInputOption, trả range đã cập nhật; không blind-retry POST. appendData cũ không đổi contract.
- GoogleSheet.applyReport(document, options?, spreadsheetId?) được bổ sung trong Wave A, import ReportDocument chỉ ở type-level. Ghi mixed scalar/formula bằng typed updateCells, không parse mọi chuỗi qua USER_ENTERED. Tạo tab thiếu, áp dụng format cần thiết, preview bảo vệ vùng có dữ liệu/formula không thuộc report. Metadata đánh dấu template và managed extent để rerun ngắn hơn chỉ xóa phần output do report sở hữu; chunk failure và concurrency limits phải được báo rõ.
- XlsxWorkbook.load(path): Promise<XlsxWorkbook>; constructor/new tạo workbook trống. inspect(): metadata/capabilities/warnings; read(range, mode?): matrix và thông tin formula freshness (mode formatted/unformatted/formula); apply(document, options?): preview/diff và cập nhật in-memory, mặc định bảo vệ formula đã tồn tại; save(path, options?): ghi file an toàn với overwrite?, inPlace?, backup?, expectedHash?; implementer gửi signatures chính xác cho integration trước wave B.
- Tên method XLSX có thể điều chỉnh một lần trong receipt wave A trước khi wave B bắt đầu; cấu trúc ReportDocument cố định. Named-range mapping phải được giải quyết rõ ràng trong template/apply, không thay đổi vùng ngoài quyền ghi.
- Không dùng any mới; type assertion phải có lý do. Phụ thuộc mới dự kiến: exceljs@4.4.0, decimal.js@10.6.0, csv-parse@7.0.2, jszip@3.10.1 (ZIP feature preflight). Chỉ Main/package owner thay dependency.

### Wave B JIT contract — CLI surface
- Existing data:get thêm --valueRenderOption/--dateTimeRenderOption, mặc định không đổi. data:update/data:append giữ positional DATA nhưng nhận thêm --input file hoặc '-' và --inputFormat json/csv, mutually exclusive. Preview/formula overwrite flags chỉ mở rộng có chủ ý; legacy default không đổi âm thầm.
- New data:batch-get dùng --ranges JSON array và spreadsheet auth flags; data:batch-update dùng inline/file/stdin array {range,values}; data:append-table dùng native append contract riêng.
- New workbook:inspect/read/write là plain Command, nhận --file và --range; write dùng input hoặc report document, --output, --dryRun, --overwrite, --overwriteFormulas và explicit --inPlace. Không OAuth/network trong local path.
- New report:run nhận --template JSON; input từ đúng một trong --input (file hoặc '-'), --sourceWorkbook với --ranges, --sourceSpreadsheet với --ranges. Output đúng một --output file.xlsx hoặc --spreadsheetId target Google. Template --workbookTemplate tùy chọn giữ format local có sẵn. --dryRun, --overwrite, --overwriteFormulas là rõ ràng. Plain Command chỉ authenticate khi cloud input/output được chọn; luôn factory.createGoogleSheet().
- Source spreadsheet đọc batch UNFORMATTED_VALUE và schema xử lý ngày/decimal; không ghép header tháng như data row. Source workbook đọc giá trị và từ chối dùng formula cache không rõ freshness để tính báo cáo nếu người dùng chưa opt-in rõ ràng.
- CLI file/stdin errors, malformed template/ranges và mutually exclusive source/output phải fail trước network/write. Machine-readable output không trộn spinner/log vào stdout. Help và generated docs theo oclif.
- Report/XLSX library exports nằm subpath riêng, không re-export eagerly từ root hoặc ./sheet; giữ dependency graph Google-only nhẹ.

## Verification
- AC-01 Compatibility: existing offline suite vẫn pass, public methods/exports cũ không mất; tests chạy với HOME trong workspace và không live credentials.
- AC-02 Data: formatted/unformatted/formula khác nhau; CSV quoted commas/newlines, UTF-8, leading zero, ngày/decimal/null/error được xử lý xác định; text công thức không bị thực thi ngầm.
- AC-03 Safe writes: preview không mutate; không ghi đè formula/file vô ý; overwrite có chủ ý, in-place conflict/backup; unsupported XLSX features không bị loại âm thầm.
- AC-04 Batch/performance: 12 bounded ranges bằng một values batch read; metadata không bị đọc lại mỗi range. 10,000 x 20 fixture có số request, thời gian, RSS và kiểm tra dữ liệu. Native append không client-side read-count race; unknown outcome không retry mù.
- AC-05 XLSX: tạo/đọc/điền template/export offline không OAuth/network; hỗ trợ styles/freeze/number formats/named ranges trong phạm vi; cache formula luôn có provenance/freshness rõ.
- AC-06 Reports: tài chính có FX, refunds, duplicate IDs, period filtering, opening/closing reconciliation và truy nguyên; manpower có role/module totals, rate/effort policy, contingency và scenarios. Example 5 x 2 x 1.2 x 1.2 =14.4 days; rate2.5m=36m.
- AC-07 Parity/idempotence: cùng normalized inputs tạo equal totals Google fake backend và XLSX, hai lần chạy không duplicate; template/input hash/version retained. Blank/shorter reruns không giữ số liệu dư sai lệch.
- AC-08 Delivery: actual CLI local report và batch fake backend scenarios chạy được; full applicable suite/test receipt; one consolidated reviewer, blockers closed. Docs/help/examples đồng bộ. Không tuyên bố live Google, Excel GUI hoặc LibreOffice đã verify nếu chưa chạy.

## Execution checklist
- [x] T-01 Chuyển roadmap thành kế hoạch và khóa hợp đồng (AC-08).
- [x] T-02 Khảo sát seams/compatibility/isolation (AC-01).
- [x] T-03 Read modes — accepted by normal offline/integration suite (AC-02).
- [x] T-04 Input file/stdin/schema — accepted by parsing, validation and actual CLI checks (AC-02).
- [x] T-05 Preview/formula overwrite safety — reviewed, repaired and verified (AC-03).
- [x] T-06 Batch/chunking/metadata — accepted by batch/integration suite (AC-04).
- [x] T-07 Read materialization performance — implementation and benchmark verified (AC-04).
- [x] T-08 Native table append và retry contract — normal command/bulk suite passed (AC-04).
- [x] T-09 XLSX read/write/template — offline suite and actual CLI passed (AC-05).
- [x] T-10 XLSX original protection/unsupported features — reviewed and verified (AC-03/05).
- [x] T-11 Declarative templates/transforms — core/integration validation passed (AC-06).
- [x] T-12 Finance report — arithmetic, reconciliation and readback passed (AC-06).
- [x] T-13 Manpower/scenarios — arithmetic, policy and readback passed (AC-06).
- [x] T-14 CLI integration, dual-backend parity, performance evidence (AC-07/08); actual CLI + full smoke + 200,000-cell benchmark passed.
- [x] T-15 Full verification and single review, consolidated repair closure (AC-01..08); official build/unit/full gates passed.
- [x] T-16 Handoff and conditional P4 boundary documented; temporary verification scripts/configuration removed, no commit/push (AC-08).

## Evidence and handoff
- Approved assessment read before exploration. Prior evidence is context, not proof of new features.
- Primary npm registry confirms ExcelJS4.4.0, decimal.js10.6.0, csv-parse7.0.2; licenses MIT. ExcelJS cannot evaluate arbitrary formulas.
- Discovery completed: types belong to GoogleSheetCli namespace; CLI factory and fake-sheets are existing seams. OAuth tests directly use homedir config and require isolation before validation. Scout baseline claim contradicted disk and was rejected.
- Authoritative baseline `git status --short`: only untracked .sot/, docs/plans/ and a client_secret_*.json file; no tracked code changes. The credential file is user-owned, was not opened, and must never be staged or included in broad JSON input fixtures.
- Dependencies installed with `npm install --save-exact --ignore-scripts --no-audit --no-fund exceljs@4.4.0 decimal.js@10.6.0 csv-parse@7.0.2 jszip@3.10.1`, HOME/cache isolated under workspace tmp. Exit 0, added80/removed1. Transitive deprecation warnings recorded; dependency audit is part of final verification, no automatic broad upgrades.
- Wave A dispatched: GoogleData owns GoogleSheet/bulk; ReportCore owns input/schema/reports/examples; XlsxBackend owns local adapter. All skip tests/build/lint/format until integration settles. No implementation acceptance yet.
- Wave A receipts received: GoogleData added render modes/rawOnly, batch APIs, native table append, typed applyReport and focused tests. XlsxBackend added xlsx.ts, xlsx-types/range/zip/file helpers and tests. ReportCore added input/schema/transforms/finance/manpower/types/index, three runnable synthetic example pairs and tests. Receipts claim completion but are NOT acceptance proof; compile/runtime/reviewer gates remain.
- Fixed interfaces allowed Wave B CLI integration to overlap remaining Wave A work without shared writes. Wave A now settled. Wave B: CliIntegration owns commands/base/runner; IntegrationCoverage owns fake backend/integration tests/OAuth test isolation; PackageDocs owns package exports/test discovery/manual README.
- P0 backend render/input/preview present in receipts, but tasks T-03/04/05 stay open until CLI wiring and safety verification. Formula-to-formula replacement, unowned data overwrite, stale cache and shrink-column clearing explicitly requested as integration regression scenarios.
- Wave B settled: CliIntegration added seven domain commands plus existing command flag/input extensions and runner; IntegrationCoverage extended fake API, command inventory, report integration coverage and OAuth-test isolation; PackageDocs added ./report and ./xlsx subpaths, published examples, broadened offline command-test discovery and manual README instructions.
- New test/smoke scripts have NOT passed yet. FinalVerification assigned build, unit/full offline-capable suite, CLI smoke, benchmark, generated docs, package/module graph and production dependency audit with isolated HOME. ConsolidatedReview (single Reviewer invocation) reviews settled source/tests concurrently; no source repairs until both evidence streams settle.
- Graph reconciliation after Wave A: 14 unchanged/0 failed; after integrated src/test: 70 unchanged/0 failed. This is graph consistency evidence, not compilation or runtime acceptance.
- Verification run 1 FAIL: build has six TypeScript diagnostics at five sites (base-class generic parse, Promise.withResolvers incompatible existing lib target, report formula result narrowing, ExcelJS Buffer boundary). Full output artifact://60.
- Standard unit run blocked by compilation; transpile-only blocked by test/oauth.test.ts monkeypatch of readonly namespace homedir. Diagnostic fallback excluding OAuth/live suite: 263 passing, 29 failing (corrected test-runner count), artifact://64. This is NOT a passing acceptance gate.
- Runtime failures include native append/fake final state, formula overwrite/freshness expectations, invalid integration template/runner fixtures, brittle CLI-help/forwarded-option/error-wording assertions. Repair must preserve observable contracts, not re-pin incidental wording.
- Actual CLI help and synthetic finance/manpower XLSX generation worked; this only proves structural/runtime output. Required specific 14.4 days/36m equation remains unverified because the dedicated integration fixture failed. Smoke script passed wrong runner option; benchmark generated 200,000 cells but checked a wrong sheet title, so benchmark is not accepted.
- Generated docs completed using oclif + clean.sh without staging. Package dry-run:102 files/121.1kB, no secret/tmp included; ./sheet import had no ExcelJS/report runtime modules. These results must be reconfirmed after repair if affected.
- Production audit found2 moderate nodes for ExcelJS->uuid, no high/critical. Advisory GHSA-w5hq-g745-h8pq affects v3/v5/v6 custom-buffer methods, not v4; primary advisory lists patched uuid11.1.1. Consider a scoped compatible override after verifying ExcelJS callsites, not broad audit-fix.
- ConsolidatedReview completed once, verdict incorrect, 23 findings; full receipt agent://ConsolidatedReview. Critical confirmed issue: Google report clear omitted endRowIndex/endColumnIndex and could clear outside report. Also compile gates, misleading preview counts, stale metadata, collisions, broken documentation/examples and integration fixtures. No real workbook was mutated by verification.
- Main adjudication: single-quoting valid A1 sheet names is legal; repair semantic range identity/parser/fake handling rather than change valid behavior solely to satisfy string tests. Shared-formula repair must translate relative references (ExcelJS getter) and must not copy a master's formula verbatim. Native append uses explicit INSERT_ROWS for the new API but no exactly-once guarantee or unproven live-concurrency claim.
- Repair owners reported R01–R23/V01–V03 addressed; all acceptance remains pending independent test-runner evidence. Additional boundaries: reject non-representable >15-significant-digit financial numeric outputs; serial dates before1900-03-01 rejected as ambiguous rather than silently aliased.
- Verification run2 still FAIL: duplicate toSafeNumericCell star re-export, missing ReportScalar imports, original XLSX load Buffer typing site not yet fixed, unterminated integration-test try/finally, smoke TypeScript fake import and benchmark null-cell fixture. These are being closed by IntegrationFixups precision worker within the same repair wave; no new Reviewer invocation.
- Run2 actual finance/manpower example CLI output still generated5-sheet workbooks, but dedicated14.4/36m smoke remains blocked. Production workspace audit now0 vulnerabilities after scoped exceljs->uuid11.1.1 override.
- Packaging caveat confirmed: no npm-shrinkwrap.json exists/is packed; downstream package consumers do not inherit root npm overrides. The reviewed ExcelJS usage is uuid.v4 (outside the advisory's v3/v5/v6 buffer paths). Workspace audit0 must not be represented as proof of downstream dependency resolution.
- Verification run3: `npm run build` PASS. Normal unit/full suites still blocked by test/report-integration.ts factory-call type mismatches, missing local readBack and a remaining block closure; source build no longer blocked. TestContractFix handles these without changing intended assertions.
- `node tmp/report-benchmark.cjs` PASS:10,000x20=200,000 data cells;1,348.56ms measured (CSV134.57ms/build125.09ms/apply153.17ms/save935.73ms), observed peakRSS489.83MB, heap269.50MB,0 network requests; boundary values verified. This measures a synthetic full-workbook pipeline, not a comparative speedup or a streaming-memory guarantee.
- Local smoke arithmetic now PASS: manpower base12 days,total14.4 days,cost36,000,000VND; finance inflow95m/outflow15m/net80m/opening100m/closing180m,reconciled. Refund reduces signed inflow in this fixture; earlier100m/20m expectation was incorrect classification, not a changed net balance.
- Google fake smoke still FAIL: requested sheetId100000 from applyReport was not honored by fake addSheet, causing missing-ID updateSheetProperties error. FakeBatchFix is correcting requested IDs/intra-batch semantics; no production workaround to satisfy fake.
- Actual CLI example generation/readback and safe docs regeneration passed; workspace audit remains0. Overall acceptance awaits normal unit/full gates and complete cloud-fake parity smoke.
- Verification run4 normal ts-node suites reached runtime: unit311 PASS/7FAIL (`artifact://152`); full311 PASS/7 offlineFAIL +2 missing-credential live hooksFAIL (`artifact://154`). No failing tests excluded. Remaining fixture issues: guard wording/precedence, malformed generic template schema/tab, nullable uncached-result sentinel, two forwarding/default dictionary assertions, JSON argv splitting. Live hooks contradicted documented skip behavior and are being gated only on missing required live configuration.
- Full `node tmp/report-smoke.cjs` now PASS: finance/manpower arithmetic and XLSX readback above; real GoogleSheet through FakeSheets writes five report sheets and exact manpower formulas `=H2*(1+I2)` / `=J2*K2`. This proves adapter request/formula parity, not recalculation by Google's real service.
- Latest normal build PASS; docs generation PASS; pack102files/121.1kB excludes secret/tmp; `./sheet` imports neither ExcelJS nor report runtime. Benchmark result retained because implementation unchanged. Offline reported coverage from full failed run:88.3% statements/62.27% branches/89.98% functions/90.15% lines; this is diagnostic coverage, not final acceptance.
- Final fixture-closure owners: ReportFixtureClosure (integration/XLSX tests), CommandFixtureClosure (offline command tests), LiveSkipClosure (live suite guards). Production source unchanged in this closure; no Reviewer re-invocation.
- Final fixture corrections settled; full test-project typecheck and normal suites are running again. Existing live suites now reuse canonical `describeLive` from commands/helper, retaining real auth failures when all required live settings are provided.
- Source graph reconcile after settled edits: `sot_reconcile {paths:["src","test"],workers:2}` →70 unchanged,0 failed.

### Final official gate receipt
- Environment isolated with `env -i`, preserved executable PATH, `HOME=$PWD/tmp/test-home`, `npm_config_cache=$PWD/tmp/npm-cache`; no live credentials or external OAuth configuration.
- `npm run build` →exit0, no TypeScript diagnostics.
- `npm run test:unit` →exit0,320 passing/0 failures/0 pending; log `artifact://158`.
- `npm test` →exit0,320 passing/0 failures/29 pending live tests; log `artifact://160`. Overall coverage88.77% statements,63.82% branches,90.01% functions,90.65% lines. No claims of production-module100% coverage inferred from test-file rows.
- Actual CLI finance/manpower generation and workbook readback passed; full business smoke and benchmark evidence above retained. Temporary smoke/benchmark scripts removed after proof; permanent regression fixtures/tests remain.
- `./node_modules/.bin/oclif readme --multi && sh ./bin/clean.sh` →exit0; generated docs regenerated, not hand-edited/staged.
- `npm pack --dry-run` →102 files,121.1kB; no credentials/tmp artifacts. `./sheet` imports no ExcelJS/report runtime. `npm audit --omit=dev --json` →0 workspace vulnerabilities, with the downstream override caveat already documented.
- Additional ad-hoc `tsc --noEmit -p test/tsconfig.json` returned TS6306 because existing test references point to a non-composite root; baseline configuration left unchanged. Supplemental standalone program check `tsc -p tmp/report-typecheck.B8N7In/tsconfig.json` inherited the same compiler settings, cleared project references, used repository rootDir/noEmit/incremental:false →exit0,0 diagnostics across all test sources,2.25s. No diagnostic suppression or source edits; temporary config directory removed after proof.
- AC-01..AC-08 accepted within the explicitly offline/synthetic scope by official gates, CLI smoke, backend parity tests and benchmark. Real Google recalculation, live edits, private templates and Excel GUI are not certified.
- One consolidated Reviewer pass completed; R01–R23 and V01–V03 closed via repair receipts and independent normal build/tests/smoke. No re-review was invoked. No commit, tag, publish or push performed.

## Consolidated repair wave
One repair wave, three independent owners; all skip mid-flight validation. Main owns acceptance. Public report contract remains compatible; optional policy fields only.
- RepairGoogleCli: GoogleSheet/sheet-batch/base-class/cli-input/commands/runner plus test/google-sheet-bulk.test.ts. No fake or other tests/package edits.
- RepairReportsXlsx: report core except runner, all XLSX helpers, test/report-core.test.ts and test/xlsx.test.ts; optional shared types additions must be announced. No commands/package/fake edits.
- RepairEvidence: integration/fake/OAuth/offline/output tests, tmp smoke/benchmark scripts, README/package/lock. No application source edits. Coordinate fixtures with repaired contracts.

### Finding closure checklist
- [x] R01 Build diagnostics: normal `npm run build` passed after import/export and Buffer-boundary correction; test-harness diagnostics tracked under R07.
- [x] R02 Restore intended offline command-test discovery.
- [x] R03 Bound Google clearing to previously owned extent; preserve unrelated cells.
- [x] R04 Canonical/quoted A1 semantics and fake parsing, not incidental string spellings.
- [x] R05 Associate preview formula conflicts with their actual update.
- [x] R06 Correct documented library/CLI examples to actual APIs.
- [x] R07 Repair integration fixtures and replace incidental wording/forwarding assertions with behavior; prove financial/manpower math.
- [x] R08 Validate each batch update width independently.
- [x] R09 Identical formula reapply allowed; changed formula protected on both backends.
- [x] R10 Input/template/output file collision requires explicit inPlace and backup.
- [x] R11 New native append INSERT_ROWS; server-confirmed receipt, fake supports actual semantics.
- [x] R12 Eliminate no-op overwrite flag; keep legacy behavior explicit and protected new paths real.
- [x] R13 Report accurate formula overwrite/protection counts and ragged-cell semantics.
- [x] R14 Update existing ownership metadata and accurately handle partial structural writes.
- [x] R15 Fake numeric rendering models decimals and text identifiers correctly.
- [x] R16 Resolve translated shared formulas or report unavailable; never return master address as formula.
- [x] R17 ZIP size accounting fails closed if bounds cannot be determined.
- [x] R18 Ragged rows do not spuriously conflict or silently clear omitted cells.
- [x] R19 Workbook inline input uses the same safe parser/validation.
- [x] R20 Multi-range datasets cannot silently overwrite same-sheet aliases.
- [x] R21 Scenario baseline independent of display names; ratios display as percentages.
- [x] R22 rawOnly skips formatted construction.
- [x] R23 Reject/normalize invalid double-quoted A1 input consistently.
- [x] V01 OAuth test isolation without readonly namespace patch.
- [x] V02 Full smoke and 200,000-cell benchmark passed with corrected real interfaces; see verification run4 evidence.
- [x] V03 Scoped UUID override passes workspace audit; downstream resolution limitation and advisory reachability documented.

## Assumptions and contingencies
- Representative private customer templates were not supplied; synthetic fixtures in the repository cover the important financial/manpower structures. Real-template fidelity remains an explicit acceptance limitation, not permission to claim arbitrary workbook support.
- No permission to access ~/.config/google-sheet-cli or other external paths in this task. All runtime checks stay workspace with synthetic data; live Google verification requires explicit permission and a disposable target.
- OAuth test isolation is verified via workspace HOME and scoped restoration; global OAuth files were not accessed.
- True concurrent Google edit isolation cannot be guaranteed by client-side preflight. Expose safety limits and partial writes honestly.
- Scope P4 engine/native charts/pivots is conditional per approved proposal; do not install LibreOffice/HyperFormula or request broader OAuth scopes by default.
