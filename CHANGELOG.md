# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

---

## [3.1.0] — 2026-09-25

### Added

Mười một command mới cùng một cơ chế lỗi có cấu trúc dùng chung, đưa CLI từ 30 lên 42 commands:

- **Structured JSON errors cho mọi command:** `--json`/`-j` báo lỗi dưới dạng một envelope JSON duy nhất trên stderr, exit code 1 (`code` ổn định, `message` an toàn, `retryable`, `retryAfterMs?`, `issues?` liên kết tọa độ) — output thành công giữ nguyên; `--rawOutput`/`-r` in cả success lẫn failure dưới dạng JSON. Bộ mã lỗi: `USAGE` (sai argv), `AUTH_REQUIRED`/`UNAUTHORIZED`/`FORBIDDEN` (xác thực), `NOT_FOUND`/`CONFLICT`/`REQUEST_INVALID`/`RATE_LIMITED`/`UPSTREAM`/`NETWORK` (transport/API), `VALIDATION`/`SCHEMA_INVALID`/`DATA_INVALID` (hợp đồng dữ liệu), `INTERNAL`. Envelope không bao giờ chứa credentials, request/response internals hay error object tùy ý.
- **`capabilities`:** self-description cho agent — liệt kê toàn bộ commands, input/output forms, destructive guards và error codes mà không cần đọc docs.
- **`data:schema`:** khám phá schema cột của worksheet (read-only) trên một mẫu giới hạn (mặc định dòng 1–100, cột A–Z; đổi bằng `--minRow/--minCol/--maxRow/--maxCol`, dòng mẫu đầu là header row). Báo header thực tế kèm tọa độ A1 tuyệt đối, kiểu dữ liệu suy luận kèm số mẫu, số ô chứa formula, data-validation rules và named ranges; empty/duplicate header và mixed type được báo minh bạch. Kiểu suy luận chỉ mô tả mẫu — Google lưu date dưới dạng number, không phải schema chính thức.
- **`data:validate`:** kiểm tra dữ liệu worksheet chống lại một TableSchema (`--schema` JSON inline hoặc `--schemaFile`/stdin), tái sử dụng đúng semantics của report (`required`, `type`, `enum`, `unique`, `min`/`max`). Mỗi violation báo kèm dòng và tọa độ A1 chính xác; dữ liệu sai exit 1 với mã `DATA_INVALID`, schema malformed fail với `SCHEMA_INVALID` trước khi bất kỳ request nào rời tiến trình. Không bao giờ ghi.
- **`data:clear`:** xóa giá trị (values-only) trong một range A1 có giới hạn hai chiều, giữ nguyên number format và mọi thuộc tính ô khác. Ô formula từ chối bị xóa trừ khi có `--overwriteFormulas`; `--dryRun` xem trước số ô sẽ bị xóa với zero mutation. Hỗ trợ cả file `.xlsx` local qua `--workbook`.
- **`data:upsert`:** upsert theo một cột key. Key tồn tại → chỉ cập nhật đúng các ô input cung cấp (cột bị bỏ qua và formula hiện có được giữ nguyên); key mới → append dưới bảng. Input header-first dạng JSON/CSV/stdin, đúng một cột `--key` bắt buộc, so khớp key theo kiểu dữ liệu (string `"001"` không bao giờ khớp number `1`, leading zero được giữ), ghi `RAW` mặc định, từ chối key rỗng/trùng, và từ chối trước khi ghi nếu range giới hạn che giấu dữ liệu bảng phía sau. `--dryRun` báo added/updated/unchanged và planned ranges. Chỉ single-writer: chạy lại cùng input không sinh bản ghi trùng, nhưng không có transaction, không blind-retry sau failure mơ hồ, không hỗ trợ nhiều writer đồng thời.
- **`data:export-csv`:** export một range worksheet ra CSV trên stdout hoặc file, UTF-8, escaping chuẩn RFC 4180.
- **`spreadsheet:list`:** liệt kê các spreadsheet truy cập được qua Drive API (`drive.file` scope), phân trang `--limit`/`--pageToken`.
- **`spreadsheet:copy`:** nhân bản spreadsheet qua Drive API. Source không đổi, receipt trả id và title do server gán; CLI không tự cấp lại sharing grants của nguồn. Kiểm tra quyền hiệu lực của bản sao trước khi chia sẻ dữ liệu nhạy cảm.
- **`worksheet:copy`:** copy một worksheet sang spreadsheet đích tường minh qua Sheets API. Values và formulas được mang theo, source không đổi; khi đích đã có sheet trùng tên, server tự gán title duy nhất và receipt báo title cuối cùng cùng sheet id mới.
- **`spreadsheet:export`:** export spreadsheet ra file PDF hoặc XLSX local qua Drive API. Ghi binary an toàn, atomic (file tạm + rename; file output sẵn có sống sót qua failure), chỉ thay thế khi có `--overwrite` tường minh, receipt kèm mime type, số byte và đường dẫn. Chịu giới hạn export 10 MB của Drive, và quyền truy cập nằm trong scope `drive.file` (chỉ các file app này tạo hoặc đã mở) — không phải toàn bộ Drive.
- **`workbook:names`:** quản lý defined names của file `.xlsx` local — `list`, `add` (`--name`, `--range`), `remove` (`--name`).

#### Local XLSX engine mở rộng (`--workbook`)

- **Structural mutations:** `grid:insert`/`grid:delete` chèn/xóa dòng-cột trên `.xlsx` với merge-range management (unmerge → splice → re-merge; merge cắt ranh giới splice → từ chối `CONFLICT` trừ khi `--force`); `format:merge --workbook`; `workbook:find` tìm ô theo điều kiện; `workbook:write --cells` ghi từng ô rời rạc.
- **`--updateRefs` trên `grid:insert`/`grid:delete`:** rewrite tham chiếu A1 cùng sheet và defined names sau splice — ref nằm trọn trong vùng xóa hoặc bị đẩy quá mép sheet → `#REF!`, receipt báo `refsRewritten`/`refsBroken`. Không có cờ này, refs giữ nguyên và receipt cảnh báo `formulasAtRisk`. Cross-sheet refs chưa hỗ trợ (phase 2).
- **Formatting:** `format:cells --workbook` hỗ trợ font (bold/italic/underline/strikethrough/size/family/color), fill, alignment, `--wrapText`, `--numberFormat`, borders; các cờ không có tương đương ExcelJS (`--numberFormatType`, `--wrapStrategy` khác WRAP) từ chối bằng `USAGE` thay vì silent no-op.
- **Grid:** `grid:freeze --workbook` (merge vào view hiện có, giữ axis không chỉ định và các thuộc tính view khác), `grid:resize --workbook` (`--pixels` hoặc `--auto`), `grid:hide --workbook` (`--unhide`).
- **Sheet management:** `worksheet:add|remove|rename --workbook` — từ chối trùng title, xóa sheet visible cuối cùng, rename va chạm; `remove` dọn defined names thuộc sheet bị xóa.
- **Safety:** mọi mutation `--workbook` hỗ trợ `--dryRun` và `-o/--inPlace`; save atomic qua file tạm + backup `.bak` + SHA-256; file chứa macro/chart/pivot/external link… bị từ chối ghi trừ khi `--discardUnsupported`.

#### Agent-gateway hardening

- **`--redacted` / `GSHEET_REDACTED`:** che credentials, token và đường dẫn nhạy cảm trong mọi output.
- **`GSHEET_CONFIG_DIR`:** đổi thư mục cấu hình/token mặc định.
- **Atomic token writes:** token OAuth ghi atomic với quyền 0600.
- **`error.mutation` envelope:** receipt mutation thất bại trả structured error thay vì partial output.

### Fixed

- Lỗi refresh OAuth token giờ giữ lại HTTP status và error code của Google (cùng `cause` non-enumerable) để structured errors phân loại đúng, đồng thời giữ nguyên message lỗi cũ.
- Lỗi xác thực/input cục bộ được phân loại là `AUTH_REQUIRED`, `VALIDATION` hoặc `NETWORK` thay vì `INTERNAL` chung chung trong structured error mode.
- `data:validate` thu thập cả lỗi trùng giá trị (`unique`) lẫn lỗi kiểu/required trong cùng một lần kiểm tra, thay vì để lỗi ở một dòng che khuất duplicate ở các dòng khác. Kết quả chỉ áp dụng cho phạm vi mẫu đã chọn (mặc định 100 dòng × 26 cột).
- Giới hạn structured errors: lỗi unknown command/topic xảy ra trước khi command class được nạp vẫn dùng output human-readable của oclif.

---

## [Fork Additions] — 2026-09-18 to 2026-09-19 (minhgv)

Tất cả tính năng được phát triển sau khi fork từ upstream (`jroehl/google-sheet-cli`), đưa CLI từ một công cụ CRUD cơ bản 9 lệnh thành **AI Agent Data Gateway** với 30 commands, hỗ trợ offline sandbox, batch APIs, định dạng bảng, thao tác cấu trúc và phân quyền Drive.

### 🚀 Added

#### 🔐 Authentication (`auth:*`)
- `auth:login`: Đăng nhập tương tác qua OAuth 2.0 Desktop Flow (local loopback server cổng 3000-3002/8080-8081). Lưu refresh token tại `~/.config/google-sheet-cli/token.json`.
- `auth:logout`: Thu hồi và xóa token lưu trên máy.
- `auth:status`: Kiểm tra trạng thái xác thực và thời hạn token.
- Hỗ trợ song song cả Service Account JWT và OAuth 2.0 User Credentials qua cờ `--useOauth`.

#### 📊 Batch & Table Data Operations (`data:*`)
- `data:batch-get`: Đọc đồng thời nhiều range trong 1 request HTTP, trả về `{ranges: [{range, values}]}`, tránh rate-limit Google API.
- `data:batch-update`: Ghi dữ liệu đồng thời vào nhiều range trong 1 request atomic; hỗ trợ `--dryRun`.
- `data:append-table`: Append dữ liệu tự động map theo header cột, tự đồng bộ schema dòng.
- `data:find`: Tìm kiếm tọa độ ô / dòng theo điều kiện (text, regex, exact, số học); trả về `{a1, row, column, value, rowValues}`.
- Formula Protection: Cơ chế bảo vệ tự động chặn ghi đè lên các ô chứa công thức (`FORMULA_OVERWRITE_GUARD`), trừ khi có cờ `--overwriteFormulas`.
- Nhận input linh hoạt: positional JSON, đọc file `-i <path>` (hỗ trợ cả `.json` và `.csv`), hoặc qua stdin (`-i -`).
- Thêm cờ `--valueRenderOption` (`FORMATTED_VALUE`, `UNFORMATTED_VALUE`, `FORMULA`) và `--dryRun` cho toàn bộ nhóm ghi dữ liệu.

#### 🎨 Formatting & Visual Styling (`format:*`)
- `format:cells`: Định dạng cell an toàn (chỉ tác động `userEnteredFormat`, không làm mất dữ liệu): in đậm (`--bold`), màu chữ (`--textColor`), màu nền (`--backgroundColor`), kẻ khung (`--borders all|outer|top|bottom|left|right`), định dạng số/ngày/tiền tệ (`--numberFormat`).
- `format:merge`: Hỗ trợ gộp (`MERGE_ALL`, `MERGE_COLUMNS`, `MERGE_ROWS`) và bỏ gộp (`--unmerge`) ô.

#### 📐 Grid Structural Mutations (`grid:*`)
- `grid:insert`: Chèn dòng hoặc cột (`--dimension ROWS|COLUMNS`, `--start`, `--count`).
- `grid:delete`: Xóa dòng/cột; hỗ trợ `--dryRun` trích xuất và in preview chính xác dữ liệu sắp bị xóa trước khi thực thi.
- `grid:hide`: Ẩn hoặc hiện dải dòng/cột (`--hidden` / `--unhide`).
- `grid:resize`: Thay đổi kích thước dòng/cột theo pixel (`--pixels`) hoặc tự co giãn theo nội dung (`--auto`).
- `grid:freeze`: Cố định tiêu đề dòng/cột (`--rows`, `--columns`, truyền `0` để unfreeze).

#### 📄 Drive Sharing & Permissions (`spreadsheet:*`)
- Tích hợp Google Drive API v3 qua phạm vi `drive.file` (an toàn, không đòi quyền truy cập toàn bộ Drive).
- `spreadsheet:share`: Cấp quyền truy cập cho email (`--email`), Google Workspace Domain (`--domain`) hoặc public link (`--anyone`); phân quyền `--role reader|commenter|writer`; mặc định tắt spam thông báo (`--notify=false`).
- `spreadsheet:permissions`: Liệt kê danh sách tất cả quyền truy cập đã cấp trên bảng tính.
- `spreadsheet:unshare`: Thu hồi quyền theo `permissionId` hoặc tự động resolve theo địa chỉ `email`.

#### 📁 Offline Local XLSX Engine (`workbook:*`)
- Hoạt động 100% offline không cần mạng, không cần credentials Google:
  - `workbook:read`: Đọc dữ liệu từ file Excel `.xlsx` local theo cú pháp A1 range.
  - `workbook:write`: Ghi dữ liệu vào file `.xlsx` local với cơ chế ghi atomic (qua file tạm `.tmp`) và tính hash SHA-256 bảo đảm tính toàn vẹn.
  - `workbook:inspect`: Xem cấu trúc metadata, danh sách sheet, kích thước dòng/cột của file `.xlsx`.

#### 📋 Declarative Report Automation (`report:*`)
- `report:run`: Tạo bảng báo cáo hoàn chỉnh từ schema JSON khai báo trước (`finance`, `manpower`, `sales`), tự động tạo sheet, đổ dữ liệu, gán công thức tính tổng và áp dụng bộ màu, kẻ viền hoàn chỉnh.
- Cung cấp bộ templates và dữ liệu mẫu thực tế tại `examples/reports/`.

#### 🤖 AI Agent Integration
- Cung cấp file skill chuẩn `skill/SKILL.md` và tài liệu tích hợp `docs/agents.md` dành cho Claude Code / OpenCode / Oh My Pi.
- Chuẩn hóa đầu ra máy đọc (`--rawOutput` JSON, `--csv`), hỗ trợ piping qua `jq`.

### 🛡️ Fixed
- `f0ff600`: Sửa lỗi oclif 5 tự động nuốt stdin vào argument đầu tiên gây xung đột với cờ `-i -`.
- `f5da889`: Sửa script build tự động dọn dẹp thư mục `lib/` trước khi chạy `tsc`.

---

## [3.0.0] — Upstream Milestones (jroehl / Johann Emrich)
Phiên bản nền tảng trước khi fork về:
- **Node 22 Platform Lift**: Nâng sàn Node.js lên `>=22`, chuyển CLI lên framework `@oclif/core` v5.
- **Slim Dependency**: Thay thế toàn bộ thư viện cồng kềnh `googleapis` bằng gói chuyên biệt `@googleapis/sheets`.
- **Error Consistency**: Chuẩn hóa toàn bộ lỗi ném ra thành `Error` instance thay vì string literal.
- **Quota Backoff**: Tự động retry và backoff theo cấp số nhân khi gặp lỗi giới hạn quota Sheets API (HTTP 429).
- **Offline Command Testing Layer**: Xây dựng `FakeSheets` mock server nội bộ chạy 180 unit test không cần internet hay credentials.
- **Subpath Export**: Xuất module `./sheet` độc lập, tách biệt khỏi runtime của oclif.

---

## [2.x] — Upstream Legacy Milestones
- **2.3.0**: Hỗ trợ private key chuẩn PKCS#1, nạp credentials từ file JSON qua cờ `-f / --credentialsFile`.
- **2.2.0**: Cơ chế tự động mở rộng lưới (`growSheetIfNeeded`) trước khi ghi vượt kích thước bảng tính.
- **2.0.0 - 2.1.x**: 9 commands CRUD cơ bản (`data:get|append|update`, `spreadsheet:add|get`, `worksheet:add|get|remove|rename`). Chỉ hỗ trợ Service Account credentials.
