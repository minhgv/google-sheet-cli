# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
