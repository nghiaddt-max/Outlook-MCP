# Outlook-MCP

MCP server cho Claude đọc mail, lịch, danh bạ trên **Outlook cổ điển (Office 16)** qua COM và soạn **thư nháp**.
Không có tool nào gửi mail: thư soạn ra nằm trong Drafts, bạn tự mở ra kiểm tra rồi bấm Send.

> Công cụ nội bộ của RIKCO JSC, công khai để tiện cài đặt. Dùng trên máy và hộp thư của chính bạn,
> tự chịu trách nhiệm khi sử dụng; không có cam kết hỗ trợ hay bảo hành.

## 1. Yêu cầu

- Windows 10/11, **Outlook cổ điển** (bản Office) đã cài và đăng nhập mail. **New Outlook không hỗ trợ COM**,
  nên không dùng được.
- **Outlook phải đang mở** khi dùng. Server không tự khởi động Outlook (Outlook khởi động ngầm có thể gửi đi
  thư đang nằm trong Outbox mà bạn không biết).
- **Python 3.14** từ [python.org](https://www.python.org/downloads/windows/) (khi cài tick *Add python.exe to PATH*).
  Phiên bản thư viện trong `constraints.txt` được kiểm thử với Python 3.14; Python ≥ 3.10 cũ hơn có thể cài không được.
- Một client MCP: Claude Desktop, Claude Code, hoặc client khác hỗ trợ MCP qua stdio.

## 2. Cài đặt nhanh (người dùng)

1. Tải bản mới nhất ở trang **[Releases](https://github.com/nghiaddt-max/Outlook-MCP/releases/latest)**:
   file `outlook-mcp-<phiên bản>-windows.zip` và `.zip.sha256`.
2. Đối chiếu mã SHA256 (giá trị in ra phải trùng nội dung file `.sha256`):
   ```bash
   Get-FileHash .\outlook-mcp-0.3.2-windows.zip
   ```
3. Giải nén, mở PowerShell trong thư mục giải nén, chạy:
   ```bash
   powershell -ExecutionPolicy Bypass -File .\install.ps1
   ```
   Script tự kiểm checksum từng file trong gói, tạo môi trường Python riêng ở `%LOCALAPPDATA%\outlook-mcp`,
   cài đúng phiên bản thư viện đã kiểm thử và tạo file cấu hình mẫu `mcp-config.json`.
4. Kiểm tra (Outlook đang mở):
   ```bash
   & "$env:LOCALAPPDATA\outlook-mcp\.venv\Scripts\outlook-mcp.exe" --check
   ```
5. Đăng ký với Claude. **Claude Code và Claude Desktop (Chat, Cowork) dùng hai file cấu hình riêng**:
   đăng ký ở chỗ nào thì chỉ chỗ đó thấy tool. Muốn dùng ở cả hai thì đăng ký cả hai.
   - **Claude Code** (ghi vào `%USERPROFILE%\.claude.json`):
     ```bash
     claude mcp add --scope user outlook -- "%LOCALAPPDATA%\outlook-mcp\.venv\Scripts\outlook-mcp.exe"
     ```
     Kiểm tra: `claude mcp get outlook` phải báo `Connected`.
   - **Claude Desktop – Chat / Cowork:** thêm mục `mcpServers` vào `claude_desktop_config.json`
     (Settings → Developer → Edit Config). Vị trí file tùy bản cài:
     - bản cài từ claude.ai: `%APPDATA%\Claude\claude_desktop_config.json`
     - một số bản Microsoft Store: `%LOCALAPPDATA%\Claude-Data\claude_<mã>\claude_desktop_config.json`

     File này đã có sẵn các thiết lập khác (`preferences`…): **sao lưu trước, chỉ thêm khóa `mcpServers`**,
     không xóa nội dung cũ:
     ```json
     {
       "preferences": { "...": "giữ nguyên" },
       "mcpServers": {
         "outlook": {
           "command": "C:\\Users\\<tên>\\AppData\\Local\\outlook-mcp\\.venv\\Scripts\\outlook-mcp.exe",
           "args": []
         }
       }
     }
     ```
     Rồi **thoát hẳn** Claude Desktop (chuột phải icon ở khay hệ thống → Quit, đóng cửa sổ là chưa đủ) và mở lại.
     Claude Desktop ghi file này liên tục khi đang chạy; nếu mở lại không thấy tool `outlook`, mở file kiểm tra
     mục `mcpServers` còn không (có thể bị ghi đè), thêm lại khi app đã thoát hẳn.
   - **Lưu ý đường dẫn khi cài từ bên trong Claude Desktop bản Microsoft Store:** nếu bạn nhờ Claude (tab Code /
     Cowork) chạy `install.ps1` hoặc tạo venv, Windows chuyển hướng `%LOCALAPPDATA%` vào thư mục riêng của app,
     nên server thực tế nằm ở
     `%LOCALAPPDATA%\Packages\Claude_<mã>\LocalCache\Local\outlook-mcp\.venv\Scripts\outlook-mcp.exe`.
     Đăng ký bằng **đường dẫn thực tế này** (xem dòng `Actual location` khi tạo venv). Gỡ app Claude sẽ xóa luôn
     thư mục đó, cần cài lại. Chạy `install.ps1` từ PowerShell thường (ngoài Claude) thì không bị chuyển hướng.
6. Thử: *"Liệt kê 5 thư mới nhất trong Inbox"*. Sau khi đăng ký phải mở phiên/cuộc trò chuyện mới mới thấy tool.

Tùy chọn của `install.ps1`: `-InstallDir` (nơi cài), `-Python` (chỉ định python.exe), `-Wheelhouse` (cài offline từ
thư mục chứa sẵn file `.whl` phụ thuộc), `-Force` (xóa môi trường cũ rồi cài lại). Cập nhật: chạy `install.ps1`
của zip mới.

## 3. Cài từ mã nguồn (phát triển)

```bash
git clone https://github.com/nghiaddt-max/Outlook-MCP.git
```

```bash
cd Outlook-MCP
```

```bash
py -3 -m venv .venv
```

```bash
.venv\Scripts\python.exe -m pip install --only-binary=:all: -c constraints.txt -e .
```

```bash
.venv\Scripts\outlook-mcp.exe --check
```

Đóng gói bản phát hành (kết quả trong `dist\`: zip gồm `.whl`, `install.ps1`, `README.md`, `constraints.txt`,
`SHA256SUMS.txt`, kèm file `.zip.sha256`), rồi đính kèm vào một GitHub Release:

```bash
powershell -ExecutionPolicy Bypass -File .\build_dist.ps1
```

Sau khi sửa code hoặc đăng ký lần đầu, client MCP phải kết nối lại (mở phiên mới / khởi động lại app) mới thấy tool mới.

## 4. Danh sách tool

| Tool | Loại | Việc làm |
|---|---|---|
| `outlook_info` | đọc | Phiên bản, profile, tài khoản, hộp thư, số thư Inbox, thư mục được phép lưu file |
| `outlook_list_folders` | đọc | Cây thư mục kèm số thư và số thư chưa đọc |
| `outlook_search_mail` | đọc | Tìm thư theo nội dung, người gửi, tiêu đề, ngày, chưa đọc, có đính kèm; có phân trang |
| `outlook_get_mail` | đọc | Đọc đầy đủ một thư: người nhận, nội dung, danh sách đính kèm |
| `outlook_list_events` | đọc | Lịch hẹn trong khoảng ngày, có trải lịch lặp lại |
| `outlook_search_contacts` | đọc | Tìm danh bạ theo tên, email, công ty |
| `outlook_save_attachments` | ghi file | Lưu đính kèm ra thư mục được phép trên máy, không ghi đè (trùng tên thì thêm ` (1)`) |
| `outlook_create_draft` | nháp | Soạn thư mới, lưu vào Drafts |
| `outlook_reply_draft` | nháp | Soạn trả lời / trả lời tất cả, lưu vào Drafts |
| `outlook_forward_draft` | nháp | Soạn chuyển tiếp (giữ đính kèm gốc), lưu vào Drafts |

Thư mục nhập bằng tên tắt (`inbox`, `sent`, `drafts`, `deleted`, `junk`, `calendar`, `contacts`), không phụ thuộc
ngôn ngữ giao diện, hoặc đường dẫn như `inbox/Du an A`, hoặc đường dẫn đầy đủ lấy từ `outlook_list_folders`.
Ngày nhập dạng `YYYY-MM-DD` hoặc `YYYY-MM-DD HH:MM`, theo giờ máy.

Ví dụ câu lệnh cho Claude:

- "Tìm mail của nhà cung cấp có tên miền example.com.vn từ đầu tháng 10 có đính kèm"
- "Đọc thư mới nhất trong Inbox/Du an A rồi lưu đính kèm vào thư mục Downloads\Outlook-MCP\Du an A"
- "Soạn nháp trả lời thư này: cảm ơn, bên em sẽ phản hồi trước thứ Sáu"

## 5. Bảo mật

- **Chỉ đọc/ghi file trong thư mục được phép.** Mặc định **chỉ** `%USERPROFILE%\Downloads\Outlook-MCP`.
  Áp dụng cho `outlook_save_attachments` (nơi lưu) và tham số `attachments` của các tool soạn nháp (file lấy để đính kèm).
  Thư mục khác phải khai báo rõ bằng biến môi trường `OUTLOOK_MCP_ALLOWED_DIRS` (cách nhau `;`) trong cấu hình MCP.
  Chỉ khai báo thư mục **dữ liệu** (hồ sơ, bản vẽ), không khai báo thư mục chứa mã nguồn/công cụ:
  ```json
  "outlook": { "command": "...", "args": [], "env": { "OUTLOOK_MCP_ALLOWED_DIRS": "D:\\Du an;%USERPROFILE%\\Downloads\\Outlook-MCP" } }
  ```
- **Chỉ nhận file tài liệu** (danh sách được phép, không phải danh sách chặn): pdf, doc/docx/docm, xls/xlsx/xlsm/xlsb,
  csv, txt, rtf, ppt/pptx, odt/ods/odp, mpp, vsd/vsdx, dwg/dxf/dwf/dwfx, ifc, nwd/nwc, rvt/rfa, ảnh
  (jpg/png/gif/bmp/tif/webp/heic), zip/rar/7z, msg/eml. Mọi loại khác bị từ chối: `.exe`, `.ps1`, `.py`/`.pyd`/`.pth`,
  `.js`/`.mjs`/`.ts`, `.lsp`, `.md` (kể cả `CLAUDE.md`, `AGENTS.override.md`), `.json`, `.lnk`/`.url`…
- Đường dẫn do model gửi phải là đường dẫn **đầy đủ**: không mở rộng `%BIEN%` hay `~` (tránh lộ biến môi trường qua
  thông báo lỗi); chặn đường dẫn mạng `\\máy\share` / WebDAV / `\\?\`, ký tự ẩn, thư mục/file ẩn bắt đầu bằng `.`,
  thư mục cài đặt và mã nguồn của chính server, file đính kèm > 25 MB.
- File đính kèm lưu xuống được gắn **Mark-of-the-Web** (vùng Internet): Windows/Office cảnh báo khi mở.
  Tên file được làm sạch: bỏ ký tự đảo chiều (RTLO) giả đuôi file, ký tự ẩn, tên thiết bị `CON`/`NUL`/`COM¹`
  (kể cả dạng `CON .txt`), cắt độ dài.
- **Thư nháp:** kết quả trả về mục `warnings` liệt kê BCC, người nhận ngoài công ty, file đính kèm — đọc kỹ trước khi
  bấm Send. HTML (`is_html=true`) chỉ giữ thẻ định dạng cơ bản; bỏ `style` (chặn giấu chữ), ảnh (chặn ảnh từ xa báo tin),
  script, form, comment. Không xác định được tên miền tài khoản gửi thì coi mọi người nhận là bên ngoài.
- **Chống prompt injection:** mọi chuỗi trả về được lọc ký tự ẩn (bidi, zero-width, Unicode Tags "ASCII smuggling",
  variation selector, control C1); `outlook_get_mail` kèm `content_warning`; tìm kiếm mặc định không trả preview nội
  dung (`preview_chars=0`); `include_subfolders` bỏ qua Junk.
- Giới hạn tham số khai trong schema: `limit` ≤ 100, `offset` ≤ 5000, `preview_chars` ≤ 1000, `max_body_chars` ≤ 50000;
  truy vấn quá nặng (số thư mục × số thư cần đọc > 20 000) bị từ chối để Outlook không bị đơ.
- Annotations: tool soạn nháp khai `openWorldHint=true` (Drafts đồng bộ lên mail server).
- Gói cài: phụ thuộc pin trong `constraints.txt`, chỉ cài wheel, `install.ps1` kiểm `SHA256SUMS.txt` trước khi cài.

Lớp chặn cuối cùng vẫn là **bạn duyệt từng lệnh gọi tool**: đừng bấm Allow cho `outlook_save_attachments` hay tool soạn
nháp nếu bạn không yêu cầu việc đó.

## 6. Lưu ý

- **Chữ ký mặc định** (File → Options → Mail → Signatures) được chèn tự động dưới nội dung cho thư mới,
  trả lời và chuyển tiếp, theo tài khoản gửi. Tắt bằng `signature=false`. `signature_added=false` trong kết quả
  nghĩa là tài khoản chưa đặt chữ ký mặc định. Dùng `display=true` để mở cửa sổ soạn thư và kiểm tra trước khi gửi.
- `unresolved_recipients` trong kết quả là tên/địa chỉ Outlook không nhận ra, cần sửa trước khi gửi.
- Tài khoản IMAP: lịch và danh bạ là loại *This computer only* (chỉ lưu trên máy), nên thường trống.
  Lịch/danh bạ công ty trên Exchange/Microsoft 365 chỉ thấy khi profile Outlook có tài khoản Exchange.
- Nếu Outlook hiện hộp thoại *"A program is trying to access e-mail address information"*, tool sẽ chờ tới khi bạn
  bấm Allow. Hộp thoại này thường không xuất hiện khi antivirus (Windows Defender) đang hoạt động.
- **Dữ liệu đi đâu:** server không tự lưu hay gửi nội dung thư đi đâu, nhưng mọi kết quả tool (nội dung thư, địa chỉ,
  danh bạ) được chuyển tới nhà cung cấp mô hình AI của client đang gọi (ví dụ Anthropic với Claude, OpenAI với Codex).
  Thư nháp lưu vào Drafts được đồng bộ lên mail server.

## 7. Trạng thái kiểm thử

Đã kiểm thử trên Windows 10, Outlook 16 (Microsoft 365 Apps), tài khoản IMAP, Python 3.14, `mcp` 2.3.0:

- Lọc theo ngày, chưa đọc, có đính kèm, nội dung, người gửi khớp với duyệt tay toàn bộ Inbox; thư báo lỗi gửi (NDR)
  không có `ReceivedTime` vẫn được lọc và sắp xếp theo giờ giao thư; phân trang không trùng, mới nhất trước.
- Soạn thư mới / trả lời / chuyển tiếp: vào đúng Drafts, chưa gửi, giữ tiếng Việt có dấu, đúng tài khoản gửi và
  chữ ký mặc định (nội dung → chữ ký → trích dẫn).
- Bảo mật: 92 kịch bản tấn công chạy offline đạt (đường dẫn mạng/đặc biệt, thư mục ẩn, file thực thi và file chỉ dẫn
  của công cụ AI, biến môi trường, ký tự ẩn, tên thiết bị, HTML giấu nội dung); chạy thật lưu đính kèm kèm
  Mark-of-the-Web, thư nháp có `warnings` đầy đủ.
- Gọi qua giao thức MCP stdio: initialize, list_tools (10 tool, schema có min/max), lỗi trả về `isError`.
- Gói cài: cài từ zip vào môi trường mới, kiểm checksum, phát hiện được file trong gói bị sửa.
- Python 3.12.10 (Windows 11, Outlook 16, `mcp` 2.3.0): cài từ mã nguồn, `--check`, `outlook_search_mail`
  (thư chưa đọc), `outlook_get_mail` chạy đúng; đăng ký Claude Code (`claude mcp add`) báo `Connected`.

Chưa kiểm chứng:

- `outlook_list_events` và `outlook_search_contacts` trên dữ liệu thật (máy thử có lịch và danh bạ trống).
- Tài khoản Exchange/Microsoft 365 (đổi địa chỉ dạng `/o=...` sang SMTP) và profile nhiều hộp thư.
- Các tool còn lại trên Python 3.12; Python 3.10–3.11; cài offline bằng `-Wheelhouse`.
- Claude Desktop Chat/Cowork nạp server qua `claude_desktop_config.json` (đã thêm cấu hình, chưa thử gọi tool từ Cowork).
- Đọc `Body` thư IMAP có đổi trạng thái đã đọc / gửi read receipt không; CVE của phụ thuộc (chưa chạy `pip-audit`).

## 8. Giấy phép

Chưa gắn giấy phép mã nguồn mở: bạn được xem và tự cài để dùng, nhưng mọi quyền khác (sao chép, sửa, phân phối lại)
thuộc tác giả. Liên hệ tác giả qua GitHub Issues nếu cần.
