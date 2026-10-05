# AGENTS.md — Hướng dẫn cho Agent (Office Workspace)

Project này chuyên xử lý các tác vụ **văn phòng**: tạo, đọc, chỉnh sửa file Excel, PowerPoint và Word.

---

## 1. Nguyên tắc chung

- **Ngôn ngữ phản hồi**: Trả lời bằng tiếng Việt là chính; dùng tiếng Anh cho thuật ngữ kỹ thuật (e.g., `openpyxl`, `pptxgenjs`, `OOXML`).
- **Quy tắc lưu file và tổ chức thư mục (Folder & File Output)**:
  - **Mỗi lần làm việc BẮT BUỘC tạo 1 folder con riêng** bên trong thư mục tương ứng theo định dạng file:
    - `excel/YYYYMMDD_ten_tac_vu/`: Lưu bảng tính và dữ liệu dạng bảng (`.xlsx`, `.xlsm`, `.xltx`, `.csv`, `.tsv`).
    - `powerpoint/YYYYMMDD_ten_tac_vu/`: Lưu bài thuyết trình / slide (`.pptx`, `.potx`).
    - `word/YYYYMMDD_ten_tac_vu/`: Lưu tài liệu văn bản (`.docx`, `.dotx`).
    - *Quy ước đặt tên folder con*: `YYYYMMDD_ten_tac_vu/` (ví dụ: `20261005_mo_mang_cham/`). Tự động tạo thư mục nếu chưa tồn tại.
  - **Lưu trữ script hỗ trợ**: Nếu có viết script (Node.js, Python...) để xử lý dữ liệu hoặc sinh file Office, lưu script ngay bên trong thư mục con của tác vụ đó cùng với file kết quả để tiện theo dõi và tái lập.
  - **File chỉnh sửa**: Giữ nguyên đường dẫn/vị trí của file gốc (ghi đè hoặc tạo bản sao cùng thư mục với file gốc theo yêu cầu), trừ khi người dùng yêu cầu di chuyển.
- **Tuyệt đối KHÔNG TẠO FILE THỪA (PDF, hình ảnh...)**:
  - Không tự ý tạo hoặc lưu lại các file phụ trợ như `.pdf`, hình ảnh (`.png`, `.jpg`, `.jpeg`), HTML... trừ khi người dùng có yêu cầu rõ ràng.
  - *Đối với bước kiểm định trực quan (Visual QA)*: Được phép render tạm PDF/ảnh trong thư mục tạm (`scratch/`), nhưng **bắt buộc phải xóa sạch** trước khi bàn giao. Thư mục bàn giao chỉ giữ file Office chính (kèm script nếu có).
- **Không hỏi lại những gì đã rõ**: Nếu yêu cầu đủ rõ và có đầy đủ dữ liệu thực tế, thực hiện ngay và báo cáo kết quả sau.

---

## 2. Nguyên tắc BẤT BIẾN về Dữ liệu và Thông tin (Data Integrity)

> **ĐẶC BIỆT QUAN TRỌNG**: Agent tuyệt đối **KHÔNG ĐƯỢC PHÉP TỰ Ý BỊA SỐ LIỆU VÀ THÔNG TIN** dưới bất kỳ hình thức nào.

1. **Tuyệt đối không bịa đặt**:
   - Cấm tự chế các con số thống kê, doanh thu, lợi nhuận, chi phí, chỉ số KPI, ngày tháng, tên công ty/khách hàng hoặc thông tin nhân sự giả lập trông như thật khi người dùng chưa cung cấp.
2. **Xử lý khi thiếu số liệu/thông tin đầu vào**:
   - Khi người dùng yêu cầu làm báo cáo, tính toán, lập slide hay tài liệu nhưng chưa cung cấp đủ dữ liệu thực tế, Agent **bắt buộc phải dừng lại và hỏi người dùng** để nhận dữ liệu đầu vào trước khi tiến hành tạo file.
3. **Quy định đối với File Mẫu (Template / Khung mẫu)**:
   - Chỉ khi người dùng **chỉ định rõ ràng** là cần tạo "file mẫu", "template", hoặc "khung bảng tính", Agent mới tạo file khi chưa có số liệu thực tế.
   - Khi đó, các ô dữ liệu BẮT BUỘC phải dùng nhãn giữ chỗ rõ ràng (ví dụ: `[CẦN ĐIỀN]`, `[Chưa có dữ liệu]`, `TODO`, `N/A`) kèm công thức mẫu, tuyệt đối không gõ số giả lập gây nhầm lẫn là số liệu thật.
4. **Bảo toàn 100% độ chính xác khi trích xuất**:
   - Khi đọc từ file nguồn hoặc chuyển đổi giữa các định dạng (ví dụ từ Excel sang PowerPoint hay Word), phải bảo toàn nguyên vẹn số liệu gốc, không tự ý làm tròn số vô căn cứ, không tự suy diễn hay ngoại suy thông tin.

---

## 3. Kích hoạt Skill — Bắt buộc đọc trước khi thực hiện

> **QUAN TRỌNG**: Trước khi xử lý bất kỳ tác vụ văn phòng nào, Agent **phải đọc file `SKILL.md` tương ứng** bằng `view_file`. Không bỏ qua bước này dù đã quen với skill.

| Loại file | Skill cần đọc | Khi nào kích hoạt |
|---|---|---|
| `.xlsx`, `.xlsm`, `.xltx`, `.csv`, `.tsv` | `.agents/skills/xlsx/SKILL.md` | Tạo, đọc, chỉnh sửa bảng tính; làm sạch dữ liệu; tạo chart |
| `.pptx`, `.potx` | `.agents/skills/pptx/SKILL.md` | Tạo slide, sửa deck, đọc nội dung presentation |
| `.docx`, `.dotx` | `.agents/skills/docx/SKILL.md` | Tạo, đọc, sửa tài liệu Word; báo cáo, thư, memo |

### Ví dụ kích hoạt đúng

- *"Tạo file báo cáo doanh thu Q3"* → đọc `xlsx/SKILL.md`
- *"Làm slide pitch deck cho startup"* → đọc `pptx/SKILL.md`
- *"Viết hợp đồng dưới dạng Word"* → đọc `docx/SKILL.md`
- *"Đọc nội dung file presentation.pptx"* → đọc `pptx/SKILL.md` (kể cả khi chỉ đọc)

---

## 4. Phân biệt giữa các loại tác vụ

Một số yêu cầu có thể nhầm lẫn giữa các skill:

| Tình huống | Skill đúng | Lý do |
|---|---|---|
| Xuất dữ liệu từ database ra file Excel | `xlsx` | Deliverable là file `.xlsx` |
| Tạo báo cáo HTML có bảng số liệu | Không dùng skill nào | Deliverable là HTML, không phải Office file |
| Tóm tắt nội dung file `.pptx` | `pptx` | File `.pptx` là input, cần đọc theo đúng cách |
| Tạo tài liệu kỹ thuật dạng Markdown | Không dùng skill nào | Deliverable là `.md`, không phải `.docx` |

---

## 5. Xử lý lỗi khi file bị hỏng hoặc không mở được

Khi gặp lỗi liên quan đến file Office, thực hiện theo thứ tự sau:

### 5.1 File `.xlsx` / `.xlsm`
1. **Thử `markitdown file.xlsx`** để kiểm tra file có đọc được không.
2. Nếu `markitdown` lỗi → thử `pandas.read_excel()` với `engine='openpyxl'`.
3. Nếu vẫn lỗi → **unzip thủ công** (`unzip -o file.xlsx -d ./extracted_xlsx`) và kiểm tra `xl/workbook.xml`.
4. Báo cáo lỗi cụ thể cho người dùng, kèm đề xuất (e.g., "File có thể bị corrupt ở sheet X").

### 5.2 File `.pptx` / `.potx`
1. **Thử `markitdown file.pptx`** để đọc nội dung text.
2. Nếu lỗi → **unzip thủ công** (`unzip -o file.pptx -d ./extracted_pptx`) và kiểm tra `ppt/presentation.xml`.
3. Kiểm tra từng `ppt/slides/slideN.xml` nếu cần xác định slide bị lỗi.
4. Báo cáo lỗi cụ thể và đề xuất hướng khắc phục.

### 5.3 File `.docx` / `.dotx`
1. **Thử `pandoc -t markdown file.docx`** để đọc nội dung.
2. Nếu lỗi → **unzip thủ công** (`unzip -o file.docx -d ./extracted_docx`) và kiểm tra `word/document.xml`.
3. Báo cáo lỗi và đề xuất khắc phục.

### Nguyên tắc chung khi xử lý lỗi
- **Không tự ý xóa hoặc ghi đè file gốc** khi chưa được phép.
- Luôn **sao lưu file gốc** trước khi chỉnh sửa trực tiếp XML: `cp file.xlsx file.xlsx.bak`.
- Trình bày lỗi rõ ràng: tên file, loại lỗi, dòng/vị trí lỗi (nếu có).

---

## 6. Cấu trúc thư mục project

```
office/
├── AGENTS.md              ← File này
├── excel/                 ← Lưu bảng tính/dữ liệu (.xlsx, .xlsm, .xltx, .csv, .tsv)
│   └── YYYYMMDD_ten_tac_vu/   ← Mỗi lần làm việc tạo 1 folder con riêng
│       ├── bang_tinh.xlsx
│       └── script.py (nếu có)
├── powerpoint/            ← Lưu slide thuyết trình (.pptx, .potx)
│   └── YYYYMMDD_ten_tac_vu/   ← Mỗi lần làm việc tạo 1 folder con riêng
│       ├── presentation.pptx
│       └── create_deck.js (nếu có)
├── word/                  ← Lưu tài liệu văn bản (.docx, .dotx)
│   └── YYYYMMDD_ten_tac_vu/   ← Mỗi lần làm việc tạo 1 folder con riêng
│       ├── van_ban.docx
│       └── script.py (nếu có)
└── .agents/
    └── skills/
        ├── xlsx/
        │   └── SKILL.md   ← Skill xử lý Excel/CSV
        ├── pptx/
        │   └── SKILL.md   ← Skill xử lý PowerPoint
        └── docx/
            └── SKILL.md   ← Skill xử lý Word
```

---

## 7. Checklist trước khi hoàn thành tác vụ

Trước khi báo cáo hoàn thành, Agent tự kiểm tra:

- [ ] Đã đọc `SKILL.md` tương ứng trước khi thực hiện?
- [ ] **Trung thực dữ liệu**: Đảm bảo 100% số liệu và thông tin trung thực, không tự ý bịa đặt (đặc biệt quan trọng)?
- [ ] **Tổ chức thư mục con**: File tạo mới đã được đặt trong folder con riêng `YYYYMMDD_ten_tac_vu/` thuộc `excel/`, `powerpoint/`, hoặc `word/`?
- [ ] **Lưu script kèm theo**: Nếu có viết script xử lý hoặc sinh file, script đã được lưu trong folder con cùng với file Office?
- [ ] **Không sót file thừa**: Đã xóa sạch các file tạm (PDF, ảnh slide, XML tạm...) sau khi QA? Thư mục không chứa file thừa ngoài file Office chính và script?
- [ ] File output có thể mở được (không bị corrupt)?
- [ ] Đã xử lý hoặc báo cáo mọi lỗi phát sinh trong quá trình?
- [ ] Phản hồi bằng tiếng Việt?
