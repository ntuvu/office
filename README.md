# Office Workspace

Workspace tự động hóa và xử lý các tác vụ tài liệu văn phòng: tạo, đọc, và chỉnh sửa file Microsoft Excel, PowerPoint, và Word.

## Cấu trúc thư mục

```
office/
├── AGENTS.md              # Hướng dẫn và quy tắc cho AI Agent
├── .gitignore             # Danh sách loại trừ file tạm, cache
├── README.md              # Tài liệu giới thiệu dự án
├── excel/                 # Thư mục lưu bảng tính (.xlsx, .xlsm, .csv...)
├── powerpoint/            # Thư mục lưu bài thuyết trình (.pptx, .potx...)
├── word/                  # Thư mục lưu tài liệu văn bản (.docx, .dotx...)
└── .agents/
    └── skills/            # Bộ kỹ năng (Skills) tự động hóa
        ├── xlsx/          # Kỹ năng xử lý Excel/CSV
        ├── pptx/          # Kỹ năng xử lý PowerPoint
        └── docx/          # Kỹ năng xử lý Word
```

## Các kỹ năng hỗ trợ

- **Excel (`xlsx`)**: Tạo, đọc, xử lý dữ liệu bảng tính, tính toán công thức, định dạng biểu đồ và làm sạch dữ liệu.
- **PowerPoint (`pptx`)**: Tạo bài thuyết trình chuyên nghiệp, thiết kế bố cục slide, trích xuất và chỉnh sửa nội dung presentation.
- **Word (`docx`)**: Tạo tài liệu văn bản chuẩn cấu trúc, hợp đồng, báo cáo, quản lý comments và tracked changes.
