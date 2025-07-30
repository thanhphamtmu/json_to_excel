# 🧾 JSON to Excel Converter (Python)

Công cụ chuyển đổi file JSON sang Excel (.xlsx) bằng Python, hỗ trợ:

- ✅ Dữ liệu JSON dạng `dict` hoặc `list`
- ✅ Gộp danh sách lồng nhau thành bảng
- ✅ Tách nhiều sheet trong Excel nếu cần
- ✅ Xử lý tên sheet hợp lệ cho Excel
- ✅ Chạy từ dòng lệnh, không cần GUI

---

## 🚀 Tính năng chính

| Tính năng                    | Mô tả                                                       |
| ---------------------------- | ----------------------------------------------------------- |
| ✅ Hỗ trợ dict hoặc list     | Chuyển đổi linh hoạt mọi cấu trúc JSON phổ biến             |
| ✅ Tách nhiều sheet          | Mỗi danh sách con (nested list) thành 1 sheet riêng         |
| ✅ Gộp danh sách             | Gộp danh sách lồng nhau với phần tử gốc vào 1 bảng duy nhất |
| ✅ Chọn nhiều file           | Cho phép chọn nhiều file từ thư mục                         |
| ✅ Tránh lỗi tên sheet       | Tự động sửa tên sheet hợp lệ với Excel                      |
| ✅ Xuất nhiều file từ 1 JSON | Nếu dữ liệu là danh sách chứa nhiều dict                    |

---

## ⚙️ Yêu cầu cài đặt

- Python >= 3.8
- Các thư viện:

```bash
pip install pandas openpyxl
```

---

## 🧑‍💻 Cách sử dụng

### 1. Chạy tool:

```bash
python json2excel.py
```

### 2. Nhập đường dẫn:

```bash
📂 Nhập đường dẫn file hoặc thư mục JSON: ./du_lieu/
```

### 3. Lựa chọn:

- Nếu là file:
  ```text
  1. Convert file
  2. Convert folder
  q. Thoát
  👉
  ```
- Nếu là folder:
  - Hiển thị danh sách file `.json`
  - Bạn có thể chọn tất cả hoặc nhập số tương ứng (vd: 1,3,5)

### 4. Khi gặp JSON có danh sách lồng nhau:

```text
📌 File 'data.json' có danh sách lồng nhau trong dict.
1. 1 sheet (gộp danh sách)
2. Xuất nhiều sheet
q. Thoát
👉
```

---

## 📦 Kết quả xuất ra

- File Excel `.xlsx` sẽ được tạo trong cùng thư mục với file JSON.
- Tên file tương ứng với tên file JSON ban đầu.

---

## 📝 Ví dụ

### JSON dạng dict:

```json
{
  "name": "Dự án A",
  "tasks": [
    { "task": "Viết code", "hours": 5 },
    { "task": "Test", "hours": 2 }
  ]
}
```

#### ✅ Kết quả `1 sheet (gộp danh sách)`:

| name    | task      | hours |
| ------- | --------- | ----- |
| Dự án A | Viết code | 5     |
| Dự án A | Test      | 2     |

#### ✅ Kết quả `nhiều sheet`:

- Sheet 1: `main_info` → `name`
- Sheet 2: `tasks` → bảng chứa task

---

## ⚠️ Lưu ý

- Excel chỉ cho phép tên sheet dài tối đa 31 ký tự và không chứa các ký tự: `: \ / ? * [ ]`
- Nếu sheet name bị trùng, chương trình sẽ tự thêm hậu tố `_1`, `_2`, v.v.
- Nếu file JSON chứa danh sách với nhiều dict phức tạp → có thể chọn xuất từng item thành file riêng biệt (chế độ nhiều sheet)

---

## 📂 Tổ chức thư mục

```
json2excel/
├── json2excel.py       # Mã nguồn chính
├── README.md           # Hướng dẫn sử dụng
├── du_lieu/            # Chứa các file JSON để convert
```

---

## 🧩 Gợi ý mở rộng

- Cho phép export CSV
- Cho phép chọn encoding
- Hỗ trợ dòng lệnh (`argparse`) để tự động hóa
- Tích hợp giao diện web nhỏ bằng Streamlit

---

## 📫 Liên hệ

Nếu bạn gặp lỗi hoặc cần tính năng mở rộng, vui lòng tạo issue hoặc liên hệ mình qua email/zalo.
