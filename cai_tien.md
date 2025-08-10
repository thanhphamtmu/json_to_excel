Để xuất file PCB_320902158_2021-12-16 08-43.json tốt hơn (dễ phân tích, không bị rối sheet, không mất dữ liệu lồng nhau), nên:
- Ưu tiên: mỗi list lớn (Contract, NotGrantedContract, Reference, Document, Profiles, ...) thành 1 sheet riêng, tên sheet rõ ràng, không trùng lặp.
- Sheet chính (main_info) chỉ chứa thông tin tổng quan, không nhồi các trường lồng nhau.
- Sheet con phải có cột liên kết về sheet cha (ví dụ: Contract phải có CBContractCode, Reference phải có IDCard hoặc CBSubjectCode...).

Để làm được:
1. Khi flatten_dict, nếu gặp list thì bỏ qua, không flatten vào main_info.
2. Khi xuất sheet con, thêm cột liên kết về cha (nếu có).
3. Sheet name: dùng key cha + key con nếu trùng tên (ví dụ: Subject.Matched.Document).

Tôi sẽ sửa code hàm export_multi_sheets và flatten_dict để:
- Không flatten list vào main_info.
- Sheet con có cột liên kết về cha.
- Sheet name rõ ràng, không trùng.

Tôi sẽ sửa trực tiếp file json2excel.py. 
```


Made changes.


```
Đã sửa code để:
- main_info chỉ chứa trường primitive, không flatten dict/list lồng nhau.
- Sheet con (ví dụ: Contract, Document, Reference, Profiles, ...) có tên rõ ràng, không trùng, theo dạng key path.
- Sheet con tự động thêm cột liên kết về cha (parent__...), giúp trace dữ liệu gốc.

Kết quả xuất file sẽ dễ phân tích, không bị rối, không mất dữ liệu lồng nhau, và dễ join giữa các sheet. Bạn chỉ cần chạy lại script với file JSON này để kiểm tra.
