import json
import re
import traceback
from pathlib import Path

import pandas as pd


def normalize_to_list(obj):
    return obj if isinstance(obj, list) else [obj] if isinstance(obj, dict) else []


def flatten_dict(d, parent_key="", sep="."):
    items = []
    for k, v in d.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            items.extend(flatten_dict(v, new_key, sep=sep).items())
        elif isinstance(v, list):
            # Không flatten list vào main_info
            continue
        else:
            items.append((new_key, v))
    return dict(items)


def find_all_lists_in_dict(data: dict) -> dict:
    result = {}

    def recursive(d, parent_key=""):
        if isinstance(d, dict):
            for k, v in d.items():
                new_key = f"{parent_key}.{k}" if parent_key else k
                if isinstance(v, dict) and any(
                    isinstance(i, (dict, list)) for i in v.values()
                ):
                    recursive(v, new_key)
                elif isinstance(v, list):
                    normalized = normalize_to_list(v)
                    if normalized and all(isinstance(i, dict) for i in normalized):
                        result[new_key] = normalized
                elif isinstance(v, dict):
                    result[new_key] = [v]
        elif isinstance(d, list):
            for idx, item in enumerate(d):
                recursive(item, f"{parent_key}[{idx}]")

    recursive(data)
    return result


def sanitize_sheet_name(name):
    # Ưu tiên phần cuối key path, rút gọn về 31 ký tự
    name = re.sub(r"[:\\/?*\[\]]", "_", name)
    if len(name) <= 31:
        return name
    # Nếu quá dài, lấy phần cuối (sau dấu _), nếu vẫn dài thì lấy 31 ký tự cuối
    parts = name.split('_')
    for i in range(len(parts)-1, -1, -1):
        candidate = '_'.join(parts[i:])
        if len(candidate) <= 31:
            return candidate
    return name[-31:]


def export_multi_sheets(data, output_path):
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            # main_info chỉ lấy trường primitive, không flatten dict/list
            def get_primitives(d):
                return {k: v for k, v in d.items() if not isinstance(v, (dict, list))}

            flat_info = get_primitives(data)
            list_fields = find_all_lists_in_dict(data)

            if flat_info:
                pd.DataFrame([flat_info]).to_excel(
                    writer, sheet_name="main_info", index=False
                )

            used_names = set(["main_info"])
            for key, value in list_fields.items():
                # Tên sheet: lấy đầy đủ key path, thay . bằng _
                sheet = sanitize_sheet_name(key.replace(".", "_"))
                base = sheet
                i = 1
                while sheet in used_names:
                    sheet = f"{base}_{i}"
                    i += 1
                used_names.add(sheet)

                # Tìm key cha để liên kết (nếu có)
                parent_keys = key.split(".")[:-1]
                parent_info = data
                for pk in parent_keys:
                    if isinstance(parent_info, dict):
                        parent_info = parent_info.get(pk, {})
                    else:
                        parent_info = {}
                parent_flat = flatten_dict(parent_info) if isinstance(parent_info, dict) else {}

                rows = []
                for item in normalize_to_list(value):
                    row = flatten_dict(item)
                    # Thêm cột liên kết về cha nếu có
                    for pk, pv in parent_flat.items():
                        if pk not in row:
                            row["parent__" + pk] = pv
                    rows.append(row)
                df = pd.DataFrame(rows)
                df.to_excel(writer, sheet_name=sheet, index=False)
    except Exception as e:
        print(f"❌ Lỗi xuất sheet: {e}")
        traceback.print_exc()


def handle_dict_data(data: dict, output_path: Path):
    lists_found = find_all_lists_in_dict(data)
    if lists_found:
        choice = (
            input("1. 1 sheet (gộp danh sách)\n2. Xuất nhiều sheet\nq. Thoát\n👉 ")
            .strip()
            .lower()
        )
        if choice == "2":
            export_multi_sheets(data, output_path)
            print(f"✅ Xuất ra: {output_path.name}")
            return
        elif choice in ["q", "quit"]:
            print("❌ Đã bỏ qua.")
            return
        else:
            flat_info = flatten_dict(
                {k: v for k, v in data.items() if not isinstance(v, (list, dict))}
            )
            records = []

            for key, value in lists_found.items():
                for item in normalize_to_list(value):
                    combined = {**flat_info, **flatten_dict(item)}
                    records.append(combined)

            if not records:
                records.append(flat_info)

            pd.DataFrame(records).to_excel(output_path, index=False, engine="openpyxl")
            print(f"✅ Xuất ra: {output_path.name}")
    else:
        pd.DataFrame([flatten_dict(data)]).to_excel(
            output_path, index=False, engine="openpyxl"
        )
        print(f"✅ Xuất ra: {output_path.name}")


def handle_list_data(data: list, json_path: Path, output_dir: Path):
    nested_keys = [k for k, v in data[0].items() if isinstance(v, list)]
    output_path = output_dir / (json_path.stem + ".xlsx")

    if nested_keys and len(nested_keys) > 1:
        choice = (
            input("1. 1 sheet\n2. Nhiều sheet (1 entry/sheet)\nq. Thoát\n👉 ")
            .strip()
            .lower()
        )
        if choice == "2":
            for i, entry in enumerate(data):
                export_multi_sheets(
                    entry, output_dir / f"{json_path.stem}_{i + 1}.xlsx"
                )
            return
        elif choice in ["q", "quit"]:
            print("❌ Đã bỏ qua.")
            return

    records = []
    for entry in data:
        flat_info = flatten_dict(
            {k: v for k, v in entry.items() if not isinstance(v, list)}
        )
        if nested_keys:
            for key in nested_keys:
                nested = normalize_to_list(entry.get(key, []))
                for item in nested:
                    records.append({**flat_info, **flatten_dict(item)})
        else:
            records.append(flat_info)

    if records:
        pd.DataFrame(records).to_excel(output_path, index=False, engine="openpyxl")
        print(f"✅ Xuất ra: {output_path.name}")
    else:
        print("⚠️ Không có dữ liệu để ghi.")


def convert_json_to_excel(json_path: Path, output_dir: Path):
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"❌ Lỗi đọc file {json_path.name}: {e}")
        return

    if isinstance(data, dict):
        handle_dict_data(data, output_dir / f"{json_path.stem}.xlsx")
    elif isinstance(data, list) and data and isinstance(data[0], dict):
        handle_list_data(data, json_path, output_dir)
    else:
        print("⚠️ JSON không hợp lệ.")


def convert_folder(folder: Path):
    json_files = sorted(folder.glob("*.json"))
    if not json_files:
        print("❌ Không tìm thấy file JSON.")
        return

    print("\n📄 File JSON có:")
    for i, file in enumerate(json_files, 1):
        print(f"{i}. {file.name}")

    selection = input(
        "\n🔢 Nhập số file cần convert (ngăn cách bằng phẩy), trống = tất cả: "
    ).strip()
    if selection.lower() in ["q", "quit"]:
        return

    if selection:
        try:
            indices = [int(i.strip()) - 1 for i in selection.split(",")]
            json_files = [json_files[i] for i in indices]
        except Exception as e:
            print(f"❌ Lỗi chọn file: {e}")
            return

    print(f"\n🔄 Đang convert {len(json_files)} file...\n")
    for file in json_files:
        convert_json_to_excel(file, folder)


def main():
    path_str = input("📂 Nhập đường dẫn file hoặc thư mục JSON: ").strip('" ')
    if path_str.lower() in ["q", "quit"]:
        return

    path = Path(path_str)
    if not path.exists():
        print("❌ Đường dẫn không tồn tại.")
        return

    if path.is_file():
        choice = (
            input("1. Convert file\n2. Convert folder\nq. Thoát\n👉 ").strip().lower()
        )
        if choice == "1":
            convert_json_to_excel(path, path.parent)
        elif choice == "2":
            convert_folder(path.parent)
    elif path.is_dir():
        convert_folder(path)
    else:
        print("❌ Định dạng không hỗ trợ.")


if __name__ == "__main__":
    main()
