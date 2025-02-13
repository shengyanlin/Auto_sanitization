import os
import csv
import time
import pandas as pd
from openpyxl import load_workbook

# ===== 1. 建立字典查表 (forward/backward) 用於字元位移 =====

_forward_map = {}
_backward_map = {}

# 小寫 a-z
for i in range(ord('a'), ord('z') + 1):
    c = chr(i)
    nxt_c = chr(i + 1) if i < ord('z') else 'a'
    _forward_map[c] = nxt_c
    _backward_map[nxt_c] = c  # 反向

# 大寫 A-Z
for i in range(ord('A'), ord('Z') + 1):
    c = chr(i)
    nxt_c = chr(i + 1) if i < ord('Z') else 'A'
    _forward_map[c] = nxt_c
    _backward_map[nxt_c] = c

# 數字 0-9
digits = '0123456789'
for i in range(10):
    c = digits[i]
    nxt_c = digits[(i + 1) % 10]  # 9 -> 0
    _forward_map[c] = nxt_c
    _backward_map[nxt_c] = c

def shift_char_forward(c: str) -> str:
    if c == '_':
        return c
    return _forward_map.get(c, c)

def shift_char_backward(c: str) -> str:
    if c == '_':
        return c
    return _backward_map.get(c, c)

# ===== 2. sanitize / desanitize 單一字串 =====

def sanitize(s: str) -> str:
    if s is None:
        return "N/A"

    s = str(s)
    if not s.strip() or s.lower() in ("nan", "none"):
        return "N/A"

    arr = list(s)
    length = len(arr)

    # 前 2
    i = 0
    shifted_start = 0
    while shifted_start < 2 and i < length:
        if arr[i] != ' ':
            arr[i] = shift_char_forward(arr[i])
            shifted_start += 1
        i += 1

    # 後 2
    j = length - 1
    shifted_end = 0
    while shifted_end < 2 and j >= 0:
        if arr[j] != ' ':
            arr[j] = shift_char_forward(arr[j])
            shifted_end += 1
        j -= 1

    return ''.join(arr)

def desanitize(s: str) -> str:
    if s is None:
        return ""

    s = str(s)
    arr = list(s)
    length = len(arr)

    # 前 2
    i = 0
    shifted_start = 0
    while shifted_start < 2 and i < length:
        if arr[i] != ' ':
            arr[i] = shift_char_backward(arr[i])
            shifted_start += 1
        i += 1

    # 後 2
    j = length - 1
    shifted_end = 0
    while shifted_end < 2 and j >= 0:
        if arr[j] != ' ':
            arr[j] = shift_char_backward(arr[j])
            shifted_end += 1
        j -= 1

    return ''.join(arr)

# ===== 3. memoization 快取 (在大量重複值時可顯著加速) =====

def memo_sanitize(val, cache):
    if val in cache:
        return cache[val]
    out = sanitize(val)
    cache[val] = out
    return out

def memo_desanitize(val, cache):
    if val in cache:
        return cache[val]
    out = desanitize(val)
    cache[val] = out
    return out

# ===== 4. 在 Pandas DataFrame 中插入 *_sanitized 或 External Part =====

def insert_sanitized_columns(df: pd.DataFrame, columns_to_add: list):
    """
    保留原欄位，並在右側插入 col_sanitized
    """
    columns_with_locs = [(col, df.columns.get_loc(col)) for col in columns_to_add]
    columns_with_locs.sort(key=lambda x: x[1], reverse=True)
    
    for col, loc_val in columns_with_locs:
        df[col] = df[col].fillna('N/A') \
                         .replace(r'^\s*$', 'N/A', regex=True) \
                         .replace(['none', 'nan'], 'N/A')
        cache = {}
        new_col_name = f"{col}_sanitized"
        df.insert(loc_val + 1, new_col_name, df[col].apply(lambda x: memo_sanitize(x, cache)))


def insert_desanitized_columns(df: pd.DataFrame, columns_to_process: list):
    """
    保留原欄位 'External part ID'，右邊插一個 'External Part' 放 desanitize 後的值
    """
    columns_with_locs = [(col, df.columns.get_loc(col)) for col in columns_to_process]
    columns_with_locs.sort(key=lambda x: x[1], reverse=True)

    for col, loc_val in columns_with_locs:
        if col == 'External part ID':
            cache = {}
            df.insert(loc_val + 1, 'External Part', df[col].apply(lambda x: memo_desanitize(x, cache)))

# ===== 5. Streaming模式下 (read_only=True) —— 保留原值，並在右側插欄 =====

def process_xlsx_sanitization_streaming(input_path: str, output_folder: str):
    """
    只讀模式逐 Sheet 讀取 .xlsx，
    如果該檔只有1個 Sheet => 檔名不加 sheetName
    否則 => baseName_sheetName_sanitized.csv

    **在需要sanitize的欄位右側，插入 col_sanitized 欄位**。
    """
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    wb_in = load_workbook(filename=input_path, read_only=True, data_only=True)

    # 要插入 sanitized 的欄位
    columns_to_add = [
        'External Part',
        'Internal Part',
        'Internal Part(Old)',
        'Internal Part(New)',
        'ex_part'
    ]

    sheet_names = wb_in.sheetnames
    sheet_count = len(sheet_names)

    for i, sheet_name in enumerate(sheet_names, start=1):
        ws_in = wb_in[sheet_name]

        # 輸出檔名
        if sheet_count == 1:
            out_filename = f"{base_name}_sanitized.csv"
        else:
            out_filename = f"{base_name}_{sheet_name}_sanitized.csv"
        out_path = os.path.join(output_folder, out_filename)

        print(f"[SANITIZE] Processing sheet {i}/{sheet_count}: {sheet_name} -> {out_filename}")

        with open(out_path, mode='w', newline='', encoding='utf-8-sig') as f_out:
            writer = csv.writer(f_out)

            row_idx = 0
            # 用於判斷哪幾個欄位需要多加一個 {col}_sanitized
            sanitize_indexes = []
            # 建個 cache_dict 供 memo_sanitize 使用
            cache_dict = {}

            for row_tuple in ws_in.iter_rows(values_only=True):
                row_list = list(row_tuple) if row_tuple else []

                if row_idx == 0:
                    # 第一行 => 表頭
                    new_header = []
                    for col_i, col_name in enumerate(row_list):
                        new_header.append(col_name)  # 保留原欄位名
                        if col_name in columns_to_add:
                            # 在右邊插入 col_name + "_sanitized"
                            sanitize_indexes.append(col_i)
                            new_header.append(col_name + "_sanitized")
                    writer.writerow(new_header)

                else:
                    new_row = []
                    for col_i, val in enumerate(row_list):
                        # 先把原值放進去
                        new_row.append(val)
                        # 若該欄位需要 sanitize，就多插入一格
                        if col_i in sanitize_indexes:
                            new_val = memo_sanitize(val, cache_dict)
                            new_row.append(new_val)
                    writer.writerow(new_row)

                row_idx += 1

    wb_in.close()


def process_xlsx_desanitization_streaming(input_path: str, output_folder: str):
    """
    只讀模式逐 Sheet 讀取 .xlsx，
    如果該檔只有1個 Sheet => 檔名不加 sheetName
    否則 => baseName_sheetName_desanitized.csv

    針對 'External part ID' 欄位，在其右側插入 'External Part'
    """
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    wb_in = load_workbook(filename=input_path, read_only=True, data_only=True)
    sheet_names = wb_in.sheetnames
    sheet_count = len(sheet_names)

    for i, sheet_name in enumerate(sheet_names, start=1):
        ws_in = wb_in[sheet_name]
        if sheet_count == 1:
            out_filename = f"{base_name}_desanitized.csv"
        else:
            out_filename = f"{base_name}_{sheet_name}_desanitized.csv"
        out_path = os.path.join(output_folder, out_filename)

        print(f"[DESANITIZE] Processing sheet {i}/{sheet_count}: {sheet_name} -> {out_filename}")

        with open(out_path, mode='w', newline='', encoding='utf-8-sig') as f_out:
            writer = csv.writer(f_out)

            row_idx = 0
            desanitize_indexes = []
            cache_dict = {}

            for row_tuple in ws_in.iter_rows(values_only=True):
                row_list = list(row_tuple) if row_tuple else []

                if row_idx == 0:
                    # 表頭
                    new_header = []
                    for col_i, col_name in enumerate(row_list):
                        new_header.append(col_name)
                        # 如果欄位名 == 'External part ID'，右邊插入 'External Part'
                        if col_name == 'External part ID':
                            desanitize_indexes.append(col_i)
                            new_header.append("External Part")
                    writer.writerow(new_header)
                else:
                    new_row = []
                    for col_i, val in enumerate(row_list):
                        new_row.append(val)
                        if col_i in desanitize_indexes:
                            # 產生 External Part 欄位
                            new_val = memo_desanitize(val, cache_dict)
                            new_row.append(new_val)
                    writer.writerow(new_row)

                row_idx += 1

    wb_in.close()

# ===== 6. 若是 .xlsx => streaming / 若是 .csv => DataFrame =====

def process_file_sanitization(file_path: str, output_folder: str):
    base_name, ext = os.path.splitext(os.path.basename(file_path))

    if ext.lower() == '.xlsx':
        process_xlsx_sanitization_streaming(file_path, output_folder)
    elif ext.lower() == '.csv':
        try:
            df = pd.read_csv(file_path, encoding='utf-8')
        except:
            df = pd.read_csv(file_path, encoding='gbk')

        # 對 CSV 用 DataFrame 插欄
        columns_to_add = []
        if 'External Part' in df.columns:
            columns_to_add.append('External Part')
        if 'Internal Part' in df.columns:
            columns_to_add.append('Internal Part')
        if 'Internal Part(Old)' in df.columns:
            columns_to_add.append('Internal Part(Old)')
        if 'Internal Part(New)' in df.columns:
            columns_to_add.append('Internal Part(New)')

        if columns_to_add:
            insert_sanitized_columns(df, columns_to_add)

        out_filename = base_name + "_sanitized.csv"
        out_path = os.path.join(output_folder, out_filename)
        df.to_csv(out_path, index=False, encoding='utf-8-sig')


def process_file_desanitization(file_path: str, output_folder: str):
    base_name, ext = os.path.splitext(os.path.basename(file_path))

    if ext.lower() == '.xlsx':
        process_xlsx_desanitization_streaming(file_path, output_folder)
    elif ext.lower() == '.csv':
        try:
            df = pd.read_csv(file_path, encoding='utf-8')
        except:
            df = pd.read_csv(file_path, encoding='gbk')

        columns_to_process = []
        if 'External part ID' in df.columns:
            columns_to_process.append('External part ID')

        if columns_to_process:
            insert_desanitized_columns(df, columns_to_process)
        else:
            print(f"The file {file_path} does not contain sanitized columns.")

        out_filename = base_name + "_desanitized.csv"
        out_path = os.path.join(output_folder, out_filename)
        df.to_csv(out_path, index=False, encoding='utf-8-sig')

# ===== 7. Main workflow =====

def sanitize_data():
    print("\nSanitizing process started...")

    input_folder = 'Unsanitized'
    output_folder = 'Sanitized'
    os.makedirs(output_folder, exist_ok=True)

    file_list = [f for f in os.listdir(input_folder) if f.endswith(('.xlsx', '.csv'))]
    if not file_list:
        print("No .xlsx or .csv files found in 'Unsanitized' folder.")
        return

    print("\nFollowing files will be sanitized:")
    for file_name in file_list:
        print(f" - {file_name}")
    print()

    for file_name in file_list:
        file_path = os.path.join(input_folder, file_name)
        print(f"Processing file: {file_name}")
        process_file_sanitization(file_path, output_folder)

    print("\nSanitizing process completed.")


def desanitize_data():
    print("\nDesanitizing process started...")

    input_folder = 'Undesanitized'
    output_folder = 'Desanitized'
    os.makedirs(output_folder, exist_ok=True)

    file_list = [f for f in os.listdir(input_folder) if f.endswith(('.xlsx', '.csv'))]
    if not file_list:
        print("No .xlsx or .csv files found in 'Undesanitized' folder.")
        return

    print("\nFollowing files will be desanitized:")
    for file_name in file_list:
        print(f" - {file_name}")
    print()

    for file_name in file_list:
        file_path = os.path.join(input_folder, file_name)
        process_file_desanitization(file_path, output_folder)

    print("\nDesanitizing process completed.")


def main():
    print("Want to sanitize data? (y/n)")
    user_input = input().strip().lower()

    start_time = time.time()

    if user_input == "y":
        sanitize_data()
    elif user_input == "n":
        desanitize_data()
    else:
        print("Invalid input, please enter 'y' or 'n'.")

    #print time taken in minutes
    print(f"\nTime taken: {round((time.time() - start_time) / 60, 2)} minutes.")
    input("\nProcessing complete. Press Enter to exit...")

if __name__ == "__main__":
    main()
