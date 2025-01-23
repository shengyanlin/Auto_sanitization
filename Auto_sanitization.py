import os
import pandas as pd

# ===== Character shifting utility functions =====

def shift_char_forward(c: str) -> str:
    """
    Shift letters and digits forward by one.
    For example: 'a' -> 'b', 'z' -> 'a', '9' -> '0'.
    Underscore '_' is unaffected; other symbols are returned as is.
    """
    if c.isalpha():
        if c.islower():
            return 'a' if c == 'z' else chr(ord(c) + 1)
        else:
            return 'A' if c == 'Z' else chr(ord(c) + 1)
    elif c.isdigit():
        return '0' if c == '9' else str(int(c) + 1)
    else:
        # '_' remains unaffected; other symbols are returned as is
        return c

def shift_char_backward(c: str) -> str:
    """
    Shift letters and digits backward by one.
    For example: 'b' -> 'a', 'a' -> 'z', '0' -> '9'.
    Underscore '_' is unaffected; other symbols are returned as is.
    """
    if c.isalpha():
        if c.islower():
            return 'z' if c == 'a' else chr(ord(c) - 1)
        else:
            return 'Z' if c == 'A' else chr(ord(c) - 1)
    elif c.isdigit():
        return '9' if c == '0' else str(int(c) - 1)
    else:
        # '_' remains unaffected; other symbols are returned as is
        return c

# ===== sanitize / desanitize 單一字串 =====

def sanitize(s: str) -> str:
    """
    Shift the first two and the last two non-space characters forward by one.
    如果是空白、NaN 或 None，則回傳 "N/A"。
    """
    # 先把 s 轉成字串
    s = str(s)

    # 如果為空字串、只有空白、或是 "nan"/"none" 都直接回傳 "N/A"
    # 雖然我們在外部流程中也會先做 fillna/replace，但這裡保留容錯處理。
    if not s.strip() or s.lower() in ("nan", "none"):
        return "N/A"

    s_list = list(s)
    length = len(s_list)

    # Process the first two non-space characters
    i = 0
    shifted_start = 0
    while shifted_start < 2 and i < length:
        if s_list[i] != ' ':
            s_list[i] = shift_char_forward(s_list[i])
            shifted_start += 1
        i += 1
    
    # Process the last two non-space characters
    j = length - 1
    shifted_end = 0
    while shifted_end < 2 and j >= 0:
        if s_list[j] != ' ':
            s_list[j] = shift_char_forward(s_list[j])
            shifted_end += 1
        j -= 1
    
    return ''.join(s_list)

def desanitize(s: str) -> str:
    """
    Shift the first two and the last two non-space characters backward by one.
    """
    s = str(s)
    s_list = list(s)
    length = len(s_list)
    
    # Process the first two non-space characters
    i = 0
    shifted_start = 0
    while shifted_start < 2 and i < length:
        if s_list[i] != ' ':
            s_list[i] = shift_char_backward(s_list[i])
            shifted_start += 1
        i += 1
    
    # Process the last two non-space characters
    j = length - 1
    shifted_end = 0
    while shifted_end < 2 and j >= 0:
        if s_list[j] != ' ':
            s_list[j] = shift_char_backward(s_list[j])
            shifted_end += 1
        j -= 1
    
    return ''.join(s_list)

# ===== Utility functions for column operations =====

def insert_sanitized_columns(df: pd.DataFrame, columns_to_add: list):
    """
    在 df 中，對於 columns_to_add 裏列出的欄位，緊接在它後面插入一個對應的 {col}_sanitized 欄位
    (透過 sanitize() 函式來產生新值)。
    依欄位的 index 由大到小插入，避免影響後面欄位的插入位置。

    除了對指定欄位做字串位移，會先以向量化方式把空值、空白、nan、none 轉成 'N/A'。
    """
    # 先依照原欄位在 df.columns 中的索引位置，從「大到小」排序
    columns_with_locs = [(col, df.columns.get_loc(col)) for col in columns_to_add]
    columns_with_locs.sort(key=lambda x: x[1], reverse=True)
    
    for col, loc_val in columns_with_locs:
        # 先用向量化方法處理空白、nan、none → 'N/A'
        df[col] = df[col].fillna('N/A') \
                         .replace(r'^\s*$', 'N/A', regex=True) \
                         .replace(['none', 'nan'], 'N/A')

        new_col_name = f"{col}_sanitized"
        # 再逐列呼叫 sanitize()
        df.insert(loc_val + 1, new_col_name, df[col].apply(sanitize))

def insert_desanitized_columns(df: pd.DataFrame, columns_to_process: list):
    """
    在 df 中，對於 columns_to_process 裏列出的欄位，緊接在它後面插入一個對應的 'External Part' 欄位
    (透過 desanitize() 函式來產生新值)。
    依欄位的 index 由大到小插入，避免影響後面欄位的插入位置。
    """
    columns_with_locs = [(col, df.columns.get_loc(col)) for col in columns_to_process]
    columns_with_locs.sort(key=lambda x: x[1], reverse=True)

    for col, loc_val in columns_with_locs:
        if col == 'External part ID':
            df.insert(loc_val + 1, 'External Part', df[col].apply(desanitize))

# ===== New: 逐檔讀寫檔案 (不再一次全部讀進記憶體) =====

def process_file_sanitization(file_path: str, output_folder: str):
    """
    讀取單一檔案 (xlsx/csv)，對需要的欄位執行 insert_sanitized_columns 後，立刻寫出。
    避免大量檔案同時存在記憶體中。
    """
    base_name, ext = os.path.splitext(os.path.basename(file_path))
    output_file_name = base_name + "_sanitized" + ext
    output_path = os.path.join(output_folder, output_file_name)

    if ext.lower() == '.xlsx':
        excel_file = pd.ExcelFile(file_path, engine='openpyxl')
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                
                #print all column names in this sheet
                print(f"Sheet: {sheet_name} - Columns: {df.columns}")

                # 需要插入 sanitized 欄位的欄位清單
                columns_to_add = []
                if 'External Part' in df.columns:
                    columns_to_add.append('External Part')
                if 'Internal Part' in df.columns:
                    columns_to_add.append('Internal Part')
                if 'Internal Part(Old)' in df.columns:
                    columns_to_add.append('Internal Part(Old)')
                if 'Internal Part(New)' in df.columns:
                    columns_to_add.append('Internal Part(New)')
                if "ex_part" in df.columns:
                    columns_to_add.append("ex_part")

                print("Columns to add: ", columns_to_add)
                if columns_to_add:
                    insert_sanitized_columns(df, columns_to_add)

                df.to_excel(writer, sheet_name=sheet_name, index=False)

    elif ext.lower() == '.csv':
        try:
            df = pd.read_csv(file_path, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(file_path, encoding='gbk')

        columns_to_add = []
        if 'External Part' in df.columns:
            columns_to_add.append('External Part')
        if 'Internal Part' in df.columns:
            columns_to_add.append('Internal Part')
        if 'Internal Part(Old)' in df.columns:
            columns_to_add.append('Internal Part(Old)')
        if 'Internal Part(New)' in df.columns:
            columns_to_add.append('Internal Part(New)')
        if "ex_part" in df.columns:
            columns_to_add.append("ex_part")

        if columns_to_add:
            insert_sanitized_columns(df, columns_to_add)

        df.to_csv(output_path, index=False, encoding='utf-8-sig')

def process_file_desanitization(file_path: str, output_folder: str):
    """
    讀取單一檔案 (xlsx/csv)，對需要的欄位執行 insert_desanitized_columns 後，立刻寫出。
    注意此需求中，檔名不改變 (不加 _sanitized)。
    """
    base_name, ext = os.path.splitext(os.path.basename(file_path))
    output_path = os.path.join(output_folder, base_name + ext)  # 不改變檔名

    if ext.lower() == '.xlsx':
        excel_file = pd.ExcelFile(file_path, engine='openpyxl')
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)

                columns_to_process = []
                if 'External part ID' in df.columns:
                    columns_to_process.append('External part ID')

                if columns_to_process:
                    insert_desanitized_columns(df, columns_to_process)
                else:
                    print(f"The file {file_path} (Sheet: {sheet_name}) does not contain sanitized columns.")

                df.to_excel(writer, sheet_name=sheet_name, index=False)

    elif ext.lower() == '.csv':
        try:
            df = pd.read_csv(file_path, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(file_path, encoding='gbk')

        columns_to_process = []
        if 'External part ID' in df.columns:
            columns_to_process.append('External part ID')

        if columns_to_process:
            insert_desanitized_columns(df, columns_to_process)
        else:
            print(f"The file {file_path} does not contain sanitized columns.")

        df.to_csv(output_path, index=False, encoding='utf-8-sig')

# ===== Main workflow functions =====

def sanitize_data():
    """
    1. 從 'Unsanitized' 資料夾讀取所有 .xlsx/.csv 檔案 (檔案大也不怕，一次處理一檔)
    2. 如果 'External Part'/'Internal Part'/'Internal Part(Old)'/'Internal Part(New)' 欄位存在，就插入對應的 '*_sanitized' 欄位
    3. 將結果輸出到 'Sanitized' 資料夾，檔名加上 "_sanitized"
    """
    print("\nSanitizing process started...")

    input_folder = 'Unsanitized'
    output_folder = 'Sanitized'

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 收集所有要處理的檔案
    file_list = [
        f for f in os.listdir(input_folder)
        if f.endswith(('.xlsx', '.csv'))
    ]

    if not file_list:
        print("No .xlsx or .csv files found in 'Unsanitized' folder.")
        return

    print("\nFollowing files will be sanitized:")
    for file_name in file_list:
        print(f" - {file_name}")
    # print()

    # 逐檔處理並輸出
    for file_name in file_list:
        file_path = os.path.join(input_folder, file_name)
        print(f"\nProcessing file: {file_name}")
        process_file_sanitization(file_path, output_folder)

    print("\nSanitizing process completed.")
    

def desanitize_data():
    """
    1. 從 'Undesanitized' 資料夾讀取所有 .xlsx/.csv 檔案 (檔案大也不怕，一次處理一檔)
    2. 如果 'External part ID' 欄位存在，就插入對應的 'External Part' 欄位 (desanitize)
    3. 將結果輸出到 'Desanitized' 資料夾 (檔名不更動)
    """
    print("\nDesanitizing process started...")

    input_folder = 'Undesanitized'
    output_folder = 'Desanitized'

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    file_list = [
        f for f in os.listdir(input_folder)
        if f.endswith(('.xlsx', '.csv'))
    ]

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
    """
    Main function to prompt user input and run either Sanitization or Desanitization process.
    """
    print("Want to sanitize data? (y/n)")
    user_input = input().strip().lower()

    if user_input == "y":
        sanitize_data()
    elif user_input == "n":
        desanitize_data()
    else:
        print("Invalid input, please enter 'y' or 'n'.")
    
    input("\nProcessing complete. Press Enter to exit...")

if __name__ == "__main__":
    main()

# ex_part、