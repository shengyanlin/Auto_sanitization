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

def sanitize(s: str) -> str:
    """
    Shift the first two and the last two non-space characters forward by one.
    如果是空白、NaN 或 None，則回傳 "N/A"
    """
    # 先把 s 轉成字串
    s = str(s)

    # 如果為空字串、只有空白、或是 "nan"/"none" 都直接回傳 "N/A"
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

# ===== Utility functions for reading/writing data =====

def read_data_frames(input_folder: str) -> dict:
    """
    從指定資料夾中讀取所有 .xlsx 與 .csv 檔，回傳一個 dict 結構：
    {
      "file1.xlsx": {
          'format': 'xlsx',
          'sheets': {
              'Sheet1': df1,
              'Sheet2': df2,
          }
      },
      "file2.csv": {
          'format': 'csv',
          'df': df_csv
      }
    }
    也會在過程中印出讀取進度。
    """
    if not os.path.exists(input_folder):
        os.makedirs(input_folder)

    file_list = [f for f in os.listdir(input_folder) if f.endswith(('.xlsx', '.csv'))]

    data_frames = {}
    
    print("\nFollowing files are read:")
    for file in file_list:
        file_path = os.path.join(input_folder, file)
        if file.endswith('.xlsx'):
            excel_file = pd.ExcelFile(file_path, engine='openpyxl')
            sheet_dict = {}
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                sheet_dict[sheet_name] = df
            data_frames[file] = {
                'format': 'xlsx',
                'sheets': sheet_dict
            }
            print(f"{file_path} (Sheets: {excel_file.sheet_names})")
        else:  # CSV
            try:
                df = pd.read_csv(file_path, encoding='utf-8')
            except UnicodeDecodeError:
                df = pd.read_csv(file_path, encoding='gbk')
            data_frames[file] = {
                'format': 'csv',
                'df': df
            }
            print(file_path)

    return data_frames

def write_data_frames(data_frames: dict, output_folder: str, sanitized: bool = False):
    """
    將 data_frames 寫入指定的資料夾中。
    若 sanitized=True，則輸出的檔名會在副檔名前加上 "_sanitized"。
    否則維持原檔名輸出。
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    print("\nFollowing files are saved:")
    for file_name, content in data_frames.items():
        base_name, ext = os.path.splitext(file_name)
        # 如果要在檔名加上 _sanitized
        if sanitized:
            output_file_name = base_name + "_sanitized" + ext
        else:
            output_file_name = file_name

        output_path = os.path.join(output_folder, output_file_name)
        
        if content['format'] == 'xlsx':
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for sheet_name, df in content['sheets'].items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(output_path)
        else:  # CSV
            df = content['df']
            df.to_csv(output_path, index=False, encoding='utf-8-sig')
            print(output_path)

# ===== Utility functions for column operations =====

def insert_sanitized_columns(df: pd.DataFrame, columns_to_add: list):
    """
    在 df 中，對於 columns_to_add 裏列出的欄位，緊接在它後面插入一個對應的 {col}_sanitized 欄位
    (透過 sanitize() 函式來產生新值)。
    依欄位的 index 由大到小插入，避免影響後面欄位的插入位置。
    """
    # 先依照原欄位在 df.columns 中的索引位置，從「大到小」排序
    columns_with_locs = [(col, df.columns.get_loc(col)) for col in columns_to_add]
    columns_with_locs.sort(key=lambda x: x[1], reverse=True)
    
    for col, loc_val in columns_with_locs:
        new_col_name = f"{col}_sanitized"
        df.insert(loc_val + 1, new_col_name, df[col].apply(sanitize))

def insert_desanitized_columns(df: pd.DataFrame, columns_to_process: list):
    """
    在 df 中，對於 columns_to_process 裏列出的欄位，緊接在它後面插入一個對應的 External Part欄位
    (透過 desanitize() 函式來產生新值)。
    依欄位的 index 由大到小插入，避免影響後面欄位的插入位置。
    """
    columns_with_locs = [(col, df.columns.get_loc(col)) for col in columns_to_process]
    columns_with_locs.sort(key=lambda x: x[1], reverse=True)

    for col, loc_val in columns_with_locs:
        if col == 'External part ID':
            df.insert(loc_val + 1, 'External Part', df[col].apply(desanitize))

# ===== Main workflow functions =====

def sanitize_data():
    """
    1. 從 'Unsanitized' 資料夾讀取所有 .xlsx/.csv 檔案
    2. 如果 'External Part'/'Internal Part'/'Internal Part(Old)'/'Internal Part(New)' 欄位存在，就插入對應的 '*_sanitized' 欄位
    3. 將結果輸出到 'Sanitized_當天日期' 資料夾，檔名加上 "_sanitized"
    """
    print("\nSanitizing process started...")

    input_folder = 'Unsanitized'
    output_folder = 'Sanitized'

    data_frames = read_data_frames(input_folder)

    print("\nFollowing files (and sheets) are sanitized:")
    for file_name, content in data_frames.items():
        if content['format'] == 'xlsx':
            for sheet_name, df in content['sheets'].items():
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
                    print(f"Unsanitized\\{file_name} - Sheet: {sheet_name}")
        else:
            df = content['df']
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
                print(f"Unsanitized\\{file_name}")

    # 另存檔案
    write_data_frames(data_frames, output_folder, sanitized=True)


def desanitize_data():
    """
    1. 從 'Undesanitized' 資料夾讀取所有 .xlsx/.csv 檔案
    2. 如果 'external_id_sanitized'/'internal_id_sanitized' 欄位存在，就插入對應的 '*_desanitized' 欄位
    3. 將結果輸出到 'Desanitized' 資料夾 (檔名不更動)
    """
    print("\nDesanitizing process started...")

    input_folder = 'Undesanitized'
    output_folder = 'Desanitized'

    data_frames = read_data_frames(input_folder)

    print("\nFollowing files are desanitized:")
    for file_name, content in data_frames.items():
        if content['format'] == 'xlsx':
            for sheet_name, df in content['sheets'].items():
                columns_to_process = []
                if 'External part ID' in df.columns:
                    columns_to_process.append('External part ID')

                if columns_to_process:
                    insert_desanitized_columns(df, columns_to_process)
                    print(f"Undesanitized\\{file_name} - Sheet: {sheet_name}")
                else:
                    print(f"The file {file_name} (Sheet: {sheet_name}) does not contain sanitized columns.")
        else:
            df = content['df']
            columns_to_process = []
            if 'External part ID' in df.columns:
                columns_to_process.append('External part ID')

            if columns_to_process:
                insert_desanitized_columns(df, columns_to_process)
                print(f"Undesanitized\\{file_name}")
            else:
                print(f"The file {file_name} does not contain sanitized columns.")

    # 另存檔案
    write_data_frames(data_frames, output_folder, sanitized=False)


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
