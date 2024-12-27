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
    """
    s = str(s)
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

# ===== Main workflow functions =====

def sanitize_data():
    """
    1. Read .xlsx and .csv files from 'Unsanitized' folder.
    2. If the file is .xlsx, read every sheet. 
       If the file is .csv, read it directly.
    3. For each DataFrame (every sheet of .xlsx / single .csv):
       - If 'external_id' column exists, create 'external_id_sanitized' column
         immediately after 'external_id'.
       - If 'internal_id' column exists, create 'internal_id_sanitized' column
         immediately after 'internal_id'.
       (The original 'external_id' / 'internal_id' is not modified.)
    4. Save the results into the 'Sanitized' folder with file name appended "_sanitized".
    """
    print("\nSanitizing process started...")
    
    input_folder = 'Unsanitized'
    output_folder = 'Sanitized'

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    file_list = [f for f in os.listdir(input_folder) if f.endswith(('.xlsx', '.csv'))]

    # 用來存所有讀取到的 DataFrame（或多個 Sheet）
    data_frames = {}
    
    print("\nFollowing files are read:")
    for file in file_list:
        file_path = os.path.join(input_folder, file)

        if file.endswith('.xlsx'):
            # 對於 .xlsx，先取得所有 sheets
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

        elif file.endswith('.csv'):
            # 對於 .csv
            try:
                df = pd.read_csv(file_path, encoding='utf-8')
            except UnicodeDecodeError:
                df = pd.read_csv(file_path, encoding='gbk')

            data_frames[file] = {
                'format': 'csv',
                'df': df
            }
            print(file_path)

    print("\nFollowing files (and sheets) are sanitized:")
    for file_name, content in data_frames.items():
        if content['format'] == 'xlsx':
            # 逐一處理該 xlsx 的每個 sheet
            for sheet_name, df in content['sheets'].items():
                # 收集可能要新增 sanitized 欄位的欄位名稱
                columns_to_add = []
                if 'external_id' in df.columns:
                    columns_to_add.append('external_id')
                if 'internal_id' in df.columns:
                    columns_to_add.append('internal_id')

                # 先依照原欄位在 df 中的索引位置，從「大到小」排序
                # 避免先插入前面欄位後，後面欄位的索引被改動影響
                columns_with_locs = [
                    (col, df.columns.get_loc(col)) for col in columns_to_add
                ]
                columns_with_locs.sort(key=lambda x: x[1], reverse=True)

                # 依照排序後的順序插入 _sanitized 欄位
                for col, loc_val in columns_with_locs:
                    df.insert(
                        loc_val + 1,
                        f"{col}_sanitized",
                        df[col].apply(sanitize)
                    )
                
                print(f"Unsanitized\\{file_name} - Sheet: {sheet_name}")

        else:  # CSV
            df = content['df']
            columns_to_add = []
            if 'external_id' in df.columns:
                columns_to_add.append('external_id')
            if 'internal_id' in df.columns:
                columns_to_add.append('internal_id')

            columns_with_locs = [
                (col, df.columns.get_loc(col)) for col in columns_to_add
            ]
            columns_with_locs.sort(key=lambda x: x[1], reverse=True)

            for col, loc_val in columns_with_locs:
                df.insert(
                    loc_val + 1,
                    f"{col}_sanitized",
                    df[col].apply(sanitize)
                )

            print(f"Unsanitized\\{file_name}")

    # TODO: Whether to delete some columns or not

    print("\nFollowing files are saved:")
    for file_name, content in data_frames.items():
        # 拆分「檔名」與「副檔名」
        base_name, ext = os.path.splitext(file_name)

        # 在檔名後面加上 "_sanitized"
        sanitized_file_name = base_name + "_sanitized" + ext

        # 重新組出完整路徑
        output_path = os.path.join(output_folder, sanitized_file_name)

        if content['format'] == 'xlsx':
            # 將多個 sheets 存回同一個 .xlsx
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for sheet_name, df in content['sheets'].items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(output_path)

        elif content['format'] == 'csv':
            df = content['df']
            df.to_csv(output_path, index=False, encoding='utf-8-sig')
            print(output_path)


def desanitize_data():
    """
    1. Read .xlsx and .csv files from 'Undesanitized' folder.
    2. If 'external_id_sanitized' or 'internal_id_sanitized' columns exist,
       apply desanitize() on them, and create new columns:
         - external_desanitized (right after external_id_sanitized)
         - internal_desanitized (right after internal_id_sanitized)
    3. Save the files into the 'Desanitized' folder.
    """
    print("\nDesanitizing process started...")

    input_folder = 'Undesanitized'
    output_folder = 'Desanitized'

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 只讀取 .xlsx 或 .csv
    file_list = [f for f in os.listdir(input_folder) if f.endswith(('.xlsx', '.csv'))]

    # 用來存所有讀取到的 DataFrame（或多工作表）
    # 結構範例：
    # data_frames = {
    #   "file1.xlsx": {
    #       'format': 'xlsx',
    #       'sheets': {
    #           'Sheet1': df1,
    #           'Sheet2': df2,
    #       }
    #   },
    #   "file2.csv": {
    #       'format': 'csv',
    #       'df': df_csv
    #   }
    # }
    data_frames = {}

    print("\nFollowing files are read:")
    for file in file_list:
        file_path = os.path.join(input_folder, file)
        file_format = 'xlsx' if file.endswith('.xlsx') else 'csv'

        if file_format == 'xlsx':
            # 讀取多個 sheet
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

    # 對每個檔案的每個 DataFrame 進行 desanitize
    print("\nFollowing files are desanitized:")
    for file_name, content in data_frames.items():
        if content['format'] == 'xlsx':
            # 逐一處理每個 sheet
            for sheet_name, df in content['sheets'].items():
                # 收集要處理的欄位 (sanitized 欄位)
                columns_to_process = []
                if 'external_id_sanitized' in df.columns:
                    columns_to_process.append('external_id_sanitized')
                if 'internal_id_sanitized' in df.columns:
                    columns_to_process.append('internal_id_sanitized')

                if columns_to_process:
                    # 依欄位位置由大到小排序，避免插入新欄位時互相影響
                    col_with_loc = [(col, df.columns.get_loc(col)) for col in columns_to_process]
                    col_with_loc.sort(key=lambda x: x[1], reverse=True)

                    # 逐一插入新欄位
                    for col, loc_val in col_with_loc:
                        if col == 'external_id_sanitized':
                            # 在 external_id_sanitized 後面插入 external_desanitized
                            df.insert(
                                loc_val + 1,
                                'external_desanitized',
                                df[col].apply(desanitize)
                            )
                        elif col == 'internal_id_sanitized':
                            # 在 internal_id_sanitized 後面插入 internal_desanitized
                            df.insert(
                                loc_val + 1,
                                'internal_desanitized',
                                df[col].apply(desanitize)
                            )
                    print(f"Undesanitized\\{file_name} - Sheet: {sheet_name}")
                else:
                    print(f"The file {file_name} (Sheet: {sheet_name}) does not contain 'external_id_sanitized' or 'internal_id_sanitized' columns")

        else:  # CSV
            df = content['df']
            columns_to_process = []
            if 'external_id_sanitized' in df.columns:
                columns_to_process.append('external_id_sanitized')
            if 'internal_id_sanitized' in df.columns:
                columns_to_process.append('internal_id_sanitized')

            if columns_to_process:
                col_with_loc = [(col, df.columns.get_loc(col)) for col in columns_to_process]
                col_with_loc.sort(key=lambda x: x[1], reverse=True)

                for col, loc_val in col_with_loc:
                    if col == 'external_id_sanitized':
                        df.insert(
                            loc_val + 1,
                            'external_desanitized',
                            df[col].apply(desanitize)
                        )
                    elif col == 'internal_id_sanitized':
                        df.insert(
                            loc_val + 1,
                            'internal_desanitized',
                            df[col].apply(desanitize)
                        )
                print(f"Undesanitized\\{file_name}")
            else:
                print(f"The file {file_name} does not contain 'external_id_sanitized' or 'internal_id_sanitized' columns")

    # 將結果存檔
    print("\nFollowing files are saved:")
    for file_name, content in data_frames.items():
        file_format = content['format']
        output_path = os.path.join(output_folder, file_name)

        if file_format == 'xlsx':
            # 寫回多個 sheet
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for sheet_name, df in content['sheets'].items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(output_path)
        else:  # CSV
            df = content['df']
            df.to_csv(output_path, index=False, encoding='utf-8-sig')
            print(output_path)

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

## cmd to package the script => pyinstaller -F Auto_sanitization.py