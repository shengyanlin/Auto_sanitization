# Auto Sanitization Script

This script provides two main features:
1. **Data Sanitization** — Reads `.xlsx` / `.csv` files from a specified folder, creates new “sanitized” columns (e.g., `external_id_sanitized`, `internal_id_sanitized`), and applies character shifting on the first two and last two non-space characters.
2. **Data Desanitization** — Reads `.xlsx` / `.csv` files from another folder, looks for “sanitized” columns, and reverts them into new columns (e.g., `external_desanitized`, `internal_desanitized`).

---

## Table of Contents

- [Auto Sanitization Script](#auto-sanitization-script)
  - [Table of Contents](#table-of-contents)
  - [Folder Structure](#folder-structure)
  - [Usage](#usage)
    - [Dependencies](#dependencies)
  - [Running the Script](#running-the-script)
  - [Creating an Executable](#creating-an-executable)
  - [Detailed Behavior](#detailed-behavior)
    - [Character Shifting](#character-shifting)
      - [Forward Shift (`shift_char_forward`):](#forward-shift-shift_char_forward)
      - [Backward Shift (`shift_char_backward`):](#backward-shift-shift_char_backward)
    - [Sanitizing Process](#sanitizing-process)
    - [Desanitizing Process](#desanitizing-process)
  - [License](#license)

---

## Folder Structure

Make sure the following folders exist so that the script can run correctly:

├── Auto_sanitization.py ├── Unsanitized ├── Sanitized ├── Undesanitized └── Desanitized


- **`Auto_sanitization.py`**: The main script (the code you provided).
- **`Unsanitized`**: Holds the `.xlsx` / `.csv` files that you want to sanitize.
- **`Sanitized`**: The folder where the newly “sanitized” files will be saved.
- **`Undesanitized`**: Holds the `.xlsx` / `.csv` files that you want to desanitize.
- **`Desanitized`**: The folder where the newly “desanitized” files will be saved.

If you need to use different folder names, simply edit the `input_folder` and `output_folder` variables in the script.

---

## Usage

### Dependencies

- Python 3.6 or higher  
- [pandas](https://pandas.pydata.org/)  
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)

If you have not installed them, run:

```bash
pip install pandas openpyxl

```
## Running the Script

1. **Prepare Files**:
   - Place your `.xlsx` or `.csv` files in the `Unsanitized` folder if you want to sanitize them.
   - Place your files in the `Undesanitized` folder if you want to desanitize them.

2. **Run the Script**:
   - Open a terminal or command prompt in the same directory as `Auto_sanitization.py`.
   - Execute the script:
     ```bash
     python Auto_sanitization.py
     ```

3. **Follow the Prompt**:
   - Type `y` to run the sanitizing process on files in `Unsanitized`, which will produce files in `Sanitized`.
   - Type `n` to run the desanitizing process on files in `Undesanitized`, which will produce files in `Desanitized`.

---

## Creating an Executable

If you wish to run the script on a machine without Python installed, use `PyInstaller` to create a standalone executable:

1. **Install PyInstaller**:
   ```bash
   pip install pyinstaller
   ```

2. **Create the Executable**:
   - In the same directory as `Auto_sanitization.py`, run:
     ```bash
     pyinstaller -F --icon=xxx.ico Auto_sanitization.py
     ```

3. **Executable Output**:
   - After PyInstaller finishes, the `dist` folder will contain `Auto_sanitization.exe` (on Windows).
   - Place the executable along with the four folders (`Unsanitized`, `Sanitized`, `Undesanitized`, `Desanitized`) in the same location. The program can then run on systems without Python installed.

---

## Detailed Behavior

### Character Shifting

#### Forward Shift (`shift_char_forward`):
- `a → b`, `z → a`
- `A → B`, `Z → A`
- `9 → 0`
- `_` remains unchanged; other symbols remain as is.

#### Backward Shift (`shift_char_backward`):
- `b → a`, `a → z`
- `B → A`, `A → Z`
- `0 → 9`
- `_` remains unchanged; other symbols remain as is.

---

### Sanitizing Process

1. Read all `.xlsx` and `.csv` files from the `Unsanitized` folder.
2. For `.xlsx` files, process every sheet in the workbook.
3. For `.csv` files, process the entire file.
4. Search for `external_id` and `internal_id` columns.
5. Insert new columns (`external_id_sanitized`, `internal_id_sanitized`) right after each corresponding ID column.
6. Apply the `sanitize` function to shift the first two and last two non-space characters forward by one.
7. Save the results into the `Sanitized` folder, appending `_sanitized` to the file name (e.g., `my_data.xlsx → my_data_sanitized.xlsx`).

---

### Desanitizing Process

1. Read all `.xlsx` and `.csv` files from the `Undesanitized` folder.
2. Search for `external_id_sanitized` and `internal_id_sanitized` columns.
3. Insert new columns (`external_desanitized`, `internal_desanitized`) right after the corresponding sanitized columns.
4. Apply the `desanitize` function to shift the first two and last two non-space characters backward by one.
5. Save the results into the `Desanitized` folder, retaining the original file name structure (e.g., `my_data.csv` remains `my_data.csv`).

---

## License

No specific license is provided. You are free to modify, distribute, and use this script. For questions or improvements, feel free to share suggestions.
