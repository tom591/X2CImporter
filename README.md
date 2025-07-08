# 📊 X2CImporter | XLSX → CSV Converter (with optional GUI)
by Tomas Salaj

---

## 🧩 Overview

**X2CImporter** is a Python utility for converting Excel spreadsheets (`.xlsx`) into a unified `.csv` file.  
It supports multiple sheets, selected columns, and adds row numbers to all rows for later database import (e.g. into MySQL).

Includes a simple GUI (Tkinter) for selecting and converting files without writing code.

This tool is intentionally very simple – but for basic purposes, such as converting `.xlsx` tables into `.csv` for database import, it works well.  
I personally use it at work for fast conversions of a specific internal spreadsheet.

---

## 🖼️ Screenshot

![GUI Preview](screen.png)

---

## 🚀 Features

- Reads multiple Excel sheets from a single file
- Selects specific columns (e.g. [A, B, C, D, E, F, G, I])
- Automatically adds sequential row numbers
- Skips the first two rows (`skiprows=2`) – assumes headers or metadata
- Combines data from all sheets into one CSV
- Exports to UTF-8 `.csv` with semicolon `;` separator
- Includes a minimal graphical interface
- No installation required – just run with Python

---

## 📁 Folder Structure

```
X2CImporter/
├── X2CImporter_GUI.py       # GUI version of the tool
├── python_base.py           # Terminal-based CLI script
├── _internal/               # Contains icons and other static assets
│   └── icons/
│       ├── xlsx_icon.png
│       └── csv_icon.png
├── screen.png               # Screenshot of the application
├── requirements.txt         # Required Python packages
└── README.md
```

---

## ⚙️ Customization

The input Excel file is expected to have:
- Up to 4 sheets
- Data starting from **row 3** (due to `skiprows=2`)
- Specific columns in use: 0, 1, 2, 3, 4, 5, 6, 8  
  *(Note: Python uses 0-based indexing)*

You may freely modify:
- The number of sheets read
- The columns selected
- How rows are skipped
- Any data transformation logic

---

## 💾 Installation & Requirements

Make sure you have Python 3 installed. Then run:

```bash
pip install -r requirements.txt
```

This installs:

- `pandas`
- `openpyxl`

> `tkinter` is included by default in most Python installations.

---

## ▶️ Running the Program

### GUI version

```bash
python X2CImporter_GUI.py
```

### Terminal version

```bash
python python_base.py
```

---

## 🧱 Optional: Create Executable (.exe)

You can build a standalone Windows executable using [`auto-py-to-exe`](https://github.com/brentvollebregt/auto-py-to-exe).

### Recommended settings:
- **Output mode**: Directory (not single file)
- **Include folder**: `_internal/`
- **Entry script**: `X2CImporter_GUI.py`
- **Console window**: Disabled (Window-based)

> Be aware: including `pandas` and `openpyxl` results in a large `.exe` (typically 80–120 MB)

---

## 📝 License & Credits

This project is released under the [MIT License](LICENSE).  
Use freely, modify as needed. No warranty provided.  
© Tomas Salaj, 2024

---
