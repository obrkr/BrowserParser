# 🌐 Browser History Parser

A Python tool for extracting and exporting browsing history from **Chrome, Edge, Brave, Opera, Vivaldi, and other Chromium-based browsers**. Parses the browser's SQLite `History` database and exports everything to a formatted `.xlsx` spreadsheet.

---

## ✨ Features

- Parses **browsing history** (URL, title, timestamp, visit count)
- Extracts **search queries** from `keyword_search_terms`
- Extracts **download history** (file path, URL, timestamp, state)
- Lists installed **browser extensions** (name, version, Chrome Web Store link) (If you have the folder.)
- Handles **locked databases** by working on a temporary copy
- Adapts to different **Chrome schema versions** automatically
- Exports to a clean, formatted **multi-sheet Excel workbook**
- GUI **folder picker** with console hints for all supported browsers and platforms

---

## 📊 Output

The script generates a `browser_history.xlsx` file (saved alongside the script) with four sheets:

| Sheet | Columns |
|---|---|
| **History** | Timestamp, Title, URL, Visit Count |
| **Searches** | Timestamp, Search Term, URL ID |
| **Downloads** | Timestamp, File Path, URL, State |
| **Extensions** | Name, Version, Install Date, Store URL |

If `browser_history.xlsx` already exists, the output is saved as `browser_history_1.xlsx`, `browser_history_2.xlsx`, etc.

---

## 🖥️ Supported Browsers

Any Chromium-based browser is supported. Tested profile paths:

**macOS**
```
Chrome:   ~/Library/Application Support/Google/Chrome/Default/
Edge:     ~/Library/Application Support/Microsoft Edge/Default/
Brave:    ~/Library/Application Support/BraveSoftware/Brave-Browser/Default/
Opera:    ~/Library/Application Support/com.operasoftware.Opera/Default/
Vivaldi:  ~/Library/Application Support/Vivaldi/Default/
Chromium: ~/Library/Application Support/Chromium/Default/
```

**Windows**
```
Chrome:   %LOCALAPPDATA%\Google\Chrome\User Data\Default\
Edge:     %LOCALAPPDATA%\Microsoft\Edge\User Data\Default\
Brave:    %LOCALAPPDATA%\BraveSoftware\Brave-Browser\User Data\Default\
Opera:    %APPDATA%\Opera Software\Opera Stable\
Vivaldi:  %LOCALAPPDATA%\Vivaldi\User Data\Default\
Chromium: %LOCALAPPDATA%\Chromium\User Data\Default\
```

**Linux**
```
Chrome:   ~/.config/google-chrome/Default/
Edge:     ~/.config/microsoft-edge/Default/
Brave:    ~/.config/BraveSoftware/Brave-Browser/Default/
Chromium: ~/.config/chromium/Default/
```

---

## 📦 Requirements

- Python 3.7+
- [openpyxl](https://openpyxl.readthedocs.io/)
- tkinter (usually bundled with Python)

Install dependencies:

```bash
pip install -r requirements.txt
```

> **Note:** `tkinter` is included with most Python installations. On Linux you may need to install it separately with `sudo apt install python3-tk`.

---

## 🚀 Usage

1. **Clone the repository:**

```bash
git clone https://github.com/obrkr/BrowserParser.git
cd BrowserParser
pip install -r requirements.txt
```

2. **Close your browser** (or the script will work from a copy, but closing it ensures the most up-to-date data).

3. Run the script:

```bash
python browser_history.py
```

4. A **folder picker dialog** will open. Navigate to your history is located. and click **Select**.

5. The script will parse your history and save `browser_history.xlsx` in the same directory as the script.

---

## 📁 Project Structure

```
BrowserParser/
├── browser_history.py   # Main script
├── requirements.txt     # Python dependencies
├── .gitignore
└── README.md
```

---

## ⚠️ Notes

- The script **never modifies** your original browser database — it always works on a temporary copy.
- The temporary copy (`history_copy.db`) is deleted automatically after parsing.
- Timestamps are stored by Chromium in microseconds since **1 Jan 1601** and are converted to UTC.
- Extension names that use localisation placeholders (e.g. `__MSG_appName__`) will fall back to displaying the extension ID.

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).

---

## 👤 Author

**obrkr**  
GitHub: [@obrkr](https://github.com/obrkr)