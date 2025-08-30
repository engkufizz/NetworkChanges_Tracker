# 📟 Network Changes Tracker

A simple desktop tool built with **PySide6** and **openpyxl** to help network engineers track **network changes** (Change Requests, Work Permits).
It provides a clean UI to add, view, and manage records directly in an Excel file (`network_changes.xlsx`).

---

## ✨ Features

* 📅 **Date Picker** – Select or auto-fill today’s approval date.
* 🔢 **Request Number Field** – Enter CR/NC reference numbers (optional or can be enforced).
* 📝 **Multi-line Description** – Write detailed work notes; multiple lines are auto-joined into a single line.
* 📊 **Excel Integration** – All data is stored in `network_changes.xlsx` with separate sheets for:

  * **CR** (Change Requests)
  * **WP** (Work Permits)
* 🔄 **Quick Actions**

  * Add new record
  * Clear input fields
  * Refresh table view
  * Open Excel directly from the app
* ⚡ **Keyboard Shortcuts**

  * `Ctrl+Enter` → Add record
  * `Ctrl+L` → Clear input
  * `Ctrl+T` → Set today’s date
  * `Ctrl+O` → Open Excel file
  * `F5` → Refresh records
* 📋 **Context Menu** – Right-click a row to copy it to clipboard.
* 🎨 **Modern UI** – Clean interface with alternating row colors.

---

## 📂 File Structure

* `network_changes.xlsx` → Auto-created if not found, contains 2 sheets: `CR`, `WP`.
* `app.py` (the app) → Runs the GUI.

---

## 🚀 Getting Started

### 1. Clone or Download

```bash
git clone https://github.com/your-username/network-changes-tracker.git
cd network-changes-tracker
```

### 2. Install Requirements

```bash
pip install PySide6 openpyxl
```

### 3. Run the App

```bash
python app.py
```

---

## 📑 Excel Format

Each sheet (`CR`, `WP`) has the following headers:

| Approval Date | Request Number | Description of Work                  |
| ------------- | -------------- | ------------------------------------ |
| 2025-08-30    | CR/ENP/1234    | Router upgrade, configuration backup |

---

## 🖥️ Platform Support

* ✅ Windows (with proper taskbar icon grouping)
* ✅ macOS
* ✅ Linux

---

## ⚠️ Common Issues

* **Excel file won’t save** → Close it if already open in Excel.
* **PermissionError** → Move the app and Excel file to a writable folder (e.g., Desktop/Documents).
* **Old Excel files with 2 columns** → The app will migrate them to the new 3-column format automatically.

---

## 📜 License

MIT License – free to use, modify, and distribute.
