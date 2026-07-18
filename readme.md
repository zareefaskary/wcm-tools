# YSAM SAP Automation Suite

**Created by:** Zareef Askary

## 📖 Overview
The **YSAM SAP Automation Suite** is a set of Python-based GUI tools designed to streamline data extraction and validation for the SAP **YSAM** transaction. It eliminates manual data entry, bulk-exports reports to Excel, and automatically validates the output against a master plant list.

This repository contains both the **compiled executables (`.exe`)** for easy, no-install use, and the **Python source code (`.py`)** for transparency and customization.

---

## 📑 Table of Contents
1. [Quick Start (Using `.exe` files)](#-quick-start-using-exe-files)
2. [Running from Source (Using `.py` files)](#-running-from-source-using-py-files)
3. [How to Build the `.exe` Yourself](#-how-to-build-the-exe-yourself)
4. [Tool 1: YSAM Export Tool](#-tool-1-ysam-export-tool)
5. [Tool 2: YSAM Checker Tool](#-tool-2-ysam-checker-tool)
6. [Troubleshooting & SAP Requirements](#️-troubleshooting--sap-requirements)

---

## 📥 Quick Start (Using `.exe` files)
*Best for users who just want to run the tool without installing Python.*

1. Go to the **[Releases](https://github.com/zareefaskary/Export-Automation-tool/releases)** section of this repository.
2. Download the latest executables:
   - `Sapautogui2.exe` (For exporting SAP data)
   - `Checkgui.exe` (For validating the exported data)
3. Double-click the `.exe` to run. No installation or Python setup is required!

> **Note:** If your corporate IT blocks the `.exe` files, please use the "Running from Source" method below or ask your IT admin to whitelist them. The source code is provided in this repo for full transparency.

---

## 🐍 Running from Source (Using `.py` files)
*Best for developers or users who want to customize the scripts.*

### Prerequisites
1. **SAP GUI** installed with **Scripting Enabled** (both server and local client).
2. **Python 3.x** installed on your Windows machine.

### Installation
1. Clone or download this repository.
2. Open your command prompt/terminal and install the required dependencies:

```bash
pip install -r requirements.txt
```

*(The `requirements.txt` includes: `pandas`, `openpyxl`, and `pywin32`. Note: `tkinter` is built into standard Python).*

---

## 🛠️ How to Build the `.exe` Yourself
If you want to compile the Python scripts into `.exe` files for your own machine or to share with colleagues, you can use `pyinstaller`.

1. Install PyInstaller: 

```bash
pip install pyinstaller
```

2. Build the Export Tool: 

```bash
pyinstaller --onefile --windowed Sapautogui2.py
```

3. Build the Checker Tool: 

```bash
pyinstaller --onefile --windowed Checkgui.py
```

*(The compiled `.exe` files will be located in the newly created `dist/` folder).*

---

## 🚀 Tool 1: YSAM Export Tool
**Files:** `Sapautogui2.py` / `Sapautogui2.exe`

This tool reads an Excel file of Company/Plant codes and exports YSAM reports directly from SAP.

**Expected Excel Format:**
- Must contain at least two columns: **Company Code** (4-digit) and **Plant Code** (4-digit).
- Default sheet name: `Plant list_All` (configurable in the GUI).

**How to Use:**
1. **Excel File**: Browse and select your input Excel file.
2. **Output Folder**: Choose where the exported `.XLSX` files will be saved.
3. **SAP System**: Enter the exact name of your SAP connection (as it appears in SAP Logon).
4. **Parameters**: Select the **Posting Period** (01-12), enter the **Year** (e.g., 2026), and specify the SAP **Layout** variant.
5. **Execute**: Click the green **Run Script** button. 
   - Files will be generated as: `YSAM_{YEAR}0{POSTING_PERIOD}_{company}.XLSX`
6. **Stop**: Click the red **Stop Script** button to safely halt the process after the current company finishes.

---

## 🔍 Tool 2: YSAM Checker Tool
**Files:** `Checkgui.py` / `Checkgui.exe`

This tool acts as Quality Assurance (QA) to ensure the exported files are 100% accurate before you use the data.

**How to Use:**
1. **Master Excel File**: Select your master list Excel file.
2. **YSAM Folder**: Select the folder containing the exported YSAM files.
3. **Parameters**: Enter the **Year** and **Posting Period** you are checking.
4. **Execute**: Click the blue **Run Check** button.
5. **Validation**: The tool will instantly report:
   - Missing export files.
   - Missing or extra plants compared to the master list.
   - Mismatches in the `FYr previous period` or `Month prev. period` columns.

---

## ⚠️ Troubleshooting & SAP Requirements

- **SAP Scripting Must Be Enabled**: 
  - *Local Client*: Go to SAP GUI Options → Accessibility & Scripting → Scripting → Check "Enable scripting" and uncheck "Notify when a script attaches".
  - *Server-Side*: Your company's SAP BASIS team must have server-side scripting enabled. If you get a "Scripting is disabled on the server" error, contact your IT helpdesk.
- **"Navigation screens not found"**: The export script attempts to bypass initial YSAM warning/info screens. If your SAP environment has unique popups, you may need to adjust the `try/except` block in the `run_script` function.
- **File Overwriting**: If a file with the same name already exists in the Output Folder, SAP will prompt an overwrite confirmation. The script attempts to handle standard save popups, but it is best to use an empty output folder.
- **SAP Connection Conflicts**: Ensure you are not already logged into the exact same SAP system in another window, as this can cause SAP GUI scripting connection conflicts.

## 📜 License
Free to use and modify for internal business automation purposes.
