# TERA Report Generator

Automated PDF report generator for **TERA** (Transcriptome based Endometrial
Receptivity Assessment) — Anderson Diagnostics & Labs in-house pipeline.

---

## Installation

### Step 1 — Install Python (one-time, skip if already installed)

| OS | Download |
|----|----------|
| Windows | [python.org/downloads](https://www.python.org/downloads/) — **check "Add Python to PATH"** during install |
| Ubuntu/Debian | `sudo apt install python3 python3-pip python3-venv` |
| Fedora | `sudo dnf install python3` |

Requires **Python 3.10 or higher**.

### Step 2 — Download / clone this repository

```
git clone https://github.com/andersoninhousepipeline-dot/TERA-Report-Automation.git
cd TERA-Report-Automation
```

Or download the ZIP from GitHub and extract it.

### Step 3 — Run the installer (one-time)

**Windows** — double-click `install.bat`

**Linux** — open a terminal in the folder and run:
```bash
bash install.sh
```

This creates a local `venv/` folder and installs all dependencies automatically.

---

## Running the App

**Windows** — double-click `launch_tera.bat`

**Linux**
```bash
bash launch_tera.sh
```

---

## Features

- Generates 3-page patient PDF reports classifying endometrial receptivity as
  **Receptive**, **Pre-receptive**, or **Post-receptive**
- **Manual Entry** tab — fill fields, live PDF preview, draft save/load
- **Bulk Upload** tab — import Excel (`.xls`), table view, double-click to edit, batch generate
- **With Logo / Without Logo** — with logo includes Anderson header & footer images; without logo is plain
- **PDF Comparison** tab — side-by-side viewer + auto-diff text report
- Draft save/load (JSON)

---

## Output Filename Format

```
{Patient Name}_{Nth biopsy}_TERA_report_with logo.pdf
{Patient Name}_{Nth biopsy}_TERA_report_without logo.pdf
```

Example: `Hemalatha Venkatesh_1st biopsy_TERA_report_with logo.pdf`

---

## File Structure

```
TERA-Report-Automation/
├── tera_template.py          # PDF engine (ReportLab)
├── tera_report_generator.py  # PyQt6 desktop GUI
├── tera_assets.py            # Base64-encoded image assets (header, footer, charts)
├── fonts/                    # TTF font files
├── requirements.txt          # Python dependencies
├── install.bat               # Windows: first-time setup
├── install.sh                # Linux: first-time setup
├── launch_tera.bat           # Windows: run the app
└── launch_tera.sh            # Linux: run the app
```

---

## Notes

- Input file: `.xls` TERA automation report (Excel 97-2003 format)
- PDF viewer: Evince, Okular, or any standard PDF reader
