# TERA Report Generator

Automated PDF report generator for **TERA** (Transcriptome based Endometrial
Receptivity Assessment) — an Anderson Diagnostics & Labs in-house pipeline.

## Features
- Generates 3-page patient reports classifying endometrial receptivity as
  **Receptive**, **Pre-receptive**, or **Post-receptive**
- Bulk Excel import with per-row inline editor
- Live PDF preview (pypdfium2)
- With Logo / Without Logo export modes
- Auto-diff comparison: compares two PDFs and summarises every text difference
- Draft save/load (JSON)

## Quick Start

```bash
pip install -r requirements.txt
python tera_report_generator.py
```

## Requirements
- Python 3.10+
- See `requirements.txt`

## File Structure
```
TERA-Report-Automation/
├── tera_template.py          # PDF engine (ReportLab)
├── tera_report_generator.py  # PyQt6 desktop GUI
├── tera_assets.py            # Base64-encoded image assets
├── generate_assets_py.py     # Utility: regenerate tera_assets.py from PNGs
├── test_run.py               # Single-row smoke test
├── fonts/                    # TTF font files
│   ├── Calibri.ttf
│   ├── GillSansMT-Bold.ttf
│   └── ...
├── requirements.txt
├── launch_tera.sh            # Linux launcher
└── launch_tera.bat           # Windows launcher
```

## Output Filename Format
```
{Patient Name}_{Nth biopsy}_TERA_report_with logo.pdf
{Patient Name}_{Nth biopsy}_TERA_report_without logo.pdf
```
Example: `Hemalatha Venkatesh_1st biopsy_TERA_report_without logo.pdf`

## Header
Uses the shared Anderson Diagnostics header (same as PGTA reports).
Place the PGTA project at `../PGTA-Report/` for automatic header sharing.

## Notes
- PDF viewer: Evince, Okular, or Firefox. **Not** VS Code (shows binary).
- Input: `.xls` TERA automation report (xlrd required for `.xls`).
