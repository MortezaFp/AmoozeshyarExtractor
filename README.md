# Amoozeshyar Offered Courses Extractor

Author: Morteza Fattah Pour

This project automates extraction of offered course classes from Amoozeshyar, saves a raw Excel export, and generates filtered Excel/PDF reports.

## Output Files

All files are saved in the same folder as `main.py`:

- `لیست دروس ارائه شده آموزشیار.xlsx` (raw export, overwritten each run)
- `لیست دروس تخصصی.xlsx` (overwritten each run)
- `لیست دروس عمومی.xlsx` (overwritten each run)
- `لیست دروس تخصصی.pdf` (overwritten each run)
- `لیست دروس عمومی.pdf` (overwritten each run)

## Requirements

- Python 3.10+
- Google Chrome
- Amoozeshyar account access

## Install

Install Python dependencies from `requirements.txt`:

```bash
pip install -r requirements.txt
```

Install Playwright Chrome runtime (first-time setup):

```bash
python -m playwright install chrome
```

## Run

From this folder:

```bash
python main.py
```

Windows example with full interpreter path:

```powershell
C:/Users/<YourUser>/AppData/Local/Programs/Python/Python313/python.exe main.py
```

## Execution Flow

1. Chrome opens.
2. Log in to Amoozeshyar manually.
3. Return to terminal and press `Enter`.
4. The script runs search/pagination and generates all outputs.

## Font Note

Keep `B_Nazanin_Bold.ttf` in the same folder as `main.py` for correct PDF rendering.
