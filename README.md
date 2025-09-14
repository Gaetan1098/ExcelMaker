PiramieExcelMaker

Excel automation pipeline with Python (pandas, openpyxl, tkinter).

Overview

PiramieExcelMaker streamlines monthly data consolidation in Excel. It allows you to ingest a monthly purchases report into a master workbook without overwriting existing rows, automatically preventing duplicates, creating backups, and expanding filters/tables for sorting.

A simple drag-and-drop GUI makes it easy to select files and run either:

Core Only → Builds summary and pivoted reports from the master.

Ingest + Core → Appends a monthly report into the master, then generates updated summaries.

The tool demonstrates:

Excel data ingestion and normalization

Header mapping and row de-duplication

Automated pivot/summarization logic

GUI design with tkinter + tkinterDnD2

File backup and safe write operations

⚠️ Note: This repo uses synthetic sample Excel files. No real company data is included.

Tech Stack

Python 3.11

pandas

openpyxl

tkinter / tkinterDnD2 (optional for drag & drop)

Usage

Clone the repo.

Install dependencies:

pip install pandas openpyxl tkinterdnd2


Run the GUI:

python PiramieExcelMaker_gui.py


Select your Master workbook (.xlsm or .xlsx) and optionally a Monthly report (.xlsx), then choose an action.

Sample Data

This repo provides:

Master_Workbook_Sample.xlsx (with APR Bundle + Molo Molo sheets)

Purchases_Report_Sample_May_v2.xlsx (monthly report with headers on row 4)

These are synthetic examples that mirror real-world structure so you can demo the pipeline safely.

Demo


(Replace with your own screenshot using the sample data)
