# ETL Data Pipeline | Excel Automation

A Python ETL pipeline that processes multiple Excel files automatically using a reusable function library.

## Features
- Batch scan and validate Excel files (missing columns, empty files detected)
- Read and align columns across multiple files
- Merge all valid files into one dataset
- Auto data cleaning: remove duplicates and missing values
- Export polished Excel report with formatted headers and zebra stripes
- Generate interactive salary chart (HTML, hover to see values)

## Requirements
pip install -r requirements.txt

## Usage
1. Set your paths in main.py:
input: FOLDER_PATH = r"your/input/folder"
output: OUTPUT_PATH = r"your/output/file.xlsx"

2. Run:
python main.py

## Deliverables
- main.py: main pipeline script
- data_tools.py: reusable function library (5 core functions)
- requirements.txt: dependencies
- input/: sample input Excel files
- output/: cleaned Excel report + interactive HTML chart
