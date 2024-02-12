# Python_Excel_Data_Matching_Copy_V2
 
# Excel Data Processing Scripts

This repository contains two Python scripts for processing data in Excel files using the `openpyxl` library. These scripts are designed to work with specific Excel file structures and perform particular data manipulation tasks.

## Scripts

### 1. Data Matching and Updating Script (`match_and_update.py`)

This script matches and updates data between two sheets in an Excel file.

#### Features:
- Reads data from 'Sayfa1' and 'Sayfa2' within the specified Excel file.
- Matches values from column C in 'Sayfa1' against values in column B of 'Sayfa2'.
- If a match is found (direct match or after trimming spaces), it writes the corresponding value from column A of 'Sayfa2' into column B of 'Sayfa1'.
- Skips updating if no match is found.

### 2. Column Update Script (`update_column.py`)

This script updates the contents of a specific column in an Excel sheet.

#### Features:
- Processes data in column C of 'Sayfa2' in the specified Excel file.
- Trims each value in column C up to the first '-' character.
- Updates the cell with the trimmed value.

## Prerequisites

- Python 3
- `openpyxl` library

## Installation

1. Ensure Python 3 is installed on your system.
2. Install `openpyxl` using pip:

   ```bash
   pip install openpyxl
Usage
To run these scripts, navigate to the script's directory and run:
python match_and_update.py
or
python update_column.py
Make sure to modify the file_path variable in each script to point to your Excel file.