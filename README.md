# Spreadsheet-Scanner
A simple Python script that compares cells of similar spreadsheets. A possible application might be to check for original responses in several homework submissions that are based on a template.


## Usage:
1. Get Python, at least 3.10
2. Get dependencies with `python -m pip install levenshtein openpyxl`
3. Download all the submissions. The script can filter out any duplicates with the same values for the name cell.
4. Place this script in that folder. Configure the options on lines 20-28 if need be
5. Run the script. It may take a few minutes.
