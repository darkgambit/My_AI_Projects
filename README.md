# SheetScout — Excel name search

A local web tool for finding names across many Excel files. Searches every sheet and column, returns complete matching rows with workbook/sheet/Excel row references, and exports results as CSV. Supports `.xlsx`, `.xlsm` and legacy `.xls`. Existing scripts and spreadsheets are unchanged.

## Run on your computer

Install Python 3.10 or newer, then from this project folder:

```sh
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS / Linux instead:
# source .venv/bin/activate
pip install -r requirements.txt
python -m excel_search.app --folder "C:\Users\YourName\Documents\Excel Files"
```

Open **http://127.0.0.1:5000**. On macOS/Linux, supply your own folder path. Omit `--folder` to search the current directory. Subfolders are included; hidden directories and Excel lock files are skipped. Restart with a different `--folder` to change the search directory. Upload mode also lets you select multiple files without moving them into that folder.

## Search and result format

1. Choose folder files or upload Excel workbooks (100 MB total maximum).
2. Enter one name per line. A row matching any entered name is returned once.
3. Select **Contains** (partial name) or **Exact cell match**. Both ignore case, normalize Unicode, and collapse whitespace. Search is literal, not fuzzy.
4. Set the header row. Default 1 treats the first row of every sheet as column labels. Rows up to and including that row are skipped; use 0 to search every row with column-letter labels.
5. Click **Search workbooks**. Matching cells are highlighted. **Export CSV** downloads all displayed records, with workbook, sheet, row number, searched names, and the original fields as separate columns.

Results show at most 1,000 matching rows and warn if more exist. Corrupt or password-protected files are reported individually without stopping the other files. Workbook headers can differ; their fields are combined for CSV output. Excel formulas use cached values from the last time Excel saved the file; formulas are not recalculated. Macros are never executed. Excel files are never modified. Uploaded files are temporary and deleted after a search. No external services or AI APIs are used.

Searches read files each time, so large folders can take a while. `.xls` dates are returned in their underlying Excel representation. Headers are configured globally; search differently structured workbooks separately if necessary.

## Hosted preview / security

```sh
python -m excel_search.app --host 0.0.0.0 --port 5000 --folder .
```

This is a single-user tool, not an authenticated public service. Anyone who can reach the server can search the configured folder and upload files. Only expose it on a trusted network or behind authentication. Uploading files to a hosted preview sends their contents to that server; for sensitive workbooks, run locally instead.

## Tests

```sh
python -m unittest discover -s tests -v
```
