# SheetScout for Excel VBA (Windows)

Requires desktop Microsoft Excel on Windows. No Python, browser, or additional VBA references are required. Excel Online cannot run VBA. This version uses Windows filesystem objects and is not compatible with Mac Excel.

## One-time setup

1. Open a **new blank workbook** in Excel. Save it as **SheetScout.xlsm**, choosing **Excel Macro-Enabled Workbook (*.xlsm)**. Do not use one of your original data files as the search tool.
2. Press **Alt+F11** to open the VBA editor.
3. Choose **File > Import File**, then select the included **SheetScout.bas**. (Import the file, rather than pasting its `Attribute` line into the editor.)
4. Return to Excel with **Alt+F11**.
5. Press **Alt+F8**, select **SetupSheetScout**, and click **Run**.
6. Save the workbook. Only enable macros for this workbook if you trust the code. Do not globally disable Excel macro security. If Windows blocks a downloaded file, review it before using its Properties > Unblock option, subject to your organization's policy.

## Searching

- In the **Name Search** sheet, enter one or more names in **B3**, separated by commas or Alt+Enter line breaks.
- Click **Choose folder**, then select your Excel data folder. All subfolders are included except filesystem junctions/symbolic links.
- Set **B5** to `Contains` for partial names, or `Exact` for a whole-cell match. Both ignore capitalization and repeated whitespace.
- **B6** is the header row, applied to every sheet. Rows through that row are skipped. Set 0 if your data has no headers.
- Click **Search names** and confirm replacing the previous results.

The **Results** sheet contains one row per matching Excel row:

| Workbook | Sheet | Excel row | Names searched | Matching columns | Record details | Full file path |
|---|---|---|---|---|---|---|
| Staff.xlsx | Employees | 12 | Ahmed Ali | B | ID: 17 \| Name: Ahmed Ali \| Branch: Tripoli | C:\Data\Staff.xlsx |

Record details combine nonempty original fields with their headers. You can filter this table and save a separate copy or CSV using normal Excel features. Dates, formatted numbers, and formulas are searched using their underlying values; date cells can appear as Excel serial numbers. Detail text is limited to 32,000 characters per result. Results stop after 10,000 matching rows; narrow your search if the limit is reached. Very large used ranges can take time and memory.

## Safety and limitations

Supports `.xlsx`, `.xlsm`, `.xls`, and `.xlsb`. Source workbooks open read-only with automatic link updates disabled and VBA macros disabled, and close without saving. Existing open workbooks and the search tool itself are skipped. Close data workbooks before searching. File errors, password-protected files, and inaccessible folders are listed in **Search Log**. Some matches from a partially read workbook may remain if that file later produces an error.

Excel may still show security/legacy Excel 4 macro or external-connection warnings for untrusted workbooks; do not approve unexpected prompts. Only scan files you trust. The tool does not remove workbook passwords. Excel settings changed by the macro are restored on success or error; do not force-close Excel during the search.

`SetupSheetScout` resets the Name Search form. `SearchNames` replaces Results and Search Log after confirmation. Use a dedicated workbook without unrelated sheets with those names.

This module was prepared in a Linux environment without Microsoft Excel; it has not been runtime-tested in Excel. If Excel reports a compile or runtime error, send the message and the highlighted line.
