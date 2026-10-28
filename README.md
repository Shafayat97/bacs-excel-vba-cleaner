# BACS Excel VBA Cleaner
Stop fixing BACS payment files by hand. Type or paste your payment data into Excel, and this VBA tool automatically formats every cell, then exports a clean CSV ready to upload to your banking or payroll software.

Demo

The problem

Preparing a BACS payment run in Excel usually means fixing the same problems every time:

Excel drops leading zeros, so sort code 05-02-12 becomes 50212 and the bank rejects the file.
Account numbers arrive with 7 digits instead of 8.
Names and references contain special characters or are longer than the 18-character BACS limit.
Amounts are typed as £1,250.50 instead of a plain number.
Long invoice numbers turn into scientific notation like 1.23E+11.

Fixing all of this manually is slow, boring and easy to get wrong. This tool does it automatically.

What it does

As soon as you type or paste data into the sheet, each column is cleaned:

Column	Field	What the macro does
A	Sort code	Keeps digits only and formats to exactly 6 digits (restores lost leading zeros)
B	Name	Removes special characters and double spaces, cuts to 18 characters
C	Account number	Keeps digits only and formats to exactly 8 digits
D	Amount	Converts to a number with 2 decimal places
E	Reference / invoice no.	Removes special characters and double spaces, cuts to 18 characters

Sort codes and account numbers are stored as text, so Excel can't strip the zeros again.

Example
	Sort code	Name	Account no.	Amount	Reference
Before	20-00-00	Mr. John O'Brien-Smith (Ltd)	1234567	£1,250.50	INV-2026/001
After	200000	Mr John OBrienSmit	01234567	1250.50	INV2026001
How to use
Download BACS_Cleaner.xlsm and open it in Excel.
Click Enable Content when Excel asks about macros.
Type or paste your payment data into the sheet, starting at cell A1 (no header row).
Run one of the export macros (Alt + F8):
SaveAsCSV creates output.csv
SaveAsNotepad creates output.txt from the CSV
ExportAll creates both
The files are saved in the same folder as the workbook.
Try it with sample data

sample-data/raw_data.xlsx contains 12 rows of deliberately messy, completely fake payment data. Copy it into the workbook and watch it get cleaned. sanitized_data.csv and sanitized_data.txt show the exported result.

Files
bacs-excel-vba-cleaner/
├── BACS_Cleaner.xlsm            the workbook with the macros built in
├── src/
│   ├── Sheet1_AutoClean.vb      cleans cells as you type (sheet module)
│   └── Export_Macros.bas        SaveAsCSV, SaveAsNotepad, ExportAll
├── sample-data/
│   ├── raw_data.xlsx            messy fake data (before)
│   ├── sanitized_data.csv       exported CSV (after)
│   └── sanitized_data.txt       exported TXT (after)
└── screenshots/
    ├── before.png
    ├── after.png
    └── demo.gif

The VBA code is also provided as plain text files, so you can read it before opening a macro-enabled workbook.

⚠️ Important
Always check the exported file before submitting a payment run. This tool formats data; it cannot tell whether a sort code or account number is actually correct.
Don't leave blank cells in a row before exporting to TXT. The TXT export removes empty commas, which would shift the remaining columns.
If a sort code or account number has too many digits, only the last 6 or 8 digits are kept, so double-check any unusual entries.
Requirements vary by bank and software (for example, some want amounts in pence: 125050 instead of 1250.50). Check what your platform expects.
Never upload real bank details to GitHub or share them in issues.
Requirements
Microsoft Excel for Windows (2010 or later) with macros enabled
Contributing

Suggestions and improvements are welcome. Open an issue or a pull request.

Author

Md. Shafayat Hossain LinkedIn- https://www.linkedin.com/in/shafayat-hossain-305aa5173/

License

MIT License. Free to use, modify and share. See LICENSE.

This tool is provided as is, without warranty. You are responsible for checking every payment file before it is submitted.
