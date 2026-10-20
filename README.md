# ExceltoSQL

![Windows](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white)
![Python 3.12](https://img.shields.io/badge/built%20with-Python%203.12-3776AB?logo=python&logoColor=white)
![Release 2.0](https://img.shields.io/badge/version-2.0-informational)

A small Windows tool that turns an Excel sheet of securities into a ready-to-run SQL script. Pick the `.xlsx` file, pick where to save the result, done.

> [!NOTE]
> This repository contains the compiled Windows program only (`ExcelToSQL_2.0.exe`). It is a single-purpose tool built for one specific table, not a general Excel-to-SQL converter.

## What it does

The program reads the **active sheet** of the workbook. Row 1 is treated as the header and skipped; every following row becomes one `INSERT`. The script always starts by emptying the table:

```sql
TRUNCATE TABLE OMS_S.CAMPAIGN_SECURITY;
INSERT INTO OMS_S.CAMPAIGN_SECURITY (ISIN, ISSUER, MARKET, PROPERTY) VALUES ('DE0007164600','SAP SE','XETRA','A');
INSERT INTO OMS_S.CAMPAIGN_SECURITY (ISIN, ISSUER, MARKET, PROPERTY) VALUES ('US0378331005','Apple Inc.','NASDAQ','B');
```

### Expected sheet layout

| | A | B | C | D |
|---|---|---|---|---|
| **1** | ISIN | Issuer | Market | Property |
| **2** | DE0007164600 | SAP SE | XETRA | A |
| **3** | US0378331005 | Apple Inc. | NASDAQ | B |

- Columns **A–D** map to `ISIN`, `ISSUER`, `MARKET`, `PROPERTY`, in this order. Header names are not checked.
- Every cell must contain **text**. An empty cell or a number stops the conversion without an error message: the success message does not appear and the script ends at that row. Format the columns as *Text* in Excel if needed.
- Values are written as they are. A value containing an apostrophe (`'`), e.g. `O'Neil`, produces invalid SQL. Double it (`O''Neil`) in the sheet first.
- Rows below the data that still have formatting count as rows too. Delete them if the conversion fails at the end of the sheet.

## Download

1. Download [`ExcelToSQL_2.0.exe`](https://github.com/FishMacc/ExceltoSQL/raw/main/ExcelToSQL_2.0.exe) (~15 MB).
2. Double-click it. No installation and no Python needed.

Windows SmartScreen may warn about an unknown publisher because the file is not signed. Click **More info → Run anyway** if you trust the source.

## Usage

1. Start `ExcelToSQL_2.0.exe` and click **Excel to SQL Converter**.
2. Confirm the hint and select the `.xlsx` file.
3. Confirm the next hint and choose where to save the script. It is saved as `.txt` by default; you can type a name ending in `.sql` instead.
4. A message shows where the file was saved. Run the script in your database client.

If no success message appears, check the sheet against the layout above. Cancelling one of the file dialogs also ends the run without writing anything.

The tool's dialogs are in German, e.g. *Beenden* closes the program.

## How it was built

Python 3.12 with [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) for the window and [openpyxl](https://openpyxl.readthedocs.io/) for reading Excel files, packaged into a single executable with [PyInstaller](https://pyinstaller.org/).

## Repository contents

| File | Description |
|------|-------------|
| `ExcelToSQL_2.0.exe` | The Windows program (version 2.0, October 2023) |
| `README.md` | This file |
| `.gitignore` | Ignore rules for local development files |

## License

No license has been chosen yet, so all rights are reserved by the author.
