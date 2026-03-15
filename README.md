# SPUploader

SharePoint Upload Tool is a Windows GUI app that validates an Excel file, previews its contents, and uploads rows into a SharePoint list using a Microsoft Access bridge. It keeps Access hidden, runs a macro to append rows, and logs activity to a rotating log file.

## What the app does
- Validates that required columns exist and are not empty.
- Previews the first 100 rows before upload.
- Imports the Excel file into a local Access temp table.
- Runs an Access macro to append rows to a linked SharePoint list.
- Tests connectivity and authentication against SharePoint.
- New: multi-region “New Starters Schedule” tab with separate London + Dublin names and combined week rotation.
- Auto-saves the London/Dublin name lists so they are preserved on next launch.

## Requirements
- Windows 10 or later.
- Python 3.9+.
- Microsoft Access installed (for COM automation).
- SharePoint list with permissions to read/write.

Python packages are listed in requirements.txt.

## Setup
1) Create and activate a virtual environment (recommended).
2) Install dependencies:

```bash
pip install -r requirements.txt
```

3) Create config.ini in the same folder as uploader.py (you can copy and edit the example from src/config.ini).

Example config.ini:
```ini
[access]
db_path = C:\Path\To\Your\Database.accdb
macro_name = AppendToSharePoint
linked_table = SP_Devices
temp_table = TempImportTable

[validation]
required_columns = Old Hostname,Serial Number,Model,New Hostname,AD

[logging]
log_file = logs/uploader.log
log_level = INFO
```

4) Run the app:
```bash
python uploader.py
```

## Access test case setup
This is a minimal setup to validate end-to-end uploads.

1) Create or open an Access database
- Create a new .accdb or use an existing one.
- Ensure the file path matches access.db_path in config.ini.

2) Link the SharePoint list
- Access: External Data -> New Data Source -> From Online Services -> SharePoint List.
- Enter the site URL and choose "Link to the data source".
- Select your target list and finish.
- Confirm the linked table name matches access.linked_table in config.ini.

3) Create a local temp table
- Create a local table named exactly as access.temp_table.
- Columns must match the Excel headers you will upload.
- Use compatible data types for the SharePoint list.

4) Create an append query
- Create a query in Design view.
- Add the temp table.
- Convert to Append Query and select the linked SharePoint table as the target.
- Map fields 1:1 from temp table to SharePoint table.
- Save the query (for example: AppendToSharePoint).

5) Create a macro
- Create a macro named exactly as access.macro_name.
- Add an OpenQuery action pointing to the append query.
- Save the macro.

## Create a test case
1) Create a small Excel file with the required columns from config.ini.
Example headers (row 1):
- Old Hostname
- Serial Number
- Model
- New Hostname
- AD

2) Add 2-3 test rows.
3) In the app:
- Select the Excel file.
- Click Validate File (ensure it passes).
- Click Preview Data (sanity check the rows).
- Click Upload to SharePoint.

## Notes
- The app imports into the temp table, then runs the macro to append to SharePoint.
- The Test Connection tool verifies permissions by opening the linked table.
- Logs are written to logs/uploader.log.
