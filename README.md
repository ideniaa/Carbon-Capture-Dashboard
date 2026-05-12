# Local Team Dashboard

This is a Flask-based internal dashboard prototype meant to stay on your local team network.

## What it includes

- Live dashboard cards for machine status
- Recent readings chart and history table
- Protected control buttons with confirmation prompts
- XLSX export for Excel
- Manual chlorine measurement entry with timestamp and note
- SQL Server-backed storage for chlorine, results, titrator history, and voltage data

## Run locally

1. Create a virtual environment if you want:
   `python -m venv .venv`
2. Activate it:
   `.\.venv\Scripts\Activate.ps1`
3. Install packages:
   `pip install -r requirements.txt`
4. Start the app:
   `python app.py`
5. Open:
   `http://127.0.0.1:5000`

## Local-team access

When you are ready to share it only inside your team:

- Run the Flask app on a machine inside your local network
- Use a local IP or machine name like `http://your-pc-name:5000`
- Do not expose the port publicly
- Restrict access with your firewall or network rules

## SQL Server setup

For the titrator, the website defaults to the same SQL table used by `C:\Titrolyzer\Titrolyzer.py`:

- Server: set via `SQL_SERVER` env var (or read from `Titrolyzer.py` at startup)
- Database: set via `SQL_DATABASE` env var
- Table: `dbo.Titrolyzer`
- Columns: `RecordID`, `ReadingTime`, `Timestamp`, `Result_ID`, `Value`

If `C:\Titrolyzer\Titrolyzer.py` exists, the website reads those SQL constants from that script at startup. Environment variables still override the script values.

The app now supports two SQL connection styles:

- A full `SQL_SERVER_CONNECTION_STRING`
- Or separate environment variables for server, database, and login

Example PowerShell setup with SQL auth:

```powershell
$env:SQL_SERVER="YOUR_SERVER\SQLEXPRESS"
$env:SQL_DATABASE="YOUR_DATABASE"
$env:SQL_USERNAME="your_username"
$env:SQL_PASSWORD="your_password"
```

If you prefer, you can still use one full connection string instead:

```powershell
$env:SQL_SERVER_CONNECTION_STRING="DRIVER={SQL Server};SERVER=YOUR_SERVER\SQLEXPRESS;DATABASE=YOUR_DATABASE;UID=your_username;PWD=your_password;"
```

Behavior:

- `dbo.Chlorine` is created automatically if it does not exist.
- `dbo.Results` is created automatically if it does not exist.
- `dbo.Titrolyzer` is created automatically if it does not exist and is read using these default columns:
  `RecordID`, `ReadingTime`, `Timestamp`, `Result_ID`, `Value`
- `dbo.Voltage` is reused on every start. If it does not exist yet, WINDAQ-MSSQL creates it with the connected device channels.

Voltage defaults:

- Fixed table: `Voltage`
- Auto-detect table pattern: `V%` only if fixed-table behavior is disabled in code
- Row id column: `row_num`
- Timestamp column: `date_time`
- Channel columns: read from the actual `dbo.Voltage` table; the website does not create guessed voltage channels
- Note column: optional and blank by default

If your voltage table uses different names, set these environment variables before starting Flask:

```powershell
$env:SQL_VOLTAGE_TABLE="Voltage"
$env:SQL_VOLTAGE_TABLE_PATTERN="V%"
$env:SQL_VOLTAGE_ROW_ID_COLUMN="row_num"
$env:SQL_VOLTAGE_TIMESTAMP_COLUMN="date_time"
$env:SQL_VOLTAGE_CH1_COLUMN="channel_1"
$env:SQL_VOLTAGE_CH2_COLUMN="channel_2"
$env:SQL_VOLTAGE_CH3_COLUMN="channel_3"
$env:SQL_VOLTAGE_CH4_COLUMN="channel_4"
$env:SQL_VOLTAGE_NOTE_COLUMN=""
```

Notes:

- If SQL is not configured, the app falls back to local seeded/manual data.
- The chlorine and results tables can be cleared from the dashboard.
- The titrator clear action deletes rows from the SQL-backed titrator table.
- The voltage clear action deletes rows from `dbo.Voltage`.
- The chlorine table SQL is also included in `sql/chlorine_measurements.sql`.
- The voltage chart reads channel columns from `dbo.Voltage`, excluding the timestamp, row id, and optional note columns. Once rows exist, it only shows columns that actually contain data.
- If your SQL voltage table does not have a note column, voltage notes will stay in the UI but will not persist to SQL unless you add one.

## WINDAQ-MSSQL control

The voltage page can try to launch and stop `WINDAQ-MSSQL` from the website if you tell the app where the recorder executable lives.

```powershell
$env:WINDAQ_MSSQL_EXE="C:\Path\To\WINDAQ-MSSQL.exe"
$env:WINDAQ_MSSQL_PROCESS_NAME="WINDAQ-MSSQL"
$env:WINDAQ_MSSQL_PID_FILE="C:\Path\To\windaq_mssql.pid"
$env:WINDAQ_MSSQL_AUTO_START="1"
```

Notes:

- `WINDAQ_MSSQL_EXE` is used for the Start button.
- `WINDAQ_MSSQL_PROCESS_NAME` is used for the Stop button. If you leave it unset, the app derives the process name from the executable path.
- `WINDAQ_MSSQL_AUTO_START=1` makes the website select the first listed DATAQ device, fill the SQL fields from the website settings, click Connect, and send WINDAQ-MSSQL's Start command.
- Before launching, the website sets WINDAQ-MSSQL's saved Table Name setting to `dbo.Voltage`. If the table does not exist yet, WINDAQ-MSSQL creates it with the connected device's real channel columns.
- If `dbo.Voltage` exists but has zero rows, the Start button lets WINDAQ-MSSQL recreate it so the column count matches the connected device instead of an old guessed schema.
- The Start button will not open a second WINDAQ-MSSQL session if one is already running. The app tracks the started process with a local PID file, defaulting to `windaq_mssql.pid` in the website folder.
- WINDAQ-MSSQL still requires WinDaq Acquisition to be running. If that program is not acquiring data, WINDAQ-MSSQL may open a prompt after the website sends Start.
- The voltage Excel export always uses the website layout: `Timestamp`, voltage channel columns, and `Note`.

## Titrolyzer app control

The Titrator tab has App commands for starting and stopping the Titrolyzer app. By default it launches:

```powershell
C:\Titrolyzer\Titrolyzer.py
```

Override the script path or Python executable if needed:

```powershell
$env:TITROLYZER_APP_PATH="C:\Titrolyzer\Titrolyzer.py"
$env:TITROLYZER_PYTHON_EXE="C:\Path\To\python.exe"
```

The Stop button closes the process that was started from the website by using a local PID file. It will not kill unrelated Python processes.

## Titrator machine commands

The Titrator tab writes TRUE to the Program Control coils from the communication manual. The website displays the manual addresses, while the Python Modbus client writes to zero-based offsets internally:

- Start Sequence: manual coil `1`
- Stop Sequence: manual coil `2`
- Break Sequence: manual coil `3`
- Trigger Program A: manual coil `33`
- Trigger Program B: manual coil `34`
- Trigger Program X: manual coil `35`


## Optional SQL Server driver

The dashboard can run without SQL Server first, but live SQL viewing requires a SQL Server ODBC driver installed on Windows plus the Python packages in `requirements.txt`.
