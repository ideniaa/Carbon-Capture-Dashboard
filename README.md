# Carbon Capture Dashboard

A web dashboard for tracking carbon capture lab data. It runs on your computer and lets your whole team view and log readings from one shared webpage.

---

## What does it do?

The dashboard has six pages (tabs) you can switch between:

| Page | What it's for |
|---|---|
| **Home** | See a summary of everything at once — latest readings, mini charts, and team announcements |
| **Titrator** | View live titrator readings (R1 & R2), send machine commands, and see a chart of recent values |
| **Chlorine** | Log chlorine measurements (mg/L) with timestamps and notes |
| **Results** | Enter experiment results like pH, current, flow rates, and CIER |
| **Voltage** | See voltage channel readings from the WINDAQ recorder |
| **WinCC** | View process data pulled directly from WinCC SQL tables |

You can also download any table as an Excel file from within each page.

---

## How to run it

> You need Python installed. If you don't have it, download it from [python.org](https://www.python.org/downloads/).

### Step 1 — Download or clone the project

If you have Git:
```bash
git clone https://github.com/ideniaa/Carbon-Capture-Dashboard.git
cd Carbon-Capture-Dashboard
```

Or just download the ZIP from GitHub and unzip it.

### Step 2 — Create a virtual environment (optional but recommended)

This keeps the project's packages separate from the rest of your system.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

On Mac/Linux:
```bash
python3 -m venv .venv
source .venv/bin/activate
```

### Step 3 — Install the required packages

```bash
pip install -r requirements.txt
```

### Step 4 — Set up your settings

```bash
copy .env.example .env
```

Then open `.env` in any text editor and fill in your details (see the Configuration section below).

### Step 5 — Start the app

```bash
python app.py
```

### Step 6 — Open in your browser

Go to: **http://127.0.0.1:5000**

> To share with teammates on the same network, use your computer's local IP address instead, like `http://192.168.1.50:5000`.

---

## Configuration (the .env file)

When you copied `.env.example` to `.env`, it created a settings file. Here's what each section means:

### Basic setting

```
SECRET_KEY=change-me-to-a-random-string
```
Change this to any random string. It keeps your session secure.

### Connecting to the titrator machine (Modbus)

```
MODBUS_HOST=           ← IP address of the titrator (e.g. 10.0.0.50)
MODBUS_PORT=502        ← leave this as 502 unless told otherwise
MODBUS_UNIT_ID=2       ← unit ID of the device
```

If you leave `MODBUS_HOST` blank, the buttons still work in the UI — they just won't send real commands to a machine. Good for testing.

### Connecting to SQL Server (for live data)

Fill in your database details:

```
SQL_SERVER=YOUR_SERVER\SQLEXPRESS
SQL_DATABASE=YOUR_DATABASE
SQL_USERNAME=your_username
SQL_PASSWORD=your_password
```

> If you don't have SQL Server set up, that's fine — the app uses built-in sample data instead and still loads normally.

### WinDAQ recorder (Voltage page)

```
WINDAQ_MSSQL_EXE=C:\WINDAQMSSQL\WINDAQMSSQL.exe
```

Set this to the path of your WinDAQ-MSSQL app. The Voltage page can then start and stop it from the browser.

### Titrolyzer app (Titrator page)

```
TITROLYZER_APP_PATH=C:\Titrolyzer\Titrolyzer.py
```

Set this to the path of your Titrolyzer script. The Titrator page can then start and stop it from the browser.

---

## Files in this project

```
app.py               The main app — all the logic lives here
requirements.txt     List of Python packages needed
.env.example         Template for your settings file
index.html           Older standalone demo page (no Flask needed)
script.js            JavaScript for the standalone demo page

templates/           HTML pages for each tab
  home.html          Home page
  titrator.html      Titrator page
  chlorine.html      Chlorine page
  results.html       Results page
  voltage.html       Voltage page
  wincc.html         WinCC page
  base.html          Shared layout (nav bar, etc.)

static/
  styles.css         Styling for the whole site
  script.js          Live updates for the titrator page
  inline-notes.js    Click-to-edit notes in tables

sql/
  chlorine_measurements.sql   SQL script to create the Chlorine table manually
```

---

## Common questions

**The app loaded but I don't see any real data.**
That's normal if SQL Server isn't set up. The app shows sample/seeded data so you can still explore the interface.

**How do I share it with my team?**
Run the app on one computer, then everyone on the same network can open `http://<your-ip>:5000` in their browser.

**Where does my data get saved?**
- If SQL Server is configured: data is saved to the database.
- If not: titrator and chlorine data stay in memory while the app is running. Notes and command logs are saved as small JSON files next to `app.py`.

**How do I export data?**
Each tab has a "Download Excel Sheet" button. The Home page also has a bulk download where you can pick which tables to export.

**What Python version do I need?**
Python 3.10 or newer is recommended.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `pip install` fails | Make sure your virtual environment is activated first |
| App won't start | Check that all packages installed without errors |
| Can't connect to SQL | Double-check the server name, database name, and credentials in `.env` |
| Teammates can't reach the site | Make sure your firewall allows connections on port 5000 |
| WinDAQ won't start from browser | Set `WINDAQ_MSSQL_EXE` to the correct path in `.env` |
