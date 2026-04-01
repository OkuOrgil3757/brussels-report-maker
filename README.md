# Brussels Report Maker

A local website that takes any Excel file and instantly generates interactive charts — bar charts, pie charts, heatmaps, treemaps, trend lines, and more. No internet required after setup. Runs entirely on your computer.

---

## What it looks like

Upload an Excel file → pick a sheet → click **Generate all charts** → get 15+ interactive charts you can zoom, hover, and download as PNG.

Charts it creates automatically:
- Daily revenue trend with 7-day rolling average
- Cumulative total over time
- Revenue by day of week (highlights best day)
- Monthly comparison
- Revenue by category (bar + donut pie)
- Top 20 items by value or quantity
- Employee performance comparison
- Category mix % per employee (stacked 100% bar)
- Weekly heatmap (week × category)
- Treemap (nested area by category/item)
- Surplus vs deficit diverging bar (for inventory data)
- Scatter plot for two numeric columns
- Distribution histogram + box plot
- Summary statistics table

---

## Step-by-step setup on Windows

### Step 1 — Install Python

1. Open your browser and go to: **https://www.python.org/downloads/**
2. Click the big yellow **Download Python 3.x.x** button (any version 3.10 or newer is fine)
3. Run the downloaded `.exe` installer
4. **IMPORTANT:** On the very first screen of the installer, check the box that says **"Add Python to PATH"** (it is unchecked by default — you must check it)
5. Click **Install Now**
6. Wait for it to finish, then click **Close**

To verify Python installed correctly:
- Press `Win + R`, type `cmd`, press Enter
- In the black window that opens, type:
  ```
  python --version
  ```
- You should see something like `Python 3.12.0`. If you see an error, Python is not on PATH — reinstall and make sure to check "Add Python to PATH"

---

### Step 2 — Download this project

**Option A — Download as ZIP (easiest, no Git needed)**

1. Go to this project's GitHub page
2. Click the green **Code** button
3. Click **Download ZIP**
4. Find the downloaded ZIP file (usually in your `Downloads` folder)
5. Right-click the ZIP → **Extract All**
6. Choose where to extract it (e.g. `C:\Users\YourName\Desktop\`)
7. You now have a folder called something like `brussels-report-maker-main`

**Option B — Clone with Git**

If you have Git installed:
```
git clone https://github.com/OkuOrgil3757/brussels-report-maker.git
```

---

### Step 3 — Open the project folder in Command Prompt

1. Press `Win + R`, type `cmd`, press Enter
2. Navigate to the project folder using the `cd` command.

   For example, if you extracted to your Desktop:
   ```
   cd C:\Users\YourName\Desktop\brussels-report-maker-main
   ```
   Replace `YourName` with your actual Windows username and the folder name with whatever yours is called.

   **Tip:** You can also type `cd ` (with a space after), then drag-and-drop the folder from Explorer into the Command Prompt window — it will paste the path automatically.

3. Press Enter. Your command prompt should now show the project folder path.

---

### Step 4 — Install the required Python packages

In the same Command Prompt window, type this exactly and press Enter:

```
pip install -r requirements.txt
```

This downloads and installs Flask, Pandas, Plotly, and other packages the app needs. It may take 1–3 minutes depending on your internet speed. You will see a lot of text scrolling — that is normal.

When it finishes you should see something like:
```
Successfully installed flask-3.x.x pandas-2.x.x plotly-5.x.x ...
```

If you see `pip is not recognized`, try:
```
python -m pip install -r requirements.txt
```

---

### Step 5 — Start the website

In the same Command Prompt window, type:

```
python app.py
```

You should see:
```
============================================================
  Brussels Report Maker
  Open http://localhost:8080 in your browser
============================================================
 * Running on http://0.0.0.0:8080
```

**Do not close this Command Prompt window** — it must stay open while you use the website.

---

### Step 6 — Open the website

1. Open any web browser (Chrome, Edge, Firefox)
2. In the address bar, type:
   ```
   http://localhost:8080
   ```
3. Press Enter
4. You should see the Brussels Report Maker homepage

---

### Step 7 — Use it

1. Drag and drop your Excel file (`.xlsx` or `.xls`) onto the upload area, or click the area to browse for a file
2. A sheet selector will appear — pick which sheet contains your data
3. Click **Preview columns** to see how the app detected your data (dates, numbers, categories)
4. If your headers are not on row 1, set **Skip rows** to the correct number
5. Click **Generate all charts**
6. All charts appear below — you can:
   - Hover over any chart for exact values
   - Zoom in by clicking and dragging
   - Double-click to zoom back out
   - Click the download button (⬇️) on any chart to save it as a PNG image

---

## Stopping the server

When you are done, go back to the Command Prompt window and press `Ctrl + C`. The server will stop.

---

## Starting it again next time

You do not need to run `pip install` again. Just:

1. Open Command Prompt
2. Navigate to the project folder:
   ```
   cd C:\Users\YourName\Desktop\brussels-report-maker-main
   ```
3. Run:
   ```
   python app.py
   ```
4. Open `http://localhost:8080` in your browser

Or just double-click **`start.bat`** — it does all of this automatically.

---

## Quick start (alternative — double-click)

Instead of using Command Prompt at all, you can simply:

1. Open the project folder in File Explorer
2. Double-click **`start.bat`**
3. A black window opens, installs anything missing, and starts the server
4. Open `http://localhost:8080` in your browser

---

## Troubleshooting

**`python` is not recognized as a command**
- You did not check "Add Python to PATH" during installation
- Uninstall Python from Control Panel, then reinstall it and check that box

**`pip install` fails with a permissions error**
- Try running Command Prompt as Administrator: press `Win`, search `cmd`, right-click → "Run as administrator", then repeat Step 3–4

**Port 8080 is already in use**
- Something else on your computer is already using port 8080
- Open `app.py` in Notepad, find the last line with `port=8080`, change it to `port=8081`, save, then use `http://localhost:8081` instead

**The page shows but charts don't appear**
- Make sure your Excel file has at least one sheet with actual data
- Try clicking "Preview columns" first — if it shows your columns correctly, then click Generate
- Check that your data has a mix of text, number, and/or date columns

**Charts look wrong / weird columns detected**
- Use the "Skip rows" option if your Excel file has title rows or blank rows before the actual headers
- The app auto-detects column types — if a number column shows as "text", it means the column has mixed content (some cells have text like "N/A" mixed with numbers)

**I only see a black Command Prompt window and it closes immediately**
- Python is not installed or not on PATH — go back to Step 1
- Or there is an error — try running `python app.py` manually to see the error message

---

## File size and privacy

- Your Excel files are processed **entirely on your computer** — nothing is uploaded to the internet
- The app temporarily saves the last uploaded file in the `uploads/` folder inside the project
- Max file size: 50 MB

---

## Requirements (installed automatically)

| Package | Purpose |
|---------|---------|
| Flask | Web server |
| Pandas | Excel reading and data processing |
| Plotly | Interactive charts |
| OpenPyXL | .xlsx file support |
| xlrd | .xls (old format) file support |
| NumPy | Numerical calculations |

Python 3.10 or newer required.

---

## Project structure

```
brussels-report-maker/
├── app.py              ← Main application (Flask backend + chart logic)
├── templates/
│   └── index.html      ← Web interface
├── requirements.txt    ← Python package list
├── start.bat           ← Windows one-click launcher
├── uploads/            ← Temporary folder for uploaded files (auto-created)
└── README.md           ← This file
```
