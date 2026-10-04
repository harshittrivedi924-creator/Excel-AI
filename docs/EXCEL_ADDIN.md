# Excel Add-in (Task Pane)

The `excel-addin/` folder contains a Microsoft Excel Add-in that lets you type
natural-language / Hinglish commands against your **currently open** workbook and
applies the results with Office.js.

```
Excel Workbook
      ↓
Excel Add-in / Task Pane (Office.js reads workbook context / selection)
      ↓
POST /api/excel/command  (commander + compact workbook snapshot)
      ↓
Existing Copilot → Parser → Validator → Engine → Handler pipeline
      ↓
Structured response (operation, parameters, result, machine-readable `writes`)
      ↓
Add-in applies the `writes` back into Excel with Office.js
      ↓
Excel updates
```

The Python backend never touches your open Excel instance or any file on disk.
Only the minimal snapshot (headers, data rows, origin coordinate, selection) is
sent to the API.

## Requirements

- Windows 10/11 with Microsoft Excel (2016+ or Microsoft 365).
- Python 3.10+ (the project `venv` already has everything).
- Node.js 18+ and npm (for the Add-in dev server).
- The Add-in API endpoints live on the existing Flask app; no extra service is
  needed.

## 1. Start the Flask backend

Use the existing venv:

```powershell
cd C:\ExcelOP\Excel-AI
.\venv\Scripts\activate
pip install -r requirements.txt
python main.py --serve --port 5000
```

Verify the Add-in endpoints are up (a second terminal):

```powershell
curl.exe http://127.0.0.1:5000/api/excel/health
```

Expected: `{"status":"ok","service":"excel-ai-copilot-addin-api","version":"1.0.0"}`

The Add-in API is CORS-enabled for the local dev-server origin
(`https://localhost:3000`) only. Everything else keeps its original same-origin
behaviour. To add another allowed origin:

```powershell
$env:EXCEL_ADDIN_ALLOWED_ORIGINS="https://localhost:3000,https://localhost:3001"
python main.py --serve --port 5000
```

## 2. Generate the HTTPS dev certificates (once)

The Add-in task pane must be served over HTTPS. Generate the local certificates
once per machine:

```powershell
cd C:\ExcelOP\Excel-AI\excel-addin
npx office-addin-dev-certs install
```

Choose **Yes** when it asks whether "localhost" can send traffic / the root CA
should be trusted.

## 3. Install dependencies and build (once)

```powershell
cd C:\ExcelOP\Excel-AI\excel-addin
npm install
npm run build
npm run validate     # validates the manifest against Microsoft's schema
```

`npm run validate` should end with "The manifest is valid."

## 4. Start the Add-in dev server

```powershell
cd C:\ExcelOP\Excel-AI\excel-addin
npm run dev-server
```

The task pane is now served at `https://localhost:3000/taskpane.html`. Keep this
terminal running.

> The manifest (`excel-addin/manifest.xml`) already points at
> `https://localhost:3000` — no manifest edits are required for local use.

## 5. Load the Add-in into Excel (sideloading)

For Excel for Windows with a Microsoft 365 subscription, sideload for testing:

1. With the dev server running from step 4, open the `excel-addin` folder in
   **Command Prompt** (not PowerShell) and run:

   ```bat
   cd C:\ExcelOP\Excel-AI\excel-addin
   npm start
   ```

   `office-addin-debugging` starts its own web server, sideloads the manifest and
   opens Excel with the Add-in already installed. If it reports "Vertex" content,
   ignore it.

2. If `npm start` does not open Excel automatically, or you prefer a manual check:
   - Open a blank workbook in Excel.
   - Go to **Insert → Add-ins → My Add-ins → Manage: Upload My Add-in**.
   - Browse to `C:\ExcelOP\Excel-AI\excel-addin\manifest.xml` and select it.
   - Wait for "Your add-in is ready" in the task pane.

3. From the **Home** ribbon, click **Open AI Copilot** in the **AI Copilot** group
   to show the task pane. It should report "backend online" in the top-right when
   the Flask server is running.

Stop the debug session later with:

```powershell
npm run stop
```

If your organisation blocks `office-addin-debugging`, sideload manually instead:
see "Upload My Add-in" above — Excel will keep the manifest registered for that
session.

## 6. Test with a real workbook

Open the sample workbook `C:\ExcelOP\Excel-AI\finance_sample.xlsx` (sheet `Sheet`,
headers `Month | Revenue | Expense | Profit`). Select a data range and try:

| Command | Expected result |
| --- | --- |
| `Revenue ka total karo` | Sum of Revenue (6600) |
| `Revenue ka average nikalo` | Average of Revenue (1320) |
| `Revenue ka maximum batao` | Max of Revenue (1800) |
| `Revenue ka percentage karo` | Each Revenue value as a % of the column total |
| `Revenue minus Expense karke Profit nikalo` | Profit formulas `=B2-C2` … written |
| `Revenue ka chart bana do` | A clustered column chart added in the sheet |
| `inka total karo` (with a numeric selection) | Sum of the selected cells, with a confirm to write it below the range |
| `bjsh fhewk qw` | Friendly "couldn't understand" error |
| `Revenue ka total karo aur B2 mein daal do` | Overwrite confirmation (B2 already has data), only writes after you accept |

Never-overwrite safety: the Add-in asks for confirmation before replacing any
occupied cell — either via the backend `overwrite_needed` signal or by checking
the destination cell through Office.js first.

### 6a. Test the AI Data Analyst

With the same table selected, the **🔍 Analyze data** chip (or typing
`Analyze this data`) runs a read-only analysis and renders a collapsible card
with key insights, data quality, a statistics table, growth, trends, anomalies
and a recommended chart. Follow it with a question such as `which month had the
highest revenue?`, or click **📈 Create Chart** to have the existing chart
pipeline draw the recommendation in Excel.

| Command | Expected result |
| --- | --- |
| `Analyze this data` | Analysis card for the selection, or the used range if nothing is selected |
| `Is data mein kya important hai?` | Same analysis (Hinglish trigger) |
| `Which month had the highest revenue?` | Answered from the stored report, no re-analysis |
| `Show revenue growth` | First → last change and % per numeric column |
| `Are there any unusual values?` | Outliers by IQR + z-score, or an explicit "none found" |
| `Compare revenue and expenses` | Side-by-side totals, gap and ratio |
| `Create the recommended chart` / **📈 Create Chart** | The chart is created in the sheet, using the first recommended column |
| Text-only selection | "No numeric column was found, so no statistics were computed." |
| Empty worksheet | "There is no data to analyze … Add some data first." |

Analysis never writes to the workbook. Only an explicit **Create Chart** or a
normal write command modifies Excel.

**Create Chart plots one series.** The recommendation may name two columns
(`Revenue vs Expense by Month`), but the existing chart pipeline creates the
chart for the first column only — a chart follow-up says so
("Revenue is charted first; add Expense separately"). To add the second series,
run a normal chart command for that column.

## API contract (Add-in only)

```
POST /api/excel/command
{
  "command": "Revenue ka total karo",
  "sheet_name": "Sheet",
  "origin": {"row": 1, "column": 1},
  "headers": ["Month", "Revenue", "Expense"],
  "rows": [["Jan", 1000, 600], ...],
  "selection": {"address": "B2:B6", "values": [[1000], ...], "row": 2,
                "column": 2, "row_count": 5, "column_count": 1},
  "allow_overwrite": false
}

GET /api/excel/health
```

Response (always structured — the Add-in never parses free text to decide
what to write):

```json
{
  "success": true,
  "message": "Done. The result of Sum of B2:B6 is 6600.",
  "operation": "SUM",
  "parameters": {"operation": "SUM", "column": "B", "values": [...]},
  "result": 6600,
  "destination": "B8",
  "written_count": 0,
  "overwrite_needed": false,
  "writes": {
    "cells":  [{"row": 8, "column": 2, "formula": "=SUM(B2:B6)"}],
    "sheets": {"Filtered": {"values": [["Month", ...], ...]}},
    "charts": [{"type": "ColumnClustered", "title": "Revenue",
                "source_range": "B1:B6"}]
  }
}
```

## Project layout

```
excel-addin/
  manifest.xml               # Office add-in manifest (points to https://localhost:3000)
  package.json               # build / dev-server / sideload scripts
  webpack.config.js          # serves over HTTPS using office-addin-dev-certs
  src/
    commands/commands.html   # (reserved) command-function page
    taskpane/
      taskpane.html          # AI sidebar UI
      taskpane.css           # styling
      taskpane.js            # Office.js context reader + API client + write applier
      format.mjs             # pure response formatters (unit-tested with node --test)
      snapshot.mjs           # used-range scanner + snapshot builder (unit-tested)
  assets/                    # manifest icons
  dist/                      # generated by `npm run build`

backend/api/addin_routes.py        # POST /api/excel/command, POST /api/excel/analyze,
                                   # GET /api/excel/health
backend/excel/memory_handler.py    # in-memory workbook for the copilot pipeline
backend/excel/data_analyzer.py     # read-only profiling/statistics/trends/anomalies
backend/excel/analysis_followup.py # deterministic follow-up answers
tests/test_addin_routes.py         # API contract tests
tests/test_data_analyzer.py        # analyzer + analyze-endpoint tests
```