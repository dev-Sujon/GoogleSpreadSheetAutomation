# ERP_GoogleSheet_Automation — Hard-coded Configuration Edition

This project is a simple, production-oriented ERP report uploader for the workflow:

```text
ERP / pandas reports
        ↓
reports = [(name, dataframe), ...]
        ↓
build_master_dataframe()
        ↓
master_df (local inspection)

reports
        ↓
build_payload()
        ↓
strict JSON-safe payload
        ↓
HTTP POST
        ↓
Google Apps Script Web App
        ↓
create/replace previous-day sheet
        ↓
DD_MMM_YYYY
```

## Important change in this edition

This version intentionally does **not** use `.env`, `python-dotenv`, Script Properties, or any external configuration service.

There are only two hard-coded configuration locations:

```text
Python:
src/config.py

Google Apps Script:
apps_script/Code.gs
```

The same values must be used in both places:

```text
APPS_SCRIPT_URL / Web App /exec URL
API_TOKEN / shared secret
SPREADSHEET_ID / target Google Spreadsheet ID
```

Because credentials are hard-coded, keep this project private and do not publish `config.py` or `Code.gs` to a public GitHub repository.

---

# 1. Final directory

```text
ERP_GoogleSheet_Automation/
│
├── .gitignore
├── README.md
├── requirements.txt
│
├── notebook/
│   └── ERP_Report_Automation.ipynb
│
├── src/
│   ├── __init__.py
│   ├── config.py
│   ├── master_builder.py
│   ├── json_payload.py
│   └── google_uploader.py
│
└── apps_script/
    └── Code.gs
```

No `.env` file is required.

---

# 2. Configure Python

Open:

```text
src/config.py
```

Change only these values:

```python
APPS_SCRIPT_URL = "PASTE_YOUR_APPS_SCRIPT_EXEC_URL_HERE"
API_TOKEN = "PASTE_YOUR_API_TOKEN_HERE"
SPREADSHEET_ID = "PASTE_YOUR_SPREADSHEET_ID_HERE"
```

Also keep:

```python
REPORT_TIMEZONE = "Asia/Dhaka"
HTTP_TIMEOUT_SECONDS = 120
MAX_RETRIES = 2
RETRY_BACKOFF_SECONDS = 2
```

The validation function checks that placeholders were removed, the URL ends with `/exec`, the token is long enough, and the spreadsheet ID is non-empty.

---

# 3. Configure Apps Script

Open:

```text
apps_script/Code.gs
```

At the top, change:

```javascript
const SPREADSHEET_ID = "PASTE_YOUR_SPREADSHEET_ID_HERE";
const API_TOKEN = "PASTE_YOUR_API_TOKEN_HERE";
```

The `API_TOKEN` here must be exactly the same as `API_TOKEN` in `src/config.py`.

The `SPREADSHEET_ID` is the ID from the target Google Sheets URL.

Example URL:

```text
https://docs.google.com/spreadsheets/d/XXXXXXXXXXXX/edit
```

The spreadsheet ID is the value between `/d/` and `/edit`.

---

# 4. Install dependencies

From Anaconda Prompt:

```bat
cd C:\path\to\ERP_GoogleSheet_Automation
python -m pip install -r requirements.txt
```

Required packages:

```text
pandas
numpy
requests
tzdata
jupyterlab
```

`python-dotenv` is intentionally not included.

---

# 5. Apps Script deployment

Open the target spreadsheet:

```text
Extensions → Apps Script
```

Replace `Code.gs` with the supplied `apps_script/Code.gs`.

Edit the two hard-coded values.

Save.

Deploy:

```text
Deploy → New deployment → Web app
```

Use the appropriate execution identity and access setting for your Google environment. The Python client is not a browser session, so the deployed Web App must be accessible to the Python request.

Copy the deployed URL ending in:

```text
/exec
```

Put it into `src/config.py`.

Do not use the `/dev` testing URL for the normal notebook upload.

---

# 6. First test from Apps Script

In the Apps Script editor, run:

```text
testWrite()
```

This creates:

```text
__ERP_AUTOMATION_TEST__
```

and writes two sample reports.

Then verify the sheet.

After successful testing, run:

```text
deleteTestSheet()
```

---

# 7. Health test from Jupyter

Open:

```text
notebook/ERP_Report_Automation.ipynb
```

The first setup cell calls:

```python
validate_config()
```

Then:

```python
health = check_apps_script_health()
print(health)
```

Expected:

```text
{'success': True, ...}
```

If this fails, do not continue to the report upload. Fix the deployment/configuration first.

---

# 8. Daily date rule

The notebook does not use the computer's local timezone directly.

It uses:

```python
REPORT_TIMEZONE = "Asia/Dhaka"
```

Then:

```python
previous_day = get_previous_report_date()
```

So if the Bangladesh business date is:

```text
21-Sep-2026
```

the target report date is:

```text
20-Sep-2026
```

and the target Google Sheet tab is:

```text
20_Sep_2026
```

This is the only daily-tab rule used by the project.

---

# 9. Report structure

Keep your existing ERP report generation code in the notebook.

Then build one list:

```python
reports = [
    ("Export Invoice Report", export_invoice_df),
    ("Ex Factory Report", ex_factory_df),
    ("Production Report", production_df),
]
```

Each DataFrame may have:

- different columns
- different column order
- different row counts
- different data types
- empty cells
- `NaN`
- `NaT`
- `None`
- NumPy scalar values

No common schema is required.

---

# 10. Master DataFrame

```python
master_df = build_master_dataframe(reports)
```

This is primarily a local inspection representation.

It preserves each report's headers/order and inserts exactly two completely blank rows between reports.

The original DataFrames are not modified.

---

# 11. JSON payload

```python
payload = build_payload(
    reports=reports,
    report_date=previous_day,
)
```

The payload contains:

```text
version
run_id
report_date
reports[]
```

Each report contains:

```text
name
headers[]
rows[][]
```

Before HTTP upload the payload is serialized with:

```python
json.dumps(payload, allow_nan=False)
```

Therefore invalid JSON values fail locally rather than reaching Apps Script.

---

# 12. Upload behavior

The notebook calls:

```python
result = upload_to_google_apps_script(payload)
```

The uploader:

1. validates configuration
2. adds the API token to a copy of the payload
3. sends one POST request
4. follows redirects
5. retries short-lived 429/5xx/network failures
6. parses the response as JSON
7. checks `success`
8. verifies the returned `run_id` when available

The same daily write is idempotent because Apps Script clears and rewrites the target daily tab.

---

# 13. Google Sheet write behavior

For each report:

```text
Report title
Header
Data
```

Then, between reports only:

```text
blank
blank
```

The next report begins after those two rows.

Every report uses its own width.

For example:

```text
Report A → 4 columns
Report B → 7 columns
Report C → 3 columns
```

There is no attempt to force these into one common Google range.

---

# 14. Re-run behavior

Running the notebook multiple times for the same report date does not create duplicate daily tabs.

Example:

```text
20_Sep_2026
```

First run:

```text
create sheet
write reports
```

Second run:

```text
find existing sheet
clear contents
rewrite reports
```

Third run behaves the same way.

---

# 15. Recommended production workflow

Use this order when integrating your real ERP reports:

```text
1. Configure Python constants
        ↓
2. Configure Code.gs constants
        ↓
3. Deploy /exec Web App
        ↓
4. Run Apps Script testWrite()
        ↓
5. Confirm target spreadsheet write
        ↓
6. Run notebook health check
        ↓
7. Run three small sample DataFrames
        ↓
8. Confirm 20_Sep_2026-style sheet name
        ↓
9. Run the same payload twice
        ↓
10. Confirm no duplicate sheet and no duplicate rows
        ↓
11. Replace samples with real ERP reports
        ↓
12. Run the complete notebook
```

---

# 16. Troubleshooting sequence

If notebook says `UPLOAD SUCCESS` but you cannot find the result:

```text
1. Confirm the returned sheet name.
2. Search that exact tab in Google Sheets.
3. Run Apps Script testWrite().
4. Check Apps Script Executions.
5. Check that SPREADSHEET_ID is the correct file.
6. Check API_TOKEN matches on both sides.
7. Check Web App URL ends in /exec.
```

If notebook gets:

```text
HTTP 403
```

check Web App access/deployment permissions.

If notebook gets:

```text
invalid JSON
```

the Web App likely returned HTML, a Google login page, or another non-JSON response.

If Apps Script reports:

```text
Unauthorized: invalid API token.
```

compare these exact values:

```text
src/config.py
        API_TOKEN

apps_script/Code.gs
        API_TOKEN
```

If Apps Script reports:

```text
SPREADSHEET_ID is not configured
```

edit the constant at the top of `Code.gs`.

---

# 17. Security note

Hard-coding credentials is technically simple, but it has a real consequence: the credentials are source-code secrets.

Therefore:

```text
DO NOT
- publish the repository publicly
- push config.py to a public GitHub repository
- share Code.gs containing the real token
- paste the token into screenshots/logs
```

The project deliberately follows your requested hard-coded design; it does not pretend this has the same secret-management properties as a `.env`/secret manager approach.

---

# 18. Final operating model

The notebook should eventually feel like this:

```python
# Generate ERP reports
export_invoice_df = ...
ex_factory_df = ...
production_df = ...

# Register reports
reports = [
    ("Export Invoice Report", export_invoice_df),
    ("Ex Factory Report", ex_factory_df),
    ("Production Report", production_df),
]

# Preview
master_df = build_master_dataframe(reports)
display(master_df)

# Previous Bangladesh business day
previous_day = get_previous_report_date()

# JSON-safe payload
payload = build_payload(reports, previous_day)

# Upload
result = upload_to_google_apps_script(payload)
print(result)
```

That is the intended steady-state workflow.
