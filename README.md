# Vendor Transaction Tracker

Drag in a monthly bank XLS → classifies every debit by vendor → pushes three tabs to a new Google Sheet.

---

## Folder structure

```
vendor-tracker/
├── app.py                  ← the app
├── List_of_Vendors.csv     ← your keyword map (you maintain this)
├── credentials.json        ← Google Service Account key (you place this once)
├── requirements.txt
└── README.md
```

---

## One-time setup

### 1 — Python environment

```bash
cd vendor-tracker
pip install -r requirements.txt
```

Python 3.10 or newer recommended.

---

### 2 — Google Service Account (do this once, takes ~5 minutes)

The app writes to Google Sheets on your behalf using a Service Account —
no browser login needed after this setup.

1. Go to [console.cloud.google.com](https://console.cloud.google.com) and sign in.
2. Click **Select a project** → **New Project** → give it any name → **Create**.
3. In the left menu go to **APIs & Services → Library**.
4. Search for **Google Sheets API** → click it → **Enable**.
5. Search for **Google Drive API** → click it → **Enable**.
6. Go to **APIs & Services → Credentials** → **Create Credentials → Service Account**.
7. Give the service account any name → **Done** (no extra roles needed).
8. Click the service account you just created → **Keys** tab → **Add Key → Create new key → JSON → Create**.
9. A `credentials.json` file downloads automatically.
10. **Place `credentials.json` in the `vendor-tracker/` folder** (same level as `app.py`).

> ⚠️ Never commit `credentials.json` to Git. Add it to `.gitignore`.

---

### 3 — Place your vendor CSV

Copy your `List_of_Vendors.csv` into the `vendor-tracker/` folder.

The file must have these three columns (exact names):

| Column | What it contains |
|---|---|
| `Vendor Name` | Clean name shown in the Google Sheet output |
| `Description` | A sample transaction description (for your reference only) |
| `Key Word` | The token the app matches against each transaction's description |

**Important rules:**
- If the same `Key Word` appears on multiple rows (e.g. `Prestige` for three properties), the app automatically consolidates them — the keyword itself becomes the display name.
- Keywords are matched **case-insensitively** with fuzzy tolerance (threshold: 85/100), so minor variations in bank descriptions are handled automatically.
- To add a new vendor: add a new row to the CSV and restart the app (or click "Reload vendor list" in the sidebar).

---

## Running the app

```bash
cd vendor-tracker
streamlit run app.py
```

The app opens in your browser at `http://localhost:8501`.

---

## How to use it

1. **Sidebar** — enter your Google account email. The new Sheet will be shared with you.
2. **Step 1** — drag and drop your monthly XLS bank statement.
3. **Step 2** — pick which sheet to process from the dropdown (one sheet = one month).
4. **Step 3** — type the Google Sheet file name and the tab name.
5. Click **Process & Push to Google Sheets**.

The app will:
- Show you a preview of matched and unclassified transactions before pushing.
- Create a **brand-new** Google Sheet with three tabs:
  - `[Your tab name]` — every matched debit with Vendor Name, Date, Txn ID, Amount, Description
  - `Unclassified` — debits that didn't match any keyword (review these to update your CSV)
  - `Pivot` — one row per vendor: Total Amount and No. of Transactions

> Running the same month twice creates a second separate sheet — nothing is overwritten.

---

## Updating the vendor keyword map

Open `List_of_Vendors.csv` in VS Code or Excel, add/edit rows, save.
Back in the app, click **🔄 Reload vendor list** in the sidebar — no restart needed.

---

## Troubleshooting

| Error | Fix |
|---|---|
| `credentials.json not found` | Place the file next to `app.py` |
| `List_of_Vendors.csv not found` | Place the file next to `app.py` |
| `INVALID_ARGUMENT` from Google | Check that Sheets API and Drive API are both enabled in your Google Cloud project |
| `403 Forbidden` from Google | The service account may not have Drive access — confirm both APIs are enabled |
| Transaction missing from output | It may be a CR entry (filtered out) or unclassified — check the Unclassified tab |
| Vendor matched incorrectly | Lower the fuzzy threshold in `app.py` (`FUZZY_THRESHOLD = 85`) or make the keyword more specific in the CSV |
