# YOOWA Weekly Report Generator

Turns the **Lagos 2025** workbook into the filled **Weekly Accounting Report**
`.docx` in well under 10 minutes. Reads the spreadsheet, applies conversions,
totals every section, and writes the report in the exact template format.

## What it maps (validated against the 08–14 June report)

| Report section | Source |
|---|---|
| Lagos (Lefkosa) revenue | sheet **HAMITKOY SALES** |
| Abuja (Magusa) revenue | sheet **MAGUSA SALES** |
| Naira / Cash / POS+IBAN / USDT / Redot / USD | columns located by header text |
| Expenses (Section 7) | **EXPENSE SHEET**, per item, HAMITKOY=Lagos / MAGUSA=Abuja |

Anything that changes weekly (exchange rate, USDT/USD rates) is entered in the
UI, so it's confirmed once per run rather than guessed.

---

## Requirements
- **Python 3.10+**
- **Node.js 18+**

## Setup

### 1. Backend (the parsing + .docx engine)
```bash
cd backend
python -m venv venv
# Windows:  venv\Scripts\activate
# macOS/Linux:  source venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```
Leave this terminal running. It serves the API at `http://localhost:8000`.

### 2. Frontend (the interface) — in a second terminal
```bash
cd frontend
npm install
npm run dev
```
Open the URL it prints (default `http://localhost:5173`).

---

## Using it (the weekly routine)
1. Download the Google Sheet as `.xlsx` (File → Download → Microsoft Excel).
2. In the app, drop in that **workbook** and the blank **report template** `.docx`
   (use `YOOWA_Weekly_Accounting_Report_Template.docx` — the complete one).
3. Pick the **week start/end** dates.
4. Confirm the **rates** (Naira `32 = ₺1`, USDT, USD…).
5. Fill **Section 8 deductions** — these are *not* in the spreadsheet; they're
   the weekly judgement entries (Greep, advances, picked cash payments). The app
   pre-loads the usual rows; type the amounts, add/remove rows as needed.
   Gross and the NET row are calculated automatically.
6. Click **Preview totals** — sanity-check revenue, deductions, and net.
7. Click **Generate & download report** → the filled `.docx` lands in Downloads.

## Why Section 8 is manual
Sections 1–7 come straight from the spreadsheet. Section 8's deductions
(Greep ₺38,750, Advance for Jadesola, "Other Expenses", the picked Cash
Expense Payment figure) are round, hand-decided numbers that appear nowhere in
the workbook — so the app can't compute them, it gives you fields to enter them
and then does the Gross − deductions = NET arithmetic for you.

## Testing it's correct
Run it for **08/06 to 14/06 2026** and compare against the already-submitted
report for that week. Naira (both sites), Magusa cash, and named expense items
(Fuel, Rice, Turkey…) should match. Small differences on a section usually mean
that week's submitted report used a hand-shifted day set — the app uses the true
calendar days, which is the more accurate version.

## Notes / next steps
- **All local** — files never leave the machine.
- The expense item-name differences (e.g. sheet "Goat Meat" vs report "Goat")
  are handled by an alias map in `backend/report_engine.py` → `EXPENSE_ALIAS`.
  Add to it if new items appear.
- **Phase 2 (optional):** swap the manual `.xlsx` upload for a live Google
  Sheets API pull, so step 1 disappears. The engine stays identical.
- Confirm with the boss's data owner: USDT/USD/foreign rates source, and whether
  the expense item list ever changes — both are the only soft spots.
