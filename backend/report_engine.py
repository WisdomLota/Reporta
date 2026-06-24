"""
YOOWA Weekly Report Engine
--------------------------
Reads the 'Lagos 2025' workbook and fills the Weekly Accounting Report .docx.

Confirmed mappings (validated against the 08-14 June report):
  Lagos (Lefkosa)  -> sheet 'HAMITKOY SALES'
  Abuja (Magusa)   -> sheet 'MAGUSA SALES'
  Sales tabs: column B = date. Money columns located by header text on row 2:
      NAIRA, CASH, POS + IBAN, USDT, REDOT, USD
  EXPENSE SHEET: row 1 = dates; each date spans 2 sub-columns
      (row 3 sub-header) HAMITKOY | MAGUSA. Items in column A, category col B.

Anything uncertain (exchange rate, USDT/USD/foreign rates) is passed in as
parameters from the UI so the user confirms once, not guessed by the code.
"""

import datetime
import io
import re
from copy import deepcopy

import openpyxl
from docx import Document


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def _as_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    return None


def _num(v):
    """Coerce a cell into a float, treating blanks/None/text as 0."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if s in ("", "-"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _money(v):
    """Format a number the way the report does: ₺ 1,234.56 (trim trailing .00)."""
    v = round(float(v), 2)
    if v == int(v):
        return f"₺ {int(v):,}"
    return f"₺ {v:,.2f}"


def _find_money_columns(ws):
    """Locate money columns by scanning row 2 (and row 3 fallback) header text."""
    wanted = {
        "NAIRA": ["NAIRA"],
        "CASH": ["CASH"],
        "POS": ["POS + IBAN", "POS+IBAN", "POS"],
        "USDT": ["USDT"],
        "REDOT": ["REDOT"],
        "USD": ["USD"],
    }
    found = {}
    for col in range(1, ws.max_column + 1):
        for hdr_row in (2, 3):
            raw = ws.cell(hdr_row, col).value
            if not raw:
                continue
            txt = str(raw).strip().upper()
            for key, variants in wanted.items():
                if key in found:
                    continue
                if any(txt == v.upper() for v in variants):
                    found[key] = col
    return found


def _table_by_header(doc, header_text):
    """Find the first table whose top-left or any first-column cell contains header_text."""
    for t in doc.tables:
        if not t.rows:
            continue
        if header_text in t.cell(0, 0).text:
            return t
        for r in range(len(t.rows)):
            if header_text in t.cell(r, 0).text:
                return t
    return None


def _daily_rows(ws, start, end):
    """Return list of (date, row_index) for dates within [start, end], sorted."""
    out = []
    for row in range(1, ws.max_row + 1):
        d = _as_date(ws.cell(row, 2).value)
        if d and start <= d <= end:
            out.append((d, row))
    out.sort(key=lambda x: x[0])
    return out


# ----------------------------------------------------------------------------
# Sales extraction
# ----------------------------------------------------------------------------
def extract_sales(ws, start, end):
    """Return a dict of 7-day lists for naira, cash, pos, usdt, redot, usd."""
    cols = _find_money_columns(ws)
    rows = _daily_rows(ws, start, end)
    data = {k: [] for k in ["naira", "cash", "pos", "usdt", "redot", "usd"]}
    for _, r in rows:
        data["naira"].append(_num(ws.cell(r, cols.get("NAIRA", 0)).value) if cols.get("NAIRA") else 0.0)
        data["cash"].append(_num(ws.cell(r, cols.get("CASH", 0)).value) if cols.get("CASH") else 0.0)
        data["pos"].append(_num(ws.cell(r, cols.get("POS", 0)).value) if cols.get("POS") else 0.0)
        data["usdt"].append(_num(ws.cell(r, cols.get("USDT", 0)).value) if cols.get("USDT") else 0.0)
        data["redot"].append(_num(ws.cell(r, cols.get("REDOT", 0)).value) if cols.get("REDOT") else 0.0)
        data["usd"].append(_num(ws.cell(r, cols.get("USD", 0)).value) if cols.get("USD") else 0.0)
    # pad to 7 days
    for k in data:
        while len(data[k]) < 7:
            data[k].append(0.0)
    data["_dates"] = [d for d, _ in rows]
    data["_columns_found"] = cols
    return data


# ----------------------------------------------------------------------------
# Expense extraction
# ----------------------------------------------------------------------------
# Map report item label -> spreadsheet item label (when they differ)
EXPENSE_ALIAS = {
    "Goat": "Goat Meat",
    "Mixed Vegetable": "Mix Vegetable",
    "Sweet Pepper Paste": "Sweet pepper paste",
    "Aluminium foil": "Aluminium Foil",
}


def extract_expenses(ws, start, end, report_items):
    """
    For each report expense item, sum HAMITKOY (Lagos) and MAGUSA (Abuja)
    over the week. Returns {item: (lagos_total, abuja_total)}.
    """
    # locate date columns + their HAMITKOY/MAGUSA sub-columns
    day_cols = []  # (date, hamit_col, magusa_col)
    for col in range(1, ws.max_column + 1):
        d = _as_date(ws.cell(1, col).value)
        if d and start <= d <= end:
            sub_a = str(ws.cell(3, col).value or "").strip().upper()
            sub_b = str(ws.cell(3, col + 1).value or "").strip().upper()
            hamit = col if sub_a == "HAMITKOY" else col
            magusa = col + 1 if sub_b == "MAGUSA" else col + 1
            day_cols.append((d, hamit, magusa))

    # index spreadsheet items by name (column A), rows 6..max
    sheet_items = {}
    for r in range(6, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if name and str(name).strip():
            sheet_items[str(name).strip().lower()] = r

    results = {}
    for item in report_items:
        sheet_name = EXPENSE_ALIAS.get(item, item)
        r = sheet_items.get(sheet_name.strip().lower())
        lagos = abuja = 0.0
        if r:
            for _, hc, mc in day_cols:
                lagos += _num(ws.cell(r, hc).value)
                abuja += _num(ws.cell(r, mc).value)
        results[item] = (lagos, abuja)
    return results


# ----------------------------------------------------------------------------
# Report builder
# ----------------------------------------------------------------------------
def build_report(xlsx_bytes, template_bytes, *, start, end, rates, header, section8=None):
    """
    xlsx_bytes / template_bytes : raw file bytes
    start, end : datetime.date  (7-day window)
    rates : dict -> naira (e.g. 32 means 32 naira = 1 TL), usdt, usd, eur, gbp
    header : dict -> prepared_by, date_prepared, exchange_rate_text
    Returns filled .docx as bytes.
    """
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    doc = Document(io.BytesIO(template_bytes))

    naira_rate = float(rates.get("naira", 32)) or 32.0

    lagos = extract_sales(wb["HAMITKOY SALES"], start, end)
    abuja = extract_sales(wb["MAGUSA SALES"], start, end)

    # ---- Section header (table 0) ----
    t = doc.tables[0]
    t.cell(0, 1).text = f"{start.strftime('%d/%m/%Y')} to {end.strftime('%d/%m/%Y')}"
    t.cell(0, 3).text = header.get("prepared_by", "FESTUS")
    t.cell(1, 1).text = header.get("exchange_rate_text", f"{int(naira_rate)} = ₺1")
    t.cell(1, 3).text = header.get("date_prepared", datetime.date.today().strftime("%d/%m/%Y"))

    # ---- Section 1: Naira -> TL (tables 1=Lagos, 2=Abuja) ----
    def fill_naira(table, naira_list):
        total_n = total_tl = 0.0
        for i in range(7):
            n = naira_list[i]
            tl = n / naira_rate
            table.cell(i + 1, 1).text = f"{int(round(n)):,}" if n else "0"
            table.cell(i + 1, 2).text = _money(tl)
            total_n += n
            total_tl += tl
        table.cell(8, 1).text = f"{int(round(total_n)):,}"
        table.cell(8, 2).text = _money(total_tl)
        return total_tl

    lagos_naira_tl = fill_naira(doc.tables[1], lagos["naira"])
    abuja_naira_tl = fill_naira(doc.tables[2], abuja["naira"])

    # ---- Section 2: Cash (tables 3=Lagos, 4=Abuja) ----
    def fill_cash(table, cash_list):
        total = 0.0
        for i in range(7):
            c = cash_list[i]
            table.cell(i + 1, 1).text = _money(c)
            total += c
        table.cell(8, 1).text = _money(total)
        return total

    lagos_cash = fill_cash(doc.tables[3], lagos["cash"])
    abuja_cash = fill_cash(doc.tables[4], abuja["cash"])

    # ---- Section 3: POS+IBAN (tables 5=Lagos, 6=Abuja) ----
    def fill_pos(table, pos_list):
        total = 0.0
        for i in range(7):
            p = pos_list[i]
            table.cell(i + 1, 1).text = _money(p)
            total += p
        table.cell(8, 1).text = _money(total)
        return total

    lagos_pos = fill_pos(doc.tables[5], lagos["pos"])
    abuja_pos = fill_pos(doc.tables[6], abuja["pos"])

    # ---- Section 4: USDT + Redot (table 7) ----
    usdt_rate = float(rates.get("usdt", 43))
    lagos_usdt = sum(lagos["usdt"])
    abuja_usdt = sum(abuja["usdt"])
    lagos_redot = sum(lagos["redot"])
    abuja_redot = sum(abuja["redot"])
    t7 = doc.tables[7]
    t7.cell(1, 1).text = f"${lagos_usdt:,.2f}"
    t7.cell(1, 2).text = f"${abuja_usdt:,.2f}"
    t7.cell(1, 3).text = f"${lagos_usdt + abuja_usdt:,.2f}"
    t7.cell(3, 1).text = f"${lagos_redot:,.2f}"
    t7.cell(3, 2).text = f"${abuja_redot:,.2f}"
    t7.cell(3, 3).text = f"${lagos_redot + abuja_redot:,.2f}"
    usdt_tl_l = lagos_usdt * usdt_rate
    usdt_tl_a = abuja_usdt * usdt_rate
    redot_tl_l = lagos_redot * usdt_rate
    redot_tl_a = abuja_redot * usdt_rate

    # ---- Section 5: Foreign currencies (table 8) just records rates given ----
    usd_rate = float(rates.get("usd", 44.5))
    lagos_usd = sum(lagos["usd"])
    abuja_usd = sum(abuja["usd"])
    foreign_tl_l = lagos_usd * usd_rate
    foreign_tl_a = abuja_usd * usd_rate

    # ---- Section 6: Revenue summary (table 9) ----
    t9 = doc.tables[9]

    def row(idx, l, a):
        t9.cell(idx, 1).text = _money(l)
        t9.cell(idx, 2).text = _money(a)
        t9.cell(idx, 3).text = _money(l + a)

    row(1, lagos_naira_tl, abuja_naira_tl)
    row(2, lagos_cash, abuja_cash)
    row(3, lagos_pos, abuja_pos)
    row(4, usdt_tl_l, usdt_tl_a)
    row(5, redot_tl_l, redot_tl_a)
    row(6, foreign_tl_l, foreign_tl_a)
    gross_l = lagos_naira_tl + lagos_cash + lagos_pos + usdt_tl_l + redot_tl_l + foreign_tl_l
    gross_a = abuja_naira_tl + abuja_cash + abuja_pos + usdt_tl_a + redot_tl_a + foreign_tl_a
    row(7, gross_l, gross_a)

    # ---- Section 7: Expenses (table 10) ----
    t10 = doc.tables[10]
    report_items = [t10.cell(r, 0).text.strip() for r in range(1, len(t10.rows) - 1)]
    exp = extract_expenses(wb["EXPENSE SHEET"], start, end, report_items)
    tot_l = tot_a = 0.0
    for i, item in enumerate(report_items, start=1):
        l, a = exp.get(item, (0.0, 0.0))
        t10.cell(i, 1).text = _money(l)
        t10.cell(i, 2).text = _money(a)
        tot_l += l
        tot_a += a
    last = len(t10.rows) - 1
    t10.cell(last, 1).text = _money(tot_l)
    t10.cell(last, 2).text = _money(tot_a)

    # ---- Section 8: Net revenue calculation ----
    # These deductions are NOT in the spreadsheet — they are manual, weekly
    # judgement entries (Greep, advances, picked cash payments). They arrive
    # from the UI as `section8`. Gross and NET are computed; deductions are
    # whatever the user entered. If none are given, Section 8 falls back to a
    # simple gross-minus-expenses net so the report is never left blank.
    section8 = section8 or {}
    deductions = section8.get("deductions", [])  # [{label, lagos, abuja}]
    debt_payable = section8.get("debt_payable", "")        # e.g. "96,610"
    cash_payments_made = section8.get("cash_payments_made", "")  # e.g. "66,503"

    # Detailed net-calc table (the 8-row table: Gross / Less:… / NET)
    t12 = _table_by_header(doc, "Gross Total Revenue")
    if t12 is not None:
        # row 1 = Gross
        t12.cell(1, 1).text = _money(gross_l)
        t12.cell(1, 2).text = _money(gross_a)
        less_l = less_a = 0.0
        # rows 2..(n-1) are the Less: lines. Match by trailing label text if
        # present in template, else write the user's deductions in order.
        body_rows = list(range(2, len(t12.rows) - 1))
        for ridx, ded in zip(body_rows, deductions):
            label = ded.get("label", "").strip()
            l = _num(ded.get("lagos", 0))
            a = _num(ded.get("abuja", 0))
            t12.cell(ridx, 0).text = f"Less: {label}" if label else t12.cell(ridx, 0).text
            t12.cell(ridx, 1).text = _money(l) if (l or ded.get("lagos") not in (None, "")) else "₺ "
            t12.cell(ridx, 2).text = _money(a) if (a or ded.get("abuja") not in (None, "")) else ""
            less_l += l
            less_a += a
        # blank any leftover template Less: rows the user didn't fill
        for ridx in body_rows[len(deductions):]:
            t12.cell(ridx, 1).text = "₺ "
            t12.cell(ridx, 2).text = ""
        net_l = gross_l - less_l
        net_a = gross_a - less_a
        last12 = len(t12.rows) - 1
        t12.cell(last12, 1).text = _money(net_l)
        t12.cell(last12, 2).text = _money(net_a)
        combined_net = net_l + net_a
    else:
        # fallback
        combined_net = (gross_l + gross_a) - (tot_l + tot_a)
        net_l = gross_l - tot_l
        net_a = gross_a - tot_a
        less_l = less_a = 0.0

    # Small 2-row table: Outstanding Debt / Cash Payments Made This Week
    t11 = _table_by_header(doc, "Outstanding Debt (Payable to Suppliers)")
    if t11 is not None:
        if debt_payable:
            t11.cell(0, 1).text = f"₺ {debt_payable}"
        if cash_payments_made:
            t11.cell(1, 1).text = f"₺ {cash_payments_made}"

    # Outstanding-debt summary line (single-cell table 13) + combined paragraph
    if debt_payable:
        for t in doc.tables:
            if len(t.rows) and "Outstanding Debt" in t.cell(0, 0).text and len(t.columns) == 1:
                t.cell(0, 0).text = (
                    f"Outstanding Debt (Payable to Suppliers)"
                    f"                                               =       ₺ {debt_payable}"
                )
                break

    for p in doc.paragraphs:
        if "COMBINED NET REVENUE BALANCE" in p.text:
            txt = f"COMBINED NET REVENUE BALANCE\t\t\t\t\t=\t{_money(combined_net)}"
            if p.runs:
                p.runs[0].text = txt
                for r_ in p.runs[1:]:
                    r_.text = ""
            else:
                p.text = txt
            break

    # Section 9 notes (table 14) — optional
    notes = (section8.get("notes_lagos", ""), section8.get("notes_abuja", ""))
    if any(notes):
        for t in doc.tables:
            if len(t.rows) == 1 and "Notes:" in t.cell(0, 0).text:
                nl = notes[0] or "NIL"
                na = notes[1] or "NIL"
                t.cell(0, 0).text = (
                    f"Lagos (Lefkosa) Notes: {nl}\n \n \n \n"
                    f"Abuja (Magusa) Notes: {na}\n \n \n "
                )
                break

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.read(), {
        "gross_lagos": gross_l, "gross_abuja": gross_a,
        "expenses_lagos": tot_l, "expenses_abuja": tot_a,
        "deductions_lagos": less_l, "deductions_abuja": less_a,
        "net_lagos": net_l, "net_abuja": net_a,
        "net": combined_net,
        "lagos_columns": lagos["_columns_found"],
        "abuja_columns": abuja["_columns_found"],
        "days_found": len(lagos["_dates"]),
    }
