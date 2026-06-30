"""
FastAPI server for the YOOWA Weekly Report generator.

POST /generate
  multipart form:
    workbook : the Lagos 2025 .xlsx
    template : the blank Weekly Accounting Report .docx
    start    : YYYY-MM-DD
    end      : YYYY-MM-DD
    naira_rate, usdt_rate, usd_rate, eur_rate, gbp_rate : floats
    prepared_by, exchange_rate_text, date_prepared : strings
  -> returns the filled .docx

POST /preview  -> same inputs, returns JSON summary (no file) for the UI to show
                  before downloading.
"""

import datetime
import io
import json

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse

import report_engine as R

app = FastAPI(title="Reporta API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def _parse_date(s: str) -> datetime.date:
    return datetime.datetime.strptime(s.strip(), "%Y-%m-%d").date()


async def _run(workbook, template, start, end, form):
    xlsx_bytes = await workbook.read()
    tpl_bytes = await template.read()
    rates = {
        "naira": float(form.get("naira_rate", 32) or 32),
        "usdt": float(form.get("usdt_rate", 43) or 43),
        "usd": float(form.get("usd_rate", 44.5) or 44.5),
        "eur": float(form.get("eur_rate", 0) or 0),
        "gbp": float(form.get("gbp_rate", 0) or 0),
    }
    header = {
        "prepared_by": form.get("prepared_by", "FESTUS"),
        "exchange_rate_text": form.get(
            "exchange_rate_text", f"{int(rates['naira'])} = ₺1"
        ),
        "date_prepared": form.get(
            "date_prepared", datetime.date.today().strftime("%d/%m/%Y")
        ),
    }
    section8 = {}
    raw = form.get("section8", "")
    if raw:
        try:
            section8 = json.loads(raw)
        except Exception:
            section8 = {}
    return R.build_report(
        xlsx_bytes, tpl_bytes,
        start=_parse_date(start), end=_parse_date(end),
        rates=rates, header=header, section8=section8,
    )


@app.post("/preview")
async def preview(
    workbook: UploadFile = File(...),
    template: UploadFile = File(...),
    start: str = Form(...),
    end: str = Form(...),
    naira_rate: str = Form("32"),
    usdt_rate: str = Form("43"),
    usd_rate: str = Form("44.5"),
    eur_rate: str = Form("0"),
    gbp_rate: str = Form("0"),
    prepared_by: str = Form("FESTUS"),
    exchange_rate_text: str = Form(""),
    date_prepared: str = Form(""),
    section8: str = Form(""),
):
    form = dict(
        naira_rate=naira_rate, usdt_rate=usdt_rate, usd_rate=usd_rate,
        eur_rate=eur_rate, gbp_rate=gbp_rate, prepared_by=prepared_by,
        exchange_rate_text=exchange_rate_text, date_prepared=date_prepared,
        section8=section8,
    )
    try:
        _, meta = await _run(workbook, template, start, end, form)
        meta = {k: (round(v, 2) if isinstance(v, float) else v) for k, v in meta.items()}
        return JSONResponse({"ok": True, "summary": meta})
    except Exception as e:  # surface a readable error to the UI
        return JSONResponse({"ok": False, "error": str(e)}, status_code=400)


@app.post("/generate")
async def generate(
    workbook: UploadFile = File(...),
    template: UploadFile = File(...),
    start: str = Form(...),
    end: str = Form(...),
    naira_rate: str = Form("32"),
    usdt_rate: str = Form("43"),
    usd_rate: str = Form("44.5"),
    eur_rate: str = Form("0"),
    gbp_rate: str = Form("0"),
    prepared_by: str = Form("FESTUS"),
    exchange_rate_text: str = Form(""),
    date_prepared: str = Form(""),
    section8: str = Form(""),
):
    form = dict(
        naira_rate=naira_rate, usdt_rate=usdt_rate, usd_rate=usd_rate,
        eur_rate=eur_rate, gbp_rate=gbp_rate, prepared_by=prepared_by,
        exchange_rate_text=exchange_rate_text, date_prepared=date_prepared,
        section8=section8,
    )
    doc_bytes, _ = await _run(workbook, template, start, end, form)
    fname = f"YOOWA_Report_{start}_to_{end}.docx"
    return StreamingResponse(
        io.BytesIO(doc_bytes),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


@app.get("/")
def root():
    return {"Reporta API running"}
