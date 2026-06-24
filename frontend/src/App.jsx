import { useState, useRef } from 'react'

const API = '/api' // proxied to FastAPI in dev (see vite.config.js)

// Sunday-to-Saturday default isn't assumed; the user picks both dates.
function todayISO(offset = 0) {
  const d = new Date()
  d.setDate(d.getDate() + offset)
  return d.toISOString().slice(0, 10)
}

function FileDrop({ label, hint, accept, file, onFile }) {
  const ref = useRef(null)
  return (
    <div
      className={'drop' + (file ? ' filled' : '')}
      onClick={() => ref.current?.click()}
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => {
        e.preventDefault()
        if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0])
      }}
    >
      <div className="name">{label}</div>
      <div className="hint">{hint}</div>
      {file && <div className="file">✓ {file.name}</div>}
      <input
        ref={ref}
        type="file"
        accept={accept}
        hidden
        onChange={(e) => e.target.files[0] && onFile(e.target.files[0])}
      />
    </div>
  )
}

function Field({ label, note, ...props }) {
  return (
    <div className="field">
      <label>{label}</label>
      <input {...props} />
      {note && <div className="note">{note}</div>}
    </div>
  )
}

const money = (n) =>
  '₺ ' + Number(n).toLocaleString(undefined, { maximumFractionDigits: 2 })

export default function App() {
  const [workbook, setWorkbook] = useState(null)
  const [template, setTemplate] = useState(null)
  const [start, setStart] = useState(todayISO(-6))
  const [end, setEnd] = useState(todayISO(0))
  const [rates, setRates] = useState({ naira: '32', usdt: '43', usd: '44.5', eur: '', gbp: '' })
  const [preparedBy, setPreparedBy] = useState('FESTUS')
  // Section 8 deductions are manual weekly judgement entries — not in the sheet.
  const [deductions, setDeductions] = useState([
    { label: 'Cash Expense Payments', lagos: '', abuja: '' },
    { label: 'Chicken', lagos: '', abuja: '' },
    { label: 'Greep', lagos: '', abuja: '' },
    { label: 'Advance for Jadesola', lagos: '', abuja: '' },
    { label: 'Other Expenses', lagos: '', abuja: '' },
  ])
  const [debtPayable, setDebtPayable] = useState('')
  const [cashPayments, setCashPayments] = useState('')
  const [notesLagos, setNotesLagos] = useState('NIL')
  const [notesAbuja, setNotesAbuja] = useState('NIL')
  const [status, setStatus] = useState({ kind: '', msg: '' })
  const [summary, setSummary] = useState(null)
  const [busy, setBusy] = useState(false)

  const ready = workbook && template && start && end

  function setDed(i, key, val) {
    setDeductions((ds) => ds.map((d, idx) => (idx === i ? { ...d, [key]: val } : d)))
  }
  function addDed() {
    setDeductions((ds) => [...ds, { label: '', lagos: '', abuja: '' }])
  }
  function removeDed(i) {
    setDeductions((ds) => ds.filter((_, idx) => idx !== i))
  }

  function buildForm() {
    const fd = new FormData()
    fd.append('workbook', workbook)
    fd.append('template', template)
    fd.append('start', start)
    fd.append('end', end)
    fd.append('naira_rate', rates.naira || '32')
    fd.append('usdt_rate', rates.usdt || '43')
    fd.append('usd_rate', rates.usd || '44.5')
    fd.append('eur_rate', rates.eur || '0')
    fd.append('gbp_rate', rates.gbp || '0')
    fd.append('prepared_by', preparedBy)
    fd.append('exchange_rate_text', `${rates.naira || 32} = ₺1`)
    fd.append(
      'section8',
      JSON.stringify({
        deductions: deductions.filter((d) => d.label.trim()),
        debt_payable: debtPayable.trim(),
        cash_payments_made: cashPayments.trim(),
        notes_lagos: notesLagos.trim(),
        notes_abuja: notesAbuja.trim(),
      })
    )
    return fd
  }

  async function preview() {
    setBusy(true); setSummary(null)
    setStatus({ kind: 'busy', msg: 'Reading workbook…' })
    try {
      const r = await fetch(`${API}/preview`, { method: 'POST', body: buildForm() })
      const j = await r.json()
      if (!j.ok) throw new Error(j.error || 'Failed')
      setSummary(j.summary)
      setStatus({ kind: 'ok', msg: `Read ${j.summary.days_found} days. Review totals below.` })
    } catch (e) {
      setStatus({ kind: 'err', msg: e.message })
    } finally {
      setBusy(false)
    }
  }

  async function generate() {
    setBusy(true)
    setStatus({ kind: 'busy', msg: 'Filling report…' })
    try {
      const r = await fetch(`${API}/generate`, { method: 'POST', body: buildForm() })
      if (!r.ok) {
        const j = await r.json().catch(() => ({}))
        throw new Error(j.error || `Server error ${r.status}`)
      }
      const blob = await r.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `YOOWA_Report_${start}_to_${end}.docx`
      a.click()
      URL.revokeObjectURL(url)
      setStatus({ kind: 'ok', msg: 'Report downloaded.' })
    } catch (e) {
      setStatus({ kind: 'err', msg: e.message })
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="shell">
      <div className="masthead">
        <h1>YOOWA <span>Weekly Report</span></h1>
        <div className="tag">Lefkosa · Magusa · ledger console</div>
      </div>

      <div className="sheet">
        {/* Files */}
        <div className="block">
          <p className="step-label">Source files</p>
          <div className="drops">
            <FileDrop
              label="Lagos 2025 workbook"
              hint="The .xlsx exported from the Google Sheet"
              accept=".xlsx"
              file={workbook}
              onFile={setWorkbook}
            />
            <FileDrop
              label="Report template"
              hint="The blank Weekly Accounting Report .docx"
              accept=".docx"
              file={template}
              onFile={setTemplate}
            />
          </div>
        </div>

        {/* Week */}
        <div className="block">
          <p className="step-label">Reporting week</p>
          <div className="grid two">
            <Field label="Week start" type="date" value={start}
              onChange={(e) => setStart(e.target.value)} />
            <Field label="Week end" type="date" value={end}
              onChange={(e) => setEnd(e.target.value)} />
          </div>
        </div>

        {/* Rates */}
        <div className="block">
          <p className="step-label">Conversion rates · confirm each week</p>
          <div className="grid rates">
            <Field label="Naira ÷" value={rates.naira} note="32 = ₺1"
              onChange={(e) => setRates({ ...rates, naira: e.target.value })} />
            <Field label="USDT ×" value={rates.usdt} note="→ TL"
              onChange={(e) => setRates({ ...rates, usdt: e.target.value })} />
            <Field label="USD ×" value={rates.usd} note="→ TL"
              onChange={(e) => setRates({ ...rates, usd: e.target.value })} />
            <Field label="EUR ×" value={rates.eur} note="optional"
              onChange={(e) => setRates({ ...rates, eur: e.target.value })} />
            <Field label="GBP ×" value={rates.gbp} note="optional"
              onChange={(e) => setRates({ ...rates, gbp: e.target.value })} />
          </div>
        </div>

        <div className="grid two">
          <Field label="Prepared by" value={preparedBy}
            onChange={(e) => setPreparedBy(e.target.value)} />
        </div>

        {/* Section 8 — manual deductions */}
        <div className="divider" />
        <div className="block">
          <p className="step-label">Section 8 · net revenue deductions</p>
          <p className="s8-help">
            These aren't in the spreadsheet — they're your weekly entries
            (Greep, advances, picked cash payments). Gross and Net are
            calculated for you; leave a row blank to skip it.
          </p>
          <div className="ded-head">
            <span>Deduction</span><span>Lagos (₺)</span><span>Abuja (₺)</span><span />
          </div>
          {deductions.map((d, i) => (
            <div className="ded-row" key={i}>
              <input value={d.label} placeholder="Label"
                onChange={(e) => setDed(i, 'label', e.target.value)} />
              <input value={d.lagos} placeholder="0" inputMode="decimal"
                onChange={(e) => setDed(i, 'lagos', e.target.value)} />
              <input value={d.abuja} placeholder="—" inputMode="decimal"
                onChange={(e) => setDed(i, 'abuja', e.target.value)} />
              <button type="button" className="ded-x" onClick={() => removeDed(i)}>×</button>
            </div>
          ))}
          <button type="button" className="ded-add" onClick={addDed}>+ Add deduction</button>

          <div className="grid two" style={{ marginTop: 16 }}>
            <Field label="Outstanding debt — payable to suppliers (₺)"
              value={debtPayable} placeholder="e.g. 96,610"
              onChange={(e) => setDebtPayable(e.target.value)} />
            <Field label="Cash payments made this week (₺)"
              value={cashPayments} placeholder="e.g. 66,503"
              onChange={(e) => setCashPayments(e.target.value)} />
          </div>
          <div className="grid two" style={{ marginTop: 14 }}>
            <Field label="Notes — Lagos (Lefkosa)" value={notesLagos}
              onChange={(e) => setNotesLagos(e.target.value)} />
            <Field label="Notes — Abuja (Magusa)" value={notesAbuja}
              onChange={(e) => setNotesAbuja(e.target.value)} />
          </div>
        </div>

        <div className="divider" />

        {/* Actions */}
        <div className="actions">
          <button className="btn-preview" disabled={!ready || busy} onClick={preview}>
            Preview totals
          </button>
          <button className="btn-generate" disabled={!ready || busy} onClick={generate}>
            Generate &amp; download report
          </button>
          {status.msg && (
            <span className={'status ' + status.kind}>{status.msg}</span>
          )}
        </div>

        {/* Summary */}
        {summary && (
          <div className="summary">
            <h3>Preview — week revenue & expenses</h3>
            <table>
              <tbody>
                <tr><td className="k">Gross revenue — Lefkosa</td><td className="v">{money(summary.gross_lagos)}</td></tr>
                <tr><td className="k">Gross revenue — Magusa</td><td className="v">{money(summary.gross_abuja)}</td></tr>
                <tr><td className="k">Expenses (Section 7) — Lefkosa</td><td className="v">{money(summary.expenses_lagos)}</td></tr>
                <tr><td className="k">Expenses (Section 7) — Magusa</td><td className="v">{money(summary.expenses_abuja)}</td></tr>
                <tr><td className="k">Section 8 deductions — Lefkosa</td><td className="v">{money(summary.deductions_lagos)}</td></tr>
                <tr><td className="k">Net remaining — Lefkosa</td><td className="v">{money(summary.net_lagos)}</td></tr>
                <tr><td className="k">Net remaining — Magusa</td><td className="v">{money(summary.net_abuja)}</td></tr>
                <tr className="net"><td className="k">Combined net revenue</td><td className="v">{money(summary.net)}</td></tr>
              </tbody>
            </table>
          </div>
        )}
      </div>

      <div className="foot">
        Reads HAMITKOY SALES → Lefkosa · MAGUSA SALES → Magusa · EXPENSE SHEET by item.
        Nothing leaves your machine.
      </div>
    </div>
  )
}
