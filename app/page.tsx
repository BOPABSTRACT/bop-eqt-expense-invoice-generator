'use client'

import { useState, useRef } from 'react'

const LOGO = "https://i.imgur.com/szjzoxt.png"

const COMPANY_OPTIONS = [
  'BOP Abstract, LLC',
  'BOP Acquisition, LLC',
]

const MANAGER_OPTIONS = [
  'Andrew Restanio',
  'Bryan Hollihan',
  'Carrick Tuck',
  'Eric Strouth',
  'Fred Rousch',
  'J.J. Courie',
  'Kristina Hancock',
  'Kurt Stephens',
  'Mitchell Shwartz',
  'Sean Cotter',
  'Torey Sochaki',
  'Wesley Rosenbaugh',
]

const AFE_OPTIONS = [
  { label: 'EQT Greene Mineral Purchasing', afe: 'N34621.2401.1014' },
  { label: 'EQT Greene County Title/Curative AFE', afe: 'N34621.2401.1011' },
  { label: 'EQT Washington Mineral Purchasing AFE', afe: 'N34622.2401.1014' },
  { label: 'EQT Washington County Title/Curative AFE', afe: 'N34622.2401.1011' },
  { label: 'EQT Allegheny Mineral Purchasing AFE', afe: 'N34620.2401.1014' },
  { label: 'EQT Allegheny County Title/Curative AFE', afe: 'N34620.2401.1011' },
  { label: 'EQT Westmoreland Mineral Purchasing AFE', afe: 'N34623.2401.1014' },
  { label: 'EQT Westmoreland Title/Curative AFE', afe: 'N34623.2401.1011' },
  { label: 'EQT Monongalia WV Mineral Purchasing AFE', afe: 'N34627.2401.1014' },
  { label: 'EQT Monongalia WV Title/Curative AFE', afe: 'N34627.2401.1011' },
  { label: 'EQT Wetzel County WV Mineral Purchasing', afe: 'N34629.2401.1014' },
  { label: 'EQT Wetzel County WV Title/Curative AFE', afe: 'N34629.2401.1011' },
  { label: 'EQT Fayette Mineral Purchasing', afe: 'NTBD.2401.1014' },
  { label: 'EQT Fayette County Title/Curative AFE', afe: 'NTBD.2401.1011' },
  { label: 'EQT Belmont Mineral Purchasing', afe: 'N34619.2401.1014' },
  { label: 'EQT Belmont County Title/Curative AFE', afe: 'N34619.2401.1011' },
  { label: 'EQT Marion Mineral Purchasing', afe: 'N34625.2401.1014' },
  { label: 'EQT Marion County Title/Curative AFE', afe: 'N34625.2401.1011' },
  { label: 'EQT Marshall Mineral Purchasing', afe: 'N34626.2401.1014' },
  { label: 'EQT Marshall County Title/Curative AFE', afe: 'N34626.2401.1011' },
]

const COUNTY_OPTIONS = [
  'Allegheny',
  'Fayette',
  'Greene',
  'Lycoming',
  'Marion',
  'Monongalia',
  'Washington',
  'Wetzel',
  'Westmoreland',
]

export default function Home() {
  const [authenticated, setAuthenticated] = useState(false)
  const [passwordInput, setPasswordInput] = useState('')
  const [passwordError, setPasswordError] = useState(false)
  const [excelFiles, setExcelFiles] = useState<File[]>([])
  const [receiptFiles, setReceiptFiles] = useState<File[]>([])
  const [companyName, setCompanyName] = useState('')
  const [manager, setManager] = useState('')
  const [afeKey, setAfeKey] = useState('')
  const [county, setCounty] = useState('')
  const [invoiceDate, setInvoiceDate] = useState('')
  const [status, setStatus] = useState<'idle' | 'loading' | 'done' | 'error'>('idle')
  const [message, setMessage] = useState('')
  const excelRef = useRef<HTMLInputElement>(null)
  const receiptRef = useRef<HTMLInputElement>(null)

  const handlePasswordSubmit = () => {
    if (passwordInput === 'BOP2026') {
      setAuthenticated(true)
      setPasswordError(false)
    } else {
      setPasswordError(true)
    }
  }

  const handleGenerate = async () => {
    if (excelFiles.length === 0) {
      setMessage('Please upload at least one Excel file.')
      setStatus('error')
      return
    }
    if (!companyName) {
      setMessage('Please select a Company Name.')
      setStatus('error')
      return
    }
    if (!invoiceDate.trim()) {
      setMessage('Please enter an invoice date.')
      setStatus('error')
      return
    }
    setStatus('loading')
    setMessage('Generating invoices...')

    // Find the AFE number from the selected key
    const selectedAfe = AFE_OPTIONS.find(o => o.afe === afeKey)
    const afeNumber = selectedAfe?.afe || ''

    const formData = new FormData()
    excelFiles.forEach(f => formData.append('excel', f))
    receiptFiles.forEach(f => formData.append('receipts', f))
    formData.append('companyName', companyName)
    formData.append('manager', manager)
    formData.append('afe', afeNumber)
    formData.append('county', county)
    formData.append('invoiceDate', invoiceDate.trim())

    try {
      const res = await fetch('/api/generate', { method: 'POST', body: formData })
      if (!res.ok) {
        const err = await res.json()
        throw new Error(err.error || 'Generation failed')
      }
      const blob = await res.blob()
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = `eqt-expense-invoices-${Date.now()}.zip`
      a.click()
      URL.revokeObjectURL(url)
      setStatus('done')
      setMessage(`✅ ${excelFiles.length} invoice${excelFiles.length !== 1 ? 's' : ''} generated and downloaded!`)
    } catch (err: unknown) {
      setStatus('error')
      setMessage(err instanceof Error ? err.message : 'Something went wrong.')
    }
  }

  if (!authenticated) {
    return (
      <main style={{
        minHeight: '100vh', background: '#0f1117', fontFamily: "'Georgia', serif",
        color: '#e8e0d0', display: 'flex', alignItems: 'center', justifyContent: 'center',
      }}>
        <div style={{
          background: '#0d0f14', border: '1px solid #2a2a3a', borderRadius: 12,
          padding: '48px 40px', width: '100%', maxWidth: 400, textAlign: 'center',
        }}>
          <img src={LOGO} alt="BOP Logo"
            style={{ width: 140, height: 140, objectFit: 'contain', margin: '0 auto 24px', display: 'block' }} />
          <div style={{ fontSize: 20, fontWeight: 600, color: '#c8a96e', marginBottom: 4 }}>
            EQT - Expense Invoice Generator
          </div>
          <div style={{ fontSize: 12, color: '#666', letterSpacing: '0.12em', textTransform: 'uppercase', marginBottom: 32 }}>
            BOP Abstract / BOP Acquisition
          </div>
          <input
            type="password" placeholder="Enter password" value={passwordInput}
            onChange={e => { setPasswordInput(e.target.value); setPasswordError(false) }}
            onKeyDown={e => e.key === 'Enter' && handlePasswordSubmit()}
            style={{
              width: '100%', padding: '12px 16px', background: '#0f1117',
              border: `1px solid ${passwordError ? '#8b2020' : '#2a2a3a'}`,
              borderRadius: 6, color: '#e8e0d0', fontSize: 15,
              fontFamily: "'Georgia', serif", boxSizing: 'border-box', marginBottom: 12, outline: 'none',
            }}
          />
          {passwordError && (
            <div style={{ color: '#e07070', fontSize: 13, marginBottom: 12 }}>
              Incorrect password. Please try again.
            </div>
          )}
          <button onClick={handlePasswordSubmit} style={{
            width: '100%', padding: '12px 32px',
            background: 'linear-gradient(135deg, #c8a96e, #8b6914)',
            color: '#fff', border: 'none', borderRadius: 6, fontSize: 15,
            fontFamily: "'Georgia', serif", cursor: 'pointer', letterSpacing: '0.04em',
          }}>Enter</button>
        </div>
      </main>
    )
  }

  const selectStyle = {
    width: '100%', padding: '11px 16px', background: '#0d0f14',
    border: '1px solid #2a2a3a', borderRadius: 6, color: '#e8e0d0', fontSize: 14,
    fontFamily: "'Georgia', serif", boxSizing: 'border-box' as const, outline: 'none',
    cursor: 'pointer', appearance: 'none' as const,
    backgroundImage: `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%23c8a96e' d='M6 8L1 3h10z'/%3E%3C/svg%3E")`,
    backgroundRepeat: 'no-repeat' as const,
    backgroundPosition: 'right 14px center',
    paddingRight: '36px',
  }

  const inputStyle = {
    width: '100%', padding: '11px 16px', background: '#0d0f14',
    border: '1px solid #2a2a3a', borderRadius: 6, color: '#e8e0d0', fontSize: 14,
    fontFamily: "'Georgia', serif", boxSizing: 'border-box' as const, outline: 'none',
  }

  const labelStyle = { fontSize: 13, color: '#c8a96e', marginBottom: 6, fontWeight: 500 }

  return (
    <main style={{ minHeight: '100vh', background: '#0f1117', fontFamily: "'Georgia', serif", color: '#e8e0d0' }}>
      <header style={{
        borderBottom: '1px solid #2a2a3a', padding: '16px 48px',
        display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: '#0d0f14',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
          <img src={LOGO} alt="BOP Logo" style={{ width: 52, height: 52, objectFit: 'contain', flexShrink: 0 }} />
          <div>
            <div style={{ fontSize: 18, fontWeight: 600, letterSpacing: '0.04em', color: '#c8a96e' }}>
              EQT - Expense Invoice Generator
            </div>
            <div style={{ fontSize: 11, color: '#666', letterSpacing: '0.12em', textTransform: 'uppercase' }}>
              BOP Abstract / BOP Acquisition
            </div>
          </div>
        </div>
        <a href="/help" style={{
          color: '#c8a96e', fontSize: 13, textDecoration: 'none',
          border: '1px solid #333', padding: '6px 14px', borderRadius: 4, letterSpacing: '0.04em',
        }}>User Guide</a>
      </header>

      <div style={{ maxWidth: 820, margin: '0 auto', padding: '48px 32px' }}>
        <div style={{ marginBottom: 48 }}>
          <h1 style={{ fontSize: 36, fontWeight: 400, color: '#e8e0d0', margin: '0 0 12px 0', letterSpacing: '-0.01em', lineHeight: 1.2 }}>
            Generate Expense Invoices
          </h1>
          <p style={{ color: '#888', fontSize: 15, margin: 0, lineHeight: 1.6 }}>
            Upload your Excel invoice files and any receipt PDFs. Fill in the invoice details below.
            One complete PDF invoice will be generated per Excel file.
          </p>
        </div>

        <Section number="1" title="Upload Excel Invoice Files">
          <MultiUploadBox
            label="Drop one or more .xlsx invoice files here or click to browse"
            accept=".xlsx,.xls"
            files={excelFiles}
            onChange={e => setExcelFiles(Array.from(e.target.files || []))}
            inputRef={excelRef}
            icon="📊"
          />
        </Section>

        <Section number="2" title="Upload Receipt PDFs (Optional)">
          <MultiUploadBox
            label="Drop receipt PDFs here or click to browse — optional"
            accept=".pdf"
            files={receiptFiles}
            onChange={e => setReceiptFiles(Array.from(e.target.files || []))}
            inputRef={receiptRef}
            icon="🧾"
          />
          <div style={{ fontSize: 12, color: '#666', marginTop: 8 }}>
            Receipts are matched to invoices by invoice number and appended to the end.
          </div>
        </Section>

        <Section number="3" title="Invoice Details">
          <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>

            <div>
              <div style={labelStyle}>Company Name *</div>
              <select value={companyName} onChange={e => setCompanyName(e.target.value)} style={selectStyle}>
                <option value="">— Select company —</option>
                {COMPANY_OPTIONS.map(o => <option key={o} value={o}>{o}</option>)}
              </select>
            </div>

            <div>
              <div style={labelStyle}>Manager / Attn</div>
              <select value={manager} onChange={e => setManager(e.target.value)} style={selectStyle}>
                <option value="">— Select manager —</option>
                {MANAGER_OPTIONS.map(o => <option key={o} value={o}>{o}</option>)}
              </select>
            </div>

            <div>
              <div style={labelStyle}>AFE</div>
              <select value={afeKey} onChange={e => setAfeKey(e.target.value)} style={selectStyle}>
                <option value="">— Select AFE —</option>
                {AFE_OPTIONS.map(o => (
                  <option key={o.afe} value={o.afe}>
                    {o.label} — {o.afe}
                  </option>
                ))}
              </select>
              {afeKey && (
                <div style={{ fontSize: 11, color: '#c8a96e', marginTop: 4 }}>
                  AFE number on invoice: {afeKey}
                </div>
              )}
            </div>

            <div>
              <div style={labelStyle}>County</div>
              <select value={county} onChange={e => setCounty(e.target.value)} style={selectStyle}>
                <option value="">— Select county —</option>
                {COUNTY_OPTIONS.map(o => <option key={o} value={o}>{o}</option>)}
              </select>
            </div>

            <div>
              <div style={labelStyle}>Invoice Date *</div>
              <input
                type="text"
                placeholder="e.g. May 8, 2026"
                value={invoiceDate}
                onChange={e => setInvoiceDate(e.target.value)}
                style={inputStyle}
              />
              <div style={{ fontSize: 11, color: '#555', marginTop: 4 }}>Any format accepted</div>
            </div>

          </div>
        </Section>

        <Section number="4" title="Generate Invoices">
          <button
            onClick={handleGenerate}
            disabled={status === 'loading'}
            style={{
              width: '100%', padding: '16px 32px',
              background: status === 'loading' ? '#2a2a3a' : 'linear-gradient(135deg, #c8a96e, #8b6914)',
              color: status === 'loading' ? '#666' : '#fff',
              border: 'none', borderRadius: 6, fontSize: 16,
              fontFamily: "'Georgia', serif", letterSpacing: '0.04em',
              cursor: status === 'loading' ? 'not-allowed' : 'pointer', transition: 'all 0.2s',
            }}
          >
            {status === 'loading' ? '⏳ Generating...' : '⬇ Generate & Download ZIP'}
          </button>

          {message && (
            <div style={{
              marginTop: 16, padding: '12px 16px', borderRadius: 6,
              background: status === 'error' ? 'rgba(200,60,60,0.1)' : 'rgba(60,180,100,0.1)',
              border: `1px solid ${status === 'error' ? '#8b2020' : '#2a6640'}`,
              color: status === 'error' ? '#e07070' : '#70c090', fontSize: 14,
            }}>{message}</div>
          )}
        </Section>

        <div style={{ marginTop: 48, padding: 24, background: '#0d0f14', borderRadius: 8, border: '1px solid #1e1e2e' }}>
          <div style={{ fontSize: 11, color: '#c8a96e', letterSpacing: '0.1em', textTransform: 'uppercase', marginBottom: 16 }}>
            How It Works
          </div>
          <div style={{ fontSize: 13, color: '#888', lineHeight: 1.8 }}>
            Each Excel file becomes one invoice PDF with 2 pages:<br />
            <span style={{ color: '#c8a96e' }}>Page 1</span> — Invoice summary with broker/expense table<br />
            <span style={{ color: '#c8a96e' }}>Page 2+</span> — Work detail log entries<br />
            <span style={{ color: '#c8a96e' }}>Final pages</span> — Receipt PDFs appended if matched by invoice number
          </div>
        </div>
      </div>
    </main>
  )
}

function Section({ number, title, children }: { number: string; title: string; children: React.ReactNode }) {
  return (
    <div style={{ marginBottom: 40 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16 }}>
        <div style={{
          width: 28, height: 28, borderRadius: '50%',
          background: 'rgba(200,169,110,0.15)', border: '1px solid #c8a96e',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          fontSize: 13, color: '#c8a96e', fontWeight: 600, flexShrink: 0,
        }}>{number}</div>
        <h2 style={{ margin: 0, fontSize: 17, fontWeight: 500, color: '#e8e0d0', letterSpacing: '0.01em' }}>{title}</h2>
      </div>
      {children}
    </div>
  )
}

function MultiUploadBox({ label, accept, files, onChange, inputRef, icon }: {
  label: string
  accept: string
  files: File[]
  onChange: (e: React.ChangeEvent<HTMLInputElement>) => void
  inputRef: React.RefObject<HTMLInputElement>
  icon: string
}) {
  return (
    <div onClick={() => inputRef.current?.click()} style={{
      border: `2px dashed ${files.length > 0 ? '#c8a96e' : '#2a2a3a'}`,
      borderRadius: 8, padding: '28px 24px', textAlign: 'center', cursor: 'pointer',
      background: files.length > 0 ? 'rgba(200,169,110,0.04)' : '#0d0f14', transition: 'all 0.2s',
    }}>
      <input ref={inputRef} type="file" accept={accept} multiple onChange={onChange} style={{ display: 'none' }} />
      {files.length === 0 ? (
        <>
          <div style={{ fontSize: 28, marginBottom: 8 }}>{icon}</div>
          <div style={{ color: '#888', fontSize: 14 }}>{label}</div>
          <div style={{ color: '#555', fontSize: 12, marginTop: 4 }}>{accept.toUpperCase().replace(/\./g, '').replace(/,/g, ' / ')}</div>
        </>
      ) : (
        <div style={{ textAlign: 'left' }}>
          {files.map(f => (
            <div key={f.name} style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 4 }}>
              <span style={{ fontSize: 16 }}>✅</span>
              <span style={{ color: '#c8a96e', fontSize: 14 }}>{f.name}</span>
              <span style={{ color: '#555', fontSize: 12 }}>({(f.size / 1024).toFixed(1)} KB)</span>
            </div>
          ))}
          <div style={{ color: '#555', fontSize: 12, marginTop: 8 }}>Click to change</div>
        </div>
      )}
    </div>
  )
}
