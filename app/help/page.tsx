'use client'

const LOGO = "https://i.imgur.com/szjzoxt.png"

export default function HelpPage() {
  return (
    <main style={{ minHeight: '100vh', background: '#0f1117', fontFamily: "'Georgia', serif", color: '#e8e0d0' }}>
      <header style={{
        borderBottom: '1px solid #2a2a3a', padding: '16px 48px',
        display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: '#0d0f14',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
          <img src={LOGO} alt="BOP Logo" style={{ width: 52, height: 52, objectFit: 'contain' }} />
          <div>
            <div style={{ fontSize: 18, fontWeight: 600, letterSpacing: '0.04em', color: '#c8a96e' }}>EQT - Expense Invoice Generator</div>
            <div style={{ fontSize: 11, color: '#666', letterSpacing: '0.12em', textTransform: 'uppercase' }}>BOP Abstract / BOP Acquisition</div>
          </div>
        </div>
        <a href="/" style={{ color: '#c8a96e', fontSize: 13, textDecoration: 'none', border: '1px solid #333', padding: '6px 14px', borderRadius: 4 }}>
          ← Back to App
        </a>
      </header>

      <div style={{ maxWidth: 820, margin: '0 auto', padding: '48px 32px' }}>
        <h1 style={{ fontSize: 32, fontWeight: 400, color: '#e8e0d0', marginBottom: 8 }}>User Guide</h1>
        <p style={{ color: '#888', fontSize: 15, marginBottom: 40 }}>How to use the EQT Expense Invoice Generator</p>

        {[
          {
            title: '1. Prepare Your Excel Files',
            content: 'Each Excel file represents one invoice. It must have two sheets:\n\n• Summary sheet — BOP header, Invoice #, Date, Period, broker/expense table\n• Work Detail sheet — line-by-line landman work entries\n\nYou can upload multiple Excel files at once for batch processing.'
          },
          {
            title: '2. Upload Receipt PDFs (Optional)',
            content: 'Upload any receipt PDFs (Amazon orders, USPS receipts, etc.). The app matches receipts to invoices by invoice number in the filename. Matched receipts are appended as additional pages at the end of the invoice.'
          },
          {
            title: '3. Fill In Invoice Details',
            content: 'Company Name — appears in the BOP header (e.g. BOP Abstract, LLC or BOP Acquisition, LLC)\n\nManager — the Attn: name in the Bill To section (e.g. Andrew Restanio)\n\nCounty — the county for the invoice (e.g. Allegheny)\n\nInvoice Date — the date to appear on all invoices (any format accepted)'
          },
          {
            title: '4. Generate & Download',
            content: 'Click "Generate & Download ZIP". The app creates one PDF per Excel file and packages them into a ZIP for download.'
          },
          {
            title: 'Output PDF Structure',
            content: 'Each invoice PDF contains:\n• Page 1 — Invoice summary with expense/mileage table\n• Page 2+ — Work detail log entries\n• Final pages — Receipt PDFs appended if matched by invoice number'
          },
          {
            title: 'File Naming',
            content: 'Output files are named:\n[Company] - EQT [County] County Expenses - [Period].pdf\n\nExample: BOP Acquisition - EQT Allegheny County Expenses - 4-1 to 4-30-2026.pdf'
          },
        ].map((section, i) => (
          <div key={i} style={{ marginBottom: 32, padding: 24, background: '#0d0f14', borderRadius: 8, border: '1px solid #1e1e2e' }}>
            <div style={{ fontSize: 15, fontWeight: 600, color: '#c8a96e', marginBottom: 12 }}>{section.title}</div>
            <div style={{ fontSize: 14, color: '#aaa', lineHeight: 1.8, whiteSpace: 'pre-line' }}>{section.content}</div>
          </div>
        ))}
      </div>
    </main>
  )
}
