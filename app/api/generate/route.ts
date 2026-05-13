import { NextRequest, NextResponse } from 'next/server'
import * as XLSX from 'xlsx'
import JSZip from 'jszip'

const BILL_TO = {
  company: 'EQT Production Company',
  address: '2200 Energy Drive',
  city: 'Canonsburg, Pennsylvania 15317',
}

function formatDate(val: unknown): string {
  if (!val) return ''
  if (val instanceof Date) return val.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' })
  const d = new Date(String(val))
  if (!isNaN(d.getTime())) return d.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' })
  return String(val)
}

function normalizeInvoiceDate(input: string): string {
  const d = new Date(input)
  if (!isNaN(d.getTime())) {
    return d.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })
  }
  return input
}

function sanitize(name: string): string {
  return name.replace(/[^a-zA-Z0-9_\-. #,]/g, '_').trim()
}

function fmtCurrency(val: unknown): string {
  const n = Number(val ?? 0)
  return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(n)
}

function fmtNum(val: unknown, decimals = 1): string {
  const n = Number(val ?? 0)
  return new Intl.NumberFormat('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(n)
}

function matchReceipt(invoiceNum: string, receiptFiles: { name: string; buffer: Buffer }[]): Buffer | null {
  for (const r of receiptFiles) {
    if (r.name.includes(invoiceNum)) return r.buffer
  }
  return null
}

async function buildInvoicePdf(
  excelBuffer: Buffer,
  companyName: string,
  manager: string,
  countyOverride: string,
  invoiceDateOverride: string,
  matchedReceiptBuffer: Buffer | null,
): Promise<Buffer> {
  const { jsPDF } = await import('jspdf')
  const autoTable = (await import('jspdf-autotable')).default

  const workbook = XLSX.read(excelBuffer, { type: 'buffer', cellDates: true })
  const summarySheet = workbook.Sheets['Summary']
  const detailSheet = workbook.Sheets['Work Detail']

  if (!summarySheet) throw new Error('No Summary sheet found')

  const summaryRows = XLSX.utils.sheet_to_json(summarySheet, { header: 1, defval: '' }) as unknown[][]

  const bopCompany = companyName || String((summaryRows[0] as unknown[])?.[0] ?? 'BOP Abstract, LLC')
  const invoiceNum = String((summaryRows[7] as unknown[])?.[5] ?? '')
  const invoiceDate = invoiceDateOverride || String((summaryRows[7] as unknown[])?.[1] ?? '')
  const afe = String((summaryRows[9] as unknown[])?.[5] ?? '')
  const attn = manager ? `Attn: ${manager}` : String((summaryRows[10] as unknown[])?.[1] ?? '')
  const project = String((summaryRows[13] as unknown[])?.[5] ?? '')
  const period = String((summaryRows[15] as unknown[])?.[1] ?? '')
  const county = countyOverride || ''

  const headerRow = (summaryRows[17] as unknown[]).map(h => String(h ?? '').toLowerCase())
  const hasMisc = headerRow.some(h => h.includes('misc'))
  const hasCopies = headerRow.some(h => h.includes('cop'))

  const brokerRows: unknown[][] = []
  for (let i = 18; i < summaryRows.length; i++) {
    const row = summaryRows[i] as unknown[]
    if (row && row[0] && String(row[0]).trim()) brokerRows.push(row)
  }
  const brokerDataRows = brokerRows.filter(r => String((r as unknown[])[0]).toLowerCase() !== 'totals')
  const brokerTotalsRow = brokerRows.find(r => String((r as unknown[])[0]).toLowerCase() === 'totals')

  const detailRows = detailSheet
    ? (XLSX.utils.sheet_to_json(detailSheet, { header: 1, defval: '' }) as unknown[][]).slice(2)
    : []
  const detailDataRows = detailRows.filter(r => {
    const row = r as unknown[]
    return row[0] && String(row[0]).trim() !== ''
  })

  const detailHasMisc = detailSheet
    ? (XLSX.utils.sheet_to_json(detailSheet, { header: 1, defval: '' }) as unknown[][])[1]
        ?.some((h: unknown) => String(h).toLowerCase().includes('misc'))
    : false

  const black = [0, 0, 0] as [number, number, number]
  const red = [255, 0, 0] as [number, number, number]
  const headerBg = [242, 220, 219] as [number, number, number]
  const totalsBg = [255, 255, 0] as [number, number, number]
  const white = [255, 255, 255] as [number, number, number]
  const lightGray = [100, 100, 100] as [number, number, number]

  const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'letter' })

  // ============ PAGE 1: SUMMARY ============
  doc.setFillColor(...white)
  doc.rect(0, 0, 612, 792, 'F')

  doc.setFont('helvetica', 'bold')
  doc.setFontSize(11)
  doc.setTextColor(...black)
  doc.text(bopCompany, 306, 45, { align: 'center' })
  doc.setFont('helvetica', 'normal')
  doc.setFontSize(9)
  doc.text('2547 Washington Rd. Bldg. 700, Ste. 720', 306, 57, { align: 'center' })
  doc.text('Pittsburgh, PA, 15241', 306, 69, { align: 'center' })
  doc.text('724-747-1594', 306, 81, { align: 'center' })

  doc.setFont('helvetica', 'bold')
  doc.setFontSize(13)
  doc.text('INVOICE', 306, 105, { align: 'center' })

  doc.setDrawColor(...black)
  doc.setLineWidth(0.5)
  doc.line(40, 112, 572, 112)

  doc.setFontSize(9)
  const lx = 40
  const rx = 320
  let ly = 128
  let ry = 128

  function leftLabel(label: string, value: string) {
    doc.setFont('helvetica', 'bold'); doc.setTextColor(...lightGray)
    doc.text(label, lx, ly)
    doc.setFont('helvetica', 'normal'); doc.setTextColor(...black)
    doc.text(value, lx + 52, ly)
    ly += 13
  }

  function rightLabel(label: string, value: string) {
    doc.setFont('helvetica', 'bold'); doc.setTextColor(...lightGray)
    doc.text(label, rx, ry)
    doc.setFont('helvetica', 'normal'); doc.setTextColor(...black)
    doc.text(value, rx + 65, ry)
    ry += 13
  }

  leftLabel('Date:', invoiceDate)
  rightLabel('Invoice #:', invoiceNum)
  ly += 4; ry += 4

  doc.setFont('helvetica', 'bold'); doc.setTextColor(...lightGray)
  doc.text('Bill To:', lx, ly)
  doc.setFont('helvetica', 'bold'); doc.setTextColor(...black)
  doc.text(BILL_TO.company, lx + 52, ly); ly += 13
  doc.setFont('helvetica', 'normal'); doc.setTextColor(...black)
  if (attn) { doc.text(attn, lx + 52, ly); ly += 13 }
  doc.text(BILL_TO.address, lx + 52, ly); ly += 13
  doc.text(BILL_TO.city, lx + 52, ly); ly += 13

  if (afe) rightLabel('AFE#:', afe)
  if (project) rightLabel('Project:', project)
  if (county) rightLabel('County:', `${county}, PA`)

  ly += 8
  doc.setFont('helvetica', 'bold'); doc.setTextColor(...lightGray)
  doc.text('Period:', lx, ly)
  doc.setFont('helvetica', 'normal'); doc.setTextColor(...black)
  doc.text(period, lx + 52, ly)

  doc.setTextColor(...red)
  doc.setFont('helvetica', 'bold')
  doc.text('DUE UPON RECEIPT', 572, ly, { align: 'right' })
  doc.setTextColor(...black)

  const tableStartY = Math.max(ly + 20, ry + 20)

  let summaryHead: string[][]
  let summaryBody: string[][]
  let totalColIndex: number

  if (hasMisc && hasCopies) {
    summaryHead = [['Broker', 'Miles\nDriven', 'Mileage Amt\n@ 0.7250/mile', 'Miscellaneous', 'Copies', 'TOTAL']]
    totalColIndex = 5
    summaryBody = brokerDataRows.map(r => {
      const row = r as unknown[]
      return [String(row[0] ?? ''), fmtNum(row[4]), fmtCurrency(row[5]), fmtCurrency(row[6]), fmtCurrency(row[7]), fmtCurrency(row[8])]
    })
    if (brokerTotalsRow) {
      const t = brokerTotalsRow as unknown[]
      summaryBody.push(['Totals', fmtNum(t[4]), fmtCurrency(t[5]), fmtCurrency(t[6]), fmtCurrency(t[7]), fmtCurrency(t[8])])
    }
  } else if (hasMisc) {
    summaryHead = [['Broker', 'Miles\nDriven', 'Mileage Amt\n@ 0.7250/mile', 'Miscellaneous', 'TOTAL']]
    totalColIndex = 4
    summaryBody = brokerDataRows.map(r => {
      const row = r as unknown[]
      return [String(row[0] ?? ''), fmtNum(row[4]), fmtCurrency(row[5]), fmtCurrency(row[6]), fmtCurrency(row[7])]
    })
    if (brokerTotalsRow) {
      const t = brokerTotalsRow as unknown[]
      summaryBody.push(['Totals', fmtNum(t[4]), fmtCurrency(t[5]), fmtCurrency(t[6]), fmtCurrency(t[7])])
    }
  } else {
    summaryHead = [['Broker', 'Miles\nDriven', 'Mileage Amt\n@ 0.7250/mile', 'TOTAL']]
    totalColIndex = 3
    summaryBody = brokerDataRows.map(r => {
      const row = r as unknown[]
      return [String(row[0] ?? ''), fmtNum(row[4]), fmtCurrency(row[5]), fmtCurrency(row[6])]
    })
    if (brokerTotalsRow) {
      const t = brokerTotalsRow as unknown[]
      summaryBody.push(['Totals', fmtNum(t[4]), fmtCurrency(t[5]), fmtCurrency(t[6])])
    }
  }

  const totalsRowIndex = summaryBody.length - 1

  autoTable(doc, {
    startY: tableStartY,
    head: summaryHead,
    body: summaryBody,
    theme: 'grid',
    styles: { font: 'helvetica', fontSize: 8, textColor: black, fillColor: white, cellPadding: 4, lineColor: black, lineWidth: 0.3, overflow: 'linebreak' },
    headStyles: { textColor: black, fillColor: headerBg, fontStyle: 'bold', halign: 'center', valign: 'middle' },
    columnStyles: {
      0: { cellWidth: 110 },
      1: { halign: 'center', cellWidth: 55 },
      2: { halign: 'right', cellWidth: 90 },
      3: { halign: 'right', cellWidth: hasMisc ? 80 : 90 },
      4: { halign: 'right', cellWidth: 80 },
      5: { halign: 'right', cellWidth: 80 },
    },
    didParseCell: function(data) {
      if (data.section === 'body' && data.row.index === totalsRowIndex) {
        data.cell.styles.fontStyle = 'bold'
        if (data.column.index === totalColIndex) {
          data.cell.styles.fillColor = totalsBg
        }
      }
    },
    margin: { left: 40, right: 40 },
  })

  const finalY = (doc as unknown as { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 20
  doc.setFontSize(8)
  doc.setTextColor(...lightGray)
  doc.setFont('helvetica', 'italic')
  doc.text('Please contact our accounting department with any questions regarding invoices', 306, finalY, { align: 'center' })

  // ============ PAGE 2: WORK DETAIL ============
  doc.addPage()
  doc.setFillColor(...white)
  doc.rect(0, 0, 612, 792, 'F')

  doc.setFont('helvetica', 'bold')
  doc.setFontSize(11)
  doc.setTextColor(...black)
  doc.text('Work Detail', 40, 45)
  doc.setLineWidth(0.5)
  doc.line(40, 52, 572, 52)

  let totalMiles = 0, totalMileageAmt = 0, totalMisc = 0, totalTotal = 0

  detailDataRows.forEach(r => {
    const row = r as unknown[]
    totalMiles += Number(row[8] ?? 0)
    totalMileageAmt += Number(row[9] ?? 0)
    if (detailHasMisc) {
      totalMisc += Number(row[11] ?? 0)
      totalTotal += Number(row[13] ?? 0)
    } else {
      totalTotal += Number(row[11] ?? 0)
    }
  })

  let detailHead: string[][]
  let detailBody: string[][]
  let detailTotalCol: number

  if (detailHasMisc) {
    detailHead = [['Landman', 'Date', 'Prospect', 'Legal', 'Miles', 'Mileage\n0.7250/mi', 'Misc.', 'Total', 'Description']]
    detailTotalCol = 7
    detailBody = detailDataRows.map(r => {
      const row = r as unknown[]
      return [
        String(row[0] ?? ''),
        formatDate(row[1]),
        String(row[2] ?? ''),
        String(row[3] ?? ''),
        fmtNum(row[8], 1),
        fmtCurrency(row[9]),
        fmtCurrency(row[11]),
        fmtCurrency(row[13]),
        String(row[14] ?? ''),
      ]
    })
    detailBody.push(['Totals', '', '', '', fmtNum(totalMiles, 1), fmtCurrency(totalMileageAmt), fmtCurrency(totalMisc), fmtCurrency(totalTotal), ''])
  } else {
    detailHead = [['Landman', 'Date', 'Prospect', 'Legal', 'Miles', 'Mileage\n0.7250/mi', 'Total', 'Description']]
    detailTotalCol = 6
    detailBody = detailDataRows.map(r => {
      const row = r as unknown[]
      return [
        String(row[0] ?? ''),
        formatDate(row[1]),
        String(row[2] ?? ''),
        String(row[3] ?? ''),
        fmtNum(row[8], 1),
        fmtCurrency(row[9]),
        fmtCurrency(row[11]),
        String(row[12] ?? ''),
      ]
    })
    detailBody.push(['Totals', '', '', '', fmtNum(totalMiles, 1), fmtCurrency(totalMileageAmt), fmtCurrency(totalTotal), ''])
  }

  const detailTotalsIndex = detailBody.length - 1

  autoTable(doc, {
    startY: 60,
    head: detailHead,
    body: detailBody,
    theme: 'grid',
    styles: {
      font: 'helvetica', fontSize: 7, textColor: black, fillColor: white,
      cellPadding: 3, lineColor: black, lineWidth: 0.3,
      overflow: 'linebreak', valign: 'top',
    },
    headStyles: { textColor: black, fillColor: headerBg, fontStyle: 'bold', halign: 'center', valign: 'middle' },
    columnStyles: detailHasMisc ? {
      0: { cellWidth: 55, overflow: 'linebreak' },
      1: { cellWidth: 45 },
      2: { cellWidth: 60, overflow: 'linebreak' },
      3: { cellWidth: 55, overflow: 'linebreak' },
      4: { cellWidth: 30, halign: 'center' },
      5: { cellWidth: 52, halign: 'right' },
      6: { cellWidth: 45, halign: 'right' },
      7: { cellWidth: 45, halign: 'right' },
      8: { cellWidth: 145, overflow: 'linebreak' },
    } : {
      0: { cellWidth: 55, overflow: 'linebreak' },
      1: { cellWidth: 45 },
      2: { cellWidth: 65, overflow: 'linebreak' },
      3: { cellWidth: 60, overflow: 'linebreak' },
      4: { cellWidth: 30, halign: 'center' },
      5: { cellWidth: 55, halign: 'right' },
      6: { cellWidth: 55, halign: 'right' },
      7: { cellWidth: 167, overflow: 'linebreak' },
    },
    didParseCell: function(data) {
      if (data.section === 'body' && data.row.index === detailTotalsIndex) {
        data.cell.styles.fontStyle = 'bold'
        if (data.column.index === detailTotalCol) {
          data.cell.styles.fillColor = totalsBg
        }
      }
    },
    margin: { left: 40, right: 40 },
  })

  // ============ APPEND RECEIPTS ============
  try {
    const { PDFDocument } = await import('pdf-lib')
    const mainPdfBytes = doc.output('arraybuffer')
    const mainPdf = await PDFDocument.load(mainPdfBytes)

    if (matchedReceiptBuffer) {
      const receiptPdf = await PDFDocument.load(matchedReceiptBuffer)
      const receiptPages = await mainPdf.copyPages(receiptPdf, receiptPdf.getPageIndices())
      receiptPages.forEach(p => mainPdf.addPage(p))
    }

    const finalBytes = await mainPdf.save()
    return Buffer.from(finalBytes)
  } catch {
    return Buffer.from(doc.output('arraybuffer'))
  }
}

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData()
    const excelFiles = formData.getAll('excel') as File[]
    const receiptFiles = formData.getAll('receipts') as File[]
    const companyName = (formData.get('companyName') as string) || 'BOP Abstract, LLC'
    const manager = (formData.get('manager') as string) || ''
    const county = (formData.get('county') as string) || ''
    const rawDate = (formData.get('invoiceDate') as string) || ''
    const invoiceDateOverride = normalizeInvoiceDate(rawDate)

    if (!excelFiles.length) return NextResponse.json({ error: 'No Excel files uploaded' }, { status: 400 })

    const receiptData: { name: string; buffer: Buffer }[] = []
    for (const r of receiptFiles) {
      receiptData.push({ name: r.name, buffer: Buffer.from(await r.arrayBuffer()) })
    }

    const zip = new JSZip()

    for (const excelFile of excelFiles) {
      const excelBuffer = Buffer.from(await excelFile.arrayBuffer())
      const wb = XLSX.read(excelBuffer, { type: 'buffer', cellDates: true })
      const ws = wb.Sheets['Summary']
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as unknown[][]
      const invoiceNum = String((rows[7] as unknown[])?.[5] ?? '').trim() ||
        (excelFile.name.match(/(\d{5,})/)?.[1] ?? 'UNKNOWN')
      const period = String((rows[15] as unknown[])?.[1] ?? '').trim()

      const matchedReceipt = matchReceipt(invoiceNum, receiptData)

      const shortCompany = companyName.replace(/, LLC|, Inc\.?/gi, '').trim()
      const countyPart = county ? `${county} County ` : ''
      const outputName = sanitize(`${shortCompany} - EQT ${countyPart}Expenses - ${period}`)

      try {
        const pdfBuffer = await buildInvoicePdf(
          excelBuffer, companyName, manager, county, invoiceDateOverride, matchedReceipt
        )
        zip.file(`${outputName}.pdf`, pdfBuffer)
      } catch (err) {
        return NextResponse.json({ error: `Failed for invoice ${invoiceNum}: ${String(err)}` }, { status: 500 })
      }
    }

    const zipBuffer = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' })
    return new NextResponse(zipBuffer, {
      headers: {
        'Content-Type': 'application/zip',
        'Content-Disposition': 'attachment; filename="eqt-expense-invoices.zip"',
      },
    })
  } catch (err) {
    console.error(err)
    return NextResponse.json({ error: String(err) }, { status: 500 })
  }
}
