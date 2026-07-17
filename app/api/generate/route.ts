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

function filenameDateFormat(input: string): string {
  const d = new Date(input)
  if (!isNaN(d.getTime())) {
    const month = d.getMonth() + 1
    const day = String(d.getDate()).padStart(2, '0')
    const year = String(d.getFullYear()).slice(-2)
    return `${month}.${day}.${year}`
  }
  return input.replace(/[^0-9.]/g, '_')
}

function extractPeriodEndDate(period: string): string {
  const match = period.match(/to\s+(\d{1,2}-\d{1,2}-\d{4})/i)
  if (match) return filenameDateFormat(match[1].replace(/-/g, '/'))
  const dates = [...period.matchAll(/\d{1,2}-\d{1,2}-\d{4}/g)]
  if (dates.length > 0) return filenameDateFormat(dates[dates.length - 1][0].replace(/-/g, '/'))
  return period
}

function sanitize(name: string): string {
  return name.replace(/[^a-zA-Z0-9_\-. #,]/g, '_').trim()
}

function fmtCurrency(val: unknown): string {
  const n = Number(val ?? 0)
  if (isNaN(n)) return '$0.00'
  return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(n)
}

function fmtNum(val: unknown, decimals = 1): string {
  const n = Number(val ?? 0)
  if (isNaN(n)) return '0.0'
  return new Intl.NumberFormat('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals }).format(n)
}

function matchReceipt(invoiceNum: string, receiptFiles: { name: string; buffer: Buffer }[]): Buffer | null {
  for (const r of receiptFiles) {
    if (r.name.includes(invoiceNum)) return r.buffer
  }
  return null
}

// Find the summary header row index by scanning for 'broker' in column 0
function findSummaryHeaderRow(summaryRows: unknown[][]): number {
  for (let i = 0; i < summaryRows.length; i++) {
    const row = summaryRows[i] as unknown[]
    if (row && String(row[0] ?? '').toLowerCase().trim() === 'broker') return i
  }
  return 17 // fallback
}

async function buildInvoicePdf(
  excelBuffer: Buffer,
  companyName: string,
  manager: string,
  afeOverride: string,
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
  const afeFromSheet = String((summaryRows[9] as unknown[])?.[5] ?? '')
  const afe = afeOverride || afeFromSheet
  const attn = manager ? `Attn: ${manager}` : String((summaryRows[10] as unknown[])?.[1] ?? '')
  const county = countyOverride || ''

  // Find period — scan for 'period' label in col 0
  let period = ''
  for (let i = 0; i < summaryRows.length; i++) {
    if (String((summaryRows[i] as unknown[])?.[0] ?? '').toLowerCase().includes('period')) {
      period = String((summaryRows[i] as unknown[])?.[1] ?? '')
      break
    }
  }

  // Find header row dynamically
  const headerRowIdx = findSummaryHeaderRow(summaryRows)
  const summaryHeaderRow = (summaryRows[headerRowIdx] as unknown[] ?? []).map(h => String(h ?? '').toLowerCase())
  const isDayrateTemplate = summaryHeaderRow.some(h => h.includes('day') || h.includes('professional service')) &&
    !summaryHeaderRow.some(h => h.includes('mile') || h.includes('mileage'))

  // Collect broker data rows (start after header row)
  const brokerRows: unknown[][] = []
  for (let i = headerRowIdx + 1; i < summaryRows.length; i++) {
    const row = summaryRows[i] as unknown[]
    if (row && row[0] && String(row[0]).trim()) brokerRows.push(row)
  }
  const brokerDataRows = brokerRows.filter(r => String((r as unknown[])[0]).toLowerCase() !== 'totals')
  const brokerTotalsRow = brokerRows.find(r => String((r as unknown[])[0]).toLowerCase() === 'totals')

  // Detail sheet
  const detailAllRows = detailSheet
    ? (XLSX.utils.sheet_to_json(detailSheet, { header: 1, defval: '' }) as unknown[][])
    : []
  const detailHeaderRow = (detailAllRows[1] ?? []).map((h: unknown) => String(h).toLowerCase())
  const detailIsDayrate = detailHeaderRow.some(h => h.includes('day') || h.includes('dayrate') || h.includes('labor')) &&
    !detailHeaderRow.some(h => h.includes('mile') || h.includes('mileage'))
  const detailHasMisc = detailHeaderRow.some(h => h.includes('misc'))

  // Detect if detail has a separate dayrate/labor column or goes straight Days→Misc→Total
  // If header has 'day rate' or 'labor total' it's the full dayrate layout; if just 'days' then simplified
  const detailHasDayRate = detailHeaderRow.some(h => h.includes('day rate') || h.includes('dayrate') || h.includes('labor total') || h.includes('amt. per day'))

  const detailDataRows = detailAllRows.slice(2).filter(r => {
    const row = r as unknown[]
    return row[0] && String(row[0]).trim() !== ''
  })

  const black: [number, number, number] = [0, 0, 0]
  const red: [number, number, number] = [255, 0, 0]
  const headerBg: [number, number, number] = [242, 220, 219]
  const totalsBg: [number, number, number] = [255, 255, 0]
  const white: [number, number, number] = [255, 255, 255]
  const lightGray: [number, number, number] = [100, 100, 100]

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

  // ============ SUMMARY TABLE ============
  // Always display mileage-style headers; map dayrate cols correctly
  // Dayrate full (col[4]=Misc, col[5]=TOTAL)
  // Mileage (col[4]=Miles, col[5]=MileageAmt, col[6]=Misc, col[7]=TOTAL)

  const summaryHasMisc = summaryHeaderRow.some(h => h.includes('misc'))
  const summaryHasCopies = summaryHeaderRow.some(h => h.includes('cop'))

  let summaryHead: string[][]
  let summaryBody: string[][]
  let totalColIndex: number

  if (summaryHasMisc && summaryHasCopies) {
    summaryHead = [['Broker', 'Miles\nDriven', 'Mileage Amt\n@ 0.7250/mile', 'Miscellaneous', 'Copies', 'TOTAL']]
    totalColIndex = 5
    summaryBody = brokerDataRows.map(r => {
      const row = r as unknown[]
      if (isDayrateTemplate) {
        return [String(row[0] ?? ''), '0.0', '$0.00', fmtCurrency(row[3]), fmtCurrency(row[4]), fmtCurrency(row[5])]
      }
      return [String(row[0] ?? ''), fmtNum(row[4]), fmtCurrency(row[5]), fmtCurrency(row[6]), fmtCurrency(row[7]), fmtCurrency(row[8])]
    })
    if (brokerTotalsRow) {
      const t = brokerTotalsRow as unknown[]
      if (isDayrateTemplate) {
        summaryBody.push(['Totals', '0.0', '$0.00', fmtCurrency(t[3]), fmtCurrency(t[4]), fmtCurrency(t[5])])
      } else {
        summaryBody.push(['Totals', fmtNum(t[4]), fmtCurrency(t[5]), fmtCurrency(t[6]), fmtCurrency(t[7]), fmtCurrency(t[8])])
      }
    }
  } else if (summaryHasMisc) {
    summaryHead = [['Broker', 'Miles\nDriven', 'Mileage Amt\n@ 0.7250/mile', 'Miscellaneous', 'TOTAL']]
    totalColIndex = 4
    summaryBody = brokerDataRows.map(r => {
      const row = r as unknown[]
      if (isDayrateTemplate) {
        return [String(row[0] ?? ''), '0.0', '$0.00', fmtCurrency(row[4]), fmtCurrency(row[5])]
      }
      return [String(row[0] ?? ''), fmtNum(row[4]), fmtCurrency(row[5]), fmtCurrency(row[6]), fmtCurrency(row[7])]
    })
    if (brokerTotalsRow) {
      const t = brokerTotalsRow as unknown[]
      if (isDayrateTemplate) {
        summaryBody.push(['Totals', '0.0', '$0.00', fmtCurrency(t[4]), fmtCurrency(t[5])])
      } else {
        summaryBody.push(['Totals', fmtNum(t[4]), fmtCurrency(t[5]), fmtCurrency(t[6]), fmtCurrency(t[7])])
      }
    }
  } else {
    summaryHead = [['Broker', 'Miles\nDriven', 'Mileage Amt\n@ 0.7250/mile', 'TOTAL']]
    totalColIndex = 3
    summaryBody = brokerDataRows.map(r => {
      const row = r as unknown[]
      if (isDayrateTemplate) {
        return [String(row[0] ?? ''), '0.0', '$0.00', fmtCurrency(row[5])]
      }
      return [String(row[0] ?? ''), fmtNum(row[4]), fmtCurrency(row[5]), fmtCurrency(row[6])]
    })
    if (brokerTotalsRow) {
      const t = brokerTotalsRow as unknown[]
      if (isDayrateTemplate) {
        summaryBody.push(['Totals', '0.0', '$0.00', fmtCurrency(t[5])])
      } else {
        summaryBody.push(['Totals', fmtNum(t[4]), fmtCurrency(t[5]), fmtCurrency(t[6])])
      }
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
      3: { halign: 'right', cellWidth: summaryHasMisc ? 80 : 90 },
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

  let detailHead: string[][]
  let detailBody: string[][]
  let detailTotalCol: number

  if (detailHasMisc) {
    detailHead = [['Landman', 'Date', 'Prospect', 'Legal', 'Miles', 'Mileage\n0.7250/mi', 'Misc.', 'Total', 'Description']]
    detailTotalCol = 7
    let totalMiles = 0, totalMileageAmt = 0, totalMisc = 0, totalTotal = 0
    detailDataRows.forEach(r => {
      const row = r as unknown[]
      if (detailIsDayrate) {
        if (detailHasDayRate) {
          // Full dayrate: [0]=Landman [1]=Date [2]=Prospect [3]=Legal [4]=Focus [5]=Days [6]=Rate [7]=Labor [8]=Misc [9]=MiscDesc [10]=Total [11]=Desc
          totalMisc += Number(row[8] ?? 0)
          totalTotal += Number(row[10] ?? 0)
        } else {
          // Simplified dayrate: [0]=Landman [1]=Date [2]=Prospect [3]=Legal [4]=Days [5]=Misc [6]=Total [7]=Desc
          totalMisc += Number(row[5] ?? 0)
          totalTotal += Number(row[6] ?? 0)
        }
      } else {
        totalMiles += Number(row[8] ?? 0)
        totalMileageAmt += Number(row[9] ?? 0)
        totalMisc += Number(row[11] ?? 0)
        totalTotal += Number(row[13] ?? 0)
      }
    })
    detailBody = detailDataRows.map(r => {
      const row = r as unknown[]
      if (detailIsDayrate) {
        if (detailHasDayRate) {
          return [String(row[0] ?? ''), formatDate(row[1]), String(row[2] ?? ''), String(row[3] ?? ''), '0.0', '$0.00', fmtCurrency(row[8]), fmtCurrency(row[10]), String(row[11] ?? '')]
        } else {
          return [String(row[0] ?? ''), formatDate(row[1]), String(row[2] ?? ''), String(row[3] ?? ''), '0.0', '$0.00', fmtCurrency(row[5]), fmtCurrency(row[6]), String(row[7] ?? '')]
        }
      }
      return [String(row[0] ?? ''), formatDate(row[1]), String(row[2] ?? ''), String(row[3] ?? ''), fmtNum(row[8], 1), fmtCurrency(row[9]), fmtCurrency(row[11]), fmtCurrency(row[13]), String(row[14] ?? '')]
    })
    detailBody.push(['Totals', '', '', '', fmtNum(totalMiles, 1), fmtCurrency(totalMileageAmt), fmtCurrency(totalMisc), fmtCurrency(totalTotal), ''])
  } else {
    detailHead = [['Landman', 'Date', 'Prospect', 'Legal', 'Miles', 'Mileage\n0.7250/mi', 'Total', 'Description']]
    detailTotalCol = 6
    let totalMiles = 0, totalMileageAmt = 0, totalTotal = 0
    detailDataRows.forEach(r => {
      const row = r as unknown[]
      if (detailIsDayrate) {
        totalTotal += Number(row[7] ?? 0)
      } else {
        totalMiles += Number(row[8] ?? 0)
        totalMileageAmt += Number(row[9] ?? 0)
        totalTotal += Number(row[11] ?? 0)
      }
    })
    detailBody = detailDataRows.map(r => {
      const row = r as unknown[]
      if (detailIsDayrate) {
        return [String(row[0] ?? ''), formatDate(row[1]), String(row[2] ?? ''), String(row[3] ?? ''), '0.0', '$0.00', fmtCurrency(row[7]), String(row[11] ?? '')]
      }
      return [String(row[0] ?? ''), formatDate(row[1]), String(row[2] ?? ''), String(row[3] ?? ''), fmtNum(row[8], 1), fmtCurrency(row[9]), fmtCurrency(row[11]), String(row[12] ?? '')]
    })
    detailBody.push(['Totals', '', '', '', fmtNum(totalMiles, 1), fmtCurrency(totalMileageAmt), fmtCurrency(totalTotal), ''])
  }

  const detailTotalsIndex = detailBody.length - 1

  const detailMiscCols = { 0: { cellWidth: 55 }, 1: { cellWidth: 45 }, 2: { cellWidth: 60 }, 3: { cellWidth: 55 }, 4: { cellWidth: 30, halign: 'center' as const }, 5: { cellWidth: 52, halign: 'right' as const }, 6: { cellWidth: 45, halign: 'right' as const }, 7: { cellWidth: 45, halign: 'right' as const }, 8: { cellWidth: 145 } }
  const detailNoMiscCols = { 0: { cellWidth: 55 }, 1: { cellWidth: 45 }, 2: { cellWidth: 65 }, 3: { cellWidth: 60 }, 4: { cellWidth: 30, halign: 'center' as const }, 5: { cellWidth: 55, halign: 'right' as const }, 6: { cellWidth: 55, halign: 'right' as const }, 7: { cellWidth: 167 } }

  autoTable(doc, {
    startY: 60,
    head: detailHead,
    body: detailBody,
    theme: 'grid',
    styles: { font: 'helvetica', fontSize: 7, textColor: black, fillColor: white, cellPadding: 3, lineColor: black, lineWidth: 0.3, overflow: 'linebreak', valign: 'top' },
    headStyles: { textColor: black, fillColor: headerBg, fontStyle: 'bold', halign: 'center', valign: 'middle' },
    columnStyles: detailHasMisc ? detailMiscCols : detailNoMiscCols,
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
    const afe = (formData.get('afe') as string) || ''
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

      // Find period dynamically
      let period = ''
      for (let i = 0; i < rows.length; i++) {
        if (String((rows[i] as unknown[])?.[0] ?? '').toLowerCase().includes('period')) {
          period = String((rows[i] as unknown[])?.[1] ?? '').trim()
          break
        }
      }
      const fileDateStr = period ? extractPeriodEndDate(period) : filenameDateFormat(rawDate)

      const matchedReceipt = matchReceipt(invoiceNum, receiptData)

      const shortCompany = companyName.replace(/, LLC|, Inc\.?/gi, '').trim()
      const parts = [shortCompany, 'EQT']
      if (manager) parts.push(manager)
      if (county) parts.push(`${county} County`)
      parts.push('Expense Invoice')
      parts.push(fileDateStr)
      const outputName = sanitize(parts.join(' - '))

      try {
        const pdfBuffer = await buildInvoicePdf(
          excelBuffer, companyName, manager, afe, county, invoiceDateOverride, matchedReceipt
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
