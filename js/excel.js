// Excel / Word / PDF I/O via CDN libs

function num(v) {
  if (v === null || v === undefined || isNaN(v)) return 0;
  return Math.round(Number(v) * 100) / 100;
}
function fmtM(v) {
  if (v === null || v === undefined || isNaN(v)) return '-';
  return Number(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtP(v) {
  if (v === null || v === undefined || isNaN(v)) return '-';
  return (Number(v) * 100).toFixed(2) + '%';
}

export function downloadScheduleAsExcel(filename, schedule, meta = {}) {
  const aoa = buildScheduleAoA(schedule, meta);
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = aoa[aoa.length - 1].map(() => ({ wch: 16 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Amortization');
  XLSX.writeFile(wb, filename);
}

function buildScheduleAoA(schedule, meta) {
  const aoa = [];
  if (meta.title) aoa.push([meta.title]);
  if (meta.subtitle) aoa.push([meta.subtitle]);
  if (meta.title || meta.subtitle) aoa.push([]);
  if (meta.summary) {
    Object.entries(meta.summary).forEach(([k, v]) => aoa.push([k, v]));
    aoa.push([]);
  }

  const hasDate = schedule.rows && schedule.rows.length && schedule.rows[0].date !== undefined;
  const hasIDP = schedule.rows && schedule.rows.some(r => (r.idpReceivable || 0) > 0);
  const headers = ['Sl.', ...(hasDate ? ['Date'] : []), 'Installment', 'Interest', 'Principal', 'URPA'];
  if (hasIDP) headers.push('Accrued Interest');
  aoa.push(headers);

  for (const r of schedule.rows) {
    const row = [r.sl];
    if (hasDate) row.push(r.date || '');
    row.push(num(r.installment), num(r.interest), num(r.principal), num(r.urpa));
    if (hasIDP) row.push(num(r.idpReceivable || 0));
    aoa.push(row);
  }
  return aoa;
}

// Word
export async function downloadScheduleAsWord(filename, schedule, meta = {}) {
  if (typeof docx === 'undefined') { alert('Word export library not loaded.'); return; }
  const { Document, Packer, Paragraph, Table, TableRow, TableCell, HeadingLevel, WidthType, TextRun } = docx;
  const hasDate = schedule.rows[0]?.date !== undefined;
  const hasIDP = schedule.rows.some(r => (r.idpReceivable || 0) > 0);
  const headers = ['Sl.', ...(hasDate ? ['Date'] : []), 'Installment', 'Interest', 'Principal', 'URPA'];
  if (hasIDP) headers.push('Accrued Interest');

  const headerRow = new TableRow({
    children: headers.map(h => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: h, bold: true })] })] })),
  });
  const dataRows = schedule.rows.map(r => {
    const cells = [String(r.sl)];
    if (hasDate) cells.push(r.date || '');
    cells.push(fmtM(r.installment), fmtM(r.interest), fmtM(r.principal), fmtM(r.urpa));
    if (hasIDP) cells.push(fmtM(r.idpReceivable || 0));
    return new TableRow({ children: cells.map(c => new TableCell({ children: [new Paragraph(c)] })) });
  });

  const children = [];
  if (meta.title) children.push(new Paragraph({ text: meta.title, heading: HeadingLevel.HEADING_1 }));
  if (meta.subtitle) children.push(new Paragraph(meta.subtitle));
  if (meta.summary) {
    Object.entries(meta.summary).forEach(([k, v]) =>
      children.push(new Paragraph({ children: [new TextRun({ text: `${k}: `, bold: true }), new TextRun(String(v))] })));
  }
  children.push(new Paragraph(''));
  children.push(new Paragraph({ text: 'Amortization Schedule', heading: HeadingLevel.HEADING_2 }));
  children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [headerRow, ...dataRows] }));

  const doc = new Document({ sections: [{ properties: {}, children }] });
  const blob = await Packer.toBlob(doc);
  saveBlob(blob, filename);
}

// PDF (schedule)
export function downloadScheduleAsPDF(filename, schedule, meta = {}) {
  const { jsPDF } = window.jspdf || {};
  if (!jsPDF) { alert('PDF export library not loaded.'); return; }
  const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
  let y = 40;
  if (meta.title) { doc.setFontSize(14); doc.text(meta.title, 40, y); y += 20; }
  if (meta.subtitle) { doc.setFontSize(10); doc.text(meta.subtitle, 40, y); y += 16; }
  if (meta.summary) {
    doc.setFontSize(10);
    Object.entries(meta.summary).forEach(([k, v]) => { doc.text(`${k}: ${v}`, 40, y); y += 12; });
    y += 6;
  }
  const hasDate = schedule.rows[0]?.date !== undefined;
  const headers = ['Sl.', ...(hasDate ? ['Date'] : []), 'Installment', 'Interest', 'Principal', 'URPA'];
  const body = schedule.rows.map(r => {
    const row = [String(r.sl)];
    if (hasDate) row.push(r.date || '');
    row.push(fmtM(r.installment), fmtM(r.interest), fmtM(r.principal), fmtM(r.urpa));
    return row;
  });
  doc.autoTable({
    head: [headers], body, startY: y,
    styles: { fontSize: 8, cellPadding: 3 },
    headStyles: { fillColor: [37, 70, 224], textColor: 255 },
    columnStyles: { 0: { halign: 'right' }, 1: { halign: 'center' } },
    margin: { left: 40, right: 40 },
  });
  doc.save(filename);
}

// =====================================================================
// Verify Calculation Excel — replicates "Sample format for verification.xlsx"
// Structure:
//   Sheet "Inputs & Results":
//     A1: "ERR Calculator – Verification"
//     A2: <page title>
//     A4..: inputs as label|value pairs
//     A17..: Results with formulas referencing input rows and schedule cells
//   Sheet "Schedule":
//     A1..G1 headers; data rows starting A2; TOTAL row at end
// =====================================================================
export function downloadVerificationExcel(filename, ctx) {
  const wb = XLSX.utils.book_new();
  const { schedule, inputs, metrics, pageTitle, pageType, params = {} } = ctx;

  // ----- Schedule sheet first (so we know row range)
  const hasDate = schedule.rows[0]?.date !== undefined;
  const cof = params.cofRate || 0;
  const schedHeaders = ['Sl.', 'Installment', 'Interest', 'Principal', 'URPA', 'Int. Expense (URPA*COF/12)', 'Accrued Interest'];
  const schedAoA = [schedHeaders];
  schedule.rows.forEach(r => {
    schedAoA.push([
      r.sl,
      num(r.installment),
      num(r.interest),
      num(r.principal),
      num(r.urpa),
      num((r.urpa || 0) * (cof / 12) * (r.sl === 0 ? 0 : 1)),
      num(r.idpReceivable || 0),
    ]);
  });
  const lastDataRow = schedule.rows.length + 1; // row 1 is header, data starts row 2
  // TOTAL row
  schedAoA.push([
    'TOTAL',
    { f: `SUM(B2:B${lastDataRow})` },
    { f: `SUM(C2:C${lastDataRow})` },
    { f: `SUM(D2:D${lastDataRow})` },
    '',
    { f: `SUM(F2:F${lastDataRow})` },
    '',
  ]);
  const wsSched = XLSX.utils.aoa_to_sheet(schedAoA);
  const totalRow = lastDataRow + 1;
  // Re-apply formulas (aoa_to_sheet may store as text otherwise)
  function setF(ws, addr, f) { ws[addr] = { t: 'n', f }; }
  setF(wsSched, `B${totalRow}`, `SUM(B2:B${lastDataRow})`);
  setF(wsSched, `C${totalRow}`, `SUM(C2:C${lastDataRow})`);
  setF(wsSched, `D${totalRow}`, `SUM(D2:D${lastDataRow})`);
  setF(wsSched, `F${totalRow}`, `SUM(F2:F${lastDataRow})`);
  wsSched['!cols'] = schedHeaders.map(() => ({ wch: 18 }));

  // ----- Inputs & Results sheet
  const inputAoA = [['ERR Calculator – Verification']];
  inputAoA.push([pageTitle]);
  inputAoA.push([]);
  inputAoA.push(['Inputs', 'Values']);

  // Map inputs by page type (mirror sample format for Structured)
  // Pass derived security amount so EMI/EQI/Installment-based security shows the actual computed value.
  const inputsForExport = {
    ...inputs,
    _derived: {
      derivedSecurityAmount: metrics.derivedSecurityAmount,
      derivedSecurityRate: metrics.derivedSecurityRate,
    },
  };
  const inputRowsStart = 5;
  const inputLines = collectInputLinesFor(pageType, inputsForExport);
  inputLines.forEach(([label, value]) => inputAoA.push([label, value]));
  // Tenor years derived row (always present)
  const tenorYearsRow = inputAoA.length + 1; // 1-indexed
  // Indices of key inputs (for formulas)
  const idx = inputIndex(inputLines, inputRowsStart);

  inputAoA.push([]);
  inputAoA.push(['Results', '']);
  const resStart = inputAoA.length + 1; // Results row (header) is at this index, data after

  // Default cell refs (may not exist for revisionCustomized — handled separately)
  const offered = idx.offeredRate, loan = idx.loanAmount;
  const tenor = idx.tenorMonths || idx.loanTenor;
  const cofRow = idx.totalCof;
  const csAmt = idx.csAmount;
  const csRate = idx.csRate;

  // Results rows
  function addRes(label, formulaOrVal) { inputAoA.push([label, formulaOrVal]); }
  const lastSchedRow = lastDataRow; // last data row in schedule
  addRes('Total Interest Received', { f: `Schedule!C${totalRow}` });
  addRes('Total Interest Expense', { f: `Schedule!F${totalRow}` });
  addRes('CS Benefit', csAmt && cofRow && tenor
    ? { f: `B${csAmt}*(B${cofRow}-B${csRate})*B${tenor}/12` }
    : 0);
  addRes('Net Interest Income', { f: `B${resStart + 1} + B${resStart + 3} - B${resStart + 2}` });
  addRes('Avg Portfolio', { f: `AVERAGE(Schedule!E2:E${lastSchedRow})` });
  // NIM = NII / AvgPortfolio / TenorYears
  if (tenor) {
    addRes('NIM', { f: `IF(B${resStart + 5}*(B${tenor}/12)=0, 0, B${resStart + 4} / B${resStart + 5} / (B${tenor}/12))` });
  } else {
    addRes('NIM', 0);
  }
  addRes('Effective Rate (ERR)', cofRow
    ? { f: `B${cofRow}+B${resStart + 6}` }
    : 0);

  const wsInputs = XLSX.utils.aoa_to_sheet(inputAoA);
  wsInputs['!cols'] = [{ wch: 44 }, { wch: 22 }];
  // Format percentage cells
  function setPct(ws, addr) {
    if (ws[addr]) ws[addr].z = '0.00%';
  }
  function setMoney(ws, addr) {
    if (ws[addr]) ws[addr].z = '#,##0.00';
  }
  if (offered) setPct(wsInputs, `B${offered}`);
  if (loan) setMoney(wsInputs, `B${loan}`);
  if (cofRow) setPct(wsInputs, `B${cofRow}`);
  if (csAmt) setMoney(wsInputs, `B${csAmt}`);
  if (csRate) setPct(wsInputs, `B${csRate}`);
  // Result formula cells: re-apply
  function applyResultFormulas() {
    setF(wsInputs, `B${resStart + 1}`, `Schedule!C${totalRow}`);
    setF(wsInputs, `B${resStart + 2}`, `Schedule!F${totalRow}`);
    if (csAmt && cofRow && tenor) setF(wsInputs, `B${resStart + 3}`, `B${csAmt}*(B${cofRow}-B${csRate})*B${tenor}/12`);
    setF(wsInputs, `B${resStart + 4}`, `B${resStart + 1} + B${resStart + 3} - B${resStart + 2}`);
    setF(wsInputs, `B${resStart + 5}`, `AVERAGE(Schedule!E2:E${lastSchedRow})`);
    if (tenor) setF(wsInputs, `B${resStart + 6}`, `IF(B${resStart + 5}*(B${tenor}/12)=0, 0, B${resStart + 4} / B${resStart + 5} / (B${tenor}/12))`);
    if (cofRow) setF(wsInputs, `B${resStart + 7}`, `B${cofRow}+B${resStart + 6}`);
    setMoney(wsInputs, `B${resStart + 1}`);
    setMoney(wsInputs, `B${resStart + 2}`);
    setMoney(wsInputs, `B${resStart + 3}`);
    setMoney(wsInputs, `B${resStart + 4}`);
    setMoney(wsInputs, `B${resStart + 5}`);
    setPct(wsInputs, `B${resStart + 6}`);
    setPct(wsInputs, `B${resStart + 7}`);
  }
  applyResultFormulas();

  XLSX.utils.book_append_sheet(wb, wsInputs, 'Inputs & Results');
  XLSX.utils.book_append_sheet(wb, wsSched, 'Schedule');

  XLSX.writeFile(wb, filename);
}

// If security type is installment-based, the input csAmount is empty —
// substitute the model's derived amount stored on metrics (passed as `inp._derived`).
function securityAmtFor(inp, key) {
  return inp._derived && inp._derived[key] !== undefined ? inp._derived[key] : null;
}

// Build list of [label, value] for the Inputs section based on page type
function collectInputLinesFor(pageType, inp) {
  const yesNo = (v) => (v === 'Yes' || v === true ? 'Yes' : 'No');
  if (pageType === 'regular') {
    return [
      ['Offered Rate', inp.offeredRate ?? 0],
      ['Loan Amount', inp.loanAmount ?? 0],
      ['Moratorium Available?', yesNo(inp.moratoriumAvail)],
      ['Moratorium Period (Months)', inp.moratoriumPeriod ?? 0],
      ['Loan Tenor including Moratorium (Months)', inp.loanTenor ?? 0],
      ['Payment Mode', inp.paymentMode ?? ''],
      ['Total COF (COF/ISC + OPEX)', inp.totalCof ?? 0],
      ['Funded Security Type', inp.fundedSecurityType ?? ''],
      ['Number of Installments (security)', inp.numInst ?? 0],
      ['Cash Security / FDR Amount', securityAmtFor(inp, 'derivedSecurityAmount') ?? (inp.csAmount ?? 0)],
      ['Cash Security / FDR Rate', inp.csRate ?? 0],
    ];
  }
  if (pageType === 'customized') {
    return [
      ['Offered Rate', inp.offeredRate ?? 0],
      ['Loan Amount', inp.loanAmount ?? 0],
      ['Moratorium Available?', yesNo(inp.moratoriumAvail)],
      ['Moratorium Period (Months)', inp.moratoriumPeriod ?? 0],
      ['Loan Tenor including Moratorium (Months)', inp.loanTenor ?? 0],
      ['Payment Layers', (inp.paymentLayers || []).length + ' layer(s)'],
      ['Total COF (COF/ISC + OPEX)', inp.totalCof ?? 0],
      ['Funded Security Type', inp.fundedSecurityType ?? ''],
      ['Number of Installments (security)', inp.numInst ?? 0],
      ['Cash Security / FDR Amount', securityAmtFor(inp, 'derivedSecurityAmount') ?? (inp.csAmount ?? 0)],
      ['Cash Security / FDR Rate', inp.csRate ?? 0],
    ];
  }
  if (pageType === 'revisionStructured') {
    return [
      ['Initial Loan Amount', inp.initialAmount ?? 0],
      ['Disbursement Date', inp.disbursementDate ?? ''],
      ['Moratorium Given at Disbursement?', yesNo(inp.moratoriumAvail)],
      ['Moratorium Period (Months)', inp.moratoriumPeriod ?? 0],
      ['Payment Modality', inp.paymentModality ?? ''],
      ['Loan Tenor including Moratorium at Disbursement (Months)', inp.tenorMonths ?? 0],
      ['Lending Rate Layers', (inp.rateLayers || []).length + ' layer(s)'],
      ['Loan Security Layers', (inp.securityLayers || []).length + ' layer(s)'],
      ['Show NIM Comparison?', yesNo(inp.nimComparison)],
      ['Cost of Fund Layers', (inp.cofLayers || []).length + ' layer(s)'],
    ];
  }
  // revisionCustomized
  return [
    ['Uploaded Schedule', (inp.uploadedRowsCount || 0) + ' rows'],
    ['Loan Security Layers', (inp.securityLayers || []).length + ' layer(s)'],
    ['Show NIM Comparison?', yesNo(inp.nimComparison)],
    ['Cost of Fund Layers', (inp.cofLayers || []).length + ' layer(s)'],
  ];
}

function inputIndex(lines, start) {
  const idx = {};
  lines.forEach(([label], i) => {
    const row = start + i;
    const key = label.toLowerCase();
    if (key === 'offered rate') idx.offeredRate = row;
    else if (key === 'loan amount' || key === 'initial loan amount') idx.loanAmount = row;
    else if (key.startsWith('loan tenor')) idx.loanTenor = row;
    else if (key.startsWith('total cof')) idx.totalCof = row;
    else if (key === 'cash security / fdr amount') idx.csAmount = row;
    else if (key === 'cash security / fdr rate') idx.csRate = row;
    else if (key.startsWith('loan tenor including moratorium at disbursement')) idx.tenorMonths = row;
  });
  return idx;
}

// =====================================================================
// Download Report — PDF with inputs + results + schedule (replication-ready)
// =====================================================================
export function downloadReportPDF(filename, ctx) {
  const { jsPDF } = window.jspdf || {};
  if (!jsPDF) { alert('PDF library not loaded.'); return; }
  const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
  let y = 36;
  doc.setFontSize(16); doc.setFont(undefined, 'bold');
  doc.text(ctx.pageTitle, 40, y); y += 22;
  doc.setFontSize(10); doc.setFont(undefined, 'normal');
  doc.text('Generated: ' + new Date().toLocaleString(), 40, y); y += 16;

  // Inputs
  doc.setFontSize(12); doc.setFont(undefined, 'bold');
  doc.text('Inputs', 40, y); y += 10;
  const inputLines = collectInputLinesFor(ctx.pageType, ctx.inputs);
  // Add layer detail
  const extras = [];
  if (Array.isArray(ctx.inputs.paymentLayers)) {
    ctx.inputs.paymentLayers.forEach((L, i) => {
      extras.push([`Payment Layer ${i + 1}`, `${L.paymentType}: Month ${L.fromInstallment}–${L.toInstallment}` + (L.customPrincipal ? ` @ ${fmtM(L.customPrincipal)}/inst` : '')]);
    });
  }
  if (Array.isArray(ctx.inputs.rateLayers)) {
    ctx.inputs.rateLayers.forEach((L, i) => {
      extras.push([`Rate Layer ${i + 1}`, `${L.fromDate} – ${L.toDate}: ${fmtP(L.activeRate)}`]);
    });
  }
  if (Array.isArray(ctx.inputs.securityLayers) && ctx.inputs.securityLayers.length) {
    ctx.inputs.securityLayers.forEach((L, i) => {
      extras.push([`Security Layer ${i + 1}`, `${L.fromDate} – ${L.toDate}: amt ${fmtM(L.amount)} @ ${fmtP(L.activeRate)}`]);
    });
  }
  if (Array.isArray(ctx.inputs.cofLayers) && ctx.inputs.cofLayers.length) {
    ctx.inputs.cofLayers.forEach((L, i) => {
      extras.push([`COF Layer ${i + 1}`, `${L.fromDate} – ${L.toDate}: ${fmtP(L.cofRate)}`]);
    });
  }
  if (Array.isArray(ctx.inputs.idpFlags) && ctx.inputs.idpFlags.some(Boolean)) {
    const months = ctx.inputs.idpFlags.map((v, i) => v ? i + 1 : null).filter(Boolean);
    extras.push(['Interest-paid Moratorium Months', months.map(m => `M${String(m).padStart(2, '0')}`).join(', ')]);
  }

  // Render input rows
  const formatLine = ([k, v]) => {
    let val = v;
    if (typeof v === 'number') {
      if (k.toLowerCase().includes('rate') || k.toLowerCase().startsWith('total cof')) val = fmtP(v);
      else val = fmtM(v);
    }
    return [k, String(val)];
  };
  doc.autoTable({
    head: [['Field', 'Value']],
    body: [...inputLines.map(formatLine), ...extras],
    startY: y,
    styles: { fontSize: 9, cellPadding: 4 },
    headStyles: { fillColor: [37, 70, 224], textColor: 255 },
    margin: { left: 40, right: 40 },
    columnStyles: { 0: { cellWidth: 240 } },
  });
  y = doc.lastAutoTable.finalY + 18;

  // Results
  doc.setFont(undefined, 'bold'); doc.setFontSize(12);
  doc.text('Results', 40, y); y += 10;
  doc.autoTable({
    head: [['Metric', 'Value']],
    body: [
      ['Effective Rate (ERR)', fmtP(ctx.metrics.effectiveRate)],
      ['NIM', fmtP(ctx.metrics.nim)],
      ['Net Interest Income', fmtM(ctx.metrics.nii)],
      ['Avg Portfolio', fmtM(ctx.metrics.avgPortfolio)],
      ['Total Interest', fmtM(ctx.metrics.totalInterest)],
      ['Tenor (years)', String(Number(ctx.metrics.tenorYears || 0).toFixed(2))],
    ],
    startY: y,
    styles: { fontSize: 9, cellPadding: 4 },
    headStyles: { fillColor: [37, 70, 224], textColor: 255 },
    margin: { left: 40, right: 40 },
    columnStyles: { 0: { cellWidth: 240 } },
  });
  y = doc.lastAutoTable.finalY + 18;

  // Schedule
  doc.setFont(undefined, 'bold'); doc.setFontSize(12);
  doc.text('Amortization Schedule', 40, y); y += 10;
  const hasDate = ctx.schedule.rows[0]?.date !== undefined;
  const headers = ['Sl.', ...(hasDate ? ['Date'] : []), 'Installment', 'Interest', 'Principal', 'URPA'];
  const body = ctx.schedule.rows.map(r => {
    const row = [String(r.sl)];
    if (hasDate) row.push(r.date || '');
    row.push(fmtM(r.installment), fmtM(r.interest), fmtM(r.principal), fmtM(r.urpa));
    return row;
  });
  doc.autoTable({
    head: [headers], body,
    startY: y,
    styles: { fontSize: 8, cellPadding: 3 },
    headStyles: { fillColor: [37, 70, 224], textColor: 255 },
    margin: { left: 40, right: 40 },
  });

  doc.save(filename);
}

export function downloadSampleAmortization() {
  const aoa = [
    ['Date', 'Installment Amount', 'Interest Amount', 'Principal Amount', 'URPA'],
    ['NOTE: First data row = DISBURSEMENT. Enter disbursement date and disbursement amount in URPA. Leave Installment/Interest/Principal blank or 0.', '', '', '', ''],
    ['2024-01-15', 0, 0, 0, 50000000],
    ['2024-02-15', 1112222, 500000, 612222, 49387778],
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 18 }, { wch: 20 }, { wch: 18 }, { wch: 18 }, { wch: 18 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Amortization');
  XLSX.writeFile(wb, 'Sample_Amortization_Schedule.xlsx');
}

export function readUploadedSchedule(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array', cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
        const out = [];
        for (const r of rows) {
          const date = r['Date'] ?? r['date'];
          if (!date) continue;
          let dateStr;
          if (date instanceof Date) dateStr = date.toISOString().slice(0, 10);
          else if (typeof date === 'string') {
            const d = new Date(date);
            if (isNaN(d)) continue;
            dateStr = d.toISOString().slice(0, 10);
          } else continue;
          out.push({
            date: dateStr,
            installmentAmount: Number(r['Installment Amount'] ?? r['installment'] ?? 0) || 0,
            interestAmount: Number(r['Interest Amount'] ?? r['interest'] ?? 0) || 0,
            principalAmount: Number(r['Principal Amount'] ?? r['principal'] ?? 0) || 0,
            urpa: Number(r['URPA'] ?? r['urpa'] ?? 0) || 0,
          });
        }
        if (out.length < 2) return reject(new Error('Uploaded file must contain disbursement row + at least one installment row.'));
        resolve(out);
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error('Failed to read file.'));
    reader.readAsArrayBuffer(file);
  });
}

function saveBlob(blob, filename) {
  if (window.saveAs) return window.saveAs(blob, filename);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 0);
}
