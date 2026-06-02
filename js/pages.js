// Page builders for the four calculation flows
import {
  el, numberField, percentField, optionField, dateField,
  monthBoxesField, layeredField, toast, infoIcon, parseDDMMMYYYY, formatDDMMMYYYY,
} from './components.js?v=20260603a';
import { isoToDDMMMYYYY } from './formatting.js?v=20260603a';
import {
  buildStructuredSchedule, buildCustomizedSchedule,
  buildRateRevisionStructured, computeMetrics,
  computeRevisionMetrics, computeRevisionCustomizedMetrics, buildCofData,
} from './calculations.js?v=20260603a';
import { formatMoney, formatPercent } from './formatting.js?v=20260603a';
import { saveSummary, listSummaries, getMax, saveDraft, loadDraft } from './storage.js?v=20260603a';
import {
  downloadScheduleAsExcel, downloadSampleAmortization, readUploadedSchedule,
  downloadScheduleAsWord, downloadScheduleAsPDF, downloadVerificationExcel, downloadReportPDF,
  downloadCofSample, readUploadedCof,
} from './excel.js?v=20260603a';

const IDP_TOOLTIP = 'Tick the months in which the borrower actually pays interest during moratorium. Unticked months accrue and are collected at the next paid month — or rolled into the first installment after moratorium.';

// Cached page state by tab key (also persisted via storage saveDraft)
const tabState = {};

// ============================================================
// REGULAR (STRUCTURED) LOAN FACILITY
// ============================================================
export function renderRegularLoan(root) {
  root.innerHTML = '';
  root.appendChild(pageTitle('Loan Facilities — Structured'));

  const section = el('div', { class: 'section-card' });
  section.appendChild(el('h2', {}, 'Loan Inputs'));

  const loanAmount = numberField({ label: 'Loan Amount', name: 'loanAmount' });
  const offeredRate = percentField({ label: 'Offered Rate', name: 'offeredRate' });

  const moratoriumAvail = optionField({
    label: 'Moratorium Available?', name: 'moratoriumAvail', options: ['No', 'Yes'], value: 'No',
    onChange: refresh,
  });
  const moratoriumPeriod = numberField({
    label: 'Moratorium Period (Months)', name: 'moratoriumPeriod', integerOnly: true, min: 1,
  });
  moratoriumPeriod.input.addEventListener('input', () => { refresh(); });

  const idpField = monthBoxesField({
    name: 'idpFlags', label: 'Interest During Moratorium Period', tooltip: IDP_TOOLTIP,
    getCount: () => moratoriumPeriod.getValue() || 0, selectAll: true,
  });

  const loanTenor = numberField({ label: 'Loan Tenor (Months)', name: 'loanTenor', integerOnly: true, min: 1 });
  const paymentMode = optionField({
    label: 'Payment Mode', name: 'paymentMode',
    options: ['EMI', 'EQI', 'Equal Principal + Interest (Monthly)', 'Equal Principal + Interest (Quarterly)'],
    value: 'EMI',
    onChange: () => refresh(),
  });

  const totalCof = percentField({ label: 'Total Cost of Fund [COF/ISC + OPEX]', name: 'totalCof' });

  function securityOptions() {
    const pm = paymentMode.getValue();
    const opts = ['FDR', 'Cash Security'];
    if (pm === 'EMI') opts.push('EMI after Moratorium');
    else if (pm === 'EQI') opts.push('EQI after Moratorium');
    else opts.push('Installment');
    return opts;
  }
  const fundedSecurityType = optionField({
    label: 'Funded Security Type', name: 'fundedSecurityType', options: securityOptions(), value: 'FDR',
    onChange: refresh,
  });

  const csAmount = numberField({ label: 'Cash Security / FDR Amount', name: 'csAmount' });
  const csRate = percentField({ label: 'Cash Security / FDR Rate', name: 'csRate' });
  const numInst = numberField({ label: 'Number of Installments', name: 'numInst', integerOnly: true, min: 1 });

  section.appendChild(el('div', { class: 'form-row' }, loanAmount, offeredRate));
  section.appendChild(el('div', { class: 'form-row' }, moratoriumAvail, moratoriumPeriod));
  const moraSection = el('div', { class: 'form-row full hidden' });
  moraSection.appendChild(idpField);
  section.appendChild(moraSection);
  section.appendChild(el('div', { class: 'form-row' }, loanTenor, paymentMode));
  section.appendChild(el('div', { class: 'form-row' }, totalCof, fundedSecurityType));
  // CS Amount and Rate — equal width 2 column row when FDR/CS; Number of Installments takes own full row otherwise
  const secRow = el('div', { class: 'form-row' }, csAmount, csRate);
  const numInstRow = el('div', { class: 'form-row full' }, numInst);
  section.appendChild(secRow);
  section.appendChild(numInstRow);

  function refresh() {
    const moraYes = moratoriumAvail.getValue() === 'Yes';
    moratoriumPeriod.classList.toggle('hidden', !moraYes);
    const months = moraYes ? (moratoriumPeriod.getValue() || 0) : 0;
    moraSection.classList.toggle('hidden', months === 0);
    if (months > 0) idpField.refresh();

    paymentMode.setLabel(moraYes ? 'Payment Mode after Moratorium' : 'Payment Mode');
    loanTenor.setLabel(moraYes ? 'Loan Tenor including Moratorium (Months)' : 'Loan Tenor (Months)');

    fundedSecurityType.setOptions(securityOptions());
    const secType = fundedSecurityType.getValue();
    const showCs = (secType === 'FDR' || secType === 'Cash Security');
    secRow.classList.toggle('hidden', !showCs);
    numInstRow.classList.toggle('hidden', showCs);
  }
  refresh();

  const calcBtn = el('button', { class: 'primary-btn', type: 'button' }, 'Calculate ERR');
  section.appendChild(el('div', { class: 'action-bar' }, calcBtn));
  root.appendChild(section);
  const resultsPanel = el('div');
  root.appendChild(resultsPanel);

  // Restore draft
  restoreDraft('regular', {
    loanAmount, offeredRate, moratoriumAvail, moratoriumPeriod, idpField,
    loanTenor, paymentMode, totalCof, fundedSecurityType, csAmount, csRate, numInst,
  });
  refresh();
  attachDraftAutosave('regular', section, () => collectRegularInputs({
    loanAmount, offeredRate, moratoriumAvail, moratoriumPeriod, idpField,
    loanTenor, paymentMode, totalCof, fundedSecurityType, csAmount, csRate, numInst,
  }));

  calcBtn.addEventListener('click', () => {
    const inputs = collectRegularInputs({
      loanAmount, offeredRate, moratoriumAvail, moratoriumPeriod, idpField,
      loanTenor, paymentMode, totalCof, fundedSecurityType, csAmount, csRate, numInst,
    });
    if (!validateRegular(inputs)) return;
    const moraMonths = inputs.moratoriumAvail === 'Yes' ? inputs.moratoriumPeriod : 0;
    const isCs = inputs.fundedSecurityType === 'FDR' || inputs.fundedSecurityType === 'Cash Security';
    const params = {
      loanAmount: inputs.loanAmount,
      ratePerYear: inputs.offeredRate,
      tenorMonths: inputs.loanTenor,
      paymentMode: inputs.paymentMode,
      moratoriumMonths: moraMonths,
      idpFlags: inputs.idpFlags,
      cofRate: inputs.totalCof,
      securityAmount: isCs ? (inputs.csAmount || 0) : 0,
      securityRate: isCs ? (inputs.csRate || 0) : 0,
      securityKind: inputs.fundedSecurityType,
      numInst: inputs.numInst,
    };
    const schedule = buildStructuredSchedule(params);
    const metrics = computeMetrics(schedule, params);
    const ctx = { pageType: 'regular', pageTitle: 'Loan Facilities — Structured', inputs, params, schedule, metrics };
    autoSaveSummary(ctx);
    renderResults(resultsPanel, ctx);
    resultsPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
}

function collectRegularInputs(f) {
  return {
    offeredRate: f.offeredRate.getValue(),
    loanAmount: f.loanAmount.getValue(),
    moratoriumAvail: f.moratoriumAvail.getValue(),
    moratoriumPeriod: f.moratoriumPeriod.getValue() || 0,
    idpFlags: f.idpField.getValue(),
    loanTenor: f.loanTenor.getValue(),
    paymentMode: f.paymentMode.getValue(),
    totalCof: f.totalCof.getValue(),
    fundedSecurityType: f.fundedSecurityType.getValue(),
    csAmount: f.csAmount.getValue(),
    csRate: f.csRate.getValue(),
    numInst: f.numInst.getValue(),
  };
}

function validateRegular(i) {
  if (i.offeredRate === null) return fail('Enter Offered Rate.');
  if (!i.loanAmount) return fail('Enter Loan Amount.');
  if (!i.loanTenor) return fail('Enter Loan Tenor.');
  if (i.moratoriumAvail === 'Yes' && !i.moratoriumPeriod) return fail('Enter Moratorium Period.');
  if (i.moratoriumAvail === 'Yes' && i.loanTenor <= i.moratoriumPeriod) return fail('Loan Tenor must exceed Moratorium Period.');
  if (i.totalCof === null) return fail('Enter Total Cost of Fund.');
  if (i.fundedSecurityType === 'FDR' || i.fundedSecurityType === 'Cash Security') {
    if (i.csRate === null) return fail('Enter Cash Security / FDR Rate.');
    if (i.csAmount === null) return fail('Enter Cash Security / FDR Amount.');
  }
  if (!(i.fundedSecurityType === 'FDR' || i.fundedSecurityType === 'Cash Security') && !i.numInst)
    return fail('Enter Number of Installments.');
  return true;
}
function fail(msg) { toast(msg, 'error'); return false; }

// ============================================================
// CUSTOMIZED LOAN FACILITY
// ============================================================
export function renderCustomizedLoan(root) {
  root.innerHTML = '';
  root.appendChild(pageTitle('Loan Facilities — Customized'));

  const section = el('div', { class: 'section-card' });
  section.appendChild(el('h2', {}, 'Loan Inputs'));

  const loanAmount = numberField({ label: 'Loan Amount', name: 'loanAmount' });
  const offeredRate = percentField({ label: 'Offered Rate', name: 'offeredRate' });
  const moratoriumAvail = optionField({ label: 'Moratorium Available?', name: 'moratoriumAvail', options: ['No', 'Yes'], value: 'No', onChange: refresh });
  const moratoriumPeriod = numberField({ label: 'Moratorium Period (Months)', name: 'moratoriumPeriod', integerOnly: true, min: 1 });
  moratoriumPeriod.input.addEventListener('input', () => { refresh(); refreshLayerOpts(); paymentLayers.applyLayerRules(); });
  const idpField = monthBoxesField({
    name: 'idpFlags', label: 'Interest During Moratorium Period', tooltip: IDP_TOOLTIP,
    getCount: () => moratoriumPeriod.getValue() || 0, selectAll: true,
  });
  const loanTenor = numberField({ label: 'Loan Tenor (Months)', name: 'loanTenor', integerOnly: true, min: 1 });
  loanTenor.input.addEventListener('input', () => { refreshLayerOpts(); refresh(); paymentLayers.applyLayerRules(); });

  function fromOptions() {
    const tenor = loanTenor.getValue() || 0;
    const mora = moratoriumAvail.getValue() === 'Yes' ? (moratoriumPeriod.getValue() || 0) : 0;
    const arr = [];
    for (let i = mora + 1; i <= tenor; i++) arr.push({ value: String(i), label: `Month ${String(i).padStart(2, '0')}` });
    return arr;
  }
  function toOptions() {
    const tenor = loanTenor.getValue() || 0;
    const mora = moratoriumAvail.getValue() === 'Yes' ? (moratoriumPeriod.getValue() || 0) : 0;
    const arr = [];
    if (tenor) arr.push({ value: 'LAST', label: `Last Month (Month ${String(tenor).padStart(2, '0')})` });
    for (let i = mora + 1; i <= tenor; i++) arr.push({ value: String(i), label: `Month ${String(i).padStart(2, '0')}` });
    return arr;
  }

  const paymentLayers = layeredField({
    label: 'Payment Layers',
    name: 'paymentLayers',
    schema: [
      { key: 'fromInstallment', label: 'From Date', type: 'option', options: fromOptions, allowEmpty: true, placeholder: '— select —', width: '1fr' },
      { key: 'toInstallment', label: 'To Date', type: 'option', options: toOptions, allowEmpty: true, placeholder: '— select —', width: '1fr' },
      { key: 'paymentType', label: 'Payment Type', type: 'option', options: [
          'Customized Principal', 'EMI', 'EQI',
          'Equal Principal + Interest (Monthly)', 'Equal Principal + Interest (Quarterly)',
        ], width: '1.4fr' },
      { key: 'customPrincipal', label: 'Custom Principal', type: 'number', width: '1.2fr' },
    ],
    addLabel: '+ Add Payment Layer',
    minRows: 2,
    initialRows: 2,
    cascadingFromKey: 'fromInstallment',
    cascadingToKey: 'toInstallment',
    getMaturity: () => {
      const tenor = loanTenor.getValue();
      return { value: tenor ? String(tenor) : null, kind: 'month' };
    },
    getAnchor: () => {
      const mora = moratoriumAvail.getValue() === 'Yes' ? (moratoriumPeriod.getValue() || 0) : 0;
      return { value: String(mora + 1), kind: 'month' };
    },
    allowFromEqualTo: true,
    onChange: () => {
      paymentLayers.rows.forEach((row) => {
        const ptype = row.inputs.paymentType.value;
        const cp = row.inputs.customPrincipal;
        const isCustom = ptype === 'Customized Principal';
        cp.disabled = !isCustom;
        cp.style.opacity = isCustom ? '1' : '0.4';
        if (!isCustom) cp.value = '';
      });
    },
  });
  // Toast on the layered field's "cannot add" callback (e.g. last layer already ends at maturity)
  paymentLayers.onCannotAdd = (msg) => toast(msg, 'error');
  function refreshLayerOpts() { paymentLayers.refreshOptions(); }

  const totalCof = percentField({ label: 'Total Cost of Fund [COF/ISC + OPEX]', name: 'totalCof' });
  const fundedSecurityType = optionField({
    label: 'Funded Security Type', name: 'fundedSecurityType',
    options: ['FDR', 'Cash Security', 'EMI after Moratorium', 'EQI after Moratorium'], value: 'FDR', onChange: refresh,
  });
  const csAmount = numberField({ label: 'Cash Security / FDR Amount', name: 'csAmount' });
  const csRate = percentField({ label: 'Cash Security / FDR Rate', name: 'csRate' });
  const numInst = numberField({ label: 'Number of Installments', name: 'numInst', integerOnly: true, min: 1 });

  section.appendChild(el('div', { class: 'form-row' }, loanAmount, offeredRate));
  section.appendChild(el('div', { class: 'form-row' }, moratoriumAvail, moratoriumPeriod));
  const moraSection = el('div', { class: 'form-row full hidden' });
  moraSection.appendChild(idpField);
  section.appendChild(moraSection);
  section.appendChild(el('div', { class: 'form-row' }, loanTenor));
  section.appendChild(el('div', { class: 'form-row full' }, paymentLayers));
  section.appendChild(el('div', { class: 'form-row' }, totalCof, fundedSecurityType));
  const secRow = el('div', { class: 'form-row' }, csAmount, csRate);
  const numInstRow = el('div', { class: 'form-row full' }, numInst);
  section.appendChild(secRow);
  section.appendChild(numInstRow);

  function refresh() {
    const moraYes = moratoriumAvail.getValue() === 'Yes';
    moratoriumPeriod.classList.toggle('hidden', !moraYes);
    const months = moraYes ? (moratoriumPeriod.getValue() || 0) : 0;
    moraSection.classList.toggle('hidden', months === 0);
    if (months > 0) idpField.refresh();
    loanTenor.setLabel(moraYes ? 'Loan Tenor including Moratorium (Months)' : 'Loan Tenor (Months)');
    const secType = fundedSecurityType.getValue();
    const showCs = (secType === 'FDR' || secType === 'Cash Security');
    secRow.classList.toggle('hidden', !showCs);
    numInstRow.classList.toggle('hidden', showCs);
  }
  refresh();

  const calcBtn = el('button', { class: 'primary-btn', type: 'button' }, 'Calculate ERR');
  section.appendChild(el('div', { class: 'action-bar' }, calcBtn));
  root.appendChild(section);
  const resultsPanel = el('div');
  root.appendChild(resultsPanel);

  restoreDraft('customized', {
    loanAmount, offeredRate, moratoriumAvail, moratoriumPeriod, idpField,
    loanTenor, paymentLayers, totalCof, fundedSecurityType, csAmount, csRate, numInst,
  });
  refresh(); refreshLayerOpts();
  attachDraftAutosave('customized', section, () => collectCustomizedInputs({
    loanAmount, offeredRate, moratoriumAvail, moratoriumPeriod, idpField,
    loanTenor, paymentLayers, totalCof, fundedSecurityType, csAmount, csRate, numInst,
  }));

  calcBtn.addEventListener('click', () => {
    const inputs = collectCustomizedInputs({
      loanAmount, offeredRate, moratoriumAvail, moratoriumPeriod, idpField,
      loanTenor, paymentLayers, totalCof, fundedSecurityType, csAmount, csRate, numInst,
    });
    const err = validateCustomized(inputs);
    if (err) return toast(err, 'error');

    const mora = inputs.moratoriumAvail === 'Yes' ? inputs.moratoriumPeriod : 0;
    const isCs = inputs.fundedSecurityType === 'FDR' || inputs.fundedSecurityType === 'Cash Security';
    const params = {
      loanAmount: inputs.loanAmount,
      ratePerYear: inputs.offeredRate,
      tenorMonths: inputs.loanTenor,
      moratoriumMonths: mora,
      idpFlags: inputs.idpFlags,
      cofRate: inputs.totalCof,
      layers: inputs.paymentLayers,
    };
    const schedule = buildCustomizedSchedule(params);
    const metrics = computeMetrics(schedule, {
      ...params, paymentMode: 'EMI',
      securityAmount: isCs ? (inputs.csAmount || 0) : 0,
      securityRate: isCs ? (inputs.csRate || 0) : 0,
      securityKind: inputs.fundedSecurityType,
      numInst: inputs.numInst,
    });
    const ctx = { pageType: 'customized', pageTitle: 'Loan Facilities — Customized', inputs, params, schedule, metrics };
    autoSaveSummary(ctx);
    renderResults(resultsPanel, ctx);
    resultsPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
}

function collectCustomizedInputs(f) {
  const tenor = f.loanTenor.getValue();
  return {
    offeredRate: f.offeredRate.getValue(),
    loanAmount: f.loanAmount.getValue(),
    moratoriumAvail: f.moratoriumAvail.getValue(),
    moratoriumPeriod: f.moratoriumPeriod.getValue() || 0,
    idpFlags: f.idpField.getValue(),
    loanTenor: tenor,
    paymentLayers: f.paymentLayers.getValue().map(r => ({
      fromInstallment: r.fromInstallment ? Number(r.fromInstallment) : null,
      // "LAST" sentinel -> last tenor month
      toInstallment: r.toInstallment === 'LAST' ? Number(tenor || 0) : (r.toInstallment ? Number(r.toInstallment) : null),
      paymentType: r.paymentType,
      customPrincipal: r.customPrincipal,
    })),
    totalCof: f.totalCof.getValue(),
    fundedSecurityType: f.fundedSecurityType.getValue(),
    csAmount: f.csAmount.getValue(),
    csRate: f.csRate.getValue(),
    numInst: f.numInst.getValue(),
  };
}

function validateCustomized(i) {
  if (i.offeredRate === null) return 'Enter Offered Rate.';
  if (!i.loanAmount) return 'Enter Loan Amount.';
  if (!i.loanTenor) return 'Enter Loan Tenor.';
  if (i.totalCof === null) return 'Enter Total Cost of Fund.';
  if (i.moratoriumAvail === 'Yes') {
    if (!i.moratoriumPeriod) return 'Enter Moratorium Period.';
    if (i.loanTenor <= i.moratoriumPeriod) return 'Loan Tenor must exceed Moratorium Period.';
  }
  if (!i.paymentLayers.length) return 'Add at least one Payment Layer.';
  // Layer count cap
  if (i.paymentLayers.length > i.loanTenor) return `Cannot have more than ${i.loanTenor} payment layers (tenor).`;

  const mora = i.moratoriumAvail === 'Yes' ? i.moratoriumPeriod : 0;
  const startMonth = mora + 1;
  const endMonth = i.loanTenor;

  // Per-layer validation
  for (let k = 0; k < i.paymentLayers.length; k++) {
    const L = i.paymentLayers[k];
    if (!L.paymentType) return `Layer ${k + 1}: select a Payment Type.`;
    if (!L.fromInstallment || !L.toInstallment) return `Layer ${k + 1}: select both From and To months.`;
    if (L.toInstallment < L.fromInstallment) return `Layer ${k + 1}: To must be ≥ From.`;
    if (L.fromInstallment < startMonth) return `Layer ${k + 1}: From must be ≥ Month ${String(startMonth).padStart(2, '0')}.`;
    if (L.toInstallment > endMonth) return `Layer ${k + 1}: To must be ≤ Month ${String(endMonth).padStart(2, '0')}.`;
    if (L.paymentType === 'Customized Principal' && !L.customPrincipal)
      return `Layer ${k + 1}: enter Custom Principal for "Customized Principal" type.`;
  }
  // Sort and check for overlaps / gaps
  const sorted = i.paymentLayers.slice().sort((a, b) => a.fromInstallment - b.fromInstallment);
  for (let k = 0; k < sorted.length; k++) {
    if (k === 0 && sorted[k].fromInstallment !== startMonth)
      return `Payment layers must start at Month ${String(startMonth).padStart(2, '0')}. First layer starts at Month ${String(sorted[k].fromInstallment).padStart(2, '0')}.`;
    if (k > 0) {
      const prev = sorted[k - 1];
      const cur = sorted[k];
      if (cur.fromInstallment <= prev.toInstallment)
        return `Layers overlap: Layer ending at Month ${String(prev.toInstallment).padStart(2, '0')} conflicts with layer starting at Month ${String(cur.fromInstallment).padStart(2, '0')}.`;
      if (cur.fromInstallment !== prev.toInstallment + 1)
        return `Gap between layers: nothing covers Month ${String(prev.toInstallment + 1).padStart(2, '0')} to Month ${String(cur.fromInstallment - 1).padStart(2, '0')}.`;
    }
  }
  if (sorted[sorted.length - 1].toInstallment !== endMonth)
    return `Last payment layer must end at Month ${String(endMonth).padStart(2, '0')} (loan tenor). Currently ends at Month ${String(sorted[sorted.length - 1].toInstallment).padStart(2, '0')}.`;
  return null;
}

// ============================================================
// RATE REVISION — STRUCTURED
// ============================================================
export function renderRateRevisionStructured(root) {
  root.innerHTML = '';
  root.appendChild(pageTitle('Rate Revision — Structured'));

  const section = el('div', { class: 'section-card' });
  section.appendChild(el('h2', {}, 'Loan Inputs'));

  const initialAmount = numberField({ label: 'Initial Loan Amount', name: 'initialAmount' });
  const disbursementDate = dateField({
    label: 'Disbursement Date', name: 'disbursementDate',
    disableFn: (d) => {
      const day = d.getDay(); // 5=Fri, 6=Sat
      return day === 5 || day === 6;
    },
    // flatpickr does NOT fire a native 'change' event on date pick — route its
    // onChange here so the layer cascade (first From = disbursement, last To = maturity) updates.
    onChange: () => rerunLayerRules(),
  });

  const moratoriumAvail = optionField({ label: 'Moratorium Given at Disbursement?', name: 'moratoriumAvail', options: ['No', 'Yes'], value: 'No', onChange: refresh });
  const moratoriumPeriod = numberField({ label: 'Moratorium Period (Months)', name: 'moratoriumPeriod', integerOnly: true, min: 1 });
  moratoriumPeriod.input.addEventListener('input', refresh);

  const idpField = monthBoxesField({
    name: 'idpFlags', label: 'Interest During Moratorium Period', tooltip: IDP_TOOLTIP,
    getCount: () => moratoriumPeriod.getValue() || 0, selectAll: true,
  });

  const paymentModality = optionField({
    label: 'Payment Modality', name: 'paymentModality',
    options: ['EMI', 'EQI', 'Equal Principal + Interest (Monthly)', 'Equal Principal + Interest (Quarterly)'], value: 'EMI',
  });
  const tenorMonths = numberField({ label: 'Loan Tenor at Disbursement (Months)', name: 'tenorMonths', integerOnly: true, min: 1 });

  // Actual loan maturity = disbursement + tenor months − 1 day.
  // e.g. 01-Jan-2020 + 60 months → 01-Jan-2025, minus 1 day → 31-Dec-2024.
  // Used only for the last layer's To Date (Lending Rate + Loan Security layers).
  function maturityISO() {
    const d = disbursementDate.getValue();
    const t = tenorMonths.getValue();
    if (!d || !t) return null;
    const dt = new Date(d);
    dt.setMonth(dt.getMonth() + t);
    dt.setDate(dt.getDate() - 1);
    return dt.toISOString().slice(0, 10);
  }

  const rateLayers = layeredField({
    label: 'Lending Rate Layers',
    name: 'rateLayers',
    schema: [
      { key: 'fromDate', label: 'From Date', type: 'date' },
      { key: 'toDate', label: 'To Date', type: 'date' },
      { key: 'activeRate', label: 'Active Rate', type: 'percent' },
    ],
    addLabel: '+ Add Lending Rate Layer',
    minRows: 2,
    initialRows: 2,
    cascadingFromKey: 'fromDate',
    cascadingToKey: 'toDate',
    getAnchor: () => ({ value: disbursementDate.getValue(), kind: 'date' }),
    getMaturity: () => ({ value: maturityISO(), kind: 'date' }),
  });

  const securityLayers = layeredField({
    label: 'Loan Security Layers',
    name: 'securityLayers',
    schema: [
      { key: 'fromDate', label: 'From Date', type: 'date' },
      { key: 'toDate', label: 'To Date', type: 'date' },
      { key: 'amount', label: 'Amount', type: 'number' },
      { key: 'activeRate', label: 'Active Rate', type: 'percent' },
    ],
    addLabel: '+ Add Security Layer',
    minRows: 2,
    initialRows: 2,
    cascadingFromKey: 'fromDate',
    cascadingToKey: 'toDate',
    getAnchor: () => ({ value: disbursementDate.getValue(), kind: 'date' }),
    getMaturity: () => ({ value: maturityISO(), kind: 'date' }),
  });

  // External inputs (disbursement / tenor) feed the cascade engine — re-run on change.
  // Disbursement is wired via dateField's onChange above (flatpickr quirk); tenor is a
  // plain text input so 'input' works.
  function rerunLayerRules() {
    rateLayers.applyLayerRules();
    securityLayers.applyLayerRules();
  }
  tenorMonths.input.addEventListener('input', rerunLayerRules);

  // COF Data Upload — replaces the old NIM-comparison toggle + COF layers.
  let cofUploadedRows = null;
  const cofGetSample = el('button', { class: 'secondary-btn', type: 'button' }, '⬇ Get Sample COF Excel');
  cofGetSample.addEventListener('click', downloadCofSample);
  const cofFileInput = el('input', { type: 'file', accept: '.xlsx,.xls', style: 'display:none' });
  const cofUploadBtn = el('button', { class: 'secondary-btn', type: 'button' }, '⬆ Upload COF Data');
  cofUploadBtn.addEventListener('click', () => cofFileInput.click());
  const cofStatus = el('span', { class: 'help' }, 'No COF file uploaded — interest expense will be 0.');
  cofFileInput.addEventListener('change', async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      cofUploadedRows = await readUploadedCof(f);
      cofStatus.textContent = `${f.name} — ${cofUploadedRows.length} COF record(s) loaded.`;
      toast('COF data parsed successfully.', 'success');
    } catch (err) { cofUploadedRows = null; toast(err.message, 'error'); cofStatus.textContent = 'Upload failed: ' + err.message; }
  });
  const cofField = el('div', { class: 'field' });
  const cofLbl = el('label', {}, 'COF Data Upload');
  cofLbl.appendChild(infoIcon('Upload the Cost of Fund (COF/ISC + OPEX) effective-date schedule. Each COF rate is effective from its date until the day before the next. Download the sample, edit only the input values, then upload.'));
  cofField.appendChild(cofLbl);
  cofField.appendChild(el('div', { style: 'display:flex; gap:10px; align-items:center; flex-wrap:wrap' },
    cofGetSample, cofUploadBtn, cofFileInput, cofStatus));

  section.appendChild(el('div', { class: 'form-row' }, initialAmount, disbursementDate));
  section.appendChild(el('div', { class: 'form-row' }, moratoriumAvail, moratoriumPeriod));
  const moraSection = el('div', { class: 'form-row full hidden' });
  moraSection.appendChild(idpField);
  section.appendChild(moraSection);
  section.appendChild(el('div', { class: 'form-row' }, paymentModality, tenorMonths));
  section.appendChild(el('div', { class: 'form-row full' }, rateLayers));
  section.appendChild(el('div', { class: 'form-row full' }, securityLayers));
  section.appendChild(el('div', { class: 'form-row full' }, cofField));

  function refresh() {
    const moraYes = moratoriumAvail.getValue() === 'Yes';
    moratoriumPeriod.classList.toggle('hidden', !moraYes);
    const months = moraYes ? (moratoriumPeriod.getValue() || 0) : 0;
    moraSection.classList.toggle('hidden', months === 0);
    if (months > 0) idpField.refresh();
    paymentModality.setLabel(moraYes ? 'Payment Modality after Moratorium' : 'Payment Modality');
    tenorMonths.setLabel(moraYes ? 'Loan Tenor including Moratorium at Disbursement (Months)' : 'Loan Tenor at Disbursement (Months)');
  }
  refresh();

  const calcBtn = el('button', { class: 'primary-btn', type: 'button' }, 'Calculate ERR');
  section.appendChild(el('div', { class: 'action-bar' }, calcBtn));
  root.appendChild(section);
  const resultsPanel = el('div');
  root.appendChild(resultsPanel);

  restoreDraft('revisionStructured', {
    initialAmount, disbursementDate, moratoriumAvail, moratoriumPeriod, idpField,
    paymentModality, tenorMonths, rateLayers, securityLayers,
  });
  refresh();
  setTimeout(() => { rateLayers.applyLayerRules(); securityLayers.applyLayerRules(); }, 150);
  attachDraftAutosave('revisionStructured', section, () => collectRevisionStructuredInputs({
    initialAmount, disbursementDate, moratoriumAvail, moratoriumPeriod, idpField,
    paymentModality, tenorMonths, rateLayers, securityLayers,
  }));

  calcBtn.addEventListener('click', () => {
    const inputs = collectRevisionStructuredInputs({
      initialAmount, disbursementDate, moratoriumAvail, moratoriumPeriod, idpField,
      paymentModality, tenorMonths, rateLayers, securityLayers,
    });
    if (!inputs.initialAmount) return toast('Enter Initial Loan Amount.', 'error');
    if (!inputs.disbursementDate) return toast('Enter Disbursement Date.', 'error');
    const dow = new Date(inputs.disbursementDate).getDay();
    if (dow === 5 || dow === 6) return toast('Disbursement Date cannot be Friday or Saturday.', 'error');
    if (!inputs.tenorMonths) return toast('Enter Loan Tenor.', 'error');
    if (!inputs.rateLayers.length) return toast('Add at least one Lending Rate Layer.', 'error');

    const mat = maturityISO();
    const lastTo = inputs.rateLayers[inputs.rateLayers.length - 1].toDate;
    if (lastTo > mat) return toast(`Last Lending Rate Layer "To Date" (${lastTo}) exceeds maturity (${mat}).`, 'error');

    // Build COF effective-date data from the uploaded file (0% before first record; cut at maturity).
    const { cofData, warning } = buildCofData(cofUploadedRows, inputs.disbursementDate, mat);
    if (warning) toast(warning, 'warn', 6000);

    const mora = inputs.moratoriumAvail === 'Yes' ? inputs.moratoriumPeriod : 0;
    const params = {
      initialLoanAmount: inputs.initialAmount,
      disbursementDate: inputs.disbursementDate,
      moratoriumMonths: mora,
      idpFlags: inputs.idpFlags,
      paymentModality: inputs.paymentModality,
      tenorMonths: inputs.tenorMonths,
      rateLayers: inputs.rateLayers,
      securityLayers: inputs.securityLayers,
      cofData,
      maturityDate: mat,
    };
    const schedule = buildRateRevisionStructured(params);
    const metrics = computeRevisionMetrics(schedule);
    const ctx = { pageType: 'revisionStructured', pageTitle: 'Rate Revision — Structured',
      inputs: { ...inputs, cofRecordCount: cofData.length }, params, schedule, metrics };
    autoSaveSummary(ctx);
    renderResults(resultsPanel, ctx);
    resultsPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
}

function collectRevisionStructuredInputs(f) {
  return {
    initialAmount: f.initialAmount.getValue(),
    disbursementDate: f.disbursementDate.getValue(),
    moratoriumAvail: f.moratoriumAvail.getValue(),
    moratoriumPeriod: f.moratoriumPeriod.getValue() || 0,
    idpFlags: f.idpField.getValue(),
    paymentModality: f.paymentModality.getValue(),
    tenorMonths: f.tenorMonths.getValue(),
    rateLayers: f.rateLayers.getValue().filter(r => r.fromDate && r.toDate && r.activeRate !== null),
    securityLayers: f.securityLayers.getValue().filter(r => r.fromDate && r.toDate && r.amount),
  };
}

// ============================================================
// RATE REVISION — CUSTOMIZED
// ============================================================
export function renderRateRevisionCustomized(root) {
  root.innerHTML = '';
  root.appendChild(pageTitle('Rate Revision — Customized'));

  const section = el('div', { class: 'section-card' });
  section.appendChild(el('h2', {}, 'Upload Amortization Schedule'));

  const downloadSample = el('button', { class: 'secondary-btn', type: 'button' }, '⬇ Get Sample Excel');
  downloadSample.addEventListener('click', downloadSampleAmortization);

  const fileInput = el('input', { type: 'file', accept: '.xlsx,.xls', style: 'display:none' });
  const uploadBtn = el('button', { class: 'secondary-btn', type: 'button' }, '⬆ Upload Excel');
  uploadBtn.addEventListener('click', () => fileInput.click());
  const uploadedLabel = el('span', { class: 'help' }, 'No file uploaded');

  let uploadedRows = null;
  fileInput.addEventListener('change', async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      uploadedRows = await readUploadedSchedule(f);
      uploadedLabel.textContent = `${f.name} — ${uploadedRows.length} rows`;
      toast('Excel parsed successfully.', 'success');
    } catch (err) { toast(err.message, 'error'); }
  });

  const uploadField = el('div', { class: 'field' });
  uploadField.appendChild(el('label', {}, 'Sample template / Upload your filled schedule'));
  uploadField.appendChild(el('div', { style: 'display:flex; gap:10px; align-items:center; flex-wrap:wrap' },
    downloadSample, uploadBtn, fileInput, uploadedLabel));
  uploadField.appendChild(el('span', { class: 'help' },
    'Download the sample, fill in Date / Installment / Interest / Principal / URPA (first row = disbursement), then upload.'));
  section.appendChild(el('div', { class: 'form-row full' }, uploadField));

  const securityLayers = layeredField({
    label: 'Loan Security Layers',
    name: 'securityLayers',
    schema: [
      { key: 'fromDate', label: 'From Date', type: 'date' },
      { key: 'toDate', label: 'To Date', type: 'date' },
      { key: 'amount', label: 'Amount', type: 'number' },
      { key: 'activeRate', label: 'Active Rate', type: 'percent' },
    ],
    initialRows: 0, addLabel: '+ Add Security Layer',
  });
  const nimComparison = optionField({ label: 'Want to show NIM margin comparison?', name: 'nimComparison', options: ['No', 'Yes'], value: 'No', onChange: refresh });
  const cofLayers = layeredField({
    label: 'Cost of Fund Layers',
    name: 'cofLayers',
    schema: [
      { key: 'fromDate', label: 'From Date', type: 'date' },
      { key: 'toDate', label: 'To Date', type: 'date' },
      { key: 'cofRate', label: 'COF (COF/ISC + OPEX)', type: 'percent' },
    ],
    initialRows: 0, addLabel: '+ Add COF Layer',
  });

  section.appendChild(el('div', { class: 'form-row full' }, securityLayers));
  section.appendChild(el('div', { class: 'form-row' }, nimComparison));
  const cofRow = el('div', { class: 'form-row full' }, cofLayers);
  section.appendChild(cofRow);

  function refresh() { cofRow.classList.toggle('hidden', nimComparison.getValue() !== 'Yes'); }
  refresh();

  const calcBtn = el('button', { class: 'primary-btn', type: 'button' }, 'Calculate ERR');
  section.appendChild(el('div', { class: 'action-bar' }, calcBtn));
  root.appendChild(section);
  const resultsPanel = el('div');
  root.appendChild(resultsPanel);

  calcBtn.addEventListener('click', () => {
    if (!uploadedRows) return toast('Upload an amortization schedule first.', 'error');
    const inputs = {
      securityLayers: securityLayers.getValue().filter(r => r.fromDate && r.toDate),
      nimComparison: nimComparison.getValue(),
      cofLayers: cofLayers.getValue().filter(r => r.fromDate && r.toDate && r.cofRate !== null),
    };
    const metrics = computeRevisionCustomizedMetrics(uploadedRows, {
      securityLayers: inputs.securityLayers,
      cofLayers: inputs.nimComparison === 'Yes' ? inputs.cofLayers : null,
      hasNimComparison: inputs.nimComparison === 'Yes',
    });
    const schedule = {
      rows: uploadedRows.map((r, i) => ({
        sl: i, date: r.date,
        installment: r.installmentAmount,
        interest: r.interestAmount,
        principal: r.principalAmount,
        urpa: r.urpa,
        interestExpense: 0,
      })),
    };
    const ctx = {
      pageType: 'revisionCustomized', pageTitle: 'Rate Revision — Customized',
      inputs: { ...inputs, uploadedRowsCount: uploadedRows.length },
      params: {}, schedule, metrics,
    };
    autoSaveSummary(ctx);
    renderResults(resultsPanel, ctx);
    resultsPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
}

// ============================================================
// Shared helpers
// ============================================================
function pageTitle(text) {
  return el('div', { class: 'page-title' }, el('h1', {}, text));
}

function autoSaveSummary(ctx) {
  const result = saveSummary({
    pageType: ctx.pageType, pageTitle: ctx.pageTitle,
    inputs: JSON.parse(JSON.stringify(ctx.inputs)),
    metrics: {
      effectiveRate: ctx.metrics.effectiveRate, nim: ctx.metrics.nim, nii: ctx.metrics.nii,
      avgPortfolio: ctx.metrics.avgPortfolio, totalInterest: ctx.metrics.totalInterest, tenorYears: ctx.metrics.tenorYears,
    },
    label: ctx.pageTitle + ' @ ' + new Date().toLocaleString(),
  });
  if (result.saved) {
    toast(`Saved (${listSummaries().length}/${getMax()}).`, 'success');
    window.dispatchEvent(new CustomEvent('summary-saved'));
  } else if (result.reason === 'full') {
    toast(`Storage is full (${getMax()} max). Remove a saved summary to add this one.`, 'warn');
  }
  // duplicate: silently no-op
}

function renderResults(panel, ctx) {
  panel.innerHTML = '';
  const card = el('div', { class: 'section-card' });
  card.appendChild(el('h2', {}, 'Results'));

  const m = ctx.metrics;
  const grid = el('div', { class: 'results-grid' });
  const metric = (label, value, primary = false) => el('div', { class: 'metric-card' + (primary ? ' primary' : '') },
    el('div', { class: 'label' }, label),
    el('div', { class: 'value' }, value));
  grid.appendChild(metric('Effective Rate (ERR)', formatPercent(m.effectiveRate), true));
  grid.appendChild(metric('NIM', formatPercent(m.nim)));
  grid.appendChild(metric('Net Interest Income', formatMoney(m.nii)));
  // If model derived a security amount (EMI/EQI after Moratorium / Installment), show it for transparency
  const sk = String(ctx.inputs.fundedSecurityType || '');
  if ((sk.startsWith('EMI') || sk.startsWith('EQI') || sk === 'Installment') && m.derivedSecurityAmount > 0) {
    grid.appendChild(metric(
      `Security Amount (${sk} × ${ctx.inputs.numInst || 1})`,
      formatMoney(m.derivedSecurityAmount)
    ));
  }
  card.appendChild(grid);

  // Top-right actions: Download Report (left), Verify Calculation (right)
  const baseFname = ctx.pageTitle.replace(/[^a-z0-9]+/gi, '_') + '_' + new Date().toISOString().slice(0, 10);
  const reportBtn = el('button', { class: 'secondary-btn', type: 'button' }, '📄 Download Report');
  const verifyBtn = el('button', { class: 'primary-btn', type: 'button' }, '🧮 Verify Calculation');
  reportBtn.addEventListener('click', () => downloadReportPDF(baseFname + '_Report.pdf', ctx));
  verifyBtn.addEventListener('click', () => downloadVerificationExcel(baseFname + '_Verification.xlsx', ctx));
  card.appendChild(el('div', { class: 'results-actions' }, reportBtn, verifyBtn));

  panel.appendChild(card);

  // Amortization schedule + totals + downloads
  const tableCard = el('div', { class: 'section-card' });
  tableCard.appendChild(el('h2', {}, 'Amortization Schedule'));
  const wrap = el('div', { class: 'table-wrap' });
  const table = el('table', { class: 'schedule' });
  const hasDate = ctx.schedule.rows[0]?.date !== undefined;
  const hasIDP = ctx.schedule.rows.some(r => (r.idpReceivable || 0) > 0);
  const headers = ['Sl.', ...(hasDate ? ['Date'] : []), 'Installment', 'Interest', 'Principal', 'URPA'];
  if (hasIDP) headers.push('Accrued Interest');
  const thead = el('thead', {}, el('tr', {}, ...headers.map(h => el('th', {}, h))));
  const tbody = el('tbody');
  let totPay = 0, totInt = 0, totPrin = 0;
  ctx.schedule.rows.forEach((r) => {
    totPay += r.installment || 0;
    totInt += r.interest || 0;
    totPrin += r.principal || 0;
    const cells = [String(r.sl)];
    if (hasDate) cells.push(r.date || '');
    cells.push(formatMoney(r.installment), formatMoney(r.interest), formatMoney(r.principal), formatMoney(r.urpa));
    if (hasIDP) cells.push(formatMoney(r.idpReceivable || 0));
    tbody.appendChild(el('tr', {}, ...cells.map(c => el('td', {}, c))));
  });
  const totalCells = ['TOTAL'];
  if (hasDate) totalCells.push('');
  totalCells.push(formatMoney(totPay), formatMoney(totInt), formatMoney(totPrin), '');
  if (hasIDP) totalCells.push('');
  tbody.appendChild(el('tr', { class: 'totals-row' }, ...totalCells.map(c => el('td', {}, c))));
  table.appendChild(thead);
  table.appendChild(tbody);
  wrap.appendChild(table);
  tableCard.appendChild(wrap);

  const summaryRow = el('div', { class: 'totals-summary' },
    el('div', { class: 'totals-card' },
      el('div', { class: 'label' }, 'Total Principal Paid'),
      el('div', { class: 'value' }, formatMoney(totPrin))),
    el('div', { class: 'totals-card' },
      el('div', { class: 'label' }, 'Total Interest Paid'),
      el('div', { class: 'value' }, formatMoney(totInt))),
    el('div', { class: 'totals-card' },
      el('div', { class: 'label' }, 'Total Payment'),
      el('div', { class: 'value' }, formatMoney(totPay))),
  );
  tableCard.appendChild(summaryRow);

  const meta = {
    title: ctx.pageTitle, subtitle: 'Generated ' + new Date().toLocaleString(),
    summary: {
      'Effective Rate (ERR)': formatPercent(m.effectiveRate),
      'NIM': formatPercent(m.nim),
      'Net Interest Income': formatMoney(m.nii),
      'Total Principal Paid': formatMoney(totPrin),
      'Total Interest Paid': formatMoney(totInt),
      'Total Payment': formatMoney(totPay),
    },
  };
  const dlLabel = el('span', { class: 'dl-label' },
    'Download the Schedule',
    el('span', { class: 'dl-arrow' }, '⟶'));
  const dlExcel = el('button', { class: 'secondary-btn', type: 'button' }, '⬇ Excel');
  const dlWord = el('button', { class: 'secondary-btn', type: 'button' }, '⬇ Word');
  const dlPdf = el('button', { class: 'secondary-btn', type: 'button' }, '⬇ PDF');
  const dlBar = el('div', { class: 'download-bar' }, dlLabel, dlExcel, dlWord, dlPdf);
  tableCard.appendChild(dlBar);
  dlExcel.addEventListener('click', () => downloadScheduleAsExcel(baseFname + '.xlsx', ctx.schedule, meta));
  dlWord.addEventListener('click', () => downloadScheduleAsWord(baseFname + '.docx', ctx.schedule, meta));
  dlPdf.addEventListener('click', () => downloadScheduleAsPDF(baseFname + '.pdf', ctx.schedule, meta));

  panel.appendChild(tableCard);
}

// ============================================================
// Draft preservation
// ============================================================
function attachDraftAutosave(tabKey, sectionEl, collector) {
  let last = '';
  let timer = null;
  function doSave() {
    try {
      const data = collector();
      const ser = JSON.stringify(data);
      if (ser !== last) {
        last = ser;
        saveDraft(tabKey, data);
      }
    } catch {}
  }
  function schedule() {
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => { timer = null; doSave(); }, 300);
  }
  function flush() {
    if (timer) { clearTimeout(timer); timer = null; }
    doSave();
  }
  // Debounce keystrokes; force-save on navigation
  sectionEl.addEventListener('input', schedule);
  sectionEl.addEventListener('change', schedule);
  window.addEventListener('beforeunload', flush);
  setTimeout(doSave, 200);
}

function restoreDraft(tabKey, fields) {
  const data = loadDraft(tabKey);
  if (!data) return;
  // Apply each known field
  try {
    if (fields.loanAmount && data.loanAmount !== undefined) fields.loanAmount.setValue(data.loanAmount);
    if (fields.offeredRate && data.offeredRate !== undefined) fields.offeredRate.setValue(data.offeredRate);
    if (fields.initialAmount && data.initialAmount !== undefined) fields.initialAmount.setValue(data.initialAmount);
    if (fields.disbursementDate && data.disbursementDate) fields.disbursementDate.setValue(data.disbursementDate);
    if (fields.moratoriumAvail && data.moratoriumAvail) fields.moratoriumAvail.setValue(data.moratoriumAvail);
    if (fields.moratoriumPeriod && data.moratoriumPeriod) fields.moratoriumPeriod.setValue(data.moratoriumPeriod);
    if (fields.idpField && Array.isArray(data.idpFlags)) fields.idpField.setValue(data.idpFlags);
    if (fields.loanTenor && data.loanTenor) fields.loanTenor.setValue(data.loanTenor);
    if (fields.tenorMonths && data.tenorMonths) fields.tenorMonths.setValue(data.tenorMonths);
    if (fields.paymentMode && data.paymentMode) fields.paymentMode.setValue(data.paymentMode);
    if (fields.paymentModality && data.paymentModality) fields.paymentModality.setValue(data.paymentModality);
    if (fields.totalCof && data.totalCof !== undefined) fields.totalCof.setValue(data.totalCof);
    if (fields.fundedSecurityType && data.fundedSecurityType) fields.fundedSecurityType.setValue(data.fundedSecurityType);
    if (fields.csAmount && data.csAmount !== undefined) fields.csAmount.setValue(data.csAmount);
    if (fields.csRate && data.csRate !== undefined) fields.csRate.setValue(data.csRate);
    if (fields.numInst && data.numInst) fields.numInst.setValue(data.numInst);
    if (fields.nimComparison && data.nimComparison) fields.nimComparison.setValue(data.nimComparison);
    if (fields.paymentLayers && Array.isArray(data.paymentLayers) && data.paymentLayers.length) {
      // Convert ISO->display etc. handled by layered field's number/percent setters internally
      fields.paymentLayers.setValue(data.paymentLayers.map(L => ({
        fromInstallment: L.fromInstallment != null ? String(L.fromInstallment) : '',
        toInstallment: L.toInstallment != null ? String(L.toInstallment) : '',
        paymentType: L.paymentType,
        customPrincipal: L.customPrincipal,
      })));
    }
    if (fields.rateLayers && Array.isArray(data.rateLayers) && data.rateLayers.length) {
      fields.rateLayers.setValue(data.rateLayers.map(L => ({
        fromDate: isoToDDMMMYYYY(L.fromDate),
        toDate: isoToDDMMMYYYY(L.toDate),
        activeRate: L.activeRate,
      })));
    }
    if (fields.securityLayers && Array.isArray(data.securityLayers) && data.securityLayers.length) {
      fields.securityLayers.setValue(data.securityLayers.map(L => ({
        fromDate: isoToDDMMMYYYY(L.fromDate),
        toDate: isoToDDMMMYYYY(L.toDate),
        amount: L.amount,
        activeRate: L.activeRate,
      })));
    }
    if (fields.cofLayers && Array.isArray(data.cofLayers) && data.cofLayers.length) {
      fields.cofLayers.setValue(data.cofLayers.map(L => ({
        fromDate: isoToDDMMMYYYY(L.fromDate),
        toDate: isoToDDMMMYYYY(L.toDate),
        cofRate: L.cofRate,
      })));
    }
  } catch (err) {
    console.warn('Draft restore failed', err);
  }
}
