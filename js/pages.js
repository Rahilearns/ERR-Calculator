// Page builders for the four calculation flows
import {
  el, numberField, percentField, optionField, dateField,
  monthBoxesField, layeredField, toast, infoIcon, parseDDMMMYYYY, formatDDMMMYYYY,
  openModal, closeModal,
} from './components.js?v=20260603zze';
import { isoToDDMMMYYYY } from './formatting.js?v=20260603zze';
import {
  buildStructuredSchedule, buildCustomizedSchedule,
  buildRateRevisionStructured, computeMetrics,
  computeRevisionMetrics, computeRevisionCustomizedMetrics, buildCofData,
  addMonthsDue,
} from './calculations.js?v=20260603zze';
import { formatMoney, formatPercent } from './formatting.js?v=20260603zze';
import { saveSummary, listSummaries, getMax, saveDraft, loadDraft, clearDraft } from './storage.js?v=20260603zze';
import {
  downloadScheduleAsExcel, downloadSampleAmortization, readUploadedSchedule,
  downloadScheduleAsWord, downloadScheduleAsPDF, downloadVerificationExcel, downloadReportPDF,
  downloadCofSample, readUploadedCof,
  downloadCustomizedRevisionSample, readCustomizedRevisionFile,
} from './excel.js?v=20260603zze';

const IDP_TOOLTIP = 'Tick the months in which the borrower actually pays interest during moratorium. Unticked months accrue and are collected at the next paid month — or rolled into the first installment after moratorium.';

// Cached page state by tab key (also persisted via storage saveDraft)
const tabState = {};

// Render a field's label across two lines: a primary phrase plus a secondary part.
// The secondary part stays inline on desktop (reads as one line, unchanged look) and
// drops onto its own line on mobile (CSS .lbl-line2). Forcing both paired fields to a
// matching two-line height keeps their input boxes aligned on the same row.
function setTwoLineLabel(field, line1, line2) {
  const lbl = field.querySelector('label');
  if (!lbl) return;
  const icon = lbl.querySelector('.info-icon');
  lbl.textContent = line1 + ' ';
  lbl.appendChild(el('span', { class: 'lbl-line2' }, line2));
  if (icon) lbl.appendChild(icon);
}

// Reset button + professional confirmation modal. On confirm, the tab's saved draft is
// cleared and the page re-rendered from scratch (blank fields, no uploads, no results).
// getState (optional) returns the page's serializable input state; it is snapshotted at
// render time — the form is built blank and drafts are restored only afterwards — so a
// click on a still-pristine page does nothing except a gentle note (no modal).
function resetButton(tabKey, rerender, getState = null) {
  const btn = el('button', { class: 'reset-btn', type: 'button' }, 'Reset');
  const pristine = getState ? JSON.stringify(getState()) : null;
  btn.addEventListener('click', () => {
    if (pristine !== null && JSON.stringify(getState()) === pristine) {
      toast('There is nothing to reset — no inputs have been made yet.');
      return;
    }
    const yes = el('button', { class: 'danger-btn', type: 'button' }, 'Yes, Reset');
    const back = el('button', { class: 'ghost-btn modal-ghost', type: 'button' }, 'Go Back');
    yes.addEventListener('click', () => { clearDraft(tabKey); closeModal(); rerender(); });
    back.addEventListener('click', closeModal);
    openModal(el('div', { class: 'confirm-card' },
      el('h3', {}, 'Reset all fields?'),
      el('p', {}, 'This will clear every input on this page — including any uploaded files and calculated results — so you can start a fresh calculation. This action cannot be undone.'),
      el('div', { class: 'confirm-actions' }, back, yes),
    ));
  });
  return btn;
}

// ============================================================
// REGULAR (STRUCTURED) LOAN FACILITY
// ============================================================
export function renderRegularLoan(root) {
  root.innerHTML = '';
  const section = el('div', { class: 'section-card' });

  const loanAmount = numberField({ label: 'Loan Amount', name: 'loanAmount' });
  const offeredRate = percentField({ label: 'Offered Rate', name: 'offeredRate' });

  const moratoriumAvail = optionField({
    label: 'Moratorium Available?', name: 'moratoriumAvail',
    options: [{ label: '— select —', value: '' }, 'No', 'Yes'], value: '',
    onChange: refresh,
  });
  const moratoriumPeriod = numberField({
    label: 'Moratorium Period (Months)', name: 'moratoriumPeriod', integerOnly: true, min: 1,
  });
  moratoriumPeriod.input.addEventListener('input', () => { refresh(); });

  const idpField = monthBoxesField({
    name: 'idpFlags', label: 'Interest During Moratorium Period', tooltip: IDP_TOOLTIP,
    getCount: () => moratoriumPeriod.getValue() || 0, selectAll: true, capitalizable: true,
  });

  const loanTenor = numberField({ label: 'Loan Tenor (Months)', name: 'loanTenor', integerOnly: true, min: 1 });
  const paymentMode = optionField({
    label: 'Payment Mode', name: 'paymentMode',
    options: [{ label: '— select —', value: '' }, 'EMI', 'EQI', 'Equal Principal + Interest (Monthly)', 'Equal Principal + Interest (Quarterly)'],
    value: '',
    onChange: () => refresh(),
  });

  const totalCof = percentField({ label: 'Total Cost of Fund [COF/ISC + OPEX]', name: 'totalCof' });
  setTwoLineLabel(totalCof, 'Total Cost of Fund', '(COF/ISC + OPEX)');

  function securityOptions() {
    const pm = paymentMode.getValue();
    const opts = ['No Funded Security', 'FDR', 'Cash Security'];
    if (pm === 'EMI') opts.push('EMI after Moratorium');
    else if (pm === 'EQI') opts.push('EQI after Moratorium');
    else if (pm) opts.push('Installment');
    return [{ label: '— select —', value: '' }, ...opts];
  }
  const fundedSecurityType = optionField({
    label: 'Funded Security Type', name: 'fundedSecurityType', options: securityOptions(), value: '',
    onChange: refresh,
  });
  setTwoLineLabel(fundedSecurityType, 'Funded Security', 'Type');

  const csAmount = numberField({ label: 'Cash Security / FDR Amount', name: 'csAmount' });
  const csRate = percentField({ label: 'Cash Security / FDR Rate', name: 'csRate' });
  const numInst = numberField({ label: 'Number of Installments', name: 'numInst', integerOnly: true, min: 1 });

  section.appendChild(el('div', { class: 'form-row' }, loanAmount, offeredRate));
  section.appendChild(el('div', { class: 'form-row' }, moratoriumAvail, moratoriumPeriod));
  const moraSection = el('div', { class: 'form-row full hidden' });
  moraSection.appendChild(el('div', { class: 'sub-card' }, idpField));
  section.appendChild(moraSection);
  section.appendChild(el('div', { class: 'form-row' }, loanTenor, paymentMode));
  section.appendChild(el('div', { class: 'form-row' }, totalCof, fundedSecurityType));
  // Security detail row — fields depend on Funded Security Type (rebuilt in refresh()):
  //   FDR / Cash Security       -> Cash Security / FDR Amount + Cash Security / FDR Rate
  //   EMI/EQI after Moratorium  -> Number of Installments + Funded Security Rate
  //   Installment               -> Number of Installments
  const secDetailRow = el('div', { class: 'form-row' });
  section.appendChild(secDetailRow);
  function rebuildSecurityRow() {
    const t = fundedSecurityType.getValue();
    secDetailRow.innerHTML = '';
    if (!t || t === 'No Funded Security') return; // No Funded Security => no detail fields, security = 0
    if (t === 'FDR' || t === 'Cash Security') {
      csRate.setLabel('Cash Security / FDR Rate');
      secDetailRow.append(csAmount, csRate);
    } else if (t === 'EMI after Moratorium' || t === 'EQI after Moratorium') {
      csRate.setLabel('Funded Security Rate');
      secDetailRow.append(numInst, csRate);
    } else {
      // "Installment" funded security (Equal-Principal loans): Number of Installments + its rate
      csRate.setLabel('Funded Security Rate');
      secDetailRow.append(numInst, csRate);
    }
  }

  function refresh() {
    const moraYes = moratoriumAvail.getValue() === 'Yes';
    moratoriumPeriod.classList.toggle('hidden', !moraYes);
    const months = moraYes ? (moratoriumPeriod.getValue() || 0) : 0;
    moraSection.classList.toggle('hidden', months === 0);
    if (months > 0) idpField.refresh();

    paymentMode.setLabel(moraYes ? 'Payment Mode after Moratorium' : 'Payment Mode');
    loanTenor.setLabel(moraYes ? 'Loan Tenor including Moratorium (Months)' : 'Loan Tenor (Months)');

    fundedSecurityType.setOptions(securityOptions());
    rebuildSecurityRow();
  }
  refresh();

  const calcBtn = el('button', { class: 'primary-btn', type: 'button' }, 'Calculate ERR');
  section.appendChild(el('div', { class: 'action-bar' },
    resetButton('regular', () => renderRegularLoan(root), () => collectRegularInputs({
      loanAmount, offeredRate, moratoriumAvail, moratoriumPeriod, idpField,
      loanTenor, paymentMode, totalCof, fundedSecurityType, csAmount, csRate, numInst,
    })), calcBtn));
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
      capFlags: inputs.capFlags,
      cofRate: inputs.totalCof,
      securityAmount: isCs ? (inputs.csAmount || 0) : 0,
      // FDR/Cash use their rate; EMI/EQI after Moratorium use the Funded Security Rate; Installment has none.
      securityRate: inputs.csRate || 0,
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
    idpFlags: f.idpField.getPaidFlags(),
    capFlags: f.idpField.getCapFlags(),
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
  if (!i.loanAmount) return fail('Enter Loan Amount.');
  if (i.offeredRate === null) return fail('Enter Offered Rate.');
  if (!i.moratoriumAvail) return fail('Select whether a moratorium is available.');
  if (i.moratoriumAvail === 'Yes' && !i.moratoriumPeriod) return fail('Enter Moratorium Period.');
  if (!i.loanTenor) return fail('Enter Loan Tenor.');
  if (i.moratoriumAvail === 'Yes' && i.loanTenor <= i.moratoriumPeriod) return fail('Loan Tenor must exceed Moratorium Period.');
  if (!i.paymentMode) return fail('Select a Payment Mode.');
  if (i.totalCof === null) return fail('Enter Total Cost of Fund.');
  if (!i.fundedSecurityType) return fail('Select a Funded Security Type.');
  if (i.fundedSecurityType === 'FDR' || i.fundedSecurityType === 'Cash Security') {
    if (i.csRate === null) return fail('Enter Cash Security / FDR Rate.');
    if (i.csAmount === null) return fail('Enter Cash Security / FDR Amount.');
  }
  if (['EMI after Moratorium', 'EQI after Moratorium', 'Installment'].includes(i.fundedSecurityType) && !i.numInst)
    return fail('Enter Number of Installments.');
  return true;
}
function fail(msg) { toast(msg, 'error'); return false; }

// ============================================================
// CUSTOMIZED LOAN FACILITY
// ============================================================
export function renderCustomizedLoan(root) {
  root.innerHTML = '';
  const section = el('div', { class: 'section-card' });

  const loanAmount = numberField({ label: 'Loan Amount', name: 'loanAmount' });
  const offeredRate = percentField({ label: 'Offered Rate', name: 'offeredRate' });
  const moratoriumAvail = optionField({ label: 'Moratorium Available?', name: 'moratoriumAvail', options: [{ label: '— select —', value: '' }, 'No', 'Yes'], value: '', onChange: refresh });
  const moratoriumPeriod = numberField({ label: 'Moratorium Period (Months)', name: 'moratoriumPeriod', integerOnly: true, min: 1 });
  moratoriumPeriod.input.addEventListener('input', () => { refresh(); refreshLayerOpts(); paymentLayers.applyLayerRules(); });
  const idpField = monthBoxesField({
    name: 'idpFlags', label: 'Interest During Moratorium Period', tooltip: IDP_TOOLTIP,
    getCount: () => moratoriumPeriod.getValue() || 0, selectAll: true, capitalizable: true,
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
  // From and To share the same month list (mora+1 .. tenor) — mirrors the Lending Rate Layers
  // in Rate Revision — Structured. The last layer's To auto-defaults to the final month (the
  // universal cascade engine fills it from getMaturity) and stays editable.
  function toOptions() {
    const tenor = loanTenor.getValue() || 0;
    const mora = moratoriumAvail.getValue() === 'Yes' ? (moratoriumPeriod.getValue() || 0) : 0;
    const arr = [];
    for (let i = mora + 1; i <= tenor; i++) arr.push({ value: String(i), label: `Month ${String(i).padStart(2, '0')}` });
    return arr;
  }

  const paymentLayers = layeredField({
    label: 'Payment Layers',
    name: 'paymentLayers',
    schema: [
      { key: 'fromInstallment', label: 'From Date', type: 'option', options: fromOptions, allowEmpty: true, placeholder: '', width: '0.8fr', readOnly: true },
      { key: 'toInstallment', label: 'To Date', type: 'option', options: toOptions, allowEmpty: true, placeholder: '— select —', width: '0.8fr' },
      { key: 'paymentType', label: 'Payment Type', type: 'option', allowEmpty: true, placeholder: '— select —', options: [
          'Customized Principal (Monthly)', 'Customized Principal (Quarterly)', 'EMI', 'EQI',
          'Equal Principal + Interest (Monthly)', 'Equal Principal + Interest (Quarterly)',
        ], width: '1.2fr' },
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
        const isCustom = !!ptype && ptype.startsWith('Customized Principal');
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
  setTwoLineLabel(totalCof, 'Total Cost of Fund', '(COF/ISC + OPEX)');
  const fundedSecurityType = optionField({
    label: 'Funded Security Type', name: 'fundedSecurityType',
    options: [{ label: '— select —', value: '' }, 'No Funded Security', 'FDR', 'Cash Security', 'EMI after Moratorium', 'EQI after Moratorium'], value: '', onChange: refresh,
  });
  setTwoLineLabel(fundedSecurityType, 'Funded Security', 'Type');
  const csAmount = numberField({ label: 'Cash Security / FDR Amount', name: 'csAmount' });
  const csRate = percentField({ label: 'Cash Security / FDR Rate', name: 'csRate' });
  const numInst = numberField({ label: 'Number of Installments', name: 'numInst', integerOnly: true, min: 1 });

  section.appendChild(el('div', { class: 'form-row' }, loanAmount, offeredRate));
  section.appendChild(el('div', { class: 'form-row' }, moratoriumAvail, moratoriumPeriod));
  const moraSection = el('div', { class: 'form-row full hidden' });
  moraSection.appendChild(el('div', { class: 'sub-card' }, idpField));
  section.appendChild(moraSection);
  section.appendChild(el('div', { class: 'form-row' }, loanTenor));
  section.appendChild(el('div', { class: 'sub-card' }, paymentLayers));
  section.appendChild(el('div', { class: 'form-row' }, totalCof, fundedSecurityType));
  // Security detail row — depends on Funded Security Type (rebuilt in refresh()).
  const secDetailRow = el('div', { class: 'form-row' });
  section.appendChild(secDetailRow);
  function rebuildSecurityRow() {
    const t = fundedSecurityType.getValue();
    secDetailRow.innerHTML = '';
    if (!t || t === 'No Funded Security') return; // No Funded Security => no detail fields, security = 0
    if (t === 'FDR' || t === 'Cash Security') {
      csRate.setLabel('Cash Security / FDR Rate');
      secDetailRow.append(csAmount, csRate);
    } else if (t === 'EMI after Moratorium' || t === 'EQI after Moratorium') {
      csRate.setLabel('Funded Security Rate');
      secDetailRow.append(numInst, csRate);
    } else {
      // "Installment" funded security (Equal-Principal loans): Number of Installments + its rate
      csRate.setLabel('Funded Security Rate');
      secDetailRow.append(numInst, csRate);
    }
  }

  function refresh() {
    const moraYes = moratoriumAvail.getValue() === 'Yes';
    moratoriumPeriod.classList.toggle('hidden', !moraYes);
    const months = moraYes ? (moratoriumPeriod.getValue() || 0) : 0;
    moraSection.classList.toggle('hidden', months === 0);
    if (months > 0) idpField.refresh();
    loanTenor.setLabel(moraYes ? 'Loan Tenor including Moratorium (Months)' : 'Loan Tenor (Months)');
    rebuildSecurityRow();
  }
  refresh();

  const calcBtn = el('button', { class: 'primary-btn', type: 'button' }, 'Calculate ERR');
  section.appendChild(el('div', { class: 'action-bar' },
    resetButton('customized', () => renderCustomizedLoan(root), () => collectCustomizedInputs({
      loanAmount, offeredRate, moratoriumAvail, moratoriumPeriod, idpField,
      loanTenor, paymentLayers, totalCof, fundedSecurityType, csAmount, csRate, numInst,
    })), calcBtn));
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
      capFlags: inputs.capFlags,
      cofRate: inputs.totalCof,
      layers: inputs.paymentLayers,
    };
    const schedule = buildCustomizedSchedule(params);
    const metrics = computeMetrics(schedule, {
      ...params, paymentMode: 'EMI',
      securityAmount: isCs ? (inputs.csAmount || 0) : 0,
      // FDR/Cash use their rate; EMI/EQI after Moratorium use the Funded Security Rate; Installment has none.
      securityRate: inputs.csRate || 0,
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
    idpFlags: f.idpField.getPaidFlags(),
    capFlags: f.idpField.getCapFlags(),
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
  if (!i.fundedSecurityType) return 'Select a Funded Security Type.';
  if (!i.moratoriumAvail) return 'Select whether a moratorium is available.';
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
    if (L.paymentType && L.paymentType.startsWith('Customized Principal') && !L.customPrincipal)
      return `Layer ${k + 1}: enter Custom Principal for the "${L.paymentType}" type.`;
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
  const section = el('div', { class: 'section-card' });

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

  const moratoriumAvail = optionField({ label: 'Moratorium Given at Disbursement?', name: 'moratoriumAvail', options: [{ label: '— select —', value: '' }, 'No', 'Yes'], value: '', onChange: refresh });
  const moratoriumPeriod = numberField({ label: 'Moratorium Period (Months)', name: 'moratoriumPeriod', integerOnly: true, min: 1 });
  moratoriumPeriod.input.addEventListener('input', refresh);

  const idpField = monthBoxesField({
    name: 'idpFlags', label: 'Interest During Moratorium Period', tooltip: IDP_TOOLTIP,
    getCount: () => moratoriumPeriod.getValue() || 0, selectAll: true, capitalizable: true,
  });

  const paymentModality = optionField({
    label: 'Payment Modality', name: 'paymentModality',
    options: [{ label: '— select —', value: '' }, 'EMI', 'EQI', 'Equal Principal + Interest (Monthly)', 'Equal Principal + Interest (Quarterly)'], value: '',
  });
  const tenorMonths = numberField({ label: 'Loan Tenor at Disbursement (Months)', name: 'tenorMonths', integerOnly: true, min: 1 });

  // Actual loan maturity = disbursement + tenor months − 1 day.
  // e.g. 01-Jan-2020 + 60 months → 01-Jan-2025, minus 1 day → 31-Dec-2024.
  // Used only for the last layer's To Date (Lending Rate + Loan Security layers).
  function maturityISO() {
    const d = disbursementDate.getValue();
    const t = tenorMonths.getValue();
    if (!d || !t) return null;
    // Month-t due date per the due-day convention (Feb -> last day for 28th-31st
    // anchors; no rollover into March), then one day back.
    const dt = addMonthsDue(d, t);
    dt.setDate(dt.getDate() - 1);
    return dt.toISOString().slice(0, 10);
  }

  // From-only layers: each layer's rate/security applies from its From Date until the day
  // before the next layer's From (the final layer extends to maturity). No To Date field.
  const rateLayers = layeredField({
    label: 'Lending Rate Layers',
    name: 'rateLayers',
    schema: [
      { key: 'fromDate', label: 'From Date', type: 'date' },
      { key: 'activeRate', label: 'Active Rate', type: 'percent' },
    ],
    addLabel: '+ Add Lending Rate Layer',
    minRows: 2,
    initialRows: 2,
    cascadingFromKey: 'fromDate',
    getAnchor: () => ({ value: disbursementDate.getValue(), kind: 'date' }),
    getMaturity: () => ({ value: maturityISO(), kind: 'date' }),
  });

  const securityLayers = layeredField({
    label: 'Loan Security Layers',
    name: 'securityLayers',
    schema: [
      { key: 'fromDate', label: 'From Date', type: 'date' },
      { key: 'amount', label: 'Amount', type: 'number' },
      { key: 'activeRate', label: 'Active Rate', type: 'percent' },
    ],
    addLabel: '+ Add Security Layer',
    minRows: 1,
    initialRows: 1,
    cascadingFromKey: 'fromDate',
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

  // COF Data Upload — upload button on the left, "Download Sample File" link below it.
  const cofUpload = cofUploadField();
  const cofField = cofUpload.field;

  section.appendChild(el('div', { class: 'form-row' }, initialAmount, disbursementDate));
  section.appendChild(el('div', { class: 'form-row' }, moratoriumAvail, moratoriumPeriod));
  const moraSection = el('div', { class: 'form-row full hidden' });
  moraSection.appendChild(el('div', { class: 'sub-card' }, idpField));
  section.appendChild(moraSection);
  section.appendChild(el('div', { class: 'form-row' }, paymentModality, tenorMonths));
  section.appendChild(el('div', { class: 'sub-card' }, rateLayers));
  section.appendChild(el('div', { class: 'sub-card' }, securityLayers));
  section.appendChild(el('div', { class: 'sub-card' }, cofField));

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
  section.appendChild(el('div', { class: 'action-bar' },
    resetButton('revisionStructured', () => renderRateRevisionStructured(root), () => ({
      ...collectRevisionStructuredInputs({
        initialAmount, disbursementDate, moratoriumAvail, moratoriumPeriod, idpField,
        paymentModality, tenorMonths, rateLayers, securityLayers,
      }),
      cofRows: (cofUpload.getRows() || []).length,
    })), calcBtn));
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
    if (!inputs.moratoriumAvail) return toast('Select whether a moratorium is given at disbursement.', 'error');
    if (!inputs.paymentModality) return toast('Select a Payment Modality.', 'error');
    if (!inputs.rateLayers.length) return toast('Add at least one Lending Rate Layer.', 'error');

    const mat = maturityISO();
    const lastRateFrom = inputs.rateLayers[inputs.rateLayers.length - 1].fromDate;
    if (lastRateFrom && mat && lastRateFrom >= mat) return toast(`The last Lending Rate Layer's From Date (${lastRateFrom}) must be earlier than loan maturity (${mat}).`, 'error');
    if (inputs.securityLayers.length) {
      const lastSecFrom = inputs.securityLayers[inputs.securityLayers.length - 1].fromDate;
      if (lastSecFrom && mat && lastSecFrom >= mat) return toast(`The last Loan Security Layer's From Date (${lastSecFrom}) must be earlier than loan maturity (${mat}).`, 'error');
    }

    // Build COF effective-date data from the uploaded file (0% before first record; cut at maturity).
    const { cofData, warning } = buildCofData(cofUpload.getRows(), inputs.disbursementDate, mat);
    if (warning) toast(warning, 'warn');

    const mora = inputs.moratoriumAvail === 'Yes' ? inputs.moratoriumPeriod : 0;
    const params = {
      initialLoanAmount: inputs.initialAmount,
      disbursementDate: inputs.disbursementDate,
      moratoriumMonths: mora,
      idpFlags: inputs.idpFlags,
      capFlags: inputs.capFlags,
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
    idpFlags: f.idpField.getPaidFlags(),
    capFlags: f.idpField.getCapFlags(),
    paymentModality: f.paymentModality.getValue(),
    tenorMonths: f.tenorMonths.getValue(),
    rateLayers: f.rateLayers.getValue().filter(r => r.fromDate && r.activeRate !== null),
    securityLayers: f.securityLayers.getValue().filter(r => r.fromDate && r.amount),
  };
}

// ============================================================
// RATE REVISION — CUSTOMIZED
// ============================================================
export function renderRateRevisionCustomized(root) {
  root.innerHTML = '';
  const section = el('div', { class: 'section-card' });

  const fileInput = el('input', { type: 'file', accept: '.xlsx,.xls', style: 'display:none' });
  const uploadBtn = el('button', { class: 'secondary-btn', type: 'button' }, '⬆ Upload Excel');
  uploadBtn.addEventListener('click', () => fileInput.click());
  const uploadedLabel = el('span', { class: 'help' }, 'No file uploaded');
  const sampleLink = el('a', { class: 'link-btn', href: '#', role: 'button' }, 'Download Sample File');
  sampleLink.addEventListener('click', (e) => { e.preventDefault(); downloadCustomizedRevisionSample(); });

  let uploadedRows = null;   // amortization schedule rows (Schedule sheet)
  let uploadedCof = null;    // COF records (COF Layers sheet of the same file)
  fileInput.addEventListener('change', async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      const parsed = await readCustomizedRevisionFile(f);
      uploadedRows = parsed.scheduleRows;
      uploadedCof = parsed.cofRows;
      uploadedLabel.textContent = `${f.name} — ${uploadedRows.length} schedule rows, ${uploadedCof.length} COF record(s)`;
      toast('File parsed: amortization schedule and COF layers loaded.', 'success');
    } catch (err) { uploadedRows = null; uploadedCof = null; toast(err.message, 'error'); }
  });

  // Distinguished upload-zone panel (dashed border, icon badge, title + hint).
  const uploadZone = el('div', { class: 'upload-zone' },
    el('div', { class: 'uz-head' },
      el('span', { class: 'uz-icon' }, '⬆'),
      el('div', { class: 'uz-titles' },
        el('div', { class: 'uz-title' }, 'Upload Amortization Schedule and COF Layers'),
        el('div', { class: 'uz-sub' }, 'One Excel file (.xlsx) with the Schedule and COF Layers sheets — use the sample as the template'))),
    el('div', { class: 'uz-actions' }, uploadBtn, fileInput, uploadedLabel),
    el('div', { class: 'uz-sample' }, sampleLink));
  section.appendChild(el('div', { class: 'sub-card' }, uploadZone));

  // From-only security layers (no To Date) — each applies from its From until the next layer's From.
  const securityLayers = layeredField({
    label: 'Loan Security Layers',
    name: 'securityLayers',
    schema: [
      { key: 'fromDate', label: 'From Date', type: 'date' },
      { key: 'amount', label: 'Amount', type: 'number' },
      { key: 'activeRate', label: 'Active Rate', type: 'percent' },
    ],
    minRows: 1, initialRows: 1, addLabel: '+ Add Security Layer',
  });

  section.appendChild(el('div', { class: 'sub-card' }, securityLayers));

  const calcBtn = el('button', { class: 'primary-btn', type: 'button' }, 'Calculate ERR');
  section.appendChild(el('div', { class: 'action-bar' },
    resetButton('revisionCustomized', () => renderRateRevisionCustomized(root), () => ({
      uploadedRows: uploadedRows ? uploadedRows.length : 0,
      uploadedCof: uploadedCof ? uploadedCof.length : 0,
      securityLayers: securityLayers.getValue(),
    })), calcBtn));
  root.appendChild(section);
  const resultsPanel = el('div');
  root.appendChild(resultsPanel);

  calcBtn.addEventListener('click', () => {
    if (!uploadedRows) return toast('Upload the amortization schedule + COF layers file first.', 'error');
    // COF effective-date list from the uploaded file's COF Layers sheet, clipped to the schedule's
    // span (first row = disbursement, last row = maturity), mirroring Rate Revision — Structured.
    const firstDate = uploadedRows[0] && uploadedRows[0].date;
    const lastDate = uploadedRows[uploadedRows.length - 1] && uploadedRows[uploadedRows.length - 1].date;
    const { cofData, warning } = buildCofData(uploadedCof, firstDate, lastDate);
    if (warning) toast(warning, 'warn');
    const inputs = {
      securityLayers: securityLayers.getValue().filter(r => r.fromDate),
      cofRecords: (cofData || []).length,
    };
    const metrics = computeRevisionCustomizedMetrics(uploadedRows, {
      securityLayers: inputs.securityLayers,
      cofData,
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
      // cofData + security layers ride along so the Verify Excel can compute the
      // per-row NIM/ERR (yield to maturity) columns with the same day-count method.
      params: { cofData, securityLayers: inputs.securityLayers }, schedule, metrics,
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

// Header shown before a downloaded schedule (Excel/Word/PDF): only a few key facts.
// Loan Amount (money, 2dp), Interest Rate (%, 2dp), Tenor, Moratorium Period (if moratorium),
// and Payment Modality (labelled "… After Moratorium" when a moratorium exists).
function buildScheduleMeta(ctx) {
  const i = ctx.inputs || {};
  const moraYes = i.moratoriumAvail === 'Yes' || (Number(i.moratoriumPeriod) || 0) > 0;
  const header = {};
  const loanAmt = i.loanAmount != null ? i.loanAmount : i.initialAmount;
  if (loanAmt != null) header['Loan Amount'] = formatMoney(loanAmt);
  if (i.offeredRate != null) header['Interest Rate'] = formatPercent(i.offeredRate);
  const tenor = i.loanTenor != null ? i.loanTenor : i.tenorMonths;
  if (tenor != null) header['Tenor (Months)'] = String(tenor);
  if (moraYes && i.moratoriumPeriod) header['Moratorium Period (Months)'] = String(i.moratoriumPeriod);
  const modality = i.paymentMode || i.paymentModality;
  if (modality) header[moraYes ? 'Payment Modality After Moratorium' : 'Payment Modality'] = modality;
  return { header };
}

// Shared "COF Data Upload" widget (Rate Revision — Structured & Customized).
// Upload button sits on the left; a blue "Download Sample File" link sits below it.
// Returns { field, getRows } where getRows() yields the parsed COF rows (or null).
function cofUploadField(onParsed) {
  let rows = null;
  const fileInput = el('input', { type: 'file', accept: '.xlsx,.xls', style: 'display:none' });
  const uploadBtn = el('button', { class: 'secondary-btn', type: 'button' }, '⬆ Upload COF Data');
  uploadBtn.addEventListener('click', () => fileInput.click());
  const status = el('span', { class: 'help' }, 'No COF file uploaded — interest expense will be 0.');
  const sampleLink = el('a', { class: 'link-btn', href: '#', role: 'button' }, 'Download Sample File');
  sampleLink.addEventListener('click', (e) => { e.preventDefault(); downloadCofSample(); });
  fileInput.addEventListener('change', async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    try {
      rows = await readUploadedCof(f);
      status.textContent = `${f.name} — ${rows.length} COF record(s) loaded.`;
      toast('COF data parsed successfully.', 'success');
      if (onParsed) onParsed(rows);
    } catch (err) { rows = null; status.textContent = 'Upload failed: ' + err.message; toast(err.message, 'error'); }
  });
  // Distinguished upload-zone panel (dashed border, icon badge, title + hint).
  const title = el('div', { class: 'uz-title' }, 'COF Data Upload');
  title.appendChild(infoIcon('Upload the Cost of Fund (COF/ISC + OPEX) effective-date schedule. Each COF rate is effective from its date until the day before the next. Download the sample, edit only the input values, then upload.'));
  const field = el('div', { class: 'upload-zone' },
    el('div', { class: 'uz-head' },
      el('span', { class: 'uz-icon' }, '⬆'),
      el('div', { class: 'uz-titles' }, title,
        el('div', { class: 'uz-sub' }, 'Excel (.xlsx) — the filled-in COF sample workbook'))),
    el('div', { class: 'uz-actions' }, uploadBtn, fileInput, status),
    el('div', { class: 'uz-sample' }, sampleLink));
  return { field, getRows: () => rows };
}

function autoSaveSummary(ctx) {
  return; // Compare feature hidden for now — delete this line to restore auto-save + toasts.
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
  const verifyBtn = el('button', { class: 'verify-btn', type: 'button' }, 'Verify Calculation');
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

  const meta = buildScheduleMeta(ctx);
  const dlExcel = el('button', { class: 'secondary-btn', type: 'button' }, '⬇ Excel');
  const dlWord = el('button', { class: 'secondary-btn', type: 'button' }, '⬇ Word');
  const dlPdf = el('button', { class: 'secondary-btn', type: 'button' }, '⬇ PDF');
  const dlBar = el('div', { class: 'download-box' },
    el('div', { class: 'dl-title' }, 'Download the Schedule'),
    el('div', { class: 'download-buttons' }, dlExcel, dlWord, dlPdf),
  );
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
    // Section replaced (e.g. after Reset): never save from a detached form.
    if (!document.body.contains(sectionEl)) return;
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
    if (fields.idpField && (Array.isArray(data.idpFlags) || Array.isArray(data.capFlags))) {
      const idp = data.idpFlags || [], cap = data.capFlags || [];
      const n = Math.max(idp.length, cap.length);
      fields.idpField.setValue(Array.from({ length: n }, (_, i) => cap[i] ? 2 : (idp[i] ? 1 : 0)));
    }
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
