// Core financial math + amortization + ERR engine
// All rates passed as decimals (0.12 = 12%). Periodic = monthly unless stated.

export function PMT(rate, nper, pv, fv = 0, type = 0) {
  if (rate === 0) return -(pv + fv) / nper;
  const pvif = Math.pow(1 + rate, nper);
  return -rate * (pv * pvif + fv) / ((pvif - 1) * (1 + rate * type));
}

export function IRR(cashflows, guess = 0.1) {
  let rate = guess;
  for (let i = 0; i < 80; i++) {
    let npv = 0, dnpv = 0;
    for (let t = 0; t < cashflows.length; t++) {
      const denom = Math.pow(1 + rate, t);
      npv += cashflows[t] / denom;
      dnpv += -t * cashflows[t] / (denom * (1 + rate));
    }
    if (Math.abs(npv) < 1e-9) return rate;
    if (dnpv === 0) break;
    const next = rate - npv / dnpv;
    if (!isFinite(next)) break;
    if (Math.abs(next - rate) < 1e-12) return next;
    rate = next;
  }
  let lo = -0.999, hi = 10;
  const npvAt = (r) => cashflows.reduce((s, c, t) => s + c / Math.pow(1 + r, t), 0);
  const fl = npvAt(lo), fh = npvAt(hi);
  if (fl * fh > 0) return rate;
  for (let i = 0; i < 200; i++) {
    const mid = (lo + hi) / 2;
    const fm = npvAt(mid);
    if (Math.abs(fm) < 1e-9) return mid;
    if (fl * fm < 0) hi = mid; else lo = mid;
  }
  return (lo + hi) / 2;
}

function periodsPerYear(mode) {
  if (mode === 'EQI' || mode === 'Equal Principal + Interest (Quarterly)') return 4;
  return 12;
}

// Apply the spec: Int Expense at end of month N = URPA at start of month N * COF/12.
// URPA at start of month N is the post-payment URPA from row N-1.
// Row 0 (disbursement) carries no expense.
// `cofForRow(i)` lets callers vary COF per row (e.g. Rate Revision); default uses constant monthlyCof.
function applyIntExpenseAccrual(rows, monthlyCof, cofForRow = null) {
  for (let i = 0; i < rows.length; i++) {
    if (i === 0) { rows[i].interestExpense = 0; continue; }
    const cof = cofForRow ? cofForRow(i) : monthlyCof;
    rows[i].interestExpense = (rows[i - 1].urpa || 0) * cof;
  }
}

// Number of quarterly payment months in [from..to] (months divisible by 3)
function countQuarterlyMonths(from, to) {
  let n = 0;
  for (let m = from; m <= to; m++) if (m % 3 === 0) n++;
  return n;
}

// ============================================================
// Schedule generation
// ============================================================
// Moratorium logic (per user spec):
//   - In months marked IDP-paid: that month's interest is paid; any accrued interest from prior unpaid months is ALSO paid
//   - In months not marked: interest accrues, no payment
//   - If accrued interest remains at the end of moratorium, it is added to the FIRST regular installment
//
// Row shape: { sl, installment, interest, principal, urpa, interestExpense, idpReceivable, date? }
// ============================================================

export function buildStructuredSchedule(p) {
  const {
    loanAmount, ratePerYear, tenorMonths,
    paymentMode,
    moratoriumMonths = 0,
    idpFlags = [],
    cofRate = 0,
  } = p;

  const ppy = periodsPerYear(paymentMode);
  const regularTenor = tenorMonths - moratoriumMonths;
  const regularPeriods = (ppy === 12) ? regularTenor : Math.round(regularTenor / 3);

  const ratePerPeriod = ratePerYear / ppy;
  const monthlyRate = ratePerYear / 12;
  const monthlyCof = cofRate / 12;

  const rows = [];
  let urpa = loanAmount;
  // Row 0 = disbursement; no interest expense
  rows.push({ sl: 0, installment: 0, interest: 0, principal: 0, urpa, interestExpense: 0, idpReceivable: 0 });

  // Moratorium accrual
  let accruedReceivable = 0;
  for (let m = 1; m <= moratoriumMonths; m++) {
    const interest = urpa * monthlyRate;
    const paid = !!idpFlags[m - 1];
    let installment = 0;
    if (paid) {
      installment = interest + accruedReceivable;
      accruedReceivable = 0;
    } else {
      accruedReceivable += interest;
    }
    rows.push({
      sl: m, installment, interest, principal: 0, urpa,
      interestExpense: urpa * monthlyCof, idpReceivable: accruedReceivable,
    });
  }

  // Regular installments
  if (regularPeriods <= 0) {
    applyIntExpenseAccrual(rows, monthlyCof);
    return { rows, accruedReceivable };
  }

  // Compute the "base" regular installment without accrued-receivable add-on (used as the EMI/installment-size for security calc)
  let baseInstallment = 0;
  if (paymentMode === 'EMI' || paymentMode === 'EQI') {
    const pmt = PMT(ratePerPeriod, regularPeriods, -urpa);
    baseInstallment = pmt;
    let paymentCounter = 0;
    for (let m = moratoriumMonths + 1; m <= tenorMonths; m++) {
      let installment = 0, interest = 0, principal = 0;
      const monthsSinceMora = m - moratoriumMonths;
      const isPaymentMonth = (ppy === 12) || (monthsSinceMora % 3 === 0);
      if (isPaymentMonth) {
        interest = urpa * ratePerPeriod;
        installment = pmt;
        if (paymentCounter === regularPeriods - 1) {
          principal = urpa;
          installment = principal + interest + accruedReceivable;
          accruedReceivable = 0;
        } else {
          principal = installment - interest;
        }
        if (paymentCounter === 0 && accruedReceivable > 0) {
          installment += accruedReceivable;
          accruedReceivable = 0;
        }
        urpa = Math.max(0, urpa - principal);
        paymentCounter++;
      }
      rows.push({
        sl: m, installment, interest, principal, urpa,
        interestExpense: urpa * monthlyCof, idpReceivable: accruedReceivable,
        paymentType: paymentMode, isPaymentMonth,
      });
    }
  } else {
    // Equal Principal + Interest
    const principalPer = urpa / regularPeriods;
    let paymentCounter = 0;
    // Base installment for security size = principal-per + interest on initial URPA at first payment
    const firstInterest = urpa * ratePerPeriod;
    baseInstallment = principalPer + firstInterest;
    for (let m = moratoriumMonths + 1; m <= tenorMonths; m++) {
      let installment = 0, interest = 0, principal = 0;
      const monthsSinceMora = m - moratoriumMonths;
      const isPaymentMonth = (ppy === 12) || (monthsSinceMora % 3 === 0);
      if (isPaymentMonth) {
        interest = urpa * ratePerPeriod;
        principal = principalPer;
        installment = principal + interest;
        if (paymentCounter === regularPeriods - 1) {
          principal = urpa;
          installment = principal + interest + accruedReceivable;
          accruedReceivable = 0;
        }
        if (paymentCounter === 0 && accruedReceivable > 0) {
          installment += accruedReceivable;
          accruedReceivable = 0;
        }
        urpa = Math.max(0, urpa - principal);
        paymentCounter++;
      }
      rows.push({
        sl: m, installment, interest, principal, urpa,
        interestExpense: urpa * monthlyCof, idpReceivable: accruedReceivable,
        paymentType: paymentMode, isPaymentMonth,
      });
    }
  }
  applyIntExpenseAccrual(rows, monthlyCof);
  return { rows, accruedReceivable, baseInstallment };
}

// Customized Loan
export function buildCustomizedSchedule(p) {
  const {
    loanAmount, ratePerYear, tenorMonths,
    moratoriumMonths = 0,
    idpFlags = [],
    cofRate = 0,
    layers,
  } = p;

  const monthlyRate = ratePerYear / 12;
  const monthlyCof = cofRate / 12;

  const rows = [];
  let urpa = loanAmount;
  rows.push({ sl: 0, installment: 0, interest: 0, principal: 0, urpa, interestExpense: 0, idpReceivable: 0 });

  let accruedReceivable = 0;
  for (let m = 1; m <= moratoriumMonths; m++) {
    const interest = urpa * monthlyRate;
    const paid = !!idpFlags[m - 1];
    let installment = 0;
    if (paid) {
      installment = interest + accruedReceivable;
      accruedReceivable = 0;
    } else {
      accruedReceivable += interest;
    }
    rows.push({
      sl: m, installment, interest, principal: 0, urpa,
      interestExpense: urpa * monthlyCof, idpReceivable: accruedReceivable,
    });
  }

  // Layers' fromInstallment/toInstallment are ABSOLUTE month numbers (Month NN values)
  // E.g. if moratorium=6 and tenor=36, layers have from=7, to=36
  const sorted = layers.slice().sort((a, b) => (a.fromInstallment ?? 0) - (b.fromInstallment ?? 0));

  // Track per-layer base installment for security-size lookup
  const layerInstallments = {}; // { 'EMI': base_pmt, 'EQI': base_pmt, ... }

  for (let li = 0; li < sorted.length; li++) {
    const L = sorted[li];
    const from = L.fromInstallment;
    const to = L.toInstallment;
    const count = to - from + 1;
    const isLastLayer = (li === sorted.length - 1);

    let pmt = 0;
    let ppy = 12;
    if (L.paymentType === 'EMI' || L.paymentType === 'Equal Principal + Interest (Monthly)') ppy = 12;
    else if (L.paymentType === 'EQI' || L.paymentType === 'Equal Principal + Interest (Quarterly)') ppy = 4;

    // Equal Principal + Interest layers: the per-period principal is CONSTANT and sized over
    // the periods remaining to MATURITY (not just the layer's own span). So a layer that ends
    // before maturity only partially amortizes, leaving a balance for the next layer.
    //   principal = (balance at layer start) / periods-to-maturity
    //   periods-to-maturity (monthly)   = tenor - from + 1
    //   periods-to-maturity (quarterly) = (tenor - from + 1) / 3
    const layerStartBalance = urpa;
    let epiConstPrincipal = 0;
    if (L.paymentType === 'Equal Principal + Interest (Monthly)') {
      const periodsToMaturity = tenorMonths - from + 1;
      epiConstPrincipal = periodsToMaturity > 0 ? layerStartBalance / periodsToMaturity : layerStartBalance;
    } else if (L.paymentType === 'Equal Principal + Interest (Quarterly)') {
      const periodsToMaturity = (tenorMonths - from + 1) / 3;
      epiConstPrincipal = periodsToMaturity > 0 ? layerStartBalance / periodsToMaturity : layerStartBalance;
    }

    if (L.paymentType === 'EMI') {
      pmt = PMT(ratePerYear / 12, count, -urpa);
      if (!layerInstallments['EMI']) layerInstallments['EMI'] = pmt;
    } else if (L.paymentType === 'EQI') {
      // Count payment months in [from..to] where m%3===0
      const qCount = countQuarterlyMonths(from, to);
      pmt = PMT(ratePerYear / 4, Math.max(1, qCount), -urpa);
      if (!layerInstallments['EQI']) layerInstallments['EQI'] = pmt;
    } else if (L.paymentType === 'Equal Principal + Interest (Monthly)' && !layerInstallments['Installment']) {
      layerInstallments['Installment'] = (urpa / count) + urpa * monthlyRate;
    } else if (L.paymentType === 'Equal Principal + Interest (Quarterly)' && !layerInstallments['Installment']) {
      const qPeriods = Math.max(1, countQuarterlyMonths(from, to));
      layerInstallments['Installment'] = (urpa / qPeriods) + urpa * (ratePerYear / 4);
    } else if (L.paymentType === 'Customized Principal' && !layerInstallments['Customized']) {
      layerInstallments['Customized'] = (L.customPrincipal || 0) + urpa * monthlyRate;
    }

    let paymentCounter = 0;
    for (let m = from; m <= to; m++) {
      let installment = 0, interest = 0, principal = 0;
      const monthsInLayer = m - from + 1;
      const isPmtMonth = (ppy === 12) || ((monthsInLayer - 1) % 3 === 0);
      const lastMonthInLayer = (m === to);
      const isLastInstallmentOfLoan = isLastLayer && lastMonthInLayer;

      if (L.paymentType === 'Customized Principal') {
        interest = urpa * monthlyRate;
        principal = Math.min(L.customPrincipal || 0, urpa);
        installment = principal + interest;
      } else if (L.paymentType === 'EMI') {
        interest = urpa * monthlyRate;
        principal = pmt - interest;
        installment = pmt;
      } else if (L.paymentType === 'EQI') {
        // EQI quarterly payments fall on absolute month divisible by 3 (Mar, Jun, Sep, Dec)
        if (m % 3 === 0) {
          interest = urpa * (ratePerYear / 4);
          principal = pmt - interest;
          installment = pmt;
        }
      } else if (L.paymentType === 'Equal Principal + Interest (Monthly)') {
        interest = urpa * monthlyRate;
        principal = Math.min(epiConstPrincipal, urpa);
        installment = principal + interest;
      } else if (L.paymentType === 'Equal Principal + Interest (Quarterly)') {
        if (m % 3 === 0) {
          interest = urpa * (ratePerYear / 4);
          principal = Math.min(epiConstPrincipal, urpa);
          installment = principal + interest;
        }
      }

      // Final month of LOAN: settle remaining principal + accrued interest
      if (isLastInstallmentOfLoan) {
        const periodRate = (ppy === 4 ? ratePerYear / 4 : monthlyRate);
        interest = urpa * periodRate;
        principal = urpa;
        installment = principal + interest + accruedReceivable;
        accruedReceivable = 0;
      } else if (li === 0 && paymentCounter === 0 && accruedReceivable > 0 && installment > 0) {
        // Apply accrued moratorium interest to first installment of first layer
        installment += accruedReceivable;
        accruedReceivable = 0;
      }
      urpa = Math.max(0, urpa - principal);

      rows.push({
        sl: m, installment, interest, principal, urpa,
        interestExpense: urpa * monthlyCof, idpReceivable: accruedReceivable,
        paymentType: L.paymentType,
      });

      if (installment > 0) paymentCounter++;
    }
  }
  applyIntExpenseAccrual(rows, monthlyCof);
  return { rows, accruedReceivable, layerInstallments };
}

// ============================================================
// Metrics — keep only what's needed; surface ERR primarily, plus NIM% and NII$
// ============================================================
export function computeMetrics(schedule, params) {
  const { loanAmount, ratePerYear, cofRate = 0, paymentMode, securityKind, numInst = 0,
          tenorMonths = 0, moratoriumMonths = 0 } = params;
  let securityAmount = params.securityAmount || 0;
  let securityRate = params.securityRate || 0;
  const ppy = periodsPerYear(paymentMode || 'EMI');
  const rows = schedule.rows;

  // Installment-equivalent funded security.
  //   "EMI after Moratorium" = PMT(rate/12, regularMonths, -loan)   (B14 in the sample)
  //   "EQI after Moratorium" = PMT(rate/4,  regularQuarters, -loan)
  //   "Installment"          = the post-moratorium Equal-Principal installment (baseInstallment)
  // These are the hypothetical EMI/EQI size after moratorium, independent of the actual layers.
  const kind = String(securityKind || '');
  const regularMonths = Math.max(0, tenorMonths - moratoriumMonths);
  if (kind.startsWith('EMI') || kind.startsWith('EQI') || kind === 'Installment') {
    let unitInstallment = 0;
    if (kind.startsWith('EMI')) {
      unitInstallment = regularMonths > 0 ? PMT(ratePerYear / 12, regularMonths, -loanAmount) : 0;
    } else if (kind.startsWith('EQI')) {
      const q = Math.max(1, Math.round(regularMonths / 3));
      unitInstallment = PMT(ratePerYear / 4, q, -loanAmount);
    } else if (schedule.baseInstallment) {
      // Structured Equal-Principal: first installment size
      unitInstallment = schedule.baseInstallment;
    }
    securityAmount = unitInstallment * (numInst || 1);
    securityRate = 0;
  }

  // Avg portfolio = avg of URPA at start of each month (rows 0..N-1)
  const urpaSeries = rows.slice(0, -1).map(r => r.urpa);
  const avgPortfolio = urpaSeries.length ? urpaSeries.reduce((s, v) => s + v, 0) / urpaSeries.length : 0;

  const totalInterest = rows.reduce((s, r) => s + r.interest, 0);
  // Int expense accrues at end of every month from sl 1 through sl N (last month).
  const totalInterestExpense = rows.slice(1).reduce((s, r) => s + (r.interestExpense || 0), 0);
  const totalMonths = rows.length - 1;
  const tenorYears = totalMonths / 12;

  const csBenefit = (cofRate - securityRate) * (securityAmount || 0) * tenorYears;
  const netInterestExpense = totalInterestExpense - csBenefit;

  // NII ($) per user: Interest received + CS savings - Interest paid
  const nii = totalInterest + csBenefit - totalInterestExpense;
  const nim = avgPortfolio > 0 && tenorYears > 0 ? (nii / avgPortfolio) / tenorYears : 0;

  // ERR — match Excel I4 = E6 + I3 = COF + NIM. Internally computed.
  const effectiveRate = cofRate + nim;

  return {
    effectiveRate,
    nim,
    nii,
    // Auxiliary (not displayed, used for context)
    avgPortfolio,
    totalInterest,
    totalInterestExpense,
    csBenefit,
    netInterestExpense,
    tenorYears,
    derivedSecurityAmount: securityAmount,
    derivedSecurityRate: securityRate,
  };
}

// ============================================================
// Rate Revision — Structured
// ============================================================
export function buildRateRevisionStructured(p) {
  const {
    initialLoanAmount,
    disbursementDate,
    moratoriumMonths = 0,
    idpFlags = [],
    paymentModality,
    tenorMonths,
    rateLayers,
    securityLayers,
    cofLayers = null,
  } = p;

  const ppy = periodsPerYear(paymentModality);
  const start = new Date(disbursementDate);

  function addMonths(d, m) { const x = new Date(d); x.setMonth(x.getMonth() + m); return x; }
  function fmt(d) { return d.toISOString().slice(0, 10); }
  function getRateOn(dateStr) {
    if (!rateLayers || !rateLayers.length) return 0;
    for (const r of rateLayers) {
      if (r.fromDate && r.toDate && dateStr >= r.fromDate && dateStr <= r.toDate) return r.activeRate || 0;
    }
    return rateLayers[0].activeRate || 0;
  }
  function getCofOn(dateStr) {
    if (!cofLayers || !cofLayers.length) return 0;
    for (const r of cofLayers) {
      if (r.fromDate && r.toDate && dateStr >= r.fromDate && dateStr <= r.toDate) return r.cofRate || 0;
    }
    return 0;
  }
  function getSecurityOn(dateStr) {
    if (!securityLayers || !securityLayers.length) return { amount: 0, rate: 0 };
    for (const r of securityLayers) {
      if (r.fromDate && r.toDate && dateStr >= r.fromDate && dateStr <= r.toDate) {
        return { amount: r.amount || 0, rate: r.activeRate || 0 };
      }
    }
    return { amount: 0, rate: 0 };
  }

  const rows = [];
  let urpa = initialLoanAmount;
  const date0 = fmt(start);
  rows.push({
    sl: 0, date: date0, installment: 0, interest: 0, principal: 0, urpa,
    interestExpense: 0,
    rate: getRateOn(date0), cof: getCofOn(date0),
    securityAmount: getSecurityOn(date0).amount, securityRate: getSecurityOn(date0).rate,
  });

  let accruedReceivable = 0;

  for (let m = 1; m <= tenorMonths; m++) {
    const d = fmt(addMonths(start, m));
    const ratePm = getRateOn(d);
    const cofPm = getCofOn(d);
    const sec = getSecurityOn(d);
    const monthlyRate = ratePm / 12;
    const monthlyCof = cofPm / 12;

    let installment = 0, interest = 0, principal = 0;
    if (m <= moratoriumMonths) {
      interest = urpa * monthlyRate;
      const paid = !!idpFlags[m - 1];
      if (paid) {
        installment = interest + accruedReceivable;
        accruedReceivable = 0;
      } else {
        accruedReceivable += interest;
      }
    } else {
      const monthsSinceMora = m - moratoriumMonths;
      const isPmtMonth = (ppy === 12) || (monthsSinceMora % 3 === 0);
      if (isPmtMonth) {
        const remainingMonths = tenorMonths - m + 1;
        const remainingPayments = (ppy === 12) ? remainingMonths : Math.ceil(remainingMonths / 3);
        const periodRate = ratePm / ppy;
        if (paymentModality === 'EMI' || paymentModality === 'EQI') {
          installment = remainingPayments > 0 ? PMT(periodRate, remainingPayments, -urpa) : urpa * (1 + periodRate);
          interest = urpa * periodRate;
          principal = installment - interest;
        } else {
          principal = urpa / remainingPayments;
          interest = urpa * periodRate;
          installment = principal + interest;
        }
        if (m === tenorMonths) {
          principal = urpa;
          installment = principal + interest;
        }
        // first regular installment carries accrued
        if (monthsSinceMora === 1 && accruedReceivable > 0) {
          installment += accruedReceivable;
          accruedReceivable = 0;
        } else if (ppy === 4 && monthsSinceMora === 3 && accruedReceivable > 0) {
          installment += accruedReceivable;
          accruedReceivable = 0;
        }
        urpa = Math.max(0, urpa - principal);
      }
    }

    rows.push({
      sl: m, date: d, installment, interest, principal, urpa,
      interestExpense: 0, // populated below per applyIntExpenseAccrual
      rate: ratePm, cof: cofPm,
      securityAmount: sec.amount, securityRate: sec.rate,
      idpReceivable: accruedReceivable,
    });
  }

  applyIntExpenseAccrual(rows, 0, (i) => (rows[i].cof || 0) / 12);
  return { rows };
}

export function computeRevisionMetrics(schedule, { cofLayers, securityLayers, hasNimComparison }) {
  const rows = schedule.rows;
  const urpaSeries = rows.slice(0, -1).map(r => r.urpa);
  const avgPortfolio = urpaSeries.length ? urpaSeries.reduce((s, v) => s + v, 0) / urpaSeries.length : 0;
  const totalInterest = rows.reduce((s, r) => s + r.interest, 0);
  const totalInterestExpense = rows.slice(1).reduce((s, r) => s + (r.interestExpense || 0), 0);
  const totalMonths = rows.length - 1;
  const tenorYears = totalMonths / 12;

  let csBenefit = 0;
  if (hasNimComparison) {
    csBenefit = rows.reduce((s, r) => s + ((r.cof || 0) - (r.securityRate || 0)) * (r.securityAmount || 0) / 12, 0);
  }
  const nii = totalInterest + csBenefit - totalInterestExpense;
  const nim = avgPortfolio > 0 && tenorYears > 0 ? (nii / avgPortfolio) / tenorYears : 0;

  let effectiveRate = 0;
  if (hasNimComparison) {
    const avgCof = rows.reduce((s, r) => s + (r.cof || 0), 0) / rows.length;
    effectiveRate = avgCof + nim;
  } else {
    effectiveRate = avgPortfolio > 0 && tenorYears > 0 ? (totalInterest / avgPortfolio) / tenorYears : 0;
  }

  return {
    effectiveRate, nim, nii,
    avgPortfolio, totalInterest, totalInterestExpense, csBenefit,
    tenorYears,
  };
}

export function computeRevisionCustomizedMetrics(uploadedRows, { securityLayers = [], cofLayers = null, hasNimComparison = false } = {}) {
  const rows = uploadedRows;
  if (!rows || rows.length < 2) return null;

  const dates = rows.map(r => new Date(r.date));
  const periodicRates = [];
  const exposure = [];
  for (let i = 1; i < rows.length; i++) {
    const prev = rows[i - 1];
    const cur = rows[i];
    const days = Math.max(1, (dates[i] - dates[i - 1]) / 86400000);
    const r = prev.urpa > 0 ? (cur.interestAmount / prev.urpa) : 0;
    const annualR = r * (365 / days);
    periodicRates.push(annualR);
    exposure.push(prev.urpa * days);
  }
  const totalExposure = exposure.reduce((s, v) => s + v, 0);
  const yolWeighted = totalExposure > 0
    ? periodicRates.reduce((s, r, i) => s + r * exposure[i], 0) / totalExposure
    : 0;

  const totalInterest = rows.reduce((s, r) => s + (r.interestAmount || 0), 0);
  const totalDays = (dates[dates.length - 1] - dates[0]) / 86400000;
  const avgPortfolio = totalDays > 0 ? totalExposure / totalDays : 0;
  const tenorYears = totalDays / 365;

  let csBenefit = 0, totalInterestExpense = 0, effectiveRate = yolWeighted;
  if (hasNimComparison && cofLayers && cofLayers.length) {
    const getCofOn = (dateStr) => {
      for (const r of cofLayers) if (r.fromDate && r.toDate && dateStr >= r.fromDate && dateStr <= r.toDate) return r.cofRate || 0;
      return 0;
    };
    const getSecOn = (dateStr) => {
      for (const r of (securityLayers || []))
        if (r.fromDate && r.toDate && dateStr >= r.fromDate && dateStr <= r.toDate)
          return { amount: r.amount || 0, rate: r.activeRate || 0 };
      return { amount: 0, rate: 0 };
    };
    for (let i = 1; i < rows.length; i++) {
      const prev = rows[i - 1];
      const days = Math.max(1, (dates[i] - dates[i - 1]) / 86400000);
      const dateStr = rows[i - 1].date;
      const cof = getCofOn(dateStr);
      const sec = getSecOn(dateStr);
      totalInterestExpense += prev.urpa * cof * (days / 365);
      csBenefit += (cof - sec.rate) * sec.amount * (days / 365);
    }
    const nii = totalInterest + csBenefit - totalInterestExpense;
    const nim = avgPortfolio > 0 && tenorYears > 0 ? (nii / avgPortfolio) / tenorYears : 0;
    const avgCof = cofLayers.reduce((s, r) => s + (r.cofRate || 0), 0) / cofLayers.length;
    effectiveRate = avgCof + nim;
    return { effectiveRate, nim, nii, avgPortfolio, totalInterest, totalInterestExpense, csBenefit, tenorYears };
  }
  const nii = totalInterest - totalInterestExpense + csBenefit;
  const nim = avgPortfolio > 0 && tenorYears > 0 ? (nii / avgPortfolio) / tenorYears : 0;
  return { effectiveRate, nim, nii, avgPortfolio, totalInterest, totalInterestExpense, csBenefit, tenorYears };
}
