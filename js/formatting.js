// Number/percentage formatting + live input helpers

export function formatNumber(value, { decimals = 2 } = {}) {
  if (value === '' || value === null || value === undefined || isNaN(value)) return '';
  const num = Number(value);
  return num.toLocaleString('en-US', {
    minimumFractionDigits: 0,
    maximumFractionDigits: decimals,
  });
}

export function formatMoney(value, { decimals = 2 } = {}) {
  if (value === '' || value === null || value === undefined || isNaN(value)) return '-';
  return Number(value).toLocaleString('en-US', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

export function formatPercent(value, { decimals = 4 } = {}) {
  if (value === '' || value === null || value === undefined || isNaN(value)) return '-';
  return (Number(value) * 100).toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: decimals,
  }) + '%';
}

// strip commas, return raw number string
export function unformat(str) {
  if (str === null || str === undefined) return '';
  return String(str).replace(/,/g, '').trim();
}

export function parseNumber(str) {
  const raw = unformat(str);
  if (raw === '' || raw === '-' || raw === '.') return NaN;
  return Number(raw);
}

// Apply live comma formatting while preserving cursor position
export function attachCommaFormatter(input, { allowDecimal = true, integerOnly = false } = {}) {
  function format(raw) {
    if (!raw) return '';
    // keep only digits and one dot (if decimals allowed)
    let cleaned = raw.replace(/,/g, '');
    if (integerOnly) cleaned = cleaned.replace(/[^0-9]/g, '');
    else cleaned = cleaned.replace(/[^0-9.]/g, '');
    // collapse multiple dots
    if (allowDecimal && !integerOnly) {
      const firstDot = cleaned.indexOf('.');
      if (firstDot !== -1) {
        cleaned = cleaned.slice(0, firstDot + 1) +
                  cleaned.slice(firstDot + 1).replace(/\./g, '');
      }
    }
    if (!cleaned) return '';
    const [intPart, decPart] = cleaned.split('.');
    const intFormatted = intPart === '' ? '' : Number(intPart).toLocaleString('en-US');
    if (decPart === undefined) return intFormatted;
    return intFormatted + '.' + decPart.slice(0, 2);
  }

  input.addEventListener('input', (e) => {
    const target = e.target;
    const oldVal = target.value;
    const cursor = target.selectionStart;
    const digitsBeforeCursor = (oldVal.slice(0, cursor).match(/[0-9.]/g) || []).length;

    const newVal = format(oldVal);
    target.value = newVal;

    // restore cursor — find position with same count of digits/dots
    let count = 0;
    let newCursor = newVal.length;
    for (let i = 0; i < newVal.length; i++) {
      if (/[0-9.]/.test(newVal[i])) count++;
      if (count >= digitsBeforeCursor) { newCursor = i + 1; break; }
    }
    target.setSelectionRange(newCursor, newCursor);

    target.classList.toggle('empty', newVal === '');
    target.dispatchEvent(new CustomEvent('formatted', { detail: { rawValue: parseNumber(newVal) } }));
  });

  input.addEventListener('blur', () => {
    input.classList.toggle('empty', input.value === '');
  });
}

// Attach a "%" suffix that appears only when input has value
export function attachPercentSuffix(wrapper) {
  const input = wrapper.querySelector('input');
  function refresh() {
    input.classList.toggle('empty', input.value.trim() === '');
  }
  input.addEventListener('input', refresh);
  input.addEventListener('blur', refresh);
  refresh();
}
