// Reusable UI component builders (returns DOM nodes)
import { attachCommaFormatter, sanitizeDecimalString, formatTwoDecimalsOnBlur } from './formatting.js';

let uid = 0;
const nextId = () => `f${++uid}`;

export function el(tag, attrs = {}, ...children) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') node.className = v;
    else if (k === 'html') node.innerHTML = v;
    else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2), v);
    else if (v === true) node.setAttribute(k, '');
    else if (v !== false && v !== null && v !== undefined) node.setAttribute(k, v);
  }
  for (const child of children.flat()) {
    if (child === null || child === undefined || child === false) continue;
    node.appendChild(typeof child === 'string' ? document.createTextNode(child) : child);
  }
  return node;
}

// Info-icon — click toggles a popover; closes on outside click or Escape
export function infoIcon(text) {
  const ic = el('span', { class: 'info-icon', tabindex: '0' }, 'i');
  ic.addEventListener('click', (e) => {
    e.stopPropagation();
    closeAllInfoPopovers();
    const pop = el('div', { class: 'info-popover' }, text);
    document.body.appendChild(pop);
    const r = ic.getBoundingClientRect();
    pop.style.left = (window.scrollX + r.right + 6) + 'px';
    pop.style.top = (window.scrollY + r.top - 4) + 'px';
    // If overflow right, flip to left
    const pr = pop.getBoundingClientRect();
    if (pr.right > window.innerWidth - 8) {
      pop.style.left = (window.scrollX + r.left - pr.width - 6) + 'px';
    }
    setTimeout(() => {
      const onDoc = (ev) => {
        if (!pop.contains(ev.target) && ev.target !== ic) closePopover();
      };
      const onKey = (ev) => { if (ev.key === 'Escape') closePopover(); };
      function closePopover() {
        pop.remove();
        document.removeEventListener('mousedown', onDoc);
        document.removeEventListener('keydown', onKey);
      }
      pop._close = closePopover;
      document.addEventListener('mousedown', onDoc);
      document.addEventListener('keydown', onKey);
    }, 0);
  });
  return ic;
}
function closeAllInfoPopovers() {
  document.querySelectorAll('.info-popover').forEach(p => p._close ? p._close() : p.remove());
}

// Number field
export function numberField({ label, name, placeholder = '', help = '', integerOnly = false, min = null, tooltip = '' }) {
  const id = nextId();
  const input = el('input', {
    id, type: 'text', inputmode: integerOnly ? 'numeric' : 'decimal',
    placeholder, autocomplete: 'off', 'data-name': name, class: 'numeric-input',
  });
  attachCommaFormatter(input, { integerOnly });
  if (!integerOnly) input.addEventListener('blur', () => formatTwoDecimalsOnBlur(input));
  const field = el('div', { class: 'field' });
  const lbl = labelWithTooltip(id, label, tooltip);
  field.appendChild(lbl);
  field.appendChild(input);
  if (help) field.appendChild(el('span', { class: 'help' }, help));

  field.getValue = () => {
    const v = input.value.replace(/,/g, '');
    if (v === '') return null;
    const n = Number(v);
    if (isNaN(n)) return null;
    if (min !== null && n < min) return null;
    return n;
  };
  field.setValue = (v) => {
    if (v === null || v === undefined || v === '') input.value = '';
    else if (integerOnly) input.value = Number(v).toLocaleString('en-US', { maximumFractionDigits: 0 });
    else input.value = Number(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  };
  field.input = input;
  field.setLabel = (text) => updateLabel(field, text);
  return field;
}

function labelWithTooltip(id, text, tooltip) {
  const lbl = el('label', { for: id }, text);
  if (tooltip) lbl.appendChild(infoIcon(tooltip));
  return lbl;
}
function updateLabel(field, text) {
  const l = field.querySelector('label');
  if (!l) return;
  l.firstChild.nodeValue = text;
  const ic = l.querySelector('.info-icon');
  if (ic) l.appendChild(ic);
}

// Percent field
export function percentField({ label, name, placeholder = '', help = '', tooltip = '' }) {
  const id = nextId();
  const wrapper = el('div', { class: 'field percent-field' });
  wrapper.appendChild(labelWithTooltip(id, label, tooltip));

  const inputWrap = el('div', { class: 'percent-input-fixed' });
  const input = el('input', {
    id, type: 'text', inputmode: 'decimal',
    placeholder, autocomplete: 'off', 'data-name': name, class: 'numeric-input',
  });
  const suffix = el('span', { class: 'percent-suffix-fixed' }, '%');
  inputWrap.appendChild(input);
  inputWrap.appendChild(suffix);
  wrapper.appendChild(inputWrap);
  if (help) wrapper.appendChild(el('span', { class: 'help' }, help));

  function refreshEmpty() { inputWrap.classList.toggle('empty', input.value === ''); }
  input.addEventListener('input', (e) => { e.target.value = sanitizeDecimalString(e.target.value); refreshEmpty(); });
  input.addEventListener('blur', () => {
    if (input.value === '' || input.value === '.') input.value = '';
    else { const n = Number(input.value); if (!isNaN(n)) input.value = n.toFixed(2); }
    refreshEmpty();
  });
  refreshEmpty();

  wrapper.getValue = () => {
    const v = input.value.trim();
    if (v === '' || v === '.') return null;
    const n = Number(v);
    return isNaN(n) ? null : n / 100;
  };
  wrapper.setValue = (v) => {
    if (v === null || v === undefined || v === '') input.value = '';
    else input.value = (Number(v) * 100).toFixed(2);
    refreshEmpty();
  };
  wrapper.input = input;
  wrapper.setLabel = (text) => updateLabel(wrapper, text);
  return wrapper;
}

// Option field (native select, centered)
export function optionField({ label, name, options, value = null, onChange = null, help = '', tooltip = '' }) {
  const id = nextId();
  const wrapper = el('div', { class: 'field' });
  wrapper.appendChild(labelWithTooltip(id, label, tooltip));
  const select = el('select', { id, 'data-name': name, class: 'centered-input' });
  function fillOptions(opts) {
    select.innerHTML = '';
    opts.forEach((opt) => {
      const text = typeof opt === 'string' ? opt : opt.label;
      const val = typeof opt === 'string' ? opt : opt.value;
      select.appendChild(el('option', { value: val }, text));
    });
  }
  fillOptions(options);
  if (value !== null) select.value = value;
  if (onChange) select.addEventListener('change', () => onChange(select.value));
  wrapper.appendChild(select);
  if (help) wrapper.appendChild(el('span', { class: 'help' }, help));

  wrapper.getValue = () => select.value;
  wrapper.setValue = (v) => { select.value = v; };
  wrapper.setOptions = (opts) => {
    const prev = select.value;
    fillOptions(opts);
    if ([...select.options].some(o => o.value === prev)) select.value = prev;
    else if (select.options.length) select.value = select.options[0].value;
  };
  wrapper.setLabel = (text) => updateLabel(wrapper, text);
  wrapper.select = select;
  return wrapper;
}

// Date field — flatpickr, DD-Mmm-YYYY, Year dropdown
export function dateField({ label, name, placeholder = 'dd-Mmm-yyyy', tooltip = '', onChange = null, disableFn = null }) {
  const id = nextId();
  const input = el('input', { id, type: 'text', 'data-name': name, placeholder, autocomplete: 'off', class: 'centered-input date-input' });
  const field = el('div', { class: 'field' });
  field.appendChild(labelWithTooltip(id, label, tooltip));
  field.appendChild(input);

  let fp;
  requestAnimationFrame(() => {
    if (typeof flatpickr === 'undefined') return;
    fp = flatpickr(input, {
      dateFormat: 'd-M-Y',
      allowInput: true,
      disable: disableFn ? [disableFn] : [],
      onChange: () => { if (onChange) onChange(input.value); },
      onReady: applyYearDropdown,
      onMonthChange: applyYearDropdown,
      onYearChange: applyYearDropdown,
    });
    field._fp = fp;
  });

  field.getValue = () => {
    if (!input.value) return null;
    const d = parseDDMMMYYYY(input.value);
    return d ? d.toISOString().slice(0, 10) : null;
  };
  field.setValue = (v) => {
    if (!v) { input.value = ''; if (fp) fp.clear(); return; }
    if (fp) fp.setDate(v, true);
    else input.value = formatDDMMMYYYY(new Date(v));
  };
  field.input = input;
  return field;
}

// Pre-built year <option> HTML — built once, cloned per flatpickr instance.
const YEAR_OPTIONS_HTML = (() => {
  const now = new Date().getFullYear();
  let html = '';
  for (let y = now - 100; y <= now + 100; y++) html += `<option value="${y}">${y}</option>`;
  return html;
})();

// Convert flatpickr's year input into a styled select dropdown.
function applyYearDropdown(selectedDates, dateStr, instance) {
  const yearWrap = instance.calendarContainer.querySelector('.numInputWrapper');
  if (!yearWrap || yearWrap.dataset.dropdownized) return;
  yearWrap.dataset.dropdownized = '1';
  const input = yearWrap.querySelector('.cur-year');
  const sel = document.createElement('select');
  sel.className = 'fp-year-select';
  sel.innerHTML = YEAR_OPTIONS_HTML;
  sel.value = input.value;
  sel.addEventListener('change', () => instance.changeYear(Number(sel.value)));
  yearWrap.insertBefore(sel, input);
  input.style.display = 'none';
  yearWrap.querySelectorAll('.arrowUp, .arrowDown').forEach(a => a.style.display = 'none');
}

export function parseDDMMMYYYY(s) {
  if (!s) return null;
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) { const d = new Date(s); return isNaN(d) ? null : d; }
  const m = s.match(/^(\d{1,2})[-\s\/](\w{3})[-\s\/](\d{4})$/);
  if (!m) { const d = new Date(s); return isNaN(d) ? null : d; }
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const mi = months.findIndex(x => x.toLowerCase() === m[2].toLowerCase());
  if (mi < 0) return null;
  return new Date(Date.UTC(Number(m[3]), mi, Number(m[1])));
}
export function formatDDMMMYYYY(d) {
  if (!d || isNaN(d)) return '';
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const day = String(d.getDate()).padStart(2, '0');
  return `${day}-${months[d.getMonth()]}-${d.getFullYear()}`;
}

// Month boxes — selectable equal-width boxes (max 6/row). Optional "Select all" checkbox.
export function monthBoxesField({ name, getCount, tooltip = '', label = '', selectAll = true }) {
  const wrapper = el('div', { class: 'field' });
  if (label) {
    const head = el('div', { class: 'label-row' });
    const lbl = el('label', {}, label);
    if (tooltip) lbl.appendChild(infoIcon(tooltip));
    head.appendChild(lbl);
    if (selectAll) {
      const allLabel = el('label', { class: 'select-all-toggle' });
      const allCb = el('input', { type: 'checkbox' });
      allLabel.appendChild(allCb);
      allLabel.appendChild(document.createTextNode(' Select all'));
      head.appendChild(allLabel);
      allCb.addEventListener('change', () => {
        states = states.map(() => allCb.checked);
        render();
      });
      wrapper._allCb = allCb;
    }
    wrapper.appendChild(head);
  }
  const grid = el('div', { class: 'month-boxes', 'data-name': name });
  wrapper.appendChild(grid);

  let states = [];
  function syncAllCb() {
    if (!wrapper._allCb) return;
    const n = states.length;
    wrapper._allCb.checked = n > 0 && states.every(Boolean);
    wrapper._allCb.indeterminate = !wrapper._allCb.checked && states.some(Boolean);
  }
  function render() {
    const n = Math.max(0, getCount() || 0);
    states = Array.from({ length: n }, (_, i) => states[i] ?? false);
    grid.innerHTML = '';
    for (let i = 0; i < n; i++) {
      const box = el('div', { class: 'month-box' + (states[i] ? ' selected' : ''), 'data-month': i + 1 },
        el('div', { class: 'mb-num' }, String(i + 1).padStart(2, '0')),
        el('div', { class: 'mb-lbl' }, 'Month'),
      );
      box.addEventListener('click', () => {
        states[i] = !states[i];
        box.classList.toggle('selected', states[i]);
        syncAllCb();
      });
      grid.appendChild(box);
    }
    syncAllCb();
  }
  wrapper.refresh = render;
  wrapper.getValue = () => states.slice();
  wrapper.setValue = (arr) => { states = (arr || []).slice(); render(); };
  render();
  return wrapper;
}

// ============================================================
// Layered field with universal logic
// Options:
//   schema, addLabel, help, onChange
//   minRows: number of initial undeletable rows (default 0)
//   maxRows: hard cap (default Infinity)
//   getMaturity(): returns {value: <max from-key>, kind: 'date'|'month'} — used for last-row To cascading
//   cascadingFromKey, cascadingToKey: keys used for cascading (e.g. 'fromDate'/'toDate' or 'fromInstallment'/'toInstallment')
//   advanceUnit: 'day' | 'month' (for cascading +1)
//   allowFromEqualTo: bool (default true)
//   validate(rows, maturity): return null or error message
// ============================================================
export function layeredField(opts) {
  const {
    label, name, schema, addLabel = '+ Add layer', help = '', initialRows = 1, onChange = null,
    minRows = 0, maxRows = Infinity,
    cascadingFromKey, cascadingToKey,
    getMaturity = null,
    getAnchor = null, // first-row cascadingFromKey source of truth (e.g. disbursement date)
    allowFromEqualTo = true,
  } = opts;
  // Hoisted once — schema is closed-over and never changes
  const fromKind = cascadingFromKey ? (schema.find(s => s.key === cascadingFromKey)?.type) : null;
  const toKind = cascadingToKey ? (schema.find(s => s.key === cascadingToKey)?.type) : null;

  const wrapper = el('div', { class: 'field' });
  wrapper.appendChild(el('label', {}, label));

  const layers = el('div', { class: 'layers', 'data-name': name });
  wrapper.appendChild(layers);
  if (help) wrapper.appendChild(el('span', { class: 'help' }, help));

  const cols = schema.map(s => s.width || '1fr').join(' ') + ' 36px';

  const header = el('div', { class: 'layer-header', style: `grid-template-columns: ${cols}` });
  schema.forEach(s => header.appendChild(el('div', { class: 'layer-th' }, s.label)));
  header.appendChild(el('div', {}, ''));
  layers.appendChild(header);

  const rows = [];
  function fireChange() {
    if (onChange) onChange(wrapper.getValue(), rows);
    applyLayerRules();
  }
  function syncHeaderVisibility() {
    header.classList.toggle('hidden', rows.length === 0);
  }

  function addRow(values = {}, opts = {}) {
    const row = el('div', { class: 'layer-row', style: `grid-template-columns: ${cols}` });
    const inputs = {};
    schema.forEach((s) => buildCell(s, row, inputs, values));
    const del = el('button', { type: 'button', class: 'row-del', title: 'Remove' }, '×');
    del.addEventListener('click', () => removeRow(rowApi));
    row.appendChild(del);
    layers.appendChild(row);
    const rowApi = { row, inputs, deletable: !opts.undeletable, errors: new Set() };
    if (!rowApi.deletable) del.style.visibility = 'hidden';
    rows.push(rowApi);
    syncHeaderVisibility();
    return rowApi;
  }

  function buildCell(s, row, inputs, values) {
    let inp;
    if (s.type === 'option') {
      inp = el('select', { class: 'centered-input' });
      const opts = typeof s.options === 'function' ? s.options() : s.options;
      if (s.allowEmpty) inp.appendChild(el('option', { value: '' }, s.placeholder || '— select —'));
      opts.forEach((o) => {
        const val = typeof o === 'string' ? o : o.value;
        const txt = typeof o === 'string' ? o : o.label;
        inp.appendChild(el('option', { value: val }, txt));
      });
      if (values[s.key] !== undefined && values[s.key] !== null) inp.value = values[s.key];
      else if (s.allowEmpty) inp.value = '';
      inp.addEventListener('change', fireChange);
    } else if (s.type === 'date') {
      inp = el('input', { type: 'text', placeholder: 'dd-Mmm-yyyy', class: 'centered-input date-input' });
      if (values[s.key]) inp.value = values[s.key];
      requestAnimationFrame(() => {
        if (typeof flatpickr === 'undefined') return;
        flatpickr(inp, {
          dateFormat: 'd-M-Y',
          allowInput: true,
          onChange: () => { inp.dataset.userSet = '1'; fireChange(); },
          onReady: applyYearDropdown,
          onMonthChange: applyYearDropdown,
          onYearChange: applyYearDropdown,
        });
      });
      inp.addEventListener('change', () => { inp.dataset.userSet = '1'; fireChange(); });
    } else if (s.type === 'percent') {
      const pwrap = el('div', { class: 'percent-input-fixed inline' });
      inp = el('input', { type: 'text', inputmode: 'decimal', class: 'numeric-input' });
      const sfx = el('span', { class: 'percent-suffix-fixed' }, '%');
      pwrap.appendChild(inp);
      pwrap.appendChild(sfx);
      if (values[s.key] !== undefined && values[s.key] !== null) inp.value = (values[s.key] * 100).toFixed(2);
      const refreshEmpty = () => pwrap.classList.toggle('empty', inp.value === '');
      inp.addEventListener('input', () => {
        inp.value = sanitizeDecimalString(inp.value);
        refreshEmpty();
        fireChange();
      });
      inp.addEventListener('blur', () => {
        if (inp.value === '' || inp.value === '.') inp.value = '';
        else inp.value = Number(inp.value).toFixed(2);
        refreshEmpty();
        fireChange();
      });
      refreshEmpty();
      row.appendChild(pwrap);
      inputs[s.key] = inp;
      return;
    } else {
      inp = el('input', { type: 'text', inputmode: 'decimal', placeholder: '0', class: 'numeric-input' });
      if (values[s.key] !== undefined && values[s.key] !== null) {
        inp.value = Number(values[s.key]).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      }
      attachCommaFormatter(inp, { integerOnly: !!s.integerOnly });
      if (!s.integerOnly) inp.addEventListener('blur', () => formatTwoDecimalsOnBlur(inp));
      inp.addEventListener('input', fireChange);
    }
    inputs[s.key] = inp;
    row.appendChild(inp);
  }

  function removeRow(rowApi) {
    if (!rowApi.deletable) return;
    const idx = rows.indexOf(rowApi);
    if (idx < 0) return;
    rows.splice(idx, 1);
    rowApi.row.remove();
    syncHeaderVisibility();
    fireChange();
  }

  const addBtn = el('button', { type: 'button', class: 'add-layer' }, addLabel);
  addBtn.addEventListener('click', () => {
    if (rows.length >= maxRows) {
      if (wrapper.onCannotAdd) wrapper.onCannotAdd('Maximum layers reached.');
      return;
    }
    if (!canAddNew()) {
      if (wrapper.onCannotAdd) wrapper.onCannotAdd('Complete the existing layer(s) and ensure the last "To" is before maturity before adding a new layer.');
      return;
    }
    addRow();
    fireChange();
  });
  wrapper.appendChild(addBtn);

  // Seed initial rows + minRows (undeletable)
  const totalInitial = Math.max(initialRows, minRows);
  for (let i = 0; i < totalInitial; i++) addRow({}, { undeletable: i < minRows });
  // Apply rules once initial rows have mounted (after rAF so flatpickr is attached)
  requestAnimationFrame(() => applyLayerRules());

  // ============ Universal layer rules ============
  function getMaturityValue() {
    if (!getMaturity) return null;
    const m = getMaturity();
    return m && m.value ? m.value : null;
  }
  function readVal(inp, kind) {
    if (!inp) return null;
    const v = inp.value;
    if (!v) return null;
    if (kind === 'date') {
      const d = parseDDMMMYYYY(v);
      return d ? d.toISOString().slice(0, 10) : null;
    }
    return Number(v);
  }
  function setVal(inp, v, kind) {
    if (!inp) return;
    if (kind === 'date') {
      if (inp._flatpickr) inp._flatpickr.setDate(v, false);
      else inp.value = v ? formatDDMMMYYYY(new Date(v)) : '';
    } else {
      inp.value = (v === null || v === undefined) ? '' : String(v);
    }
  }
  function advanceOne(v, kind) {
    if (kind === 'date') {
      const d = new Date(v);
      d.setDate(d.getDate() + 1);
      return d.toISOString().slice(0, 10);
    }
    return Number(v) + 1;
  }

  function rowsComplete() {
    if (!cascadingFromKey || !cascadingToKey) return true;
    return rows.every((r) =>
      readVal(r.inputs[cascadingFromKey], fromKind) != null &&
      readVal(r.inputs[cascadingToKey], toKind) != null
    );
  }
  function canAddNew() {
    if (rows.length === 0) return true;
    if (!cascadingFromKey || !cascadingToKey) return true;
    if (!rowsComplete()) return false;
    const lastTo = readVal(rows[rows.length - 1].inputs[cascadingToKey], toKind);
    const maturity = getMaturityValue();
    if (lastTo && maturity && lastTo >= maturity) return false;
    return true;
  }

  function applyLayerRules() {
    if (!cascadingFromKey || !cascadingToKey) return;
    const maturity = getMaturityValue();
    const anchor = getAnchor ? (getAnchor()?.value || null) : null;

    rows.forEach((r, i) => {
      const fromInp = r.inputs[cascadingFromKey];
      const toInp = r.inputs[cascadingToKey];
      r.errors.clear();
      [fromInp, toInp].forEach(inp => inp?.classList.remove('field-error'));

      if (i === 0) {
        // First row: From = anchor (always, when provided)
        if (anchor && fromInp && readVal(fromInp, fromKind) !== anchor) setVal(fromInp, anchor, fromKind);
        if (fromInp) fromInp.disabled = !!anchor; // user can't edit if anchored
      } else {
        // Subsequent rows: From = prev.To + 1, disabled until prev.To set
        const prevTo = readVal(rows[i - 1].inputs[cascadingToKey], toKind);
        if (prevTo) {
          const desired = advanceOne(prevTo, fromKind);
          if (readVal(fromInp, fromKind) !== desired) setVal(fromInp, desired, fromKind);
        }
        if (fromInp) {
          fromInp.disabled = true; // always cascaded
          if (fromInp._flatpickr) fromInp._flatpickr.set('clickOpens', false);
        }
      }

      // Last row: force To = maturity
      const isLast = i === rows.length - 1;
      if (isLast && maturity) {
        if (readVal(toInp, toKind) !== maturity) setVal(toInp, maturity, toKind);
        if (toInp) toInp.disabled = true;
      } else if (toInp) {
        toInp.disabled = false;
      }
    });

    // Validate logical consistency
    rows.forEach((r, i) => {
      const fromInp = r.inputs[cascadingFromKey];
      const toInp = r.inputs[cascadingToKey];
      const fv = readVal(fromInp, fromKind);
      const tv = readVal(toInp, toKind);
      const isLast = i === rows.length - 1;

      if (fv && tv && (allowFromEqualTo ? tv < fv : tv <= fv)) {
        flagError(toInp, r); flagError(fromInp, r);
      }
      if (tv && maturity && tv > maturity) flagError(toInp, r);
      if (!isLast && tv && maturity && tv === maturity) flagError(toInp, r);
      if (i > 0) {
        const prevTo = readVal(rows[i - 1].inputs[cascadingToKey], toKind);
        if (prevTo && fv && fv <= prevTo) flagError(fromInp, r);
      }
    });

    function flagError(inp, r) {
      if (!inp) return;
      inp.classList.add('field-error');
      r.errors.add(inp);
    }
  }
  function hasErrors() {
    return rows.some(r => r.errors && r.errors.size > 0);
  }

  // Public API
  wrapper.getValue = () => rows.map(({ inputs }) => {
    const row = {};
    schema.forEach((s) => {
      const v = inputs[s.key].value;
      if (s.type === 'percent') row[s.key] = v === '' ? null : Number(v) / 100;
      else if (s.type === 'number') {
        const raw = v.replace(/,/g, '');
        row[s.key] = raw === '' ? null : Number(raw);
      } else if (s.type === 'date') {
        const d = parseDDMMMYYYY(v);
        row[s.key] = d ? d.toISOString().slice(0, 10) : null;
      } else row[s.key] = v || null;
    });
    return row;
  });
  wrapper.setValue = (arr) => {
    rows.length = 0;
    layers.querySelectorAll('.layer-row').forEach(r => r.remove());
    (arr || [{}]).forEach((v, i) => addRow(v, { undeletable: i < minRows }));
    syncHeaderVisibility();
    applyLayerRules();
    fireChange();
  };
  wrapper.addRow = addRow;
  wrapper.rows = rows;
  wrapper.addBtn = addBtn;
  wrapper.refreshOptions = () => {
    rows.forEach(({ inputs }) => {
      schema.forEach((s) => {
        if (s.type === 'option' && typeof s.options === 'function') {
          const inp = inputs[s.key];
          const prev = inp.value;
          inp.innerHTML = '';
          if (s.allowEmpty) inp.appendChild(el('option', { value: '' }, s.placeholder || '— select —'));
          s.options().forEach((o) => {
            const val = typeof o === 'string' ? o : o.value;
            const txt = typeof o === 'string' ? o : o.label;
            inp.appendChild(el('option', { value: val }, txt));
          });
          if ([...inp.options].some(o => o.value === prev)) inp.value = prev;
        }
      });
    });
  };
  wrapper.applyLayerRules = applyLayerRules;
  wrapper.hasErrors = hasErrors;
  wrapper.canAddNew = canAddNew;
  // Allow page code to disable inheritance (e.g., for Customized Payment Layers that have non-cascading payment types)
  return wrapper;
}

// Toast + modal
export function toast(message, type = 'info', duration = 2400) {
  const root = document.getElementById('toast-root');
  const t = el('div', { class: `toast ${type}` }, message);
  root.appendChild(t);
  setTimeout(() => {
    t.style.transition = 'opacity 0.2s ease';
    t.style.opacity = '0';
    setTimeout(() => t.remove(), 220);
  }, duration);
}

export function openModal(node) {
  const root = document.getElementById('modal-root');
  const card = document.getElementById('modal-card');
  card.innerHTML = '';
  card.appendChild(node);
  root.classList.remove('hidden');
  root.querySelector('.modal-backdrop').onclick = closeModal;
}
export function closeModal() {
  document.getElementById('modal-root').classList.add('hidden');
}
