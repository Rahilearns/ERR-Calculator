// Reusable UI component builders (returns DOM nodes)
import { attachCommaFormatter, sanitizeDecimalString, formatTwoDecimalsOnBlur } from './formatting.js?v=20260603zt';

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
    const dateObj = isoToLocalDate(v);
    if (fp) fp.setDate(dateObj, false);
    else input.value = formatDDMMMYYYY(dateObj);
  };
  field.input = input;
  return field;
}

// Parse an ISO 'YYYY-MM-DD' string to a Date at LOCAL midnight (avoids the UTC
// off-by-one that `new Date('2020-01-01')` causes in negative-offset timezones).
export function isoToLocalDate(iso) {
  if (iso instanceof Date) return iso;
  const m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) { const d = new Date(iso); return isNaN(d) ? null : d; }
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
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
export function monthBoxesField({ name, getCount, tooltip = '', label = '', selectAll = true, capitalizable = false }) {
  const wrapper = el('div', { class: 'field' });
  let states = []; // per month: 0 = none (accrue), 1 = paid, 2 = capitalized
  const maxState = capitalizable ? 2 : 1;
  if (label) {
    const head = el('div', { class: 'label-row' });
    const lbl = el('label', {}, label);
    if (tooltip) lbl.appendChild(infoIcon(tooltip));
    head.appendChild(lbl);
    if (selectAll) {
      const allLabel = el('label', { class: 'select-all-toggle' });
      const allCb = el('input', { type: 'checkbox' });
      allLabel.appendChild(allCb);
      allLabel.appendChild(document.createTextNode(' Select all (paid)'));
      head.appendChild(allLabel);
      allCb.addEventListener('change', () => {
        const v = allCb.checked ? 1 : 0;
        states = states.map(() => v);
        render();
      });
      wrapper._allCb = allCb;
    }
    wrapper.appendChild(head);
  }
  const grid = el('div', { class: 'month-boxes', 'data-name': name });
  wrapper.appendChild(grid);

  // Explanatory note below the boxes (only when capitalization is offered).
  if (capitalizable) {
    wrapper.appendChild(el('div', { class: 'mora-note' },
      'Click a month to set how its accrued-but-unpaid interest is treated — ',
      el('b', { class: 'mn-paid' }, 'one click = Paid'),
      ' at that month-end, ',
      el('b', { class: 'mn-cap' }, 'two clicks = Capitalized'),
      ' into principal at that month-end (it then earns interest); a third click clears it. ',
      'Unmarked months keep accruing until the next Paid/Capitalized month, or the first installment after the moratorium.',
    ));
  }

  function syncAllCb() {
    if (!wrapper._allCb) return;
    const n = states.length;
    wrapper._allCb.checked = n > 0 && states.every(s => s > 0);
    wrapper._allCb.indeterminate = !wrapper._allCb.checked && states.some(s => s > 0);
  }
  const CLS = { 1: 'paid', 2: 'capitalized' };
  function render() {
    const n = Math.max(0, getCount() || 0);
    states = Array.from({ length: n }, (_, i) => Math.min(Number(states[i]) || 0, maxState));
    grid.innerHTML = '';
    for (let i = 0; i < n; i++) {
      const cls = CLS[states[i]] || '';
      const box = el('div', {
        class: 'month-box' + (cls ? ' ' + cls : ''),
        'data-month': i + 1,
      },
        el('div', { class: 'mb-num' }, String(i + 1).padStart(2, '0')),
        el('div', { class: 'mb-lbl' }, 'Month'),
      );
      box.addEventListener('click', () => {
        states[i] = (states[i] + 1) % (maxState + 1);
        render();
      });
      grid.appendChild(box);
    }
    syncAllCb();
  }

  wrapper.refresh = render;
  wrapper.getValue = () => states.slice();
  wrapper.setValue = (arr) => {
    states = (arr || []).map(v => v === true ? 1 : v === false ? 0 : Math.min(Number(v) || 0, maxState));
    render();
  };
  wrapper.getPaidFlags = () => states.map(s => s === 1);
  wrapper.getCapFlags = () => states.map(s => s === 2);
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

  // Columns fill the full width (no reserved delete column — the × is overlaid at the
  // row's corner instead, so the boxes stretch edge-to-edge and stay equal width).
  const cols = schema.map(s => s.width || '1fr').join(' ');

  const header = el('div', { class: 'layer-header', style: `grid-template-columns: ${cols}` });
  schema.forEach(s => header.appendChild(el('div', { class: 'layer-th' }, s.label)));
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
    let inp, cellNode = null;
    if (s.type === 'option') {
      inp = el('select', { class: 'centered-input' });
      const opts = typeof s.options === 'function' ? s.options() : s.options;
      if (s.allowEmpty) inp.appendChild(el('option', { value: '' }, s.placeholder ?? '— select —'));
      opts.forEach((o) => {
        const val = typeof o === 'string' ? o : o.value;
        const txt = typeof o === 'string' ? o : o.label;
        inp.appendChild(el('option', { value: val }, txt));
      });
      if (values[s.key] !== undefined && values[s.key] !== null) inp.value = values[s.key];
      else if (s.allowEmpty) inp.value = '';
      inp.addEventListener('change', () => { inp.dataset.userSet = '1'; fireChange(); });
      // Auto-derived, user-uneditable field (still set programmatically by applyLayerRules).
      if (s.readOnly) { inp.disabled = true; inp.classList.add('readonly-cell'); }
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
      cellNode = pwrap;
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
    // Wrap each control with its column label so layers can stack on mobile.
    // `.layer-cell { display: contents }` keeps the desktop grid identical.
    row.appendChild(el('div', { class: 'layer-cell', 'data-label': s.label }, cellNode || inp));
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
      // clear(false) / setDate(date, false): the `false` suppresses flatpickr's onChange so
      // programmatic updates don't get mistaken for user edits (which would set userSet).
      if (!v) { if (inp._flatpickr) inp._flatpickr.clear(false); else inp.value = ''; return; }
      // v is an ISO 'YYYY-MM-DD' string. flatpickr is configured with dateFormat 'd-M-Y',
      // so passing the ISO string would be mis-parsed — pass a real Date (local midnight).
      const dateObj = isoToLocalDate(v);
      if (inp._flatpickr) inp._flatpickr.setDate(dateObj, false);
      else inp.value = formatDDMMMYYYY(dateObj);
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
    if (!cascadingFromKey) return true;
    if (!cascadingToKey) return rows.every((r) => readVal(r.inputs[cascadingFromKey], fromKind) != null);
    return rows.every((r) =>
      readVal(r.inputs[cascadingFromKey], fromKind) != null &&
      readVal(r.inputs[cascadingToKey], toKind) != null
    );
  }
  function canAddNew() {
    if (rows.length === 0) return true;
    if (!cascadingFromKey) return true;
    if (!rowsComplete()) return false;
    const maturity = getMaturityValue();
    if (!cascadingToKey) {
      // from-only: can add while the last From is before maturity
      const lastFrom = readVal(rows[rows.length - 1].inputs[cascadingFromKey], fromKind);
      if (lastFrom && maturity && lastFrom >= maturity) return false;
      return true;
    }
    const lastTo = readVal(rows[rows.length - 1].inputs[cascadingToKey], toKind);
    if (lastTo && maturity && lastTo >= maturity) return false;
    return true;
  }

  function applyLayerRules() {
    if (!cascadingFromKey) return;
    const maturity = getMaturityValue();
    const anchor = getAnchor ? (getAnchor()?.value || null) : null;
    // From-only mode (no To column): a layer runs from its From until the day/month before the
    // next layer's From (the last layer extends to maturity). Used by Rate Revision layers.
    const fromOnly = !cascadingToKey;

    // Every field stays enabled (editable). Derived fields are AUTO-FILLED unless the
    // user has manually edited that specific cell (tracked via dataset.userSet, which is
    // only set by a genuine flatpickr/user change — never by programmatic setVal).
    rows.forEach((r, i) => {
      const fromInp = r.inputs[cascadingFromKey];
      const toInp = cascadingToKey ? r.inputs[cascadingToKey] : null;
      r.errors.clear();
      [fromInp, toInp].forEach(inp => inp && inp.classList.remove('field-error'));

      // FROM
      if (fromInp && fromInp.dataset.userSet !== '1') {
        if (i === 0) {
          // First row anchored to the external source (e.g. disbursement date / mora+1 month)
          setVal(fromInp, anchor, fromKind);
        } else if (!fromOnly) {
          // To-mode: cascade from the previous row's To + 1 unit (cleared if prev To empty)
          const prevTo = readVal(rows[i - 1].inputs[cascadingToKey], toKind);
          setVal(fromInp, prevTo ? advanceOne(prevTo, fromKind) : null, fromKind);
        }
        // from-only: each subsequent From is user-entered (no auto-fill).
      }

      // TO — only the LAST row auto-defaults to maturity (when not user-edited).
      if (!fromOnly) {
        const isLast = i === rows.length - 1;
        if (isLast && toInp && toInp.dataset.userSet !== '1') setVal(toInp, maturity, toKind);
      }
    });

    // Validate logical consistency
    rows.forEach((r, i) => {
      const fromInp = r.inputs[cascadingFromKey];
      const fv = readVal(fromInp, fromKind);
      if (fromOnly) {
        // Each From strictly increasing; every From (incl. the final layer) must be before maturity.
        if (fv && maturity && fv >= maturity) flagError(fromInp, r);
        if (i > 0) {
          const prevFrom = readVal(rows[i - 1].inputs[cascadingFromKey], fromKind);
          if (prevFrom && fv && fv <= prevFrom) flagError(fromInp, r);
        }
        return;
      }
      const toInp = r.inputs[cascadingToKey];
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
          if (s.allowEmpty) inp.appendChild(el('option', { value: '' }, s.placeholder ?? '— select —'));
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
export function toast(message, type = 'info', duration) {
  const root = document.getElementById('toast-root');
  const t = el('div', { class: `toast ${type}` });
  const msg = el('span', { class: 'toast-msg' }, message);
  const close = el('button', { class: 'toast-close', type: 'button', 'aria-label': 'Dismiss' }, '×');
  let timer = null;
  const dismiss = () => {
    if (timer) { clearTimeout(timer); timer = null; }
    t.style.transition = 'opacity 0.2s ease';
    t.style.opacity = '0';
    setTimeout(() => t.remove(), 220);
  };
  close.addEventListener('click', dismiss);
  t.appendChild(msg);
  t.appendChild(close);
  root.appendChild(t);
  // Warnings & errors persist until the user dismisses them via the × button; info/success
  // auto-dismiss after a comfortable delay but can still be closed early.
  const persist = type === 'warn' || type === 'warning' || type === 'error';
  if (!persist) timer = setTimeout(dismiss, duration != null ? duration : 5000);
  return t;
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
