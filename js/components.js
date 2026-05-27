// Reusable UI component builders (returns DOM nodes)
import { attachCommaFormatter } from './formatting.js';

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

// Info-icon tooltip: small (i) icon with hover text
export function infoIcon(text) {
  const ic = el('span', { class: 'info-icon', title: text }, 'i');
  return ic;
}

// Label with optional info tooltip
function labeled(field, label, tooltip) {
  const wrap = el('div', { class: 'label-row' }, el('span', {}, label));
  if (tooltip) wrap.appendChild(infoIcon(tooltip));
  field.insertBefore(wrap, field.firstChild);
  return field;
}

// Number field — right-aligned, 2-decimal blur formatting (unless integerOnly)
export function numberField({ label, name, placeholder = '', help = '', integerOnly = false, min = null, tooltip = '' }) {
  const id = nextId();
  const input = el('input', {
    id, type: 'text', inputmode: integerOnly ? 'numeric' : 'decimal',
    placeholder, autocomplete: 'off', 'data-name': name,
    class: 'numeric-input',
  });
  attachCommaFormatter(input, { integerOnly });
  if (!integerOnly) {
    input.addEventListener('blur', () => {
      const raw = input.value.replace(/,/g, '');
      if (raw === '' || raw === '.') return;
      const n = Number(raw);
      if (isNaN(n)) return;
      input.value = n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    });
  }
  const field = el('div', { class: 'field' });
  const lbl = el('label', { for: id }, label);
  if (tooltip) lbl.appendChild(infoIcon(tooltip));
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
  field.setLabel = (text) => {
    const l = field.querySelector('label');
    if (l) {
      l.firstChild.nodeValue = text;
      // re-append tooltip after rename
      const ic = l.querySelector('.info-icon');
      if (ic) { l.appendChild(ic); }
    }
  };
  return field;
}

// Percent field — % fixed at far-right, right-aligned input, 2-decimal blur
export function percentField({ label, name, placeholder = '', help = '', tooltip = '' }) {
  const id = nextId();
  const wrapper = el('div', { class: 'field percent-field' });
  const lbl = el('label', { for: id }, label);
  if (tooltip) lbl.appendChild(infoIcon(tooltip));
  wrapper.appendChild(lbl);

  const inputWrap = el('div', { class: 'percent-input-fixed' });
  const input = el('input', {
    id, type: 'text', inputmode: 'decimal',
    placeholder, autocomplete: 'off', 'data-name': name,
    class: 'numeric-input',
  });
  const suffix = el('span', { class: 'percent-suffix-fixed' }, '%');
  inputWrap.appendChild(input);
  inputWrap.appendChild(suffix);
  wrapper.appendChild(inputWrap);
  if (help) wrapper.appendChild(el('span', { class: 'help' }, help));

  function sanitize(v) {
    v = v.replace(/[^0-9.]/g, '');
    const firstDot = v.indexOf('.');
    if (firstDot !== -1) v = v.slice(0, firstDot + 1) + v.slice(firstDot + 1).replace(/\./g, '');
    return v;
  }
  function refreshEmpty() { inputWrap.classList.toggle('empty', input.value === ''); }
  input.addEventListener('input', (e) => { e.target.value = sanitize(e.target.value); refreshEmpty(); });
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
  wrapper.setLabel = (text) => {
    const l = wrapper.querySelector('label');
    if (l) {
      l.firstChild.nodeValue = text;
      const ic = l.querySelector('.info-icon');
      if (ic) l.appendChild(ic);
    }
  };
  return wrapper;
}

// Option field as native <select> — center-aligned text
export function optionField({ label, name, options, value = null, onChange = null, help = '', tooltip = '' }) {
  const id = nextId();
  const wrapper = el('div', { class: 'field' });
  const lbl = el('label', { for: id }, label);
  if (tooltip) lbl.appendChild(infoIcon(tooltip));
  wrapper.appendChild(lbl);
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
  wrapper.setLabel = (text) => {
    const l = wrapper.querySelector('label');
    if (l) {
      l.firstChild.nodeValue = text;
      const ic = l.querySelector('.info-icon');
      if (ic) l.appendChild(ic);
    }
  };
  wrapper.select = select;
  return wrapper;
}

// Date field — flatpickr-powered, format DD-MMM-YYYY
export function dateField({ label, name, placeholder = 'dd-Mmm-yyyy', tooltip = '', onChange = null, disableFn = null }) {
  const id = nextId();
  const input = el('input', { id, type: 'text', 'data-name': name, placeholder, autocomplete: 'off', class: 'centered-input date-input' });
  const field = el('div', { class: 'field' });
  const lbl = el('label', { for: id }, label);
  if (tooltip) lbl.appendChild(infoIcon(tooltip));
  field.appendChild(lbl);
  field.appendChild(input);

  let fp;
  requestAnimationFrame(() => {
    if (typeof flatpickr === 'undefined') return;
    fp = flatpickr(input, {
      dateFormat: 'd-M-Y',
      altInput: false,
      allowInput: true,
      disable: disableFn ? [disableFn] : [],
      onChange: () => { if (onChange) onChange(input.value); },
    });
    field._fp = fp;
  });

  field.getValue = () => {
    // Convert DD-MMM-YYYY -> ISO YYYY-MM-DD
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

export function parseDDMMMYYYY(s) {
  if (!s) return null;
  // Try ISO first
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

// Month boxes — selectable equal-width boxes, max 6 per row (always visible, no dropdown)
export function monthBoxesField({ name, getCount, tooltip = '', label = '' }) {
  const wrapper = el('div', { class: 'field' });
  if (label) {
    const lbl = el('label', {}, label);
    if (tooltip) lbl.appendChild(infoIcon(tooltip));
    wrapper.appendChild(lbl);
  }
  const grid = el('div', { class: 'month-boxes', 'data-name': name });
  wrapper.appendChild(grid);

  let states = [];
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
      });
      grid.appendChild(box);
    }
  }
  wrapper.refresh = render;
  wrapper.getValue = () => states.slice();
  wrapper.setValue = (arr) => { states = (arr || []).slice(); render(); };
  render();
  return wrapper;
}

// Layered field — header center-aligned with cells; right-align numeric/percent; center select/date
export function layeredField({ label, name, schema, addLabel = '+ Add layer', help = '', initialRows = 1, onChange = null }) {
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
  function syncHeaderVisibility() {
    header.classList.toggle('hidden', rows.length === 0);
  }

  function fireChange() { if (onChange) onChange(wrapper.getValue(), rows); }

  function addRow(values = {}) {
    const row = el('div', { class: 'layer-row', style: `grid-template-columns: ${cols}` });
    const inputs = {};
    schema.forEach((s) => {
      let inp;
      if (s.type === 'option') {
        inp = el('select', { class: 'centered-input' });
        const opts = typeof s.options === 'function' ? s.options() : s.options;
        // Allow empty placeholder option
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
        // attach flatpickr after mount
        requestAnimationFrame(() => {
          if (typeof flatpickr === 'undefined') return;
          flatpickr(inp, { dateFormat: 'd-M-Y', allowInput: true, onChange: fireChange });
        });
        inp.addEventListener('change', fireChange);
      } else if (s.type === 'percent') {
        const pwrap = el('div', { class: 'percent-input-fixed inline' });
        inp = el('input', { type: 'text', inputmode: 'decimal', class: 'numeric-input' });
        const sfx = el('span', { class: 'percent-suffix-fixed' }, '%');
        pwrap.appendChild(inp);
        pwrap.appendChild(sfx);
        if (values[s.key] !== undefined && values[s.key] !== null) inp.value = (values[s.key] * 100).toFixed(2);
        const refreshEmpty = () => pwrap.classList.toggle('empty', inp.value === '');
        inp.addEventListener('input', () => {
          let v = inp.value.replace(/[^0-9.]/g, '');
          const firstDot = v.indexOf('.');
          if (firstDot !== -1) v = v.slice(0, firstDot + 1) + v.slice(firstDot + 1).replace(/\./g, '');
          inp.value = v;
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
      } else { // number
        inp = el('input', { type: 'text', inputmode: 'decimal', placeholder: '0', class: 'numeric-input' });
        if (values[s.key] !== undefined && values[s.key] !== null) {
          inp.value = Number(values[s.key]).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        }
        attachCommaFormatter(inp, { integerOnly: !!s.integerOnly });
        if (!s.integerOnly) {
          inp.addEventListener('blur', () => {
            const raw = inp.value.replace(/,/g, '');
            if (raw === '' || raw === '.') return;
            const n = Number(raw);
            if (isNaN(n)) return;
            inp.value = n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
          });
        }
        inp.addEventListener('input', fireChange);
      }
      inputs[s.key] = inp;
      row.appendChild(inp);
    });
    const del = el('button', { type: 'button', class: 'row-del', title: 'Remove' }, '×');
    del.addEventListener('click', () => {
      const idx = rows.indexOf(rowApi);
      if (idx >= 0) {
        rows.splice(idx, 1);
        row.remove();
        syncHeaderVisibility();
        fireChange();
      }
    });
    row.appendChild(del);
    layers.appendChild(row);
    const rowApi = { row, inputs };
    rows.push(rowApi);
    syncHeaderVisibility();
    return rowApi;
  }

  const addBtn = el('button', { type: 'button', class: 'add-layer' }, addLabel);
  addBtn.addEventListener('click', () => {
    if (wrapper.canAdd && !wrapper.canAdd()) return;
    addRow();
    fireChange();
  });
  wrapper.appendChild(addBtn);

  for (let i = 0; i < initialRows; i++) addRow();
  syncHeaderVisibility();

  wrapper.getValue = () => rows.map(({ inputs }) => {
    const row = {};
    schema.forEach((s) => {
      const v = inputs[s.key].value;
      if (s.type === 'percent') row[s.key] = v === '' ? null : Number(v) / 100;
      else if (s.type === 'number') {
        const raw = v.replace(/,/g, '');
        row[s.key] = raw === '' ? null : Number(raw);
      } else if (s.type === 'date') {
        // Stored as ISO YYYY-MM-DD for calc
        const d = parseDDMMMYYYY(v);
        row[s.key] = d ? d.toISOString().slice(0, 10) : null;
      } else row[s.key] = v || null;
    });
    return row;
  });
  wrapper.setValue = (arr) => {
    rows.length = 0;
    layers.querySelectorAll('.layer-row').forEach(r => r.remove());
    (arr || [{}]).forEach(v => addRow(v));
    syncHeaderVisibility();
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
