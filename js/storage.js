// localStorage-backed saved summaries (max 5) + draft state per tab

const KEY = 'err_summaries_v1';
const DRAFT_KEY = 'err_drafts_v1';
const MAX = 5;

export function listSummaries() {
  try { return JSON.parse(localStorage.getItem(KEY) || '[]'); } catch { return []; }
}

function fingerprint(summary) {
  const obj = { pageType: summary.pageType, inputs: summary.inputs };
  return JSON.stringify(obj, Object.keys(obj).sort());
}

// Save iff fingerprint not already present. Silently no-op on duplicate.
// Returns { saved: bool, reason: 'duplicate'|'full'|null, summary }.
export function saveSummary(summary) {
  const all = listSummaries();
  const fp = fingerprint(summary);
  if (all.some(s => s._fp === fp || fingerprint(s) === fp)) {
    return { saved: false, reason: 'duplicate' };
  }
  if (all.length >= MAX) {
    return { saved: false, reason: 'full' };
  }
  summary.id = 's' + Date.now() + '_' + Math.random().toString(36).slice(2, 6);
  summary.savedAt = new Date().toISOString();
  summary._fp = fp;
  all.push(summary);
  localStorage.setItem(KEY, JSON.stringify(all));
  return { saved: true, summary };
}

export function deleteSummary(id) {
  const all = listSummaries().filter(s => s.id !== id);
  localStorage.setItem(KEY, JSON.stringify(all));
}

export function clearSummaries() { localStorage.removeItem(KEY); }
export function getMax() { return MAX; }

// ---- Draft preservation (per-tab in-progress inputs) ----
export function saveDraft(tabKey, state) {
  let all = {};
  try { all = JSON.parse(localStorage.getItem(DRAFT_KEY) || '{}'); } catch {}
  all[tabKey] = state;
  localStorage.setItem(DRAFT_KEY, JSON.stringify(all));
}
export function loadDraft(tabKey) {
  try { return JSON.parse(localStorage.getItem(DRAFT_KEY) || '{}')[tabKey] || null; } catch { return null; }
}
export function clearDraft(tabKey) {
  let all = {};
  try { all = JSON.parse(localStorage.getItem(DRAFT_KEY) || '{}'); } catch {}
  delete all[tabKey];
  localStorage.setItem(DRAFT_KEY, JSON.stringify(all));
}
