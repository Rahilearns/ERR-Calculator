// User store. Backed by a JSON file (server/data/users.json) so it runs with zero setup.
// The interface below is deliberately small and DB-shaped: to move to Postgres/MySQL later,
// reimplement these functions against your database and keep the same signatures — nothing
// else in the server needs to change. The data file holds password HASHES, never plaintext,
// and is git-ignored.
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const DATA_DIR = path.join(__dirname, 'data');
const DATA_FILE = path.join(DATA_DIR, 'users.json');

let users = [];

function load() {
  try {
    if (fs.existsSync(DATA_FILE)) {
      users = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8')) || [];
    } else {
      fs.mkdirSync(DATA_DIR, { recursive: true });
      users = [];
      persist();
    }
  } catch (err) {
    console.error('[store] failed to load users.json — starting empty:', err.message);
    users = [];
  }
}
function persist() {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(DATA_FILE, JSON.stringify(users, null, 2));
}

const normCif = (c) => String(c == null ? '' : c).trim();
const normEmail = (e) => String(e == null ? '' : e).trim().toLowerCase();

function findByCif(cif) {
  const c = normCif(cif).toLowerCase();
  return users.find((u) => u.cif.toLowerCase() === c) || null;
}
function findByEmail(email) {
  const e = normEmail(email);
  return users.find((u) => u.email === e) || null;
}
function findById(id) {
  return users.find((u) => u.id === id) || null;
}

// data: { cif, email, passwordHash, role, status }
function create(data) {
  const now = new Date().toISOString();
  const user = {
    id: crypto.randomUUID(),
    cif: normCif(data.cif),
    email: normEmail(data.email),
    passwordHash: data.passwordHash,
    role: data.role || 'user',
    status: data.status || 'pending',
    createdAt: now,
    approvedAt: null,
    approvedBy: null,
  };
  users.push(user);
  persist();
  return user;
}

function setStatus(id, status, approvedBy) {
  const u = findById(id);
  if (!u) return null;
  u.status = status;
  u.approvedAt = (status === 'approved' || status === 'rejected') ? new Date().toISOString() : null;
  u.approvedBy = approvedBy || null;
  persist();
  return u;
}

function list(status) {
  const rows = status ? users.filter((u) => u.status === status) : users.slice();
  return rows.sort((a, b) => (a.createdAt < b.createdAt ? 1 : -1)); // newest first
}

function count() {
  return users.length;
}

module.exports = { load, findByCif, findByEmail, findById, create, setStatus, list, count };
