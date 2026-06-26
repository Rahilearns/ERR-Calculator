// ERR Calculator — authentication + admin-approval API.
//
// Flow: a user registers with CIF + IDLC email + password (stored hashed, status "pending").
// An admin approves them from the admin portal. Once approved, the user logs in with CIF +
// password and receives a token the website sends on every request.
//
// Run:  cp .env.example .env  &&  (edit .env)  &&  npm install  &&  npm start
require('dotenv').config();
const express = require('express');
const cors = require('cors');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const store = require('./store');

const PORT = Number(process.env.PORT || 4000);
const JWT_SECRET = process.env.JWT_SECRET || '';
const TOKEN_TTL = process.env.TOKEN_TTL || '12h';
const ALLOWED_DOMAIN = (process.env.ALLOWED_EMAIL_DOMAIN || '').trim().toLowerCase();
const MIN_PW = Number(process.env.MIN_PASSWORD_LENGTH || 8);
const ORIGINS = (process.env.CORS_ORIGIN || '*').split(',').map((s) => s.trim()).filter(Boolean);

if (!JWT_SECRET || JWT_SECRET === 'change-me-to-a-long-random-string') {
  console.error('\n[FATAL] JWT_SECRET is not set. Copy .env.example to .env and set a long random JWT_SECRET.\n');
  process.exit(1);
}

store.load();

const app = express();
app.use(express.json({ limit: '64kb' }));
app.use(cors({
  origin: (origin, cb) => {
    if (!origin || ORIGINS.includes('*') || ORIGINS.includes(origin)) return cb(null, true);
    return cb(new Error('Origin not allowed by CORS: ' + origin));
  },
}));

// ---- helpers ----
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const publicUser = (u) => ({ id: u.id, cif: u.cif, email: u.email, role: u.role, status: u.status, createdAt: u.createdAt, approvedAt: u.approvedAt });

function signToken(u) {
  return jwt.sign({ sub: u.id, cif: u.cif, role: u.role }, JWT_SECRET, { expiresIn: TOKEN_TTL });
}
function auth(req, res, next) {
  const h = req.headers.authorization || '';
  const token = h.startsWith('Bearer ') ? h.slice(7) : null;
  if (!token) return res.status(401).json({ error: 'Not signed in.' });
  try {
    const payload = jwt.verify(token, JWT_SECRET);
    const u = store.findById(payload.sub);
    if (!u) return res.status(401).json({ error: 'Account not found.' });
    if (u.status !== 'approved') return res.status(403).json({ error: 'Account is not approved.' });
    req.user = u;
    next();
  } catch {
    return res.status(401).json({ error: 'Session expired — please sign in again.' });
  }
}
function adminOnly(req, res, next) {
  if (req.user.role !== 'admin') return res.status(403).json({ error: 'Admins only.' });
  next();
}

// Simple in-memory brute-force slowdown for login (per CIF). Resets on restart.
const attempts = new Map();
function tooManyAttempts(key) {
  const rec = attempts.get(key);
  if (!rec) return false;
  if (Date.now() > rec.until) { attempts.delete(key); return false; }
  return rec.count >= 8;
}
function noteFailure(key) {
  const rec = attempts.get(key) || { count: 0, until: 0 };
  rec.count += 1;
  rec.until = Date.now() + 15 * 60 * 1000; // 15-minute window
  attempts.set(key, rec);
}
function clearAttempts(key) { attempts.delete(key); }

// ---- routes ----
app.get('/api/health', (req, res) => res.json({ ok: true, users: store.count() }));

// Register → creates a PENDING account (no token issued; admin must approve).
app.post('/api/register', async (req, res) => {
  const cif = String(req.body.cif || '').trim();
  const email = String(req.body.email || '').trim().toLowerCase();
  const password = String(req.body.password || '');

  if (!cif) return res.status(400).json({ error: 'CIF is required.' });
  if (!EMAIL_RE.test(email)) return res.status(400).json({ error: 'Enter a valid email address.' });
  if (ALLOWED_DOMAIN && !email.endsWith('@' + ALLOWED_DOMAIN)) {
    return res.status(400).json({ error: `Email must be an @${ALLOWED_DOMAIN} address.` });
  }
  if (password.length < MIN_PW) return res.status(400).json({ error: `Password must be at least ${MIN_PW} characters.` });
  if (store.findByCif(cif)) return res.status(409).json({ error: 'An account with this CIF already exists.' });
  if (store.findByEmail(email)) return res.status(409).json({ error: 'An account with this email already exists.' });

  const passwordHash = await bcrypt.hash(password, 11);
  store.create({ cif, email, passwordHash, role: 'user', status: 'pending' });
  return res.status(201).json({ message: 'Registration received. You can sign in once an administrator approves your account.' });
});

// Login → only succeeds for APPROVED accounts.
app.post('/api/login', async (req, res) => {
  const cif = String(req.body.cif || '').trim();
  const password = String(req.body.password || '');
  const key = cif.toLowerCase();
  if (!cif || !password) return res.status(400).json({ error: 'Enter your CIF and password.' });
  if (tooManyAttempts(key)) return res.status(429).json({ error: 'Too many attempts. Try again in a few minutes.' });

  const u = store.findByCif(cif);
  const ok = u && await bcrypt.compare(password, u.passwordHash);
  if (!u || !ok) { noteFailure(key); return res.status(401).json({ error: 'Invalid CIF or password.' }); }
  clearAttempts(key);

  if (u.status === 'pending') return res.status(403).json({ error: 'Your account is awaiting administrator approval.' });
  if (u.status === 'rejected') return res.status(403).json({ error: 'Your registration was not approved. Contact the administrator.' });

  return res.json({ token: signToken(u), user: publicUser(u) });
});

app.get('/api/me', auth, (req, res) => res.json({ user: publicUser(req.user) }));

// ---- admin ----
// List users, optionally filtered by ?status=pending|approved|rejected
app.get('/api/admin/users', auth, adminOnly, (req, res) => {
  const status = req.query.status ? String(req.query.status) : null;
  res.json({ users: store.list(status).map(publicUser) });
});
app.post('/api/admin/approve', auth, adminOnly, (req, res) => {
  const u = store.findById(String(req.body.id || ''));
  if (!u) return res.status(404).json({ error: 'User not found.' });
  if (u.role === 'admin') return res.status(400).json({ error: 'Admin accounts are already active.' });
  res.json({ user: publicUser(store.setStatus(u.id, 'approved', req.user.cif)) });
});
app.post('/api/admin/reject', auth, adminOnly, (req, res) => {
  const u = store.findById(String(req.body.id || ''));
  if (!u) return res.status(404).json({ error: 'User not found.' });
  if (u.role === 'admin') return res.status(400).json({ error: 'Cannot reject an admin account.' });
  res.json({ user: publicUser(store.setStatus(u.id, 'rejected', req.user.cif)) });
});

// Seed the first admin account from env (idempotent).
async function seedAdmin() {
  const cif = (process.env.ADMIN_CIF || '').trim();
  const password = process.env.ADMIN_PASSWORD || '';
  const email = (process.env.ADMIN_EMAIL || 'admin@idlc.com').trim().toLowerCase();
  if (!cif || !password) { console.warn('[seed] ADMIN_CIF / ADMIN_PASSWORD not set — no admin created.'); return; }
  if (store.findByCif(cif)) return;
  const passwordHash = await bcrypt.hash(password, 11);
  store.create({ cif, email, passwordHash, role: 'admin', status: 'approved' });
  console.log(`[seed] Admin account created (CIF "${cif}"). Change its password after first login.`);
}

seedAdmin().then(() => {
  app.listen(PORT, () => {
    console.log(`ERR Calculator auth API listening on http://localhost:${PORT}`);
    console.log(`Allowed origins: ${ORIGINS.join(', ')}`);
  });
});
