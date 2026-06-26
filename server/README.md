# ERR Calculator — Auth Server

Small backend that adds **user login with admin approval** to the ERR Calculator:

1. A user **registers** with their **CIF**, **IDLC email**, and a **password** → account is saved as **pending**.
2. An **admin** approves (or rejects) the account from the admin portal in the website.
3. Once approved, the user **logs in with CIF + password** and gets a session token the site sends on every request.

Passwords are stored **hashed** (bcrypt). The site never sees or stores plaintext passwords.

---

## Run it (local)

Requires **Node.js 18+**.

```bash
cd server
cp .env.example .env          # then edit .env (see below)
npm install
npm start
```

It prints e.g. `... listening on http://localhost:4000`.

### Configure `.env`
- `JWT_SECRET` — **required.** A long random string. Generate one:
  `node -e "console.log(require('crypto').randomBytes(48).toString('hex'))"`
- `CORS_ORIGIN` — comma‑separated list of website origins allowed to call the API
  (your live site + your local dev URL). `*` allows any origin (local dev only).
- `ADMIN_CIF` / `ADMIN_PASSWORD` / `ADMIN_EMAIL` — the first admin account, created automatically on
  first start. Log in with this CIF + password to reach the admin portal. **Change the password after first login.**
- `ALLOWED_EMAIL_DOMAIN` — restrict registration emails to this domain (e.g. `idlc.com`); blank = any valid email.
- `MIN_PASSWORD_LENGTH` — minimum password length (default 8).

### Point the website at it
In the frontend, set `API_BASE_URL` in **`js/config.js`** to this server's URL
(e.g. `http://localhost:4000` for local, or the IDLC‑hosted URL in production).
While `API_BASE_URL` is blank, the website skips login entirely and works as before.

---

## API

| Method | Path | Who | Body | Purpose |
|---|---|---|---|---|
| GET | `/api/health` | anyone | — | liveness check |
| POST | `/api/register` | anyone | `{ cif, email, password }` | create a pending account |
| POST | `/api/login` | anyone | `{ cif, password }` | returns `{ token, user }` if approved |
| GET | `/api/me` | logged in | — | current user |
| GET | `/api/admin/users?status=pending` | admin | — | list users (optional status filter) |
| POST | `/api/admin/approve` | admin | `{ id }` | approve a user |
| POST | `/api/admin/reject` | admin | `{ id }` | reject a user |

Send the token as `Authorization: Bearer <token>`.

---

## Data store & moving to a real database

For zero‑setup, users are kept in `server/data/users.json` (git‑ignored — it holds password hashes).
That's fine for a handful of users and for testing the whole flow. For production, reimplement the
small interface in **`store.js`** (`findByCif`, `findByEmail`, `findById`, `create`, `setStatus`, `list`,
`count`) against a proper database (e.g. Postgres). Nothing else in the server changes.

> Note: the in‑memory login throttle and JSON store assume a **single** server process. If IDLC runs
> multiple instances behind a load balancer, move both to the shared database.
