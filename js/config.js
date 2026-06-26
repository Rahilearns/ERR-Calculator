// ── Site configuration ────────────────────────────────────────────────────────
// Where the authentication API lives (the server in /server). Examples:
//   ''                         → login DISABLED; the calculator works for everyone (current behaviour)
//   'http://localhost:4000'    → local testing against the bundled server
//   'https://auth.idlc.com'    → production, once IDLC hosts the backend
// No trailing slash.
export const API_BASE_URL = '';

// Derived: when an API base is set, the site requires login.
export const AUTH_ENABLED = !!API_BASE_URL;
