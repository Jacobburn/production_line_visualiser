# Staging Deploy Checklist

## Backend (Render/Railway/Fly/etc.)
1. Deploy `/backend` as a Node service.
2. Set env vars:
   - `NODE_ENV=production`
   - `API_PORT=4000`
   - `SUPABASE_DB_URL=<supabase postgres url>`
   - `DB_SSL=true`
   - `DB_SSL_REJECT_UNAUTHORIZED=false`
   - `JWT_SECRET=<strong secret>`
   - `JWT_EXPIRES_IN=12h`
   - `FRONTEND_ORIGINS=<comma-separated frontend urls>`
     - Example: `FRONTEND_ORIGINS=https://clever-squirrel-e035dc.netlify.app,https://production-line-supervisor-app.netlify.app`
3. Run migrations:
   - `npm --prefix backend run migrate`
4. Seed once (optional for test users):
   - `npm --prefix backend run seed`

## Frontend (Netlify/Vercel/Static host)
1. Deploy repo root as static site (`index.html`, `app.js`, `styles.css`).
2. After deploy, set API base in browser console on staging URL:
   - `localStorage.setItem("production-line-api-base","https://YOUR-STAGING-API")`
   - `location.reload()`

## Smoke test
1. `GET /api/health` from deployed API.
2. Supervisor login works.
3. Add shift/run/downtime from supervisor UI.
4. Confirm manager data tab sees those rows.
