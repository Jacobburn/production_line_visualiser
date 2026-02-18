# Backend (Production Line API)

## 1) Prerequisites
- Node.js 20+
- npm 10+
- Docker Desktop (optional, for local Postgres)
- Supabase project (recommended for staging/production)

## 2) Choose database mode

### Option A: Local Postgres (docker)
From the project root:

```bash
docker compose up -d postgres
```

### Option B: Supabase Postgres
1. Create a Supabase project.
2. In Supabase, open Project Settings -> Database.
3. Copy a Postgres connection string (direct or session pooler).
4. Use that value as `SUPABASE_DB_URL` in `backend/.env`.

## 3) Configure backend env
Copy env template:

```bash
cp backend/.env.example backend/.env
```

For local docker Postgres, `DATABASE_URL` already points to:
`postgresql://postgres:postgres@localhost:5432/production_line`.

For Supabase, set:
- `SUPABASE_DB_URL=<your supabase postgres connection string>`
- `DB_SSL=true`
- `DB_SSL_REJECT_UNAUTHORIZED=true`

Always set a strong JWT secret in `backend/.env`:
- `JWT_SECRET=<random 32+ character secret>`

Optional resilience/timeouts:
- `DB_CONNECTION_TIMEOUT_MS=10000`
- `DB_QUERY_TIMEOUT_MS=15000`
- `DB_STATEMENT_TIMEOUT_MS=20000`
- `DB_MAX_RETRIES=1` (read-only query retries)
- `DB_RETRY_DELAY_MS=250`
- `HTTP_REQUEST_TIMEOUT_MS=30000`
- `HTTP_HEADERS_TIMEOUT_MS=35000`
- `HTTP_KEEP_ALIVE_TIMEOUT_MS=5000`

## 4) Install backend dependencies

```bash
cd backend
npm install
cd ..
```

## 5) Run migrations + seed

```bash
npm run backend:migrate
npm run backend:seed
```

## 6) Run API

```bash
npm run backend:dev
```

Health check:

```bash
curl http://localhost:4000/api/health
```

## 7) Seed login users
- Manager: `manager` / `manager123`
- Supervisor: `supervisor` / `supervisor123`

## API routes (current)
- `POST /api/auth/login`
- `GET /api/me`
- `GET /api/lines`
- `POST /api/lines` (manager only)
- `GET /api/lines/:lineId`
- `POST /api/logs/shifts`
- `POST /api/logs/runs`
- `POST /api/logs/downtime`
- `GET /api/health`
