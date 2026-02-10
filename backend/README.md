# Backend (Production Line API)

## 1) Prerequisites
- Node.js 20+
- npm 10+
- Docker Desktop (recommended for local Postgres)

## 2) Start local infrastructure
From the project root:

```bash
docker compose up -d postgres
```

## 3) Configure backend env
Copy env template:

```bash
cp backend/.env.example backend/.env
```

Default local DB is already set to `postgresql://postgres:postgres@localhost:5432/production_line`.

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
