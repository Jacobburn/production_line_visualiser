# Production Backend Migration Plan

## Goal
Move from localStorage-only persistence to a production backend with safe rollout and a permanent test environment.

## Environments
- `local`: Docker Postgres + local Node API + local frontend.
- `staging`: cloud Postgres + deployed API + preview frontend (used for QA/UAT).
- `production`: cloud Postgres + deployed API + live frontend.

Yes, you keep a test environment (`staging`) for ongoing feature work and safe releases.

## Rollout phases

### Phase 1: Backend foundation (in progress)
- Database schema + migrations
- Auth (manager/supervisor)
- Core logging endpoints
- Audit events

### Phase 2: Dual-write bridge
- Keep current UI behavior
- On create/edit/log actions:
  - write to backend first (or in parallel)
  - keep localStorage fallback
- Add API connectivity status banner in UI

### Phase 3: Read path migration
- Switch initial app load to backend data
- Keep localStorage import fallback for one release window
- Add one-click local data import to backend

### Phase 4: Hard cutover
- Disable local-only writes
- Keep export/backup tooling
- Enforce auth for all data operations

### Phase 5: Production hardening
- Rate limits
- Centralized logs + alerting
- Daily backups + restore drill
- Sentry/error monitoring

## Immediate next tasks
1. Add frontend API client and auth session handling.
2. Wire supervisor/manager log forms to backend endpoints.
3. Add line/stage/layout save endpoints and connect builder/editor.
4. Add staging deployment pipeline (API + frontend + database migrations).
