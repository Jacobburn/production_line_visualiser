ALTER TABLE run_logs
  ADD COLUMN IF NOT EXISTS run_crewing_pattern JSONB NOT NULL DEFAULT '{}'::jsonb;
