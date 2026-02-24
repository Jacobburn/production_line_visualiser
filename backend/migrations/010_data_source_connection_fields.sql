ALTER TABLE data_sources
  ADD COLUMN IF NOT EXISTS connection_mode TEXT NOT NULL DEFAULT 'api';

ALTER TABLE data_sources
  ADD COLUMN IF NOT EXISTS api_base_url TEXT;

ALTER TABLE data_sources
  ADD COLUMN IF NOT EXISTS api_key TEXT;

ALTER TABLE data_sources
  ADD COLUMN IF NOT EXISTS api_secret TEXT;

ALTER TABLE data_sources
  ADD COLUMN IF NOT EXISTS sql_host TEXT;

ALTER TABLE data_sources
  ADD COLUMN IF NOT EXISTS sql_port INTEGER;

ALTER TABLE data_sources
  ADD COLUMN IF NOT EXISTS sql_database TEXT;

ALTER TABLE data_sources
  ADD COLUMN IF NOT EXISTS sql_username TEXT;

ALTER TABLE data_sources
  ADD COLUMN IF NOT EXISTS sql_password TEXT;

UPDATE data_sources
SET connection_mode = 'sql'
WHERE provider = 'sql'
  AND (connection_mode IS NULL OR connection_mode = '' OR connection_mode = 'api');
