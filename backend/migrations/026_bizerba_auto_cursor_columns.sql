ALTER TABLE bizerba_auto_line_states
  ADD COLUMN IF NOT EXISTS last_processed_timestamp TIMESTAMPTZ;

ALTER TABLE bizerba_auto_line_states
  ADD COLUMN IF NOT EXISTS last_processed_bizerba_id_text TEXT NOT NULL DEFAULT '';
