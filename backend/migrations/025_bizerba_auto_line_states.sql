CREATE TABLE IF NOT EXISTS bizerba_auto_line_states (
  line_id UUID PRIMARY KEY REFERENCES production_lines(id) ON DELETE CASCADE,
  device_name TEXT NOT NULL DEFAULT '',
  last_processed_bizerba_id BIGINT NOT NULL DEFAULT 0,
  pending_article_number TEXT NOT NULL DEFAULT '',
  pending_count INTEGER NOT NULL DEFAULT 0,
  active_run_log_id UUID REFERENCES run_logs(id) ON DELETE SET NULL,
  active_article_number TEXT NOT NULL DEFAULT '',
  active_product TEXT NOT NULL DEFAULT '',
  active_run_date DATE,
  active_start_timestamp TIMESTAMPTZ,
  active_last_date DATE,
  active_last_timestamp TIMESTAMPTZ,
  active_units_produced NUMERIC(14,2) NOT NULL DEFAULT 0,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  CONSTRAINT bizerba_auto_line_states_pending_count_check CHECK (pending_count >= 0),
  CONSTRAINT bizerba_auto_line_states_active_units_check CHECK (active_units_produced >= 0)
);

CREATE INDEX IF NOT EXISTS idx_bizerba_auto_line_states_updated_at
  ON bizerba_auto_line_states (updated_at DESC);

CREATE INDEX IF NOT EXISTS idx_bizerba_auto_line_states_active_run
  ON bizerba_auto_line_states (active_run_log_id)
  WHERE active_run_log_id IS NOT NULL;
