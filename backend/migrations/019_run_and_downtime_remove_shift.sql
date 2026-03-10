ALTER TABLE run_logs DROP CONSTRAINT IF EXISTS run_logs_shift_check;
ALTER TABLE downtime_logs DROP CONSTRAINT IF EXISTS downtime_logs_shift_check;

DROP INDEX IF EXISTS idx_run_logs_line_date_shift;
DROP INDEX IF EXISTS idx_down_logs_line_date_shift;

ALTER TABLE run_logs
  DROP COLUMN IF EXISTS shift;

ALTER TABLE downtime_logs
  DROP COLUMN IF EXISTS shift;

CREATE INDEX IF NOT EXISTS idx_run_logs_line_date ON run_logs(line_id, date);
CREATE INDEX IF NOT EXISTS idx_down_logs_line_date ON downtime_logs(line_id, date);
