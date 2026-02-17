CREATE TABLE IF NOT EXISTS shift_break_logs (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  shift_log_id UUID NOT NULL REFERENCES shift_logs(id) ON DELETE CASCADE,
  line_id UUID NOT NULL REFERENCES production_lines(id) ON DELETE CASCADE,
  date DATE NOT NULL,
  shift TEXT NOT NULL CHECK (shift IN ('Day', 'Night', 'Full Day')),
  break_start TIME NOT NULL,
  break_finish TIME,
  submitted_by_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
  submitted_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_shift_break_logs_shift_log ON shift_break_logs(shift_log_id);
CREATE INDEX IF NOT EXISTS idx_shift_break_logs_line_date_shift ON shift_break_logs(line_id, date, shift);
