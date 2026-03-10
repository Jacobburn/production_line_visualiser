ALTER TABLE shift_logs DROP CONSTRAINT IF EXISTS shift_logs_shift_check;
ALTER TABLE shift_break_logs DROP CONSTRAINT IF EXISTS shift_break_logs_shift_check;

ALTER TABLE shift_logs
  ADD CONSTRAINT shift_logs_shift_check
  CHECK (shift IN ('Day', 'Night', 'Full Day'));

ALTER TABLE shift_break_logs
  ADD CONSTRAINT shift_break_logs_shift_check
  CHECK (shift IN ('Day', 'Night', 'Full Day'));
