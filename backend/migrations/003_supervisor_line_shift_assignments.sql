CREATE TABLE IF NOT EXISTS supervisor_line_shift_assignments (
  supervisor_user_id UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  line_id UUID NOT NULL REFERENCES production_lines(id) ON DELETE CASCADE,
  shift TEXT NOT NULL CHECK (shift IN ('Day', 'Night')),
  assigned_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  assigned_by_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
  PRIMARY KEY (supervisor_user_id, line_id, shift)
);

INSERT INTO supervisor_line_shift_assignments(supervisor_user_id, line_id, shift, assigned_by_user_id)
SELECT a.supervisor_user_id, a.line_id, s.shift, a.assigned_by_user_id
FROM supervisor_line_assignments a
CROSS JOIN (VALUES ('Day'), ('Night')) AS s(shift)
ON CONFLICT (supervisor_user_id, line_id, shift) DO NOTHING;

CREATE INDEX IF NOT EXISTS idx_shift_assignments_supervisor_line
  ON supervisor_line_shift_assignments(supervisor_user_id, line_id);

CREATE INDEX IF NOT EXISTS idx_shift_assignments_line
  ON supervisor_line_shift_assignments(line_id);
