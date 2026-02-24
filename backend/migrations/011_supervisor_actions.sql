CREATE TABLE IF NOT EXISTS supervisor_actions (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  supervisor_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
  supervisor_username TEXT NOT NULL,
  supervisor_name TEXT NOT NULL DEFAULT '',
  line_id UUID REFERENCES production_lines(id) ON DELETE SET NULL,
  title TEXT NOT NULL,
  description TEXT NOT NULL DEFAULT '',
  priority TEXT NOT NULL DEFAULT 'Medium' CHECK (priority IN ('Low', 'Medium', 'High', 'Critical')),
  status TEXT NOT NULL DEFAULT 'Open' CHECK (status IN ('Open', 'In Progress', 'Blocked', 'Completed')),
  due_date DATE,
  related_equipment_stage_id UUID REFERENCES line_stages(id) ON DELETE SET NULL,
  related_reason_category TEXT,
  related_reason_detail TEXT,
  created_by_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
  created_by_name TEXT NOT NULL DEFAULT '',
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_supervisor_actions_assignee
  ON supervisor_actions (LOWER(supervisor_username), created_at DESC);

CREATE INDEX IF NOT EXISTS idx_supervisor_actions_line
  ON supervisor_actions (line_id, created_at DESC);

CREATE INDEX IF NOT EXISTS idx_supervisor_actions_due_date
  ON supervisor_actions (due_date ASC NULLS LAST, created_at DESC);
