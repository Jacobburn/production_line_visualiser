CREATE TABLE IF NOT EXISTS line_groups (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  name TEXT NOT NULL,
  display_order INTEGER NOT NULL DEFAULT 0,
  is_active BOOLEAN NOT NULL DEFAULT TRUE,
  created_by_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

ALTER TABLE production_lines
  ADD COLUMN IF NOT EXISTS group_id UUID REFERENCES line_groups(id) ON DELETE SET NULL;

CREATE UNIQUE INDEX IF NOT EXISTS idx_line_groups_name_active_unique
  ON line_groups (LOWER(name))
  WHERE is_active = TRUE;

CREATE INDEX IF NOT EXISTS idx_line_groups_display_order
  ON line_groups (display_order ASC, created_at ASC);

CREATE INDEX IF NOT EXISTS idx_production_lines_group_id
  ON production_lines (group_id);
