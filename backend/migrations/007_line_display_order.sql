ALTER TABLE production_lines
  ADD COLUMN IF NOT EXISTS display_order INTEGER NOT NULL DEFAULT 0;

WITH ordered_lines AS (
  SELECT
    id,
    ROW_NUMBER() OVER (PARTITION BY group_id ORDER BY created_at ASC, id ASC) - 1 AS next_display_order
  FROM production_lines
)
UPDATE production_lines AS line
SET display_order = ordered_lines.next_display_order
FROM ordered_lines
WHERE line.id = ordered_lines.id;

CREATE INDEX IF NOT EXISTS idx_production_lines_group_display_order
  ON production_lines (group_id, display_order ASC, created_at ASC)
  WHERE is_active = TRUE;
