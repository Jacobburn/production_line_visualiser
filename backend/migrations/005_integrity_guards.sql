-- Align any historical break rows to their parent shift metadata.
UPDATE shift_break_logs b
SET
  line_id = s.line_id,
  date = s.date,
  shift = s.shift
FROM shift_logs s
WHERE s.id = b.shift_log_id
  AND (
    b.line_id IS DISTINCT FROM s.line_id
    OR b.date IS DISTINCT FROM s.date
    OR b.shift IS DISTINCT FROM s.shift
  );

-- Resolve pre-existing race duplicates by keeping one open break per shift log.
WITH ranked_open_breaks AS (
  SELECT
    id,
    ROW_NUMBER() OVER (
      PARTITION BY shift_log_id
      ORDER BY created_at ASC, id ASC
    ) AS rn
  FROM shift_break_logs
  WHERE break_finish IS NULL
)
UPDATE shift_break_logs b
SET break_finish = b.break_start
FROM ranked_open_breaks r
WHERE b.id = r.id
  AND r.rn > 1;

CREATE UNIQUE INDEX IF NOT EXISTS idx_shift_break_logs_one_open_per_shift
  ON shift_break_logs(shift_log_id)
  WHERE break_finish IS NULL;

-- Clean up any historical downtime rows with cross-line stage references.
UPDATE downtime_logs d
SET equipment_stage_id = NULL
WHERE d.equipment_stage_id IS NOT NULL
  AND NOT EXISTS (
    SELECT 1
    FROM line_stages s
    WHERE s.id = d.equipment_stage_id
      AND s.line_id = d.line_id
  );

CREATE OR REPLACE FUNCTION validate_shift_break_consistency()
RETURNS trigger AS $$
BEGIN
  IF NOT EXISTS (
    SELECT 1
    FROM shift_logs s
    WHERE s.id = NEW.shift_log_id
      AND s.line_id = NEW.line_id
      AND s.date = NEW.date
      AND s.shift = NEW.shift
  ) THEN
    RAISE EXCEPTION 'shift_break_logs row must match parent shift log metadata'
      USING ERRCODE = '23514';
  END IF;

  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_validate_shift_break_consistency ON shift_break_logs;
CREATE TRIGGER trg_validate_shift_break_consistency
BEFORE INSERT OR UPDATE OF shift_log_id, line_id, date, shift
ON shift_break_logs
FOR EACH ROW
EXECUTE FUNCTION validate_shift_break_consistency();

CREATE OR REPLACE FUNCTION validate_downtime_stage_scope()
RETURNS trigger AS $$
BEGIN
  IF NEW.equipment_stage_id IS NULL THEN
    RETURN NEW;
  END IF;

  IF NOT EXISTS (
    SELECT 1
    FROM line_stages s
    WHERE s.id = NEW.equipment_stage_id
      AND s.line_id = NEW.line_id
  ) THEN
    RAISE EXCEPTION 'equipment stage does not belong to this line'
      USING ERRCODE = '23514';
  END IF;

  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_validate_downtime_stage_scope ON downtime_logs;
CREATE TRIGGER trg_validate_downtime_stage_scope
BEFORE INSERT OR UPDATE OF line_id, equipment_stage_id
ON downtime_logs
FOR EACH ROW
EXECUTE FUNCTION validate_downtime_stage_scope();
