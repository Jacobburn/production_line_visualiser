ALTER TABLE production_lines
  ADD COLUMN IF NOT EXISTS device_name TEXT;

CREATE INDEX IF NOT EXISTS idx_production_lines_device_name_active
  ON production_lines (LOWER(device_name))
  WHERE is_active = TRUE
    AND device_name IS NOT NULL
    AND BTRIM(device_name) <> '';
