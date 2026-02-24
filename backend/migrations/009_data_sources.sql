CREATE TABLE IF NOT EXISTS data_sources (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  source_key TEXT NOT NULL UNIQUE,
  source_name TEXT NOT NULL,
  machine_no TEXT,
  device_name TEXT,
  device_id TEXT,
  scale_number TEXT,
  provider TEXT NOT NULL DEFAULT 'sql',
  is_active BOOLEAN NOT NULL DEFAULT TRUE,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

ALTER TABLE line_stages
  ADD COLUMN IF NOT EXISTS data_source_id UUID REFERENCES data_sources(id) ON DELETE SET NULL;

CREATE UNIQUE INDEX IF NOT EXISTS idx_line_stages_data_source_unique
  ON line_stages(data_source_id)
  WHERE data_source_id IS NOT NULL;

CREATE INDEX IF NOT EXISTS idx_data_sources_active_name
  ON data_sources(is_active, LOWER(source_name), created_at);

INSERT INTO data_sources (
  source_key,
  source_name,
  machine_no,
  device_name,
  device_id,
  scale_number,
  provider,
  is_active
)
VALUES
  (
    'bizerba-proseal-line-1',
    'ProSeal Line 1 - MASTER',
    '1',
    'ProSeal Line 1 - MASTER',
    '{2081B032-ECC1-4350-88A4-51867EBDFBDE}',
    '1',
    'sql',
    TRUE
  ),
  (
    'bizerba-multivac-line',
    'Multivac Line - MASTER',
    '3',
    'Multivac Line - MASTER',
    '{4341C2FD-EFBC-4b14-A0CC-3E59C1750D1E}',
    '1',
    'sql',
    TRUE
  ),
  (
    'bizerba-ulma-line',
    'Ulma Line - MASTER',
    '4',
    'Ulma Line -  MASTER',
    '{EE6A8E2F-9068-4bda-9DAF-83C81A322FA0}',
    '1',
    'sql',
    TRUE
  )
ON CONFLICT (source_key) DO UPDATE
SET
  source_name = EXCLUDED.source_name,
  machine_no = EXCLUDED.machine_no,
  device_name = EXCLUDED.device_name,
  device_id = EXCLUDED.device_id,
  scale_number = EXCLUDED.scale_number,
  provider = EXCLUDED.provider,
  is_active = EXCLUDED.is_active,
  updated_at = NOW();
