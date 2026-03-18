ALTER TABLE public."Bizerba_ID"
  ADD COLUMN IF NOT EXISTS source_id TEXT;

DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1
    FROM pg_constraint
    WHERE conname = 'bizerba_id_source_id_key'
      AND conrelid = 'public."Bizerba_ID"'::regclass
  ) THEN
    ALTER TABLE public."Bizerba_ID"
      ADD CONSTRAINT bizerba_id_source_id_key UNIQUE (source_id);
  END IF;
END
$$;
