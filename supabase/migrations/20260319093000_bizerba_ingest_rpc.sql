CREATE OR REPLACE FUNCTION public.ingest_bizerba_rows(batch_rows JSONB)
RETURNS TABLE(inserted_count INTEGER, skipped_count INTEGER)
LANGUAGE plpgsql
SET search_path = public
AS $$
DECLARE
  total_count INTEGER := 0;
BEGIN
  IF batch_rows IS NULL OR jsonb_typeof(batch_rows) <> 'array' THEN
    RAISE EXCEPTION 'batch_rows must be a JSON array';
  END IF;

  total_count := jsonb_array_length(batch_rows);

  RETURN QUERY
  WITH normalized_rows AS (
    SELECT
      NULLIF(BTRIM(entry->>'source_id'), '') AS source_id,
      (entry->>'ActualNetWeightValue')::NUMERIC(14,3) AS "ActualNetWeightValue",
      BTRIM(entry->>'DeviceName') AS "DeviceName",
      (entry->>'Date')::DATE AS "Date",
      (entry->>'Timestamp')::TIMESTAMPTZ AS "Timestamp",
      BTRIM(entry->>'ArticleNumber') AS "ArticleNumber"
    FROM jsonb_array_elements(batch_rows) AS entry
  ),
  inserted_rows AS (
    INSERT INTO public."Bizerba_ID" (
      source_id,
      "ActualNetWeightValue",
      "DeviceName",
      "Date",
      "Timestamp",
      "ArticleNumber"
    )
    SELECT
      source_id,
      "ActualNetWeightValue",
      "DeviceName",
      "Date",
      "Timestamp",
      "ArticleNumber"
    FROM normalized_rows
    ON CONFLICT (source_id) DO NOTHING
    RETURNING 1
  )
  SELECT
    COUNT(*)::INTEGER AS inserted_count,
    (total_count - COUNT(*))::INTEGER AS skipped_count
  FROM inserted_rows;
END;
$$;
