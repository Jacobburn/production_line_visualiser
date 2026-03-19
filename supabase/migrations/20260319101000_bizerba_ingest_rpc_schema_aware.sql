CREATE OR REPLACE FUNCTION public.ingest_bizerba_rows(batch_rows JSONB)
RETURNS TABLE(inserted_count INTEGER, skipped_count INTEGER)
LANGUAGE plpgsql
SET search_path = public
AS $$
DECLARE
  total_count INTEGER := 0;
  id_udt TEXT;
  id_default TEXT;
  stmt TEXT;
BEGIN
  IF batch_rows IS NULL OR jsonb_typeof(batch_rows) <> 'array' THEN
    RAISE EXCEPTION 'batch_rows must be a JSON array';
  END IF;

  total_count := jsonb_array_length(batch_rows);

  SELECT
    columns.udt_name,
    columns.column_default
  INTO
    id_udt,
    id_default
  FROM information_schema.columns AS columns
  WHERE columns.table_schema = 'public'
    AND columns.table_name = 'Bizerba_ID'
    AND columns.column_name = 'id';

  IF id_udt IS NULL THEN
    RAISE EXCEPTION 'public."Bizerba_ID".id column was not found';
  END IF;

  IF id_udt = 'int8' AND id_default IS NOT NULL THEN
    stmt := $sql$
      WITH normalized_rows AS (
        SELECT
          NULLIF(BTRIM(entry->>'source_id'), '') AS source_id,
          (entry->>'ActualNetWeightValue')::NUMERIC(14,3) AS "ActualNetWeightValue",
          BTRIM(entry->>'DeviceName') AS "DeviceName",
          (entry->>'Date')::DATE AS "Date",
          (entry->>'Timestamp')::TIMESTAMPTZ AS "Timestamp",
          BTRIM(entry->>'ArticleNumber') AS "ArticleNumber"
        FROM jsonb_array_elements($1) AS entry
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
        ($2 - COUNT(*))::INTEGER AS skipped_count
      FROM inserted_rows
    $sql$;
  ELSIF id_udt = 'uuid' THEN
    stmt := $sql$
      WITH normalized_rows AS (
        SELECT
          NULLIF(BTRIM(entry->>'source_id'), '') AS source_id,
          (entry->>'ActualNetWeightValue')::NUMERIC(14,3) AS "ActualNetWeightValue",
          BTRIM(entry->>'DeviceName') AS "DeviceName",
          (entry->>'Date')::DATE AS "Date",
          (entry->>'Timestamp')::TIMESTAMPTZ AS "Timestamp",
          BTRIM(entry->>'ArticleNumber') AS "ArticleNumber"
        FROM jsonb_array_elements($1) AS entry
      ),
      inserted_rows AS (
        INSERT INTO public."Bizerba_ID" (
          id,
          source_id,
          "ActualNetWeightValue",
          "DeviceName",
          "Date",
          "Timestamp",
          "ArticleNumber"
        )
        SELECT
          source_id::UUID,
          source_id,
          "ActualNetWeightValue",
          "DeviceName",
          "Date",
          "Timestamp",
          "ArticleNumber"
        FROM normalized_rows
        ON CONFLICT (id) DO NOTHING
        RETURNING 1
      )
      SELECT
        COUNT(*)::INTEGER AS inserted_count,
        ($2 - COUNT(*))::INTEGER AS skipped_count
      FROM inserted_rows
    $sql$;
  ELSIF id_udt IN ('text', 'varchar', 'bpchar') THEN
    stmt := $sql$
      WITH normalized_rows AS (
        SELECT
          NULLIF(BTRIM(entry->>'source_id'), '') AS source_id,
          (entry->>'ActualNetWeightValue')::NUMERIC(14,3) AS "ActualNetWeightValue",
          BTRIM(entry->>'DeviceName') AS "DeviceName",
          (entry->>'Date')::DATE AS "Date",
          (entry->>'Timestamp')::TIMESTAMPTZ AS "Timestamp",
          BTRIM(entry->>'ArticleNumber') AS "ArticleNumber"
        FROM jsonb_array_elements($1) AS entry
      ),
      inserted_rows AS (
        INSERT INTO public."Bizerba_ID" (
          id,
          source_id,
          "ActualNetWeightValue",
          "DeviceName",
          "Date",
          "Timestamp",
          "ArticleNumber"
        )
        SELECT
          source_id,
          source_id,
          "ActualNetWeightValue",
          "DeviceName",
          "Date",
          "Timestamp",
          "ArticleNumber"
        FROM normalized_rows
        ON CONFLICT (id) DO NOTHING
        RETURNING 1
      )
      SELECT
        COUNT(*)::INTEGER AS inserted_count,
        ($2 - COUNT(*))::INTEGER AS skipped_count
      FROM inserted_rows
    $sql$;
  ELSE
    RAISE EXCEPTION 'Unsupported public."Bizerba_ID".id type: %', id_udt;
  END IF;

  RETURN QUERY EXECUTE stmt USING batch_rows, total_count;
END;
$$;
