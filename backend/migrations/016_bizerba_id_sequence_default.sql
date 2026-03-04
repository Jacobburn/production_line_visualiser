DO $$
DECLARE
  max_id BIGINT;
BEGIN
  IF EXISTS (
    SELECT 1
    FROM information_schema.columns
    WHERE table_schema = 'public'
      AND table_name = 'Bizerba_ID'
      AND column_name = 'id'
      AND udt_name = 'int8'
  ) THEN
    CREATE SEQUENCE IF NOT EXISTS public."Bizerba_ID_id_seq";

    SELECT COALESCE(MAX(id), 0)
      INTO max_id
      FROM public."Bizerba_ID";

    IF max_id > 0 THEN
      PERFORM setval('public."Bizerba_ID_id_seq"', max_id, true);
    ELSE
      PERFORM setval('public."Bizerba_ID_id_seq"', 1, false);
    END IF;

    ALTER TABLE public."Bizerba_ID"
      ALTER COLUMN id SET DEFAULT nextval('public."Bizerba_ID_id_seq"');

    ALTER SEQUENCE public."Bizerba_ID_id_seq"
      OWNED BY public."Bizerba_ID".id;
  END IF;
END
$$;
