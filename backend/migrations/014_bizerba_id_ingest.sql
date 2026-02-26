CREATE TABLE IF NOT EXISTS public."Bizerba_ID" (
  id BIGSERIAL PRIMARY KEY,
  "ArticleName" TEXT NOT NULL,
  "ActualNetWeightValue" NUMERIC(14,3) NOT NULL,
  "DeviceName" TEXT NOT NULL,
  "Date" DATE NOT NULL,
  "Timestamp" TIMESTAMPTZ NOT NULL,
  "ArticleNumber" TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_bizerba_id_timestamp
  ON public."Bizerba_ID" ("Timestamp" DESC);

CREATE INDEX IF NOT EXISTS idx_bizerba_id_device_name
  ON public."Bizerba_ID" ("DeviceName");

ALTER TABLE public."Bizerba_ID" ENABLE ROW LEVEL SECURITY;
