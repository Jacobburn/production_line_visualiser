CREATE TABLE IF NOT EXISTS product_catalog_entries (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  catalog_values JSONB NOT NULL DEFAULT '[]'::jsonb,
  line_ids UUID[] NOT NULL DEFAULT ARRAY[]::UUID[],
  all_lines BOOLEAN NOT NULL DEFAULT TRUE,
  created_by_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_product_catalog_entries_created
  ON product_catalog_entries (created_at ASC, id ASC);

CREATE INDEX IF NOT EXISTS idx_product_catalog_entries_all_lines
  ON product_catalog_entries (all_lines, created_at ASC);

CREATE INDEX IF NOT EXISTS idx_product_catalog_entries_line_ids
  ON product_catalog_entries
  USING GIN (line_ids);
