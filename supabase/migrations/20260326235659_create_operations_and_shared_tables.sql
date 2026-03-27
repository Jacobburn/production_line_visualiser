CREATE SCHEMA IF NOT EXISTS shared;
CREATE SCHEMA IF NOT EXISTS operations;

CREATE TABLE IF NOT EXISTS shared.production_areas (
  id UUID NOT NULL,
  name TEXT NOT NULL,
  display_order INTEGER NOT NULL,
  created_by_user_id UUID NOT NULL,
  created_at TIMESTAMPTZ NOT NULL,
  updated_at TIMESTAMPTZ NOT NULL,
  area_crewing_total NUMERIC,
  CONSTRAINT production_areas_pkey PRIMARY KEY (id)
);

CREATE TABLE IF NOT EXISTS shared.production_lines (
  id UUID NOT NULL,
  name TEXT NOT NULL,
  secret_key TEXT NOT NULL,
  is_active BOOLEAN NOT NULL,
  created_by_user_id UUID NOT NULL,
  created_at TIMESTAMPTZ NOT NULL,
  updated_at TIMESTAMPTZ NOT NULL,
  group_id UUID NOT NULL REFERENCES shared.production_areas(id),
  display_order INTEGER NOT NULL,
  device_name TEXT,
  crewing_pattern JSONB,
  CONSTRAINT production_lines_pkey PRIMARY KEY (id)
);

CREATE TABLE IF NOT EXISTS shared.products (
  id UUID NOT NULL,
  code TEXT NOT NULL,
  desc_short TEXT NOT NULL,
  desc_long TEXT NOT NULL,
  tray TEXT NOT NULL,
  production_line_type TEXT NOT NULL,
  "KG_per_tray" NUMERIC NOT NULL,
  trays_per_outer NUMERIC NOT NULL,
  pack_type TEXT NOT NULL,
  "KG_per_outer" NUMERIC NOT NULL,
  outer_per_pallet NUMERIC NOT NULL,
  optimised_standard_tdm NUMERIC,
  optimised_crewing NUMERIC,
  CONSTRAINT products_pkey PRIMARY KEY (id)
);

CREATE TABLE IF NOT EXISTS operations.shift_logs (
  id UUID NOT NULL,
  line_id UUID NOT NULL REFERENCES shared.production_lines(id),
  date DATE NOT NULL,
  shift TEXT NOT NULL,
  crew_on_shift INTEGER NOT NULL,
  start_time TIME WITHOUT TIME ZONE NOT NULL,
  finish_time TIME WITHOUT TIME ZONE,
  submitted_by_user_id UUID NOT NULL,
  created_at TIMESTAMPTZ NOT NULL,
  notes TEXT,
  CONSTRAINT shift_logs_pkey PRIMARY KEY (id)
);

CREATE TABLE IF NOT EXISTS operations.production_runs (
  id UUID NOT NULL,
  line_id UUID NOT NULL REFERENCES shared.production_lines(id),
  date DATE NOT NULL,
  product TEXT NOT NULL,
  start_time TIME WITHOUT TIME ZONE NOT NULL,
  finish_time TIME WITHOUT TIME ZONE NOT NULL,
  units_produced NUMERIC NOT NULL,
  submitted_by_user_id UUID NOT NULL,
  created_at TIMESTAMPTZ NOT NULL,
  run_crewing_pattern JSONB NOT NULL,
  mean_weight NUMERIC NOT NULL,
  median_weight NUMERIC NOT NULL,
  stdev NUMERIC NOT NULL,
  notes TEXT,
  CONSTRAINT production_runs_pkey PRIMARY KEY (id)
);

CREATE TABLE IF NOT EXISTS operations.downtime_logs (
  id UUID NOT NULL,
  line_id UUID NOT NULL REFERENCES shared.production_lines(id),
  date DATE NOT NULL,
  downtime_start TIME WITHOUT TIME ZONE NOT NULL,
  downtime_finish TIME WITHOUT TIME ZONE NOT NULL,
  equipment_stage_id UUID NOT NULL,
  reason TEXT NOT NULL,
  submitted_by_user_id UUID NOT NULL,
  created_at TIMESTAMPTZ NOT NULL,
  notes TEXT,
  exclude_from_calculation BOOLEAN NOT NULL,
  CONSTRAINT downtime_logs_pkey PRIMARY KEY (id)
);

CREATE TABLE IF NOT EXISTS operations._downtime_types (
  id UUID NOT NULL,
  category TEXT NOT NULL,
  sub_category TEXT NOT NULL,
  is_break BOOLEAN NOT NULL,
  CONSTRAINT downtime_types_pkey PRIMARY KEY (id)
);
