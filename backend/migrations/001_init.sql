CREATE EXTENSION IF NOT EXISTS pgcrypto;

CREATE TABLE IF NOT EXISTS schema_migrations (
  id BIGSERIAL PRIMARY KEY,
  filename TEXT UNIQUE NOT NULL,
  applied_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS users (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  name TEXT NOT NULL,
  username TEXT UNIQUE NOT NULL,
  password_hash TEXT NOT NULL,
  role TEXT NOT NULL CHECK (role IN ('manager', 'supervisor')),
  is_active BOOLEAN NOT NULL DEFAULT TRUE,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS production_lines (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  name TEXT NOT NULL UNIQUE,
  secret_key TEXT NOT NULL,
  is_active BOOLEAN NOT NULL DEFAULT TRUE,
  created_by_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS line_stages (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  line_id UUID NOT NULL REFERENCES production_lines(id) ON DELETE CASCADE,
  stage_order INTEGER NOT NULL,
  stage_name TEXT NOT NULL,
  stage_type TEXT NOT NULL CHECK (stage_type IN ('main', 'prep', 'transfer')),
  day_crew INTEGER NOT NULL DEFAULT 0 CHECK (day_crew >= 0),
  night_crew INTEGER NOT NULL DEFAULT 0 CHECK (night_crew >= 0),
  max_throughput_per_crew NUMERIC(10,2) NOT NULL DEFAULT 0 CHECK (max_throughput_per_crew >= 0),
  x NUMERIC(8,3) NOT NULL DEFAULT 0,
  y NUMERIC(8,3) NOT NULL DEFAULT 0,
  w NUMERIC(8,3) NOT NULL DEFAULT 9.5,
  h NUMERIC(8,3) NOT NULL DEFAULT 15,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  UNIQUE(line_id, stage_order)
);

CREATE TABLE IF NOT EXISTS line_layout_guides (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  line_id UUID NOT NULL REFERENCES production_lines(id) ON DELETE CASCADE,
  guide_type TEXT NOT NULL CHECK (guide_type IN ('line', 'arrow', 'shape')),
  x NUMERIC(8,3) NOT NULL DEFAULT 0,
  y NUMERIC(8,3) NOT NULL DEFAULT 0,
  w NUMERIC(8,3) NOT NULL DEFAULT 12,
  h NUMERIC(8,3) NOT NULL DEFAULT 2,
  angle NUMERIC(8,3) NOT NULL DEFAULT 0,
  src TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS supervisor_line_assignments (
  supervisor_user_id UUID NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  line_id UUID NOT NULL REFERENCES production_lines(id) ON DELETE CASCADE,
  assigned_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  assigned_by_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
  PRIMARY KEY (supervisor_user_id, line_id)
);

CREATE TABLE IF NOT EXISTS shift_logs (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  line_id UUID NOT NULL REFERENCES production_lines(id) ON DELETE CASCADE,
  date DATE NOT NULL,
  shift TEXT NOT NULL CHECK (shift IN ('Day', 'Night', 'Full Day')),
  crew_on_shift INTEGER NOT NULL DEFAULT 0 CHECK (crew_on_shift >= 0),
  start_time TIME NOT NULL,
  break1_start TIME,
  break2_start TIME,
  break3_start TIME,
  finish_time TIME NOT NULL,
  submitted_by_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
  submitted_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS run_logs (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  line_id UUID NOT NULL REFERENCES production_lines(id) ON DELETE CASCADE,
  date DATE NOT NULL,
  shift TEXT NOT NULL CHECK (shift IN ('Day', 'Night', 'Full Day')),
  product TEXT NOT NULL,
  setup_start_time TIME,
  production_start_time TIME NOT NULL,
  finish_time TIME NOT NULL,
  units_produced NUMERIC(14,2) NOT NULL DEFAULT 0 CHECK (units_produced >= 0),
  submitted_by_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
  submitted_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS downtime_logs (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  line_id UUID NOT NULL REFERENCES production_lines(id) ON DELETE CASCADE,
  date DATE NOT NULL,
  shift TEXT NOT NULL CHECK (shift IN ('Day', 'Night', 'Full Day')),
  downtime_start TIME NOT NULL,
  downtime_finish TIME NOT NULL,
  equipment_stage_id UUID REFERENCES line_stages(id) ON DELETE SET NULL,
  reason TEXT,
  submitted_by_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
  submitted_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS audit_events (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  line_id UUID REFERENCES production_lines(id) ON DELETE CASCADE,
  actor_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
  actor_name TEXT,
  actor_role TEXT,
  action TEXT NOT NULL,
  details TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_line_stages_line ON line_stages(line_id);
CREATE INDEX IF NOT EXISTS idx_guides_line ON line_layout_guides(line_id);
CREATE INDEX IF NOT EXISTS idx_assignments_supervisor ON supervisor_line_assignments(supervisor_user_id);
CREATE INDEX IF NOT EXISTS idx_assignments_line ON supervisor_line_assignments(line_id);
CREATE INDEX IF NOT EXISTS idx_shift_logs_line_date_shift ON shift_logs(line_id, date, shift);
CREATE INDEX IF NOT EXISTS idx_run_logs_line_date_shift ON run_logs(line_id, date, shift);
CREATE INDEX IF NOT EXISTS idx_down_logs_line_date_shift ON downtime_logs(line_id, date, shift);
CREATE INDEX IF NOT EXISTS idx_audit_events_line_created ON audit_events(line_id, created_at DESC);
