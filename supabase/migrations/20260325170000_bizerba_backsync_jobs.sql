CREATE TABLE IF NOT EXISTS public.bizerba_backsync_jobs (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  status TEXT NOT NULL DEFAULT 'queued',
  trigger TEXT NOT NULL DEFAULT 'manual',
  full_rescan BOOLEAN NOT NULL DEFAULT FALSE,
  requested_by_user_id TEXT NOT NULL DEFAULT '',
  requested_by_name TEXT NOT NULL DEFAULT '',
  summary JSONB NOT NULL DEFAULT '{}'::jsonb,
  state JSONB NOT NULL DEFAULT '{}'::jsonb,
  error TEXT NOT NULL DEFAULT '',
  worker_id TEXT NOT NULL DEFAULT '',
  lease_expires_at TIMESTAMPTZ,
  started_at TIMESTAMPTZ,
  finished_at TIMESTAMPTZ,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  CONSTRAINT bizerba_backsync_jobs_status_check CHECK (status IN ('queued', 'running', 'completed', 'failed'))
);

CREATE INDEX IF NOT EXISTS idx_bizerba_backsync_jobs_status_created
  ON public.bizerba_backsync_jobs(status, created_at ASC);

CREATE INDEX IF NOT EXISTS idx_bizerba_backsync_jobs_lease
  ON public.bizerba_backsync_jobs(lease_expires_at);

CREATE INDEX IF NOT EXISTS idx_bizerba_backsync_jobs_updated
  ON public.bizerba_backsync_jobs(updated_at DESC);

CREATE TABLE IF NOT EXISTS public.bizerba_auto_line_states (
  line_id UUID PRIMARY KEY,
  device_name TEXT NOT NULL DEFAULT '',
  last_processed_bizerba_id BIGINT NOT NULL DEFAULT 0,
  last_processed_timestamp TIMESTAMPTZ,
  last_processed_bizerba_id_text TEXT NOT NULL DEFAULT '',
  pending_article_number TEXT NOT NULL DEFAULT '',
  pending_count INTEGER NOT NULL DEFAULT 0,
  active_run_log_id UUID,
  active_article_number TEXT NOT NULL DEFAULT '',
  active_product TEXT NOT NULL DEFAULT '',
  active_run_date DATE,
  active_start_timestamp TIMESTAMPTZ,
  active_last_date DATE,
  active_last_timestamp TIMESTAMPTZ,
  active_units_produced NUMERIC(14,2) NOT NULL DEFAULT 0,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  CONSTRAINT bizerba_auto_line_states_pending_count_check CHECK (pending_count >= 0),
  CONSTRAINT bizerba_auto_line_states_active_units_check CHECK (active_units_produced >= 0)
);

CREATE INDEX IF NOT EXISTS idx_bizerba_auto_line_states_updated_at
  ON public.bizerba_auto_line_states (updated_at DESC);

CREATE INDEX IF NOT EXISTS idx_bizerba_auto_line_states_active_run
  ON public.bizerba_auto_line_states (active_run_log_id)
  WHERE active_run_log_id IS NOT NULL;

CREATE OR REPLACE FUNCTION public.claim_bizerba_backsync_job(
  p_worker_id TEXT,
  p_lease_seconds INTEGER DEFAULT 120
)
RETURNS SETOF public.bizerba_backsync_jobs
LANGUAGE plpgsql
AS $$
DECLARE
  v_job public.bizerba_backsync_jobs%ROWTYPE;
BEGIN
  UPDATE public.bizerba_backsync_jobs AS j
  SET
    status = 'running',
    started_at = COALESCE(j.started_at, NOW()),
    worker_id = COALESCE(NULLIF(BTRIM(p_worker_id), ''), 'worker'),
    lease_expires_at = NOW() + make_interval(secs => GREATEST(10, COALESCE(p_lease_seconds, 120))),
    updated_at = NOW()
  WHERE j.id = (
    SELECT q.id
    FROM public.bizerba_backsync_jobs AS q
    WHERE q.status IN ('queued', 'running')
      AND (
        q.status = 'queued'
        OR q.lease_expires_at IS NULL
        OR q.lease_expires_at <= NOW()
      )
    ORDER BY
      CASE WHEN q.status = 'running' THEN 0 ELSE 1 END,
      q.created_at ASC
    LIMIT 1
    FOR UPDATE SKIP LOCKED
  )
  RETURNING j.* INTO v_job;

  IF v_job.id IS NULL THEN
    RETURN;
  END IF;

  RETURN QUERY
  SELECT *
  FROM public.bizerba_backsync_jobs
  WHERE id = v_job.id;
END;
$$;
