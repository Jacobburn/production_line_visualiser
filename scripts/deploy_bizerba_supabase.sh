#!/usr/bin/env bash
set -euo pipefail

: "${SUPABASE_ACCESS_TOKEN:?Set SUPABASE_ACCESS_TOKEN first.}"
: "${SUPABASE_PROJECT_REF:?Set SUPABASE_PROJECT_REF first.}"
: "${SUPABASE_DB_PASSWORD:?Set SUPABASE_DB_PASSWORD first.}"
: "${BIZERBA_DEVICE_KEY:?Set BIZERBA_DEVICE_KEY first.}"

echo "Linking project ${SUPABASE_PROJECT_REF}..."
npx --yes supabase link --project-ref "${SUPABASE_PROJECT_REF}"

echo "Pushing migrations..."
npx --yes supabase db push --include-all --password "${SUPABASE_DB_PASSWORD}"

echo "Setting function secrets..."
npx --yes supabase secrets set \
  --project-ref "${SUPABASE_PROJECT_REF}" \
  BIZERBA_DEVICE_KEY="${BIZERBA_DEVICE_KEY}"

echo "Deploying edge function ingest-bizerba..."
npx --yes supabase functions deploy ingest-bizerba --project-ref "${SUPABASE_PROJECT_REF}"

echo "Done."
