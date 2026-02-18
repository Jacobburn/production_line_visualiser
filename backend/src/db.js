import { Pool } from 'pg';
import { config } from './config.js';

function shouldUseSsl(connectionString) {
  const raw = String(process.env.DB_SSL || '').trim().toLowerCase();
  if (raw === 'true') return true;
  if (raw === 'false') return false;
  try {
    const url = new URL(connectionString);
    const sslMode = (url.searchParams.get('sslmode') || '').toLowerCase();
    if (['require', 'verify-ca', 'verify-full', 'prefer'].includes(sslMode)) return true;
    if (['disable', 'allow'].includes(sslMode)) return false;
    return /\.render\.com$/i.test(url.hostname)
      || /\.onrender\.com$/i.test(url.hostname)
      || /supabase/i.test(url.hostname);
  } catch {
    return false;
  }
}

function shouldRejectUnauthorized() {
  const raw = String(process.env.DB_SSL_REJECT_UNAUTHORIZED || '').trim().toLowerCase();
  if (raw === 'true') return true;
  if (raw === 'false') return false;
  return false;
}

export const pool = new Pool({
  connectionString: config.databaseUrl,
  ssl: shouldUseSsl(config.databaseUrl) ? { rejectUnauthorized: shouldRejectUnauthorized() } : undefined,
  max: config.dbPoolMax,
  idleTimeoutMillis: config.dbIdleTimeoutMs
});

export async function dbQuery(text, params = []) {
  return pool.query(text, params);
}
