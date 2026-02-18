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
  return true;
}

export const pool = new Pool({
  connectionString: config.databaseUrl,
  ssl: shouldUseSsl(config.databaseUrl) ? { rejectUnauthorized: shouldRejectUnauthorized() } : undefined,
  max: config.dbPoolMax,
  idleTimeoutMillis: config.dbIdleTimeoutMs,
  connectionTimeoutMillis: config.dbConnectionTimeoutMs,
  query_timeout: config.dbQueryTimeoutMs,
  statement_timeout: config.dbStatementTimeoutMs,
  application_name: 'kebab-line-backend'
});

pool.on('error', (error) => {
  // Prevent idle client errors from crashing the Node process.
  console.error('[db] Unexpected idle client error:', error);
});

function isLikelyReadOnlyQuery(text) {
  const sql = String(text || '').trim();
  if (!sql) return false;
  const startsReadOnly = /^(SELECT|SHOW|EXPLAIN|WITH)\b/i.test(sql);
  if (!startsReadOnly) return false;
  return !/\b(INSERT|UPDATE|DELETE|MERGE|UPSERT|CREATE|ALTER|DROP|TRUNCATE|GRANT|REVOKE)\b/i.test(sql);
}

function isTransientDbError(error) {
  const code = String(error?.code || '').toUpperCase();
  const syscall = String(error?.syscall || '').toLowerCase();
  const message = String(error?.message || '').toLowerCase();
  if (code.startsWith('08')) return true; // Connection exception class.
  if (['57P01', '57P02', '57P03', '53300', '53400'].includes(code)) return true;
  if (['ETIMEDOUT', 'ECONNRESET', 'EPIPE', 'ECONNREFUSED', 'ENOTFOUND', 'EAI_AGAIN'].includes(code)) return true;
  if (syscall === 'getaddrinfo') return true;
  return message.includes('terminating connection')
    || message.includes('connection terminated unexpectedly')
    || message.includes('could not connect')
    || message.includes('timeout');
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export async function dbQuery(text, params = []) {
  const readOnly = isLikelyReadOnlyQuery(text);
  const maxRetries = readOnly ? config.dbMaxRetries : 0;
  let attempt = 0;
  while (attempt <= maxRetries) {
    try {
      return await pool.query(text, params);
    } catch (error) {
      if (attempt >= maxRetries || !isTransientDbError(error)) throw error;
      const delay = config.dbRetryDelayMs * (attempt + 1);
      console.warn(`[db] transient query failure, retrying (${attempt + 1}/${maxRetries})`, {
        code: error?.code,
        message: error?.message
      });
      if (delay > 0) await sleep(delay);
      attempt += 1;
    }
  }
  return pool.query(text, params);
}
