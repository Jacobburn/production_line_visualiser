import dotenv from 'dotenv';

dotenv.config();

const nodeEnv = String(process.env.NODE_ENV || 'development').trim() || 'development';
const isProduction = nodeEnv === 'production';

function requireEnv(name, fallback = '') {
  const value = process.env[name] ?? fallback;
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

function firstNonEmptyEnv(names, fallback = '') {
  for (const name of names) {
    const value = String(process.env[name] || '').trim();
    if (value) return value;
  }
  if (fallback) return fallback;
  throw new Error(`Missing required environment variable. Tried: ${names.join(', ')}`);
}

function parseNumberEnv(name, fallback, { min = Number.NEGATIVE_INFINITY } = {}) {
  const raw = process.env[name];
  if (raw === undefined || raw === null || raw === '') return fallback;
  const parsed = Number(raw);
  if (!Number.isFinite(parsed) || parsed < min) return fallback;
  return parsed;
}

function normalizeOrigin(value, { strict = false } = {}) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  if (raw === '*') return '*';
  try {
    return new URL(raw).origin;
  } catch {
    if (strict) {
      throw new Error(`Invalid frontend origin: ${raw}`);
    }
    return raw.replace(/\/+$/, '');
  }
}

function parseFrontendOrigins({ requireExplicit = false, strict = false, allowWildcard = true } = {}) {
  const raw = firstNonEmptyEnv(['FRONTEND_ORIGINS', 'FRONTEND_ORIGIN'], requireExplicit ? '' : '*');
  const trimmed = String(raw || '').trim();
  if (!trimmed || trimmed === '*') {
    if (!allowWildcard) {
      throw new Error('FRONTEND_ORIGINS must be set to one or more explicit origins in production.');
    }
    return ['*'];
  }
  const origins = trimmed
    .split(',')
    .map((entry) => normalizeOrigin(entry, { strict }))
    .filter(Boolean);
  const deduped = origins.length ? Array.from(new Set(origins)) : [];
  if (!deduped.length) {
    if (!allowWildcard) {
      throw new Error('FRONTEND_ORIGINS must contain at least one valid origin in production.');
    }
    return ['*'];
  }
  if (!allowWildcard && deduped.includes('*')) {
    throw new Error('Wildcard FRONTEND_ORIGINS is not allowed in production.');
  }
  return deduped;
}

function validatedJwtSecret() {
  const secret = requireEnv('JWT_SECRET').trim();
  if (secret === 'change-me-in-staging-and-prod' || secret === 'replace-with-a-32-plus-char-random-secret') {
    throw new Error('JWT_SECRET must be changed from the default placeholder value.');
  }
  if (secret.length < 32) {
    throw new Error('JWT_SECRET must be at least 32 characters long.');
  }
  return secret;
}

const databaseUrl = requireEnv('SUPABASE_DB_URL').trim();
let databaseHost = '';
try {
  databaseHost = String(new URL(databaseUrl).hostname || '').toLowerCase();
} catch {
  throw new Error('SUPABASE_DB_URL must be a valid PostgreSQL connection string.');
}
if (!/supabase/i.test(databaseHost)) {
  throw new Error('SUPABASE_DB_URL must target a Supabase-hosted PostgreSQL instance.');
}
if (databaseHost === 'localhost' || databaseHost === '127.0.0.1' || databaseHost === '::1') {
  throw new Error('SUPABASE_DB_URL cannot point to localhost.');
}
const databaseProvider = 'supabase';
const frontendOrigins = parseFrontendOrigins({
  requireExplicit: isProduction,
  strict: isProduction,
  allowWildcard: !isProduction
});

export const config = {
  nodeEnv,
  port: Number(process.env.API_PORT || process.env.PORT || 4000),
  databaseProvider,
  databaseUrl,
  jwtSecret: validatedJwtSecret(),
  jwtExpiresIn: process.env.JWT_EXPIRES_IN || '12h',
  frontendOrigins,
  frontendOrigin: frontendOrigins[0] || '*',
  dbPoolMax: parseNumberEnv('DB_POOL_MAX', 10, { min: 1 }),
  dbIdleTimeoutMs: parseNumberEnv('DB_IDLE_TIMEOUT_MS', 30000, { min: 0 }),
  dbConnectionTimeoutMs: parseNumberEnv('DB_CONNECTION_TIMEOUT_MS', 10000, { min: 1000 }),
  dbQueryTimeoutMs: parseNumberEnv('DB_QUERY_TIMEOUT_MS', 15000, { min: 1000 }),
  dbStatementTimeoutMs: parseNumberEnv('DB_STATEMENT_TIMEOUT_MS', 20000, { min: 1000 }),
  dbMaxRetries: parseNumberEnv('DB_MAX_RETRIES', 1, { min: 0 }),
  dbRetryDelayMs: parseNumberEnv('DB_RETRY_DELAY_MS', 250, { min: 0 }),
  httpRequestTimeoutMs: parseNumberEnv('HTTP_REQUEST_TIMEOUT_MS', 30000, { min: 1000 }),
  httpHeadersTimeoutMs: parseNumberEnv('HTTP_HEADERS_TIMEOUT_MS', 35000, { min: 1000 }),
  httpKeepAliveTimeoutMs: parseNumberEnv('HTTP_KEEP_ALIVE_TIMEOUT_MS', 5000, { min: 1000 })
};
