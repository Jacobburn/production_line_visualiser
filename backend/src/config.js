import dotenv from 'dotenv';

dotenv.config();

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

function inferDatabaseProvider(connectionString) {
  try {
    const url = new URL(connectionString);
    if (/supabase/i.test(url.hostname)) return 'supabase';
  } catch {
    // Ignore invalid URL values; fallback handled by caller.
  }
  return 'postgres';
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

const defaultLocalDatabaseUrl = 'postgresql://postgres:postgres@localhost:5432/production_line';
const databaseUrl = firstNonEmptyEnv(['SUPABASE_DB_URL', 'DATABASE_URL'], defaultLocalDatabaseUrl);
const rawProvider = String(process.env.DATABASE_PROVIDER || '').trim().toLowerCase();
const inferredProvider = inferDatabaseProvider(databaseUrl);
const databaseProvider = rawProvider === 'supabase' || rawProvider === 'postgres'
  ? rawProvider
  : inferredProvider;

export const config = {
  nodeEnv: process.env.NODE_ENV || 'development',
  port: Number(process.env.API_PORT || process.env.PORT || 4000),
  databaseProvider,
  databaseUrl,
  jwtSecret: validatedJwtSecret(),
  jwtExpiresIn: process.env.JWT_EXPIRES_IN || '12h',
  frontendOrigin: process.env.FRONTEND_ORIGIN || '*',
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
