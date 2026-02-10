import { Pool } from 'pg';
import { config } from './config.js';

export const pool = new Pool({
  connectionString: config.databaseUrl,
  max: 10,
  idleTimeoutMillis: 30000
});

export async function dbQuery(text, params = []) {
  return pool.query(text, params);
}
