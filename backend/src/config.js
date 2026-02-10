import dotenv from 'dotenv';

dotenv.config();

function requireEnv(name, fallback = '') {
  const value = process.env[name] ?? fallback;
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

export const config = {
  nodeEnv: process.env.NODE_ENV || 'development',
  port: Number(process.env.API_PORT || 4000),
  databaseUrl: requireEnv('DATABASE_URL', 'postgresql://postgres:postgres@localhost:5432/production_line'),
  jwtSecret: requireEnv('JWT_SECRET', 'change-me-in-staging-and-prod'),
  jwtExpiresIn: process.env.JWT_EXPIRES_IN || '12h',
  frontendOrigin: process.env.FRONTEND_ORIGIN || '*'
};
