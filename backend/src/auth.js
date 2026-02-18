import bcrypt from 'bcryptjs';
import jwt from 'jsonwebtoken';
import { z } from 'zod';
import { config } from './config.js';
import { dbQuery } from './db.js';

const authHeaderSchema = z.string().regex(/^Bearer\s+.+$/i);

export async function hashPassword(password) {
  return bcrypt.hash(password, 12);
}

export async function comparePassword(password, passwordHash) {
  return bcrypt.compare(password, passwordHash);
}

export function signAccessToken(user) {
  return jwt.sign(
    {
      sub: user.id,
      username: user.username,
      role: user.role,
      name: user.name
    },
    config.jwtSecret,
    { expiresIn: config.jwtExpiresIn }
  );
}

export async function getUserByUsername(username) {
  const result = await dbQuery(
    `SELECT id, name, username, password_hash, role, is_active
     FROM users
     WHERE username = $1`,
    [String(username || '').trim().toLowerCase()]
  );
  return result.rows[0] || null;
}

export async function authMiddleware(req, res, next) {
  try {
    const parsed = authHeaderSchema.safeParse(req.headers.authorization || '');
    if (!parsed.success) return res.status(401).json({ error: 'Missing or invalid authorization header' });
    const token = parsed.data.replace(/^Bearer\s+/i, '').trim();
    let payload;
    try {
      payload = jwt.verify(token, config.jwtSecret);
    } catch {
      return res.status(401).json({ error: 'Invalid or expired token' });
    }
    let userResult;
    try {
      userResult = await dbQuery(
        `SELECT id, name, username, role, is_active FROM users WHERE id = $1`,
        [payload.sub]
      );
    } catch (error) {
      console.error('[auth] User lookup failed:', error);
      return res.status(503).json({ error: 'Database unavailable' });
    }
    const user = userResult.rows[0];
    if (!user || !user.is_active) return res.status(401).json({ error: 'User inactive or not found' });
    req.user = user;
    return next();
  } catch (error) {
    console.error('[auth] Unexpected auth middleware error:', error);
    return res.status(500).json({ error: 'Authentication service unavailable' });
  }
}

export function requireRole(...roles) {
  return (req, res, next) => {
    if (!req.user) return res.status(401).json({ error: 'Not authenticated' });
    if (!roles.includes(req.user.role)) return res.status(403).json({ error: 'Forbidden' });
    return next();
  };
}
