import express from 'express';
import cors from 'cors';
import { z } from 'zod';
import { config } from './config.js';
import { dbQuery, pool } from './db.js';
import {
  authMiddleware,
  comparePassword,
  getUserByUsername,
  requireRole,
  signAccessToken
} from './auth.js';

const app = express();

app.use(cors({ origin: config.frontendOrigin === '*' ? true : config.frontendOrigin }));
app.use(express.json({ limit: '1mb' }));

const shiftValues = ['Day', 'Night'];
const timeRegex = /^([01]\d|2[0-3]):([0-5]\d)$/;
const optionalTime = z.string().regex(timeRegex).optional().or(z.literal('')).or(z.null());

const loginSchema = z.object({
  username: z.string().min(1),
  password: z.string().min(1)
});

const lineSchema = z.object({
  name: z.string().min(2).max(120),
  secretKey: z.string().min(4).max(64)
});

const shiftLogSchema = z.object({
  lineId: z.string().uuid(),
  date: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
  shift: z.enum(shiftValues),
  crewOnShift: z.number().int().min(0),
  startTime: z.string().regex(timeRegex),
  finishTime: z.string().regex(timeRegex),
  break1Start: optionalTime,
  break2Start: optionalTime,
  break3Start: optionalTime
});

const runLogSchema = z.object({
  lineId: z.string().uuid(),
  date: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
  shift: z.enum(shiftValues),
  product: z.string().min(1).max(120),
  setUpStartTime: optionalTime,
  productionStartTime: z.string().regex(timeRegex),
  finishTime: z.string().regex(timeRegex),
  unitsProduced: z.number().nonnegative()
});

const downtimeLogSchema = z.object({
  lineId: z.string().uuid(),
  date: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
  shift: z.enum(shiftValues),
  downtimeStart: z.string().regex(timeRegex),
  downtimeFinish: z.string().regex(timeRegex),
  equipmentStageId: z.string().uuid().nullable().optional(),
  reason: z.string().max(250).optional().or(z.literal('')).or(z.null())
});

async function hasLineAccess(user, lineId) {
  if (user.role === 'manager') return true;
  const result = await dbQuery(
    `SELECT 1
     FROM supervisor_line_assignments
     WHERE supervisor_user_id = $1 AND line_id = $2`,
    [user.id, lineId]
  );
  return result.rowCount > 0;
}

async function writeAudit({ lineId = null, actorUserId = null, actorName = null, actorRole = null, action, details = '' }) {
  await dbQuery(
    `INSERT INTO audit_events(line_id, actor_user_id, actor_name, actor_role, action, details)
     VALUES ($1, $2, $3, $4, $5, $6)`,
    [lineId, actorUserId, actorName, actorRole, action, details]
  );
}

app.get('/api/health', async (_req, res) => {
  await dbQuery('SELECT 1');
  res.json({ ok: true, env: config.nodeEnv, timestamp: new Date().toISOString() });
});

app.post('/api/auth/login', async (req, res) => {
  const parsed = loginSchema.safeParse(req.body);
  if (!parsed.success) {
    return res.status(400).json({ error: 'Invalid login payload' });
  }

  const user = await getUserByUsername(parsed.data.username);
  if (!user || !user.is_active) return res.status(401).json({ error: 'Invalid credentials' });

  const valid = await comparePassword(parsed.data.password, user.password_hash);
  if (!valid) return res.status(401).json({ error: 'Invalid credentials' });

  const token = signAccessToken(user);
  return res.json({
    token,
    user: {
      id: user.id,
      name: user.name,
      username: user.username,
      role: user.role
    }
  });
});

app.get('/api/me', authMiddleware, async (req, res) => {
  res.json({ user: req.user });
});

app.get('/api/lines', authMiddleware, async (req, res) => {
  const baseFields = `
    l.id,
    l.name,
    l.secret_key AS "secretKey",
    l.created_at AS "createdAt",
    (
      SELECT COUNT(*)::INT FROM line_stages s WHERE s.line_id = l.id
    ) AS "stageCount",
    (
      SELECT COUNT(*)::INT FROM shift_logs sl WHERE sl.line_id = l.id
    ) AS "shiftCount"
  `;

  const query = req.user.role === 'manager'
    ? `SELECT ${baseFields} FROM production_lines l WHERE l.is_active = TRUE ORDER BY l.created_at DESC`
    : `SELECT ${baseFields}
       FROM production_lines l
       INNER JOIN supervisor_line_assignments a ON a.line_id = l.id
       WHERE a.supervisor_user_id = $1 AND l.is_active = TRUE
       ORDER BY l.created_at DESC`;

  const result = await dbQuery(query, req.user.role === 'manager' ? [] : [req.user.id]);
  res.json({ lines: result.rows });
});

app.post('/api/lines', authMiddleware, requireRole('manager'), async (req, res) => {
  const parsed = lineSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: 'Invalid line payload' });

  try {
    const result = await dbQuery(
      `INSERT INTO production_lines(name, secret_key, created_by_user_id)
       VALUES ($1, $2, $3)
       RETURNING id, name, secret_key AS "secretKey", created_at AS "createdAt"`,
      [parsed.data.name.trim(), parsed.data.secretKey.trim(), req.user.id]
    );

    await writeAudit({
      lineId: result.rows[0].id,
      actorUserId: req.user.id,
      actorName: req.user.name,
      actorRole: req.user.role,
      action: 'CREATE_LINE',
      details: `Line created: ${result.rows[0].name}`
    });

    return res.status(201).json({ line: result.rows[0] });
  } catch (error) {
    if (String(error?.message || '').includes('duplicate key')) {
      return res.status(409).json({ error: 'Line name already exists' });
    }
    throw error;
  }
});

app.get('/api/lines/:lineId', authMiddleware, async (req, res) => {
  const lineId = req.params.lineId;
  if (!z.string().uuid().safeParse(lineId).success) return res.status(400).json({ error: 'Invalid line id' });
  if (!(await hasLineAccess(req.user, lineId))) return res.status(403).json({ error: 'Forbidden' });

  const lineResult = await dbQuery(
    `SELECT id, name, secret_key AS "secretKey", created_at AS "createdAt", updated_at AS "updatedAt"
     FROM production_lines
     WHERE id = $1 AND is_active = TRUE`,
    [lineId]
  );
  if (!lineResult.rowCount) return res.status(404).json({ error: 'Line not found' });

  const [stages, guides] = await Promise.all([
    dbQuery(
      `SELECT id, stage_order AS "stageOrder", stage_name AS "stageName", stage_type AS "stageType",
              day_crew AS "dayCrew", night_crew AS "nightCrew", max_throughput_per_crew AS "maxThroughputPerCrew",
              x, y, w, h
       FROM line_stages
       WHERE line_id = $1
       ORDER BY stage_order ASC`,
      [lineId]
    ),
    dbQuery(
      `SELECT id, guide_type AS "guideType", x, y, w, h, angle, src
       FROM line_layout_guides
       WHERE line_id = $1
       ORDER BY created_at ASC`,
      [lineId]
    )
  ]);

  return res.json({ line: lineResult.rows[0], stages: stages.rows, guides: guides.rows });
});

app.post('/api/logs/shifts', authMiddleware, async (req, res) => {
  const parsed = shiftLogSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: 'Invalid shift log payload' });
  if (!(await hasLineAccess(req.user, parsed.data.lineId))) return res.status(403).json({ error: 'Forbidden' });

  const result = await dbQuery(
    `INSERT INTO shift_logs(line_id, date, shift, crew_on_shift, start_time, break1_start, break2_start, break3_start, finish_time, submitted_by_user_id)
     VALUES ($1, $2, $3, $4, $5, NULLIF($6, '')::time, NULLIF($7, '')::time, NULLIF($8, '')::time, $9, $10)
     RETURNING id, line_id AS "lineId", date, shift, crew_on_shift AS "crewOnShift", start_time AS "startTime", finish_time AS "finishTime", submitted_at AS "submittedAt"`,
    [
      parsed.data.lineId,
      parsed.data.date,
      parsed.data.shift,
      parsed.data.crewOnShift,
      parsed.data.startTime,
      parsed.data.break1Start || '',
      parsed.data.break2Start || '',
      parsed.data.break3Start || '',
      parsed.data.finishTime,
      req.user.id
    ]
  );

  await writeAudit({
    lineId: parsed.data.lineId,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'CREATE_SHIFT_LOG',
    details: `${parsed.data.shift} ${parsed.data.date}, crew ${parsed.data.crewOnShift}`
  });

  return res.status(201).json({ shiftLog: result.rows[0] });
});

app.post('/api/logs/runs', authMiddleware, async (req, res) => {
  const parsed = runLogSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: 'Invalid run log payload' });
  if (!(await hasLineAccess(req.user, parsed.data.lineId))) return res.status(403).json({ error: 'Forbidden' });

  const result = await dbQuery(
    `INSERT INTO run_logs(line_id, date, shift, product, setup_start_time, production_start_time, finish_time, units_produced, submitted_by_user_id)
     VALUES ($1, $2, $3, $4, NULLIF($5, '')::time, $6, $7, $8, $9)
     RETURNING id, line_id AS "lineId", date, shift, product, units_produced AS "unitsProduced", submitted_at AS "submittedAt"`,
    [
      parsed.data.lineId,
      parsed.data.date,
      parsed.data.shift,
      parsed.data.product,
      parsed.data.setUpStartTime || '',
      parsed.data.productionStartTime,
      parsed.data.finishTime,
      parsed.data.unitsProduced,
      req.user.id
    ]
  );

  await writeAudit({
    lineId: parsed.data.lineId,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'CREATE_RUN_LOG',
    details: `${parsed.data.product}, units ${parsed.data.unitsProduced}`
  });

  return res.status(201).json({ runLog: result.rows[0] });
});

app.post('/api/logs/downtime', authMiddleware, async (req, res) => {
  const parsed = downtimeLogSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: 'Invalid downtime log payload' });
  if (!(await hasLineAccess(req.user, parsed.data.lineId))) return res.status(403).json({ error: 'Forbidden' });

  const result = await dbQuery(
    `INSERT INTO downtime_logs(line_id, date, shift, downtime_start, downtime_finish, equipment_stage_id, reason, submitted_by_user_id)
     VALUES ($1, $2, $3, $4, $5, $6, NULLIF($7, ''), $8)
     RETURNING id, line_id AS "lineId", date, shift, reason, submitted_at AS "submittedAt"`,
    [
      parsed.data.lineId,
      parsed.data.date,
      parsed.data.shift,
      parsed.data.downtimeStart,
      parsed.data.downtimeFinish,
      parsed.data.equipmentStageId || null,
      parsed.data.reason || '',
      req.user.id
    ]
  );

  await writeAudit({
    lineId: parsed.data.lineId,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'CREATE_DOWNTIME_LOG',
    details: parsed.data.reason || 'Downtime logged'
  });

  return res.status(201).json({ downtimeLog: result.rows[0] });
});

app.use((error, _req, res, _next) => {
  console.error(error);
  res.status(500).json({ error: 'Internal server error' });
});

const server = app.listen(config.port, () => {
  console.log(`API listening on http://localhost:${config.port}`);
});

async function shutdown() {
  server.close(async () => {
    await pool.end();
    process.exit(0);
  });
}

process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);
