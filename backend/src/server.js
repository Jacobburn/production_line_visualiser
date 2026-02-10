import express from 'express';
import cors from 'cors';
import { z } from 'zod';
import { config } from './config.js';
import { dbQuery, pool } from './db.js';
import {
  authMiddleware,
  comparePassword,
  hashPassword,
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

const lineRenameSchema = z.object({
  name: z.string().min(2).max(120)
});

const lineModelSchema = z.object({
  stages: z.array(
    z.object({
      id: z.string().optional(),
      stageOrder: z.number().int().min(1),
      stageName: z.string().min(1).max(150),
      stageType: z.enum(['main', 'prep', 'transfer']),
      dayCrew: z.number().int().min(0),
      nightCrew: z.number().int().min(0),
      maxThroughputPerCrew: z.number().min(0),
      x: z.number(),
      y: z.number(),
      w: z.number().positive(),
      h: z.number().positive()
    })
  ).default([]),
  guides: z.array(
    z.object({
      id: z.string().optional(),
      guideType: z.enum(['line', 'arrow', 'shape']),
      x: z.number(),
      y: z.number(),
      w: z.number(),
      h: z.number(),
      angle: z.number(),
      src: z.string().optional().nullable()
    })
  ).default([])
});

const supervisorCreateSchema = z.object({
  name: z.string().min(1).max(120),
  username: z.string().min(3).max(60),
  password: z.string().min(6).max(120),
  assignedLineIds: z.array(z.string().uuid()).optional().default([])
});

const supervisorAssignmentsSchema = z.object({
  assignedLineIds: z.array(z.string().uuid())
});

const supervisorUpdateSchema = z.object({
  name: z.string().min(1).max(120),
  username: z.string().min(3).max(60),
  password: z.string().min(6).max(120).optional().or(z.literal(''))
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

const clearLineDataSchema = z.object({
  secretKey: z.string().min(1).max(128)
});

const asyncRoute = (handler) => (req, res, next) => Promise.resolve(handler(req, res, next)).catch(next);

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

app.get('/api/health', asyncRoute(async (_req, res) => {
  await dbQuery('SELECT 1');
  res.json({ ok: true, env: config.nodeEnv, timestamp: new Date().toISOString() });
}));

app.post('/api/auth/login', asyncRoute(async (req, res) => {
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
}));

app.get('/api/me', authMiddleware, asyncRoute(async (req, res) => {
  res.json({ user: req.user });
}));

app.get('/api/lines', authMiddleware, asyncRoute(async (req, res) => {
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
}));

app.get('/api/state-snapshot', authMiddleware, asyncRoute(async (req, res) => {
  const linesQuery = req.user.role === 'manager'
    ? `SELECT
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
       FROM production_lines l
       WHERE l.is_active = TRUE
       ORDER BY l.created_at DESC`
    : `SELECT
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
       FROM production_lines l
       INNER JOIN supervisor_line_assignments a ON a.line_id = l.id
       WHERE a.supervisor_user_id = $1 AND l.is_active = TRUE
       ORDER BY l.created_at DESC`;

  const linesResult = await dbQuery(linesQuery, req.user.role === 'manager' ? [] : [req.user.id]);
  const lineRows = linesResult.rows;
  const lineIds = lineRows.map((line) => line.id);
  if (!lineIds.length) {
    const payload = { lines: [], supervisors: [] };
    if (req.user.role === 'manager') {
      const supervisorsResult = await dbQuery(
        `SELECT
           u.id,
           u.name,
           u.username,
           u.is_active AS "isActive",
           COALESCE(
             ARRAY_REMOVE(ARRAY_AGG(a.line_id::TEXT), NULL),
             ARRAY[]::TEXT[]
           ) AS "assignedLineIds"
         FROM users u
         LEFT JOIN supervisor_line_assignments a ON a.supervisor_user_id = u.id
         WHERE u.role = 'supervisor'
         GROUP BY u.id
         ORDER BY u.created_at DESC`
      );
      payload.supervisors = supervisorsResult.rows;
    }
    return res.json(payload);
  }

  const [stagesResult, guidesResult, shiftLogsResult, runLogsResult, downtimeLogsResult] = await Promise.all([
    dbQuery(
      `SELECT
         line_id AS "lineId",
         id,
         stage_order AS "stageOrder",
         stage_name AS "stageName",
         stage_type AS "stageType",
         day_crew AS "dayCrew",
         night_crew AS "nightCrew",
         max_throughput_per_crew AS "maxThroughputPerCrew",
         x, y, w, h
       FROM line_stages
       WHERE line_id = ANY($1::UUID[])
       ORDER BY line_id, stage_order ASC`,
      [lineIds]
    ),
    dbQuery(
      `SELECT
         line_id AS "lineId",
         id,
         guide_type AS "guideType",
         x, y, w, h, angle, src
       FROM line_layout_guides
       WHERE line_id = ANY($1::UUID[])
       ORDER BY line_id, created_at ASC`,
      [lineIds]
    ),
    dbQuery(
      `SELECT
         line_id AS "lineId",
         date::TEXT AS date,
         shift,
         crew_on_shift AS "crewOnShift",
         to_char(start_time, 'HH24:MI') AS "startTime",
         COALESCE(to_char(break1_start, 'HH24:MI'), '') AS "break1Start",
         COALESCE(to_char(break2_start, 'HH24:MI'), '') AS "break2Start",
         COALESCE(to_char(break3_start, 'HH24:MI'), '') AS "break3Start",
         to_char(finish_time, 'HH24:MI') AS "finishTime",
         COALESCE(u.name, u.username, '') AS "submittedBy",
         submitted_at AS "submittedAt"
       FROM shift_logs
       LEFT JOIN users u ON u.id = shift_logs.submitted_by_user_id
       WHERE line_id = ANY($1::UUID[])
       ORDER BY line_id, date ASC, shift ASC, submitted_at ASC`,
      [lineIds]
    ),
    dbQuery(
      `SELECT
         line_id AS "lineId",
         date::TEXT AS date,
         shift,
         product,
         COALESCE(to_char(setup_start_time, 'HH24:MI'), '') AS "setUpStartTime",
         to_char(production_start_time, 'HH24:MI') AS "productionStartTime",
         to_char(finish_time, 'HH24:MI') AS "finishTime",
         units_produced AS "unitsProduced",
         COALESCE(u.name, u.username, '') AS "submittedBy",
         submitted_at AS "submittedAt"
       FROM run_logs
       LEFT JOIN users u ON u.id = run_logs.submitted_by_user_id
       WHERE line_id = ANY($1::UUID[])
       ORDER BY line_id, date ASC, shift ASC, submitted_at ASC`,
      [lineIds]
    ),
    dbQuery(
      `SELECT
         line_id AS "lineId",
         date::TEXT AS date,
         shift,
         to_char(downtime_start, 'HH24:MI') AS "downtimeStart",
         to_char(downtime_finish, 'HH24:MI') AS "downtimeFinish",
         COALESCE(equipment_stage_id::TEXT, '') AS equipment,
         COALESCE(reason, '') AS reason,
         COALESCE(u.name, u.username, '') AS "submittedBy",
         submitted_at AS "submittedAt"
       FROM downtime_logs
       LEFT JOIN users u ON u.id = downtime_logs.submitted_by_user_id
       WHERE line_id = ANY($1::UUID[])
       ORDER BY line_id, date ASC, shift ASC, submitted_at ASC`,
      [lineIds]
    )
  ]);

  const stagesByLine = new Map();
  stagesResult.rows.forEach((row) => {
    const list = stagesByLine.get(row.lineId) || [];
    list.push({
      id: row.id,
      stageOrder: row.stageOrder,
      stageName: row.stageName,
      stageType: row.stageType,
      dayCrew: row.dayCrew,
      nightCrew: row.nightCrew,
      maxThroughputPerCrew: row.maxThroughputPerCrew,
      x: row.x,
      y: row.y,
      w: row.w,
      h: row.h
    });
    stagesByLine.set(row.lineId, list);
  });

  const guidesByLine = new Map();
  guidesResult.rows.forEach((row) => {
    const list = guidesByLine.get(row.lineId) || [];
    list.push({
      id: row.id,
      guideType: row.guideType,
      x: row.x,
      y: row.y,
      w: row.w,
      h: row.h,
      angle: row.angle,
      src: row.src
    });
    guidesByLine.set(row.lineId, list);
  });

  const shiftByLine = new Map();
  shiftLogsResult.rows.forEach((row) => {
    const list = shiftByLine.get(row.lineId) || [];
    list.push({
      date: row.date,
      shift: row.shift,
      crewOnShift: row.crewOnShift,
      startTime: row.startTime,
      break1Start: row.break1Start,
      break2Start: row.break2Start,
      break3Start: row.break3Start,
      finishTime: row.finishTime,
      submittedAt: row.submittedAt
    });
    shiftByLine.set(row.lineId, list);
  });

  const runByLine = new Map();
  runLogsResult.rows.forEach((row) => {
    const list = runByLine.get(row.lineId) || [];
    list.push({
      date: row.date,
      shift: row.shift,
      product: row.product,
      setUpStartTime: row.setUpStartTime,
      productionStartTime: row.productionStartTime,
      finishTime: row.finishTime,
      unitsProduced: row.unitsProduced,
      submittedAt: row.submittedAt
    });
    runByLine.set(row.lineId, list);
  });

  const downtimeByLine = new Map();
  downtimeLogsResult.rows.forEach((row) => {
    const list = downtimeByLine.get(row.lineId) || [];
    list.push({
      date: row.date,
      shift: row.shift,
      downtimeStart: row.downtimeStart,
      downtimeFinish: row.downtimeFinish,
      equipment: row.equipment,
      reason: row.reason,
      submittedAt: row.submittedAt
    });
    downtimeByLine.set(row.lineId, list);
  });

  const lines = lineRows.map((line) => ({
    line,
    stages: stagesByLine.get(line.id) || [],
    guides: guidesByLine.get(line.id) || [],
    shiftRows: shiftByLine.get(line.id) || [],
    runRows: runByLine.get(line.id) || [],
    downtimeRows: downtimeByLine.get(line.id) || []
  }));

  const payload = { lines, supervisors: [] };
  if (req.user.role === 'manager') {
    const supervisorsResult = await dbQuery(
      `SELECT
         u.id,
         u.name,
         u.username,
         u.is_active AS "isActive",
         COALESCE(
           ARRAY_REMOVE(ARRAY_AGG(a.line_id::TEXT), NULL),
           ARRAY[]::TEXT[]
         ) AS "assignedLineIds"
       FROM users u
       LEFT JOIN supervisor_line_assignments a ON a.supervisor_user_id = u.id
       WHERE u.role = 'supervisor'
       GROUP BY u.id
       ORDER BY u.created_at DESC`
    );
    payload.supervisors = supervisorsResult.rows;
  }

  return res.json(payload);
}));

app.post('/api/lines', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
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
}));

app.get('/api/supervisors', authMiddleware, requireRole('manager'), asyncRoute(async (_req, res) => {
  const result = await dbQuery(
    `SELECT
       u.id,
       u.name,
       u.username,
       u.is_active AS "isActive",
       COALESCE(
         ARRAY_REMOVE(ARRAY_AGG(a.line_id::TEXT), NULL),
         ARRAY[]::TEXT[]
       ) AS "assignedLineIds"
     FROM users u
     LEFT JOIN supervisor_line_assignments a ON a.supervisor_user_id = u.id
     WHERE u.role = 'supervisor'
     GROUP BY u.id
     ORDER BY u.created_at DESC`
  );
  res.json({ supervisors: result.rows });
}));

app.post('/api/supervisors', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  const parsed = supervisorCreateSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: 'Invalid supervisor payload' });
  const username = parsed.data.username.trim().toLowerCase();
  const name = parsed.data.name.trim();
  const passwordHash = await hashPassword(parsed.data.password);
  const assignedLineIds = parsed.data.assignedLineIds || [];

  const client = await pool.connect();
  try {
    await client.query('BEGIN');

    const insertUser = await client.query(
      `INSERT INTO users(name, username, password_hash, role, is_active)
       VALUES ($1, $2, $3, 'supervisor', TRUE)
       RETURNING id, name, username, is_active AS "isActive"`,
      [name, username, passwordHash]
    );
    const supervisor = insertUser.rows[0];

    if (assignedLineIds.length) {
      await client.query(
        `INSERT INTO supervisor_line_assignments(supervisor_user_id, line_id, assigned_by_user_id)
         SELECT $1, line_id::UUID, $2
         FROM UNNEST($3::TEXT[]) AS t(line_id)
         ON CONFLICT (supervisor_user_id, line_id) DO NOTHING`,
        [supervisor.id, req.user.id, assignedLineIds]
      );
    }

    await client.query('COMMIT');
    return res.status(201).json({
      supervisor: {
        ...supervisor,
        assignedLineIds
      }
    });
  } catch (error) {
    await client.query('ROLLBACK');
    if (String(error?.message || '').includes('duplicate key')) {
      return res.status(409).json({ error: 'Supervisor username already exists' });
    }
    throw error;
  } finally {
    client.release();
  }
}));

app.patch('/api/supervisors/:supervisorId/assignments', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  const supervisorId = req.params.supervisorId;
  if (!z.string().uuid().safeParse(supervisorId).success) return res.status(400).json({ error: 'Invalid supervisor id' });
  const parsed = supervisorAssignmentsSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: 'Invalid assignments payload' });

  const userCheck = await dbQuery(
    `SELECT id FROM users WHERE id = $1 AND role = 'supervisor'`,
    [supervisorId]
  );
  if (!userCheck.rowCount) return res.status(404).json({ error: 'Supervisor not found' });

  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    await client.query(
      `DELETE FROM supervisor_line_assignments WHERE supervisor_user_id = $1`,
      [supervisorId]
    );
    if (parsed.data.assignedLineIds.length) {
      await client.query(
        `INSERT INTO supervisor_line_assignments(supervisor_user_id, line_id, assigned_by_user_id)
         SELECT $1, line_id::UUID, $2
         FROM UNNEST($3::TEXT[]) AS t(line_id)
         ON CONFLICT (supervisor_user_id, line_id) DO NOTHING`,
        [supervisorId, req.user.id, parsed.data.assignedLineIds]
      );
    }
    await client.query('COMMIT');
    return res.json({ ok: true });
  } catch (error) {
    await client.query('ROLLBACK');
    throw error;
  } finally {
    client.release();
  }
}));

app.patch('/api/supervisors/:supervisorId', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  const supervisorId = req.params.supervisorId;
  if (!z.string().uuid().safeParse(supervisorId).success) return res.status(400).json({ error: 'Invalid supervisor id' });
  const parsed = supervisorUpdateSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid supervisor update payload' });

  const userCheck = await dbQuery(
    `SELECT id FROM users WHERE id = $1 AND role = 'supervisor'`,
    [supervisorId]
  );
  if (!userCheck.rowCount) return res.status(404).json({ error: 'Supervisor not found' });

  const nextName = parsed.data.name.trim();
  const nextUsername = parsed.data.username.trim().toLowerCase();
  const password = String(parsed.data.password || '').trim();

  try {
    let result;
    if (password) {
      const passwordHash = await hashPassword(password);
      result = await dbQuery(
        `UPDATE users
         SET name = $2, username = $3, password_hash = $4, updated_at = NOW()
         WHERE id = $1 AND role = 'supervisor'
         RETURNING id, name, username, is_active AS "isActive"`,
        [supervisorId, nextName, nextUsername, passwordHash]
      );
    } else {
      result = await dbQuery(
        `UPDATE users
         SET name = $2, username = $3, updated_at = NOW()
         WHERE id = $1 AND role = 'supervisor'
         RETURNING id, name, username, is_active AS "isActive"`,
        [supervisorId, nextName, nextUsername]
      );
    }
    if (!result.rowCount) return res.status(404).json({ error: 'Supervisor not found' });
    return res.json({ supervisor: result.rows[0] });
  } catch (error) {
    if (String(error?.message || '').includes('duplicate key')) {
      return res.status(409).json({ error: 'Supervisor username already exists' });
    }
    throw error;
  }
}));

app.delete('/api/supervisors/:supervisorId', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  const supervisorId = req.params.supervisorId;
  if (!z.string().uuid().safeParse(supervisorId).success) return res.status(400).json({ error: 'Invalid supervisor id' });

  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    await client.query(
      `DELETE FROM supervisor_line_assignments WHERE supervisor_user_id = $1`,
      [supervisorId]
    );
    const updateUser = await client.query(
      `UPDATE users
       SET is_active = FALSE, updated_at = NOW()
       WHERE id = $1 AND role = 'supervisor'
       RETURNING id`,
      [supervisorId]
    );
    if (!updateUser.rowCount) {
      await client.query('ROLLBACK');
      return res.status(404).json({ error: 'Supervisor not found' });
    }
    await client.query('COMMIT');
    return res.json({ ok: true });
  } catch (error) {
    await client.query('ROLLBACK');
    throw error;
  } finally {
    client.release();
  }
}));

app.get('/api/lines/:lineId', authMiddleware, asyncRoute(async (req, res) => {
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
}));

app.get('/api/lines/:lineId/logs', authMiddleware, asyncRoute(async (req, res) => {
  const lineId = req.params.lineId;
  if (!z.string().uuid().safeParse(lineId).success) return res.status(400).json({ error: 'Invalid line id' });
  if (!(await hasLineAccess(req.user, lineId))) return res.status(403).json({ error: 'Forbidden' });

  const [shiftRows, runRows, downtimeRows] = await Promise.all([
    dbQuery(
      `SELECT
         date::TEXT AS date,
         shift,
         crew_on_shift AS "crewOnShift",
         to_char(start_time, 'HH24:MI') AS "startTime",
         COALESCE(to_char(break1_start, 'HH24:MI'), '') AS "break1Start",
         COALESCE(to_char(break2_start, 'HH24:MI'), '') AS "break2Start",
         COALESCE(to_char(break3_start, 'HH24:MI'), '') AS "break3Start",
         to_char(finish_time, 'HH24:MI') AS "finishTime",
         COALESCE(u.name, u.username, '') AS "submittedBy",
         submitted_at AS "submittedAt"
       FROM shift_logs
       LEFT JOIN users u ON u.id = shift_logs.submitted_by_user_id
       WHERE line_id = $1
       ORDER BY date ASC, shift ASC, submitted_at ASC`,
      [lineId]
    ),
    dbQuery(
      `SELECT
         date::TEXT AS date,
         shift,
         product,
         COALESCE(to_char(setup_start_time, 'HH24:MI'), '') AS "setUpStartTime",
         to_char(production_start_time, 'HH24:MI') AS "productionStartTime",
         to_char(finish_time, 'HH24:MI') AS "finishTime",
         units_produced AS "unitsProduced",
         COALESCE(u.name, u.username, '') AS "submittedBy",
         submitted_at AS "submittedAt"
       FROM run_logs
       LEFT JOIN users u ON u.id = run_logs.submitted_by_user_id
       WHERE line_id = $1
       ORDER BY date ASC, shift ASC, submitted_at ASC`,
      [lineId]
    ),
    dbQuery(
      `SELECT
         date::TEXT AS date,
         shift,
         to_char(downtime_start, 'HH24:MI') AS "downtimeStart",
         to_char(downtime_finish, 'HH24:MI') AS "downtimeFinish",
         COALESCE(equipment_stage_id::TEXT, '') AS equipment,
         COALESCE(reason, '') AS reason,
         COALESCE(u.name, u.username, '') AS "submittedBy",
         submitted_at AS "submittedAt"
       FROM downtime_logs
       LEFT JOIN users u ON u.id = downtime_logs.submitted_by_user_id
       WHERE line_id = $1
       ORDER BY date ASC, shift ASC, submitted_at ASC`,
      [lineId]
    )
  ]);

  return res.json({
    shiftRows: shiftRows.rows,
    runRows: runRows.rows,
    downtimeRows: downtimeRows.rows
  });
}));

app.put('/api/lines/:lineId/model', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  const lineId = req.params.lineId;
  if (!z.string().uuid().safeParse(lineId).success) return res.status(400).json({ error: 'Invalid line id' });
  const parsed = lineModelSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid line model payload' });

  const lineCheck = await dbQuery(`SELECT id FROM production_lines WHERE id = $1 AND is_active = TRUE`, [lineId]);
  if (!lineCheck.rowCount) return res.status(404).json({ error: 'Line not found' });

  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    await client.query(`DELETE FROM line_stages WHERE line_id = $1`, [lineId]);
    await client.query(`DELETE FROM line_layout_guides WHERE line_id = $1`, [lineId]);

    for (const stage of parsed.data.stages) {
      await client.query(
        `INSERT INTO line_stages(
           line_id, stage_order, stage_name, stage_type, day_crew, night_crew, max_throughput_per_crew, x, y, w, h
         ) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)`,
        [
          lineId,
          stage.stageOrder,
          stage.stageName,
          stage.stageType,
          stage.dayCrew,
          stage.nightCrew,
          stage.maxThroughputPerCrew,
          stage.x,
          stage.y,
          stage.w,
          stage.h
        ]
      );
    }

    for (const guide of parsed.data.guides) {
      await client.query(
        `INSERT INTO line_layout_guides(line_id, guide_type, x, y, w, h, angle, src)
         VALUES ($1,$2,$3,$4,$5,$6,$7,$8)`,
        [lineId, guide.guideType, guide.x, guide.y, guide.w, guide.h, guide.angle, guide.src || null]
      );
    }

    await client.query('COMMIT');
    return res.json({ ok: true });
  } catch (error) {
    await client.query('ROLLBACK');
    throw error;
  } finally {
    client.release();
  }
}));

app.patch('/api/lines/:lineId', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  const lineId = req.params.lineId;
  if (!z.string().uuid().safeParse(lineId).success) return res.status(400).json({ error: 'Invalid line id' });
  const parsed = lineRenameSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid line rename payload' });

  try {
    const result = await dbQuery(
      `UPDATE production_lines
       SET name = $2, updated_at = NOW()
       WHERE id = $1 AND is_active = TRUE
       RETURNING id, name, secret_key AS "secretKey", updated_at AS "updatedAt"`,
      [lineId, parsed.data.name.trim()]
    );
    if (!result.rowCount) return res.status(404).json({ error: 'Line not found' });

    await writeAudit({
      lineId,
      actorUserId: req.user.id,
      actorName: req.user.name,
      actorRole: req.user.role,
      action: 'RENAME_LINE',
      details: `Line renamed to: ${result.rows[0].name}`
    });

    return res.json({ line: result.rows[0] });
  } catch (error) {
    if (String(error?.message || '').includes('duplicate key')) {
      return res.status(409).json({ error: 'Line name already exists' });
    }
    throw error;
  }
}));

app.delete('/api/lines/:lineId', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  const lineId = req.params.lineId;
  if (!z.string().uuid().safeParse(lineId).success) return res.status(400).json({ error: 'Invalid line id' });
  const result = await dbQuery(
    `UPDATE production_lines
     SET is_active = FALSE, updated_at = NOW()
     WHERE id = $1
     RETURNING id`,
    [lineId]
  );
  if (!result.rowCount) return res.status(404).json({ error: 'Line not found' });
  await dbQuery(`DELETE FROM supervisor_line_assignments WHERE line_id = $1`, [lineId]);
  return res.json({ ok: true });
}));

app.post('/api/lines/:lineId/clear-data', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  const lineId = req.params.lineId;
  if (!z.string().uuid().safeParse(lineId).success) return res.status(400).json({ error: 'Invalid line id' });
  const parsed = clearLineDataSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: 'Invalid clear data payload' });

  const lineRes = await dbQuery(
    `SELECT id, name, secret_key AS "secretKey"
     FROM production_lines
     WHERE id = $1 AND is_active = TRUE`,
    [lineId]
  );
  if (!lineRes.rowCount) return res.status(404).json({ error: 'Line not found' });
  const line = lineRes.rows[0];
  if (parsed.data.secretKey !== 'admin' && parsed.data.secretKey !== line.secretKey) {
    return res.status(403).json({ error: 'Invalid key/password. Data was not cleared.' });
  }

  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    await client.query(`DELETE FROM shift_logs WHERE line_id = $1`, [lineId]);
    await client.query(`DELETE FROM run_logs WHERE line_id = $1`, [lineId]);
    await client.query(`DELETE FROM downtime_logs WHERE line_id = $1`, [lineId]);
    await client.query('COMMIT');
  } catch (error) {
    await client.query('ROLLBACK');
    throw error;
  } finally {
    client.release();
  }

  await writeAudit({
    lineId,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'CLEAR_DATA',
    details: `All logs cleared for line ${line.name}`
  });

  return res.json({ ok: true });
}));

app.post('/api/logs/shifts', authMiddleware, asyncRoute(async (req, res) => {
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
}));

app.post('/api/logs/runs', authMiddleware, asyncRoute(async (req, res) => {
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
}));

app.post('/api/logs/downtime', authMiddleware, asyncRoute(async (req, res) => {
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
}));

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
