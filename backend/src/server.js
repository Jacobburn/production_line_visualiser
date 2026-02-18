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

const shiftValues = ['Day', 'Night', 'Full Day'];
const supervisorShiftValues = ['Day', 'Night'];
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
  assignedLineIds: z.array(z.string().uuid()).optional().default([]),
  assignedLineShifts: z.record(z.array(z.enum(supervisorShiftValues))).optional().default({})
});

const supervisorAssignmentsSchema = z.object({
  assignedLineIds: z.array(z.string().uuid()).optional().default([]),
  assignedLineShifts: z.record(z.array(z.enum(supervisorShiftValues))).optional().default({})
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
  finishTime: z.string().regex(timeRegex)
});

const runLogSchema = z.object({
  lineId: z.string().uuid(),
  date: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
  shift: z.enum(shiftValues),
  product: z.string().min(1).max(120),
  setUpStartTime: optionalTime,
  productionStartTime: z.string().regex(timeRegex),
  finishTime: z.string().regex(timeRegex),
  unitsProduced: z.number().nonnegative(),
  runCrewingPattern: z.record(z.number().int().min(0)).optional().default({})
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

const hasAtLeastOneField = (obj) => Object.values(obj).some((value) => value !== undefined);

const shiftLogUpdateSchema = z.object({
  crewOnShift: z.number().int().min(0).optional(),
  startTime: optionalTime.optional(),
  finishTime: optionalTime.optional()
}).refine(hasAtLeastOneField, { message: 'At least one field is required' });

const shiftBreakStartSchema = z.object({
  breakStart: z.string().regex(timeRegex)
});

const shiftBreakUpdateSchema = z.object({
  breakStart: z.string().regex(timeRegex).optional(),
  breakFinish: z.string().regex(timeRegex).optional()
}).refine(hasAtLeastOneField, { message: 'At least one field is required' });

const runLogUpdateSchema = z.object({
  product: z.string().min(1).max(120).optional(),
  setUpStartTime: optionalTime.optional(),
  productionStartTime: optionalTime.optional(),
  finishTime: optionalTime.optional(),
  unitsProduced: z.number().nonnegative().optional(),
  runCrewingPattern: z.record(z.number().int().min(0)).optional()
}).refine(hasAtLeastOneField, { message: 'At least one field is required' });

const downtimeLogUpdateSchema = z.object({
  downtimeStart: optionalTime.optional(),
  downtimeFinish: optionalTime.optional(),
  equipmentStageId: z.string().uuid().optional().or(z.literal('')).or(z.null()),
  reason: z.string().max(250).optional().or(z.literal('')).or(z.null())
}).refine(hasAtLeastOneField, { message: 'At least one field is required' });

const clearLineDataSchema = z.object({
  secretKey: z.string().min(1).max(128)
});

const asyncRoute = (handler) => (req, res, next) => Promise.resolve(handler(req, res, next)).catch(next);

function normalizeAssignedLineShifts(rawMap, fallbackLineIds = []) {
  const next = {};
  if (rawMap && typeof rawMap === 'object' && !Array.isArray(rawMap)) {
    Object.entries(rawMap).forEach(([lineId, shifts]) => {
      if (!z.string().uuid().safeParse(lineId).success || !Array.isArray(shifts)) return;
      const unique = Array.from(new Set(shifts.filter((shift) => supervisorShiftValues.includes(shift))));
      if (unique.length) next[lineId] = unique;
    });
  }
  if (!Object.keys(next).length && Array.isArray(fallbackLineIds)) {
    fallbackLineIds.forEach((lineId) => {
      if (!z.string().uuid().safeParse(lineId).success) return;
      next[lineId] = supervisorShiftValues.slice();
    });
  }
  return next;
}

function flattenAssignedLineShifts(lineShifts) {
  return Object.entries(lineShifts || {}).flatMap(([lineId, shifts]) =>
    (Array.isArray(shifts) ? shifts : [])
      .filter((shift) => supervisorShiftValues.includes(shift))
      .map((shift) => ({ lineId, shift }))
  );
}

async function fetchSupervisorsWithAssignments() {
  const [supervisorsResult, assignmentsResult] = await Promise.all([
    dbQuery(
      `SELECT
         u.id,
         u.name,
         u.username,
         u.is_active AS "isActive"
       FROM users u
       WHERE u.role = 'supervisor'
       ORDER BY u.created_at DESC`
    ),
    dbQuery(
      `SELECT
         supervisor_user_id AS "supervisorUserId",
         line_id::TEXT AS "lineId",
         shift
       FROM supervisor_line_shift_assignments
       ORDER BY assigned_at ASC`
    )
  ]);

  const bySupervisor = new Map();
  assignmentsResult.rows.forEach((row) => {
    if (!supervisorShiftValues.includes(row.shift)) return;
    const current = bySupervisor.get(row.supervisorUserId) || {};
    if (!current[row.lineId]) current[row.lineId] = [];
    if (!current[row.lineId].includes(row.shift)) current[row.lineId].push(row.shift);
    bySupervisor.set(row.supervisorUserId, current);
  });

  return supervisorsResult.rows.map((sup) => {
    const assignedLineShifts = bySupervisor.get(sup.id) || {};
    return {
      ...sup,
      assignedLineIds: Object.keys(assignedLineShifts),
      assignedLineShifts
    };
  });
}

async function fetchSupervisorAssignments(supervisorUserId) {
  const result = await dbQuery(
    `SELECT
       line_id::TEXT AS "lineId",
       shift
     FROM supervisor_line_shift_assignments
     WHERE supervisor_user_id = $1
     ORDER BY assigned_at ASC`,
    [supervisorUserId]
  );
  const assignedLineShifts = {};
  result.rows.forEach((row) => {
    if (!supervisorShiftValues.includes(row.shift)) return;
    if (!assignedLineShifts[row.lineId]) assignedLineShifts[row.lineId] = [];
    if (!assignedLineShifts[row.lineId].includes(row.shift)) assignedLineShifts[row.lineId].push(row.shift);
  });
  return {
    assignedLineIds: Object.keys(assignedLineShifts),
    assignedLineShifts
  };
}

async function hasLineAccess(user, lineId) {
  if (user.role === 'manager') return true;
  const result = await dbQuery(
    `SELECT 1
     FROM supervisor_line_shift_assignments
     WHERE supervisor_user_id = $1 AND line_id = $2
     LIMIT 1`,
    [user.id, lineId]
  );
  return result.rowCount > 0;
}

async function isActiveLine(lineId) {
  const result = await dbQuery(
    `SELECT 1
     FROM production_lines
     WHERE id = $1
       AND is_active = TRUE
     LIMIT 1`,
    [lineId]
  );
  return result.rowCount > 0;
}

async function hasLineShiftAccess(user, lineId, shift) {
  if (!shiftValues.includes(shift)) return false;
  if (user.role === 'manager') return isActiveLine(lineId);
  const result = await dbQuery(
    `SELECT 1
     FROM supervisor_line_shift_assignments a
     INNER JOIN production_lines l
       ON l.id = a.line_id
     WHERE a.supervisor_user_id = $1
       AND a.line_id = $2
       AND l.is_active = TRUE
       AND (
         ($3 = 'Full Day' AND a.shift IN ('Day', 'Night'))
         OR a.shift = $3
       )
     LIMIT 1`,
    [user.id, lineId, shift]
  );
  return result.rowCount > 0;
}

async function isLineStage(lineId, stageId) {
  const result = await dbQuery(
    `SELECT 1
     FROM line_stages
     WHERE id = $1 AND line_id = $2
     LIMIT 1`,
    [stageId, lineId]
  );
  return result.rowCount > 0;
}

async function findMissingActiveLineIds(lineIds) {
  const uniqueLineIds = Array.from(
    new Set((Array.isArray(lineIds) ? lineIds : []).filter((lineId) => z.string().uuid().safeParse(lineId).success))
  );
  if (!uniqueLineIds.length) return [];
  const result = await dbQuery(
    `SELECT id::TEXT AS id
     FROM production_lines
     WHERE is_active = TRUE
       AND id = ANY($1::UUID[])`,
    [uniqueLineIds]
  );
  const foundIds = new Set(result.rows.map((row) => row.id));
  return uniqueLineIds.filter((lineId) => !foundIds.has(lineId));
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
  const managerBaseFields = `
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
  const supervisorBaseFields = `
    l.id,
    l.name,
    NULL::TEXT AS "secretKey",
    l.created_at AS "createdAt",
    (
      SELECT COUNT(*)::INT FROM line_stages s WHERE s.line_id = l.id
    ) AS "stageCount",
    (
      SELECT COUNT(*)::INT
      FROM shift_logs sl
      WHERE sl.line_id = l.id
        AND EXISTS (
          SELECT 1
          FROM supervisor_line_shift_assignments a_count
          WHERE a_count.supervisor_user_id = $1
            AND a_count.line_id = sl.line_id
            AND (
              a_count.shift = sl.shift
              OR (sl.shift = 'Full Day' AND a_count.shift IN ('Day', 'Night'))
            )
        )
    ) AS "shiftCount"
  `;

  const query = req.user.role === 'manager'
    ? `SELECT ${managerBaseFields} FROM production_lines l WHERE l.is_active = TRUE ORDER BY l.created_at DESC`
    : `SELECT ${supervisorBaseFields}
       FROM production_lines l
       INNER JOIN (
         SELECT DISTINCT line_id
         FROM supervisor_line_shift_assignments
         WHERE supervisor_user_id = $1
       ) a ON a.line_id = l.id
       WHERE l.is_active = TRUE
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
         NULL::TEXT AS "secretKey",
         l.created_at AS "createdAt",
         (
           SELECT COUNT(*)::INT FROM line_stages s WHERE s.line_id = l.id
         ) AS "stageCount",
         (
           SELECT COUNT(*)::INT
           FROM shift_logs sl
           WHERE sl.line_id = l.id
             AND EXISTS (
               SELECT 1
               FROM supervisor_line_shift_assignments a_count
               WHERE a_count.supervisor_user_id = $1
                 AND a_count.line_id = sl.line_id
                 AND (
                   a_count.shift = sl.shift
                   OR (sl.shift = 'Full Day' AND a_count.shift IN ('Day', 'Night'))
                 )
             )
         ) AS "shiftCount"
       FROM production_lines l
       INNER JOIN (
         SELECT DISTINCT line_id
         FROM supervisor_line_shift_assignments
         WHERE supervisor_user_id = $1
       ) a ON a.line_id = l.id
       WHERE l.is_active = TRUE
       ORDER BY l.created_at DESC`;

  const isManager = req.user.role === 'manager';
  const linesResult = await dbQuery(linesQuery, isManager ? [] : [req.user.id]);
  const lineRows = linesResult.rows;
  const lineIds = lineRows.map((line) => line.id);
  if (!lineIds.length) {
    const payload = { lines: [], supervisors: [] };
    if (req.user.role === 'manager') {
      payload.supervisors = await fetchSupervisorsWithAssignments();
    } else if (req.user.role === 'supervisor') {
      payload.supervisorAssignments = await fetchSupervisorAssignments(req.user.id);
    }
    return res.json(payload);
  }

  const snapshotScopedParams = isManager ? [lineIds] : [req.user.id, lineIds];
  const [stagesResult, guidesResult, shiftLogsResult, breakLogsResult, runLogsResult, downtimeLogsResult] = await Promise.all([
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
      isManager
        ? `SELECT
             line_id AS "lineId",
             shift_logs.id,
             date::TEXT AS date,
             shift,
             crew_on_shift AS "crewOnShift",
             to_char(start_time, 'HH24:MI') AS "startTime",
             to_char(finish_time, 'HH24:MI') AS "finishTime",
             COALESCE(u.name, u.username, '') AS "submittedBy",
             submitted_at AS "submittedAt"
           FROM shift_logs
           LEFT JOIN users u ON u.id = shift_logs.submitted_by_user_id
           WHERE line_id = ANY($1::UUID[])
           ORDER BY line_id, date ASC, shift ASC, submitted_at ASC`
        : `SELECT
             line_id AS "lineId",
             shift_logs.id,
             date::TEXT AS date,
             shift,
             crew_on_shift AS "crewOnShift",
             to_char(start_time, 'HH24:MI') AS "startTime",
             to_char(finish_time, 'HH24:MI') AS "finishTime",
             COALESCE(u.name, u.username, '') AS "submittedBy",
             submitted_at AS "submittedAt"
           FROM shift_logs
           LEFT JOIN users u ON u.id = shift_logs.submitted_by_user_id
           WHERE line_id = ANY($2::UUID[])
             AND EXISTS (
               SELECT 1
               FROM supervisor_line_shift_assignments a
               WHERE a.supervisor_user_id = $1
                 AND a.line_id = shift_logs.line_id
                 AND (
                   a.shift = shift_logs.shift
                   OR (shift_logs.shift = 'Full Day' AND a.shift IN ('Day', 'Night'))
                 )
             )
           ORDER BY line_id, date ASC, shift ASC, submitted_at ASC`,
      snapshotScopedParams
    ),
    dbQuery(
      isManager
        ? `SELECT
             line_id AS "lineId",
             shift_break_logs.id,
             shift_log_id AS "shiftLogId",
             date::TEXT AS date,
             shift,
             to_char(break_start, 'HH24:MI') AS "breakStart",
             COALESCE(to_char(break_finish, 'HH24:MI'), '') AS "breakFinish",
             COALESCE(u.name, u.username, '') AS "submittedBy",
             submitted_at AS "submittedAt"
           FROM shift_break_logs
           LEFT JOIN users u ON u.id = shift_break_logs.submitted_by_user_id
           WHERE line_id = ANY($1::UUID[])
           ORDER BY line_id, date ASC, shift ASC, submitted_at ASC`
        : `SELECT
             line_id AS "lineId",
             shift_break_logs.id,
             shift_log_id AS "shiftLogId",
             date::TEXT AS date,
             shift,
             to_char(break_start, 'HH24:MI') AS "breakStart",
             COALESCE(to_char(break_finish, 'HH24:MI'), '') AS "breakFinish",
             COALESCE(u.name, u.username, '') AS "submittedBy",
             submitted_at AS "submittedAt"
           FROM shift_break_logs
           LEFT JOIN users u ON u.id = shift_break_logs.submitted_by_user_id
           WHERE line_id = ANY($2::UUID[])
             AND EXISTS (
               SELECT 1
               FROM supervisor_line_shift_assignments a
               WHERE a.supervisor_user_id = $1
                 AND a.line_id = shift_break_logs.line_id
                 AND (
                   a.shift = shift_break_logs.shift
                   OR (shift_break_logs.shift = 'Full Day' AND a.shift IN ('Day', 'Night'))
                 )
             )
           ORDER BY line_id, date ASC, shift ASC, submitted_at ASC`,
      snapshotScopedParams
    ),
    dbQuery(
      isManager
        ? `SELECT
             line_id AS "lineId",
             run_logs.id,
             date::TEXT AS date,
             shift,
             product,
             COALESCE(to_char(setup_start_time, 'HH24:MI'), '') AS "setUpStartTime",
             to_char(production_start_time, 'HH24:MI') AS "productionStartTime",
             to_char(finish_time, 'HH24:MI') AS "finishTime",
             units_produced AS "unitsProduced",
             COALESCE(run_crewing_pattern, '{}'::jsonb) AS "runCrewingPattern",
             COALESCE(u.name, u.username, '') AS "submittedBy",
             submitted_at AS "submittedAt"
           FROM run_logs
           LEFT JOIN users u ON u.id = run_logs.submitted_by_user_id
           WHERE line_id = ANY($1::UUID[])
           ORDER BY line_id, date ASC, shift ASC, submitted_at ASC`
        : `SELECT
             line_id AS "lineId",
             run_logs.id,
             date::TEXT AS date,
             shift,
             product,
             COALESCE(to_char(setup_start_time, 'HH24:MI'), '') AS "setUpStartTime",
             to_char(production_start_time, 'HH24:MI') AS "productionStartTime",
             to_char(finish_time, 'HH24:MI') AS "finishTime",
             units_produced AS "unitsProduced",
             COALESCE(run_crewing_pattern, '{}'::jsonb) AS "runCrewingPattern",
             COALESCE(u.name, u.username, '') AS "submittedBy",
             submitted_at AS "submittedAt"
           FROM run_logs
           LEFT JOIN users u ON u.id = run_logs.submitted_by_user_id
           WHERE line_id = ANY($2::UUID[])
             AND EXISTS (
               SELECT 1
               FROM supervisor_line_shift_assignments a
               WHERE a.supervisor_user_id = $1
                 AND a.line_id = run_logs.line_id
                 AND (
                   a.shift = run_logs.shift
                   OR (run_logs.shift = 'Full Day' AND a.shift IN ('Day', 'Night'))
                 )
             )
           ORDER BY line_id, date ASC, shift ASC, submitted_at ASC`,
      snapshotScopedParams
    ),
    dbQuery(
      isManager
        ? `SELECT
             line_id AS "lineId",
             downtime_logs.id,
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
           ORDER BY line_id, date ASC, shift ASC, submitted_at ASC`
        : `SELECT
             line_id AS "lineId",
             downtime_logs.id,
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
           WHERE line_id = ANY($2::UUID[])
             AND EXISTS (
               SELECT 1
               FROM supervisor_line_shift_assignments a
               WHERE a.supervisor_user_id = $1
                 AND a.line_id = downtime_logs.line_id
                 AND (
                   a.shift = downtime_logs.shift
                   OR (downtime_logs.shift = 'Full Day' AND a.shift IN ('Day', 'Night'))
                 )
             )
           ORDER BY line_id, date ASC, shift ASC, submitted_at ASC`,
      snapshotScopedParams
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
      id: row.id,
      date: row.date,
      shift: row.shift,
      crewOnShift: row.crewOnShift,
      startTime: row.startTime,
      finishTime: row.finishTime,
      submittedBy: row.submittedBy,
      submittedAt: row.submittedAt
    });
    shiftByLine.set(row.lineId, list);
  });

  const breakByLine = new Map();
  breakLogsResult.rows.forEach((row) => {
    const list = breakByLine.get(row.lineId) || [];
    list.push({
      id: row.id,
      shiftLogId: row.shiftLogId,
      date: row.date,
      shift: row.shift,
      breakStart: row.breakStart,
      breakFinish: row.breakFinish,
      submittedBy: row.submittedBy,
      submittedAt: row.submittedAt
    });
    breakByLine.set(row.lineId, list);
  });

  const runByLine = new Map();
  runLogsResult.rows.forEach((row) => {
    const list = runByLine.get(row.lineId) || [];
    list.push({
      id: row.id,
      date: row.date,
      shift: row.shift,
      product: row.product,
      setUpStartTime: row.setUpStartTime,
      productionStartTime: row.productionStartTime,
      finishTime: row.finishTime,
      unitsProduced: row.unitsProduced,
      runCrewingPattern: row.runCrewingPattern || {},
      submittedBy: row.submittedBy,
      submittedAt: row.submittedAt
    });
    runByLine.set(row.lineId, list);
  });

  const downtimeByLine = new Map();
  downtimeLogsResult.rows.forEach((row) => {
    const list = downtimeByLine.get(row.lineId) || [];
    list.push({
      id: row.id,
      date: row.date,
      shift: row.shift,
      downtimeStart: row.downtimeStart,
      downtimeFinish: row.downtimeFinish,
      equipment: row.equipment,
      reason: row.reason,
      submittedBy: row.submittedBy,
      submittedAt: row.submittedAt
    });
    downtimeByLine.set(row.lineId, list);
  });

  const lines = lineRows.map((line) => ({
    line,
    stages: stagesByLine.get(line.id) || [],
    guides: guidesByLine.get(line.id) || [],
    shiftRows: shiftByLine.get(line.id) || [],
    breakRows: breakByLine.get(line.id) || [],
    runRows: runByLine.get(line.id) || [],
    downtimeRows: downtimeByLine.get(line.id) || []
  }));

  const payload = { lines, supervisors: [] };
  if (req.user.role === 'manager') {
    payload.supervisors = await fetchSupervisorsWithAssignments();
  } else if (req.user.role === 'supervisor') {
    payload.supervisorAssignments = await fetchSupervisorAssignments(req.user.id);
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
  res.json({ supervisors: await fetchSupervisorsWithAssignments() });
}));

app.post('/api/supervisors', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  const parsed = supervisorCreateSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: 'Invalid supervisor payload' });
  const username = parsed.data.username.trim().toLowerCase();
  const name = parsed.data.name.trim();
  const passwordHash = await hashPassword(parsed.data.password);
  const assignedLineShifts = normalizeAssignedLineShifts(parsed.data.assignedLineShifts, parsed.data.assignedLineIds || []);
  const assignedLineIds = Object.keys(assignedLineShifts);
  const assignedPairs = flattenAssignedLineShifts(assignedLineShifts);
  const missingLineIds = await findMissingActiveLineIds(assignedLineIds);
  if (missingLineIds.length) {
    return res.status(400).json({ error: 'One or more assigned lines are invalid or inactive', lineIds: missingLineIds });
  }

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

    if (assignedPairs.length) {
      for (const pair of assignedPairs) {
        await client.query(
          `INSERT INTO supervisor_line_shift_assignments(supervisor_user_id, line_id, shift, assigned_by_user_id)
           VALUES ($1, $2, $3, $4)
           ON CONFLICT (supervisor_user_id, line_id, shift) DO NOTHING`,
          [supervisor.id, pair.lineId, pair.shift, req.user.id]
        );
      }
    }

    await client.query('COMMIT');
    return res.status(201).json({
      supervisor: {
        ...supervisor,
        assignedLineIds,
        assignedLineShifts
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
  const assignedLineShifts = normalizeAssignedLineShifts(parsed.data.assignedLineShifts, parsed.data.assignedLineIds || []);
  const assignedPairs = flattenAssignedLineShifts(assignedLineShifts);
  const missingLineIds = await findMissingActiveLineIds(Object.keys(assignedLineShifts));
  if (missingLineIds.length) {
    return res.status(400).json({ error: 'One or more assigned lines are invalid or inactive', lineIds: missingLineIds });
  }

  const userCheck = await dbQuery(
    `SELECT id FROM users WHERE id = $1 AND role = 'supervisor'`,
    [supervisorId]
  );
  if (!userCheck.rowCount) return res.status(404).json({ error: 'Supervisor not found' });

  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    await client.query(
      `DELETE FROM supervisor_line_shift_assignments WHERE supervisor_user_id = $1`,
      [supervisorId]
    );
    if (assignedPairs.length) {
      for (const pair of assignedPairs) {
        await client.query(
          `INSERT INTO supervisor_line_shift_assignments(supervisor_user_id, line_id, shift, assigned_by_user_id)
           VALUES ($1, $2, $3, $4)
           ON CONFLICT (supervisor_user_id, line_id, shift) DO NOTHING`,
          [supervisorId, pair.lineId, pair.shift, req.user.id]
        );
      }
    }
    await client.query('COMMIT');
    return res.json({ ok: true, assignedLineIds: Object.keys(assignedLineShifts), assignedLineShifts });
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
      `DELETE FROM supervisor_line_shift_assignments WHERE supervisor_user_id = $1`,
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
    req.user.role === 'manager'
      ? `SELECT id, name, secret_key AS "secretKey", created_at AS "createdAt", updated_at AS "updatedAt"
         FROM production_lines
         WHERE id = $1 AND is_active = TRUE`
      : `SELECT id, name, NULL::TEXT AS "secretKey", created_at AS "createdAt", updated_at AS "updatedAt"
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

  const isManager = req.user.role === 'manager';
  const lineLogsParams = isManager ? [lineId] : [req.user.id, lineId];
  const [shiftRows, breakRows, runRows, downtimeRows] = await Promise.all([
    dbQuery(
      isManager
        ? `SELECT
             shift_logs.id,
             date::TEXT AS date,
             shift,
             crew_on_shift AS "crewOnShift",
             to_char(start_time, 'HH24:MI') AS "startTime",
             to_char(finish_time, 'HH24:MI') AS "finishTime",
             COALESCE(u.name, u.username, '') AS "submittedBy",
             submitted_at AS "submittedAt"
           FROM shift_logs
           LEFT JOIN users u ON u.id = shift_logs.submitted_by_user_id
           WHERE line_id = $1
           ORDER BY date ASC, shift ASC, submitted_at ASC`
        : `SELECT
             shift_logs.id,
             date::TEXT AS date,
             shift,
             crew_on_shift AS "crewOnShift",
             to_char(start_time, 'HH24:MI') AS "startTime",
             to_char(finish_time, 'HH24:MI') AS "finishTime",
             COALESCE(u.name, u.username, '') AS "submittedBy",
             submitted_at AS "submittedAt"
           FROM shift_logs
           LEFT JOIN users u ON u.id = shift_logs.submitted_by_user_id
           WHERE line_id = $2
             AND EXISTS (
               SELECT 1
               FROM supervisor_line_shift_assignments a
               WHERE a.supervisor_user_id = $1
                 AND a.line_id = shift_logs.line_id
                 AND (
                   a.shift = shift_logs.shift
                   OR (shift_logs.shift = 'Full Day' AND a.shift IN ('Day', 'Night'))
                 )
             )
           ORDER BY date ASC, shift ASC, submitted_at ASC`,
      lineLogsParams
    ),
    dbQuery(
      isManager
        ? `SELECT
             shift_break_logs.id,
             shift_log_id AS "shiftLogId",
             date::TEXT AS date,
             shift,
             to_char(break_start, 'HH24:MI') AS "breakStart",
             COALESCE(to_char(break_finish, 'HH24:MI'), '') AS "breakFinish",
             COALESCE(u.name, u.username, '') AS "submittedBy",
             submitted_at AS "submittedAt"
           FROM shift_break_logs
           LEFT JOIN users u ON u.id = shift_break_logs.submitted_by_user_id
           WHERE line_id = $1
           ORDER BY date ASC, shift ASC, submitted_at ASC`
        : `SELECT
             shift_break_logs.id,
             shift_log_id AS "shiftLogId",
             date::TEXT AS date,
             shift,
             to_char(break_start, 'HH24:MI') AS "breakStart",
             COALESCE(to_char(break_finish, 'HH24:MI'), '') AS "breakFinish",
             COALESCE(u.name, u.username, '') AS "submittedBy",
             submitted_at AS "submittedAt"
           FROM shift_break_logs
           LEFT JOIN users u ON u.id = shift_break_logs.submitted_by_user_id
           WHERE line_id = $2
             AND EXISTS (
               SELECT 1
               FROM supervisor_line_shift_assignments a
               WHERE a.supervisor_user_id = $1
                 AND a.line_id = shift_break_logs.line_id
                 AND (
                   a.shift = shift_break_logs.shift
                   OR (shift_break_logs.shift = 'Full Day' AND a.shift IN ('Day', 'Night'))
                 )
             )
           ORDER BY date ASC, shift ASC, submitted_at ASC`,
      lineLogsParams
    ),
    dbQuery(
      isManager
        ? `SELECT
             run_logs.id,
             date::TEXT AS date,
             shift,
             product,
             COALESCE(to_char(setup_start_time, 'HH24:MI'), '') AS "setUpStartTime",
             to_char(production_start_time, 'HH24:MI') AS "productionStartTime",
             to_char(finish_time, 'HH24:MI') AS "finishTime",
             units_produced AS "unitsProduced",
             COALESCE(run_crewing_pattern, '{}'::jsonb) AS "runCrewingPattern",
             COALESCE(u.name, u.username, '') AS "submittedBy",
             submitted_at AS "submittedAt"
           FROM run_logs
           LEFT JOIN users u ON u.id = run_logs.submitted_by_user_id
           WHERE line_id = $1
           ORDER BY date ASC, shift ASC, submitted_at ASC`
        : `SELECT
             run_logs.id,
             date::TEXT AS date,
             shift,
             product,
             COALESCE(to_char(setup_start_time, 'HH24:MI'), '') AS "setUpStartTime",
             to_char(production_start_time, 'HH24:MI') AS "productionStartTime",
             to_char(finish_time, 'HH24:MI') AS "finishTime",
             units_produced AS "unitsProduced",
             COALESCE(run_crewing_pattern, '{}'::jsonb) AS "runCrewingPattern",
             COALESCE(u.name, u.username, '') AS "submittedBy",
             submitted_at AS "submittedAt"
           FROM run_logs
           LEFT JOIN users u ON u.id = run_logs.submitted_by_user_id
           WHERE line_id = $2
             AND EXISTS (
               SELECT 1
               FROM supervisor_line_shift_assignments a
               WHERE a.supervisor_user_id = $1
                 AND a.line_id = run_logs.line_id
                 AND (
                   a.shift = run_logs.shift
                   OR (run_logs.shift = 'Full Day' AND a.shift IN ('Day', 'Night'))
                 )
             )
           ORDER BY date ASC, shift ASC, submitted_at ASC`,
      lineLogsParams
    ),
    dbQuery(
      isManager
        ? `SELECT
             downtime_logs.id,
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
           ORDER BY date ASC, shift ASC, submitted_at ASC`
        : `SELECT
             downtime_logs.id,
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
           WHERE line_id = $2
             AND EXISTS (
               SELECT 1
               FROM supervisor_line_shift_assignments a
               WHERE a.supervisor_user_id = $1
                 AND a.line_id = downtime_logs.line_id
                 AND (
                   a.shift = downtime_logs.shift
                   OR (downtime_logs.shift = 'Full Day' AND a.shift IN ('Day', 'Night'))
                 )
             )
           ORDER BY date ASC, shift ASC, submitted_at ASC`,
      lineLogsParams
    )
  ]);

  return res.json({
    shiftRows: shiftRows.rows,
    breakRows: breakRows.rows,
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
  await dbQuery(`DELETE FROM supervisor_line_shift_assignments WHERE line_id = $1`, [lineId]);
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
  if (parsed.data.secretKey !== line.secretKey) {
    return res.status(403).json({ error: 'Invalid key/password. Data was not cleared.' });
  }

  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    await client.query(`DELETE FROM shift_break_logs WHERE line_id = $1`, [lineId]);
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
  if (!(await isActiveLine(parsed.data.lineId))) return res.status(404).json({ error: 'Line not found' });
  if (!(await hasLineShiftAccess(req.user, parsed.data.lineId, parsed.data.shift))) return res.status(403).json({ error: 'Forbidden' });

  const result = await dbQuery(
    `INSERT INTO shift_logs(line_id, date, shift, crew_on_shift, start_time, finish_time, submitted_by_user_id)
     VALUES ($1, $2, $3, $4, $5, $6, $7)
     RETURNING id, line_id AS "lineId", date, shift, crew_on_shift AS "crewOnShift", start_time AS "startTime", finish_time AS "finishTime", submitted_at AS "submittedAt"`,
    [
      parsed.data.lineId,
      parsed.data.date,
      parsed.data.shift,
      parsed.data.crewOnShift,
      parsed.data.startTime,
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

app.post('/api/logs/shifts/:logId/breaks', authMiddleware, asyncRoute(async (req, res) => {
  const logId = req.params.logId;
  if (!z.string().uuid().safeParse(logId).success) return res.status(400).json({ error: 'Invalid shift log id' });
  const parsed = shiftBreakStartSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid break start payload' });

  const shiftResult = await dbQuery(
    `SELECT
       id,
       line_id AS "lineId",
       date::TEXT AS date,
       shift,
       (start_time = finish_time) AS "isOpen",
       submitted_by_user_id AS "submittedByUserId"
     FROM shift_logs
     WHERE id = $1`,
    [logId]
  );
  if (!shiftResult.rowCount) return res.status(404).json({ error: 'Shift log not found' });
  const shiftLog = shiftResult.rows[0];

  if (!(await hasLineShiftAccess(req.user, shiftLog.lineId, shiftLog.shift))) return res.status(403).json({ error: 'Forbidden' });
  if (req.user.role !== 'manager' && shiftLog.submittedByUserId && shiftLog.submittedByUserId !== req.user.id) {
    return res.status(403).json({ error: 'Cannot update another user\'s shift log' });
  }
  if (!shiftLog.isOpen) return res.status(400).json({ error: 'Shift is already complete' });

  const openBreakResult = await dbQuery(
    `SELECT id
     FROM shift_break_logs
     WHERE shift_log_id = $1
       AND break_finish IS NULL
     LIMIT 1`,
    [logId]
  );
  if (openBreakResult.rowCount) return res.status(400).json({ error: 'An open break already exists for this shift' });

  const result = await dbQuery(
    `INSERT INTO shift_break_logs(shift_log_id, line_id, date, shift, break_start, submitted_by_user_id)
     VALUES ($1, $2, $3, $4, $5, $6)
     RETURNING
       id,
       shift_log_id AS "shiftLogId",
       line_id AS "lineId",
       date::TEXT AS date,
       shift,
       to_char(break_start, 'HH24:MI') AS "breakStart",
       COALESCE(to_char(break_finish, 'HH24:MI'), '') AS "breakFinish",
       submitted_at AS "submittedAt"`,
    [logId, shiftLog.lineId, shiftLog.date, shiftLog.shift, parsed.data.breakStart, req.user.id]
  );

  await writeAudit({
    lineId: shiftLog.lineId,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'START_SHIFT_BREAK',
    details: `${shiftLog.shift} ${shiftLog.date} break started ${parsed.data.breakStart}`
  });

  return res.status(201).json({ breakLog: result.rows[0] });
}));

app.patch('/api/logs/shifts/:logId/breaks/:breakId', authMiddleware, asyncRoute(async (req, res) => {
  const logId = req.params.logId;
  const breakId = req.params.breakId;
  if (!z.string().uuid().safeParse(logId).success) return res.status(400).json({ error: 'Invalid shift log id' });
  if (!z.string().uuid().safeParse(breakId).success) return res.status(400).json({ error: 'Invalid break log id' });
  const parsed = shiftBreakUpdateSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid break update payload' });

  const breakResult = await dbQuery(
    `SELECT
       b.id,
       b.shift_log_id AS "shiftLogId",
       b.line_id AS "lineId",
       b.date::TEXT AS date,
       b.shift,
       b.break_start AS "breakStartRaw",
       b.break_finish AS "breakFinishRaw",
       b.submitted_by_user_id AS "submittedByUserId",
       (s.start_time = s.finish_time) AS "shiftOpen"
     FROM shift_break_logs b
     INNER JOIN shift_logs s ON s.id = b.shift_log_id
     WHERE b.id = $1
       AND b.shift_log_id = $2`,
    [breakId, logId]
  );
  if (!breakResult.rowCount) return res.status(404).json({ error: 'Break log not found' });
  const breakLog = breakResult.rows[0];

  if (!(await hasLineShiftAccess(req.user, breakLog.lineId, breakLog.shift))) return res.status(403).json({ error: 'Forbidden' });
  if (req.user.role !== 'manager' && breakLog.submittedByUserId && breakLog.submittedByUserId !== req.user.id) {
    return res.status(403).json({ error: 'Cannot update another user\'s break log' });
  }
  if (!breakLog.shiftOpen) return res.status(400).json({ error: 'Shift is already complete' });

  if (parsed.data.breakFinish === undefined && parsed.data.breakStart === undefined) {
    return res.status(400).json({ error: 'At least one break field must be provided' });
  }

  const result = await dbQuery(
    `UPDATE shift_break_logs
     SET
       break_start = COALESCE($2::time, break_start),
       break_finish = COALESCE($3::time, break_finish),
       submitted_by_user_id = $4,
       submitted_at = NOW()
     WHERE id = $1
     RETURNING
       id,
       shift_log_id AS "shiftLogId",
       line_id AS "lineId",
       date::TEXT AS date,
       shift,
       to_char(break_start, 'HH24:MI') AS "breakStart",
       COALESCE(to_char(break_finish, 'HH24:MI'), '') AS "breakFinish",
       submitted_at AS "submittedAt"`,
    [breakId, parsed.data.breakStart || null, parsed.data.breakFinish || null, req.user.id]
  );

  const changedFields = [];
  if (parsed.data.breakStart !== undefined) changedFields.push(`start ${parsed.data.breakStart}`);
  if (parsed.data.breakFinish !== undefined) changedFields.push(`finish ${parsed.data.breakFinish}`);
  const isSimpleEnd = parsed.data.breakStart === undefined && parsed.data.breakFinish !== undefined && !breakLog.breakFinishRaw;

  await writeAudit({
    lineId: breakLog.lineId,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: isSimpleEnd ? 'END_SHIFT_BREAK' : 'UPDATE_SHIFT_BREAK',
    details: `${breakLog.shift} ${breakLog.date} break updated (${changedFields.join(', ')})`
  });

  return res.json({ breakLog: result.rows[0] });
}));

app.post('/api/logs/runs', authMiddleware, asyncRoute(async (req, res) => {
  const parsed = runLogSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: 'Invalid run log payload' });
  if (!(await isActiveLine(parsed.data.lineId))) return res.status(404).json({ error: 'Line not found' });
  if (!(await hasLineShiftAccess(req.user, parsed.data.lineId, parsed.data.shift))) return res.status(403).json({ error: 'Forbidden' });

  const result = await dbQuery(
    `INSERT INTO run_logs(
       line_id, date, shift, product, setup_start_time, production_start_time, finish_time, units_produced, run_crewing_pattern, submitted_by_user_id
     )
     VALUES ($1, $2, $3, $4, NULLIF($5, '')::time, $6, $7, $8, $9::jsonb, $10)
     RETURNING
       id,
       line_id AS "lineId",
       date::TEXT AS date,
       shift,
       product,
       COALESCE(to_char(setup_start_time, 'HH24:MI'), '') AS "setUpStartTime",
       to_char(production_start_time, 'HH24:MI') AS "productionStartTime",
       to_char(finish_time, 'HH24:MI') AS "finishTime",
       units_produced AS "unitsProduced",
       COALESCE(run_crewing_pattern, '{}'::jsonb) AS "runCrewingPattern",
       submitted_at AS "submittedAt"`,
    [
      parsed.data.lineId,
      parsed.data.date,
      parsed.data.shift,
      parsed.data.product,
      parsed.data.setUpStartTime || '',
      parsed.data.productionStartTime,
      parsed.data.finishTime,
      parsed.data.unitsProduced,
      JSON.stringify(parsed.data.runCrewingPattern || {}),
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
  if (!(await isActiveLine(parsed.data.lineId))) return res.status(404).json({ error: 'Line not found' });
  if (!(await hasLineShiftAccess(req.user, parsed.data.lineId, parsed.data.shift))) return res.status(403).json({ error: 'Forbidden' });
  if (parsed.data.equipmentStageId && !(await isLineStage(parsed.data.lineId, parsed.data.equipmentStageId))) {
    return res.status(400).json({ error: 'Equipment stage does not belong to this line' });
  }

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

app.patch('/api/logs/shifts/:logId', authMiddleware, asyncRoute(async (req, res) => {
  const logId = req.params.logId;
  if (!z.string().uuid().safeParse(logId).success) return res.status(400).json({ error: 'Invalid shift log id' });
  const parsed = shiftLogUpdateSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid shift log update payload' });

  const existingResult = await dbQuery(
    `SELECT id, line_id AS "lineId", shift, submitted_by_user_id AS "submittedByUserId"
     FROM shift_logs
     WHERE id = $1`,
    [logId]
  );
  if (!existingResult.rowCount) return res.status(404).json({ error: 'Shift log not found' });
  const existing = existingResult.rows[0];

  if (!(await hasLineShiftAccess(req.user, existing.lineId, existing.shift))) return res.status(403).json({ error: 'Forbidden' });
  if (req.user.role !== 'manager' && existing.submittedByUserId && existing.submittedByUserId !== req.user.id) {
    return res.status(403).json({ error: 'Cannot edit logs submitted by another user' });
  }

  const data = parsed.data;
  const result = await dbQuery(
    `UPDATE shift_logs
     SET
       crew_on_shift = COALESCE($2, crew_on_shift),
       start_time = COALESCE(NULLIF($3, '')::time, start_time),
       finish_time = COALESCE(NULLIF($4, '')::time, finish_time),
       submitted_by_user_id = $5,
       submitted_at = NOW()
     WHERE id = $1
     RETURNING
       id,
       line_id AS "lineId",
       date::TEXT AS date,
       shift,
       crew_on_shift AS "crewOnShift",
       to_char(start_time, 'HH24:MI') AS "startTime",
       to_char(finish_time, 'HH24:MI') AS "finishTime",
       submitted_at AS "submittedAt"`,
    [
      logId,
      data.crewOnShift,
      data.startTime ?? null,
      data.finishTime ?? null,
      req.user.id
    ]
  );

  await writeAudit({
    lineId: existing.lineId,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'UPDATE_SHIFT_LOG',
    details: `Shift log ${logId} updated`
  });

  return res.json({ shiftLog: result.rows[0] });
}));

app.patch('/api/logs/runs/:logId', authMiddleware, asyncRoute(async (req, res) => {
  const logId = req.params.logId;
  if (!z.string().uuid().safeParse(logId).success) return res.status(400).json({ error: 'Invalid run log id' });
  const parsed = runLogUpdateSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid run log update payload' });

  const existingResult = await dbQuery(
    `SELECT id, line_id AS "lineId", shift, submitted_by_user_id AS "submittedByUserId"
     FROM run_logs
     WHERE id = $1`,
    [logId]
  );
  if (!existingResult.rowCount) return res.status(404).json({ error: 'Run log not found' });
  const existing = existingResult.rows[0];

  if (!(await hasLineShiftAccess(req.user, existing.lineId, existing.shift))) return res.status(403).json({ error: 'Forbidden' });
  if (req.user.role !== 'manager' && existing.submittedByUserId && existing.submittedByUserId !== req.user.id) {
    return res.status(403).json({ error: 'Cannot edit logs submitted by another user' });
  }

  const data = parsed.data;
  const result = await dbQuery(
    `UPDATE run_logs
     SET
       product = COALESCE(NULLIF($2, ''), product),
       setup_start_time = CASE WHEN $3::text IS NULL THEN setup_start_time ELSE NULLIF($3, '')::time END,
       production_start_time = COALESCE(NULLIF($4, '')::time, production_start_time),
       finish_time = COALESCE(NULLIF($5, '')::time, finish_time),
       units_produced = COALESCE($6, units_produced),
       run_crewing_pattern = COALESCE($7::jsonb, run_crewing_pattern),
       submitted_by_user_id = $8,
       submitted_at = NOW()
     WHERE id = $1
     RETURNING
       id,
       line_id AS "lineId",
       date::TEXT AS date,
       shift,
       product,
       COALESCE(to_char(setup_start_time, 'HH24:MI'), '') AS "setUpStartTime",
       to_char(production_start_time, 'HH24:MI') AS "productionStartTime",
       to_char(finish_time, 'HH24:MI') AS "finishTime",
       units_produced AS "unitsProduced",
       COALESCE(run_crewing_pattern, '{}'::jsonb) AS "runCrewingPattern",
       submitted_at AS "submittedAt"`,
    [
      logId,
      data.product ?? null,
      data.setUpStartTime ?? null,
      data.productionStartTime ?? null,
      data.finishTime ?? null,
      data.unitsProduced,
      data.runCrewingPattern === undefined ? null : JSON.stringify(data.runCrewingPattern || {}),
      req.user.id
    ]
  );

  await writeAudit({
    lineId: existing.lineId,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'UPDATE_RUN_LOG',
    details: `Run log ${logId} updated`
  });

  return res.json({ runLog: result.rows[0] });
}));

app.patch('/api/logs/downtime/:logId', authMiddleware, asyncRoute(async (req, res) => {
  const logId = req.params.logId;
  if (!z.string().uuid().safeParse(logId).success) return res.status(400).json({ error: 'Invalid downtime log id' });
  const parsed = downtimeLogUpdateSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid downtime log update payload' });

  const existingResult = await dbQuery(
    `SELECT id, line_id AS "lineId", shift, submitted_by_user_id AS "submittedByUserId"
     FROM downtime_logs
     WHERE id = $1`,
    [logId]
  );
  if (!existingResult.rowCount) return res.status(404).json({ error: 'Downtime log not found' });
  const existing = existingResult.rows[0];

  if (!(await hasLineShiftAccess(req.user, existing.lineId, existing.shift))) return res.status(403).json({ error: 'Forbidden' });
  if (req.user.role !== 'manager' && existing.submittedByUserId && existing.submittedByUserId !== req.user.id) {
    return res.status(403).json({ error: 'Cannot edit logs submitted by another user' });
  }

  const data = parsed.data;
  const equipmentStageProvided = Object.prototype.hasOwnProperty.call(data, 'equipmentStageId');
  const equipmentStageValue = equipmentStageProvided ? String(data.equipmentStageId || '').trim() : '';
  if (equipmentStageValue && !(await isLineStage(existing.lineId, equipmentStageValue))) {
    return res.status(400).json({ error: 'Equipment stage does not belong to this line' });
  }
  const reasonProvided = Object.prototype.hasOwnProperty.call(data, 'reason');
  const reasonValue = reasonProvided ? String(data.reason ?? '').trim() : '';
  const result = await dbQuery(
    `UPDATE downtime_logs
     SET
       downtime_start = COALESCE(NULLIF($2, '')::time, downtime_start),
       downtime_finish = COALESCE(NULLIF($3, '')::time, downtime_finish),
       equipment_stage_id = CASE
         WHEN $4::boolean IS FALSE THEN equipment_stage_id
         WHEN NULLIF($5, '')::uuid IS NULL THEN NULL
         ELSE NULLIF($5, '')::uuid
       END,
       reason = CASE
         WHEN $6::boolean IS FALSE THEN reason
         ELSE NULLIF($7, '')
       END,
       submitted_by_user_id = $8,
       submitted_at = NOW()
     WHERE id = $1
     RETURNING
       id,
       line_id AS "lineId",
       date::TEXT AS date,
       shift,
       to_char(downtime_start, 'HH24:MI') AS "downtimeStart",
       to_char(downtime_finish, 'HH24:MI') AS "downtimeFinish",
       COALESCE(equipment_stage_id::TEXT, '') AS equipment,
       COALESCE(reason, '') AS reason,
       submitted_at AS "submittedAt"`,
    [
      logId,
      data.downtimeStart ?? null,
      data.downtimeFinish ?? null,
      equipmentStageProvided,
      equipmentStageValue,
      reasonProvided,
      reasonValue,
      req.user.id
    ]
  );

  await writeAudit({
    lineId: existing.lineId,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'UPDATE_DOWNTIME_LOG',
    details: `Downtime log ${logId} updated`
  });

  return res.json({ downtimeLog: result.rows[0] });
}));

function isTransientDatabaseError(error) {
  const code = String(error?.code || '').toUpperCase();
  const message = String(error?.message || '').toLowerCase();
  if (code.startsWith('08')) return true;
  if (['57P01', '57P02', '57P03', '53300', '53400'].includes(code)) return true;
  if (['ETIMEDOUT', 'ECONNRESET', 'EPIPE', 'ECONNREFUSED', 'ENOTFOUND', 'EAI_AGAIN'].includes(code)) return true;
  return message.includes('could not connect')
    || message.includes('connection terminated unexpectedly')
    || message.includes('timeout');
}

app.use((error, _req, res, _next) => {
  console.error(error);
  if (res.headersSent) return;
  if (isTransientDatabaseError(error)) {
    return res.status(503).json({ error: 'Database temporarily unavailable. Please retry shortly.' });
  }
  if (error?.code === '23503') {
    return res.status(400).json({ error: 'Invalid related resource reference.' });
  }
  if (error?.code === '23505') {
    return res.status(409).json({ error: 'Duplicate value violates unique constraint.' });
  }
  return res.status(500).json({ error: 'Internal server error' });
});

const server = app.listen(config.port, () => {
  console.log(`API listening on http://localhost:${config.port}`);
  console.log(`Database provider: ${config.databaseProvider}`);
});

server.requestTimeout = config.httpRequestTimeoutMs;
server.keepAliveTimeout = config.httpKeepAliveTimeoutMs;
server.headersTimeout = Math.max(config.httpHeadersTimeoutMs, config.httpKeepAliveTimeoutMs + 1000);
server.on('error', (error) => {
  console.error('[server] HTTP server error:', error);
});

let shuttingDown = false;
async function shutdown() {
  if (shuttingDown) return;
  shuttingDown = true;
  console.log('[server] Shutting down gracefully...');
  const forceExitTimer = setTimeout(() => {
    console.error('[server] Graceful shutdown timed out; forcing exit.');
    process.exit(1);
  }, 10000);
  forceExitTimer.unref();

  server.close(async () => {
    try {
      await pool.end();
    } catch (error) {
      console.error('[server] Error while closing DB pool:', error);
    } finally {
      clearTimeout(forceExitTimer);
      process.exit(0);
    }
  });
}

process.on('uncaughtException', (error) => {
  console.error('[process] Uncaught exception:', error);
});

process.on('unhandledRejection', (reason) => {
  console.error('[process] Unhandled rejection:', reason);
});

process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);
