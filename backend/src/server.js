import express from 'express';
import cors from 'cors';
import { Client } from 'pg';
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

function normalizeCorsOrigin(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  try {
    return new URL(raw).origin;
  } catch {
    return raw.replace(/\/+$/, '');
  }
}

const allowAllOrigins = (config.frontendOrigins || []).includes('*');
const allowedOrigins = new Set((config.frontendOrigins || []).map((origin) => normalizeCorsOrigin(origin)).filter(Boolean));

app.use(cors({
  origin(origin, callback) {
    if (!origin || allowAllOrigins) return callback(null, true);
    const normalizedOrigin = normalizeCorsOrigin(origin);
    return callback(null, allowedOrigins.has(normalizedOrigin));
  }
}));
app.use(express.json({ limit: '1mb' }));

const shiftValues = ['Day', 'Night', 'Full Day'];
const supervisorShiftValues = ['Day', 'Night'];
const actionPriorityValues = ['Low', 'Medium', 'High', 'Critical'];
const actionStatusValues = ['Open', 'In Progress', 'Blocked', 'Completed'];
const productCatalogColumnCount = 12;
const productCatalogAllLinesToken = '*';
const timeRegex = /^([01]\d|2[0-3]):([0-5]\d)$/;
const isoDateRegex = /^\d{4}-\d{2}-\d{2}$/;
const optionalTime = z.string().regex(timeRegex).optional().or(z.literal('')).or(z.null());
const optionalLogNotes = z.string().max(2000).optional().or(z.literal('')).or(z.null());
const optionalIsoDate = z.string().regex(isoDateRegex).optional().or(z.literal('')).or(z.null());

const loginSchema = z.object({
  username: z.string().min(1),
  password: z.string().min(1)
});

const lineSchema = z.object({
  name: z.string().min(2).max(120),
  secretKey: z.string().min(4).max(64)
});

const lineGroupCreateSchema = z.object({
  name: z.string().trim().min(1).max(120)
});

const lineGroupUpdateSchema = z.object({
  name: z.string().trim().min(1).max(120)
});

const lineGroupReorderSchema = z.object({
  groupIds: z.array(z.string().uuid()).min(1)
});

const dataSourceCreateSchema = z.object({
  sourceName: z.string().trim().min(1).max(160),
  sourceKey: z.string().trim().min(1).max(160).optional().or(z.literal('')),
  provider: z.enum(['sql', 'api']).optional().default('api'),
  connectionMode: z.enum(['sql', 'api']).optional().default('api'),
  machineNo: z.string().max(64).optional().or(z.literal('')).or(z.null()),
  deviceName: z.string().max(160).optional().or(z.literal('')).or(z.null()),
  deviceId: z.string().max(160).optional().or(z.literal('')).or(z.null()),
  scaleNumber: z.string().max(64).optional().or(z.literal('')).or(z.null()),
  apiBaseUrl: z.string().max(300).optional().or(z.literal('')).or(z.null()),
  apiKey: z.string().max(300).optional().or(z.literal('')).or(z.null()),
  apiSecret: z.string().max(300).optional().or(z.literal('')).or(z.null()),
  sqlHost: z.string().max(255).optional().or(z.literal('')).or(z.null()),
  sqlPort: z.number().int().min(1).max(65535).optional().nullable(),
  sqlDatabase: z.string().max(160).optional().or(z.literal('')).or(z.null()),
  sqlUsername: z.string().max(160).optional().or(z.literal('')).or(z.null()),
  sqlPassword: z.string().max(300).optional().or(z.literal('')).or(z.null())
});

const dataSourceConnectionTestSchema = z.object({
  provider: z.enum(['sql', 'api']).optional().default('api'),
  connectionMode: z.enum(['sql', 'api']).optional().default('api'),
  apiBaseUrl: z.string().max(300).optional().or(z.literal('')).or(z.null()),
  apiKey: z.string().max(300).optional().or(z.literal('')).or(z.null()),
  apiSecret: z.string().max(300).optional().or(z.literal('')).or(z.null()),
  sqlHost: z.string().max(255).optional().or(z.literal('')).or(z.null()),
  sqlPort: z.number().int().min(1).max(65535).optional().nullable(),
  sqlDatabase: z.string().max(160).optional().or(z.literal('')).or(z.null()),
  sqlUsername: z.string().max(160).optional().or(z.literal('')).or(z.null()),
  sqlPassword: z.string().max(300).optional().or(z.literal('')).or(z.null())
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
      dataSourceId: z.string().uuid().optional().or(z.literal('')).or(z.null()),
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
  date: z.string().regex(isoDateRegex),
  shift: z.enum(shiftValues),
  crewOnShift: z.number().int().min(0),
  startTime: z.string().regex(timeRegex),
  finishTime: z.string().regex(timeRegex),
  notes: optionalLogNotes
});

const runLogSchema = z.object({
  lineId: z.string().uuid(),
  date: z.string().regex(isoDateRegex),
  shift: z.enum(shiftValues),
  product: z.string().min(1).max(120),
  setUpStartTime: optionalTime,
  productionStartTime: z.string().regex(timeRegex),
  finishTime: optionalTime,
  unitsProduced: z.number().nonnegative(),
  notes: optionalLogNotes,
  runCrewingPattern: z.record(z.number().int().min(0)).optional().default({})
});

const downtimeLogSchema = z.object({
  lineId: z.string().uuid(),
  date: z.string().regex(isoDateRegex),
  shift: z.enum(shiftValues),
  downtimeStart: z.string().regex(timeRegex),
  downtimeFinish: z.string().regex(timeRegex),
  equipmentStageId: z.string().uuid().nullable().optional(),
  reason: z.string().max(250).optional().or(z.literal('')).or(z.null()),
  notes: optionalLogNotes
});

const supervisorActionCreateSchema = z.object({
  supervisorUsername: z.string().min(1).max(60).optional(),
  supervisorName: z.string().max(120).optional().or(z.literal('')).or(z.null()),
  lineId: z.string().uuid().optional().or(z.literal('')).or(z.null()),
  title: z.string().trim().min(1).max(180),
  description: z.string().max(4000).optional().or(z.literal('')).or(z.null()),
  priority: z.enum(actionPriorityValues).optional().default('Medium'),
  status: z.enum(actionStatusValues).optional().default('Open'),
  dueDate: optionalIsoDate,
  relatedEquipmentId: z.string().uuid().optional().or(z.literal('')).or(z.null()),
  relatedReasonCategory: z.string().max(120).optional().or(z.literal('')).or(z.null()),
  relatedReasonDetail: z.string().max(240).optional().or(z.literal('')).or(z.null())
});

const hasAtLeastOneField = (obj) => Object.values(obj).some((value) => value !== undefined);

const lineUpdateSchema = z.object({
  name: z.string().trim().min(2).max(120).optional(),
  groupId: z.string().uuid().optional().or(z.literal('')).or(z.null())
}).refine(hasAtLeastOneField, { message: 'At least one field is required' });

const lineReorderSchema = z.object({
  lineIds: z.array(z.string().uuid()).min(1),
  groupId: z.string().uuid().optional().or(z.literal('')).or(z.null())
});

const LINE_ORDER_BY_SQL = `
  CASE WHEN l.group_id IS NULL THEN 1 ELSE 0 END ASC,
  COALESCE(g.display_order, 2147483647) ASC,
  l.display_order ASC,
  LOWER(l.name) ASC,
  l.created_at ASC
`;

const shiftLogUpdateSchema = z.object({
  date: z.string().regex(isoDateRegex).optional(),
  shift: z.enum(shiftValues).optional(),
  crewOnShift: z.number().int().min(0).optional(),
  startTime: optionalTime.optional(),
  finishTime: optionalTime.optional(),
  notes: optionalLogNotes
}).refine(hasAtLeastOneField, { message: 'At least one field is required' });

const shiftBreakStartSchema = z.object({
  breakStart: z.string().regex(timeRegex),
  breakFinish: optionalTime.optional()
});

const shiftBreakUpdateSchema = z.object({
  breakStart: z.string().regex(timeRegex).optional(),
  breakFinish: z.string().regex(timeRegex).optional()
}).refine(hasAtLeastOneField, { message: 'At least one field is required' });

const runLogUpdateSchema = z.object({
  date: z.string().regex(isoDateRegex).optional(),
  shift: z.enum(shiftValues).optional(),
  product: z.string().min(1).max(120).optional(),
  setUpStartTime: optionalTime.optional(),
  productionStartTime: optionalTime.optional(),
  finishTime: optionalTime.optional(),
  unitsProduced: z.number().nonnegative().optional(),
  notes: optionalLogNotes,
  runCrewingPattern: z.record(z.number().int().min(0)).optional()
}).refine(hasAtLeastOneField, { message: 'At least one field is required' });

const downtimeLogUpdateSchema = z.object({
  date: z.string().regex(isoDateRegex).optional(),
  shift: z.enum(shiftValues).optional(),
  downtimeStart: optionalTime.optional(),
  downtimeFinish: optionalTime.optional(),
  equipmentStageId: z.string().uuid().optional().or(z.literal('')).or(z.null()),
  reason: z.string().max(250).optional().or(z.literal('')).or(z.null()),
  notes: optionalLogNotes
}).refine(hasAtLeastOneField, { message: 'At least one field is required' });

const supervisorActionUpdateSchema = z.object({
  supervisorUsername: z.string().min(1).max(60).optional(),
  supervisorName: z.string().max(120).optional().or(z.literal('')).or(z.null()),
  lineId: z.string().uuid().optional().or(z.literal('')).or(z.null()),
  title: z.string().trim().min(1).max(180).optional(),
  description: z.string().max(4000).optional().or(z.literal('')).or(z.null()),
  priority: z.enum(actionPriorityValues).optional(),
  status: z.enum(actionStatusValues).optional(),
  dueDate: optionalIsoDate,
  relatedEquipmentId: z.string().uuid().optional().or(z.literal('')).or(z.null()),
  relatedReasonCategory: z.string().max(120).optional().or(z.literal('')).or(z.null()),
  relatedReasonDetail: z.string().max(240).optional().or(z.literal('')).or(z.null())
}).refine(hasAtLeastOneField, { message: 'At least one field is required' });

const productCatalogValuesSchema = z.array(z.string()).length(productCatalogColumnCount);
const productCatalogLineIdSchema = z.union([z.string().uuid(), z.literal(productCatalogAllLinesToken)]);

const productCatalogEntrySchema = z.object({
  values: productCatalogValuesSchema,
  lineIds: z.array(productCatalogLineIdSchema).optional().default([productCatalogAllLinesToken])
});

const clearLineDataSchema = z.object({
  secretKey: z.string().min(1).max(128)
});

const loadPermanentSampleDataSchema = z.object({
  replaceExisting: z.boolean().optional().default(true)
});

const asyncRoute = (handler) => (req, res, next) => Promise.resolve(handler(req, res, next)).catch(next);
const DATA_SOURCE_CONNECTION_TEST_TIMEOUT_MS = 7000;

function normalizeDataSourceKey(sourceName, sourceKey = '') {
  const base = String(sourceKey || sourceName || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 160);
  if (base) return base;
  return `source-${Date.now()}`;
}

function optionalText(value) {
  const trimmed = String(value || '').trim();
  return trimmed || null;
}

function normalizeDataSourceConnectionPayload(data = {}) {
  const provider = data.provider === 'sql' ? 'sql' : 'api';
  const connectionMode = data.connectionMode === 'sql' ? 'sql' : 'api';
  const apiBaseUrl = optionalText(data.apiBaseUrl);
  const apiKey = connectionMode === 'api' ? optionalText(data.apiKey) : null;
  const apiSecret = connectionMode === 'api' ? optionalText(data.apiSecret) : null;
  const sqlHost = connectionMode === 'sql' ? optionalText(data.sqlHost) : null;
  const sqlPort = connectionMode === 'sql' && Number.isFinite(Number(data.sqlPort))
    ? Math.max(1, Math.min(65535, Math.floor(Number(data.sqlPort))))
    : null;
  const sqlDatabase = connectionMode === 'sql' ? optionalText(data.sqlDatabase) : null;
  const sqlUsername = connectionMode === 'sql' ? optionalText(data.sqlUsername) : null;
  const sqlPassword = connectionMode === 'sql' ? optionalText(data.sqlPassword) : null;
  return {
    provider,
    connectionMode,
    apiBaseUrl,
    apiKey,
    apiSecret,
    sqlHost,
    sqlPort,
    sqlDatabase,
    sqlUsername,
    sqlPassword
  };
}

function dataSourceConnectionErrorMessage(error) {
  const code = String(error?.code || '').toUpperCase();
  if (code === 'ENOTFOUND') return 'Host could not be resolved.';
  if (code === 'ECONNREFUSED') return 'Connection was refused by the remote host.';
  if (code === 'ECONNRESET') return 'Connection was reset by the remote host.';
  if (code === 'ETIMEDOUT') return 'Connection timed out.';
  if (code === '28P01') return 'Authentication failed. Check username and password.';
  if (code === '3D000') return 'Database does not exist or is not accessible.';
  const message = String(error?.message || '').trim();
  if (!message) return 'Connection failed.';
  const compact = message.replace(/\s+/g, ' ');
  return compact.length > 220 ? `${compact.slice(0, 217)}...` : compact;
}

async function testApiDataSourceConnection(payload = {}) {
  const apiBaseUrl = String(payload.apiBaseUrl || '').trim();
  const apiKey = String(payload.apiKey || '').trim();
  const apiSecret = String(payload.apiSecret || '').trim();
  if (!apiBaseUrl) {
    return { ok: false, mode: 'api', message: 'API Base URL is required for API connection tests.' };
  }
  if (!apiKey) {
    return { ok: false, mode: 'api', message: 'API key is required for API connection tests.' };
  }

  let targetUrl = '';
  try {
    targetUrl = new URL(apiBaseUrl).toString();
  } catch {
    return { ok: false, mode: 'api', message: 'API Base URL must be a valid absolute URL.' };
  }

  const controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
  const timeoutId = controller ? setTimeout(() => controller.abort(), DATA_SOURCE_CONNECTION_TEST_TIMEOUT_MS) : null;
  try {
    const headers = {
      Authorization: `Bearer ${apiKey}`,
      Accept: 'application/json,text/plain;q=0.9,*/*;q=0.8'
    };
    if (apiSecret) headers['X-API-Secret'] = apiSecret;
    const response = await fetch(targetUrl, {
      method: 'GET',
      headers,
      signal: controller ? controller.signal : undefined
    });
    if (response.ok) {
      return {
        ok: true,
        mode: 'api',
        statusCode: response.status,
        message: `API responded with HTTP ${response.status}.`
      };
    }
    const responseText = String(await response.text()).trim().replace(/\s+/g, ' ');
    const suffix = responseText ? ` ${responseText.slice(0, 140)}` : '';
    return {
      ok: false,
      mode: 'api',
      statusCode: response.status,
      message: `API returned HTTP ${response.status}.${suffix}`
    };
  } catch (error) {
    if (error?.name === 'AbortError') {
      return {
        ok: false,
        mode: 'api',
        message: `Connection test timed out after ${Math.round(DATA_SOURCE_CONNECTION_TEST_TIMEOUT_MS / 1000)} seconds.`
      };
    }
    return {
      ok: false,
      mode: 'api',
      message: dataSourceConnectionErrorMessage(error)
    };
  } finally {
    if (timeoutId) clearTimeout(timeoutId);
  }
}

async function testSqlDataSourceConnection(payload = {}) {
  const sqlHost = String(payload.sqlHost || '').trim();
  const sqlDatabase = String(payload.sqlDatabase || '').trim();
  const sqlUsername = String(payload.sqlUsername || '').trim();
  const sqlPassword = String(payload.sqlPassword || '').trim();
  const sqlPortRaw = Number(payload.sqlPort);
  const sqlPort = Number.isFinite(sqlPortRaw) && sqlPortRaw > 0 ? Math.floor(sqlPortRaw) : 5432;
  if (!sqlHost || !sqlDatabase || !sqlUsername || !sqlPassword) {
    return {
      ok: false,
      mode: 'sql',
      message: 'SQL host, database, username and password are required for SQL connection tests.'
    };
  }

  const client = new Client({
    host: sqlHost,
    port: sqlPort,
    database: sqlDatabase,
    user: sqlUsername,
    password: sqlPassword,
    application_name: 'kebab-line-data-source-test',
    connectionTimeoutMillis: DATA_SOURCE_CONNECTION_TEST_TIMEOUT_MS,
    query_timeout: DATA_SOURCE_CONNECTION_TEST_TIMEOUT_MS,
    statement_timeout: DATA_SOURCE_CONNECTION_TEST_TIMEOUT_MS
  });

  try {
    await client.connect();
    await client.query('SELECT 1');
    return {
      ok: true,
      mode: 'sql',
      message: `SQL connection successful (${sqlHost}:${sqlPort}/${sqlDatabase}).`
    };
  } catch (error) {
    return {
      ok: false,
      mode: 'sql',
      message: dataSourceConnectionErrorMessage(error)
    };
  } finally {
    try {
      await client.end();
    } catch {
      // Ignore close errors; the probe result already captures the primary failure.
    }
  }
}

async function testDataSourceConnection(payload = {}) {
  const connectionMode = String(payload.connectionMode || '').trim().toLowerCase() === 'sql' ? 'sql' : 'api';
  return connectionMode === 'sql'
    ? testSqlDataSourceConnection(payload)
    : testApiDataSourceConnection(payload);
}

const SAMPLE_DOWNTIME_REASON_PRESETS = {
  'Donor Meat': ['Stock Out', 'Late Delivery', 'Quality Hold', 'Temperature Hold'],
  People: ['Understaffed', 'Training', 'Handover Delay', 'Absence'],
  Materials: ['Film Shortage', 'Label Shortage', 'Tray Shortage', 'Marinade Shortage', 'Skewer shortage'],
  Other: ['Cleaning', 'QA Hold', 'Power', 'Unplanned Stop']
};

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
         AND u.is_active = TRUE
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

async function fetchLineGroups() {
  const result = await dbQuery(
    `SELECT
       id,
       name,
       display_order AS "displayOrder",
       created_at AS "createdAt",
       updated_at AS "updatedAt"
     FROM line_groups
     WHERE is_active = TRUE
     ORDER BY display_order ASC, LOWER(name) ASC, created_at ASC`
  );
  return result.rows;
}

async function fetchDataSources() {
  const result = await dbQuery(
    `SELECT
       id,
       source_key AS "sourceKey",
       source_name AS "sourceName",
       machine_no AS "machineNo",
       device_name AS "deviceName",
       device_id AS "deviceId",
       scale_number AS "scaleNumber",
       provider,
       connection_mode AS "connectionMode",
       COALESCE(api_base_url, '') AS "apiBaseUrl",
       (COALESCE(api_key, '') <> '') AS "hasApiKey",
       (
         COALESCE(sql_host, '') <> ''
         AND COALESCE(sql_database, '') <> ''
         AND COALESCE(sql_username, '') <> ''
         AND COALESCE(sql_password, '') <> ''
       ) AS "hasSqlCredentials",
       is_active AS "isActive",
       created_at AS "createdAt",
       updated_at AS "updatedAt"
     FROM data_sources
     WHERE is_active = TRUE
     ORDER BY LOWER(source_name) ASC, created_at ASC`
  );
  return result.rows;
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
  if (user.role === 'manager') return true;
  const result = await dbQuery(
    `SELECT 1
     FROM supervisor_line_shift_assignments a
     INNER JOIN production_lines l
       ON l.id = a.line_id
     WHERE a.supervisor_user_id = $1
       AND a.line_id = $2
       AND l.is_active = TRUE
       AND (
         a.shift = $3
         OR (
           $3 = 'Full Day'
           AND EXISTS (
             SELECT 1
             FROM supervisor_line_shift_assignments day_access
             WHERE day_access.supervisor_user_id = $1
                       AND day_access.line_id = $2
               AND day_access.shift = 'Day'
           )
           AND EXISTS (
             SELECT 1
             FROM supervisor_line_shift_assignments night_access
             WHERE night_access.supervisor_user_id = $1
                       AND night_access.line_id = $2
               AND night_access.shift = 'Night'
           )
         )
       )
     LIMIT 1`,
    [user.id, lineId, shift]
  );
  return result.rowCount > 0;
}

function canMutateSubmittedLog(user, submittedByUserId) {
  if (!user) return false;
  if (user.role === 'manager') return true;
  const ownerId = String(submittedByUserId || '').trim();
  const actorId = String(user.id || '').trim();
  return Boolean(ownerId) && ownerId === actorId;
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

async function findMissingActiveLineGroupIds(groupIds) {
  const uniqueGroupIds = Array.from(
    new Set((Array.isArray(groupIds) ? groupIds : []).filter((groupId) => z.string().uuid().safeParse(groupId).success))
  );
  if (!uniqueGroupIds.length) return [];
  const result = await dbQuery(
    `SELECT id::TEXT AS id
     FROM line_groups
     WHERE is_active = TRUE
       AND id = ANY($1::UUID[])`,
    [uniqueGroupIds]
  );
  const foundIds = new Set(result.rows.map((row) => row.id));
  return uniqueGroupIds.filter((groupId) => !foundIds.has(groupId));
}

async function writeAudit({ lineId = null, actorUserId = null, actorName = null, actorRole = null, action, details = '' }) {
  await dbQuery(
    `INSERT INTO audit_events(line_id, actor_user_id, actor_name, actor_role, action, details)
     VALUES ($1, $2, $3, $4, $5, $6)`,
    [lineId, actorUserId, actorName, actorRole, action, details]
  );
}

function normalizeIsoDate(value) {
  const trimmed = String(value || '').trim();
  return isoDateRegex.test(trimmed) ? trimmed : '';
}

function normalizeActionReasonCategory(value) {
  const trimmed = String(value || '').trim();
  return trimmed || '';
}

function normalizeActionReasonDetail(value, category = '') {
  const trimmed = String(value || '').trim();
  if (!category) return '';
  return trimmed;
}

function normalizeActionAssigneeUsername(value) {
  return String(value || '').trim().toLowerCase();
}

function mapSupervisorActionRow(row = {}) {
  return {
    id: row.id,
    supervisorUsername: normalizeActionAssigneeUsername(row.supervisorUsername),
    supervisorName: String(row.supervisorName || '').trim(),
    lineId: String(row.lineId || '').trim(),
    title: String(row.title || '').trim(),
    description: String(row.description || '').trim(),
    priority: actionPriorityValues.includes(String(row.priority || '').trim()) ? String(row.priority).trim() : 'Medium',
    status: actionStatusValues.includes(String(row.status || '').trim()) ? String(row.status).trim() : 'Open',
    dueDate: normalizeIsoDate(row.dueDate),
    relatedEquipmentId: String(row.relatedEquipmentId || '').trim(),
    relatedReasonCategory: normalizeActionReasonCategory(row.relatedReasonCategory),
    relatedReasonDetail: normalizeActionReasonDetail(row.relatedReasonDetail, row.relatedReasonCategory),
    createdAt: row.createdAt,
    createdBy: String(row.createdBy || '').trim()
  };
}

function normalizeProductCatalogValues(values = []) {
  const source = Array.isArray(values) ? values : [];
  return Array.from({ length: productCatalogColumnCount }, (_unused, index) => String(source[index] ?? '').trim());
}

function normalizeProductCatalogLineSelection(lineIds = []) {
  const source = Array.isArray(lineIds) ? lineIds : [];
  const normalized = [];
  let allLines = false;
  source.forEach((rawLineId) => {
    const lineId = String(rawLineId || '').trim();
    if (!lineId) return;
    if (lineId === productCatalogAllLinesToken) {
      allLines = true;
      return;
    }
    if (!z.string().uuid().safeParse(lineId).success) return;
    if (!normalized.includes(lineId)) normalized.push(lineId);
  });
  if (allLines) return { allLines: true, lineIds: [] };
  return { allLines: false, lineIds: normalized };
}

function hasProductCatalogIdentity(values = []) {
  const safeValues = normalizeProductCatalogValues(values);
  return Boolean(safeValues[0] || safeValues[1] || safeValues[2]);
}

function mapProductCatalogRow(row = {}) {
  const id = String(row.id || '').trim();
  if (!z.string().uuid().safeParse(id).success) return null;
  const values = normalizeProductCatalogValues(Array.isArray(row.values) ? row.values : []);
  const lineIds = Array.isArray(row.lineIds)
    ? row.lineIds
      .map((lineId) => String(lineId || '').trim())
      .filter((lineId) => z.string().uuid().safeParse(lineId).success)
    : [];
  if (!hasProductCatalogIdentity(values)) return null;
  return {
    id,
    values,
    lineIds: row.allLines ? [productCatalogAllLinesToken] : lineIds
  };
}

async function fetchProductCatalogForUser(user) {
  const isManager = user?.role === 'manager';
  const result = await dbQuery(
    isManager
      ? `SELECT
           p.id,
           COALESCE(p.catalog_values, '[]'::jsonb) AS values,
           p.all_lines AS "allLines",
           ARRAY(
             SELECT line_id::TEXT
             FROM unnest(COALESCE(p.line_ids, ARRAY[]::UUID[])) AS line_id
           ) AS "lineIds"
         FROM product_catalog_entries p
         ORDER BY p.created_at ASC, p.id ASC`
      : `SELECT
           p.id,
           COALESCE(p.catalog_values, '[]'::jsonb) AS values,
           p.all_lines AS "allLines",
           ARRAY(
             SELECT line_id::TEXT
             FROM unnest(COALESCE(p.line_ids, ARRAY[]::UUID[])) AS line_id
           ) AS "lineIds"
         FROM product_catalog_entries p
         WHERE p.all_lines = TRUE
            OR EXISTS (
              SELECT 1
              FROM unnest(COALESCE(p.line_ids, ARRAY[]::UUID[])) AS product_line_id
              INNER JOIN supervisor_line_shift_assignments a
                ON a.line_id = product_line_id
               AND a.supervisor_user_id = $1
            )
         ORDER BY p.created_at ASC, p.id ASC`,
    isManager ? [] : [user.id]
  );
  return result.rows
    .map((row) => mapProductCatalogRow(row))
    .filter(Boolean);
}

async function resolveSupervisorActionAssignee(usernameInput, fallbackName = '') {
  const requestedUsername = normalizeActionAssigneeUsername(usernameInput);
  if (!requestedUsername) return null;
  const supervisorResult = await dbQuery(
    `SELECT id, username, name
     FROM users
     WHERE role = 'supervisor'
       AND is_active = TRUE
       AND LOWER(username) = $1
     LIMIT 1`,
    [requestedUsername]
  );
  if (!supervisorResult.rowCount) {
    return {
      supervisorUserId: null,
      supervisorUsername: requestedUsername,
      supervisorName: String(fallbackName || requestedUsername).trim() || requestedUsername
    };
  }
  const row = supervisorResult.rows[0];
  const username = normalizeActionAssigneeUsername(row.username) || requestedUsername;
  const supervisorName = String(row.name || username).trim() || username;
  return {
    supervisorUserId: row.id,
    supervisorUsername: username,
    supervisorName
  };
}

async function fetchSupervisorActionsForUser(user) {
  const isManager = user?.role === 'manager';
  const result = await dbQuery(
    isManager
      ? `SELECT
           sa.id,
           sa.supervisor_username AS "supervisorUsername",
           sa.supervisor_name AS "supervisorName",
           COALESCE(sa.line_id::TEXT, '') AS "lineId",
           sa.title,
           COALESCE(sa.description, '') AS description,
           sa.priority,
           sa.status,
           COALESCE(sa.due_date::TEXT, '') AS "dueDate",
           COALESCE(sa.related_equipment_stage_id::TEXT, '') AS "relatedEquipmentId",
           COALESCE(sa.related_reason_category, '') AS "relatedReasonCategory",
           COALESCE(sa.related_reason_detail, '') AS "relatedReasonDetail",
           sa.created_at AS "createdAt",
           COALESCE(sa.created_by_name, '') AS "createdBy"
         FROM supervisor_actions sa
         ORDER BY sa.created_at DESC, sa.id DESC`
      : `SELECT
           sa.id,
           sa.supervisor_username AS "supervisorUsername",
           sa.supervisor_name AS "supervisorName",
           COALESCE(sa.line_id::TEXT, '') AS "lineId",
           sa.title,
           COALESCE(sa.description, '') AS description,
           sa.priority,
           sa.status,
           COALESCE(sa.due_date::TEXT, '') AS "dueDate",
           COALESCE(sa.related_equipment_stage_id::TEXT, '') AS "relatedEquipmentId",
           COALESCE(sa.related_reason_category, '') AS "relatedReasonCategory",
           COALESCE(sa.related_reason_detail, '') AS "relatedReasonDetail",
           sa.created_at AS "createdAt",
           COALESCE(sa.created_by_name, '') AS "createdBy"
         FROM supervisor_actions sa
         WHERE (
             sa.supervisor_user_id = $1
             OR LOWER(sa.supervisor_username) = $2
           )
         ORDER BY sa.created_at DESC, sa.id DESC`,
    isManager ? [] : [user.id, normalizeActionAssigneeUsername(user?.username || '')]
  );
  return result.rows.map((row) => mapSupervisorActionRow(row)).filter((row) => row && row.supervisorUsername);
}

function formatDateUtc(date) {
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date.getUTCDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function buildDowntimeReasonText(category, detail, note = '') {
  const group = String(category || '').trim();
  const detailText = String(detail || '').trim();
  const noteText = String(note || '').trim();
  if (!group) return noteText;
  if (!detailText && !noteText) return group;
  if (!noteText) return `${group} > ${detailText}`;
  if (!detailText) return `${group} > ${noteText}`;
  return `${group} > ${detailText} > ${noteText}`;
}

function pickFrom(list, index, offset = 0, fallback = '') {
  const source = Array.isArray(list) ? list : [];
  if (!source.length) return fallback;
  const position = ((index + offset) % source.length + source.length) % source.length;
  return source[position] || fallback;
}

function buildPermanentSampleData(stageRows = [], { days = 84 } = {}) {
  const shiftRows = [];
  const breakRows = [];
  const runRows = [];
  const downtimeRows = [];
  const safeDays = Math.max(1, Math.min(365, Number.isFinite(Number(days)) ? Math.floor(Number(days)) : 84));
  const start = new Date(Date.UTC(2025, 10, 1));

  const stages = Array.isArray(stageRows) ? stageRows : [];
  const stageIds = stages.map((stage) => stage.id).filter(Boolean);
  const stageNameById = new Map(stages.map((stage) => [stage.id, String(stage.stageName || 'Stage').trim() || 'Stage']));
  const dayEquipment = stageIds.filter((_, index) => index % 2 === 0);
  const nightEquipment = stageIds.filter((_, index) => index % 2 === 1);
  const fallbackEquipment = stageIds[0] || null;
  const dayRequired = Math.max(
    1,
    stages.reduce((sum, stage) => sum + Math.max(0, Number(stage.dayCrew) || 0), 0) || 14
  );
  const nightRequired = Math.max(
    1,
    stages.reduce((sum, stage) => sum + Math.max(0, Number(stage.nightCrew) || 0), 0) || 12
  );

  for (let i = 0; i < safeDays; i += 1) {
    const dateObj = new Date(start);
    dateObj.setUTCDate(start.getUTCDate() + i);
    const date = formatDateUtc(dateObj);
    const dayTrend = 0.95 + ((i % 7) - 3) * 0.01;
    const nightTrend = 0.92 + (((i + 2) % 7) - 3) * 0.01;

    shiftRows.push(
      {
        date,
        shift: 'Day',
        crewOnShift: Math.max(0, dayRequired - (i % 12 === 0 ? 2 : i % 7 === 0 ? 1 : 0)),
        startTime: '06:00',
        finishTime: '14:00'
      },
      {
        date,
        shift: 'Night',
        crewOnShift: Math.max(0, nightRequired - (i % 10 === 0 ? 1 : 0)),
        startTime: '14:00',
        finishTime: '22:00'
      }
    );

    breakRows.push(
      { date, shift: 'Day', breakStart: '09:00', breakFinish: '09:15' },
      { date, shift: 'Day', breakStart: '12:00', breakFinish: '12:30' },
      { date, shift: 'Day', breakStart: '13:40', breakFinish: '13:55' },
      { date, shift: 'Night', breakStart: '17:00', breakFinish: '17:15' },
      { date, shift: 'Night', breakStart: '20:00', breakFinish: '20:30' },
      { date, shift: 'Night', breakStart: '21:40', breakFinish: '21:55' }
    );

    runRows.push(
      {
        date,
        shift: 'Day',
        product: 'Teriyaki',
        setUpStartTime: null,
        productionStartTime: '06:10',
        finishTime: '10:35',
        unitsProduced: Math.round(2850 * dayTrend)
      },
      {
        date,
        shift: 'Day',
        product: 'Honey Soy',
        setUpStartTime: null,
        productionStartTime: '10:55',
        finishTime: '15:25',
        unitsProduced: Math.round(2600 * dayTrend)
      },
      {
        date,
        shift: 'Night',
        product: 'Peri Peri',
        setUpStartTime: null,
        productionStartTime: '14:15',
        finishTime: '18:55',
        unitsProduced: Math.round(2500 * nightTrend)
      },
      {
        date,
        shift: 'Night',
        product: 'Lemon Herb',
        setUpStartTime: null,
        productionStartTime: '21:30',
        finishTime: '00:25',
        unitsProduced: Math.round(2200 * nightTrend)
      }
    );

    const dayEquipmentId = pickFrom(dayEquipment, i, 0, fallbackEquipment);
    const dayEquipmentAltId = pickFrom(dayEquipment, i, 2, fallbackEquipment);
    const nightEquipmentId = pickFrom(nightEquipment, i, 0, fallbackEquipment);
    const nightEquipmentAltId = pickFrom(nightEquipment, i, 3, fallbackEquipment);
    const dayEquipmentName = stageNameById.get(dayEquipmentId) || 'Equipment';
    const dayEquipmentAltName = stageNameById.get(dayEquipmentAltId) || 'Equipment';
    const nightEquipmentName = stageNameById.get(nightEquipmentId) || 'Equipment';
    const nightEquipmentAltName = stageNameById.get(nightEquipmentAltId) || 'Equipment';

    const dayReasonCategory = i % 4 === 0 ? 'People' : 'Equipment';
    const dayReasonDetail = dayReasonCategory === 'Equipment'
      ? dayEquipmentName
      : SAMPLE_DOWNTIME_REASON_PRESETS.People[i % SAMPLE_DOWNTIME_REASON_PRESETS.People.length];
    const dayReasonNote = dayReasonCategory === 'Equipment' ? 'Planned maintenance' : '';

    const dayReasonCategory2 = i % 5 === 0 ? 'Materials' : 'Equipment';
    const dayReasonDetail2 = dayReasonCategory2 === 'Equipment'
      ? dayEquipmentAltName
      : SAMPLE_DOWNTIME_REASON_PRESETS.Materials[i % SAMPLE_DOWNTIME_REASON_PRESETS.Materials.length];
    const dayReasonNote2 = dayReasonCategory2 === 'Equipment' ? 'Minor stoppage' : '';

    const nightReasonCategory = i % 3 === 0 ? 'Donor Meat' : 'Equipment';
    const nightReasonDetail = nightReasonCategory === 'Equipment'
      ? nightEquipmentName
      : SAMPLE_DOWNTIME_REASON_PRESETS['Donor Meat'][i % SAMPLE_DOWNTIME_REASON_PRESETS['Donor Meat'].length];
    const nightReasonNote = nightReasonCategory === 'Equipment' ? 'Sensor reset' : '';

    const nightReasonCategory2 = i % 6 === 0 ? 'Other' : 'Equipment';
    const nightReasonDetail2 = nightReasonCategory2 === 'Equipment'
      ? nightEquipmentAltName
      : SAMPLE_DOWNTIME_REASON_PRESETS.Other[i % SAMPLE_DOWNTIME_REASON_PRESETS.Other.length];
    const nightReasonNote2 = nightReasonCategory2 === 'Equipment' ? 'Label adjustment' : '';

    downtimeRows.push(
      {
        date,
        shift: 'Day',
        downtimeStart: '08:10',
        downtimeFinish: `08:${String(22 + (i % 8)).padStart(2, '0')}`,
        equipmentStageId: dayReasonCategory === 'Equipment' ? dayEquipmentId : null,
        reason: buildDowntimeReasonText(dayReasonCategory, dayReasonDetail, dayReasonNote)
      },
      {
        date,
        shift: 'Day',
        downtimeStart: '11:20',
        downtimeFinish: `11:${String(30 + (i % 10)).padStart(2, '0')}`,
        equipmentStageId: dayReasonCategory2 === 'Equipment' ? dayEquipmentAltId : null,
        reason: buildDowntimeReasonText(dayReasonCategory2, dayReasonDetail2, dayReasonNote2)
      },
      {
        date,
        shift: 'Night',
        downtimeStart: '16:30',
        downtimeFinish: `16:${String(42 + (i % 9)).padStart(2, '0')}`,
        equipmentStageId: nightReasonCategory === 'Equipment' ? nightEquipmentId : null,
        reason: buildDowntimeReasonText(nightReasonCategory, nightReasonDetail, nightReasonNote)
      },
      {
        date,
        shift: 'Night',
        downtimeStart: '20:05',
        downtimeFinish: `20:${String(18 + (i % 11)).padStart(2, '0')}`,
        equipmentStageId: nightReasonCategory2 === 'Equipment' ? nightEquipmentAltId : null,
        reason: buildDowntimeReasonText(nightReasonCategory2, nightReasonDetail2, nightReasonNote2)
      }
    );
  }

  return { shiftRows, breakRows, runRows, downtimeRows };
}

let lineGroupSchemaReady = false;
let lineGroupSchemaPromise = null;

async function ensureLineGroupSchema() {
  if (lineGroupSchemaReady) return;
  if (lineGroupSchemaPromise) return lineGroupSchemaPromise;
  lineGroupSchemaPromise = (async () => {
    await dbQuery(
      `CREATE TABLE IF NOT EXISTS line_groups (
         id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
         name TEXT NOT NULL,
         display_order INTEGER NOT NULL DEFAULT 0,
         is_active BOOLEAN NOT NULL DEFAULT TRUE,
         created_by_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
         created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
         updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
       )`
    );
    await dbQuery(
      `ALTER TABLE production_lines
       ADD COLUMN IF NOT EXISTS group_id UUID REFERENCES line_groups(id) ON DELETE SET NULL`
    );
    await dbQuery(
      `ALTER TABLE production_lines
       ADD COLUMN IF NOT EXISTS display_order INTEGER NOT NULL DEFAULT 0`
    );
    await dbQuery(
      `CREATE UNIQUE INDEX IF NOT EXISTS idx_line_groups_name_active_unique
       ON line_groups (LOWER(name))
       WHERE is_active = TRUE`
    );
    await dbQuery(
      `CREATE INDEX IF NOT EXISTS idx_line_groups_display_order
       ON line_groups (display_order ASC, created_at ASC)`
    );
    await dbQuery(
      `CREATE INDEX IF NOT EXISTS idx_production_lines_group_id
       ON production_lines (group_id)`
    );
    await dbQuery(
      `CREATE INDEX IF NOT EXISTS idx_production_lines_group_display_order
       ON production_lines (group_id, display_order ASC, created_at ASC)
       WHERE is_active = TRUE`
    );
    lineGroupSchemaReady = true;
  })()
    .catch((error) => {
      lineGroupSchemaPromise = null;
      throw error;
    });
  return lineGroupSchemaPromise;
}

let dataSourceSchemaReady = false;
let dataSourceSchemaPromise = null;

async function ensureDataSourceSchema() {
  if (dataSourceSchemaReady) return;
  if (dataSourceSchemaPromise) return dataSourceSchemaPromise;
  dataSourceSchemaPromise = (async () => {
    await dbQuery(
      `CREATE TABLE IF NOT EXISTS data_sources (
         id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
         source_key TEXT NOT NULL UNIQUE,
         source_name TEXT NOT NULL,
         machine_no TEXT,
         device_name TEXT,
         device_id TEXT,
         scale_number TEXT,
         provider TEXT NOT NULL DEFAULT 'sql',
         connection_mode TEXT NOT NULL DEFAULT 'api',
         api_base_url TEXT,
         api_key TEXT,
         api_secret TEXT,
         sql_host TEXT,
         sql_port INTEGER,
         sql_database TEXT,
         sql_username TEXT,
         sql_password TEXT,
         is_active BOOLEAN NOT NULL DEFAULT TRUE,
         created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
         updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
       )`
    );
    await dbQuery(
      `ALTER TABLE data_sources
       ADD COLUMN IF NOT EXISTS connection_mode TEXT NOT NULL DEFAULT 'api'`
    );
    await dbQuery(
      `ALTER TABLE data_sources
       ADD COLUMN IF NOT EXISTS api_base_url TEXT`
    );
    await dbQuery(
      `ALTER TABLE data_sources
       ADD COLUMN IF NOT EXISTS api_key TEXT`
    );
    await dbQuery(
      `ALTER TABLE data_sources
       ADD COLUMN IF NOT EXISTS api_secret TEXT`
    );
    await dbQuery(
      `ALTER TABLE data_sources
       ADD COLUMN IF NOT EXISTS sql_host TEXT`
    );
    await dbQuery(
      `ALTER TABLE data_sources
       ADD COLUMN IF NOT EXISTS sql_port INTEGER`
    );
    await dbQuery(
      `ALTER TABLE data_sources
       ADD COLUMN IF NOT EXISTS sql_database TEXT`
    );
    await dbQuery(
      `ALTER TABLE data_sources
       ADD COLUMN IF NOT EXISTS sql_username TEXT`
    );
    await dbQuery(
      `ALTER TABLE data_sources
       ADD COLUMN IF NOT EXISTS sql_password TEXT`
    );
    await dbQuery(
      `ALTER TABLE line_stages
       ADD COLUMN IF NOT EXISTS data_source_id UUID REFERENCES data_sources(id) ON DELETE SET NULL`
    );
    await dbQuery(
      `CREATE UNIQUE INDEX IF NOT EXISTS idx_line_stages_data_source_unique
       ON line_stages(data_source_id)
       WHERE data_source_id IS NOT NULL`
    );
    await dbQuery(
      `CREATE INDEX IF NOT EXISTS idx_data_sources_active_name
       ON data_sources(is_active, LOWER(source_name), created_at)`
    );
    await dbQuery(
      `INSERT INTO data_sources(
         source_key,
         source_name,
         machine_no,
         device_name,
         device_id,
         scale_number,
         provider,
         connection_mode,
         is_active
       )
       VALUES
         ($1, $2, $3, $4, $5, $6, $7, $8, TRUE),
         ($9, $10, $11, $12, $13, $14, $15, $16, TRUE),
         ($17, $18, $19, $20, $21, $22, $23, $24, TRUE)
       ON CONFLICT (source_key) DO UPDATE
       SET
         source_name = EXCLUDED.source_name,
         machine_no = EXCLUDED.machine_no,
         device_name = EXCLUDED.device_name,
         device_id = EXCLUDED.device_id,
         scale_number = EXCLUDED.scale_number,
         provider = EXCLUDED.provider,
         connection_mode = EXCLUDED.connection_mode,
         is_active = EXCLUDED.is_active,
         updated_at = NOW()`,
      [
        'bizerba-proseal-line-1',
        'ProSeal Line 1 - MASTER',
        '1',
        'ProSeal Line 1 - MASTER',
        '{2081B032-ECC1-4350-88A4-51867EBDFBDE}',
        '1',
        'sql',
        'sql',
        'bizerba-multivac-line',
        'Multivac Line - MASTER',
        '3',
        'Multivac Line - MASTER',
        '{4341C2FD-EFBC-4b14-A0CC-3E59C1750D1E}',
        '1',
        'sql',
        'sql',
        'bizerba-ulma-line',
        'Ulma Line - MASTER',
        '4',
        'Ulma Line -  MASTER',
        '{EE6A8E2F-9068-4bda-9DAF-83C81A322FA0}',
        '1',
        'sql',
        'sql'
      ]
    );
    dataSourceSchemaReady = true;
  })()
    .catch((error) => {
      dataSourceSchemaPromise = null;
      throw error;
    });
  return dataSourceSchemaPromise;
}

let logSchemaReady = false;
let logSchemaPromise = null;

async function ensureLogSchema() {
  if (logSchemaReady) return;
  if (logSchemaPromise) return logSchemaPromise;
  logSchemaPromise = (async () => {
    await dbQuery(
      `ALTER TABLE shift_logs
       ADD COLUMN IF NOT EXISTS notes TEXT NOT NULL DEFAULT ''`
    );
    await dbQuery(
      `ALTER TABLE run_logs
       ADD COLUMN IF NOT EXISTS run_crewing_pattern JSONB NOT NULL DEFAULT '{}'::jsonb`
    );
    await dbQuery(
      `ALTER TABLE run_logs
       ADD COLUMN IF NOT EXISTS notes TEXT NOT NULL DEFAULT ''`
    );
    await dbQuery(
      `ALTER TABLE run_logs
       ALTER COLUMN finish_time DROP NOT NULL`
    );
    await dbQuery(
      `ALTER TABLE downtime_logs
       ADD COLUMN IF NOT EXISTS notes TEXT NOT NULL DEFAULT ''`
    );
    logSchemaReady = true;
  })()
    .catch((error) => {
      logSchemaPromise = null;
      throw error;
    });
  return logSchemaPromise;
}

let supervisorActionSchemaReady = false;
let supervisorActionSchemaPromise = null;

async function ensureSupervisorActionSchema() {
  if (supervisorActionSchemaReady) return;
  if (supervisorActionSchemaPromise) return supervisorActionSchemaPromise;
  supervisorActionSchemaPromise = (async () => {
    await dbQuery(
      `CREATE TABLE IF NOT EXISTS supervisor_actions (
         id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
         supervisor_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
         supervisor_username TEXT NOT NULL,
         supervisor_name TEXT NOT NULL DEFAULT '',
         line_id UUID REFERENCES production_lines(id) ON DELETE SET NULL,
         title TEXT NOT NULL,
         description TEXT NOT NULL DEFAULT '',
         priority TEXT NOT NULL DEFAULT 'Medium',
         status TEXT NOT NULL DEFAULT 'Open',
         due_date DATE,
         related_equipment_stage_id UUID REFERENCES line_stages(id) ON DELETE SET NULL,
         related_reason_category TEXT,
         related_reason_detail TEXT,
         created_by_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
         created_by_name TEXT NOT NULL DEFAULT '',
         created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
         updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
         CONSTRAINT supervisor_actions_priority_check CHECK (priority IN ('Low', 'Medium', 'High', 'Critical')),
         CONSTRAINT supervisor_actions_status_check CHECK (status IN ('Open', 'In Progress', 'Blocked', 'Completed'))
       )`
    );
    await dbQuery(
      `ALTER TABLE supervisor_actions
       ADD COLUMN IF NOT EXISTS supervisor_user_id UUID REFERENCES users(id) ON DELETE SET NULL`
    );
    await dbQuery(
      `ALTER TABLE supervisor_actions
       ADD COLUMN IF NOT EXISTS supervisor_username TEXT NOT NULL DEFAULT ''`
    );
    await dbQuery(
      `ALTER TABLE supervisor_actions
       ADD COLUMN IF NOT EXISTS supervisor_name TEXT NOT NULL DEFAULT ''`
    );
    await dbQuery(
      `ALTER TABLE supervisor_actions
       ADD COLUMN IF NOT EXISTS line_id UUID REFERENCES production_lines(id) ON DELETE SET NULL`
    );
    await dbQuery(
      `ALTER TABLE supervisor_actions
       ADD COLUMN IF NOT EXISTS title TEXT NOT NULL DEFAULT ''`
    );
    await dbQuery(
      `ALTER TABLE supervisor_actions
       ADD COLUMN IF NOT EXISTS description TEXT NOT NULL DEFAULT ''`
    );
    await dbQuery(
      `ALTER TABLE supervisor_actions
       ADD COLUMN IF NOT EXISTS priority TEXT NOT NULL DEFAULT 'Medium'`
    );
    await dbQuery(
      `ALTER TABLE supervisor_actions
       ADD COLUMN IF NOT EXISTS status TEXT NOT NULL DEFAULT 'Open'`
    );
    await dbQuery(
      `ALTER TABLE supervisor_actions
       ADD COLUMN IF NOT EXISTS due_date DATE`
    );
    await dbQuery(
      `ALTER TABLE supervisor_actions
       ADD COLUMN IF NOT EXISTS related_equipment_stage_id UUID REFERENCES line_stages(id) ON DELETE SET NULL`
    );
    await dbQuery(
      `ALTER TABLE supervisor_actions
       ADD COLUMN IF NOT EXISTS related_reason_category TEXT`
    );
    await dbQuery(
      `ALTER TABLE supervisor_actions
       ADD COLUMN IF NOT EXISTS related_reason_detail TEXT`
    );
    await dbQuery(
      `ALTER TABLE supervisor_actions
       ADD COLUMN IF NOT EXISTS created_by_user_id UUID REFERENCES users(id) ON DELETE SET NULL`
    );
    await dbQuery(
      `ALTER TABLE supervisor_actions
       ADD COLUMN IF NOT EXISTS created_by_name TEXT NOT NULL DEFAULT ''`
    );
    await dbQuery(
      `ALTER TABLE supervisor_actions
       ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()`
    );
    await dbQuery(
      `ALTER TABLE supervisor_actions
       ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()`
    );
    await dbQuery(
      `CREATE INDEX IF NOT EXISTS idx_supervisor_actions_assignee
       ON supervisor_actions (LOWER(supervisor_username), created_at DESC)`
    );
    await dbQuery(
      `CREATE INDEX IF NOT EXISTS idx_supervisor_actions_line
       ON supervisor_actions (line_id, created_at DESC)`
    );
    await dbQuery(
      `CREATE INDEX IF NOT EXISTS idx_supervisor_actions_due_date
       ON supervisor_actions (due_date ASC NULLS LAST, created_at DESC)`
    );
    supervisorActionSchemaReady = true;
  })()
    .catch((error) => {
      supervisorActionSchemaPromise = null;
      throw error;
    });
  return supervisorActionSchemaPromise;
}

let productCatalogSchemaReady = false;
let productCatalogSchemaPromise = null;

async function ensureProductCatalogSchema() {
  if (productCatalogSchemaReady) return;
  if (productCatalogSchemaPromise) return productCatalogSchemaPromise;
  productCatalogSchemaPromise = (async () => {
    await dbQuery(
      `CREATE TABLE IF NOT EXISTS product_catalog_entries (
         id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
         catalog_values JSONB NOT NULL DEFAULT '[]'::jsonb,
         line_ids UUID[] NOT NULL DEFAULT ARRAY[]::UUID[],
         all_lines BOOLEAN NOT NULL DEFAULT TRUE,
         created_by_user_id UUID REFERENCES users(id) ON DELETE SET NULL,
         created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
         updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
       )`
    );
    await dbQuery(
      `ALTER TABLE product_catalog_entries
       ADD COLUMN IF NOT EXISTS catalog_values JSONB NOT NULL DEFAULT '[]'::jsonb`
    );
    await dbQuery(
      `ALTER TABLE product_catalog_entries
       ADD COLUMN IF NOT EXISTS line_ids UUID[] NOT NULL DEFAULT ARRAY[]::UUID[]`
    );
    await dbQuery(
      `ALTER TABLE product_catalog_entries
       ADD COLUMN IF NOT EXISTS all_lines BOOLEAN NOT NULL DEFAULT TRUE`
    );
    await dbQuery(
      `ALTER TABLE product_catalog_entries
       ADD COLUMN IF NOT EXISTS created_by_user_id UUID REFERENCES users(id) ON DELETE SET NULL`
    );
    await dbQuery(
      `ALTER TABLE product_catalog_entries
       ADD COLUMN IF NOT EXISTS created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()`
    );
    await dbQuery(
      `ALTER TABLE product_catalog_entries
       ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()`
    );
    await dbQuery(
      `CREATE INDEX IF NOT EXISTS idx_product_catalog_entries_created
       ON product_catalog_entries (created_at ASC, id ASC)`
    );
    await dbQuery(
      `CREATE INDEX IF NOT EXISTS idx_product_catalog_entries_all_lines
       ON product_catalog_entries (all_lines, created_at ASC)`
    );
    await dbQuery(
      `CREATE INDEX IF NOT EXISTS idx_product_catalog_entries_line_ids
       ON product_catalog_entries
       USING GIN (line_ids)`
    );
    productCatalogSchemaReady = true;
  })()
    .catch((error) => {
      productCatalogSchemaPromise = null;
      throw error;
    });
  return productCatalogSchemaPromise;
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

app.get('/api/line-groups', authMiddleware, requireRole('manager'), asyncRoute(async (_req, res) => {
  await ensureLineGroupSchema();
  res.json({ lineGroups: await fetchLineGroups() });
}));

app.get('/api/data-sources', authMiddleware, requireRole('manager'), asyncRoute(async (_req, res) => {
  await ensureDataSourceSchema();
  res.json({ dataSources: await fetchDataSources() });
}));

app.post('/api/data-sources', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  await ensureDataSourceSchema();
  const parsed = dataSourceCreateSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid data source payload' });

  const sourceName = parsed.data.sourceName.trim();
  const sourceKey = normalizeDataSourceKey(sourceName, parsed.data.sourceKey || '');
  const machineNo = optionalText(parsed.data.machineNo);
  const deviceName = optionalText(parsed.data.deviceName);
  const deviceId = optionalText(parsed.data.deviceId);
  const scaleNumber = optionalText(parsed.data.scaleNumber);
  const {
    provider,
    connectionMode,
    apiBaseUrl,
    apiKey,
    apiSecret,
    sqlHost,
    sqlPort,
    sqlDatabase,
    sqlUsername,
    sqlPassword
  } = normalizeDataSourceConnectionPayload(parsed.data);

  if (connectionMode === 'api' && !apiKey) {
    return res.status(400).json({ error: 'API key is required for API connection mode.' });
  }
  if (connectionMode === 'sql' && (!sqlHost || !sqlDatabase || !sqlUsername || !sqlPassword)) {
    return res.status(400).json({ error: 'SQL host, database, username and password are required for SQL connection mode.' });
  }

  try {
    const inserted = await dbQuery(
      `INSERT INTO data_sources(
         source_key,
         source_name,
         machine_no,
         device_name,
         device_id,
         scale_number,
         provider,
         connection_mode,
         api_base_url,
         api_key,
         api_secret,
         sql_host,
         sql_port,
         sql_database,
         sql_username,
         sql_password,
         is_active,
         updated_at
       )
       VALUES (
         $1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,TRUE,NOW()
       )
       RETURNING id`,
      [
        sourceKey,
        sourceName,
        machineNo,
        deviceName,
        deviceId,
        scaleNumber,
        provider,
        connectionMode,
        apiBaseUrl,
        apiKey,
        apiSecret,
        sqlHost,
        sqlPort,
        sqlDatabase,
        sqlUsername,
        sqlPassword
      ]
    );
    const insertedId = inserted.rows?.[0]?.id;
    const result = await dbQuery(
      `SELECT
         id,
         source_key AS "sourceKey",
         source_name AS "sourceName",
         machine_no AS "machineNo",
         device_name AS "deviceName",
         device_id AS "deviceId",
         scale_number AS "scaleNumber",
         provider,
         connection_mode AS "connectionMode",
         COALESCE(api_base_url, '') AS "apiBaseUrl",
         (COALESCE(api_key, '') <> '') AS "hasApiKey",
         (
           COALESCE(sql_host, '') <> ''
           AND COALESCE(sql_database, '') <> ''
           AND COALESCE(sql_username, '') <> ''
           AND COALESCE(sql_password, '') <> ''
         ) AS "hasSqlCredentials",
         is_active AS "isActive",
         created_at AS "createdAt",
         updated_at AS "updatedAt"
       FROM data_sources
       WHERE id = $1`,
      [insertedId]
    );
    return res.status(201).json({ dataSource: result.rows[0] || null });
  } catch (error) {
    if (String(error?.message || '').includes('duplicate key')) {
      return res.status(409).json({ error: 'Data source key already exists. Please use a unique source key.' });
    }
    throw error;
  }
}));

app.post('/api/data-sources/test', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  await ensureDataSourceSchema();
  const parsed = dataSourceConnectionTestSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid data source test payload' });

  const payload = normalizeDataSourceConnectionPayload(parsed.data);
  const test = await testDataSourceConnection(payload);
  return res.json({ test });
}));

app.post('/api/data-sources/:dataSourceId/test', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  await ensureDataSourceSchema();
  const dataSourceId = String(req.params.dataSourceId || '').trim();
  if (!z.string().uuid().safeParse(dataSourceId).success) {
    return res.status(400).json({ error: 'Invalid data source id' });
  }
  const result = await dbQuery(
    `SELECT
       id::TEXT AS id,
       source_name AS "sourceName",
       provider,
       connection_mode AS "connectionMode",
       COALESCE(api_base_url, '') AS "apiBaseUrl",
       COALESCE(api_key, '') AS "apiKey",
       COALESCE(api_secret, '') AS "apiSecret",
       COALESCE(sql_host, '') AS "sqlHost",
       sql_port AS "sqlPort",
       COALESCE(sql_database, '') AS "sqlDatabase",
       COALESCE(sql_username, '') AS "sqlUsername",
       COALESCE(sql_password, '') AS "sqlPassword"
     FROM data_sources
     WHERE id = $1
       AND is_active = TRUE
     LIMIT 1`,
    [dataSourceId]
  );
  const source = result.rows?.[0];
  if (!source) return res.status(404).json({ error: 'Data source not found' });

  const payload = normalizeDataSourceConnectionPayload(source);
  const test = await testDataSourceConnection(payload);
  return res.json({
    test: {
      ...test,
      dataSourceId: source.id,
      sourceName: source.sourceName
    }
  });
}));

app.post('/api/line-groups', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  await ensureLineGroupSchema();
  const parsed = lineGroupCreateSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid line group payload' });

  const name = parsed.data.name.trim();
  try {
    const result = await dbQuery(
      `INSERT INTO line_groups(name, display_order, created_by_user_id)
       VALUES (
         $1,
         COALESCE((SELECT MAX(display_order) + 1 FROM line_groups WHERE is_active = TRUE), 0),
         $2
       )
       RETURNING
         id,
         name,
         display_order AS "displayOrder",
         created_at AS "createdAt",
         updated_at AS "updatedAt"`,
      [name, req.user.id]
    );
    return res.status(201).json({ lineGroup: result.rows[0] });
  } catch (error) {
    if (String(error?.message || '').includes('duplicate key')) {
      return res.status(409).json({ error: 'Line group name already exists' });
    }
    throw error;
  }
}));

app.patch('/api/line-groups/reorder', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  await ensureLineGroupSchema();
  const parsed = lineGroupReorderSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid line group reorder payload' });
  const nextGroupIds = parsed.data.groupIds;
  const uniqueGroupIds = Array.from(new Set(nextGroupIds));
  if (uniqueGroupIds.length !== nextGroupIds.length) {
    return res.status(400).json({ error: 'Duplicate group ids are not allowed' });
  }

  const activeGroupsResult = await dbQuery(
    `SELECT id::TEXT AS id
     FROM line_groups
     WHERE is_active = TRUE
     ORDER BY display_order ASC, LOWER(name) ASC, created_at ASC`
  );
  const activeGroupIds = activeGroupsResult.rows.map((row) => row.id);
  if (activeGroupIds.length !== uniqueGroupIds.length) {
    return res.status(400).json({ error: 'Reorder payload must include all active groups' });
  }
  const activeGroupSet = new Set(activeGroupIds);
  const hasInvalid = uniqueGroupIds.some((groupId) => !activeGroupSet.has(groupId));
  if (hasInvalid) return res.status(400).json({ error: 'One or more line group ids are invalid' });

  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    for (let index = 0; index < uniqueGroupIds.length; index += 1) {
      await client.query(
        `UPDATE line_groups
         SET display_order = $2, updated_at = NOW()
         WHERE id = $1 AND is_active = TRUE`,
        [uniqueGroupIds[index], index]
      );
    }
    await client.query('COMMIT');
    return res.json({ lineGroups: await fetchLineGroups() });
  } catch (error) {
    await client.query('ROLLBACK');
    throw error;
  } finally {
    client.release();
  }
}));

app.patch('/api/line-groups/:groupId', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  await ensureLineGroupSchema();
  const groupId = req.params.groupId;
  if (!z.string().uuid().safeParse(groupId).success) return res.status(400).json({ error: 'Invalid line group id' });
  const parsed = lineGroupUpdateSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid line group update payload' });

  try {
    const result = await dbQuery(
      `UPDATE line_groups
       SET name = $2, updated_at = NOW()
       WHERE id = $1 AND is_active = TRUE
       RETURNING
         id,
         name,
         display_order AS "displayOrder",
         created_at AS "createdAt",
         updated_at AS "updatedAt"`,
      [groupId, parsed.data.name.trim()]
    );
    if (!result.rowCount) return res.status(404).json({ error: 'Line group not found' });
    return res.json({ lineGroup: result.rows[0] });
  } catch (error) {
    if (String(error?.message || '').includes('duplicate key')) {
      return res.status(409).json({ error: 'Line group name already exists' });
    }
    throw error;
  }
}));

app.delete('/api/line-groups/:groupId', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  await ensureLineGroupSchema();
  const groupId = req.params.groupId;
  if (!z.string().uuid().safeParse(groupId).success) return res.status(400).json({ error: 'Invalid line group id' });

  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    const groupCheck = await client.query(
      `SELECT id FROM line_groups WHERE id = $1 AND is_active = TRUE`,
      [groupId]
    );
    if (!groupCheck.rowCount) {
      await client.query('ROLLBACK');
      return res.status(404).json({ error: 'Line group not found' });
    }
    await client.query(
      `WITH group_lines AS (
         SELECT
           id,
           ROW_NUMBER() OVER (ORDER BY display_order ASC, created_at ASC) - 1 AS row_idx
         FROM production_lines
         WHERE group_id = $1 AND is_active = TRUE
       ),
       ungrouped_base AS (
         SELECT COALESCE(MAX(display_order) + 1, 0) AS start_order
         FROM production_lines
         WHERE group_id IS NULL AND is_active = TRUE
       )
       UPDATE production_lines line
       SET
         group_id = NULL,
         display_order = (SELECT start_order FROM ungrouped_base) + group_lines.row_idx,
         updated_at = NOW()
       FROM group_lines
       WHERE line.id = group_lines.id`,
      [groupId]
    );
    await client.query(
      `UPDATE line_groups
       SET is_active = FALSE, updated_at = NOW()
       WHERE id = $1`,
      [groupId]
    );
    await client.query('COMMIT');
    return res.json({ ok: true });
  } catch (error) {
    await client.query('ROLLBACK');
    throw error;
  } finally {
    client.release();
  }
}));

app.get('/api/lines', authMiddleware, asyncRoute(async (req, res) => {
  await ensureLineGroupSchema();
  const managerBaseFields = `
    l.id,
    l.name,
    l.group_id::TEXT AS "groupId",
    COALESCE(g.name, '') AS "groupName",
    l.display_order AS "displayOrder",
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
    l.group_id::TEXT AS "groupId",
    COALESCE(g.name, '') AS "groupName",
    l.display_order AS "displayOrder",
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
              OR (
                   sl.shift = 'Full Day'
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments day_access
                     WHERE day_access.supervisor_user_id = $1
                       AND day_access.line_id = sl.line_id
                       AND day_access.shift = 'Day'
                   )
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments night_access
                     WHERE night_access.supervisor_user_id = $1
                       AND night_access.line_id = sl.line_id
                       AND night_access.shift = 'Night'
                   )
                 )
            )
        )
    ) AS "shiftCount"
  `;

  const query = req.user.role === 'manager'
    ? `SELECT ${managerBaseFields}
       FROM production_lines l
       LEFT JOIN line_groups g
         ON g.id = l.group_id
        AND g.is_active = TRUE
       WHERE l.is_active = TRUE
       ORDER BY ${LINE_ORDER_BY_SQL}`
    : `SELECT ${supervisorBaseFields}
       FROM production_lines l
       LEFT JOIN line_groups g
         ON g.id = l.group_id
        AND g.is_active = TRUE
       INNER JOIN (
         SELECT DISTINCT line_id
         FROM supervisor_line_shift_assignments
         WHERE supervisor_user_id = $1
       ) a ON a.line_id = l.id
       WHERE l.is_active = TRUE
       ORDER BY ${LINE_ORDER_BY_SQL}`;

  const result = await dbQuery(query, req.user.role === 'manager' ? [] : [req.user.id]);
  res.json({ lines: result.rows });
}));

app.patch('/api/lines/reorder', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  await ensureLineGroupSchema();
  const parsed = lineReorderSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid line reorder payload' });

  const nextGroupId = String(parsed.data.groupId || '').trim() || null;
  if (nextGroupId) {
    const missingGroupIds = await findMissingActiveLineGroupIds([nextGroupId]);
    if (missingGroupIds.length) return res.status(400).json({ error: 'Line group not found' });
  }

  const nextLineIds = parsed.data.lineIds;
  const uniqueLineIds = Array.from(new Set(nextLineIds));
  if (uniqueLineIds.length !== nextLineIds.length) {
    return res.status(400).json({ error: 'Duplicate line ids are not allowed' });
  }

  const existingLineResult = await dbQuery(
    `SELECT id::TEXT AS id
     FROM production_lines
     WHERE is_active = TRUE
       AND (($1::UUID IS NULL AND group_id IS NULL) OR group_id = $1::UUID)
     ORDER BY display_order ASC, created_at ASC`,
    [nextGroupId]
  );
  const existingLineIds = existingLineResult.rows.map((row) => row.id);
  if (existingLineIds.length !== uniqueLineIds.length) {
    return res.status(400).json({ error: 'Reorder payload must include all active lines in the target group' });
  }
  const existingLineSet = new Set(existingLineIds);
  const hasInvalidLine = uniqueLineIds.some((lineId) => !existingLineSet.has(lineId));
  if (hasInvalidLine) {
    return res.status(400).json({ error: 'One or more line ids are invalid for the target group' });
  }

  const client = await pool.connect();
  try {
    await client.query('BEGIN');
    for (let index = 0; index < uniqueLineIds.length; index += 1) {
      const updateResult = await client.query(
        `UPDATE production_lines
         SET display_order = $2, updated_at = NOW()
         WHERE id = $1
           AND is_active = TRUE
           AND (($3::UUID IS NULL AND group_id IS NULL) OR group_id = $3::UUID)`,
        [uniqueLineIds[index], index, nextGroupId]
      );
      if (!updateResult.rowCount) {
        throw new Error('Line reorder conflict. Please refresh and try again.');
      }
    }
    await client.query('COMMIT');
  } catch (error) {
    await client.query('ROLLBACK');
    throw error;
  } finally {
    client.release();
  }

  const reorderedResult = await dbQuery(
    `SELECT
       id,
       name,
       group_id::TEXT AS "groupId",
       display_order AS "displayOrder"
     FROM production_lines
     WHERE id = ANY($1::UUID[])`,
    [uniqueLineIds]
  );
  return res.json({ lines: reorderedResult.rows });
}));

app.get('/api/state-snapshot', authMiddleware, asyncRoute(async (req, res) => {
  await ensureLineGroupSchema();
  await ensureDataSourceSchema();
  await ensureLogSchema();
  await ensureSupervisorActionSchema();
  await ensureProductCatalogSchema();
  const linesQuery = req.user.role === 'manager'
    ? `SELECT
         l.id,
         l.name,
         l.group_id::TEXT AS "groupId",
         COALESCE(g.name, '') AS "groupName",
         l.display_order AS "displayOrder",
         l.secret_key AS "secretKey",
         l.created_at AS "createdAt",
         (
           SELECT COUNT(*)::INT FROM line_stages s WHERE s.line_id = l.id
         ) AS "stageCount",
         (
           SELECT COUNT(*)::INT FROM shift_logs sl WHERE sl.line_id = l.id
         ) AS "shiftCount"
       FROM production_lines l
       LEFT JOIN line_groups g
         ON g.id = l.group_id
        AND g.is_active = TRUE
       WHERE l.is_active = TRUE
       ORDER BY ${LINE_ORDER_BY_SQL}`
    : `SELECT
         l.id,
         l.name,
         l.group_id::TEXT AS "groupId",
         COALESCE(g.name, '') AS "groupName",
         l.display_order AS "displayOrder",
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
                   OR (
                   sl.shift = 'Full Day'
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments day_access
                     WHERE day_access.supervisor_user_id = $1
                       AND day_access.line_id = sl.line_id
                       AND day_access.shift = 'Day'
                   )
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments night_access
                     WHERE night_access.supervisor_user_id = $1
                       AND night_access.line_id = sl.line_id
                       AND night_access.shift = 'Night'
                   )
                 )
                 )
             )
         ) AS "shiftCount"
       FROM production_lines l
       LEFT JOIN line_groups g
         ON g.id = l.group_id
        AND g.is_active = TRUE
       INNER JOIN (
         SELECT DISTINCT line_id
         FROM supervisor_line_shift_assignments
         WHERE supervisor_user_id = $1
       ) a ON a.line_id = l.id
       WHERE l.is_active = TRUE
       ORDER BY ${LINE_ORDER_BY_SQL}`;

  const isManager = req.user.role === 'manager';
  const linesResult = await dbQuery(linesQuery, isManager ? [] : [req.user.id]);
  const lineRows = linesResult.rows;
  const lineIds = lineRows.map((line) => line.id);
  if (!lineIds.length) {
    const payload = { lines: [], supervisors: [], lineGroups: [], dataSources: [], supervisorActions: [], productCatalog: [] };
    [payload.supervisorActions, payload.productCatalog] = await Promise.all([
      fetchSupervisorActionsForUser(req.user),
      fetchProductCatalogForUser(req.user)
    ]);
    if (req.user.role === 'manager') {
      const [supervisors, lineGroups, dataSources] = await Promise.all([
        fetchSupervisorsWithAssignments(),
        fetchLineGroups(),
        fetchDataSources()
      ]);
      payload.supervisors = supervisors;
      payload.lineGroups = lineGroups;
      payload.dataSources = dataSources;
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
         COALESCE(data_source_id::TEXT, '') AS "dataSourceId",
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
             COALESCE(shift_logs.notes, '') AS notes,
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
             COALESCE(shift_logs.notes, '') AS notes,
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
                   OR (
                   shift_logs.shift = 'Full Day'
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments day_access
                     WHERE day_access.supervisor_user_id = $1
                       AND day_access.line_id = shift_logs.line_id
                       AND day_access.shift = 'Day'
                   )
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments night_access
                     WHERE night_access.supervisor_user_id = $1
                       AND night_access.line_id = shift_logs.line_id
                       AND night_access.shift = 'Night'
                   )
                 )
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
                   OR (
                   shift_break_logs.shift = 'Full Day'
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments day_access
                     WHERE day_access.supervisor_user_id = $1
                       AND day_access.line_id = shift_break_logs.line_id
                       AND day_access.shift = 'Day'
                   )
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments night_access
                     WHERE night_access.supervisor_user_id = $1
                       AND night_access.line_id = shift_break_logs.line_id
                       AND night_access.shift = 'Night'
                   )
                 )
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
             COALESCE(to_char(finish_time, 'HH24:MI'), '') AS "finishTime",
             units_produced AS "unitsProduced",
             COALESCE(run_crewing_pattern, '{}'::jsonb) AS "runCrewingPattern",
             COALESCE(run_logs.notes, '') AS notes,
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
             COALESCE(to_char(finish_time, 'HH24:MI'), '') AS "finishTime",
             units_produced AS "unitsProduced",
             COALESCE(run_crewing_pattern, '{}'::jsonb) AS "runCrewingPattern",
             COALESCE(run_logs.notes, '') AS notes,
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
                   OR (
                   run_logs.shift = 'Full Day'
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments day_access
                     WHERE day_access.supervisor_user_id = $1
                       AND day_access.line_id = run_logs.line_id
                       AND day_access.shift = 'Day'
                   )
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments night_access
                     WHERE night_access.supervisor_user_id = $1
                       AND night_access.line_id = run_logs.line_id
                       AND night_access.shift = 'Night'
                   )
                 )
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
             COALESCE(downtime_logs.notes, '') AS notes,
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
             COALESCE(downtime_logs.notes, '') AS notes,
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
                   OR (
                   downtime_logs.shift = 'Full Day'
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments day_access
                     WHERE day_access.supervisor_user_id = $1
                       AND day_access.line_id = downtime_logs.line_id
                       AND day_access.shift = 'Day'
                   )
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments night_access
                     WHERE night_access.supervisor_user_id = $1
                       AND night_access.line_id = downtime_logs.line_id
                       AND night_access.shift = 'Night'
                   )
                 )
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
      dataSourceId: row.dataSourceId,
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

  const payload = { lines, supervisors: [], lineGroups: [], dataSources: [], supervisorActions: [], productCatalog: [] };
  [payload.supervisorActions, payload.productCatalog] = await Promise.all([
    fetchSupervisorActionsForUser(req.user),
    fetchProductCatalogForUser(req.user)
  ]);
  if (req.user.role === 'manager') {
    const [supervisors, lineGroups, dataSources] = await Promise.all([
      fetchSupervisorsWithAssignments(),
      fetchLineGroups(),
      fetchDataSources()
    ]);
    payload.supervisors = supervisors;
    payload.lineGroups = lineGroups;
    payload.dataSources = dataSources;
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
      `INSERT INTO production_lines(name, secret_key, created_by_user_id, display_order)
       VALUES (
         $1,
         $2,
         $3,
         COALESCE((
           SELECT MAX(display_order) + 1
           FROM production_lines
           WHERE group_id IS NULL
             AND is_active = TRUE
         ), 0)
       )
       RETURNING
         id,
         name,
         group_id::TEXT AS "groupId",
         display_order AS "displayOrder",
         secret_key AS "secretKey",
         created_at AS "createdAt"`,
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
    const existingUserResult = await client.query(
      `SELECT id, role, is_active AS "isActive"
       FROM users
       WHERE username = $1
       FOR UPDATE`,
      [username]
    );
    const existingUser = existingUserResult.rows[0] || null;
    let supervisor = null;

    if (existingUser) {
      if (existingUser.role !== 'supervisor' || existingUser.isActive) {
        await client.query('ROLLBACK');
        return res.status(409).json({ error: 'Supervisor username already exists' });
      }
      const reactivated = await client.query(
        `UPDATE users
         SET name = $2,
             username = $3,
             password_hash = $4,
             role = 'supervisor',
             is_active = TRUE,
             updated_at = NOW()
         WHERE id = $1
         RETURNING id, name, username, is_active AS "isActive"`,
        [existingUser.id, name, username, passwordHash]
      );
      supervisor = reactivated.rows[0];
      await client.query(
        `DELETE FROM supervisor_line_shift_assignments WHERE supervisor_user_id = $1`,
        [supervisor.id]
      );
    } else {
      const insertUser = await client.query(
        `INSERT INTO users(name, username, password_hash, role, is_active)
         VALUES ($1, $2, $3, 'supervisor', TRUE)
         RETURNING id, name, username, is_active AS "isActive"`,
        [name, username, passwordHash]
      );
      supervisor = insertUser.rows[0];
    }

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

app.post('/api/supervisor-actions', authMiddleware, asyncRoute(async (req, res) => {
  await ensureSupervisorActionSchema();
  const parsed = supervisorActionCreateSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid action payload' });

  const data = parsed.data;
  const lineId = String(data.lineId || '').trim();
  if (lineId) {
    if (!(await isActiveLine(lineId))) return res.status(404).json({ error: 'Line not found' });
    if (!(await hasLineAccess(req.user, lineId))) return res.status(403).json({ error: 'Forbidden' });
  }

  const relatedEquipmentId = String(data.relatedEquipmentId || '').trim();
  if (relatedEquipmentId && !lineId) {
    return res.status(400).json({ error: 'Related equipment requires a valid line.' });
  }
  if (relatedEquipmentId && !(await isLineStage(lineId, relatedEquipmentId))) {
    return res.status(400).json({ error: 'Related equipment does not belong to the selected line.' });
  }

  const requestedAssigneeUsername = req.user.role === 'manager'
    ? normalizeActionAssigneeUsername(data.supervisorUsername)
    : normalizeActionAssigneeUsername(req.user.username);
  if (!requestedAssigneeUsername) {
    return res.status(400).json({ error: 'Supervisor username is required.' });
  }
  const assignee = await resolveSupervisorActionAssignee(
    requestedAssigneeUsername,
    req.user.role === 'manager' ? data.supervisorName : req.user.name
  );
  if (!assignee) return res.status(400).json({ error: 'Supervisor username is required.' });

  const title = String(data.title || '').trim();
  if (!title) return res.status(400).json({ error: 'Action title is required.' });
  const description = String(data.description || '').trim();
  const dueDate = normalizeIsoDate(data.dueDate);
  const relatedReasonCategory = normalizeActionReasonCategory(data.relatedReasonCategory);
  const relatedReasonDetail = normalizeActionReasonDetail(data.relatedReasonDetail, relatedReasonCategory);
  const createdByName = String(req.user.name || req.user.username || req.user.role).trim() || String(req.user.role || 'system').trim();

  const inserted = await dbQuery(
    `INSERT INTO supervisor_actions(
       supervisor_user_id,
       supervisor_username,
       supervisor_name,
       line_id,
       title,
       description,
       priority,
       status,
       due_date,
       related_equipment_stage_id,
       related_reason_category,
       related_reason_detail,
       created_by_user_id,
       created_by_name
     )
     VALUES (
       $1,
       $2,
       $3,
       NULLIF($4, '')::UUID,
       $5,
       $6,
       $7,
       $8,
       NULLIF($9, '')::DATE,
       NULLIF($10, '')::UUID,
       NULLIF($11, ''),
       NULLIF($12, ''),
       $13,
       $14
     )
     RETURNING
       id,
       supervisor_username AS "supervisorUsername",
       supervisor_name AS "supervisorName",
       COALESCE(line_id::TEXT, '') AS "lineId",
       title,
       COALESCE(description, '') AS description,
       priority,
       status,
       COALESCE(due_date::TEXT, '') AS "dueDate",
       COALESCE(related_equipment_stage_id::TEXT, '') AS "relatedEquipmentId",
       COALESCE(related_reason_category, '') AS "relatedReasonCategory",
       COALESCE(related_reason_detail, '') AS "relatedReasonDetail",
       created_at AS "createdAt",
       COALESCE(created_by_name, '') AS "createdBy"`,
    [
      assignee.supervisorUserId,
      assignee.supervisorUsername,
      assignee.supervisorName,
      lineId,
      title,
      description,
      data.priority,
      data.status,
      dueDate,
      relatedEquipmentId,
      relatedReasonCategory,
      relatedReasonDetail,
      req.user.id,
      createdByName
    ]
  );

  await writeAudit({
    lineId: lineId || null,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'CREATE_SUPERVISOR_ACTION',
    details: `Action created for ${assignee.supervisorUsername}: ${title}`
  });

  return res.status(201).json({ action: mapSupervisorActionRow(inserted.rows[0]) });
}));

app.patch('/api/supervisor-actions/:actionId', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  await ensureSupervisorActionSchema();
  const actionId = req.params.actionId;
  if (!z.string().uuid().safeParse(actionId).success) return res.status(400).json({ error: 'Invalid action id' });
  const parsed = supervisorActionUpdateSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid action update payload' });

  const existingResult = await dbQuery(
    `SELECT
       id,
       supervisor_user_id AS "supervisorUserId",
       supervisor_username AS "supervisorUsername",
       supervisor_name AS "supervisorName",
       COALESCE(line_id::TEXT, '') AS "lineId",
       title,
       COALESCE(description, '') AS description,
       priority,
       status,
       COALESCE(due_date::TEXT, '') AS "dueDate",
       COALESCE(related_equipment_stage_id::TEXT, '') AS "relatedEquipmentId",
       COALESCE(related_reason_category, '') AS "relatedReasonCategory",
       COALESCE(related_reason_detail, '') AS "relatedReasonDetail"
     FROM supervisor_actions
     WHERE id = $1`,
    [actionId]
  );
  if (!existingResult.rowCount) return res.status(404).json({ error: 'Action not found' });
  const existing = existingResult.rows[0];

  const data = parsed.data;
  const lineProvided = Object.prototype.hasOwnProperty.call(data, 'lineId');
  const nextLineId = lineProvided ? String(data.lineId || '').trim() : String(existing.lineId || '').trim();
  if (nextLineId && !(await isActiveLine(nextLineId))) return res.status(404).json({ error: 'Line not found' });

  const equipmentProvided = Object.prototype.hasOwnProperty.call(data, 'relatedEquipmentId');
  const nextEquipmentId = equipmentProvided ? String(data.relatedEquipmentId || '').trim() : String(existing.relatedEquipmentId || '').trim();
  if (nextEquipmentId && !nextLineId) {
    return res.status(400).json({ error: 'Related equipment requires a valid line.' });
  }
  if (nextEquipmentId && !(await isLineStage(nextLineId, nextEquipmentId))) {
    return res.status(400).json({ error: 'Related equipment does not belong to the selected line.' });
  }

  const assigneeProvided = Object.prototype.hasOwnProperty.call(data, 'supervisorUsername');
  const assignee = assigneeProvided
    ? await resolveSupervisorActionAssignee(data.supervisorUsername, data.supervisorName)
    : {
        supervisorUserId: existing.supervisorUserId || null,
        supervisorUsername: normalizeActionAssigneeUsername(existing.supervisorUsername),
        supervisorName: String(existing.supervisorName || existing.supervisorUsername || '').trim()
      };
  if (!assignee?.supervisorUsername) {
    return res.status(400).json({ error: 'Supervisor username is required.' });
  }

  const titleProvided = Object.prototype.hasOwnProperty.call(data, 'title');
  const nextTitle = titleProvided ? String(data.title || '').trim() : String(existing.title || '').trim();
  if (!nextTitle) return res.status(400).json({ error: 'Action title is required.' });

  const descriptionProvided = Object.prototype.hasOwnProperty.call(data, 'description');
  const nextDescription = descriptionProvided ? String(data.description || '').trim() : String(existing.description || '').trim();
  const nextPriorityRaw = String(data.priority || existing.priority || '').trim();
  const nextPriority = actionPriorityValues.includes(nextPriorityRaw) ? nextPriorityRaw : 'Medium';
  const nextStatusRaw = String(data.status || existing.status || '').trim();
  const nextStatus = actionStatusValues.includes(nextStatusRaw) ? nextStatusRaw : 'Open';

  const dueDateProvided = Object.prototype.hasOwnProperty.call(data, 'dueDate');
  const nextDueDate = dueDateProvided ? normalizeIsoDate(data.dueDate) : normalizeIsoDate(existing.dueDate);

  const reasonCategoryProvided = Object.prototype.hasOwnProperty.call(data, 'relatedReasonCategory');
  const reasonDetailProvided = Object.prototype.hasOwnProperty.call(data, 'relatedReasonDetail');
  const nextReasonCategory = reasonCategoryProvided
    ? normalizeActionReasonCategory(data.relatedReasonCategory)
    : normalizeActionReasonCategory(existing.relatedReasonCategory);
  const nextReasonDetail = normalizeActionReasonDetail(
    reasonDetailProvided ? data.relatedReasonDetail : existing.relatedReasonDetail,
    nextReasonCategory
  );

  const updated = await dbQuery(
    `UPDATE supervisor_actions
     SET
       supervisor_user_id = $2,
       supervisor_username = $3,
       supervisor_name = $4,
       line_id = NULLIF($5, '')::UUID,
       title = $6,
       description = $7,
       priority = $8,
       status = $9,
       due_date = NULLIF($10, '')::DATE,
       related_equipment_stage_id = NULLIF($11, '')::UUID,
       related_reason_category = NULLIF($12, ''),
       related_reason_detail = NULLIF($13, ''),
       updated_at = NOW()
     WHERE id = $1
     RETURNING
       id,
       supervisor_username AS "supervisorUsername",
       supervisor_name AS "supervisorName",
       COALESCE(line_id::TEXT, '') AS "lineId",
       title,
       COALESCE(description, '') AS description,
       priority,
       status,
       COALESCE(due_date::TEXT, '') AS "dueDate",
       COALESCE(related_equipment_stage_id::TEXT, '') AS "relatedEquipmentId",
       COALESCE(related_reason_category, '') AS "relatedReasonCategory",
       COALESCE(related_reason_detail, '') AS "relatedReasonDetail",
       created_at AS "createdAt",
       COALESCE(created_by_name, '') AS "createdBy"`,
    [
      actionId,
      assignee.supervisorUserId,
      assignee.supervisorUsername,
      assignee.supervisorName,
      nextLineId,
      nextTitle,
      nextDescription,
      nextPriority,
      nextStatus,
      nextDueDate,
      nextEquipmentId,
      nextReasonCategory,
      nextReasonDetail
    ]
  );

  await writeAudit({
    lineId: nextLineId || null,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'UPDATE_SUPERVISOR_ACTION',
    details: `Action ${actionId} updated`
  });

  return res.json({ action: mapSupervisorActionRow(updated.rows[0]) });
}));

app.delete('/api/supervisor-actions/:actionId', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  await ensureSupervisorActionSchema();
  const actionId = req.params.actionId;
  if (!z.string().uuid().safeParse(actionId).success) return res.status(400).json({ error: 'Invalid action id' });

  const existingResult = await dbQuery(
    `SELECT
       id,
       COALESCE(line_id::TEXT, '') AS "lineId",
       title
     FROM supervisor_actions
     WHERE id = $1`,
    [actionId]
  );
  if (!existingResult.rowCount) return res.status(404).json({ error: 'Action not found' });
  const existing = existingResult.rows[0];

  await dbQuery(`DELETE FROM supervisor_actions WHERE id = $1`, [actionId]);
  await writeAudit({
    lineId: existing.lineId || null,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'DELETE_SUPERVISOR_ACTION',
    details: `Action deleted: ${String(existing.title || actionId).trim() || actionId}`
  });

  return res.status(204).send();
}));

app.get('/api/product-catalog', authMiddleware, asyncRoute(async (req, res) => {
  await ensureProductCatalogSchema();
  const products = await fetchProductCatalogForUser(req.user);
  return res.json({ products });
}));

app.post('/api/product-catalog', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  await ensureProductCatalogSchema();
  const parsed = productCatalogEntrySchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid product catalog payload' });

  const values = normalizeProductCatalogValues(parsed.data.values);
  if (!hasProductCatalogIdentity(values)) {
    return res.status(400).json({ error: 'Code, Desc 1, or Desc 2 is required.' });
  }
  const lineSelection = normalizeProductCatalogLineSelection(parsed.data.lineIds);
  const missingLineIds = await findMissingActiveLineIds(lineSelection.lineIds);
  if (missingLineIds.length) {
    return res.status(400).json({ error: 'One or more assigned lines are invalid or inactive', lineIds: missingLineIds });
  }

  const inserted = await dbQuery(
    `INSERT INTO product_catalog_entries(
       catalog_values,
       line_ids,
       all_lines,
       created_by_user_id
     )
     VALUES (
       $1::jsonb,
       $2::UUID[],
       $3,
       $4
     )
     RETURNING
       id,
       COALESCE(catalog_values, '[]'::jsonb) AS values,
       all_lines AS "allLines",
       ARRAY(
         SELECT line_id::TEXT
         FROM unnest(COALESCE(line_ids, ARRAY[]::UUID[])) AS line_id
       ) AS "lineIds"`,
    [
      JSON.stringify(values),
      lineSelection.lineIds,
      lineSelection.allLines,
      req.user.id
    ]
  );
  const product = mapProductCatalogRow(inserted.rows[0]);
  if (!product) throw new Error('Saved product catalog entry could not be mapped.');

  await writeAudit({
    lineId: null,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'CREATE_PRODUCT_CATALOG_ENTRY',
    details: `Product created: ${product.values[1] || product.values[0] || product.id}`
  });

  return res.status(201).json({ product });
}));

app.patch('/api/product-catalog/:productId', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  await ensureProductCatalogSchema();
  const productId = req.params.productId;
  if (!z.string().uuid().safeParse(productId).success) return res.status(400).json({ error: 'Invalid product id' });
  const parsed = productCatalogEntrySchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid product catalog payload' });

  const values = normalizeProductCatalogValues(parsed.data.values);
  if (!hasProductCatalogIdentity(values)) {
    return res.status(400).json({ error: 'Code, Desc 1, or Desc 2 is required.' });
  }
  const lineSelection = normalizeProductCatalogLineSelection(parsed.data.lineIds);
  const missingLineIds = await findMissingActiveLineIds(lineSelection.lineIds);
  if (missingLineIds.length) {
    return res.status(400).json({ error: 'One or more assigned lines are invalid or inactive', lineIds: missingLineIds });
  }

  const updated = await dbQuery(
    `UPDATE product_catalog_entries
     SET
       catalog_values = $2::jsonb,
       line_ids = $3::UUID[],
       all_lines = $4,
       updated_at = NOW()
     WHERE id = $1
     RETURNING
       id,
       COALESCE(catalog_values, '[]'::jsonb) AS values,
       all_lines AS "allLines",
       ARRAY(
         SELECT line_id::TEXT
         FROM unnest(COALESCE(line_ids, ARRAY[]::UUID[])) AS line_id
       ) AS "lineIds"`,
    [
      productId,
      JSON.stringify(values),
      lineSelection.lineIds,
      lineSelection.allLines
    ]
  );
  if (!updated.rowCount) return res.status(404).json({ error: 'Product not found' });
  const product = mapProductCatalogRow(updated.rows[0]);
  if (!product) throw new Error('Updated product catalog entry could not be mapped.');

  await writeAudit({
    lineId: null,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'UPDATE_PRODUCT_CATALOG_ENTRY',
    details: `Product updated: ${product.values[1] || product.values[0] || product.id}`
  });

  return res.json({ product });
}));

app.get('/api/lines/:lineId', authMiddleware, asyncRoute(async (req, res) => {
  await ensureLineGroupSchema();
  await ensureDataSourceSchema();
  const lineId = req.params.lineId;
  if (!z.string().uuid().safeParse(lineId).success) return res.status(400).json({ error: 'Invalid line id' });
  if (!(await hasLineAccess(req.user, lineId))) return res.status(403).json({ error: 'Forbidden' });

  const lineResult = await dbQuery(
    req.user.role === 'manager'
      ? `SELECT
           l.id,
           l.name,
           l.group_id::TEXT AS "groupId",
           COALESCE(g.name, '') AS "groupName",
           l.display_order AS "displayOrder",
           l.secret_key AS "secretKey",
           l.created_at AS "createdAt",
           l.updated_at AS "updatedAt"
         FROM production_lines l
         LEFT JOIN line_groups g
           ON g.id = l.group_id
          AND g.is_active = TRUE
         WHERE l.id = $1 AND l.is_active = TRUE`
      : `SELECT
           l.id,
           l.name,
           l.group_id::TEXT AS "groupId",
           COALESCE(g.name, '') AS "groupName",
           l.display_order AS "displayOrder",
           NULL::TEXT AS "secretKey",
           l.created_at AS "createdAt",
           l.updated_at AS "updatedAt"
         FROM production_lines l
         LEFT JOIN line_groups g
           ON g.id = l.group_id
          AND g.is_active = TRUE
         WHERE l.id = $1 AND l.is_active = TRUE`,
    [lineId]
  );
  if (!lineResult.rowCount) return res.status(404).json({ error: 'Line not found' });

  const [stages, guides] = await Promise.all([
    dbQuery(
      `SELECT id, stage_order AS "stageOrder", stage_name AS "stageName", stage_type AS "stageType",
              day_crew AS "dayCrew", night_crew AS "nightCrew", max_throughput_per_crew AS "maxThroughputPerCrew",
              COALESCE(data_source_id::TEXT, '') AS "dataSourceId",
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
  await ensureLogSchema();
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
             COALESCE(shift_logs.notes, '') AS notes,
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
             COALESCE(shift_logs.notes, '') AS notes,
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
                   OR (
                   shift_logs.shift = 'Full Day'
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments day_access
                     WHERE day_access.supervisor_user_id = $1
                       AND day_access.line_id = shift_logs.line_id
                       AND day_access.shift = 'Day'
                   )
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments night_access
                     WHERE night_access.supervisor_user_id = $1
                       AND night_access.line_id = shift_logs.line_id
                       AND night_access.shift = 'Night'
                   )
                 )
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
                   OR (
                   shift_break_logs.shift = 'Full Day'
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments day_access
                     WHERE day_access.supervisor_user_id = $1
                       AND day_access.line_id = shift_break_logs.line_id
                       AND day_access.shift = 'Day'
                   )
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments night_access
                     WHERE night_access.supervisor_user_id = $1
                       AND night_access.line_id = shift_break_logs.line_id
                       AND night_access.shift = 'Night'
                   )
                 )
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
             COALESCE(to_char(finish_time, 'HH24:MI'), '') AS "finishTime",
             units_produced AS "unitsProduced",
             COALESCE(run_crewing_pattern, '{}'::jsonb) AS "runCrewingPattern",
             COALESCE(run_logs.notes, '') AS notes,
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
             COALESCE(to_char(finish_time, 'HH24:MI'), '') AS "finishTime",
             units_produced AS "unitsProduced",
             COALESCE(run_crewing_pattern, '{}'::jsonb) AS "runCrewingPattern",
             COALESCE(run_logs.notes, '') AS notes,
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
                   OR (
                   run_logs.shift = 'Full Day'
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments day_access
                     WHERE day_access.supervisor_user_id = $1
                       AND day_access.line_id = run_logs.line_id
                       AND day_access.shift = 'Day'
                   )
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments night_access
                     WHERE night_access.supervisor_user_id = $1
                       AND night_access.line_id = run_logs.line_id
                       AND night_access.shift = 'Night'
                   )
                 )
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
             COALESCE(downtime_logs.notes, '') AS notes,
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
             COALESCE(downtime_logs.notes, '') AS notes,
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
                   OR (
                   downtime_logs.shift = 'Full Day'
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments day_access
                     WHERE day_access.supervisor_user_id = $1
                       AND day_access.line_id = downtime_logs.line_id
                       AND day_access.shift = 'Day'
                   )
                   AND EXISTS (
                     SELECT 1
                     FROM supervisor_line_shift_assignments night_access
                     WHERE night_access.supervisor_user_id = $1
                       AND night_access.line_id = downtime_logs.line_id
                       AND night_access.shift = 'Night'
                   )
                 )
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
  await ensureDataSourceSchema();
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
           line_id, stage_order, stage_name, stage_type, day_crew, night_crew, max_throughput_per_crew, data_source_id, x, y, w, h
         ) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12)`,
        [
          lineId,
          stage.stageOrder,
          stage.stageName,
          stage.stageType,
          stage.dayCrew,
          stage.nightCrew,
          stage.maxThroughputPerCrew,
          stage.dataSourceId ? stage.dataSourceId : null,
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
    if (error?.code === '23505' && String(error?.constraint || '') === 'idx_line_stages_data_source_unique') {
      return res.status(409).json({ error: 'Data source is already connected to another equipment stage.' });
    }
    if (error?.code === '23503' && String(error?.constraint || '').includes('line_stages_data_source_id_fkey')) {
      return res.status(400).json({ error: 'Data source not found.' });
    }
    throw error;
  } finally {
    client.release();
  }
}));

app.patch('/api/lines/:lineId', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  await ensureLineGroupSchema();
  const lineId = req.params.lineId;
  if (!z.string().uuid().safeParse(lineId).success) return res.status(400).json({ error: 'Invalid line id' });
  const parsed = lineUpdateSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid line update payload' });

  const hasNameUpdate = Object.prototype.hasOwnProperty.call(parsed.data, 'name');
  const hasGroupUpdate = Object.prototype.hasOwnProperty.call(parsed.data, 'groupId');
  const nextName = hasNameUpdate ? parsed.data.name.trim() : null;
  const nextGroupId = hasGroupUpdate
    ? (String(parsed.data.groupId || '').trim() || null)
    : undefined;
  if (nextGroupId) {
    const missingGroupIds = await findMissingActiveLineGroupIds([nextGroupId]);
    if (missingGroupIds.length) return res.status(400).json({ error: 'Line group not found' });
  }

  try {
    const currentResult = await dbQuery(
      `SELECT
         l.id,
         l.name,
         l.group_id::TEXT AS "groupId",
         COALESCE(g.name, '') AS "groupName",
         l.display_order AS "displayOrder"
       FROM production_lines l
       LEFT JOIN line_groups g
         ON g.id = l.group_id
        AND g.is_active = TRUE
       WHERE l.id = $1
         AND l.is_active = TRUE`,
      [lineId]
    );
    if (!currentResult.rowCount) return res.status(404).json({ error: 'Line not found' });
    const currentLine = currentResult.rows[0];
    const groupChanged = hasGroupUpdate && String(currentLine.groupId || '') !== String(nextGroupId || '');
    const nextDisplayOrderResult = groupChanged
      ? await dbQuery(
        `SELECT COALESCE(MAX(display_order) + 1, 0) AS "nextDisplayOrder"
         FROM production_lines
         WHERE is_active = TRUE
           AND (($1::UUID IS NULL AND group_id IS NULL) OR group_id = $1::UUID)`,
        [nextGroupId || null]
      )
      : null;
    const nextDisplayOrder = groupChanged
      ? Number(nextDisplayOrderResult?.rows?.[0]?.nextDisplayOrder || 0)
      : null;

    const setParts = ['updated_at = NOW()'];
    const params = [lineId];
    if (hasNameUpdate) {
      setParts.push(`name = $${params.length + 1}`);
      params.push(nextName);
    }
    if (hasGroupUpdate) {
      setParts.push(`group_id = $${params.length + 1}`);
      params.push(nextGroupId);
      if (groupChanged) {
        setParts.push(`display_order = $${params.length + 1}`);
        params.push(nextDisplayOrder);
      }
    }

    const result = await dbQuery(
      `UPDATE production_lines
       SET ${setParts.join(', ')}
       WHERE id = $1 AND is_active = TRUE
       RETURNING id`,
      params
    );
    if (!result.rowCount) return res.status(404).json({ error: 'Line not found' });

    const updatedLineResult = await dbQuery(
      `SELECT
         l.id,
         l.name,
         l.group_id::TEXT AS "groupId",
         COALESCE(g.name, '') AS "groupName",
         l.display_order AS "displayOrder",
         l.secret_key AS "secretKey",
         l.updated_at AS "updatedAt"
       FROM production_lines l
       LEFT JOIN line_groups g
         ON g.id = l.group_id
        AND g.is_active = TRUE
       WHERE l.id = $1
         AND l.is_active = TRUE`,
      [lineId]
    );
    if (!updatedLineResult.rowCount) return res.status(404).json({ error: 'Line not found' });
    const updatedLine = updatedLineResult.rows[0];

    if (hasNameUpdate && currentLine.name !== updatedLine.name) {
      await writeAudit({
        lineId,
        actorUserId: req.user.id,
        actorName: req.user.name,
        actorRole: req.user.role,
        action: 'RENAME_LINE',
        details: `Line renamed to: ${updatedLine.name}`
      });
    }
    if (hasGroupUpdate && String(currentLine.groupId || '') !== String(updatedLine.groupId || '')) {
      const previousGroup = currentLine.groupName || 'Ungrouped';
      const nextGroup = updatedLine.groupName || 'Ungrouped';
      await writeAudit({
        lineId,
        actorUserId: req.user.id,
        actorName: req.user.name,
        actorRole: req.user.role,
        action: 'UPDATE_LINE_GROUP',
        details: `Line group changed: ${previousGroup} -> ${nextGroup}`
      });
    }

    return res.json({ line: updatedLine });
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

app.post('/api/lines/:lineId/load-sample-data', authMiddleware, requireRole('manager'), asyncRoute(async (req, res) => {
  const lineId = req.params.lineId;
  if (!z.string().uuid().safeParse(lineId).success) return res.status(400).json({ error: 'Invalid line id' });
  const parsed = loadPermanentSampleDataSchema.safeParse(req.body || {});
  if (!parsed.success) return res.status(400).json({ error: 'Invalid load sample payload' });

  const lineRes = await dbQuery(
    `SELECT id, name
     FROM production_lines
     WHERE id = $1 AND is_active = TRUE`,
    [lineId]
  );
  if (!lineRes.rowCount) return res.status(404).json({ error: 'Line not found' });

  const stageRows = (
    await dbQuery(
      `SELECT
         id,
         stage_name AS "stageName",
         day_crew AS "dayCrew",
         night_crew AS "nightCrew"
       FROM line_stages
       WHERE line_id = $1
       ORDER BY stage_order ASC`,
      [lineId]
    )
  ).rows;
  const sample = buildPermanentSampleData(stageRows, { days: 84 });
  const replaceExisting = parsed.data.replaceExisting !== false;
  const shiftPayload = sample.shiftRows.map((row) => ({
    date: row.date,
    shift: row.shift,
    crew_on_shift: row.crewOnShift,
    start_time: row.startTime,
    finish_time: row.finishTime
  }));
  const breakPayload = sample.breakRows.map((row) => ({
    date: row.date,
    shift: row.shift,
    break_start: row.breakStart,
    break_finish: row.breakFinish
  }));
  const runPayload = sample.runRows.map((row) => ({
    date: row.date,
    shift: row.shift,
    product: row.product,
    set_up_start_time: row.setUpStartTime || '',
    production_start_time: row.productionStartTime,
    finish_time: row.finishTime,
    units_produced: row.unitsProduced
  }));
  const downtimePayload = sample.downtimeRows.map((row) => ({
    date: row.date,
    shift: row.shift,
    downtime_start: row.downtimeStart,
    downtime_finish: row.downtimeFinish,
    equipment_stage_id: row.equipmentStageId || '',
    reason: row.reason || ''
  }));

  const client = await pool.connect();
  let shiftCount = 0;
  let breakCount = 0;
  let runCount = 0;
  let downtimeCount = 0;
  try {
    await client.query('BEGIN');

    if (replaceExisting) {
      await client.query(`DELETE FROM shift_break_logs WHERE line_id = $1`, [lineId]);
      await client.query(`DELETE FROM shift_logs WHERE line_id = $1`, [lineId]);
      await client.query(`DELETE FROM run_logs WHERE line_id = $1`, [lineId]);
      await client.query(`DELETE FROM downtime_logs WHERE line_id = $1`, [lineId]);
    }

    const shiftBreakInsert = await client.query(
      `WITH shift_data AS (
         SELECT
           s.date::date AS date,
           s.shift::text AS shift,
           GREATEST(0, COALESCE(s.crew_on_shift, 0)) AS crew_on_shift,
           s.start_time::time AS start_time,
           s.finish_time::time AS finish_time
         FROM jsonb_to_recordset($2::jsonb) AS s(
           date text,
           shift text,
           crew_on_shift integer,
           start_time text,
           finish_time text
         )
       ),
       inserted_shifts AS (
         INSERT INTO shift_logs(line_id, date, shift, crew_on_shift, start_time, finish_time, submitted_by_user_id)
         SELECT $1, date, shift, crew_on_shift, start_time, finish_time, $3
         FROM shift_data
         RETURNING id, date, shift
       ),
       break_data AS (
         SELECT
           b.date::date AS date,
           b.shift::text AS shift,
           b.break_start::time AS break_start,
           b.break_finish::time AS break_finish
         FROM jsonb_to_recordset($4::jsonb) AS b(
           date text,
           shift text,
           break_start text,
           break_finish text
         )
       ),
       inserted_breaks AS (
         INSERT INTO shift_break_logs(shift_log_id, line_id, date, shift, break_start, break_finish, submitted_by_user_id)
         SELECT
           s.id,
           $1,
           b.date,
           b.shift,
           b.break_start,
           b.break_finish,
           $3
         FROM break_data b
         INNER JOIN inserted_shifts s
           ON s.date = b.date
          AND s.shift = b.shift
         RETURNING id
       )
       SELECT
         (SELECT COUNT(*)::INT FROM inserted_shifts) AS "shiftCount",
         (SELECT COUNT(*)::INT FROM inserted_breaks) AS "breakCount"`,
      [lineId, JSON.stringify(shiftPayload), req.user.id, JSON.stringify(breakPayload)]
    );
    shiftCount = Number(shiftBreakInsert.rows?.[0]?.shiftCount || 0);
    breakCount = Number(shiftBreakInsert.rows?.[0]?.breakCount || 0);

    const runInsert = await client.query(
      `WITH run_data AS (
         SELECT
           r.date::date AS date,
           r.shift::text AS shift,
           r.product::text AS product,
           NULLIF(COALESCE(r.set_up_start_time, ''), '')::time AS set_up_start_time,
           r.production_start_time::time AS production_start_time,
           NULLIF(COALESCE(r.finish_time, ''), '')::time AS finish_time,
           GREATEST(0, COALESCE(r.units_produced, 0)) AS units_produced
         FROM jsonb_to_recordset($2::jsonb) AS r(
           date text,
           shift text,
           product text,
           set_up_start_time text,
           production_start_time text,
           finish_time text,
           units_produced numeric
         )
       ),
       inserted_runs AS (
         INSERT INTO run_logs(
           line_id, date, shift, product, setup_start_time, production_start_time, finish_time, units_produced, run_crewing_pattern, submitted_by_user_id
         )
         SELECT
           $1,
           date,
           shift,
           product,
           set_up_start_time,
           production_start_time,
           finish_time,
           units_produced,
           '{}'::jsonb,
           $3
         FROM run_data
         RETURNING id
       )
       SELECT COUNT(*)::INT AS count FROM inserted_runs`,
      [lineId, JSON.stringify(runPayload), req.user.id]
    );
    runCount = Number(runInsert.rows?.[0]?.count || 0);

    const downtimeInsert = await client.query(
      `WITH downtime_data AS (
         SELECT
           d.date::date AS date,
           d.shift::text AS shift,
           d.downtime_start::time AS downtime_start,
           d.downtime_finish::time AS downtime_finish,
           NULLIF(COALESCE(d.equipment_stage_id, ''), '')::uuid AS equipment_stage_id,
           NULLIF(COALESCE(d.reason, ''), '') AS reason
         FROM jsonb_to_recordset($2::jsonb) AS d(
           date text,
           shift text,
           downtime_start text,
           downtime_finish text,
           equipment_stage_id text,
           reason text
         )
       ),
       inserted_downtime AS (
         INSERT INTO downtime_logs(
           line_id, date, shift, downtime_start, downtime_finish, equipment_stage_id, reason, submitted_by_user_id
         )
         SELECT
           $1,
           date,
           shift,
           downtime_start,
           downtime_finish,
           equipment_stage_id,
           reason,
           $3
         FROM downtime_data
         RETURNING id
       )
       SELECT COUNT(*)::INT AS count FROM inserted_downtime`,
      [lineId, JSON.stringify(downtimePayload), req.user.id]
    );
    downtimeCount = Number(downtimeInsert.rows?.[0]?.count || 0);

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
    action: 'LOAD_PERMANENT_SAMPLE_DATA',
    details: `Permanent sample data loaded (${shiftCount} shifts, ${breakCount} breaks, ${runCount} runs, ${downtimeCount} downtime rows)`
  });

  return res.json({
    ok: true,
    counts: {
      shifts: shiftCount,
      breaks: breakCount,
      runs: runCount,
      downtime: downtimeCount
    }
  });
}));

app.post('/api/logs/shifts', authMiddleware, asyncRoute(async (req, res) => {
  await ensureLogSchema();
  const parsed = shiftLogSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: 'Invalid shift log payload' });
  if (!(await isActiveLine(parsed.data.lineId))) return res.status(404).json({ error: 'Line not found' });
  if (!(await hasLineShiftAccess(req.user, parsed.data.lineId, parsed.data.shift))) return res.status(403).json({ error: 'Forbidden' });

  const result = await dbQuery(
    `INSERT INTO shift_logs(line_id, date, shift, crew_on_shift, start_time, finish_time, notes, submitted_by_user_id)
     VALUES ($1, $2, $3, $4, $5, $6, COALESCE($7, ''), $8)
     RETURNING
       id,
       line_id AS "lineId",
       date,
       shift,
       crew_on_shift AS "crewOnShift",
       start_time AS "startTime",
       finish_time AS "finishTime",
       COALESCE(notes, '') AS notes,
       submitted_at AS "submittedAt"`,
    [
      parsed.data.lineId,
      parsed.data.date,
      parsed.data.shift,
      parsed.data.crewOnShift,
      parsed.data.startTime,
      parsed.data.finishTime,
      String(parsed.data.notes ?? '').trim(),
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
  const breakFinishRaw = String(parsed.data.breakFinish ?? '').trim();
  const breakFinish = breakFinishRaw || null;
  const creatingOpenBreak = !breakFinish;

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
  if (!canMutateSubmittedLog(req.user, shiftLog.submittedByUserId)) {
    return res.status(403).json({ error: 'Supervisors can only update their own shift logs' });
  }
  if (!shiftLog.isOpen && req.user.role !== 'manager') return res.status(400).json({ error: 'Shift is already complete' });
  if (!shiftLog.isOpen && creatingOpenBreak) {
    return res.status(400).json({ error: 'Completed shifts require a break finish time' });
  }

  if (creatingOpenBreak) {
    const openBreakResult = await dbQuery(
      `SELECT id
       FROM shift_break_logs
       WHERE shift_log_id = $1
         AND break_finish IS NULL
       LIMIT 1`,
      [logId]
    );
    if (openBreakResult.rowCount) return res.status(400).json({ error: 'An open break already exists for this shift' });
  }

  const result = await dbQuery(
    `INSERT INTO shift_break_logs(shift_log_id, line_id, date, shift, break_start, break_finish, submitted_by_user_id)
     VALUES ($1, $2, $3, $4, $5, NULLIF($6, '')::time, $7)
     RETURNING
       id,
       shift_log_id AS "shiftLogId",
       line_id AS "lineId",
       date::TEXT AS date,
       shift,
       to_char(break_start, 'HH24:MI') AS "breakStart",
       COALESCE(to_char(break_finish, 'HH24:MI'), '') AS "breakFinish",
       submitted_at AS "submittedAt"`,
    [logId, shiftLog.lineId, shiftLog.date, shiftLog.shift, parsed.data.breakStart, breakFinish, req.user.id]
  );

  await writeAudit({
    lineId: shiftLog.lineId,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: creatingOpenBreak ? 'START_SHIFT_BREAK' : 'CREATE_SHIFT_BREAK',
    details: creatingOpenBreak
      ? `${shiftLog.shift} ${shiftLog.date} break started ${parsed.data.breakStart}`
      : `${shiftLog.shift} ${shiftLog.date} break logged ${parsed.data.breakStart} to ${breakFinish}`
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
  if (!canMutateSubmittedLog(req.user, breakLog.submittedByUserId)) {
    return res.status(403).json({ error: 'Supervisors can only update their own break logs' });
  }
  if (!breakLog.shiftOpen && req.user.role !== 'manager') return res.status(400).json({ error: 'Shift is already complete' });

  if (parsed.data.breakFinish === undefined && parsed.data.breakStart === undefined) {
    return res.status(400).json({ error: 'At least one break field must be provided' });
  }

  const result = await dbQuery(
    `UPDATE shift_break_logs
     SET
       break_start = COALESCE($2::time, break_start),
       break_finish = COALESCE($3::time, break_finish),
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
    [breakId, parsed.data.breakStart || null, parsed.data.breakFinish || null]
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

app.delete('/api/logs/shifts/:logId/breaks/:breakId', authMiddleware, asyncRoute(async (req, res) => {
  const logId = req.params.logId;
  const breakId = req.params.breakId;
  if (!z.string().uuid().safeParse(logId).success) return res.status(400).json({ error: 'Invalid shift log id' });
  if (!z.string().uuid().safeParse(breakId).success) return res.status(400).json({ error: 'Invalid break log id' });

  const breakResult = await dbQuery(
    `SELECT
       b.id,
       b.line_id AS "lineId",
       b.date::TEXT AS date,
       b.shift,
       to_char(b.break_start, 'HH24:MI') AS "breakStart",
       COALESCE(to_char(b.break_finish, 'HH24:MI'), '') AS "breakFinish",
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
  if (!canMutateSubmittedLog(req.user, breakLog.submittedByUserId)) {
    return res.status(403).json({ error: 'Supervisors can only delete their own break logs' });
  }
  if (!breakLog.shiftOpen && req.user.role !== 'manager') return res.status(400).json({ error: 'Shift is already complete' });

  await dbQuery(`DELETE FROM shift_break_logs WHERE id = $1`, [breakId]);
  await writeAudit({
    lineId: breakLog.lineId,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'DELETE_SHIFT_BREAK',
    details: `${breakLog.shift} ${breakLog.date} break deleted (${breakLog.breakStart}${breakLog.breakFinish ? ` to ${breakLog.breakFinish}` : ''})`
  });

  return res.status(204).send();
}));

app.post('/api/logs/runs', authMiddleware, asyncRoute(async (req, res) => {
  await ensureLogSchema();
  const parsed = runLogSchema.safeParse(req.body);
  if (!parsed.success) return res.status(400).json({ error: 'Invalid run log payload' });
  if (!(await isActiveLine(parsed.data.lineId))) return res.status(404).json({ error: 'Line not found' });
  if (!(await hasLineShiftAccess(req.user, parsed.data.lineId, parsed.data.shift))) return res.status(403).json({ error: 'Forbidden' });

  const result = await dbQuery(
    `INSERT INTO run_logs(
       line_id, date, shift, product, setup_start_time, production_start_time, finish_time, units_produced, run_crewing_pattern, notes, submitted_by_user_id
     )
     VALUES ($1, $2, $3, $4, NULLIF($5, '')::time, $6, NULLIF($7, '')::time, $8, $9::jsonb, COALESCE($10, ''), $11)
     RETURNING
       id,
       line_id AS "lineId",
       date::TEXT AS date,
       shift,
       product,
       COALESCE(to_char(setup_start_time, 'HH24:MI'), '') AS "setUpStartTime",
       to_char(production_start_time, 'HH24:MI') AS "productionStartTime",
       COALESCE(to_char(finish_time, 'HH24:MI'), '') AS "finishTime",
       units_produced AS "unitsProduced",
       COALESCE(run_crewing_pattern, '{}'::jsonb) AS "runCrewingPattern",
       COALESCE(notes, '') AS notes,
       submitted_at AS "submittedAt"`,
    [
      parsed.data.lineId,
      parsed.data.date,
      parsed.data.shift,
      parsed.data.product,
      parsed.data.setUpStartTime || '',
      parsed.data.productionStartTime,
      parsed.data.finishTime ?? '',
      parsed.data.unitsProduced,
      JSON.stringify(parsed.data.runCrewingPattern || {}),
      String(parsed.data.notes ?? '').trim(),
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
  await ensureLogSchema();
  const parsed = downtimeLogSchema.safeParse(req.body);
  if (!parsed.success) {
    const details = parsed.error.issues
      .map((issue) => `${issue.path.join('.') || 'payload'}: ${issue.message}`)
      .join('; ');
    return res.status(400).json({
      error: details ? `Invalid downtime log payload: ${details}` : 'Invalid downtime log payload'
    });
  }
  if (!(await isActiveLine(parsed.data.lineId))) return res.status(404).json({ error: 'Line not found' });
  if (!(await hasLineShiftAccess(req.user, parsed.data.lineId, parsed.data.shift))) return res.status(403).json({ error: 'Forbidden' });
  if (parsed.data.equipmentStageId && !(await isLineStage(parsed.data.lineId, parsed.data.equipmentStageId))) {
    return res.status(400).json({ error: 'Equipment stage does not belong to this line' });
  }

  const result = await dbQuery(
    `INSERT INTO downtime_logs(line_id, date, shift, downtime_start, downtime_finish, equipment_stage_id, reason, notes, submitted_by_user_id)
     VALUES ($1, $2, $3, $4, $5, $6, NULLIF($7, ''), COALESCE($8, ''), $9)
     RETURNING
       id,
       line_id AS "lineId",
       date,
       shift,
       reason,
       COALESCE(notes, '') AS notes,
       submitted_at AS "submittedAt"`,
    [
      parsed.data.lineId,
      parsed.data.date,
      parsed.data.shift,
      parsed.data.downtimeStart,
      parsed.data.downtimeFinish,
      parsed.data.equipmentStageId || null,
      parsed.data.reason || '',
      String(parsed.data.notes ?? '').trim(),
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
  await ensureLogSchema();
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

  const data = parsed.data;
  const accessShift = data.shift || existing.shift;
  if (!(await hasLineShiftAccess(req.user, existing.lineId, accessShift))) return res.status(403).json({ error: 'Forbidden' });
  if (!canMutateSubmittedLog(req.user, existing.submittedByUserId)) {
    return res.status(403).json({ error: 'Supervisors can only edit their own logs' });
  }

  const notesProvided = Object.prototype.hasOwnProperty.call(data, 'notes');
  const notesValue = notesProvided ? String(data.notes ?? '').trim() : '';
  const result = await dbQuery(
    `WITH updated_shift AS (
       UPDATE shift_logs
       SET
         date = COALESCE($2::date, date),
         shift = COALESCE($3, shift),
         crew_on_shift = COALESCE($4, crew_on_shift),
         start_time = COALESCE(NULLIF($5, '')::time, start_time),
         finish_time = COALESCE(NULLIF($6, '')::time, finish_time),
         notes = CASE
           WHEN $7::boolean IS FALSE THEN notes
           ELSE $8
         END,
         submitted_at = NOW()
       WHERE id = $1
       RETURNING
         id,
         line_id AS "lineId",
         date,
         shift,
         crew_on_shift AS "crewOnShift",
         to_char(start_time, 'HH24:MI') AS "startTime",
         to_char(finish_time, 'HH24:MI') AS "finishTime",
         COALESCE(notes, '') AS notes,
         submitted_at AS "submittedAt"
     ),
     synced_breaks AS (
       UPDATE shift_break_logs b
       SET
         date = u.date::date,
         shift = u.shift
       FROM updated_shift u
       WHERE b.shift_log_id = u.id
     )
     SELECT
       id,
       "lineId",
       date::TEXT AS date,
       shift,
       "crewOnShift",
       "startTime",
       "finishTime",
       notes,
       "submittedAt"
     FROM updated_shift`,
    [
      logId,
      data.date ?? null,
      data.shift ?? null,
      data.crewOnShift,
      data.startTime ?? null,
      data.finishTime ?? null,
      notesProvided,
      notesValue
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
  await ensureLogSchema();
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

  const data = parsed.data;
  const accessShift = data.shift || existing.shift;
  if (!(await hasLineShiftAccess(req.user, existing.lineId, accessShift))) return res.status(403).json({ error: 'Forbidden' });
  if (!canMutateSubmittedLog(req.user, existing.submittedByUserId)) {
    return res.status(403).json({ error: 'Supervisors can only edit their own logs' });
  }

  const notesProvided = Object.prototype.hasOwnProperty.call(data, 'notes');
  const notesValue = notesProvided ? String(data.notes ?? '').trim() : '';
  const result = await dbQuery(
    `UPDATE run_logs
     SET
       date = COALESCE($2::date, date),
       shift = COALESCE($3, shift),
       product = COALESCE(NULLIF($4, ''), product),
       setup_start_time = CASE WHEN $5::text IS NULL THEN setup_start_time ELSE NULLIF($5, '')::time END,
       production_start_time = COALESCE(NULLIF($6, '')::time, production_start_time),
       finish_time = CASE WHEN $7::text IS NULL THEN finish_time ELSE NULLIF($7, '')::time END,
       units_produced = COALESCE($8, units_produced),
       run_crewing_pattern = COALESCE($9::jsonb, run_crewing_pattern),
       notes = CASE
         WHEN $10::boolean IS FALSE THEN notes
         ELSE $11
       END,
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
       COALESCE(to_char(finish_time, 'HH24:MI'), '') AS "finishTime",
       units_produced AS "unitsProduced",
       COALESCE(run_crewing_pattern, '{}'::jsonb) AS "runCrewingPattern",
       COALESCE(notes, '') AS notes,
       submitted_at AS "submittedAt"`,
    [
      logId,
      data.date ?? null,
      data.shift ?? null,
      data.product ?? null,
      data.setUpStartTime ?? null,
      data.productionStartTime ?? null,
      data.finishTime ?? null,
      data.unitsProduced,
      data.runCrewingPattern === undefined ? null : JSON.stringify(data.runCrewingPattern || {}),
      notesProvided,
      notesValue
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
  await ensureLogSchema();
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

  const data = parsed.data;
  const accessShift = data.shift || existing.shift;
  if (!(await hasLineShiftAccess(req.user, existing.lineId, accessShift))) return res.status(403).json({ error: 'Forbidden' });
  if (!canMutateSubmittedLog(req.user, existing.submittedByUserId)) {
    return res.status(403).json({ error: 'Supervisors can only edit their own logs' });
  }
  const equipmentStageProvided = Object.prototype.hasOwnProperty.call(data, 'equipmentStageId');
  const equipmentStageValue = equipmentStageProvided ? String(data.equipmentStageId || '').trim() : '';
  if (equipmentStageValue && !(await isLineStage(existing.lineId, equipmentStageValue))) {
    return res.status(400).json({ error: 'Equipment stage does not belong to this line' });
  }
  const reasonProvided = Object.prototype.hasOwnProperty.call(data, 'reason');
  const reasonValue = reasonProvided ? String(data.reason ?? '').trim() : '';
  const notesProvided = Object.prototype.hasOwnProperty.call(data, 'notes');
  const notesValue = notesProvided ? String(data.notes ?? '').trim() : '';
  const result = await dbQuery(
    `UPDATE downtime_logs
     SET
       date = COALESCE($2::date, date),
       shift = COALESCE($3, shift),
       downtime_start = COALESCE(NULLIF($4, '')::time, downtime_start),
       downtime_finish = COALESCE(NULLIF($5, '')::time, downtime_finish),
       equipment_stage_id = CASE
         WHEN $6::boolean IS FALSE THEN equipment_stage_id
         WHEN NULLIF($7, '')::uuid IS NULL THEN NULL
         ELSE NULLIF($7, '')::uuid
       END,
       reason = CASE
         WHEN $8::boolean IS FALSE THEN reason
         ELSE NULLIF($9, '')
       END,
       notes = CASE
         WHEN $10::boolean IS FALSE THEN notes
         ELSE $11
       END,
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
       COALESCE(notes, '') AS notes,
       submitted_at AS "submittedAt"`,
    [
      logId,
      data.date ?? null,
      data.shift ?? null,
      data.downtimeStart ?? null,
      data.downtimeFinish ?? null,
      equipmentStageProvided,
      equipmentStageValue,
      reasonProvided,
      reasonValue,
      notesProvided,
      notesValue
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

app.delete('/api/logs/shifts/:logId', authMiddleware, asyncRoute(async (req, res) => {
  const logId = req.params.logId;
  if (!z.string().uuid().safeParse(logId).success) return res.status(400).json({ error: 'Invalid shift log id' });

  const existingResult = await dbQuery(
    `SELECT
       id,
       line_id AS "lineId",
       date::TEXT AS date,
       shift,
       submitted_by_user_id AS "submittedByUserId"
     FROM shift_logs
     WHERE id = $1`,
    [logId]
  );
  if (!existingResult.rowCount) return res.status(404).json({ error: 'Shift log not found' });
  const existing = existingResult.rows[0];

  if (!(await hasLineShiftAccess(req.user, existing.lineId, existing.shift))) return res.status(403).json({ error: 'Forbidden' });
  if (!canMutateSubmittedLog(req.user, existing.submittedByUserId)) {
    return res.status(403).json({ error: 'Supervisors can only delete their own logs' });
  }

  await dbQuery(`DELETE FROM shift_logs WHERE id = $1`, [logId]);
  await writeAudit({
    lineId: existing.lineId,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'DELETE_SHIFT_LOG',
    details: `Shift log ${logId} deleted (${existing.shift} ${existing.date})`
  });

  return res.status(204).send();
}));

app.delete('/api/logs/runs/:logId', authMiddleware, asyncRoute(async (req, res) => {
  const logId = req.params.logId;
  if (!z.string().uuid().safeParse(logId).success) return res.status(400).json({ error: 'Invalid run log id' });

  const existingResult = await dbQuery(
    `SELECT
       id,
       line_id AS "lineId",
       date::TEXT AS date,
       shift,
       product,
       submitted_by_user_id AS "submittedByUserId"
     FROM run_logs
     WHERE id = $1`,
    [logId]
  );
  if (!existingResult.rowCount) return res.status(404).json({ error: 'Run log not found' });
  const existing = existingResult.rows[0];

  if (!(await hasLineShiftAccess(req.user, existing.lineId, existing.shift))) return res.status(403).json({ error: 'Forbidden' });
  if (!canMutateSubmittedLog(req.user, existing.submittedByUserId)) {
    return res.status(403).json({ error: 'Supervisors can only delete their own logs' });
  }

  await dbQuery(`DELETE FROM run_logs WHERE id = $1`, [logId]);
  await writeAudit({
    lineId: existing.lineId,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'DELETE_RUN_LOG',
    details: `Run log ${logId} deleted (${existing.product}, ${existing.date} ${existing.shift})`
  });

  return res.status(204).send();
}));

app.delete('/api/logs/downtime/:logId', authMiddleware, asyncRoute(async (req, res) => {
  const logId = req.params.logId;
  if (!z.string().uuid().safeParse(logId).success) return res.status(400).json({ error: 'Invalid downtime log id' });

  const existingResult = await dbQuery(
    `SELECT
       id,
       line_id AS "lineId",
       date::TEXT AS date,
       shift,
       COALESCE(reason, '') AS reason,
       submitted_by_user_id AS "submittedByUserId"
     FROM downtime_logs
     WHERE id = $1`,
    [logId]
  );
  if (!existingResult.rowCount) return res.status(404).json({ error: 'Downtime log not found' });
  const existing = existingResult.rows[0];

  if (!(await hasLineShiftAccess(req.user, existing.lineId, existing.shift))) return res.status(403).json({ error: 'Forbidden' });
  if (!canMutateSubmittedLog(req.user, existing.submittedByUserId)) {
    return res.status(403).json({ error: 'Supervisors can only delete their own logs' });
  }

  await dbQuery(`DELETE FROM downtime_logs WHERE id = $1`, [logId]);
  await writeAudit({
    lineId: existing.lineId,
    actorUserId: req.user.id,
    actorName: req.user.name,
    actorRole: req.user.role,
    action: 'DELETE_DOWNTIME_LOG',
    details: `Downtime log ${logId} deleted (${existing.reason || `${existing.date} ${existing.shift}`})`
  });

  return res.status(204).send();
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
  console.log(`API listening on port ${config.port}`);
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
