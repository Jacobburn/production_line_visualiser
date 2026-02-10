const STORAGE_KEY = "kebab-line-data-v2";
const STORAGE_BACKUP_KEY = "kebab-line-data-v2-backup";
const API_BASE_URL = `${
  localStorage.getItem("production-line-api-base") ||
  window.PRODUCTION_LINE_API_BASE ||
  "http://localhost:4000"
}`.replace(/\/+$/, "");
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

const STAGES = [
  { id: "s1", name: "1. Tumbler", crew: 2, group: "prep", match: ["tumbler"], x: 66.5, y: 64, w: 10, h: 26 },
  { id: "s2", name: "2. Transfer", crew: 0, group: "prep", match: ["transfer"], x: 56.5, y: 74, w: 7.5, h: 10, kind: "transfer" },
  { id: "s3", name: "3. Kebab Box Pack", crew: 4, group: "prep", match: ["box", "pack"], x: 23, y: 49, w: 14, h: 16 },
  { id: "s4", name: "4. Kebab Box Cut", crew: 1, group: "prep", match: ["box", "cut"], x: 13.5, y: 49, w: 9.5, h: 16 },
  { id: "s5", name: "5. Kebab Box Unload", crew: 1, group: "prep", match: ["box", "unload"], x: 5, y: 49, w: 8.5, h: 16 },
  { id: "s6", name: "6. Transfer", crew: 0, group: "prep", match: ["transfer"], x: 0.8, y: 34, w: 7.5, h: 10, kind: "transfer" },
  { id: "s7", name: "7. Kebab Split", crew: 1, group: "main", match: ["split"], x: 4.5, y: 14, w: 8.5, h: 15 },
  { id: "s8", name: "8. Marinate", crew: 5, group: "main", match: ["marinate"], x: 13.5, y: 14, w: 15, h: 15 },
  { id: "s9", name: "9. Wipe", crew: 1, group: "main", match: ["wipe"], x: 33.5, y: 14, w: 6, h: 15 },
  { id: "s10", name: "10. Proseal", crew: 1, group: "main", match: ["proseal"], x: 40.5, y: 14, w: 9.5, h: 15 },
  { id: "s11", name: "11. Metal Detector", crew: 1, group: "main", match: ["metal"], x: 55.5, y: 14, w: 6.8, h: 15 },
  { id: "s12", name: "12. Bottom Labeller", crew: 1, group: "main", match: ["bottom", "labeller", "labeler"], x: 69.5, y: 14, w: 9.5, h: 15 },
  { id: "s13", name: "13. Top Labeller", crew: 1, group: "main", match: ["top", "labeller", "labeler"], x: 80.5, y: 14, w: 8, h: 15 },
  { id: "s14", name: "14. Pack", crew: 2, group: "main", match: ["pack"], x: 90, y: 14, w: 5.5, h: 15 }
];

const SHIFT_COLUMNS = ["Date", "Shift", "Crew On Shift", "Start Time", "Break 1 Start", "Break 2 Start", "Break 3 Start", "Finish Time", "Total Shift Time"];
const RUN_COLUMNS = ["Date", "Shift", "Product", "Set Up Start Time", "Production Start Time", "Finish Time", "Units Produced", "Gross Production Time", "Associated Down Time", "Net Production Time", "Gross Run Rate", "Net Run Rate"];
const DOWN_COLUMNS = ["Date", "Shift", "Downtime Start", "Downtime Finish", "Downtime (mins)", "Equipment", "Reason"];
const AUDIT_COLUMNS = ["When", "Actor", "Action", "Details"];
const DASHBOARD_COLUMNS = ["Line", "Date", "Shift", "Units", "Downtime (min)", "Utilisation (%)", "Net Run Rate (u/min)", "Bottleneck", "Staffing"];
const DEFAULT_SUPERVISORS = [
  { id: "sup-1", name: "Supervisor", username: "supervisor", password: "supervisor", mode: "all" },
  { id: "sup-2", name: "Day Lead", username: "daylead", password: "day123", mode: "even" },
  { id: "sup-3", name: "Night Lead", username: "nightlead", password: "night123", mode: "odd" }
];

let appState = loadState();
let state = appState.lines[appState.activeLineId];
let visualiserDragMoved = false;
appState.supervisors = normalizeSupervisors(appState.supervisors, appState.lines);
let managerBackendSession = {
  backendToken: localStorage.getItem("production-line-manager-token") || "",
  backendLineMap: (() => {
    try {
      const parsed = JSON.parse(localStorage.getItem("production-line-manager-line-map") || "{}");
      return parsed && typeof parsed === "object" ? parsed : {};
    } catch {
      return {};
    }
  })(),
  backendStageMap: (() => {
    try {
      const parsed = JSON.parse(localStorage.getItem("production-line-manager-stage-map") || "{}");
      return parsed && typeof parsed === "object" ? parsed : {};
    } catch {
      return {};
    }
  })()
};

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function generateSecretKey() {
  const alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  let key = "";
  for (let i = 0; i < 8; i += 1) key += alphabet[Math.floor(Math.random() * alphabet.length)];
  return key;
}

function defaultStageCrew(stages = STAGES) {
  return Object.fromEntries(stages.map((stage) => [stage.id, { crew: stage.crew }]));
}

function defaultStageSettings(stages = STAGES) {
  return Object.fromEntries(
    stages.map((stage) => [
      stage.id,
      {
        maxThroughput: stage.kind === "transfer" ? 3 : 2
      }
    ])
  );
}

function defaultCrewByShift(stages = STAGES) {
  return {
    Day: defaultStageCrew(stages),
    Night: defaultStageCrew(stages)
  };
}

function normalizeCrewByShift(parsed, stages = STAGES) {
  if (parsed?.crewsByShift?.Day && parsed?.crewsByShift?.Night) {
    const day = {};
    const night = {};
    stages.forEach((stage) => {
      day[stage.id] = { crew: num(parsed.crewsByShift.Day?.[stage.id]?.crew ?? stage.crew) };
      night[stage.id] = { crew: num(parsed.crewsByShift.Night?.[stage.id]?.crew ?? stage.crew) };
    });
    return { Day: day, Night: night };
  }

  const legacy = parsed?.stages || defaultStageCrew(stages);
  return {
    Day: clone(legacy),
    Night: JSON.parse(JSON.stringify(legacy))
  };
}

function normalizeStageSettings(parsed, stages = STAGES) {
  const defaults = defaultStageSettings(stages);
  const incoming = parsed?.stageSettings || {};
  stages.forEach((stage) => {
    if (!incoming[stage.id]) incoming[stage.id] = {};
    if (incoming[stage.id].maxThroughput === undefined || incoming[stage.id].maxThroughput === null || incoming[stage.id].maxThroughput === "") {
      incoming[stage.id].maxThroughput = defaults[stage.id].maxThroughput;
    }
  });
  return incoming;
}

function normalizeFlowGuides(guides) {
  if (!Array.isArray(guides)) return [];
  return guides.map((guide, index) => ({
    id: guide?.id || `fg-${Date.now()}-${index}`,
    type: guide?.type === "arrow" ? "arrow" : guide?.type === "shape" ? "shape" : "line",
    x: Math.max(0, num(guide?.x)),
    y: Math.max(0, num(guide?.y)),
    w: Math.max(2, num(guide?.w) || 12),
    h: Math.max(1, num(guide?.h) || 2),
    angle: num(guide?.angle),
    src: guide?.type === "shape" ? String(guide?.src || "") : ""
  }));
}

function supervisorModeAssignments(mode, lines) {
  const ids = Object.keys(lines || {});
  if (!ids.length) return [];
  if (mode === "even") return ids.filter((_, i) => i % 2 === 0);
  if (mode === "odd") return ids.filter((_, i) => i % 2 === 1);
  return ids;
}

function defaultSupervisors(lines) {
  return DEFAULT_SUPERVISORS.map((sup) => ({
    id: sup.id,
    name: sup.name,
    username: sup.username.toLowerCase(),
    password: sup.password,
    assignedLineIds: supervisorModeAssignments(sup.mode, lines)
  }));
}

function normalizeSupervisors(supervisors, lines) {
  const source = Array.isArray(supervisors) && supervisors.length ? supervisors : defaultSupervisors(lines);
  const seen = new Set();
  return source
    .map((sup, index) => {
      const username = String(sup?.username || "").trim().toLowerCase();
      if (!username || seen.has(username)) return null;
      seen.add(username);
      const assigned = Array.isArray(sup?.assignedLineIds) ? sup.assignedLineIds.filter((id) => lines[id]) : [];
      return {
        id: sup?.id || `sup-${Date.now()}-${index}`,
        name: String(sup?.name || username).trim() || username,
        username,
        password: String(sup?.password || ""),
        assignedLineIds: assigned
      };
    })
    .filter(Boolean);
}

function supervisorByUsername(username) {
  const key = String(username || "").trim().toLowerCase();
  return (appState.supervisors || []).find((sup) => sup.username === key) || null;
}

function normalizeSupervisorSession(session, supervisors, lines) {
  if (!session || typeof session.username !== "string") return null;
  const sup = (supervisors || []).find((item) => item.username === String(session.username).toLowerCase());
  const assignedFromSession = Array.isArray(session.assignedLineIds) ? session.assignedLineIds.filter((id) => lines[id]) : [];
  const assigned = sup
    ? (Array.isArray(sup.assignedLineIds) ? sup.assignedLineIds.filter((id) => lines[id]) : [])
    : assignedFromSession;
  if (!sup && !assigned.length && !session.backendToken) return null;
  const backendLineMap = {};
  if (session.backendLineMap && typeof session.backendLineMap === "object") {
    Object.entries(session.backendLineMap).forEach(([localId, backendId]) => {
      if (lines[localId] && UUID_RE.test(String(backendId || ""))) backendLineMap[localId] = String(backendId);
    });
  }
  return {
    username: sup.username,
    assignedLineIds: assigned,
    backendToken: typeof session.backendToken === "string" && session.backendToken ? session.backendToken : "",
    backendLineMap
  };
}

function selectedSupervisorLineId() {
  const sel = document.getElementById("supervisorLineSelect");
  return sel ? sel.value : "";
}

function selectedSupervisorLine() {
  const id = selectedSupervisorLineId();
  return id && appState.lines[id] ? appState.lines[id] : null;
}

function supervisorEquipmentOptions(line) {
  if (!line) return `<option value="">Equipment Stage</option>`;
  const stages = line.stages?.length ? line.stages : STAGES;
  return [
    `<option value="">Equipment Stage</option>`,
    ...stages.map((stage, index) => `<option value="${stage.id}">${stageDisplayName(stage, index)}</option>`)
  ].join("");
}

function stageMaxThroughputForLine(line, stageId) {
  return Math.max(0, num(line?.stageSettings?.[stageId]?.maxThroughput));
}

function stageCrewForShiftForLine(line, stageId, shift) {
  const crew = num(line?.crewsByShift?.[shift]?.[stageId]?.crew);
  const stage = (line?.stages || []).find((s) => s.id === stageId);
  if (crew > 0) return crew;
  if (stage?.kind === "transfer") return 1;
  return 0;
}

function stageTotalMaxThroughputForLine(line, stageId, shift) {
  return stageMaxThroughputForLine(line, stageId) * stageCrewForShiftForLine(line, stageId, shift);
}

function derivedDataForLine(line) {
  const downtimeRows = (line?.downtimeRows || []).map(computeDowntimeRow);
  const shiftRows = (line?.shiftRows || []).map(computeShiftRow);
  const downtimeByShift = new Map();
  downtimeRows.forEach((row) => {
    const key = `${row.date}__${row.shift}`;
    downtimeByShift.set(key, (downtimeByShift.get(key) || 0) + num(row.downtimeMins));
  });
  const runRows = (line?.runRows || []).map((row) => computeRunRow(row, downtimeByShift));
  return { shiftRows, runRows, downtimeRows };
}

function computeLineMetrics(line, date, shift) {
  const stages = line?.stages?.length ? line.stages : STAGES;
  const data = derivedDataForLine(line || {});
  const selectedRunRows = data.runRows.filter((row) => row.date === date && row.shift === shift);
  const selectedDownRows = data.downtimeRows.filter((row) => row.date === date && row.shift === shift);
  const selectedShiftRows = data.shiftRows.filter((row) => row.date === date && row.shift === shift);
  const latestShiftRow = selectedShiftRows.length ? selectedShiftRows[selectedShiftRows.length - 1] : null;
  const shiftMins = num(latestShiftRow?.totalShiftTime);
  const hasCrewLog = Boolean(latestShiftRow) && latestShiftRow?.crewOnShift !== "" && latestShiftRow?.crewOnShift !== undefined;
  const crewOnShift = hasCrewLog ? Math.max(0, num(latestShiftRow?.crewOnShift)) : 0;
  const requiredCrew = requiredCrewForLineShift(line, shift);
  const understaffedBy = hasCrewLog ? Math.max(0, requiredCrew - crewOnShift) : 0;
  const staffingCallout = !latestShiftRow ? "No shift data" : understaffedBy > 0 ? `Understaffed by ${understaffedBy}` : "Fully staffed";
  const units = selectedRunRows.reduce((sum, row) => sum + num(row.unitsProduced), 0);
  const totalDowntime = selectedDownRows.reduce((sum, row) => sum + num(row.downtimeMins), 0);
  const totalNetTime = selectedRunRows.reduce((sum, row) => sum + num(row.netProductionTime), 0);
  const netRunRate = totalNetTime > 0 ? units / totalNetTime : 0;
  let utilAccumulator = 0;
  let utilCount = 0;
  let bottleneckStageName = "-";
  let bottleneckUtil = -1;

  stages.forEach((stage, index) => {
    const stageDowntime = selectedDownRows
      .filter((row) => matchesStage(stage, row.equipment))
      .reduce((sum, row) => sum + num(row.downtimeMins), 0);
    const uptimeRatio = shiftMins > 0 ? Math.max(0, (shiftMins - stageDowntime) / shiftMins) : 0;
    const stageRate = netRunRate * uptimeRatio;
    const totalMax = stageTotalMaxThroughputForLine(line, stage.id, shift);
    const utilisation = totalMax > 0 ? (stageRate / totalMax) * 100 : 0;
    utilAccumulator += Math.max(0, utilisation);
    utilCount += 1;
    if (utilisation > bottleneckUtil) {
      bottleneckUtil = utilisation;
      bottleneckStageName = stageDisplayName(stage, index);
    }
  });

  return {
    lineName: line?.name || "Line",
    date,
    shift,
    units,
    totalDowntime,
    lineUtil: utilCount > 0 ? utilAccumulator / utilCount : 0,
    netRunRate,
    bottleneckStageName,
    crewOnShift,
    requiredCrew,
    understaffedBy,
    staffingCallout
  };
}

function csvEscape(value) {
  const raw = value === null || value === undefined ? "" : String(value);
  if (!/[",\n]/.test(raw)) return raw;
  return `"${raw.replace(/"/g, "\"\"")}"`;
}

function toCsv(rows, columns) {
  return [columns.join(","), ...rows.map((row) => columns.map((col) => csvEscape(row[col])).join(","))].join("\n");
}

function downloadTextFile(fileName, text, mimeType = "text/plain;charset=utf-8") {
  const blob = new Blob([text], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}

function loadState() {
  const saved = localStorage.getItem(STORAGE_KEY) || localStorage.getItem(STORAGE_BACKUP_KEY);
  const defaultLine = makeDefaultLine("line-1", "Production Line 1");

  if (!saved) {
    return {
      activeView: "home",
      appMode: "manager",
      dashboardDate: todayISO(),
      dashboardShift: "Day",
      supervisorMobileMode: false,
      supervisorMainTab: "supervisorVisual",
      supervisorTab: "superShift",
      supervisorSelectedLineId: defaultLine.id,
      supervisorSelectedDate: todayISO(),
      supervisorSelectedShift: "Day",
      supervisorSession: null,
      supervisors: defaultSupervisors({ [defaultLine.id]: defaultLine }),
      activeLineId: defaultLine.id,
      lines: { [defaultLine.id]: defaultLine }
    };
  }

  try {
    const parsed = JSON.parse(saved);
    if (parsed?.lines && parsed?.activeLineId) {
      const lines = {};
      Object.entries(parsed.lines).forEach(([id, line]) => {
        lines[id] = normalizeLine(id, line);
      });
      const supervisors = normalizeSupervisors(parsed.supervisors, lines);
      const activeLineId = lines[parsed.activeLineId] ? parsed.activeLineId : Object.keys(lines)[0];
      return {
        activeView: parsed.activeView || "home",
        appMode: parsed.appMode === "supervisor" ? "supervisor" : "manager",
        dashboardDate: parsed.dashboardDate || todayISO(),
        dashboardShift: parsed.dashboardShift === "Night" ? "Night" : "Day",
        supervisorMobileMode: Boolean(parsed.supervisorMobileMode),
        supervisorMainTab: ["supervisorVisual", "supervisorData"].includes(parsed.supervisorMainTab) ? parsed.supervisorMainTab : "supervisorVisual",
        supervisorTab: ["superShift", "superRun", "superDown"].includes(parsed.supervisorTab) ? parsed.supervisorTab : "superShift",
        supervisorSelectedLineId: parsed.supervisorSelectedLineId || activeLineId,
        supervisorSelectedDate: parsed.supervisorSelectedDate || todayISO(),
        supervisorSelectedShift: parsed.supervisorSelectedShift === "Night" ? "Night" : "Day",
        supervisorSession: normalizeSupervisorSession(parsed.supervisorSession, supervisors, lines),
        supervisors,
        activeLineId,
        lines
      };
    }

    // Legacy single-line migration.
    const migrated = normalizeLine("line-1", parsed);
    return {
      activeView: "home",
      appMode: "manager",
      dashboardDate: todayISO(),
      dashboardShift: "Day",
      supervisorMobileMode: false,
      supervisorMainTab: "supervisorVisual",
      supervisorTab: "superShift",
      supervisorSelectedLineId: migrated.id,
      supervisorSelectedDate: todayISO(),
      supervisorSelectedShift: "Day",
      supervisorSession: null,
      supervisors: defaultSupervisors({ [migrated.id]: migrated }),
      activeLineId: migrated.id,
      lines: { [migrated.id]: migrated }
    };
  } catch {
    return {
      activeView: "home",
      appMode: "manager",
      dashboardDate: todayISO(),
      dashboardShift: "Day",
      supervisorMobileMode: false,
      supervisorMainTab: "supervisorVisual",
      supervisorTab: "superShift",
      supervisorSelectedLineId: defaultLine.id,
      supervisorSelectedDate: todayISO(),
      supervisorSelectedShift: "Day",
      supervisorSession: null,
      supervisors: defaultSupervisors({ [defaultLine.id]: defaultLine }),
      activeLineId: defaultLine.id,
      lines: { [defaultLine.id]: defaultLine }
    };
  }
}

function saveState() {
  if (state && state.id) {
    appState.lines[state.id] = state;
    appState.activeLineId = state.id;
  }
  const payload = JSON.stringify(appState);
  localStorage.setItem(STORAGE_BACKUP_KEY, payload);
  localStorage.setItem(STORAGE_KEY, payload);
}

function makeDefaultLine(id, name) {
  const stages = clone(STAGES);
  return {
    id,
    name,
    secretKey: generateSecretKey(),
    selectedDate: todayISO(),
    selectedShift: "Day",
    visualEditMode: false,
    flowGuides: [],
    selectedStageId: STAGES[0].id,
    activeDataTab: "dataShift",
    trendGranularity: "daily",
    trendMonth: todayISO().slice(0, 7),
    crewsByShift: defaultCrewByShift(stages),
    stageSettings: defaultStageSettings(stages),
    stages,
    shiftRows: [],
    runRows: [],
    downtimeRows: [],
    supervisorLogs: [],
    auditRows: []
  };
}

function normalizeLine(id, line) {
  const base = makeDefaultLine(id, line?.name || "Production Line");
  const stages = Array.isArray(line?.stages) && line.stages.length ? line.stages : clone(STAGES);
  return {
    ...base,
    ...line,
    id,
    name: line?.name || base.name,
    secretKey: line?.secretKey || base.secretKey,
    visualEditMode: Boolean(line?.visualEditMode),
    flowGuides: normalizeFlowGuides(line?.flowGuides),
    crewsByShift: normalizeCrewByShift(line || {}, stages),
    stageSettings: normalizeStageSettings(line || {}, stages),
    stages,
    shiftRows: line?.shiftRows || [],
    runRows: line?.runRows || [],
    downtimeRows: line?.downtimeRows || [],
    supervisorLogs: Array.isArray(line?.supervisorLogs) ? line.supervisorLogs : [],
    auditRows: Array.isArray(line?.auditRows) ? line.auditRows : []
  };
}

function getStages() {
  return state?.stages?.length ? state.stages : STAGES;
}

function makeStagesFromBuilder(rows) {
  const stripLeadingNumber = (name) => String(name || "").replace(/^\s*\d+\.\s*/, "").trim();
  const cleaned = rows
    .map((row, index) => ({
      id: `s${index + 1}`,
      name: `${index + 1}. ${stripLeadingNumber(row.name) || "Stage"}`,
      crew: Math.max(0, num(row.crew)),
      maxThroughput: Math.max(0, num(row.maxThroughput)),
      group: row.group || "main",
      kind: row.group === "transfer" ? "transfer" : undefined,
      x: num(row.x),
      y: num(row.y),
      w: num(row.w),
      h: num(row.h),
      match: String(row.name || "")
        .toLowerCase()
        .split(/[^a-z0-9]+/)
        .filter(Boolean)
    }))
    .filter((row) => row.name);

  const topCount = Math.max(1, Math.ceil(cleaned.length * 0.6));
  const topGap = topCount > 1 ? 86 / (topCount - 1) : 0;
  const bottomCount = Math.max(0, cleaned.length - topCount);
  const bottomGap = bottomCount > 1 ? 62 / (bottomCount - 1) : 0;

  return cleaned.map((stage, index) => {
    if (stage.x > 0 && stage.y > 0) {
      return {
        ...stage,
        w: stage.w > 0 ? stage.w : stage.kind === "transfer" ? 7 : 9.5,
        h: stage.h > 0 ? stage.h : stage.kind === "transfer" ? 10 : 15
      };
    }
    if (index < topCount) {
      return {
        ...stage,
        x: 4 + topGap * index,
        y: 14,
        w: stage.kind === "transfer" ? 7 : 9.5,
        h: 15
      };
    }
    const bi = index - topCount;
    return {
      ...stage,
      x: 6 + bottomGap * bi,
      y: 49,
      w: stage.kind === "transfer" ? 7.5 : 10,
      h: 16
    };
  });
}

function makeCrewByShiftFromStages(stages) {
  const rows = Object.fromEntries(stages.map((stage) => [stage.id, { crew: Math.max(0, num(stage.crew)) }]));
  return { Day: clone(rows), Night: clone(rows) };
}

function makeSettingsFromStages(stages) {
  return Object.fromEntries(
    stages.map((stage) => [
      stage.id,
      {
        maxThroughput: Math.max(0, num(stage.maxThroughput) || (stage.kind === "transfer" ? 3 : 2))
      }
    ])
  );
}

function nextAutoLineName() {
  const existing = new Set(Object.values(appState.lines || {}).map((line) => String(line?.name || "").trim().toLowerCase()));
  let n = Object.keys(appState.lines || {}).length + 1;
  while (existing.has(`production line ${n}`.toLowerCase())) n += 1;
  return `Production Line ${n}`;
}

function formatDateLocal(date) {
  const y = date.getFullYear();
  const m = `${date.getMonth() + 1}`.padStart(2, "0");
  const d = `${date.getDate()}`.padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function parseDateLocal(isoDate) {
  const [y, m, d] = String(isoDate).split("-").map(Number);
  if (!y || !m || !d) return new Date();
  return new Date(y, m - 1, d);
}

function todayISO() {
  return formatDateLocal(new Date());
}

function num(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function formatNum(value, digits = 1) {
  return Number(value || 0).toLocaleString(undefined, { maximumFractionDigits: digits });
}

function strictTimeValid(value) {
  if (value === null || value === undefined || value === "") return false;
  return /^([01]\d|2[0-3]):([0-5]\d)$/.test(String(value).trim());
}

function optionalStrictTimeValid(value) {
  if (value === null || value === undefined || value === "") return true;
  return strictTimeValid(value);
}

function rowIsValidDateShift(date, shift) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(date || "")) && (shift === "Day" || shift === "Night");
}

async function apiRequest(path, { method = "GET", token = "", body } = {}) {
  const headers = {};
  if (token) headers.Authorization = `Bearer ${token}`;
  if (body !== undefined) headers["Content-Type"] = "application/json";
  const response = await fetch(`${API_BASE_URL}${path}`, {
    method,
    headers,
    body: body !== undefined ? JSON.stringify(body) : undefined
  });
  const text = await response.text();
  let payload = null;
  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    payload = text ? { raw: text } : null;
  }
  if (!response.ok) {
    throw new Error(payload?.error || payload?.message || `API ${response.status}`);
  }
  return payload;
}

function findLocalLineIdByName(name) {
  const target = String(name || "").trim().toLowerCase();
  if (!target) return "";
  const match = Object.values(appState.lines || {}).find((line) => String(line?.name || "").trim().toLowerCase() === target);
  return match?.id || "";
}

async function ensureBackendLineId(localLineId, session) {
  if (!session?.backendToken || !appState.lines?.[localLineId]) return "";
  if (UUID_RE.test(localLineId)) return localLineId;
  if (session.backendLineMap?.[localLineId]) return session.backendLineMap[localLineId];
  const payload = await apiRequest("/api/lines", { token: session.backendToken });
  const lines = Array.isArray(payload?.lines) ? payload.lines : [];
  const localName = String(appState.lines[localLineId]?.name || "").trim().toLowerCase();
  const backendLine = lines.find((line) => String(line?.name || "").trim().toLowerCase() === localName);
  if (backendLine?.id && UUID_RE.test(backendLine.id)) {
    session.backendLineMap = session.backendLineMap || {};
    session.backendLineMap[localLineId] = backendLine.id;
    if (session === managerBackendSession) persistManagerBackendSession();
    else saveState();
    return backendLine.id;
  }
  if (session.role === "manager") {
    try {
      const created = await apiRequest("/api/lines", {
        method: "POST",
        token: session.backendToken,
        body: {
          name: appState.lines[localLineId].name,
          secretKey: appState.lines[localLineId].secretKey || generateSecretKey()
        }
      });
      const createdId = created?.line?.id;
      if (createdId && UUID_RE.test(createdId)) {
        session.backendLineMap = session.backendLineMap || {};
        session.backendLineMap[localLineId] = createdId;
        if (session === managerBackendSession) persistManagerBackendSession();
        else saveState();
        return createdId;
      }
    } catch (error) {
      console.warn("Backend line auto-create failed:", error);
    }
  }
  return "";
}

function persistManagerBackendSession() {
  localStorage.setItem("production-line-manager-token", managerBackendSession.backendToken || "");
  localStorage.setItem("production-line-manager-line-map", JSON.stringify(managerBackendSession.backendLineMap || {}));
  localStorage.setItem("production-line-manager-stage-map", JSON.stringify(managerBackendSession.backendStageMap || {}));
}

async function ensureManagerBackendSession() {
  managerBackendSession.role = "manager";
  if (managerBackendSession.backendToken) return managerBackendSession;
  const loginPayload = await apiRequest("/api/auth/login", {
    method: "POST",
    body: { username: "manager", password: "manager123" }
  });
  if (!loginPayload?.token) throw new Error("Manager API login failed");
  managerBackendSession.backendToken = loginPayload.token;
  persistManagerBackendSession();
  return managerBackendSession;
}

function stageNameCore(name) {
  return String(name || "")
    .replace(/^\s*\d+\.\s*/, "")
    .trim()
    .toLowerCase();
}

async function ensureBackendStageId(localLineId, localStageId, session) {
  if (!localLineId || !localStageId || !session?.backendToken) return "";
  const key = `${localLineId}::${localStageId}`;
  if (session.backendStageMap?.[key]) return session.backendStageMap[key];
  const backendLineId = await ensureBackendLineId(localLineId, session);
  if (!backendLineId) return "";
  const payload = await apiRequest(`/api/lines/${backendLineId}`, { token: session.backendToken });
  const backendStages = Array.isArray(payload?.stages) ? payload.stages : [];
  const localStage = (appState.lines[localLineId]?.stages || []).find((stage) => stage.id === localStageId);
  if (!localStage) return "";
  const localName = stageNameCore(localStage.name);
  const matched = backendStages.find((stage) => stageNameCore(stage.stageName) === localName);
  if (!matched?.id || !UUID_RE.test(matched.id)) return "";
  session.backendStageMap = session.backendStageMap || {};
  session.backendStageMap[key] = matched.id;
  if (session === managerBackendSession) persistManagerBackendSession();
  else saveState();
  return matched.id;
}

async function syncManagerShiftLog(payload) {
  try {
    const session = await ensureManagerBackendSession();
    const backendLineId = await ensureBackendLineId(payload.lineId, session);
    if (!backendLineId) return;
    await apiRequest("/api/logs/shifts", {
      method: "POST",
      token: session.backendToken,
      body: {
        ...payload,
        lineId: backendLineId
      }
    });
  } catch (error) {
    console.warn("Backend manager shift sync failed:", error);
  }
}

async function syncManagerRunLog(payload) {
  try {
    const session = await ensureManagerBackendSession();
    const backendLineId = await ensureBackendLineId(payload.lineId, session);
    if (!backendLineId) return;
    await apiRequest("/api/logs/runs", {
      method: "POST",
      token: session.backendToken,
      body: {
        ...payload,
        lineId: backendLineId
      }
    });
  } catch (error) {
    console.warn("Backend manager run sync failed:", error);
  }
}

async function syncManagerDowntimeLog(payload) {
  try {
    const session = await ensureManagerBackendSession();
    const backendLineId = await ensureBackendLineId(payload.lineId, session);
    if (!backendLineId) return;
    const backendEquipmentId = await ensureBackendStageId(payload.lineId, payload.equipment, session);
    await apiRequest("/api/logs/downtime", {
      method: "POST",
      token: session.backendToken,
      body: {
        lineId: backendLineId,
        date: payload.date,
        shift: payload.shift,
        downtimeStart: payload.downtimeStart,
        downtimeFinish: payload.downtimeFinish,
        equipmentStageId: backendEquipmentId || null,
        reason: payload.reason || ""
      }
    });
  } catch (error) {
    console.warn("Backend manager downtime sync failed:", error);
  }
}

async function syncSupervisorShiftLog(session, payload) {
  try {
    const backendLineId = await ensureBackendLineId(payload.lineId, session);
    if (!backendLineId) return;
    await apiRequest("/api/logs/shifts", {
      method: "POST",
      token: session.backendToken,
      body: {
        ...payload,
        lineId: backendLineId
      }
    });
  } catch (error) {
    console.warn("Backend shift sync failed:", error);
  }
}

async function syncSupervisorRunLog(session, payload) {
  try {
    const backendLineId = await ensureBackendLineId(payload.lineId, session);
    if (!backendLineId) return;
    await apiRequest("/api/logs/runs", {
      method: "POST",
      token: session.backendToken,
      body: {
        ...payload,
        lineId: backendLineId
      }
    });
  } catch (error) {
    console.warn("Backend run sync failed:", error);
  }
}

async function syncSupervisorDowntimeLog(session, payload) {
  try {
    const backendLineId = await ensureBackendLineId(payload.lineId, session);
    if (!backendLineId) return;
    const backendEquipmentId = payload.equipment;
    await apiRequest("/api/logs/downtime", {
      method: "POST",
      token: session.backendToken,
      body: {
        lineId: backendLineId,
        date: payload.date,
        shift: payload.shift,
        downtimeStart: payload.downtimeStart,
        downtimeFinish: payload.downtimeFinish,
        equipmentStageId: UUID_RE.test(String(backendEquipmentId || "")) ? backendEquipmentId : null,
        reason: payload.reason || ""
      }
    });
  } catch (error) {
    console.warn("Backend downtime sync failed:", error);
  }
}

function requiredCrewForLineShift(line, shift) {
  const stages = line?.stages?.length ? line.stages : STAGES;
  return stages.reduce((sum, stage) => sum + Math.max(0, num(line?.crewsByShift?.[shift]?.[stage.id]?.crew)), 0);
}

function nowIso() {
  return new Date().toISOString();
}

function currentActor() {
  if (appState.appMode === "supervisor" && appState.supervisorSession?.username) {
    const sup = supervisorByUsername(appState.supervisorSession.username);
    return { actor: sup?.name || appState.supervisorSession.username, actorType: "supervisor" };
  }
  return { actor: "manager", actorType: "manager" };
}

function ensureAuditRows(line) {
  if (!Array.isArray(line.auditRows)) line.auditRows = [];
  return line.auditRows;
}

function addAudit(line, action, details, actorOverride = null) {
  if (!line) return;
  const actorData = actorOverride || currentActor();
  const rows = ensureAuditRows(line);
  rows.unshift({
    id: `audit-${Date.now()}-${Math.floor(Math.random() * 100000)}`,
    at: nowIso(),
    actor: actorData.actor,
    actorType: actorData.actorType,
    action,
    details: String(details || "")
  });
  if (rows.length > 800) rows.length = 800;
}

function stageMaxThroughput(stageId) {
  return Math.max(0, num(state.stageSettings?.[stageId]?.maxThroughput));
}

function stageDisplayName(stage, index) {
  const base = String(stage?.name || `Stage ${index + 1}`).replace(/^\s*\d+\.\s*/, "").trim();
  return `${index + 1}. ${base || "Stage"}`;
}

function stageCrewForShift(stageId, shift) {
  const crew = num(state.crewsByShift?.[shift]?.[stageId]?.crew);
  const stage = getStages().find((s) => s.id === stageId);
  if (crew > 0) return crew;
  if (stage?.kind === "transfer") return 1;
  return 0;
}

function stageTotalMaxThroughput(stageId, shift) {
  return stageMaxThroughput(stageId) * stageCrewForShift(stageId, shift);
}

function stageNameById(id) {
  const stages = getStages();
  const idx = stages.findIndex((stage) => stage.id === id);
  if (idx === -1) return id || "";
  return stageDisplayName(stages[idx], idx);
}

function stageOptionListHTML() {
  return [
    `<option value="">Equipment Stage</option>`,
    ...getStages().map((stage, index) => `<option value="${stage.id}">${stageDisplayName(stage, index)}</option>`)
  ].join("");
}

function shiftKey(date, shift) {
  return `${date}__${shift}`;
}

function monthKey(isoDate) {
  return String(isoDate || "").slice(0, 7);
}

function addMonths(month, delta) {
  const [y, m] = month.split("-").map(Number);
  const dt = new Date(y, (m || 1) - 1, 1);
  dt.setMonth(dt.getMonth() + delta);
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}`;
}

function formatMonthLabel(month) {
  const [y, m] = month.split("-").map(Number);
  if (!y || !m) return month;
  return new Date(y, m - 1, 1).toLocaleDateString(undefined, { month: "long", year: "numeric" });
}

function sampleDataSet() {
  const shiftRows = [];
  const runRows = [];
  const downtimeRows = [];
  const start = new Date(2025, 10, 1);
  const days = 100;
  const stageIds = getStages().map((stage) => stage.id).filter(Boolean);
  const dayEquipment = stageIds.filter((_, idx) => idx % 2 === 0);
  const nightEquipment = stageIds.filter((_, idx) => idx % 2 === 1);
  const fallbackEquipment = stageIds[0] || "";
  const equipAt = (list, i, offset = 0) => (list.length ? list[(i + offset) % list.length] : fallbackEquipment);
  const dayRequired = requiredCrewForLineShift(state, "Day");
  const nightRequired = requiredCrewForLineShift(state, "Night");

  for (let i = 0; i < days; i += 1) {
    const d = new Date(start);
    d.setDate(start.getDate() + i);
    const date = formatDateLocal(d);
    const dayTrend = 0.95 + ((i % 7) - 3) * 0.01;
    const nightTrend = 0.92 + (((i + 2) % 7) - 3) * 0.01;

    shiftRows.push(
      {
        date,
        shift: "Day",
        crewOnShift: Math.max(0, dayRequired - (i % 12 === 0 ? 2 : i % 7 === 0 ? 1 : 0)),
        startTime: "06:00",
        break1Start: "09:00",
        break2Start: "12:00",
        break3Start: "14:00",
        finishTime: "14:00"
      },
      {
        date,
        shift: "Night",
        crewOnShift: Math.max(0, nightRequired - (i % 10 === 0 ? 1 : 0)),
        startTime: "14:00",
        break1Start: "17:00",
        break2Start: "20:00",
        break3Start: "22:00",
        finishTime: "22:00"
      }
    );

    runRows.push(
      {
        date,
        shift: "Day",
        product: "Teriyaki",
        setUpStartTime: "05:40",
        productionStartTime: "06:10",
        finishTime: "10:35",
        unitsProduced: Math.round(2850 * dayTrend)
      },
      {
        date,
        shift: "Day",
        product: "Honey Soy",
        setUpStartTime: "10:40",
        productionStartTime: "10:55",
        finishTime: "15:25",
        unitsProduced: Math.round(2600 * dayTrend)
      },
      {
        date,
        shift: "Night",
        product: "Peri Peri",
        setUpStartTime: "13:55",
        productionStartTime: "14:15",
        finishTime: "18:55",
        unitsProduced: Math.round(2500 * nightTrend)
      },
      {
        date,
        shift: "Night",
        product: "Lemon Herb",
        setUpStartTime: "21:10",
        productionStartTime: "21:30",
        finishTime: "00:25",
        unitsProduced: Math.round(2200 * nightTrend)
      }
    );

    const deq = equipAt(dayEquipment, i);
    const neq = equipAt(nightEquipment, i);
    downtimeRows.push(
      { date, shift: "Day", downtimeStart: "08:10", downtimeFinish: `08:${String(22 + (i % 8)).padStart(2, "0")}`, equipment: deq, reason: "Planned maintenance" },
      { date, shift: "Day", downtimeStart: "11:20", downtimeFinish: `11:${String(30 + (i % 10)).padStart(2, "0")}`, equipment: equipAt(dayEquipment, i, 2), reason: "Minor stoppage" },
      { date, shift: "Night", downtimeStart: "16:30", downtimeFinish: `16:${String(42 + (i % 9)).padStart(2, "0")}`, equipment: neq, reason: "Sensor reset" },
      { date, shift: "Night", downtimeStart: "20:05", downtimeFinish: `20:${String(18 + (i % 11)).padStart(2, "0")}`, equipment: equipAt(nightEquipment, i, 3), reason: "Label adjustment" }
    );
  }

  return {
    selectedDate: "2026-02-08",
    selectedShift: "Day",
    trendGranularity: "daily",
    trendMonth: "2026-02",
    shiftRows,
    runRows,
    downtimeRows
  };
}

function parseTimeToDayFraction(value) {
  if (value === null || value === undefined || value === "") return null;
  const raw = String(value).trim();
  if (!raw) return null;

  const numeric = Number(raw);
  if (Number.isFinite(numeric)) {
    if (numeric <= 1) return numeric;
    const minsAsFraction = numeric / (24 * 60);
    return minsAsFraction <= 1 ? minsAsFraction : null;
  }

  const timeMatch = raw.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!timeMatch) return null;
  const hh = Number(timeMatch[1]);
  const mm = Number(timeMatch[2]);
  const ss = Number(timeMatch[3] || 0);
  if (hh > 23 || mm > 59 || ss > 59) return null;
  return (hh * 3600 + mm * 60 + ss) / 86400;
}

function diffMinutes(startValue, finishValue) {
  const start = parseTimeToDayFraction(startValue);
  const finish = parseTimeToDayFraction(finishValue);
  if (start === null || finish === null) return 0;
  const deltaDays = finish >= start ? finish - start : finish + 1 - start;
  return deltaDays * 24 * 60;
}

function computeDowntimeRow(row) {
  const fallback = num(row.downtimeMins);
  const calc = diffMinutes(row.downtimeStart, row.downtimeFinish);
  return { ...row, downtimeMins: calc > 0 ? calc : fallback };
}

function computeShiftRow(row) {
  const fallback = num(row.totalShiftTime);
  const calc = diffMinutes(row.startTime, row.finishTime);
  return { ...row, crewOnShift: Math.max(0, num(row.crewOnShift)), totalShiftTime: calc > 0 ? calc : fallback };
}

function computeRunRow(row, downtimeByShift) {
  const grossFallback = num(row.grossProductionTime);
  const grossCalc = diffMinutes(row.productionStartTime, row.finishTime);
  const grossProductionTime = grossCalc > 0 ? grossCalc : grossFallback;

  const associatedFallback = num(row.associatedDownTime);
  const associatedFromDowntime = downtimeByShift.get(`${row.date}__${row.shift}`) ?? 0;
  const associatedDownTime = associatedFromDowntime > 0 ? associatedFromDowntime : associatedFallback;

  const netFallback = num(row.netProductionTime);
  const netCalc = Math.max(0, grossProductionTime - associatedDownTime);
  const netProductionTime = netCalc > 0 ? netCalc : netFallback;

  const unitsProduced = num(row.unitsProduced);
  const grossRunRate = grossProductionTime > 0 ? unitsProduced / grossProductionTime : num(row.grossRunRate);
  const netRunRate = netProductionTime > 0 ? unitsProduced / netProductionTime : num(row.netRunRate);

  return {
    ...row,
    grossProductionTime,
    associatedDownTime,
    netProductionTime,
    grossRunRate,
    netRunRate
  };
}

function derivedData() {
  const downtimeRows = state.downtimeRows.map(computeDowntimeRow);
  const shiftRows = state.shiftRows.map(computeShiftRow);
  const downtimeByShift = new Map();
  downtimeRows.forEach((row) => {
    const key = `${row.date}__${row.shift}`;
    downtimeByShift.set(key, (downtimeByShift.get(key) || 0) + num(row.downtimeMins));
  });
  const runRows = state.runRows.map((row) => computeRunRow(row, downtimeByShift));
  return { shiftRows, runRows, downtimeRows };
}

function clearLineTrackingData(line) {
  line.shiftRows = [];
  line.runRows = [];
  line.downtimeRows = [];
}

function matchesStage(stage, equipmentText) {
  const input = String(equipmentText || "").toLowerCase();
  if (input === stage.id.toLowerCase()) return true;
  if (input === stage.name.toLowerCase()) return true;
  return stage.match.some((token) => input.includes(token));
}

function statusClass(utilisation) {
  if (utilisation >= 85) return "good";
  if (utilisation >= 65) return "warn";
  return "bad";
}

function bindTabs() {
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach((b) => {
        b.classList.remove("active");
        b.setAttribute("aria-selected", "false");
      });
      document.querySelectorAll(".tab-panel").forEach((panel) => panel.classList.remove("active"));
      btn.classList.add("active");
      btn.setAttribute("aria-selected", "true");
      document.getElementById(btn.dataset.tab).classList.add("active");
    });
  });
}

function bindHome() {
  const goHomeBtn = document.getElementById("goHome");
  const modeManagerBtn = document.getElementById("modeManager");
  const modeSupervisorBtn = document.getElementById("modeSupervisor");
  const sidebarToggleBtn = document.getElementById("sidebarToggle");
  const sidebarBackdrop = document.getElementById("sidebarBackdrop");
  const homeSidebar = document.getElementById("homeSidebar");
  const dashboardDate = document.getElementById("dashboardDate");
  const dashboardShiftButtons = Array.from(document.querySelectorAll("[data-dash-shift]"));
  const exportDashboardCsvBtn = document.getElementById("exportDashboardCsv");
  const supervisorLoginForm = document.getElementById("supervisorLoginForm");
  const supervisorLogoutBtn = document.getElementById("supervisorLogout");
  const supervisorMobileModeBtn = document.getElementById("supervisorMobileMode");
  const supervisorLineSelect = document.getElementById("supervisorLineSelect");
  const svDateInput = document.getElementById("svSelectedDate");
  const svShiftButtons = Array.from(document.querySelectorAll("[data-sv-shift]"));
  const svPrevBtn = document.getElementById("svPrevDay");
  const svNextBtn = document.getElementById("svNextDay");
  const supervisorShiftForm = document.getElementById("supervisorShiftForm");
  const supervisorRunForm = document.getElementById("supervisorRunForm");
  const supervisorDownForm = document.getElementById("supervisorDownForm");
  const supervisorDownEquipment = document.getElementById("superDownEquipment");
  const manageSupervisorsBtn = document.getElementById("manageSupervisorsBtn");
  const addSupervisorBtn = document.getElementById("addSupervisorBtn");
  const manageSupervisorsModal = document.getElementById("manageSupervisorsModal");
  const closeManageSupervisorsModalBtn = document.getElementById("closeManageSupervisorsModal");
  const supervisorManagerList = document.getElementById("supervisorManagerList");
  const addSupervisorModal = document.getElementById("addSupervisorModal");
  const closeAddSupervisorModalBtn = document.getElementById("closeAddSupervisorModal");
  const addSupervisorForm = document.getElementById("addSupervisorForm");
  const newSupervisorLines = document.getElementById("newSupervisorLines");
  const cards = document.getElementById("lineCards");
  const openBuilderBtn = document.getElementById("openBuilder");
  const openBuilderSecondaryBtn = document.getElementById("openBuilderSecondary");
  const builderModal = document.getElementById("builderModal");
  const closeBuilderBtn = document.getElementById("closeBuilderModal");
  const addBuilderStageBtn = document.getElementById("addBuilderStage");
  const createBuilderLineBtn = document.getElementById("createBuilderLine");
  const builderLineNameInput = document.getElementById("builderLineName");
  const builderStages = document.getElementById("builderStages");
  const builderCanvas = document.getElementById("builderCanvas");
  let builderDraft = null;
  let dragState = null;

  const stripLeadingStageNumber = (name) => String(name || "").replace(/^\s*\d+\.\s*/, "").trim();
  const seedBuilderFromTemplate = () => ({
    lineName: "",
    stages: [
      {
        uid: "tmp-1",
        name: "Stage",
        group: "main",
        crew: 1,
        maxThroughput: 2,
        x: 8,
        y: 14,
        w: 9.5,
        h: 15
      }
    ]
  });

  const clamp = (value, min, max) => Math.max(min, Math.min(max, value));
  const renderBuilder = () => {
    if (!builderDraft) return;
    builderLineNameInput.value = builderDraft.lineName;

    builderStages.innerHTML = builderDraft.stages
      .map(
        (stage, i) => `
          <div class="builder-row" data-builder-row="${stage.uid}">
            <span class="builder-index">${i + 1}.</span>
            <input data-builder="name" data-uid="${stage.uid}" value="${stage.name}" placeholder="Stage name" />
            <select data-builder="group" data-uid="${stage.uid}">
              <option value="main" ${stage.group === "main" ? "selected" : ""}>Main</option>
              <option value="prep" ${stage.group === "prep" ? "selected" : ""}>Prep</option>
              <option value="transfer" ${stage.group === "transfer" ? "selected" : ""}>Transfer</option>
            </select>
            <input data-builder="crew" data-uid="${stage.uid}" type="number" min="0" step="1" value="${num(stage.crew)}" />
            <input data-builder="throughput" data-uid="${stage.uid}" type="number" min="0" step="1" value="${num(stage.maxThroughput)}" />
            <button type="button" class="danger" data-remove-builder="${stage.uid}">Del</button>
          </div>
        `
      )
      .join("");

    builderCanvas.innerHTML = builderDraft.stages
      .map(
        (stage) => `
          <div class="builder-stage-chip ${stage.group === "transfer" ? "transfer" : ""}" data-canvas-stage="${stage.uid}"
            style="left:${stage.x}%;top:${stage.y}%;width:${stage.w}%;height:${stage.h}%;">
            <span>${stage.name}</span>
            <span class="builder-resize-handle" data-resize-stage="${stage.uid}"></span>
          </div>
        `
      )
      .join("");
  };

  const openBuilderModal = () => {
    builderDraft = seedBuilderFromTemplate();
    renderBuilder();
    builderModal.classList.add("open");
    builderModal.setAttribute("aria-hidden", "false");
  };

  const closeBuilderModal = () => {
    builderModal.classList.remove("open");
    builderModal.setAttribute("aria-hidden", "true");
    builderDraft = null;
    dragState = null;
  };

  const renderSupervisorLineChecklist = (selectedIds = []) => {
    const lineIds = Object.keys(appState.lines);
    newSupervisorLines.innerHTML = lineIds
      .map((id) => {
        const line = appState.lines[id];
        return `
          <label class="supervisor-line-item">
            <input type="checkbox" value="${id}" ${selectedIds.includes(id) ? "checked" : ""} />
            <span>${line.name}</span>
          </label>
        `;
      })
      .join("");
  };

  const renderSupervisorManagerList = () => {
    const lineIds = Object.keys(appState.lines);
    const linesById = appState.lines;
    supervisorManagerList.innerHTML = (appState.supervisors || [])
      .map((sup) => {
        const options = lineIds
          .map(
            (id) =>
              `<option value="${id}" ${sup.assignedLineIds.includes(id) ? "selected" : ""}>${linesById[id].name}</option>`
          )
          .join("");
        return `
          <section class="panel supervisor-manager-row" data-supervisor-id="${sup.id}">
            <div class="action-row">
              <h3>${sup.name}</h3>
              <span class="muted">@${sup.username}</span>
            </div>
            <label>
              Assigned Lines
              <select multiple size="${Math.min(8, Math.max(3, lineIds.length))}" data-supervisor-lines="${sup.id}">
                ${options}
              </select>
            </label>
            <div class="action-row">
              <button type="button" data-supervisor-save="${sup.id}">Save Assignments</button>
              <button type="button" class="danger" data-supervisor-delete="${sup.id}">Delete Supervisor</button>
            </div>
          </section>
        `;
      })
      .join("");
    if (!appState.supervisors?.length) {
      supervisorManagerList.innerHTML = `<p class="muted">No supervisors created yet.</p>`;
    }
  };

  const openManageSupervisorsModal = () => {
    renderSupervisorManagerList();
    manageSupervisorsModal.classList.add("open");
    manageSupervisorsModal.setAttribute("aria-hidden", "false");
  };

  const closeManageSupervisorsModal = () => {
    manageSupervisorsModal.classList.remove("open");
    manageSupervisorsModal.setAttribute("aria-hidden", "true");
  };

  const openAddSupervisorModal = () => {
    addSupervisorForm.reset();
    renderSupervisorLineChecklist(Object.keys(appState.lines));
    addSupervisorModal.classList.add("open");
    addSupervisorModal.setAttribute("aria-hidden", "false");
  };

  const closeAddSupervisorModal = () => {
    addSupervisorModal.classList.remove("open");
    addSupervisorModal.setAttribute("aria-hidden", "true");
  };

  goHomeBtn.addEventListener("click", () => {
    appState.activeView = "home";
    saveState();
    renderAll();
  });

  modeManagerBtn.addEventListener("click", () => {
    appState.appMode = "manager";
    saveState();
    renderAll();
  });

  modeSupervisorBtn.addEventListener("click", () => {
    appState.appMode = "supervisor";
    saveState();
    renderAll();
  });

  if (sidebarToggleBtn && sidebarBackdrop && homeSidebar) {
    sidebarToggleBtn.addEventListener("click", () => {
      homeSidebar.classList.toggle("open");
      sidebarBackdrop.classList.toggle("hidden", !homeSidebar.classList.contains("open"));
    });
    sidebarBackdrop.addEventListener("click", () => {
      homeSidebar.classList.remove("open");
      sidebarBackdrop.classList.add("hidden");
    });
  }

  dashboardDate.addEventListener("change", () => {
    appState.dashboardDate = dashboardDate.value || todayISO();
    saveState();
    renderHome();
  });

  dashboardShiftButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const shift = btn.dataset.dashShift;
      if (!shift) return;
      appState.dashboardShift = shift === "Night" ? "Night" : "Day";
      saveState();
      renderHome();
    });
  });

  exportDashboardCsvBtn.addEventListener("click", () => {
    const rows = Object.values(appState.lines || {}).map((line) =>
      computeLineMetrics(line, appState.dashboardDate || todayISO(), appState.dashboardShift || "Day")
    );
    const csvRows = rows.map((row) => ({
      Line: row.lineName,
      Date: row.date,
      Shift: row.shift,
      Units: Number(row.units || 0).toFixed(2),
      "Downtime (min)": Number(row.totalDowntime || 0).toFixed(2),
      "Utilisation (%)": Number(row.lineUtil || 0).toFixed(2),
      "Net Run Rate (u/min)": Number(row.netRunRate || 0).toFixed(2),
      Bottleneck: row.bottleneckStageName,
      Staffing: row.staffingCallout
    }));
    downloadTextFile(`line-dashboard-${appState.dashboardDate || todayISO()}-${appState.dashboardShift || "Day"}.csv`, toCsv(csvRows, DASHBOARD_COLUMNS), "text/csv;charset=utf-8");
  });

  supervisorLoginForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const username = String(document.getElementById("supervisorUser").value || "").trim().toLowerCase();
    const password = String(document.getElementById("supervisorPass").value || "");
    const sup = supervisorByUsername(username);
    let loggedInSession = null;

    try {
      const loginPayload = await apiRequest("/api/auth/login", {
        method: "POST",
        body: { username, password }
      });
      if (loginPayload?.user?.role === "supervisor" && loginPayload?.token) {
        const linesPayload = await apiRequest("/api/lines", { token: loginPayload.token });
        const backendLines = Array.isArray(linesPayload?.lines) ? linesPayload.lines : [];
        const backendLineMap = {};
        const assignedLocalIds = [];
        backendLines.forEach((line) => {
          if (!line?.id || !line?.name) return;
          const localId = findLocalLineIdByName(line.name);
          if (!localId) return;
          backendLineMap[localId] = line.id;
          assignedLocalIds.push(localId);
        });
        const fallbackLocal = sup ? sup.assignedLineIds.filter((id) => appState.lines[id]) : [];
        loggedInSession = {
          username: sup?.username || username,
          assignedLineIds: assignedLocalIds.length ? assignedLocalIds : fallbackLocal,
          backendToken: loginPayload.token,
          backendLineMap
        };
      }
    } catch {
      // Keep local supervisor login fallback for migration period.
    }

    if (!loggedInSession) {
      if (!username || !sup || sup.password !== password) {
        alert("Invalid supervisor credentials.");
        return;
      }
      loggedInSession = {
        username: sup.username,
        assignedLineIds: sup.assignedLineIds.filter((id) => appState.lines[id]),
        backendToken: "",
        backendLineMap: {}
      };
    }

    appState.supervisorSession = loggedInSession;
    saveState();
    renderAll();
  });

  supervisorLogoutBtn.addEventListener("click", () => {
    appState.supervisorSession = null;
    saveState();
    renderAll();
  });

  supervisorMobileModeBtn.addEventListener("click", () => {
    appState.supervisorMobileMode = !appState.supervisorMobileMode;
    saveState();
    renderHome();
  });

  document.querySelectorAll("[data-supervisor-tab]").forEach((btn) => {
    btn.addEventListener("click", () => {
      appState.supervisorTab = btn.dataset.supervisorTab || "superShift";
      saveState();
      renderHome();
    });
  });

  supervisorLineSelect.addEventListener("change", () => {
    appState.supervisorSelectedLineId = supervisorLineSelect.value || "";
    supervisorDownEquipment.innerHTML = supervisorEquipmentOptions(selectedSupervisorLine());
    saveState();
    renderHome();
  });

  svDateInput.addEventListener("change", () => {
    appState.supervisorSelectedDate = svDateInput.value || todayISO();
    saveState();
    renderHome();
  });

  svShiftButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const shift = btn.dataset.svShift;
      if (!shift) return;
      appState.supervisorSelectedShift = shift === "Night" ? "Night" : "Day";
      saveState();
      renderHome();
    });
  });

  svPrevBtn.addEventListener("click", () => {
    const dt = parseDateLocal(appState.supervisorSelectedDate || todayISO());
    dt.setDate(dt.getDate() - 1);
    appState.supervisorSelectedDate = formatDateLocal(dt);
    saveState();
    renderHome();
  });

  svNextBtn.addEventListener("click", () => {
    const dt = parseDateLocal(appState.supervisorSelectedDate || todayISO());
    dt.setDate(dt.getDate() + 1);
    appState.supervisorSelectedDate = formatDateLocal(dt);
    saveState();
    renderHome();
  });

  document.querySelectorAll("[data-super-main-tab]").forEach((btn) => {
    btn.addEventListener("click", () => {
      appState.supervisorMainTab = btn.dataset.superMainTab || "supervisorVisual";
      saveState();
      renderHome();
    });
  });

  manageSupervisorsBtn.addEventListener("click", openManageSupervisorsModal);
  closeManageSupervisorsModalBtn.addEventListener("click", closeManageSupervisorsModal);
  manageSupervisorsModal.addEventListener("click", (event) => {
    if (event.target === manageSupervisorsModal) closeManageSupervisorsModal();
  });

  addSupervisorBtn.addEventListener("click", openAddSupervisorModal);
  closeAddSupervisorModalBtn.addEventListener("click", closeAddSupervisorModal);
  addSupervisorModal.addEventListener("click", (event) => {
    if (event.target === addSupervisorModal) closeAddSupervisorModal();
  });

  supervisorManagerList.addEventListener("click", (event) => {
    const saveBtn = event.target.closest("[data-supervisor-save]");
    if (saveBtn) {
      const supId = saveBtn.getAttribute("data-supervisor-save");
      const sup = (appState.supervisors || []).find((item) => item.id === supId);
      const select = supervisorManagerList.querySelector(`[data-supervisor-lines="${supId}"]`);
      if (!sup || !select) return;
      const prevAssigned = Array.isArray(sup.assignedLineIds) ? sup.assignedLineIds.slice() : [];
      sup.assignedLineIds = Array.from(select.selectedOptions).map((opt) => opt.value).filter((id) => appState.lines[id]);
      const added = sup.assignedLineIds.filter((id) => !prevAssigned.includes(id));
      const removed = prevAssigned.filter((id) => !sup.assignedLineIds.includes(id));
      added.forEach((lineId) => addAudit(appState.lines[lineId], "ASSIGN_SUPERVISOR", `${sup.name} assigned to line`));
      removed.forEach((lineId) => addAudit(appState.lines[lineId], "UNASSIGN_SUPERVISOR", `${sup.name} removed from line`));
      if (appState.supervisorSession?.username === sup.username) {
        appState.supervisorSession.assignedLineIds = sup.assignedLineIds.slice();
      }
      saveState();
      renderHome();
      openManageSupervisorsModal();
      return;
    }

    const delBtn = event.target.closest("[data-supervisor-delete]");
    if (delBtn) {
      const supId = delBtn.getAttribute("data-supervisor-delete");
      const sup = (appState.supervisors || []).find((item) => item.id === supId);
      if (!sup) return;
      if (!window.confirm(`Delete supervisor "${sup.name}"?`)) return;
      Object.values(appState.lines || {}).forEach((line) => {
        if ((sup.assignedLineIds || []).includes(line.id)) {
          addAudit(line, "DELETE_SUPERVISOR", `Supervisor ${sup.name} deleted`);
        }
      });
      appState.supervisors = (appState.supervisors || []).filter((item) => item.id !== supId);
      if (appState.supervisorSession?.username === sup.username) appState.supervisorSession = null;
      saveState();
      renderHome();
      openManageSupervisorsModal();
    }
  });

  addSupervisorForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const name = String(document.getElementById("newSupervisorName").value || "").trim();
    const username = String(document.getElementById("newSupervisorUsername").value || "").trim().toLowerCase();
    const password = String(document.getElementById("newSupervisorPassword").value || "").trim();
    const assignedLineIds = Array.from(newSupervisorLines.querySelectorAll("input[type=\"checkbox\"]:checked"))
      .map((input) => input.value)
      .filter((id) => appState.lines[id]);
    if (!name || !username || !password) return;
    if (supervisorByUsername(username)) {
      alert("Username already exists.");
      return;
    }
    appState.supervisors = Array.isArray(appState.supervisors) ? appState.supervisors : [];
    appState.supervisors.push({
      id: `sup-${Date.now()}`,
      name,
      username,
      password,
      assignedLineIds
    });
    assignedLineIds.forEach((lineId) => addAudit(appState.lines[lineId], "CREATE_SUPERVISOR", `Supervisor ${name} created and assigned`));
    saveState();
    closeAddSupervisorModal();
    renderAll();
  });

  supervisorShiftForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const session = appState.supervisorSession;
    if (!session) return;
    const lineId = selectedSupervisorLineId();
    if (!session.assignedLineIds.includes(lineId) || !appState.lines[lineId]) {
      alert("You are not assigned to that line.");
      return;
    }
    const date = document.getElementById("superShiftDate").value || todayISO();
    const shift = document.getElementById("superShiftShift").value || "Day";
    const crewOnShift = Math.max(0, Math.floor(num(document.getElementById("superShiftCrew").value || 0)));
    const startTime = document.getElementById("superShiftStart").value || "";
    const break1Start = document.getElementById("superShiftBreak1").value || "";
    const break2Start = document.getElementById("superShiftBreak2").value || "";
    const break3Start = document.getElementById("superShiftBreak3").value || "";
    const finishTime = document.getElementById("superShiftFinish").value || "";
    if (!rowIsValidDateShift(date, shift)) {
      alert("Date/shift are invalid.");
      return;
    }
    if (!strictTimeValid(startTime) || !strictTimeValid(finishTime)) {
      alert("Shift start and finish must be in HH:MM (24h).");
      return;
    }
    if (!optionalStrictTimeValid(break1Start) || !optionalStrictTimeValid(break2Start) || !optionalStrictTimeValid(break3Start)) {
      alert("Break times must be HH:MM.");
      return;
    }
    if (crewOnShift < 0) {
      alert("Crew on shift cannot be negative.");
      return;
    }
    const line = appState.lines[lineId];
    line.shiftRows.push({
      date,
      shift,
      crewOnShift,
      startTime,
      break1Start,
      break2Start,
      break3Start,
      finishTime,
      submittedBy: session.username,
      submittedAt: new Date().toISOString()
    });
    addAudit(line, "SUPERVISOR_SHIFT_LOG", `${session.username} logged ${shift} shift for ${date} (crew ${crewOnShift})`);
    await syncSupervisorShiftLog(session, {
      lineId,
      date,
      shift,
      crewOnShift,
      startTime,
      break1Start,
      break2Start,
      break3Start,
      finishTime
    });
    event.currentTarget.reset();
    document.getElementById("superShiftDate").value = date;
    saveState();
    renderAll();
  });

  supervisorRunForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const session = appState.supervisorSession;
    if (!session) return;
    const lineId = selectedSupervisorLineId();
    if (!session.assignedLineIds.includes(lineId) || !appState.lines[lineId]) {
      alert("You are not assigned to that line.");
      return;
    }
    const date = document.getElementById("superRunDate").value || todayISO();
    const shift = document.getElementById("superRunShift").value || "Day";
    const product = String(document.getElementById("superRunProduct").value || "").trim();
    const setUpStartTime = document.getElementById("superRunSetupStart").value || "";
    const productionStartTime = document.getElementById("superRunProdStart").value || "";
    const finishTime = document.getElementById("superRunFinish").value || "";
    const unitsProduced = num(document.getElementById("superRunUnits").value || 0);
    if (!rowIsValidDateShift(date, shift)) {
      alert("Date/shift are invalid.");
      return;
    }
    if (!product) return;
    if (!optionalStrictTimeValid(setUpStartTime) || !strictTimeValid(productionStartTime) || !strictTimeValid(finishTime)) {
      alert("Production start and finish must be HH:MM (24h).");
      return;
    }
    if (unitsProduced < 0) {
      alert("Units produced cannot be negative.");
      return;
    }
    if (!product) return;
    const line = appState.lines[lineId];
    line.runRows.push({
      date,
      shift,
      product,
      setUpStartTime,
      productionStartTime,
      finishTime,
      unitsProduced,
      submittedBy: session.username,
      submittedAt: new Date().toISOString()
    });
    addAudit(line, "SUPERVISOR_RUN_LOG", `${session.username} logged run ${product} (${unitsProduced} units)`);
    await syncSupervisorRunLog(session, {
      lineId,
      date,
      shift,
      product,
      setUpStartTime,
      productionStartTime,
      finishTime,
      unitsProduced
    });
    event.currentTarget.reset();
    document.getElementById("superRunDate").value = date;
    saveState();
    renderAll();
  });

  supervisorDownForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const session = appState.supervisorSession;
    if (!session) return;
    const lineId = selectedSupervisorLineId();
    if (!session.assignedLineIds.includes(lineId) || !appState.lines[lineId]) {
      alert("You are not assigned to that line.");
      return;
    }
    const date = document.getElementById("superDownDate").value || todayISO();
    const shift = document.getElementById("superDownShift").value || "Day";
    const startTime = document.getElementById("superDownStart").value || "";
    const finishTime = document.getElementById("superDownFinish").value || "";
    const equipment = document.getElementById("superDownEquipment").value || "";
    const reason = String(document.getElementById("superDownReason").value || "").trim();
    if (!rowIsValidDateShift(date, shift)) {
      alert("Date/shift are invalid.");
      return;
    }
    if (!strictTimeValid(startTime) || !strictTimeValid(finishTime)) {
      alert("Downtime start and finish must be HH:MM (24h).");
      return;
    }
    if (!equipment) {
      alert("Select equipment stage.");
      return;
    }
    const line = appState.lines[lineId];
    line.downtimeRows.push({
      date,
      shift,
      downtimeStart: startTime,
      downtimeFinish: finishTime,
      equipment,
      reason,
      submittedBy: session.username,
      submittedAt: new Date().toISOString()
    });
    addAudit(line, "SUPERVISOR_DOWNTIME_LOG", `${session.username} logged downtime on ${stageNameById(equipment)}`);
    await syncSupervisorDowntimeLog(session, {
      lineId,
      date,
      shift,
      downtimeStart: startTime,
      downtimeFinish: finishTime,
      equipment,
      reason
    });
    event.currentTarget.reset();
    document.getElementById("superDownDate").value = date;
    supervisorDownEquipment.innerHTML = supervisorEquipmentOptions(selectedSupervisorLine());
    saveState();
    renderAll();
  });

  openBuilderBtn.addEventListener("click", openBuilderModal);
  if (openBuilderSecondaryBtn) openBuilderSecondaryBtn.addEventListener("click", openBuilderModal);
  closeBuilderBtn.addEventListener("click", closeBuilderModal);
  builderModal.addEventListener("click", (event) => {
    if (event.target === builderModal) closeBuilderModal();
  });

  builderLineNameInput.addEventListener("input", () => {
    if (!builderDraft) return;
    builderDraft.lineName = builderLineNameInput.value;
  });

  addBuilderStageBtn.addEventListener("click", () => {
    if (!builderDraft) return;
    const next = builderDraft.stages.length + 1;
    builderDraft.stages.push({
      uid: `tmp-${Date.now()}-${next}`,
      name: "Stage",
      group: "main",
      crew: 1,
      maxThroughput: 2,
      x: 8 + (next % 8) * 10,
      y: next % 2 ? 14 : 52,
      w: 9.5,
      h: 15
    });
    renderBuilder();
  });

  builderStages.addEventListener("input", (event) => {
    if (!builderDraft) return;
    const target = event.target;
    const uid = target.getAttribute("data-uid");
    const field = target.getAttribute("data-builder");
    const stage = builderDraft.stages.find((s) => s.uid === uid);
    if (!stage || !field) return;
    if (field === "name") {
      stage.name = stripLeadingStageNumber(target.value) || "Stage";
      const chipLabel = builderCanvas.querySelector(`[data-canvas-stage="${uid}"] > span`);
      if (chipLabel) chipLabel.textContent = stage.name;
      return;
    }
    if (field === "crew" || field === "throughput") {
      stage[field === "crew" ? "crew" : "maxThroughput"] = num(target.value);
      return;
    }
  });

  builderStages.addEventListener("change", (event) => {
    if (!builderDraft) return;
    const target = event.target;
    const uid = target.getAttribute("data-uid");
    const field = target.getAttribute("data-builder");
    if (field !== "group") return;
    const stage = builderDraft.stages.find((s) => s.uid === uid);
    if (!stage) return;
    stage.group = target.value || "main";
    if (stage.group === "transfer") {
      stage.w = 7.5;
      stage.h = 10;
    } else {
      stage.w = 9.5;
      stage.h = 15;
    }
    renderBuilder();
  });

  builderStages.addEventListener("click", (event) => {
    if (!builderDraft) return;
    const removeBtn = event.target.closest("[data-remove-builder]");
    if (!removeBtn) return;
    const uid = removeBtn.getAttribute("data-remove-builder");
    builderDraft.stages = builderDraft.stages.filter((s) => s.uid !== uid);
    if (!builderDraft.stages.length) {
      builderDraft.stages.push({ uid: `tmp-${Date.now()}`, name: "Stage", group: "main", crew: 1, maxThroughput: 2, x: 8, y: 14, w: 9.5, h: 15 });
    }
    renderBuilder();
  });

  builderCanvas.addEventListener("mousedown", (event) => {
    if (!builderDraft) return;
    const resizeNode = event.target.closest("[data-resize-stage]");
    if (resizeNode) {
      const uid = resizeNode.getAttribute("data-resize-stage");
      const stage = builderDraft.stages.find((s) => s.uid === uid);
      if (!stage) return;
      const rect = builderCanvas.getBoundingClientRect();
      dragState = {
        mode: "resize",
        uid,
        startX: event.clientX,
        startY: event.clientY,
        startW: stage.w,
        startH: stage.h,
        canvasW: rect.width,
        canvasH: rect.height
      };
      event.preventDefault();
      return;
    }

    const node = event.target.closest("[data-canvas-stage]");
    if (!node) return;
    const uid = node.getAttribute("data-canvas-stage");
    const stage = builderDraft.stages.find((s) => s.uid === uid);
    if (!stage) return;
    const rect = builderCanvas.getBoundingClientRect();
    dragState = {
      mode: "move",
      uid,
      offsetX: event.clientX - rect.left - (stage.x / 100) * rect.width,
      offsetY: event.clientY - rect.top - (stage.y / 100) * rect.height
    };
    event.preventDefault();
  });

  window.addEventListener("mousemove", (event) => {
    if (!builderDraft || !dragState) return;
    const stage = builderDraft.stages.find((s) => s.uid === dragState.uid);
    if (!stage) return;
    const rect = builderCanvas.getBoundingClientRect();
    if (dragState.mode === "resize") {
      const deltaXPercent = ((event.clientX - dragState.startX) / dragState.canvasW) * 100;
      const deltaYPercent = ((event.clientY - dragState.startY) / dragState.canvasH) * 100;
      const minW = stage.group === "transfer" ? 5.5 : 6.5;
      const minH = stage.group === "transfer" ? 7.5 : 10;
      stage.w = clamp(dragState.startW + deltaXPercent, minW, 100 - stage.x);
      stage.h = clamp(dragState.startH + deltaYPercent, minH, 100 - stage.y);
    } else {
      const newX = ((event.clientX - rect.left - dragState.offsetX) / rect.width) * 100;
      const newY = ((event.clientY - rect.top - dragState.offsetY) / rect.height) * 100;
      stage.x = clamp(newX, 0, 100 - stage.w);
      stage.y = clamp(newY, 0, 100 - stage.h);
    }
    renderBuilder();
  });

  window.addEventListener("mouseup", () => {
    dragState = null;
  });

  createBuilderLineBtn.addEventListener("click", () => {
    if (!builderDraft) {
      alert("Builder is not ready. Please close and reopen the builder.");
      return;
    }

    try {
      const lineName = String(builderDraft.lineName || "").trim() || nextAutoLineName();
      const id = `line-${Date.now()}`;
      const builtStages = makeStagesFromBuilder(builderDraft.stages || []);
      const line = makeDefaultLine(id, lineName);
      line.stages = builtStages.length ? builtStages : clone(STAGES);
      line.crewsByShift = makeCrewByShiftFromStages(line.stages);
      line.stageSettings = makeSettingsFromStages(line.stages);
      line.selectedStageId = line.stages[0]?.id || "s1";
      addAudit(line, "CREATE_LINE", `Line created with ${line.stages.length} stages`);
      appState.lines[id] = line;
      appState.activeLineId = id;
      appState.activeView = "line";
      state = line;
      saveState();
      alert(`Line created.\nDelete key: ${line.secretKey}\nSave this key to delete the line later.`);
      closeBuilderModal();
      renderAll();
    } catch (error) {
      console.error(error);
      alert("Could not create line due to an unexpected error. Please try again.");
    }
  });

  cards.addEventListener("click", (event) => {
    const deleteBtn = event.target.closest("[data-delete-line]");
    if (deleteBtn) {
      const id = deleteBtn.getAttribute("data-delete-line");
      if (!id || !appState.lines[id]) return;
      const line = appState.lines[id];
      const entered = window.prompt(`Enter delete key for "${line.name}" (or admin password):`) || "";
      if (entered !== "admin" && entered !== line.secretKey) {
        alert("Invalid key/password. Line was not deleted.");
        return;
      }
      addAudit(line, "DELETE_LINE", "Line deleted");
      delete appState.lines[id];
      appState.supervisors = (appState.supervisors || []).map((sup) => ({
        ...sup,
        assignedLineIds: (sup.assignedLineIds || []).filter((lineId) => lineId !== id)
      }));
      if (appState.supervisorSession) {
        appState.supervisorSession.assignedLineIds = (appState.supervisorSession.assignedLineIds || []).filter((lineId) => lineId !== id);
      }
      if (!Object.keys(appState.lines).length) {
        const fallback = makeDefaultLine("line-1", "Production Line");
        appState.lines[fallback.id] = fallback;
      }
      appState.activeLineId = Object.keys(appState.lines)[0];
      state = appState.lines[appState.activeLineId];
      appState.activeView = "home";
      saveState();
      renderAll();
      return;
    }

    const btn = event.target.closest("[data-open-line]");
    if (!btn) return;
    const id = btn.getAttribute("data-open-line");
    if (!id || !appState.lines[id]) return;
    appState.activeLineId = id;
    appState.activeView = "line";
    state = appState.lines[id];
    saveState();
    renderAll();
  });
}

function bindDataSubtabs() {
  document.querySelectorAll(".data-subtab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const target = btn.dataset.dataTab;
      if (!target) return;
      state.activeDataTab = target;
      saveState();
      setActiveDataSubtab();
    });
  });
}

function bindVisualiserControls() {
  const dateInput = document.getElementById("selectedDate");
  const shiftButtons = Array.from(document.querySelectorAll(".shift-option[data-shift]"));
  const map = document.getElementById("lineMap");
  const editBtn = document.getElementById("toggleLayoutEdit");
  const addLineBtn = document.getElementById("addFlowLine");
  const addArrowBtn = document.getElementById("addFlowArrow");
  const uploadShapeBtn = document.getElementById("uploadFlowShapeBtn");
  const uploadShapeInput = document.getElementById("uploadFlowShapeInput");
  let layoutDragState = null;

  const clamp = (value, min, max) => Math.max(min, Math.min(max, value));
  const setLayoutEditButtonUI = () => {
    const active = Boolean(state.visualEditMode);
    editBtn.classList.toggle("active", active);
    editBtn.textContent = active ? "Done Editing" : "Edit Layout";
    addLineBtn.disabled = !active;
    addArrowBtn.disabled = !active;
    uploadShapeBtn.disabled = !active;
    addLineBtn.classList.toggle("hidden", !active);
    addArrowBtn.classList.toggle("hidden", !active);
    uploadShapeBtn.classList.toggle("hidden", !active);
  };

  const createGuide = (type) => ({
    id: `fg-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
    type: type === "arrow" ? "arrow" : "line",
    x: 8,
    y: 42,
    w: 14,
    h: 2,
    angle: 0
  });

  dateInput.value = state.selectedDate;
  setShiftToggleUI();
  setLayoutEditButtonUI();

  dateInput.addEventListener("change", () => {
    state.selectedDate = dateInput.value || todayISO();
    saveState();
    renderAll();
  });

  shiftButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const nextShift = btn.dataset.shift;
      if (!nextShift || nextShift === state.selectedShift) return;
      state.selectedShift = nextShift;
      setShiftToggleUI();
      saveState();
      renderAll();
    });
  });

  document.getElementById("prevShift").addEventListener("click", () => moveShift(-1));
  document.getElementById("nextShift").addEventListener("click", () => moveShift(1));
  editBtn.addEventListener("click", () => {
    state.visualEditMode = !state.visualEditMode;
    setLayoutEditButtonUI();
    saveState();
    renderVisualiser();
  });

  addLineBtn.addEventListener("click", () => {
    if (!state.visualEditMode) return;
    state.flowGuides = Array.isArray(state.flowGuides) ? state.flowGuides : [];
    state.flowGuides.push(createGuide("line"));
    addAudit(state, "LAYOUT_ADD_LINE", "Flow line guide added");
    saveState();
    renderVisualiser();
  });

  addArrowBtn.addEventListener("click", () => {
    if (!state.visualEditMode) return;
    state.flowGuides = Array.isArray(state.flowGuides) ? state.flowGuides : [];
    state.flowGuides.push(createGuide("arrow"));
    addAudit(state, "LAYOUT_ADD_ARROW", "Flow arrow guide added");
    saveState();
    renderVisualiser();
  });

  uploadShapeBtn.addEventListener("click", () => {
    if (!state.visualEditMode) return;
    uploadShapeInput.click();
  });

  uploadShapeInput.addEventListener("change", async (event) => {
    if (!state.visualEditMode) return;
    const file = event.target.files?.[0];
    if (!file) return;
    if (!file.type.startsWith("image/")) {
      alert("Please upload an image file.");
      event.target.value = "";
      return;
    }

    try {
      const dataUrl = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result || ""));
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });
      state.flowGuides = Array.isArray(state.flowGuides) ? state.flowGuides : [];
      state.flowGuides.push({
        id: `fg-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
        type: "shape",
        x: 10,
        y: 24,
        w: 18,
        h: 18,
        angle: 0,
        src: dataUrl
      });
      addAudit(state, "LAYOUT_ADD_SHAPE", "Custom background shape uploaded");
      saveState();
      renderVisualiser();
    } catch {
      alert("Could not load shape.");
    }
    event.target.value = "";
  });

  map.addEventListener("mousedown", (event) => {
    if (!state.visualEditMode) return;
    visualiserDragMoved = false;
    const deleteGuide = event.target.closest("[data-guide-delete]");
    if (deleteGuide) {
      const guideId = deleteGuide.getAttribute("data-guide-delete");
      state.flowGuides = (state.flowGuides || []).filter((guide) => guide.id !== guideId);
      addAudit(state, "LAYOUT_DELETE_GUIDE", "Flow guide deleted");
      saveState();
      renderVisualiser();
      event.preventDefault();
      return;
    }

    const guideNode = event.target.closest("[data-guide-id]");
    const guideResizeNode = event.target.closest("[data-guide-resize]");
    const guideRotateNode = event.target.closest("[data-guide-rotate]");
    if (guideNode) {
      const guideId = guideNode.getAttribute("data-guide-id");
      const guide = (state.flowGuides || []).find((item) => item.id === guideId);
      if (!guide) return;
      const rect = map.getBoundingClientRect();
      if (guideRotateNode) {
        layoutDragState = {
          target: "guide",
          mode: "rotate",
          guideId
        };
      } else if (guideResizeNode) {
        layoutDragState = {
          target: "guide",
          mode: "resize",
          guideId,
          startX: event.clientX,
          startY: event.clientY,
          startW: guide.w,
          startH: guide.h,
          canvasW: rect.width,
          canvasH: rect.height
        };
      } else {
        layoutDragState = {
          target: "guide",
          mode: "move",
          guideId,
          offsetX: event.clientX - rect.left - (guide.x / 100) * rect.width,
          offsetY: event.clientY - rect.top - (guide.y / 100) * rect.height
        };
      }
      event.preventDefault();
      return;
    }

    const resizeNode = event.target.closest("[data-stage-resize]");
    const cardNode = event.target.closest("[data-stage-id]");
    if (!cardNode) return;
    const stage = getStages().find((item) => item.id === cardNode.getAttribute("data-stage-id"));
    if (!stage) return;
    const rect = map.getBoundingClientRect();

    if (resizeNode) {
      layoutDragState = {
        target: "stage",
        mode: "resize",
        stageId: stage.id,
        startX: event.clientX,
        startY: event.clientY,
        startW: stage.w,
        startH: stage.h,
        canvasW: rect.width,
        canvasH: rect.height
      };
      event.preventDefault();
      return;
    }

    layoutDragState = {
      target: "stage",
      mode: "move",
      stageId: stage.id,
      offsetX: event.clientX - rect.left - (stage.x / 100) * rect.width,
      offsetY: event.clientY - rect.top - (stage.y / 100) * rect.height
    };
    event.preventDefault();
  });

  window.addEventListener("mousemove", (event) => {
    if (!state.visualEditMode || !layoutDragState) return;
    visualiserDragMoved = true;
    const rect = map.getBoundingClientRect();

    if (layoutDragState.target === "guide") {
      const guide = (state.flowGuides || []).find((item) => item.id === layoutDragState.guideId);
      if (!guide) return;
      if (layoutDragState.mode === "rotate") {
        const cx = rect.left + ((guide.x + guide.w / 2) / 100) * rect.width;
        const cy = rect.top + ((guide.y + guide.h / 2) / 100) * rect.height;
        guide.angle = (Math.atan2(event.clientY - cy, event.clientX - cx) * 180) / Math.PI;
      } else if (layoutDragState.mode === "resize") {
        const deltaXPercent = ((event.clientX - layoutDragState.startX) / layoutDragState.canvasW) * 100;
        const deltaYPercent = ((event.clientY - layoutDragState.startY) / layoutDragState.canvasH) * 100;
        guide.w = clamp(layoutDragState.startW + deltaXPercent, 2, 100 - guide.x);
        guide.h = clamp(layoutDragState.startH + deltaYPercent, 1, 100 - guide.y);
      } else {
        const newX = ((event.clientX - rect.left - layoutDragState.offsetX) / rect.width) * 100;
        const newY = ((event.clientY - rect.top - layoutDragState.offsetY) / rect.height) * 100;
        guide.x = clamp(newX, 0, 100 - guide.w);
        guide.y = clamp(newY, 0, 100 - guide.h);
      }
    } else {
      const stage = getStages().find((item) => item.id === layoutDragState.stageId);
      if (!stage) return;
      if (layoutDragState.mode === "resize") {
        const deltaXPercent = ((event.clientX - layoutDragState.startX) / layoutDragState.canvasW) * 100;
        const deltaYPercent = ((event.clientY - layoutDragState.startY) / layoutDragState.canvasH) * 100;
        const minW = stage.kind === "transfer" ? 5.5 : 6.5;
        const minH = stage.kind === "transfer" ? 7.5 : 10;
        stage.w = clamp(layoutDragState.startW + deltaXPercent, minW, 100 - stage.x);
        stage.h = clamp(layoutDragState.startH + deltaYPercent, minH, 100 - stage.y);
      } else {
        const newX = ((event.clientX - rect.left - layoutDragState.offsetX) / rect.width) * 100;
        const newY = ((event.clientY - rect.top - layoutDragState.offsetY) / rect.height) * 100;
        stage.x = clamp(newX, 0, 100 - stage.w);
        stage.y = clamp(newY, 0, 100 - stage.h);
      }
    }
    renderVisualiser();
  });

  window.addEventListener("mouseup", () => {
    if (layoutDragState) {
      addAudit(state, "LAYOUT_EDIT", `Layout ${layoutDragState.target} ${layoutDragState.mode}`);
      layoutDragState = null;
      saveState();
    }
  });
}

function stageBaseName(name) {
  return String(name || "").replace(/^\s*\d+\.\s*/, "").trim();
}

function stageDefaultSize(stage) {
  if (stage.group === "transfer" || stage.kind === "transfer") return { w: 7.5, h: 10 };
  return { w: 9.5, h: 15 };
}

function openStageSettingsModal(stageId) {
  const stage = getStages().find((item) => item.id === stageId);
  if (!stage) return;
  document.getElementById("stageSettingsId").value = stage.id;
  document.getElementById("stageSettingsName").value = stageBaseName(stage.name) || "Stage";
  document.getElementById("stageSettingsType").value = stage.kind === "transfer" ? "transfer" : stage.group || "main";
  document.getElementById("stageSettingsCrewDay").value = num(state.crewsByShift?.Day?.[stage.id]?.crew ?? stage.crew);
  document.getElementById("stageSettingsCrewNight").value = num(state.crewsByShift?.Night?.[stage.id]?.crew ?? stage.crew);
  document.getElementById("stageSettingsMaxThroughput").value = stageMaxThroughput(stage.id);
  document.getElementById("stageSettingsWidth").value = num(stage.w) || stageDefaultSize(stage).w;
  document.getElementById("stageSettingsHeight").value = num(stage.h) || stageDefaultSize(stage).h;

  const overlay = document.getElementById("stageSettingsModal");
  overlay.classList.add("open");
  overlay.setAttribute("aria-hidden", "false");
}

function closeStageSettingsModal() {
  const overlay = document.getElementById("stageSettingsModal");
  overlay.classList.remove("open");
  overlay.setAttribute("aria-hidden", "true");
}

function bindStageSettingsModal() {
  const overlay = document.getElementById("stageSettingsModal");
  const closeBtn = document.getElementById("closeStageSettingsModal");
  const form = document.getElementById("stageSettingsForm");

  closeBtn.addEventListener("click", closeStageSettingsModal);
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) closeStageSettingsModal();
  });

  form.addEventListener("submit", (event) => {
    event.preventDefault();
    const stageId = document.getElementById("stageSettingsId").value;
    const stage = getStages().find((item) => item.id === stageId);
    if (!stage) return;

    const name = stageBaseName(document.getElementById("stageSettingsName").value) || "Stage";
    const type = document.getElementById("stageSettingsType").value || "main";
    const crewDay = Math.max(0, num(document.getElementById("stageSettingsCrewDay").value));
    const crewNight = Math.max(0, num(document.getElementById("stageSettingsCrewNight").value));
    const maxThroughput = Math.max(0, num(document.getElementById("stageSettingsMaxThroughput").value));
    const width = Math.max(2, num(document.getElementById("stageSettingsWidth").value));
    const height = Math.max(1, num(document.getElementById("stageSettingsHeight").value));
    const defaults = stageDefaultSize(stage);

    stage.name = name;
    stage.group = type === "transfer" ? "prep" : type;
    stage.kind = type === "transfer" ? "transfer" : undefined;
    stage.w = width || defaults.w;
    stage.h = height || defaults.h;
    stage.match = name
      .toLowerCase()
      .split(/[^a-z0-9]+/)
      .filter(Boolean);

    if (!state.crewsByShift.Day[stage.id]) state.crewsByShift.Day[stage.id] = {};
    if (!state.crewsByShift.Night[stage.id]) state.crewsByShift.Night[stage.id] = {};
    state.crewsByShift.Day[stage.id].crew = crewDay;
    state.crewsByShift.Night[stage.id].crew = crewNight;
    stage.crew = num(state.crewsByShift[state.selectedShift]?.[stage.id]?.crew ?? crewDay);

    if (!state.stageSettings[stage.id]) state.stageSettings[stage.id] = {};
    state.stageSettings[stage.id].maxThroughput = maxThroughput;

    addAudit(state, "EDIT_STAGE_SETTINGS", `Stage updated: ${stageDisplayName(stage, getStages().findIndex((s) => s.id === stage.id))}`);
    saveState();
    closeStageSettingsModal();
    renderAll();
  });
}

function bindTrendModal() {
  const overlay = document.getElementById("trendModal");
  const closeBtn = document.getElementById("closeTrendModal");
  const dailyBtn = document.getElementById("trendDaily");
  const monthlyBtn = document.getElementById("trendMonthly");
  const prevBtn = document.getElementById("trendPrevMonth");
  const nextBtn = document.getElementById("trendNextMonth");

  closeBtn.addEventListener("click", closeTrendModal);

  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) closeTrendModal();
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeTrendModal();
  });

  dailyBtn.addEventListener("click", () => {
    state.trendGranularity = "daily";
    saveState();
    renderStageTrend();
  });

  monthlyBtn.addEventListener("click", () => {
    state.trendGranularity = "monthly";
    saveState();
    renderStageTrend();
  });

  prevBtn.addEventListener("click", () => {
    state.trendMonth = addMonths(state.trendMonth || monthKey(state.selectedDate), -1);
    saveState();
    renderStageTrend();
  });

  nextBtn.addEventListener("click", () => {
    state.trendMonth = addMonths(state.trendMonth || monthKey(state.selectedDate), 1);
    saveState();
    renderStageTrend();
  });
}

function moveShift(direction) {
  const date = parseDateLocal(state.selectedDate || todayISO());

  if (direction < 0) {
    date.setDate(date.getDate() - 1);
  } else {
    date.setDate(date.getDate() + 1);
  }
  state.selectedDate = formatDateLocal(date);

  document.getElementById("selectedDate").value = state.selectedDate;
  saveState();
  renderAll();
}

function setShiftToggleUI() {
  document.querySelectorAll(".shift-option[data-shift]").forEach((btn) => {
    const active = btn.dataset.shift === state.selectedShift;
    btn.classList.toggle("active", active);
    btn.setAttribute("aria-pressed", String(active));
  });
}

function bindForms() {
  const equipmentSelect = document.getElementById("downtimeEquipment");
  equipmentSelect.innerHTML = stageOptionListHTML();

  document.getElementById("shiftForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const data = Object.fromEntries(new FormData(event.currentTarget).entries());
    data.crewOnShift = Math.max(0, Math.floor(num(data.crewOnShift)));
    if (!rowIsValidDateShift(data.date, data.shift)) {
      alert("Date and shift are required.");
      return;
    }
    if (!strictTimeValid(data.startTime) || !strictTimeValid(data.finishTime)) {
      alert("Shift start and finish must be HH:MM (24h).");
      return;
    }
    if (!optionalStrictTimeValid(data.break1Start) || !optionalStrictTimeValid(data.break2Start) || !optionalStrictTimeValid(data.break3Start)) {
      alert("Break times must be HH:MM.");
      return;
    }
    if (data.crewOnShift < 0) {
      alert("Crew on shift cannot be negative.");
      return;
    }
    state.shiftRows.push(data);
    addAudit(state, "MANAGER_SHIFT_LOG", `Manager logged ${data.shift} shift for ${data.date} (crew ${data.crewOnShift})`);
    await syncManagerShiftLog({
      lineId: state.id,
      date: data.date,
      shift: data.shift,
      crewOnShift: data.crewOnShift,
      startTime: data.startTime,
      break1Start: data.break1Start || "",
      break2Start: data.break2Start || "",
      break3Start: data.break3Start || "",
      finishTime: data.finishTime
    });
    event.currentTarget.reset();
    saveState();
    renderAll();
  });

  document.getElementById("runForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const data = Object.fromEntries(new FormData(event.currentTarget).entries());
    if (!rowIsValidDateShift(data.date, data.shift)) {
      alert("Date and shift are required.");
      return;
    }
    if (!String(data.product || "").trim()) {
      alert("Product is required.");
      return;
    }
    if (!optionalStrictTimeValid(data.setUpStartTime) || !strictTimeValid(data.productionStartTime) || !strictTimeValid(data.finishTime)) {
      alert("Production start and finish must be HH:MM (24h).");
      return;
    }
    if (num(data.unitsProduced) < 0) {
      alert("Units produced cannot be negative.");
      return;
    }
    data.unitsProduced = num(data.unitsProduced);
    state.runRows.push(data);
    addAudit(state, "MANAGER_RUN_LOG", `Manager logged run ${data.product} (${data.unitsProduced} units)`);
    await syncManagerRunLog({
      lineId: state.id,
      date: data.date,
      shift: data.shift,
      product: data.product,
      setUpStartTime: data.setUpStartTime || "",
      productionStartTime: data.productionStartTime,
      finishTime: data.finishTime,
      unitsProduced: data.unitsProduced
    });
    event.currentTarget.reset();
    saveState();
    renderAll();
  });

  document.getElementById("downtimeForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const data = Object.fromEntries(new FormData(event.currentTarget).entries());
    if (!rowIsValidDateShift(data.date, data.shift)) {
      alert("Date and shift are required.");
      return;
    }
    if (!strictTimeValid(data.downtimeStart) || !strictTimeValid(data.downtimeFinish)) {
      alert("Downtime start and finish must be HH:MM (24h).");
      return;
    }
    if (!String(data.equipment || "")) {
      alert("Equipment stage is required.");
      return;
    }
    state.downtimeRows.push(data);
    addAudit(state, "MANAGER_DOWNTIME_LOG", `Manager logged downtime on ${stageNameById(data.equipment)}`);
    await syncManagerDowntimeLog({
      lineId: state.id,
      date: data.date,
      shift: data.shift,
      downtimeStart: data.downtimeStart,
      downtimeFinish: data.downtimeFinish,
      equipment: data.equipment,
      reason: data.reason || ""
    });
    event.currentTarget.reset();
    equipmentSelect.innerHTML = stageOptionListHTML();
    saveState();
    renderAll();
  });
}

function bindDataControls() {
  document.getElementById("loadSampleData").addEventListener("click", () => {
    const sample = sampleDataSet();
    state.selectedDate = sample.selectedDate;
    state.selectedShift = sample.selectedShift;
    state.trendGranularity = sample.trendGranularity;
    state.trendMonth = sample.trendMonth;
    state.shiftRows = sample.shiftRows;
    state.runRows = sample.runRows;
    state.downtimeRows = sample.downtimeRows;
    addAudit(state, "LOAD_SAMPLE_DATA", "Sample shift/run/downtime rows loaded");
    saveState();
    renderAll();
  });

  document.getElementById("exportData").addEventListener("click", () => {
    downloadTextFile(`kebab-line-data-${Date.now()}.json`, JSON.stringify(state, null, 2), "application/json");
    addAudit(state, "EXPORT_JSON", "Line JSON exported");
    saveState();
  });

  document.getElementById("exportLineCsv").addEventListener("click", () => {
    const data = derivedData();
    const rows = [
      ...data.shiftRows.map((row) => ({
        RecordType: "Shift",
        Date: row.date,
        Shift: row.shift,
        Product: "",
        Equipment: "",
        Units: "",
        DowntimeMins: "",
        Start: row.startTime || "",
        Finish: row.finishTime || "",
        Details: `Breaks: ${row.break1Start || "-"}, ${row.break2Start || "-"}, ${row.break3Start || "-"}`
      })),
      ...data.runRows.map((row) => ({
        RecordType: "Run",
        Date: row.date,
        Shift: row.shift,
        Product: row.product || "",
        Equipment: "",
        Units: Number(row.unitsProduced || 0).toFixed(2),
        DowntimeMins: Number(row.associatedDownTime || 0).toFixed(2),
        Start: row.productionStartTime || "",
        Finish: row.finishTime || "",
        Details: `Net rate ${Number(row.netRunRate || 0).toFixed(2)} u/min`
      })),
      ...data.downtimeRows.map((row) => ({
        RecordType: "Downtime",
        Date: row.date,
        Shift: row.shift,
        Product: "",
        Equipment: stageNameById(row.equipment),
        Units: "",
        DowntimeMins: Number(row.downtimeMins || 0).toFixed(2),
        Start: row.downtimeStart || "",
        Finish: row.downtimeFinish || "",
        Details: row.reason || ""
      }))
    ];
    const columns = ["RecordType", "Date", "Shift", "Product", "Equipment", "Units", "DowntimeMins", "Start", "Finish", "Details"];
    downloadTextFile(`line-report-${state.name.replace(/\s+/g, "-").toLowerCase()}-${Date.now()}.csv`, toCsv(rows, columns), "text/csv;charset=utf-8");
    addAudit(state, "EXPORT_CSV", "Line CSV report exported");
    saveState();
  });

  document.getElementById("importData").addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      state.selectedDate = parsed.selectedDate || state.selectedDate;
      state.selectedShift = parsed.selectedShift || state.selectedShift;
      state.selectedStageId = parsed.selectedStageId || state.selectedStageId;
      state.activeDataTab = parsed.activeDataTab || state.activeDataTab;
      state.trendGranularity = parsed.trendGranularity || state.trendGranularity;
      state.trendMonth = parsed.trendMonth || state.trendMonth;
      state.crewsByShift = normalizeCrewByShift(parsed);
      state.stageSettings = normalizeStageSettings(parsed);
      state.shiftRows = parsed.shiftRows || [];
      state.runRows = parsed.runRows || [];
      state.downtimeRows = parsed.downtimeRows || [];
      addAudit(state, "IMPORT_JSON", "Line JSON imported");
      saveState();
      renderAll();
    } catch {
      alert("Invalid JSON file.");
    }

    event.target.value = "";
  });

  document.getElementById("clearData").addEventListener("click", () => {
    const entered = window.prompt(`Enter secret key to clear all data for "${state.name}" (or admin):`) || "";
    if (entered !== "admin" && entered !== state.secretKey) {
      alert("Invalid key/password. Data was not cleared.");
      return;
    }
    if (!window.confirm(`Clear all tracking data for "${state.name}" only?`)) return;
    clearLineTrackingData(state);
    addAudit(state, "CLEAR_DATA", "All shift/run/downtime rows cleared for this line");
    saveState();
    renderAll();
  });
}

function renderCrewInputs() {
  const form = document.getElementById("crewForm");
  const shiftNote = document.getElementById("crewShiftNote");
  if (!form || !shiftNote) return;
  form.innerHTML = "";
  shiftNote.textContent = `Crew values for ${state.selectedShift} shift.`;
  const activeCrew = state.crewsByShift[state.selectedShift] || defaultStageCrew();

  getStages().forEach((stage, index) => {
    const row = document.createElement("label");
    row.className = "crew-item";

    const name = document.createElement("span");
    name.textContent = stage.name;

    const input = document.createElement("input");
    input.type = "number";
    input.min = "0";
    input.step = "1";
    input.value = activeCrew?.[stage.id]?.crew ?? stage.crew;
    input.addEventListener("change", () => {
      if (!state.crewsByShift[state.selectedShift]) state.crewsByShift[state.selectedShift] = defaultStageCrew();
      if (!state.crewsByShift[state.selectedShift][stage.id]) state.crewsByShift[state.selectedShift][stage.id] = {};
      state.crewsByShift[state.selectedShift][stage.id].crew = num(input.value);
      saveState();
      renderVisualiser();
      renderStageTrend();
    });

    row.append(name, input);
    form.append(row);
  });
}

function renderThroughputInputs() {
  const form = document.getElementById("throughputForm");
  if (!form) return;
  form.innerHTML = "";

  getStages().forEach((stage, index) => {
    const row = document.createElement("label");
    row.className = "crew-item";

    const name = document.createElement("span");
    name.textContent = stage.name;

    const input = document.createElement("input");
    input.type = "number";
    input.min = "0";
    input.step = "1";
    input.value = stageMaxThroughput(stage.id);
    input.addEventListener("change", () => {
      if (!state.stageSettings[stage.id]) state.stageSettings[stage.id] = {};
      state.stageSettings[stage.id].maxThroughput = num(input.value);
      saveState();
      renderVisualiser();
      renderStageTrend();
    });

    row.append(name, input);
    form.append(row);
  });
}

function renderTable(tableId, columns, rows, fieldMap) {
  const formatTableCellValue = (value) => {
    if (value === null || value === undefined || value === "") return "";
    if (typeof value === "number" && Number.isFinite(value)) {
      return value.toLocaleString(undefined, { maximumFractionDigits: 2 });
    }
    if (typeof value === "string") {
      const raw = value.trim();
      if (/^-?\d+(?:\.\d+)?$/.test(raw)) {
        const n = Number(raw);
        if (Number.isFinite(n)) return n.toLocaleString(undefined, { maximumFractionDigits: 2 });
      }
    }
    return value;
  };

  const table = document.getElementById(tableId);
  const header = `<thead><tr>${columns.map((col) => `<th>${col}</th>`).join("")}</tr></thead>`;
  const body = rows
    .map((row) => `<tr>${columns.map((col) => `<td>${formatTableCellValue(row[fieldMap[col]])}</td>`).join("")}</tr>`)
    .join("");

  table.innerHTML = `${header}<tbody>${body}</tbody>`;
}

function selectedRows(rows) {
  return rows.filter((row) => row.date === state.selectedDate && row.shift === state.selectedShift);
}

function renderVisualiser() {
  const data = derivedData();
  const selectedRunRows = selectedRows(data.runRows);
  const selectedDowntimeRows = selectedRows(data.downtimeRows);
  const selectedShiftRows = selectedRows(data.shiftRows);

  const shiftMins = num(selectedShiftRows[0]?.totalShiftTime);
  const units = selectedRunRows.reduce((sum, row) => sum + num(row.unitsProduced), 0);
  const totalDowntime = selectedDowntimeRows.reduce((sum, row) => sum + num(row.downtimeMins), 0);
  const totalNetTime = selectedRunRows.reduce((sum, row) => sum + num(row.netProductionTime), 0);
  const netRunRate = totalNetTime > 0 ? units / totalNetTime : 0;
  let utilAccumulator = 0;
  let utilCount = 0;

  document.getElementById("kpiUnits").textContent = formatNum(units, 0);
  document.getElementById("kpiDowntime").textContent = `${formatNum(totalDowntime, 1)} min`;
  document.getElementById("kpiRunRate").textContent = `${formatNum(netRunRate, 2)} u/min`;

  const map = document.getElementById("lineMap");
  map.innerHTML = "";
  map.classList.toggle("layout-editing", Boolean(state.visualEditMode));
  const activeCrew = state.crewsByShift[state.selectedShift] || defaultStageCrew();
  let bottleneckCard = null;
  let bottleneckUtilisation = -1;
  const guides = normalizeFlowGuides(state.flowGuides);
  state.flowGuides = guides;

  guides.forEach((guide) => {
    const node = document.createElement("div");
    node.className = `flow-guide flow-${guide.type}`;
    node.setAttribute("data-guide-id", guide.id);
    node.style.left = `${guide.x}%`;
    node.style.top = `${guide.y}%`;
    node.style.width = `${guide.w}%`;
    node.style.height = `${guide.h}%`;
    node.style.transform = `rotate(${guide.angle || 0}deg)`;
    if (state.visualEditMode) {
      node.innerHTML = `
        <span class="flow-guide-delete" data-guide-delete="${guide.id}">x</span>
        <span class="flow-guide-resize" data-guide-resize="${guide.id}"></span>
        <span class="flow-guide-rotate" data-guide-rotate="${guide.id}"></span>
      `;
    }
    if (guide.type === "shape" && guide.src) {
      const img = document.createElement("img");
      img.className = "flow-shape-image";
      img.src = guide.src;
      img.alt = "";
      img.draggable = false;
      node.append(img);
    }
    map.append(node);
  });

  getStages().forEach((stage, index) => {
    const stageDowntime = selectedDowntimeRows
      .filter((row) => matchesStage(stage, row.equipment))
      .reduce((sum, row) => sum + num(row.downtimeMins), 0);

    const uptimeRatio = shiftMins > 0 ? Math.max(0, (shiftMins - stageDowntime) / shiftMins) : 0;
    const stageRate = netRunRate * uptimeRatio;
    const perCrewMax = stageMaxThroughput(stage.id);
    const totalMaxThroughput = stageTotalMaxThroughput(stage.id, state.selectedShift);
    const utilisation = totalMaxThroughput > 0 ? (stageRate / totalMaxThroughput) * 100 : 0;
    utilAccumulator += Math.max(0, utilisation);
    utilCount += 1;
    const stageCrew = num(activeCrew?.[stage.id]?.crew ?? stage.crew);
    const compact = stage.w * stage.h < 140;
    const status = statusClass(utilisation);

    const card = document.createElement("article");
    card.className = `stage-card group-${stage.group}${stage.kind ? ` kind-${stage.kind}` : ""} status-${status}`;
    card.setAttribute("data-stage-id", stage.id);
    card.style.left = `${stage.x}%`;
    card.style.top = `${stage.y}%`;
    card.style.width = `${stage.w}%`;
    card.style.height = `${stage.h}%`;
    card.classList.toggle("compact", compact);
    card.classList.toggle("selected", stage.id === state.selectedStageId);
    if (utilisation > bottleneckUtilisation) {
      bottleneckUtilisation = utilisation;
      bottleneckCard = card;
    }
    if (!state.visualEditMode) {
      card.addEventListener("click", () => {
        state.selectedStageId = stage.id;
        saveState();
        renderVisualiser();
        renderStageTrend();
        openTrendModal();
      });
    } else {
      card.addEventListener("click", () => {
        if (visualiserDragMoved) {
          visualiserDragMoved = false;
          return;
        }
        openStageSettingsModal(stage.id);
      });
    }

    if (stage.kind === "transfer") {
      card.innerHTML = `
        <h3 class="stage-title">${stageDisplayName(stage, index)}</h3>
        <div class="stage-meta compact">
          <div>Util: ${formatNum(Math.max(utilisation, 0), 1)}%</div>
          <div>Down: ${formatNum(stageDowntime, 1)} min</div>
        </div>
        <div class="stage-tag-row">
          <span class="status-tag ${status}">${status.toUpperCase()}</span>
        </div>
        ${state.visualEditMode ? `<span class="stage-resize-handle" data-stage-resize="${stage.id}"></span>` : ""}
      `;
    } else {
      card.innerHTML = `
        ${stageCrew > 0 ? `<span class="crew-badge">${stageCrew}</span>` : ""}
        <h3 class="stage-title">${stageDisplayName(stage, index)}</h3>
        <div class="stage-meta compact">
          <div>Util: ${formatNum(Math.max(utilisation, 0), 1)}%</div>
          ${compact ? "" : `<div>Down: ${formatNum(stageDowntime, 1)} min</div>`}
          ${compact ? "" : `<div>ETC: ${formatNum(stageRate, 2)} u/min</div>`}
          ${compact ? "" : `<div>Max: ${formatNum(totalMaxThroughput, 0)} u/min</div>`}
          ${compact ? "" : `<div>Per Crew: ${formatNum(perCrewMax, 0)} u/min</div>`}
        </div>
        <div class="stage-tag-row">
          <span class="status-tag ${status}">${status.toUpperCase()}</span>
        </div>
        ${state.visualEditMode ? `<span class="stage-resize-handle" data-stage-resize="${stage.id}"></span>` : ""}
      `;
    }

    map.append(card);
  });

  if (bottleneckCard) {
    bottleneckCard.classList.add("bottleneck");
    const tagRow = bottleneckCard.querySelector(".stage-tag-row");
    if (tagRow) {
      const badge = document.createElement("span");
      badge.className = "bottleneck-tag";
      badge.textContent = "BOTTLENECK";
      tagRow.append(badge);
    }
  }

  const lineUtil = utilCount > 0 ? utilAccumulator / utilCount : 0;
  document.getElementById("kpiUtilisation").textContent = `${formatNum(Math.max(lineUtil, 0), 1)}%`;
}

function stageDailyMetrics(stage, date, shift, data) {
  const shiftRow = data.shiftRows.find((row) => row.date === date && row.shift === shift);
  const shiftMins = num(shiftRow?.totalShiftTime);
  const stageDowntime = data.downtimeRows
    .filter((row) => row.date === date && row.shift === shift)
    .filter((row) => matchesStage(stage, row.equipment))
    .reduce((sum, row) => sum + num(row.downtimeMins), 0);

  const runRows = data.runRows.filter((row) => row.date === date && row.shift === shift);
  const units = runRows.reduce((sum, row) => sum + num(row.unitsProduced), 0);
  const totalNetTime = runRows.reduce((sum, row) => sum + num(row.netProductionTime), 0);
  const netRunRate = totalNetTime > 0 ? units / totalNetTime : 0;
  const uptimeRatio = shiftMins > 0 ? Math.max(0, (shiftMins - stageDowntime) / shiftMins) : 0;
  const stageEtc = netRunRate * uptimeRatio;
  const totalMaxThroughput = stageTotalMaxThroughput(stage.id, shift);
  const utilisation = totalMaxThroughput > 0 ? (stageEtc / totalMaxThroughput) * 100 : 0;

  return { date, shiftMins, stageDowntime, utilisation, stageEtc };
}

function trendDatesForShift(shift, data) {
  const dates = new Set();
  data.shiftRows.forEach((row) => row.shift === shift && row.date && dates.add(row.date));
  data.runRows.forEach((row) => row.shift === shift && row.date && dates.add(row.date));
  data.downtimeRows.forEach((row) => row.shift === shift && row.date && dates.add(row.date));
  return Array.from(dates).sort();
}

function renderStageTrend() {
  const container = document.getElementById("stageTrendChart");
  const title = document.getElementById("trendTitle");
  const meta = document.getElementById("trendMeta");
  const data = derivedData();
  const stages = getStages();
  const stageIdx = stages.findIndex((item) => item.id === state.selectedStageId);
  const stage = stageIdx >= 0 ? stages[stageIdx] : stages[0];
  const allDates = trendDatesForShift(state.selectedShift, data);
  const allMonths = Array.from(new Set(allDates.map(monthKey))).sort();
  const selectedMonthFromDate = monthKey(state.selectedDate);
  if (!state.trendMonth) state.trendMonth = selectedMonthFromDate;
  if (allMonths.length && !allMonths.includes(state.trendMonth)) {
    state.trendMonth = allMonths.includes(selectedMonthFromDate) ? selectedMonthFromDate : allMonths[allMonths.length - 1];
  }

  let points = [];
  if (state.trendGranularity === "monthly") {
    points = allMonths.map((month) => {
      const monthDates = allDates.filter((d) => monthKey(d) === month);
      const dayPoints = monthDates.map((date) => stageDailyMetrics(stage, date, state.selectedShift, data));
      const utilisation = dayPoints.length ? dayPoints.reduce((s, p) => s + p.utilisation, 0) / dayPoints.length : 0;
      const stageDowntime = dayPoints.reduce((s, p) => s + p.stageDowntime, 0);
      const stageEtc = dayPoints.length ? dayPoints.reduce((s, p) => s + p.stageEtc, 0) / dayPoints.length : 0;
      return { date: month, utilisation, stageDowntime, stageEtc, label: month.slice(2) };
    });
  } else {
    const monthDates = allDates.filter((d) => monthKey(d) === state.trendMonth);
    points = monthDates.map((date) => {
      const p = stageDailyMetrics(stage, date, state.selectedShift, data);
      return { ...p, label: date.slice(5) };
    });
  }

  title.textContent = `${stageDisplayName(stage, stageIdx >= 0 ? stageIdx : 0)} Trend`;
  meta.textContent = `Shift: ${state.selectedShift} | ${state.trendGranularity === "monthly" ? "Monthly aggregated" : "Daily within selected month"} | Utilisation is based on ETC vs max throughput.`;
  setTrendControlsUI(allMonths);

  if (points.length < 2) {
    container.innerHTML = `<div class="empty-chart">Need at least 2 ${state.trendGranularity === "monthly" ? "months" : "dates in the selected month"} to draw a trend.</div>`;
    return;
  }

  const width = 1180;
  const height = 340;
  const pad = { top: 22, right: 88, bottom: 58, left: 62 };
  const chartW = width - pad.left - pad.right;
  const chartH = height - pad.top - pad.bottom;
  const maxUtil = Math.max(100, ...points.map((p) => p.utilisation));
  const maxDown = Math.max(1, ...points.map((p) => p.stageDowntime));
  const maxEtc = Math.max(1, ...points.map((p) => p.stageEtc));
  const stepX = points.length > 1 ? chartW / (points.length - 1) : chartW;

  const yUtil = (v) => pad.top + chartH - (Math.max(0, Math.min(maxUtil, v)) / maxUtil) * chartH;
  const yDown = (v) => pad.top + chartH - (Math.max(0, v) / maxDown) * chartH;
  const yEtc = (v) => pad.top + chartH - (Math.max(0, v) / maxEtc) * chartH;
  const x = (i) => pad.left + stepX * i;

  const utilPath = points.map((p, i) => `${i === 0 ? "M" : "L"}${x(i)},${yUtil(p.utilisation)}`).join(" ");
  const etcPath = points.map((p, i) => `${i === 0 ? "M" : "L"}${x(i)},${yEtc(p.stageEtc)}`).join(" ");

  const bars = points
    .map((p, i) => {
      const bw = Math.max(8, Math.min(24, stepX * 0.42));
      const bx = x(i) - bw / 2;
      const by = yDown(p.stageDowntime);
      const bh = pad.top + chartH - by;
      return `<rect x="${bx}" y="${by}" width="${bw}" height="${bh}" rx="3" class="bar-down" />`;
    })
    .join("");

  const utilTicks = [0, 0.25, 0.5, 0.75, 1].map((p) => Math.round(maxUtil * p));
  const ticks = utilTicks
    .map((t) => {
      const y = yUtil(t);
      return `<g><line x1="${pad.left}" y1="${y}" x2="${pad.left + chartW}" y2="${y}" class="grid"/><text x="${pad.left - 10}" y="${y + 4}" text-anchor="end" class="axis">${t}%</text></g>`;
    })
    .join("");

  const labelSkip = points.length > 16 ? Math.ceil(points.length / 12) : 1;
  const labels = points
    .map((p, i) => {
      const label = i % labelSkip === 0 || i === points.length - 1 ? p.label : "";
      return `<text x="${x(i)}" y="${height - 16}" text-anchor="middle" class="axis">${label}</text>`;
    })
    .join("");

  container.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" class="trend-svg" role="img" aria-label="${stage.name} performance trend">
      ${ticks}
      <line x1="${pad.left}" y1="${pad.top + chartH}" x2="${pad.left + chartW}" y2="${pad.top + chartH}" class="axis-line" />
      <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${pad.top + chartH}" class="axis-line" />
      ${bars}
      <path d="${utilPath}" class="line-util" />
      <path d="${etcPath}" class="line-etc" />
      ${points
        .map(
          (p, i) => `
          <circle cx="${x(i)}" cy="${yUtil(p.utilisation)}" r="3.5" class="dot-util">
            <title>${p.date} ${state.selectedShift}: Util ${formatNum(p.utilisation, 1)}%, Down ${formatNum(
              p.stageDowntime,
              1
            )} min, ETC ${formatNum(p.stageEtc, 2)} u/min</title>
          </circle>`
        )
        .join("")}
      ${labels}
      <text x="${pad.left}" y="${12}" class="legend util">Utilisation %</text>
      <text x="${pad.left + 120}" y="${12}" class="legend down">Downtime min (bars)</text>
      <text x="${pad.left + 300}" y="${12}" class="legend etc">ETC u/min</text>
      <text x="${width - 6}" y="${pad.top + 10}" text-anchor="end" class="axis">Downtime scale max ${formatNum(maxDown, 1)} min</text>
    </svg>
  `;
}

function setTrendControlsUI(allMonths = []) {
  const dailyBtn = document.getElementById("trendDaily");
  const monthlyBtn = document.getElementById("trendMonthly");
  const label = document.getElementById("trendMonthLabel");
  const prevBtn = document.getElementById("trendPrevMonth");
  const nextBtn = document.getElementById("trendNextMonth");
  const activeMonth = state.trendMonth || monthKey(state.selectedDate);

  dailyBtn.classList.toggle("active", state.trendGranularity === "daily");
  monthlyBtn.classList.toggle("active", state.trendGranularity === "monthly");
  label.textContent = formatMonthLabel(activeMonth);

  if (!allMonths.length) {
    prevBtn.disabled = true;
    nextBtn.disabled = true;
    return;
  }

  const idx = allMonths.indexOf(activeMonth);
  const safeIdx = idx >= 0 ? idx : allMonths.length - 1;
  prevBtn.disabled = safeIdx <= 0;
  nextBtn.disabled = safeIdx >= allMonths.length - 1;
}

function openTrendModal() {
  const overlay = document.getElementById("trendModal");
  overlay.classList.add("open");
  overlay.setAttribute("aria-hidden", "false");
}

function closeTrendModal() {
  const overlay = document.getElementById("trendModal");
  overlay.classList.remove("open");
  overlay.setAttribute("aria-hidden", "true");
}

function renderTrackingTables() {
  const data = derivedData();

  renderTable("shiftTable", SHIFT_COLUMNS, data.shiftRows, {
    Date: "date",
    Shift: "shift",
    "Crew On Shift": "crewOnShift",
    "Start Time": "startTime",
    "Break 1 Start": "break1Start",
    "Break 2 Start": "break2Start",
    "Break 3 Start": "break3Start",
    "Finish Time": "finishTime",
    "Total Shift Time": "totalShiftTime"
  });

  renderTable("runTable", RUN_COLUMNS, data.runRows, {
    Date: "date",
    Shift: "shift",
    Product: "product",
    "Set Up Start Time": "setUpStartTime",
    "Production Start Time": "productionStartTime",
    "Finish Time": "finishTime",
    "Units Produced": "unitsProduced",
    "Gross Production Time": "grossProductionTime",
    "Associated Down Time": "associatedDownTime",
    "Net Production Time": "netProductionTime",
    "Gross Run Rate": "grossRunRate",
    "Net Run Rate": "netRunRate"
  });

  const displayDowntimeRows = data.downtimeRows.map((row) => ({ ...row, equipment: stageNameById(row.equipment) }));
  renderTable("downtimeTable", DOWN_COLUMNS, displayDowntimeRows, {
    Date: "date",
    Shift: "shift",
    "Downtime Start": "downtimeStart",
    "Downtime Finish": "downtimeFinish",
    "Downtime (mins)": "downtimeMins",
    Equipment: "equipment",
    Reason: "reason"
  });
}

function renderAuditTrail() {
  const container = document.getElementById("auditTrail");
  if (!container) return;
  const rows = (ensureAuditRows(state) || []).slice(0, 40);
  if (!rows.length) {
    container.innerHTML = `<p class="muted">No audit events yet.</p>`;
    return;
  }
  const mapped = rows.map((row) => ({
    when: row.at ? new Date(row.at).toLocaleString() : "-",
    actor: row.actor || "-",
    action: row.action || "-",
    details: row.details || ""
  }));
  const table = document.createElement("table");
  table.id = "auditTable";
  container.innerHTML = "";
  container.append(table);
  renderTable("auditTable", AUDIT_COLUMNS, mapped, {
    When: "when",
    Actor: "actor",
    Action: "action",
    Details: "details"
  });
}

function renderSupervisorVisualiser(line, selectedDate, selectedShift) {
  const map = document.getElementById("supervisorLineMap");
  if (!map) return;
  const stages = line?.stages?.length ? line.stages : STAGES;
  const data = derivedDataForLine(line || {});
  const selectedRunRows = data.runRows.filter((row) => row.date === selectedDate && row.shift === selectedShift);
  const selectedDownRows = data.downtimeRows.filter((row) => row.date === selectedDate && row.shift === selectedShift);
  const selectedShiftRows = data.shiftRows.filter((row) => row.date === selectedDate && row.shift === selectedShift);
  const shiftMins = num(selectedShiftRows[0]?.totalShiftTime);
  const units = selectedRunRows.reduce((sum, row) => sum + num(row.unitsProduced), 0);
  const totalDowntime = selectedDownRows.reduce((sum, row) => sum + num(row.downtimeMins), 0);
  const totalNetTime = selectedRunRows.reduce((sum, row) => sum + num(row.netProductionTime), 0);
  const netRunRate = totalNetTime > 0 ? units / totalNetTime : 0;
  let utilAccumulator = 0;
  let utilCount = 0;
  let bottleneckCard = null;
  let bottleneckUtil = -1;
  const activeCrew = line?.crewsByShift?.[selectedShift] || defaultStageCrew(stages);

  document.getElementById("svKpiUnits").textContent = formatNum(units, 0);
  document.getElementById("svKpiDowntime").textContent = `${formatNum(totalDowntime, 1)} min`;
  document.getElementById("svKpiRunRate").textContent = `${formatNum(netRunRate, 2)} u/min`;
  map.innerHTML = "";
  map.classList.remove("layout-editing");

  normalizeFlowGuides(line?.flowGuides).forEach((guide) => {
    const node = document.createElement("div");
    node.className = `flow-guide flow-${guide.type}`;
    node.style.left = `${guide.x}%`;
    node.style.top = `${guide.y}%`;
    node.style.width = `${guide.w}%`;
    node.style.height = `${guide.h}%`;
    node.style.transform = `rotate(${guide.angle || 0}deg)`;
    if (guide.type === "shape" && guide.src) {
      const img = document.createElement("img");
      img.className = "flow-shape-image";
      img.src = guide.src;
      img.alt = "";
      node.append(img);
    }
    map.append(node);
  });

  stages.forEach((stage, index) => {
    const stageDowntime = selectedDownRows
      .filter((row) => matchesStage(stage, row.equipment))
      .reduce((sum, row) => sum + num(row.downtimeMins), 0);
    const uptimeRatio = shiftMins > 0 ? Math.max(0, (shiftMins - stageDowntime) / shiftMins) : 0;
    const stageRate = netRunRate * uptimeRatio;
    const perCrewMax = stageMaxThroughputForLine(line, stage.id);
    const totalMax = stageTotalMaxThroughputForLine(line, stage.id, selectedShift);
    const utilisation = totalMax > 0 ? (stageRate / totalMax) * 100 : 0;
    const stageCrew = num(activeCrew?.[stage.id]?.crew ?? stage.crew);
    const compact = stage.w * stage.h < 140;
    const status = statusClass(utilisation);
    utilAccumulator += Math.max(0, utilisation);
    utilCount += 1;

    const card = document.createElement("article");
    card.className = `stage-card group-${stage.group}${stage.kind ? ` kind-${stage.kind}` : ""} status-${status}`;
    card.style.left = `${stage.x}%`;
    card.style.top = `${stage.y}%`;
    card.style.width = `${stage.w}%`;
    card.style.height = `${stage.h}%`;
    card.classList.toggle("compact", compact);

    if (stage.kind === "transfer") {
      card.innerHTML = `
        <h3 class="stage-title">${stageDisplayName(stage, index)}</h3>
        <div class="stage-meta compact">
          <div>Util: ${formatNum(Math.max(utilisation, 0), 1)}%</div>
          <div>Down: ${formatNum(stageDowntime, 1)} min</div>
        </div>
        <div class="stage-tag-row">
          <span class="status-tag ${status}">${status.toUpperCase()}</span>
        </div>
      `;
    } else {
      card.innerHTML = `
        ${stageCrew > 0 ? `<span class="crew-badge">${stageCrew}</span>` : ""}
        <h3 class="stage-title">${stageDisplayName(stage, index)}</h3>
        <div class="stage-meta compact">
          <div>Util: ${formatNum(Math.max(utilisation, 0), 1)}%</div>
          ${compact ? "" : `<div>Down: ${formatNum(stageDowntime, 1)} min</div>`}
          ${compact ? "" : `<div>ETC: ${formatNum(stageRate, 2)} u/min</div>`}
          ${compact ? "" : `<div>Max: ${formatNum(totalMax, 0)} u/min</div>`}
          ${compact ? "" : `<div>Per Crew: ${formatNum(perCrewMax, 0)} u/min</div>`}
        </div>
        <div class="stage-tag-row">
          <span class="status-tag ${status}">${status.toUpperCase()}</span>
        </div>
      `;
    }

    if (utilisation > bottleneckUtil) {
      bottleneckUtil = utilisation;
      bottleneckCard = card;
    }
    map.append(card);
  });

  if (bottleneckCard) {
    bottleneckCard.classList.add("bottleneck");
    const row = bottleneckCard.querySelector(".stage-tag-row");
    if (row) {
      const badge = document.createElement("span");
      badge.className = "bottleneck-tag";
      badge.textContent = "BOTTLENECK";
      row.append(badge);
    }
  }
  const lineUtil = utilCount > 0 ? utilAccumulator / utilCount : 0;
  document.getElementById("svKpiUtilisation").textContent = `${formatNum(Math.max(lineUtil, 0), 1)}%`;
}

function renderHome() {
  const homeTitle = document.getElementById("homeTitle");
  const sidebarBackdrop = document.getElementById("sidebarBackdrop");
  const homeSidebar = document.getElementById("homeSidebar");
  const headerLogoutBtn = document.getElementById("supervisorLogout");
  const managerHome = document.getElementById("managerHome");
  const supervisorHome = document.getElementById("supervisorHome");
  const modeManagerBtn = document.getElementById("modeManager");
  const modeSupervisorBtn = document.getElementById("modeSupervisor");
  const dashboardDateInput = document.getElementById("dashboardDate");
  const dashboardShiftButtons = Array.from(document.querySelectorAll("[data-dash-shift]"));
  const dashboardTable = document.getElementById("dashboardTable");
  const loginSection = document.getElementById("supervisorLoginSection");
  const appSection = document.getElementById("supervisorAppSection");
  const supervisorMobileModeBtn = document.getElementById("supervisorMobileMode");
  const welcome = document.getElementById("supervisorWelcome");
  const lineSelect = document.getElementById("supervisorLineSelect");
  const svDateInput = document.getElementById("svSelectedDate");
  const svShiftButtons = Array.from(document.querySelectorAll("[data-sv-shift]"));
  const downEquipment = document.getElementById("superDownEquipment");
  const shiftDateInput = document.getElementById("superShiftDate");
  const runDateInput = document.getElementById("superRunDate");
  const downDateInput = document.getElementById("superDownDate");
  const entryList = document.getElementById("supervisorEntryList");
  const entryCards = document.getElementById("supervisorEntryCards");
  const superMainTabBtns = Array.from(document.querySelectorAll("[data-super-main-tab]"));
  const superMainPanels = Array.from(document.querySelectorAll(".supervisor-main-panel"));
  const session = normalizeSupervisorSession(appState.supervisorSession, appState.supervisors, appState.lines);
  appState.supervisorSession = session;
  const isSupervisor = appState.appMode === "supervisor";

  if (homeTitle) {
    homeTitle.textContent = isSupervisor ? "Production Line Hub Supervisor Portal" : "Production Line Dashboard";
  }
  if (headerLogoutBtn) {
    headerLogoutBtn.classList.toggle("hidden", !(isSupervisor && session));
  }

  managerHome.classList.toggle("hidden", isSupervisor);
  supervisorHome.classList.toggle("hidden", !isSupervisor);
  modeManagerBtn.classList.toggle("active", !isSupervisor);
  modeSupervisorBtn.classList.toggle("active", isSupervisor);
  supervisorMobileModeBtn.classList.toggle("hidden", true);

  const cards = document.getElementById("lineCards");
  const lineList = Object.values(appState.lines || {});
  const statTotalLines = document.getElementById("statTotalLines");
  const statSupervisors = document.getElementById("statSupervisors");
  const statShiftRecords = document.getElementById("statShiftRecords");
  const statRunRecords = document.getElementById("statRunRecords");
  dashboardDateInput.value = appState.dashboardDate || todayISO();
  dashboardShiftButtons.forEach((btn) => {
    const active = btn.dataset.dashShift === (appState.dashboardShift || "Day");
    btn.classList.toggle("active", active);
    btn.setAttribute("aria-pressed", String(active));
  });
  const dashboardRows = lineList.map((line) => computeLineMetrics(line, appState.dashboardDate || todayISO(), appState.dashboardShift || "Day"));
  renderTable("dashboardTable", DASHBOARD_COLUMNS, dashboardRows.map((row) => ({
    lineName: row.lineName,
    date: row.date,
    shift: row.shift,
    units: row.units,
    totalDowntime: row.totalDowntime,
    lineUtil: row.lineUtil,
    netRunRate: row.netRunRate,
    bottleneckStageName: row.bottleneckStageName,
    staffing: row.staffingCallout
  })), {
    Line: "lineName",
    Date: "date",
    Shift: "shift",
    Units: "units",
    "Downtime (min)": "totalDowntime",
    "Utilisation (%)": "lineUtil",
    "Net Run Rate (u/min)": "netRunRate",
    Bottleneck: "bottleneckStageName",
    Staffing: "staffing"
  });
  cards.innerHTML = lineList
    .map((line) => {
      const stages = line.stages?.length || STAGES.length;
      const shifts = line.shiftRows?.length || 0;
      const assignedSupCount = (appState.supervisors || []).filter((sup) => (sup.assignedLineIds || []).includes(line.id)).length;
      return `
        <article class="line-card">
          <h3>${line.name}</h3>
          <p>${stages} stages | ${shifts} shift records | ${assignedSupCount} supervisors assigned</p>
          <div class="line-card-actions">
            <button type="button" data-open-line="${line.id}">Open Line</button>
            <button type="button" class="danger" data-delete-line="${line.id}">Delete</button>
          </div>
        </article>
      `;
    })
    .join("");
  if (statTotalLines) statTotalLines.textContent = formatNum(lineList.length, 0);
  if (statSupervisors) statSupervisors.textContent = formatNum((appState.supervisors || []).length, 0);
  if (statShiftRecords) statShiftRecords.textContent = formatNum(lineList.reduce((sum, line) => sum + (line.shiftRows?.length || 0), 0), 0);
  if (statRunRecords) statRunRecords.textContent = formatNum(lineList.reduce((sum, line) => sum + (line.runRows?.length || 0), 0), 0);

  if (!isSupervisor) {
    if (homeSidebar) homeSidebar.classList.remove("open");
    if (sidebarBackdrop) sidebarBackdrop.classList.add("hidden");
  }

  if (!isSupervisor) return;

  const assignedIds = session?.assignedLineIds?.filter((id) => appState.lines[id]) || [];
  loginSection.classList.toggle("hidden", Boolean(session));
  appSection.classList.toggle("hidden", !session);
  if (!session) return;

  const activeMainTab = ["supervisorVisual", "supervisorData"].includes(appState.supervisorMainTab) ? appState.supervisorMainTab : "supervisorVisual";
  superMainTabBtns.forEach((btn) => {
    const active = btn.dataset.superMainTab === activeMainTab;
    btn.classList.toggle("active", active);
    btn.setAttribute("aria-pressed", String(active));
  });
  superMainPanels.forEach((panel) => {
    panel.classList.toggle("active", panel.id === activeMainTab);
  });

  if (!assignedIds.includes(appState.supervisorSelectedLineId)) {
    appState.supervisorSelectedLineId = assignedIds[0] || "";
  }
  if (!appState.supervisorSelectedDate) appState.supervisorSelectedDate = todayISO();
  if (!["Day", "Night"].includes(appState.supervisorSelectedShift)) appState.supervisorSelectedShift = "Day";

  const activeSupervisor = supervisorByUsername(session.username);
  welcome.textContent = `Logged in as ${activeSupervisor?.name || session.username}`;
  supervisorMobileModeBtn.classList.toggle("hidden", false);
  appSection.classList.toggle("mobile-mode", Boolean(appState.supervisorMobileMode));
  supervisorMobileModeBtn.classList.toggle("active", Boolean(appState.supervisorMobileMode));
  supervisorMobileModeBtn.textContent = appState.supervisorMobileMode ? "Mobile Mode On" : "Mobile Mode";
  lineSelect.innerHTML = assignedIds.length
    ? assignedIds.map((id) => `<option value="${id}">${appState.lines[id].name}</option>`).join("")
    : `<option value="">No assigned lines</option>`;
  lineSelect.value = appState.supervisorSelectedLineId || "";
  svDateInput.value = appState.supervisorSelectedDate;
  svShiftButtons.forEach((btn) => {
    const active = btn.dataset.svShift === appState.supervisorSelectedShift;
    btn.classList.toggle("active", active);
    btn.setAttribute("aria-pressed", String(active));
  });
  downEquipment.innerHTML = supervisorEquipmentOptions(selectedSupervisorLine());
  renderSupervisorVisualiser(selectedSupervisorLine(), appState.supervisorSelectedDate, appState.supervisorSelectedShift);

  if (!shiftDateInput.value) shiftDateInput.value = todayISO();
  if (!runDateInput.value) runDateInput.value = todayISO();
  if (!downDateInput.value) downDateInput.value = todayISO();

  const activeTab = ["superShift", "superRun", "superDown"].includes(appState.supervisorTab) ? appState.supervisorTab : "superShift";
  document.querySelectorAll("[data-supervisor-tab]").forEach((btn) => {
    const active = btn.dataset.supervisorTab === activeTab;
    btn.classList.toggle("active", active);
    btn.setAttribute("aria-pressed", String(active));
  });
  document.querySelectorAll("#supervisorAppSection .data-section").forEach((section) => {
    section.classList.toggle("active", section.id === activeTab);
  });

  const logs = assignedIds
    .flatMap((id) => {
      const line = appState.lines[id];
      const shiftItems = (line.shiftRows || [])
        .filter((row) => row.submittedAt)
        .map((row) => ({
          lineName: line.name,
          date: row.date,
          shift: row.shift,
          type: "Shift",
          summary: `${row.startTime || "-"} to ${row.finishTime || "-"}`,
          supervisor: row.submittedBy || "-",
          createdAt: row.submittedAt
        }));
      const runItems = (line.runRows || [])
        .filter((row) => row.submittedAt)
        .map((row) => ({
          lineName: line.name,
          date: row.date,
          shift: row.shift,
          type: "Run",
          summary: `${row.product || "-"} (${formatNum(row.unitsProduced, 0)} units)`,
          supervisor: row.submittedBy || "-",
          createdAt: row.submittedAt
        }));
      const downItems = (line.downtimeRows || [])
        .filter((row) => row.submittedAt)
        .map((row) => ({
          lineName: line.name,
          date: row.date,
          shift: row.shift,
          type: "Downtime",
          summary: `${stageNameById(row.equipment)}: ${row.reason || "-"}`,
          supervisor: row.submittedBy || "-",
          createdAt: row.submittedAt
        }));
      return [...shiftItems, ...runItems, ...downItems];
    })
    .sort((a, b) => String(b.createdAt || "").localeCompare(String(a.createdAt || "")))
    .slice(0, 20);

  entryList.innerHTML = logs.length
    ? `
      <table>
        <thead><tr><th>Line</th><th>Date</th><th>Shift</th><th>Type</th><th>Details</th><th>By</th></tr></thead>
        <tbody>
          ${logs
            .map(
              (log) =>
                `<tr><td>${log.lineName}</td><td>${log.date}</td><td>${log.shift}</td><td>${log.type}</td><td>${log.summary}</td><td>${log.supervisor}</td></tr>`
            )
            .join("")}
        </tbody>
      </table>
    `
    : `<p class="muted">No supervisor submissions yet.</p>`;

  entryCards.innerHTML = logs.length
    ? logs
        .map(
          (log) => `
          <article class="entry-card">
            <h4>${log.type} | ${log.lineName}</h4>
            <p>${log.date} ${log.shift} | ${log.summary}</p>
            <p>By ${log.supervisor}</p>
          </article>
        `
        )
        .join("")
    : `<p class="muted">No supervisor submissions yet.</p>`;
}

function renderAll() {
  if (!state || !appState.lines[appState.activeLineId]) {
    const first = Object.keys(appState.lines)[0];
    if (first) {
      appState.activeLineId = first;
      state = appState.lines[first];
    }
  }

  document.getElementById("homeView").classList.toggle("hidden", appState.activeView === "line");
  document.getElementById("lineWorkspace").classList.toggle("hidden", appState.activeView !== "line");
  renderHome();

  if (appState.activeView !== "line" || !state) return;
  document.getElementById("appTitle").textContent = state.name || "Production Line";
  const editBtn = document.getElementById("toggleLayoutEdit");
  const addLineBtn = document.getElementById("addFlowLine");
  const addArrowBtn = document.getElementById("addFlowArrow");
  const uploadShapeBtn = document.getElementById("uploadFlowShapeBtn");
  if (editBtn) {
    const active = Boolean(state.visualEditMode);
    editBtn.classList.toggle("active", active);
    editBtn.textContent = active ? "Done Editing" : "Edit Layout";
    if (addLineBtn) {
      addLineBtn.disabled = !active;
      addLineBtn.classList.toggle("hidden", !active);
    }
    if (addArrowBtn) {
      addArrowBtn.disabled = !active;
      addArrowBtn.classList.toggle("hidden", !active);
    }
    if (uploadShapeBtn) {
      uploadShapeBtn.disabled = !active;
      uploadShapeBtn.classList.toggle("hidden", !active);
    }
  }
  const equip = document.getElementById("downtimeEquipment");
  if (equip) equip.innerHTML = stageOptionListHTML();
  document.getElementById("selectedDate").value = state.selectedDate;
  setShiftToggleUI();
  setActiveDataSubtab();
  renderCrewInputs();
  renderThroughputInputs();
  renderTrackingTables();
  renderAuditTrail();
  renderVisualiser();
}

function setActiveDataSubtab() {
  const activeId = state.activeDataTab || "dataShift";
  document.querySelectorAll(".data-subtab-btn").forEach((btn) => {
    const active = btn.dataset.dataTab === activeId;
    btn.classList.toggle("active", active);
    btn.setAttribute("aria-pressed", String(active));
  });
  document.querySelectorAll(".data-section").forEach((section) => {
    section.classList.toggle("active", section.id === activeId);
  });
}

bindTabs();
bindHome();
bindDataSubtabs();
bindVisualiserControls();
bindTrendModal();
bindStageSettingsModal();
bindForms();
bindDataControls();
renderAll();
