const STORAGE_KEY = "kebab-line-data-v2";
const STORAGE_BACKUP_KEY = "kebab-line-data-v2-backup";
const AUTH_STORAGE_KEY = "kebab-line-auth-v1";
const ROUTE_STORAGE_KEY = "kebab-line-route-v1";
const API_BASE_URL = `${
  window.PRODUCTION_LINE_API_BASE ||
  "http://localhost:4000"
}`.replace(/\/+$/, "");
const API_REQUEST_TIMEOUT_MS = 15000;
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

const STAGES = [
  { id: "s1", name: "1. Tumbler", crew: 2, group: "prep", match: ["tumbler"], x: 66.5, y: 64, w: 10, h: 26 },
  { id: "s2", name: "2. Transfer", crew: 1, group: "prep", match: ["transfer"], x: 56.5, y: 74, w: 7.5, h: 10, kind: "transfer" },
  { id: "s3", name: "3. Kebab Box Pack", crew: 3, group: "prep", match: ["box", "pack"], x: 23, y: 49, w: 14, h: 16 },
  { id: "s4", name: "4. Kebab Box Cut", crew: 1, group: "prep", match: ["box", "cut"], x: 13.5, y: 49, w: 9.5, h: 16 },
  { id: "s5", name: "5. Kebab Box Unload", crew: 1, group: "prep", match: ["box", "unload"], x: 5, y: 49, w: 8.5, h: 16 },
  { id: "s6", name: "6. Transfer", crew: 1, group: "prep", match: ["transfer"], x: 0.8, y: 34, w: 7.5, h: 10, kind: "transfer" },
  { id: "s7", name: "7. Kebab Split", crew: 2, group: "main", match: ["split"], x: 4.5, y: 14, w: 8.5, h: 15 },
  { id: "s8", name: "8. Marinate", crew: 10, group: "main", match: ["marinate"], x: 13.5, y: 14, w: 15, h: 15 },
  { id: "s9", name: "9. Wipe", crew: 1, group: "main", match: ["wipe"], x: 33.5, y: 14, w: 6, h: 15 },
  { id: "s10", name: "10. Proseal", crew: 1, group: "main", match: ["proseal"], x: 40.5, y: 14, w: 9.5, h: 15 },
  { id: "s11", name: "11. Metal Detector", crew: 1, group: "main", match: ["metal"], x: 55.5, y: 14, w: 6.8, h: 15 },
  { id: "s12", name: "12. Bottom Labeller", crew: 1, group: "main", match: ["bottom", "labeller", "labeler"], x: 69.5, y: 14, w: 9.5, h: 15 },
  { id: "s13", name: "13. Top Labeller", crew: 1, group: "main", match: ["top", "labeller", "labeler"], x: 80.5, y: 14, w: 8, h: 15 },
  { id: "s14", name: "14. Pack", crew: 2, group: "main", match: ["pack"], x: 90, y: 14, w: 5.5, h: 15 }
];

const PERMANENT_STAGE_TEMPLATE = Object.freeze({
  tumbler: { crew: 2, maxThroughput: 50 },
  transfer: { crew: 1, maxThroughput: 100 },
  "kebab box pack": { crew: 3, maxThroughput: 15 },
  "kebab box cut": { crew: 1, maxThroughput: 50 },
  "kebab box unload": { crew: 1, maxThroughput: 100 },
  "kebab split": { crew: 2, maxThroughput: 25 },
  marinate: { crew: 10, maxThroughput: 2.2 },
  wipe: { crew: 1, maxThroughput: 25 },
  proseal: { crew: 1, maxThroughput: 35 },
  "metal detector": { crew: 1, maxThroughput: 50 },
  "bottom labeller": { crew: 1, maxThroughput: 50 },
  "top labeller": { crew: 1, maxThroughput: 50 },
  pack: { crew: 2, maxThroughput: 25 }
});

const SHIFT_COLUMNS = ["Date", "Shift", "Crew On Shift", "Start Time", "Finish Time", "Break Count", "Break Time (min)", "Total Shift Time", "Submitted By", "Action"];
const RUN_COLUMNS = ["Date", "Assigned Shift", "Product", "Production Start Time", "Finish Time", "Units Produced", "Gross Production Time", "Associated Down Time", "Net Production Time", "Gross Run Rate", "Net Run Rate", "Submitted By", "Action"];
const DOWN_COLUMNS = ["Date", "Assigned Shift", "Downtime Start", "Downtime Finish", "Downtime (mins)", "Equipment", "Reason", "Submitted By", "Action"];
const AUDIT_COLUMNS = ["When", "Actor", "Action", "Details"];
const DASHBOARD_COLUMNS = ["Line", "Date", "Shift", "Units", "Downtime (min)", "Utilisation (%)", "Net Run Rate (u/min)", "Bottleneck", "Staffing"];
const SHIFT_OPTIONS = ["Day", "Night", "Full Day"];
const LINE_TREND_RANGES = ["day", "week", "month", "quarter"];
const LOG_ASSIGNABLE_SHIFTS = ["Day", "Night"];
const SUPERVISOR_SHIFT_OPTIONS = ["Day", "Night"];
const DOWNTIME_REASON_PRESETS = {
  "Donor Meat": ["Stock Out", "Late Delivery", "Quality Hold", "Temperature Hold"],
  People: ["Understaffed", "Training", "Handover Delay", "Absence"],
  Materials: ["Film Shortage", "Label Shortage", "Tray Shortage", "Marinade Shortage", "Skewer shortage"],
  Other: ["Cleaning", "QA Hold", "Power", "Unplanned Stop"]
};
const DEFAULT_SUPERVISORS = [
  { id: "sup-1", name: "Supervisor", username: "supervisor", password: "supervisor", mode: "all", shifts: ["Day", "Night"] },
  { id: "sup-2", name: "Day Lead", username: "daylead", password: "day123", mode: "even", shifts: ["Day"] },
  { id: "sup-3", name: "Night Lead", username: "nightlead", password: "night123", mode: "odd", shifts: ["Night"] }
];

let appState = loadState();
let state = appState.lines[appState.activeLineId] || null;
let visualiserDragMoved = false;
let managerLogInlineEdit = { lineId: "", type: "", logId: "" };
let supervisorShiftTileEditId = "";
let runCrewingPatternModalState = null;
appState.supervisors = normalizeSupervisors(appState.supervisors, appState.lines);
let lineModelSyncTimer = null;
let hostedRefreshErrorShown = false;
let hostedDatabaseAvailable = true;
let dbLoadingRequestCount = 0;
let dbLoadingFailsafeTimer = null;
let managerBackendSession = {
  backendToken: "",
  backendLineMap: {},
  backendStageMap: {},
  role: "manager"
};
clearLegacyStateStorage();
restoreAuthSessionsFromStorage();
restoreRouteFromHash();
state = appState.lines[appState.activeLineId] || null;

function setManagerLogInlineEdit(lineId = "", type = "", logId = "") {
  managerLogInlineEdit = {
    lineId: String(lineId || ""),
    type: String(type || ""),
    logId: String(logId || "")
  };
}

function clearManagerLogInlineEdit() {
  setManagerLogInlineEdit("", "", "");
}

function isManagerLogInlineEditRow(lineId, type, logId) {
  return (
    managerLogInlineEdit.lineId === String(lineId || "") &&
    managerLogInlineEdit.type === String(type || "") &&
    managerLogInlineEdit.logId === String(logId || "")
  );
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function readRouteStorage() {
  try {
    const raw = window.localStorage.getItem(ROUTE_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : null;
  } catch {
    return null;
  }
}

function writeRouteStorage(value) {
  try {
    if (value && typeof value === "object") {
      window.localStorage.setItem(ROUTE_STORAGE_KEY, JSON.stringify(value));
    } else {
      window.localStorage.removeItem(ROUTE_STORAGE_KEY);
    }
  } catch {
    // Ignore storage failures (private mode, quota, disabled storage).
  }
}

function applyRouteSnapshot(route) {
  if (!route || typeof route !== "object") return;
  const mode = route.mode;
  if (mode === "manager" || mode === "supervisor") appState.appMode = mode;
  const view = route.view;
  if (view === "home" || view === "line") appState.activeView = view;
  const lineId = String(route.lineId || "");
  if (lineId) appState.activeLineId = lineId;

  if (route.supervisorMainTab === "supervisorDay" || route.supervisorMainTab === "supervisorData") {
    appState.supervisorMainTab = route.supervisorMainTab;
  }
  const supervisorTab = String(route.supervisorTab || "");
  if (["superShift", "superRun", "superDown", "superVisual"].includes(supervisorTab)) {
    appState.supervisorTab = supervisorTab;
  }
  const supervisorSelectedLineId = String(route.supervisorSelectedLineId || "");
  if (supervisorSelectedLineId) appState.supervisorSelectedLineId = supervisorSelectedLineId;
  const supervisorSelectedDate = String(route.supervisorSelectedDate || "");
  if (/^\d{4}-\d{2}-\d{2}$/.test(supervisorSelectedDate)) appState.supervisorSelectedDate = supervisorSelectedDate;
  const supervisorSelectedShift = String(route.supervisorSelectedShift || "");
  if (SHIFT_OPTIONS.includes(supervisorSelectedShift)) appState.supervisorSelectedShift = supervisorSelectedShift;
  const dashboardDate = String(route.dashboardDate || "");
  if (/^\d{4}-\d{2}-\d{2}$/.test(dashboardDate)) appState.dashboardDate = dashboardDate;
  const dashboardShift = String(route.dashboardShift || "");
  if (SHIFT_OPTIONS.includes(dashboardShift)) appState.dashboardShift = dashboardShift;
}

function currentRouteSnapshot() {
  return {
    mode: appState.appMode === "supervisor" ? "supervisor" : "manager",
    view: appState.activeView === "line" ? "line" : "home",
    lineId: appState.activeView === "line" ? String(appState.activeLineId || "") : "",
    supervisorMainTab: appState.supervisorMainTab || "supervisorDay",
    supervisorTab: appState.supervisorTab || "superShift",
    supervisorSelectedLineId: String(appState.supervisorSelectedLineId || ""),
    supervisorSelectedDate: appState.supervisorSelectedDate || todayISO(),
    supervisorSelectedShift: appState.supervisorSelectedShift || "Full Day",
    dashboardDate: appState.dashboardDate || todayISO(),
    dashboardShift: appState.dashboardShift || "Day"
  };
}

function restoreRouteFromHash() {
  const raw = String(window.location.hash || "").replace(/^#/, "");
  if (!raw) {
    applyRouteSnapshot(readRouteStorage());
    return;
  }
  const params = new URLSearchParams(raw);
  applyRouteSnapshot({
    mode: params.get("mode"),
    view: params.get("view"),
    lineId: params.get("line"),
    supervisorMainTab: params.get("smt"),
    supervisorTab: params.get("st"),
    supervisorSelectedLineId: params.get("sl"),
    supervisorSelectedDate: params.get("sd"),
    supervisorSelectedShift: params.get("ss"),
    dashboardDate: params.get("dd"),
    dashboardShift: params.get("ds")
  });
}

function syncRouteToHash() {
  const route = currentRouteSnapshot();
  const params = new URLSearchParams();
  params.set("mode", route.mode);
  params.set("view", route.view);
  if (route.view === "line" && route.lineId) params.set("line", route.lineId);
  if (route.mode === "supervisor") {
    params.set("smt", route.supervisorMainTab);
    params.set("st", route.supervisorTab);
    if (route.supervisorSelectedLineId) params.set("sl", route.supervisorSelectedLineId);
    if (route.supervisorSelectedDate) params.set("sd", route.supervisorSelectedDate);
    if (route.supervisorSelectedShift) params.set("ss", route.supervisorSelectedShift);
  } else {
    if (route.dashboardDate) params.set("dd", route.dashboardDate);
    if (route.dashboardShift) params.set("ds", route.dashboardShift);
  }
  const nextHash = `#${params.toString()}`;
  if (window.location.hash !== nextHash) history.replaceState(null, "", nextHash);
  writeRouteStorage(route);
}

function generateSecretKey() {
  const alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  let key = "";
  for (let i = 0; i < 8; i += 1) key += alphabet[Math.floor(Math.random() * alphabet.length)];
  return key;
}

function stageTemplateDefaults(stage) {
  const key = stageNameCore(stage?.name || "");
  return PERMANENT_STAGE_TEMPLATE[key] || null;
}

function defaultCrewForStage(stage) {
  const templated = stageTemplateDefaults(stage);
  if (templated && Number.isFinite(templated.crew)) return Math.max(0, num(templated.crew));
  return Math.max(0, num(stage?.crew));
}

function defaultMaxThroughputForStage(stage) {
  const templated = stageTemplateDefaults(stage);
  if (templated && Number.isFinite(templated.maxThroughput)) return Math.max(0, num(templated.maxThroughput));
  return stage?.kind === "transfer" ? 3 : 2;
}

function defaultStageCrew(stages = STAGES) {
  return Object.fromEntries(stages.map((stage) => [stage.id, { crew: defaultCrewForStage(stage) }]));
}

function defaultStageSettings(stages = STAGES) {
  return Object.fromEntries(
    stages.map((stage) => [
      stage.id,
      {
        maxThroughput: defaultMaxThroughputForStage(stage)
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
      day[stage.id] = { crew: num(parsed.crewsByShift.Day?.[stage.id]?.crew ?? defaultCrewForStage(stage)) };
      night[stage.id] = { crew: num(parsed.crewsByShift.Night?.[stage.id]?.crew ?? defaultCrewForStage(stage)) };
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

function enforcePermanentCrewAndThroughput(line) {
  if (!line || !Array.isArray(line.stages)) return;
  if (!line.crewsByShift || typeof line.crewsByShift !== "object") line.crewsByShift = {};
  if (!line.crewsByShift.Day || typeof line.crewsByShift.Day !== "object") line.crewsByShift.Day = {};
  if (!line.crewsByShift.Night || typeof line.crewsByShift.Night !== "object") line.crewsByShift.Night = {};
  if (!line.stageSettings || typeof line.stageSettings !== "object") line.stageSettings = {};

  line.stages.forEach((stage) => {
    const templated = stageTemplateDefaults(stage);
    if (!templated) return;
    const crew = defaultCrewForStage(stage);
    const maxThroughput = defaultMaxThroughputForStage(stage);
    stage.crew = crew;
    line.crewsByShift.Day[stage.id] = { ...(line.crewsByShift.Day[stage.id] || {}), crew };
    line.crewsByShift.Night[stage.id] = { ...(line.crewsByShift.Night[stage.id] || {}), crew };
    line.stageSettings[stage.id] = { ...(line.stageSettings[stage.id] || {}), maxThroughput };
  });
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

function normalizeSupervisorShifts(shifts, { fallbackToAll = true } = {}) {
  if (!Array.isArray(shifts)) return fallbackToAll ? SUPERVISOR_SHIFT_OPTIONS.slice() : [];
  const unique = Array.from(
    new Set(
      shifts.map((shift) => (shift === "Night" ? "Night" : shift === "Day" ? "Day" : shift === "Full Day" ? "Full Day" : ""))
    )
  )
    .filter(Boolean);
  return unique;
}

function normalizeSupervisorLineShifts(assignedLineShifts, lines, legacyLineIds = []) {
  const normalized = {};
  if (assignedLineShifts && typeof assignedLineShifts === "object" && !Array.isArray(assignedLineShifts)) {
    Object.entries(assignedLineShifts).forEach(([lineId, shifts]) => {
      if (!lines[lineId]) return;
      const nextShifts = normalizeSupervisorShifts(shifts, { fallbackToAll: false });
      if (nextShifts.length) normalized[lineId] = nextShifts;
    });
  }
  if (!Object.keys(normalized).length && Array.isArray(legacyLineIds)) {
    legacyLineIds.forEach((lineId) => {
      if (lines[lineId]) normalized[lineId] = SUPERVISOR_SHIFT_OPTIONS.slice();
    });
  }
  return normalized;
}

function defaultSupervisors(lines) {
  return DEFAULT_SUPERVISORS.map((sup) => {
    const lineIds = supervisorModeAssignments(sup.mode, lines);
    const lineShifts = Object.fromEntries(
      lineIds.map((lineId) => [lineId, normalizeSupervisorShifts(sup.shifts, { fallbackToAll: true })])
    );
    return {
      id: sup.id,
      name: sup.name,
      username: sup.username.toLowerCase(),
      password: sup.password,
      assignedLineIds: lineIds,
      assignedLineShifts: lineShifts
    };
  });
}

function normalizeSupervisors(supervisors, lines) {
  const source = Array.isArray(supervisors) && supervisors.length ? supervisors : defaultSupervisors(lines);
  const seen = new Set();
  return source
    .map((sup, index) => {
      const username = String(sup?.username || "").trim().toLowerCase();
      if (!username || seen.has(username)) return null;
      seen.add(username);
      const assignedLineShifts = normalizeSupervisorLineShifts(sup?.assignedLineShifts, lines, sup?.assignedLineIds || []);
      return {
        id: sup?.id || `sup-${Date.now()}-${index}`,
        name: String(sup?.name || username).trim() || username,
        username,
        password: String(sup?.password || ""),
        assignedLineIds: Object.keys(assignedLineShifts),
        assignedLineShifts
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
  const assignedLineShifts = sup
    ? normalizeSupervisorLineShifts(sup.assignedLineShifts, lines, sup.assignedLineIds || [])
    : normalizeSupervisorLineShifts(session.assignedLineShifts, lines, session.assignedLineIds || []);
  const assignedLineIds = Object.keys(assignedLineShifts);
  if (!sup && !assignedLineIds.length && !session.backendToken) return null;
  const backendLineMap = {};
  if (session.backendLineMap && typeof session.backendLineMap === "object") {
    Object.entries(session.backendLineMap).forEach(([localId, backendId]) => {
      if (lines[localId] && UUID_RE.test(String(backendId || ""))) backendLineMap[localId] = String(backendId);
    });
  }
  return {
    username: sup?.username || String(session.username || "").trim().toLowerCase(),
    assignedLineIds,
    assignedLineShifts,
    backendToken: typeof session.backendToken === "string" && session.backendToken ? session.backendToken : "",
    backendLineMap
  };
}

function selectedSupervisorLineId() {
  const sel = document.getElementById("supervisorLineSelect");
  return sel?.value || appState.supervisorSelectedLineId || "";
}

function selectedSupervisorLine() {
  const id = selectedSupervisorLineId();
  return id && appState.lines[id] ? appState.lines[id] : null;
}

function supervisorCanAccessShift(session, lineId, shift) {
  if (!shift || !lineId) return false;
  const lineShiftMap = normalizeSupervisorLineShifts(session?.assignedLineShifts, appState.lines || {}, session?.assignedLineIds || []);
  const allowed = expandedSupervisorShiftAccess(lineShiftMap[lineId]);
  return allowed.includes(shift);
}

function expandedSupervisorShiftAccess(shifts) {
  const base = normalizeSupervisorShifts(shifts, { fallbackToAll: false });
  const next = Array.from(new Set(base));
  if (next.includes("Day") && next.includes("Night") && !next.includes("Full Day")) {
    next.push("Full Day");
  }
  return next;
}

function stageMaxThroughputForLine(line, stageId) {
  return Math.max(0, num(line?.stageSettings?.[stageId]?.maxThroughput));
}

function stageCrewForShiftForLine(line, stageId, shift) {
  if (isFullDayShift(shift)) {
    const dayCrew = num(line?.crewsByShift?.Day?.[stageId]?.crew);
    const nightCrew = num(line?.crewsByShift?.Night?.[stageId]?.crew);
    const stage = (line?.stages || []).find((s) => s.id === stageId);
    const crew = Math.max(dayCrew, nightCrew);
    if (crew > 0) return crew;
    if (stage?.kind === "transfer") return 1;
    return 0;
  }
  const crew = num(line?.crewsByShift?.[shift]?.[stageId]?.crew);
  const stage = (line?.stages || []).find((s) => s.id === stageId);
  if (crew > 0) return crew;
  if (stage?.kind === "transfer") return 1;
  return 0;
}

function stageTotalMaxThroughputForLine(line, stageId, shift) {
  return stageMaxThroughputForLine(line, stageId) * stageCrewForShiftForLine(line, stageId, shift);
}

function isCrewedStage(stage) {
  if (!stage) return false;
  const kind = String(stage.kind || "").toLowerCase();
  const group = String(stage.group || "").toLowerCase();
  return kind !== "transfer" && group !== "transfer";
}

function crewedStagesForLine(line) {
  return (line?.stages || [])
    .map((stage, index) => ({ stage, index }))
    .filter(({ stage }) => isCrewedStage(stage));
}

function parseRunCrewingPattern(value) {
  if (value && typeof value === "object" && !Array.isArray(value)) return value;
  const raw = String(value || "").trim();
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
  } catch {
    return {};
  }
}

function normalizeRunCrewingPattern(pattern, line, shift, { fallbackToIdeal = false } = {}) {
  const source = parseRunCrewingPattern(pattern);
  const normalized = {};
  crewedStagesForLine(line).forEach(({ stage }) => {
    const hasValue = Object.prototype.hasOwnProperty.call(source, stage.id);
    if (!hasValue) {
      if (fallbackToIdeal) {
        normalized[stage.id] = Math.max(0, Math.floor(num(stageCrewForShiftForLine(line, stage.id, shift))));
      }
      return;
    }
    normalized[stage.id] = Math.max(0, Math.floor(num(source[stage.id])));
  });
  return normalized;
}

function runCrewingPatternFromInput(inputEl, line, shift, { fallbackToIdeal = false } = {}) {
  const raw = inputEl ? inputEl.value : "";
  return normalizeRunCrewingPattern(raw, line, shift, { fallbackToIdeal });
}

function runCrewingPatternSetCount(pattern, line) {
  const stageIds = new Set(crewedStagesForLine(line).map(({ stage }) => stage.id));
  return Object.keys(pattern || {}).filter((stageId) => stageIds.has(stageId)).length;
}

function runCrewingPatternTotalCrew(pattern = {}) {
  return Object.values(pattern || {}).reduce((sum, value) => sum + Math.max(0, Math.floor(num(value))), 0);
}

function runCrewingPatternSummaryText(pattern, line) {
  const stageCount = crewedStagesForLine(line).length;
  if (!stageCount) return "No crewed stages on this line.";
  const setCount = runCrewingPatternSetCount(pattern, line);
  if (!setCount) return "No crewing pattern set.";
  const totalCrew = runCrewingPatternTotalCrew(pattern);
  return `${setCount}/${stageCount} stages set (${formatNum(totalCrew, 0)} crew total)`;
}

function setRunCrewingPatternField(inputEl, summaryEl, line, shift, pattern, { fallbackToIdeal = false } = {}) {
  const normalized = normalizeRunCrewingPattern(pattern, line, shift, { fallbackToIdeal });
  if (inputEl) inputEl.value = Object.keys(normalized).length ? JSON.stringify(normalized) : "";
  if (summaryEl) summaryEl.textContent = runCrewingPatternSummaryText(normalized, line);
  return normalized;
}

function runCrewingPatternModalLine() {
  const lineId = String(runCrewingPatternModalState?.lineId || "");
  if (!lineId) return null;
  return appState.lines[lineId] || (state?.id === lineId ? state : null);
}

function closeRunCrewingPatternModal() {
  const modal = document.getElementById("runCrewingPatternModal");
  if (!modal) return;
  modal.classList.remove("open");
  modal.setAttribute("aria-hidden", "true");
  runCrewingPatternModalState = null;
}

function renderRunCrewingPatternModalInputs() {
  const list = document.getElementById("runCrewingPatternList");
  const meta = document.getElementById("runCrewingPatternMeta");
  if (!list || !meta || !runCrewingPatternModalState) return;
  const line = runCrewingPatternModalLine();
  const shift = runCrewingPatternModalState.shift || "Day";
  const stageRows = crewedStagesForLine(line);
  meta.textContent = `Set crew values for each non-transfer stage on this ${shift} run.`;
  if (!line || !stageRows.length) {
    list.innerHTML = `<p class="muted">No crewed stages available for this line.</p>`;
    return;
  }
  const pattern = runCrewingPatternModalState.pattern || {};
  list.innerHTML = stageRows
    .map(({ stage, index }) => {
      const value = Math.max(0, Math.floor(num(pattern[stage.id])));
      return `
        <label class="run-crewing-stage-row">
          <span class="run-crewing-stage-name">${htmlEscape(stageDisplayName(stage, index))}</span>
          <input type="number" min="0" step="1" data-run-crewing-stage="${htmlEscape(stage.id)}" value="${value}" />
        </label>
      `;
    })
    .join("");
}

function openRunCrewingPatternModal({ line, shift = "Day", inputEl = null, summaryEl = null } = {}) {
  const modal = document.getElementById("runCrewingPatternModal");
  if (!modal || !line || !inputEl) return;
  const activeShift = String(shift || "Day");
  const existingPattern = runCrewingPatternFromInput(inputEl, line, activeShift, { fallbackToIdeal: false });
  const pattern = Object.keys(existingPattern).length
    ? existingPattern
    : normalizeRunCrewingPattern({}, line, activeShift, { fallbackToIdeal: true });
  runCrewingPatternModalState = {
    lineId: line.id,
    shift: activeShift,
    inputEl,
    summaryEl,
    pattern
  };
  renderRunCrewingPatternModalInputs();
  modal.classList.add("open");
  modal.setAttribute("aria-hidden", "false");
}

function saveRunCrewingPatternFromModal() {
  if (!runCrewingPatternModalState) return;
  const list = document.getElementById("runCrewingPatternList");
  const line = runCrewingPatternModalLine();
  const shift = runCrewingPatternModalState.shift || "Day";
  if (!list || !line) {
    closeRunCrewingPatternModal();
    return;
  }
  const pattern = {};
  Array.from(list.querySelectorAll("[data-run-crewing-stage]")).forEach((input) => {
    const stageId = String(input.getAttribute("data-run-crewing-stage") || "");
    if (!stageId) return;
    pattern[stageId] = Math.max(0, Math.floor(num(input.value)));
  });
  setRunCrewingPatternField(
    runCrewingPatternModalState.inputEl,
    runCrewingPatternModalState.summaryEl,
    line,
    shift,
    pattern,
    { fallbackToIdeal: false }
  );
  closeRunCrewingPatternModal();
}

function bindRunCrewingPatternModal() {
  const modal = document.getElementById("runCrewingPatternModal");
  const closeBtn = document.getElementById("closeRunCrewingPatternModal");
  const saveBtn = document.getElementById("saveRunCrewingPattern");
  if (!modal || !closeBtn || !saveBtn) return;

  closeBtn.addEventListener("click", closeRunCrewingPatternModal);
  saveBtn.addEventListener("click", saveRunCrewingPatternFromModal);
  modal.addEventListener("click", (event) => {
    if (event.target === modal) closeRunCrewingPatternModal();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && modal.classList.contains("open")) closeRunCrewingPatternModal();
  });
}

function refreshRunCrewingPatternSummaries() {
  const managerInput = document.getElementById("runCrewingPattern");
  const managerSummary = document.getElementById("runCrewingPatternSummary");
  if (managerInput && managerSummary && state) {
    setRunCrewingPatternField(managerInput, managerSummary, state, state.selectedShift || "Day", managerInput.value, { fallbackToIdeal: false });
  }

  const supervisorInput = document.getElementById("superRunCrewingPattern");
  const supervisorSummary = document.getElementById("superRunCrewingPatternSummary");
  if (!supervisorInput || !supervisorSummary) return;
  const line = selectedSupervisorLine() || (appState.supervisorSelectedLineId ? appState.lines[appState.supervisorSelectedLineId] : null);
  if (!line) {
    supervisorInput.value = "";
    supervisorSummary.textContent = "No crewing pattern set.";
    return;
  }
  setRunCrewingPatternField(
    supervisorInput,
    supervisorSummary,
    line,
    appState.supervisorSelectedShift || "Day",
    supervisorInput.value,
    { fallbackToIdeal: false }
  );
}

function derivedDataForLine(line) {
  const downtimeRows = (line?.downtimeRows || [])
    .map(computeDowntimeRow)
    .map((row) => decorateTimedLogShift(row, line, "downtimeStart", "downtimeFinish"));
  const shiftRows = (line?.shiftRows || []).map(computeShiftRow);
  const breakRows = (line?.breakRows || []).map(computeBreakRow);
  const downtimeByShift = new Map();
  downtimeRows.forEach((row) => {
    LOG_ASSIGNABLE_SHIFTS.forEach((shiftKey) => {
      const mins = Math.max(0, num(row?.__shiftMinutes?.[shiftKey]));
      if (mins <= 0) return;
      const key = `${row.date}__${shiftKey}`;
      downtimeByShift.set(key, (downtimeByShift.get(key) || 0) + mins);
    });
  });
  const runRows = (line?.runRows || [])
    .map((row) => decorateTimedLogShift(row, line, "productionStartTime", "finishTime"))
    .map((row) => computeRunRow(row, downtimeByShift));
  return { shiftRows, breakRows, runRows, downtimeRows };
}

function computeLineMetrics(line, date, shift) {
  const stages = line?.stages?.length ? line.stages : STAGES;
  const data = derivedDataForLine(line || {});
  const selectedRunRows = selectedShiftRowsByDate(data.runRows, date, shift, { line });
  const selectedDownRows = selectedShiftRowsByDate(data.downtimeRows, date, shift, { line });
  const selectedShiftRows = selectedShiftRowsByDate(data.shiftRows, date, shift, { line });
  const shiftMins = selectedShiftRows.reduce((sum, row) => sum + num(row.totalShiftTime), 0);

  let requiredCrew = requiredCrewForLineShift(line, shift);
  let crewOnShift = 0;
  let hasCrewLog = false;
  if (isFullDayShift(shift)) {
    const latestByShift = {};
    selectedShiftRows.forEach((row) => {
      const key = row.shift || "Day";
      if (!latestByShift[key]) latestByShift[key] = row;
      const prevAt = Date.parse(latestByShift[key]?.submittedAt || "") || 0;
      const nextAt = Date.parse(row?.submittedAt || "") || 0;
      if (nextAt >= prevAt) latestByShift[key] = row;
    });
    const dayRow = latestByShift.Day || null;
    const nightRow = latestByShift.Night || null;
    requiredCrew = requiredCrewForLineShift(line, "Day") + requiredCrewForLineShift(line, "Night");
    crewOnShift = Math.max(0, num(dayRow?.crewOnShift)) + Math.max(0, num(nightRow?.crewOnShift));
    hasCrewLog = Boolean(dayRow || nightRow);
  } else {
    const latestShiftRow = selectedShiftRows.length ? selectedShiftRows[selectedShiftRows.length - 1] : null;
    hasCrewLog = Boolean(latestShiftRow) && latestShiftRow?.crewOnShift !== "" && latestShiftRow?.crewOnShift !== undefined;
    crewOnShift = hasCrewLog ? Math.max(0, num(latestShiftRow?.crewOnShift)) : 0;
  }
  const understaffedBy = hasCrewLog ? Math.max(0, requiredCrew - crewOnShift) : 0;
  const staffingCallout = !hasCrewLog ? "No shift data" : understaffedBy > 0 ? `Understaffed by ${understaffedBy}` : "Fully staffed";
  const units = selectedRunRows.reduce((sum, row) => sum + num(row.unitsProduced) * timedLogShiftWeight(row, shift), 0);
  const totalDowntime = selectedDownRows.reduce((sum, row) => sum + num(row.downtimeMins) * timedLogShiftWeight(row, shift), 0);
  const totalNetTime = selectedRunRows.reduce((sum, row) => sum + num(row.netProductionTime) * timedLogShiftWeight(row, shift), 0);
  const netRunRate = totalNetTime > 0 ? units / totalNetTime : 0;
  let utilAccumulator = 0;
  let utilCount = 0;
  let bottleneckStageName = "-";
  let bottleneckUtil = -1;

  stages.forEach((stage, index) => {
    const stageDowntime = selectedDownRows
      .filter((row) => matchesStage(stage, row.equipment))
      .reduce((sum, row) => sum + num(row.downtimeMins) * timedLogShiftWeight(row, shift), 0);
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
  return {
    activeView: "home",
    appMode: "manager",
    dashboardDate: todayISO(),
    dashboardShift: "Day",
    supervisorMobileMode: false,
    supervisorMainTab: "supervisorDay",
    supervisorTab: "superShift",
    supervisorSelectedLineId: "",
    supervisorSelectedDate: todayISO(),
    supervisorSelectedShift: "Full Day",
    supervisorSession: null,
    supervisors: [],
    activeLineId: "",
    lines: {}
  };
}

function readAuthStorage() {
  try {
    const raw = window.localStorage.getItem(AUTH_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : null;
  } catch {
    return null;
  }
}

function writeAuthStorage(value) {
  try {
    if (value && typeof value === "object") {
      window.localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify(value));
    } else {
      window.localStorage.removeItem(AUTH_STORAGE_KEY);
    }
  } catch {
    // Ignore storage failures (private mode, quota, disabled storage).
  }
}

function clearLegacyStateStorage() {
  try {
    window.localStorage.removeItem(STORAGE_KEY);
    window.localStorage.removeItem(STORAGE_BACKUP_KEY);
  } catch {
    // Ignore storage failures (private mode, quota, disabled storage).
  }
}

function persistAuthSessions() {
  const managerToken = String(managerBackendSession?.backendToken || "").trim();
  const supervisorSession = appState?.supervisorSession;
  const supervisorToken = String(supervisorSession?.backendToken || "").trim();
  const supervisorUsername = String(supervisorSession?.username || "").trim().toLowerCase();
  if (!managerToken && !(supervisorToken && supervisorUsername)) {
    writeAuthStorage(null);
    return;
  }
  writeAuthStorage({
    manager: managerToken ? { backendToken: managerToken } : null,
    supervisor:
      supervisorToken && supervisorUsername
        ? {
            username: supervisorUsername,
            backendToken: supervisorToken,
            assignedLineIds: Array.isArray(supervisorSession?.assignedLineIds) ? supervisorSession.assignedLineIds : [],
            assignedLineShifts:
              supervisorSession?.assignedLineShifts && typeof supervisorSession.assignedLineShifts === "object"
                ? supervisorSession.assignedLineShifts
                : {},
            backendLineMap:
              supervisorSession?.backendLineMap && typeof supervisorSession.backendLineMap === "object"
                ? supervisorSession.backendLineMap
                : {}
          }
        : null
  });
}

function restoreAuthSessionsFromStorage() {
  const stored = readAuthStorage();
  if (!stored) return;
  const managerToken = String(stored?.manager?.backendToken || "").trim();
  if (managerToken) {
    managerBackendSession.backendToken = managerToken;
    managerBackendSession.role = "manager";
  }
  const supervisorToken = String(stored?.supervisor?.backendToken || "").trim();
  const supervisorUsername = String(stored?.supervisor?.username || "").trim().toLowerCase();
  if (!supervisorToken || !supervisorUsername) return;
  const supervisorSession = normalizeSupervisorSession(
    {
      username: supervisorUsername,
      assignedLineIds: Array.isArray(stored?.supervisor?.assignedLineIds) ? stored.supervisor.assignedLineIds : [],
      assignedLineShifts:
        stored?.supervisor?.assignedLineShifts && typeof stored.supervisor.assignedLineShifts === "object"
          ? stored.supervisor.assignedLineShifts
          : {},
      backendToken: supervisorToken,
      backendLineMap:
        stored?.supervisor?.backendLineMap && typeof stored.supervisor.backendLineMap === "object"
          ? stored.supervisor.backendLineMap
          : {},
      role: "supervisor"
    },
    appState.supervisors,
    appState.lines
  );
  appState.supervisorSession = supervisorSession || {
    username: supervisorUsername,
    assignedLineIds: [],
    assignedLineShifts: {},
    backendToken: supervisorToken,
    backendLineMap: {},
    role: "supervisor"
  };
}

function queueLineModelSync(lineId) {
  if (!lineId || !UUID_RE.test(String(lineId))) return;
  if (lineModelSyncTimer) clearTimeout(lineModelSyncTimer);
  lineModelSyncTimer = setTimeout(async () => {
    try {
      await saveLineModelToBackend(lineId);
    } catch (error) {
      console.warn("Backend line model sync failed:", error);
      if (document.visibilityState === "visible") {
        alert(`Could not save layout/settings changes to server.\n${error?.message || "Please retry."}`);
      }
    }
  }, 300);
}

function saveState(options = {}) {
  if (state && state.id) {
    appState.lines[state.id] = state;
    appState.activeLineId = state.id;
    if (options.syncModel) queueLineModelSync(state.id);
  }
  persistAuthSessions();
  syncRouteToHash();
}

function makeDefaultLine(id, name, { seedSample = false } = {}) {
  const stages = clone(STAGES);
  const line = {
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
    lineTrendRange: "day",
    crewsByShift: defaultCrewByShift(stages),
    stageSettings: defaultStageSettings(stages),
    stages,
    shiftRows: [],
    breakRows: [],
    runRows: [],
    downtimeRows: [],
    supervisorLogs: [],
    auditRows: []
  };
  enforcePermanentCrewAndThroughput(line);
  if (seedSample) loadSampleDataIntoLine(line);
  return line;
}

function normalizeLine(id, line) {
  const base = makeDefaultLine(id, line?.name || "Production Line");
  const stages = Array.isArray(line?.stages) && line.stages.length ? line.stages : clone(STAGES);
  const normalized = {
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
    breakRows: line?.breakRows || [],
    runRows: Array.isArray(line?.runRows)
      ? line.runRows.map((row) => ({
          ...row,
          runCrewingPattern: parseRunCrewingPattern(row?.runCrewingPattern),
          shift: ""
        }))
      : [],
    downtimeRows: Array.isArray(line?.downtimeRows)
      ? line.downtimeRows.map((row) => ({
          ...row,
          shift: ""
        }))
      : [],
    supervisorLogs: Array.isArray(line?.supervisorLogs) ? line.supervisorLogs : [],
    auditRows: Array.isArray(line?.auditRows) ? line.auditRows : []
  };
  enforcePermanentCrewAndThroughput(normalized);
  ensureManagerLogRowIds(normalized);
  return normalized;
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
        maxThroughput: Math.max(0, num(stage.maxThroughput) || defaultMaxThroughputForStage(stage))
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

function isPendingRunLogRow(row) {
  return strictTimeValid(row?.productionStartTime) && strictTimeValid(row?.finishTime) && row.productionStartTime === row.finishTime;
}

function isPendingShiftLogRow(row) {
  return strictTimeValid(row?.startTime) && strictTimeValid(row?.finishTime) && row.startTime === row.finishTime;
}

function isPendingBreakLogRow(row) {
  return strictTimeValid(row?.breakStart) && !strictTimeValid(row?.breakFinish);
}

function isPendingDowntimeLogRow(row) {
  return strictTimeValid(row?.downtimeStart) && strictTimeValid(row?.downtimeFinish) && row.downtimeStart === row.downtimeFinish;
}

function rowNewestSortValue(row, primaryTimeField = "") {
  const submittedAt = Date.parse(String(row?.submittedAt || ""));
  if (Number.isFinite(submittedAt) && submittedAt > 0) return submittedAt;
  const rowDate = String(row?.date || "").trim();
  const rowTime = strictTimeValid(row?.[primaryTimeField]) ? String(row[primaryTimeField]) : "00:00";
  if (/^\d{4}-\d{2}-\d{2}$/.test(rowDate)) {
    const derived = Date.parse(`${rowDate}T${rowTime}:00`);
    if (Number.isFinite(derived) && derived > 0) return derived;
  }
  return 0;
}

function managerLogSubmittedByLabel(row) {
  const submittedBy = String(row?.submittedBy || "").trim();
  if (!submittedBy) return "Manager";
  if (submittedBy.toLowerCase() === "manager") return "Manager";
  return submittedBy;
}

function breakRowsForShift(line, shiftLogId) {
  return (line?.breakRows || [])
    .filter((row) => String(row?.shiftLogId || "") === String(shiftLogId || ""))
    .slice()
    .sort((a, b) => {
      const startCmp = String(a?.breakStart || "").localeCompare(String(b?.breakStart || ""));
      if (startCmp !== 0) return startCmp;
      const submittedCmp = String(a?.submittedAt || "").localeCompare(String(b?.submittedAt || ""));
      if (submittedCmp !== 0) return submittedCmp;
      return String(a?.id || "").localeCompare(String(b?.id || ""));
    });
}

function nowTimeHHMM() {
  const now = new Date();
  const hh = String(now.getHours()).padStart(2, "0");
  const mm = String(now.getMinutes()).padStart(2, "0");
  return `${hh}:${mm}`;
}

function selectedSupervisorShiftLogId({ line = null } = {}) {
  const inputId = String(document.getElementById("superShiftLogId")?.value || "").trim();
  const tileId = String(supervisorShiftTileEditId || "").trim();
  const selectedId = inputId || tileId;
  if (!selectedId || !line) return selectedId;
  const row = (line.shiftRows || []).find((item) => String(item?.id || "") === selectedId);
  if (!row) return "";
  return selectedId;
}

function updateSupervisorProgressButtonLabels() {
  const shiftBtn = document.getElementById("superShiftSaveProgress");
  const runBtn = document.getElementById("superRunSaveProgress");
  const downBtn = document.getElementById("superDownSaveProgress");
  const shiftSelected = Boolean(selectedSupervisorShiftLogId());
  const runSelected = Boolean(String(document.getElementById("superRunLogId")?.value || "").trim());
  const downSelected = Boolean(String(document.getElementById("superDownLogId")?.value || "").trim());

  if (shiftBtn) shiftBtn.textContent = shiftSelected ? "Save" : "Start";
  if (runBtn) runBtn.textContent = runSelected ? "Save" : "Start";
  if (downBtn) downBtn.textContent = downSelected ? "Save" : "Start";
}

function syncNowInputPrefixState(input) {
  if (!input || typeof input.closest !== "function") return;
  const wrapper = input.closest(".input-now-wrap[data-time-prefix]");
  if (!wrapper) return;
  const hasValue = String(input.value || "").trim().length > 0;
  wrapper.classList.toggle("has-value", hasValue);
  if (input.style) {
    input.style.paddingLeft = hasValue ? "3.8rem" : "0.86rem";
    input.style.paddingRight = "4.15rem";
  }
}

function syncAllNowInputPrefixStates(root = document) {
  if (!root || typeof root.querySelectorAll !== "function") return;
  root.querySelectorAll(".input-now-wrap[data-time-prefix] input").forEach((input) => {
    syncNowInputPrefixState(input);
  });
}

function isFullDayShift(shift) {
  return String(shift || "") === "Full Day";
}

function shiftKeysForSelection(shift) {
  if (isFullDayShift(shift)) return ["Day", "Night", "Full Day"];
  return [shift];
}

function mergeMinuteIntervals(intervals = []) {
  return (intervals || [])
    .map((segment) => ({ start: Math.max(0, num(segment?.start)), end: Math.min(24 * 60, num(segment?.end)) }))
    .filter((segment) => Number.isFinite(segment.start) && Number.isFinite(segment.end) && segment.end > segment.start)
    .sort((a, b) => a.start - b.start)
    .reduce((merged, segment) => {
      const last = merged[merged.length - 1];
      if (!last || segment.start > last.end) {
        merged.push({ ...segment });
      } else if (segment.end > last.end) {
        last.end = segment.end;
      }
      return merged;
    }, []);
}

function intervalOverlapMinutes(a, b) {
  if (!a || !b) return 0;
  const start = Math.max(num(a.start), num(b.start));
  const end = Math.min(num(a.end), num(b.end));
  return end > start ? end - start : 0;
}

function intervalsFromTimes(startValue, finishValue) {
  const startMins = parseTimeToMinutes(startValue);
  const finishMins = parseTimeToMinutes(finishValue);
  if (!Number.isFinite(startMins) || !Number.isFinite(finishMins)) return [];
  return splitAcrossMidnight(startMins, finishMins)
    .map((segment) => ({ start: num(segment.start), end: num(segment.end) }))
    .filter((segment) => segment.end > segment.start);
}

function shiftIntervalsForDate(line, date) {
  const byShift = Object.fromEntries(LOG_ASSIGNABLE_SHIFTS.map((shift) => [shift, []]));
  (line?.shiftRows || [])
    .filter((row) => row?.date === date && LOG_ASSIGNABLE_SHIFTS.includes(String(row?.shift || "")))
    .forEach((row) => {
      const shiftKey = String(row.shift || "");
      let intervals = intervalsFromTimes(row.startTime, row.finishTime);
      if (!intervals.length && strictTimeValid(row.startTime) && strictTimeValid(row.finishTime) && row.startTime === row.finishTime) {
        const startMins = parseTimeToMinutes(row.startTime);
        const nowCandidate = nowTimeHHMM();
        const nowMins = parseTimeToMinutes(nowCandidate);
        if (Number.isFinite(startMins) && Number.isFinite(nowMins)) {
          const effectiveEnd = startMins === nowMins ? (startMins + 1) % (24 * 60) : nowMins;
          intervals = splitAcrossMidnight(startMins, effectiveEnd);
        }
      }
      byShift[shiftKey].push(...intervals);
    });
  LOG_ASSIGNABLE_SHIFTS.forEach((shiftKey) => {
    byShift[shiftKey] = mergeMinuteIntervals(byShift[shiftKey]);
  });
  return byShift;
}

function timedLogShiftLabelFromCoverage(coverage) {
  const matches = Array.isArray(coverage?.shiftMatches) ? coverage.shiftMatches.filter((shift) => LOG_ASSIGNABLE_SHIFTS.includes(shift)) : [];
  const unassignedMinutes = Math.max(0, num(coverage?.unassignedMinutes));
  if (!matches.length) return "Unassigned";
  const base = matches.join(" + ");
  if (unassignedMinutes > 0) return `${base} + Unassigned`;
  return base;
}

function shiftCoverageForTimedLog(line, date, startValue, finishValue) {
  const empty = {
    shiftMatches: [],
    shiftMinutes: Object.fromEntries(LOG_ASSIGNABLE_SHIFTS.map((shift) => [shift, 0])),
    totalMinutes: 0,
    unassignedMinutes: 0,
    label: "Unassigned"
  };
  if (!line || !/^\d{4}-\d{2}-\d{2}$/.test(String(date || ""))) return empty;
  const startMins = parseTimeToMinutes(startValue);
  const finishMins = parseTimeToMinutes(finishValue);
  if (!Number.isFinite(startMins) || !Number.isFinite(finishMins)) return empty;
  const eventSegments = intervalsFromTimes(startValue, finishValue);
  const isPointEvent = startMins === finishMins || !eventSegments.length;
  const shiftIntervals = shiftIntervalsForDate(line, date);
  const shiftMatches = new Set();
  const shiftMinutes = Object.fromEntries(LOG_ASSIGNABLE_SHIFTS.map((shift) => [shift, 0]));

  LOG_ASSIGNABLE_SHIFTS.forEach((shiftKey) => {
    const intervals = shiftIntervals[shiftKey] || [];
    if (isPointEvent) {
      if (intervals.some((interval) => startMins >= interval.start && startMins <= interval.end)) {
        shiftMatches.add(shiftKey);
      }
      return;
    }
    const overlap = eventSegments.reduce(
      (sum, eventSegment) => sum + intervals.reduce((inner, interval) => inner + intervalOverlapMinutes(eventSegment, interval), 0),
      0
    );
    if (overlap > 0) {
      shiftMinutes[shiftKey] = overlap;
      shiftMatches.add(shiftKey);
    }
  });

  const totalMinutes = isPointEvent ? 0 : eventSegments.reduce((sum, segment) => sum + Math.max(0, num(segment.end) - num(segment.start)), 0);
  const assignedMinutes = LOG_ASSIGNABLE_SHIFTS.reduce((sum, shiftKey) => sum + Math.max(0, num(shiftMinutes[shiftKey])), 0);
  const unassignedMinutes = isPointEvent ? 0 : Math.max(0, totalMinutes - assignedMinutes);
  const coverage = {
    shiftMatches: LOG_ASSIGNABLE_SHIFTS.filter((shiftKey) => shiftMatches.has(shiftKey)),
    shiftMinutes,
    totalMinutes,
    unassignedMinutes,
    label: ""
  };
  coverage.label = timedLogShiftLabelFromCoverage(coverage);
  return coverage;
}

function decorateTimedLogShift(row, line, startField, finishField) {
  const base = { ...(row || {}) };
  const coverage = shiftCoverageForTimedLog(line, base.date, base[startField], base[finishField]);
  const shiftMatches = coverage.shiftMatches;
  const shiftMinutes = coverage.shiftMinutes;
  const totalMinutes = Math.max(0, num(coverage.totalMinutes));
  const unassignedMinutes = Math.max(0, num(coverage.unassignedMinutes));
  const label = coverage.label;
  return {
    ...base,
    shift: "",
    __shiftMatches: shiftMatches,
    __shiftMinutes: shiftMinutes,
    __totalTimedMinutes: totalMinutes,
    __unassignedMinutes: unassignedMinutes,
    __shiftLabel: label
  };
}

function timedLogShiftWeight(row, selectedShift) {
  if (isFullDayShift(selectedShift)) return 1;
  const shiftKey = selectedShift === "Night" ? "Night" : "Day";
  const total = Math.max(0, num(row?.__totalTimedMinutes));
  const covered = Math.max(0, num(row?.__shiftMinutes?.[shiftKey]));
  if (total > 0) return covered / total;
  if (Array.isArray(row?.__shiftMatches)) return row.__shiftMatches.includes(shiftKey) ? 1 : 0;
  return 0;
}

function resolveTimedLogShiftLabel(row, line, startField, finishField) {
  const explicit = String(row?.__shiftLabel || "").trim();
  if (explicit) return explicit;
  const coverage = shiftCoverageForTimedLog(line, row?.date, row?.[startField], row?.[finishField]);
  if (coverage.shiftMatches.length || coverage.unassignedMinutes > 0) return coverage.label;
  return "Unassigned";
}

function rowMatchesDateShift(row, date, shift, { line = null, startField = "", finishField = "" } = {}) {
  if (!row || row.date !== date) return false;
  if (isFullDayShift(shift)) return true;
  const requestedShift = shift === "Night" ? "Night" : "Day";
  const explicitMatches = Array.isArray(row?.__shiftMatches) ? row.__shiftMatches : null;
  if (explicitMatches) return explicitMatches.includes(requestedShift);
  if (line && startField && finishField) {
    const coverage = shiftCoverageForTimedLog(line, row.date, row[startField], row[finishField]);
    return coverage.shiftMatches.includes(requestedShift);
  }
  return shiftKeysForSelection(shift).includes(row.shift);
}

function fallbackShiftValue(shift) {
  return shift === "Night" ? "Night" : "Day";
}

function isWithinShiftWindow(targetMins, startMins, finishMins) {
  if (!Number.isFinite(targetMins) || !Number.isFinite(startMins) || !Number.isFinite(finishMins)) return false;
  if (startMins === finishMins) return true;
  if (finishMins > startMins) return targetMins >= startMins && targetMins <= finishMins;
  return targetMins >= startMins || targetMins <= finishMins;
}

function inferShiftForLog(line, date, timeValue, fallbackShift = "Day") {
  const fallback = fallbackShiftValue(fallbackShift);
  const shiftRows = (line?.shiftRows || [])
    .filter((row) => row.date === date && SHIFT_OPTIONS.includes(row.shift))
    .slice()
    .sort((a, b) => String(a.startTime || "").localeCompare(String(b.startTime || "")));
  const targetMins = parseTimeToMinutes(timeValue);

  if (Number.isFinite(targetMins)) {
    const matched = shiftRows.find((row) => {
      const startMins = parseTimeToMinutes(row.startTime);
      const finishMins = parseTimeToMinutes(row.finishTime);
      return isWithinShiftWindow(targetMins, startMins, finishMins);
    });
    if (matched?.shift) return matched.shift;
  }

  if (shiftRows.some((row) => row.shift === fallback)) return fallback;
  if (shiftRows[0]?.shift) return shiftRows[0].shift;
  return fallback;
}

function selectedShiftRowsByDate(rows, date, shift, { line = null } = {}) {
  return (rows || []).filter((row) => {
    if (row && Object.prototype.hasOwnProperty.call(row, "productionStartTime") && Object.prototype.hasOwnProperty.call(row, "finishTime")) {
      return rowMatchesDateShift(row, date, shift, { line, startField: "productionStartTime", finishField: "finishTime" });
    }
    if (row && Object.prototype.hasOwnProperty.call(row, "downtimeStart") && Object.prototype.hasOwnProperty.call(row, "downtimeFinish")) {
      return rowMatchesDateShift(row, date, shift, { line, startField: "downtimeStart", finishField: "downtimeFinish" });
    }
    return rowMatchesDateShift(row, date, shift);
  });
}

function rowIsValidDateShift(date, shift) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(date || "")) && SHIFT_OPTIONS.includes(String(shift || ""));
}

async function apiRequest(path, { method = "GET", token = "", body } = {}) {
  dbLoadingRequestCount += 1;
  showStartupLoading("Loading production data...");
  try {
    const headers = {};
    if (token) headers.Authorization = `Bearer ${token}`;
    if (body !== undefined) headers["Content-Type"] = "application/json";
    const controller = typeof AbortController !== "undefined" ? new AbortController() : null;
    const timeoutId =
      controller && typeof window !== "undefined"
        ? window.setTimeout(() => controller.abort(), API_REQUEST_TIMEOUT_MS)
        : null;
    let response;
    let text = "";
    try {
      response = await fetch(`${API_BASE_URL}${path}`, {
        method,
        headers,
        body: body !== undefined ? JSON.stringify(body) : undefined,
        signal: controller ? controller.signal : undefined
      });
      text = await response.text();
    } catch (error) {
      if (error?.name === "AbortError") {
        throw new Error(`Request timed out after ${Math.round(API_REQUEST_TIMEOUT_MS / 1000)}s.`);
      }
      throw error;
    } finally {
      if (timeoutId && typeof window !== "undefined") window.clearTimeout(timeoutId);
    }
    let payload = null;
    try {
      payload = text ? JSON.parse(text) : null;
    } catch {
      payload = text ? { raw: text } : null;
    }
    if (!response.ok) {
      const error = new Error(payload?.error || payload?.message || `API ${response.status}`);
      error.status = response.status;
      error.payload = payload;
      throw error;
    }
    return payload;
  } finally {
    dbLoadingRequestCount = Math.max(0, dbLoadingRequestCount - 1);
    hideStartupLoading();
  }
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
  persistAuthSessions();
}

function clearManagerBackendSession() {
  managerBackendSession.backendToken = "";
  managerBackendSession.backendLineMap = {};
  managerBackendSession.backendStageMap = {};
  managerBackendSession.role = "manager";
  persistAuthSessions();
}

async function ensureManagerBackendSession() {
  managerBackendSession.role = "manager";
  if (!managerBackendSession.backendToken) {
    throw new Error("Manager login required.");
  }
  return managerBackendSession;
}

function fromBackendStageType(stageType) {
  if (stageType === "transfer") return { group: "prep", kind: "transfer" };
  if (stageType === "prep") return { group: "prep", kind: undefined };
  return { group: "main", kind: undefined };
}

function toBackendStageType(stage) {
  if (stage?.kind === "transfer") return "transfer";
  if (stage?.group === "prep") return "prep";
  return "main";
}

function makeLineFromBackend(lineSummary, lineDetail, logs) {
  const lineId = lineSummary.id;
  const line = makeDefaultLine(lineId, lineSummary.name || "Production Line");
  line.secretKey = lineSummary.secretKey || line.secretKey;
  const backendStages = Array.isArray(lineDetail?.stages) ? lineDetail.stages : [];
  line.stages = backendStages.length
    ? backendStages.map((stage, index) => {
        const mappedType = fromBackendStageType(stage.stageType);
        return {
          id: stage.id,
          name: `${index + 1}. ${String(stage.stageName || "Stage").replace(/^\s*\d+\.\s*/, "").trim()}`,
          crew: Math.max(0, num(stage.dayCrew)),
          group: mappedType.group,
          kind: mappedType.kind,
          match: String(stage.stageName || "")
            .toLowerCase()
            .split(/[^a-z0-9]+/)
            .filter(Boolean),
          x: num(stage.x),
          y: num(stage.y),
          w: num(stage.w) || (mappedType.kind === "transfer" ? 7.5 : 9.5),
          h: num(stage.h) || (mappedType.kind === "transfer" ? 10 : 15)
        };
      })
    : clone(STAGES);

  line.crewsByShift = { Day: {}, Night: {} };
  line.stageSettings = {};
  line.stages.forEach((stage) => {
    const bStage = backendStages.find((item) => item.id === stage.id);
    line.crewsByShift.Day[stage.id] = { crew: Math.max(0, num(bStage?.dayCrew ?? defaultCrewForStage(stage))) };
    line.crewsByShift.Night[stage.id] = { crew: Math.max(0, num(bStage?.nightCrew ?? defaultCrewForStage(stage))) };
    line.stageSettings[stage.id] = { maxThroughput: Math.max(0, num(bStage?.maxThroughputPerCrew ?? defaultMaxThroughputForStage(stage))) };
  });

  line.flowGuides = Array.isArray(lineDetail?.guides)
    ? lineDetail.guides.map((guide) => ({
        id: guide.id,
        type: guide.guideType === "arrow" ? "arrow" : guide.guideType === "shape" ? "shape" : "line",
        x: num(guide.x),
        y: num(guide.y),
        w: Math.max(2, num(guide.w)),
        h: Math.max(1, num(guide.h)),
        angle: num(guide.angle),
        src: guide.src || ""
      }))
    : [];

  line.shiftRows = Array.isArray(logs?.shiftRows) ? logs.shiftRows : [];
  line.breakRows = Array.isArray(logs?.breakRows) ? logs.breakRows : [];
  line.runRows = Array.isArray(logs?.runRows)
    ? logs.runRows.map((row) => ({
        ...row,
        runCrewingPattern: parseRunCrewingPattern(row?.runCrewingPattern),
        shift: ""
      }))
    : [];
  line.downtimeRows = Array.isArray(logs?.downtimeRows)
    ? logs.downtimeRows.map((row) => {
        const parsedReason = parseDowntimeReasonParts(row?.reason, row?.equipment);
        return {
          ...row,
          shift: "",
          reasonCategory: row?.reasonCategory || parsedReason.reasonCategory,
          reasonDetail: row?.reasonDetail || parsedReason.reasonDetail,
          reasonNote: row?.reasonNote || parsedReason.reasonNote
        };
      })
    : [];
  line.auditRows = [];
  line.supervisorLogs = [];
  line.selectedStageId = line.stages[0]?.id || "";
  enforcePermanentCrewAndThroughput(line);
  return line;
}

async function refreshHostedState(preferredSession = null) {
  try {
    const activeSession = preferredSession || (await ensureManagerBackendSession());
    if (!activeSession?.backendToken) throw new Error("Missing backend token.");
    const snapshot = await apiRequest("/api/state-snapshot", { token: activeSession.backendToken });
    const snapshotLines = Array.isArray(snapshot?.lines) ? snapshot.lines : [];
    const hostedLines = {};
    activeSession.backendLineMap = {};
    activeSession.backendStageMap = {};
    snapshotLines.forEach((bundle) => {
      const lineSummary = bundle?.line || {};
      const lineDetail = { stages: bundle?.stages || [], guides: bundle?.guides || [] };
      const logs = {
        shiftRows: bundle?.shiftRows || [],
        breakRows: bundle?.breakRows || [],
        runRows: bundle?.runRows || [],
        downtimeRows: bundle?.downtimeRows || []
      };
      if (!lineSummary?.id) return;
      const line = makeLineFromBackend(lineSummary, lineDetail, logs);
      hostedLines[line.id] = line;
      activeSession.backendLineMap[line.id] = line.id;
      (line.stages || []).forEach((stage) => {
        activeSession.backendStageMap[`${line.id}::${stage.id}`] = stage.id;
      });
    });

    appState.lines = hostedLines;
    const ids = Object.keys(hostedLines);
    appState.activeLineId = hostedLines[appState.activeLineId] ? appState.activeLineId : ids[0] || "";
    state = appState.lines[appState.activeLineId] || null;

    const supervisors = Array.isArray(snapshot?.supervisors) ? snapshot.supervisors : [];
    const hasSnapshotSupervisorAssignments = Object.prototype.hasOwnProperty.call(snapshot || {}, "supervisorAssignments");
    const snapshotSupervisorLineShifts = hasSnapshotSupervisorAssignments
      ? normalizeSupervisorLineShifts(snapshot?.supervisorAssignments?.assignedLineShifts, hostedLines, snapshot?.supervisorAssignments?.assignedLineIds || [])
      : {};
    appState.supervisors = supervisors.map((sup) => {
      const assignedLineShifts = normalizeSupervisorLineShifts(sup.assignedLineShifts, hostedLines, sup.assignedLineIds || []);
      return {
        id: sup.id,
        name: sup.name,
        username: sup.username,
        password: "",
        assignedLineShifts,
        assignedLineIds: Object.keys(assignedLineShifts)
      };
    });

    if (appState.supervisorSession) {
      const current = appState.supervisors.find((sup) => sup.username === appState.supervisorSession.username);
      if (current) {
        appState.supervisorSession = {
          ...appState.supervisorSession,
          assignedLineIds: current.assignedLineIds.slice(),
          assignedLineShifts: clone(current.assignedLineShifts || {})
        };
      } else if (hasSnapshotSupervisorAssignments) {
        appState.supervisorSession = {
          ...appState.supervisorSession,
          assignedLineIds: Object.keys(snapshotSupervisorLineShifts),
          assignedLineShifts: clone(snapshotSupervisorLineShifts),
          backendLineMap: { ...activeSession.backendLineMap },
          backendToken: activeSession.backendToken
        };
      } else if (preferredSession?.backendToken || appState.supervisorSession?.backendToken) {
        const retainedLineShifts = normalizeSupervisorLineShifts(
          appState.supervisorSession.assignedLineShifts,
          hostedLines,
          appState.supervisorSession.assignedLineIds || []
        );
        appState.supervisorSession = {
          ...appState.supervisorSession,
          assignedLineIds: Object.keys(retainedLineShifts),
          assignedLineShifts: retainedLineShifts,
          backendLineMap: { ...activeSession.backendLineMap },
          backendToken: activeSession.backendToken
        };
      } else {
        appState.supervisorSession = null;
      }
    }

    hostedDatabaseAvailable = true;
    hostedRefreshErrorShown = false;
    saveState();
    renderAll();
    return true;
  } catch (error) {
    console.warn("Hosted state refresh failed:", error);
    hostedDatabaseAvailable = false;
    const status = Number(error?.status || 0);
    const message = String(error?.message || "").toLowerCase();
    const authFailure =
      status === 401 &&
      (message.includes("missing or invalid authorization header") || message.includes("not authenticated"));
    if (authFailure) {
      if (preferredSession === appState.supervisorSession || appState.appMode === "supervisor") {
        appState.supervisorSession = null;
      } else {
        clearManagerBackendSession();
      }
      saveState();
      renderAll();
    }
    renderAll();
    if (!hostedRefreshErrorShown) {
      hostedRefreshErrorShown = true;
      alert(`Could not refresh data from server.\n${error?.message || "Please try again."}`);
    }
    return false;
  }
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
  const session = await ensureManagerBackendSession();
  const backendLineId = await ensureBackendLineId(payload.lineId, session);
  if (!backendLineId) throw new Error("Line is not synced to server.");
  const response = await apiRequest("/api/logs/shifts", {
    method: "POST",
    token: session.backendToken,
    body: {
      ...payload,
      lineId: backendLineId
    }
  });
  return response?.shiftLog || null;
}

async function syncManagerRunLog(payload) {
  const session = await ensureManagerBackendSession();
  const backendLineId = await ensureBackendLineId(payload.lineId, session);
  if (!backendLineId) throw new Error("Line is not synced to server.");
  const response = await apiRequest("/api/logs/runs", {
    method: "POST",
    token: session.backendToken,
    body: {
      ...payload,
      lineId: backendLineId
    }
  });
  return response?.runLog || null;
}

async function syncManagerDowntimeLog(payload) {
  const session = await ensureManagerBackendSession();
  const backendLineId = await ensureBackendLineId(payload.lineId, session);
  if (!backendLineId) throw new Error("Line is not synced to server.");
  const backendEquipmentId = await ensureBackendStageId(payload.lineId, payload.equipment, session);
  const response = await apiRequest("/api/logs/downtime", {
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
  return response?.downtimeLog || null;
}

async function patchManagerShiftLog(logId, payload) {
  const session = await ensureManagerBackendSession();
  const response = await apiRequest(`/api/logs/shifts/${logId}`, {
    method: "PATCH",
    token: session.backendToken,
    body: payload
  });
  return response?.shiftLog || null;
}

async function patchManagerRunLog(logId, payload) {
  const session = await ensureManagerBackendSession();
  const response = await apiRequest(`/api/logs/runs/${logId}`, {
    method: "PATCH",
    token: session.backendToken,
    body: payload
  });
  return response?.runLog || null;
}

async function patchManagerDowntimeLog(logId, payload) {
  const session = await ensureManagerBackendSession();
  const backendEquipmentId = payload.equipment
    ? await ensureBackendStageId(payload.lineId || state.id, payload.equipment, session)
    : null;
  const response = await apiRequest(`/api/logs/downtime/${logId}`, {
    method: "PATCH",
    token: session.backendToken,
    body: {
      downtimeStart: payload.downtimeStart,
      downtimeFinish: payload.downtimeFinish,
      equipmentStageId: backendEquipmentId || null,
      reason: payload.reason || ""
    }
  });
  return response?.downtimeLog || null;
}

async function saveLineModelToBackend(lineId) {
  const session = await ensureManagerBackendSession();
  const line = appState.lines[lineId];
  if (!line) return;
  const stages = (line.stages || []).map((stage, index) => ({
    stageOrder: index + 1,
    stageName: stageBaseName(stage.name) || "Stage",
    stageType: toBackendStageType(stage),
    dayCrew: Math.max(0, num(line?.crewsByShift?.Day?.[stage.id]?.crew ?? stage.crew)),
    nightCrew: Math.max(0, num(line?.crewsByShift?.Night?.[stage.id]?.crew ?? stage.crew)),
    maxThroughputPerCrew: Math.max(0, num(line?.stageSettings?.[stage.id]?.maxThroughput)),
    x: num(stage.x),
    y: num(stage.y),
    w: Math.max(2, num(stage.w)),
    h: Math.max(1, num(stage.h))
  }));
  const guides = normalizeFlowGuides(line.flowGuides).map((guide) => ({
    guideType: guide.type === "arrow" ? "arrow" : guide.type === "shape" ? "shape" : "line",
    x: num(guide.x),
    y: num(guide.y),
    w: num(guide.w),
    h: num(guide.h),
    angle: num(guide.angle),
    src: guide.type === "shape" ? guide.src || "" : ""
  }));
  await apiRequest(`/api/lines/${lineId}/model`, {
    method: "PUT",
    token: session.backendToken,
    body: { stages, guides }
  });
}

async function createLineOnBackend(lineName, secretKey, lineModel) {
  const session = await ensureManagerBackendSession();
  const created = await apiRequest("/api/lines", {
    method: "POST",
    token: session.backendToken,
    body: {
      name: lineName,
      secretKey
    }
  });
  const lineId = created?.line?.id;
  if (!lineId) throw new Error("Backend line create failed");
  appState.lines[lineId] = lineModel;
  lineModel.id = lineId;
  await saveLineModelToBackend(lineId);
  return { lineId, created };
}

async function syncSupervisorShiftLog(session, payload) {
  const backendLineId = await ensureBackendLineId(payload.lineId, session);
  if (!backendLineId) throw new Error("Line is not synced to server.");
  const response = await apiRequest("/api/logs/shifts", {
    method: "POST",
    token: session.backendToken,
    body: {
      ...payload,
      lineId: backendLineId
    }
  });
  return response?.shiftLog || null;
}

async function patchSupervisorShiftLog(session, logId, payload) {
  const response = await apiRequest(`/api/logs/shifts/${logId}`, {
    method: "PATCH",
    token: session.backendToken,
    body: payload
  });
  return response?.shiftLog || null;
}

async function startSupervisorShiftBreak(session, shiftLogId, breakStart) {
  const response = await apiRequest(`/api/logs/shifts/${shiftLogId}/breaks`, {
    method: "POST",
    token: session.backendToken,
    body: { breakStart }
  });
  return response?.breakLog || null;
}

async function patchSupervisorShiftBreak(session, shiftLogId, breakId, payload) {
  const response = await apiRequest(`/api/logs/shifts/${shiftLogId}/breaks/${breakId}`, {
    method: "PATCH",
    token: session.backendToken,
    body: payload
  });
  return response?.breakLog || null;
}

async function endSupervisorShiftBreak(session, shiftLogId, breakId, breakFinish) {
  return patchSupervisorShiftBreak(session, shiftLogId, breakId, { breakFinish });
}

async function syncSupervisorRunLog(session, payload) {
  const backendLineId = await ensureBackendLineId(payload.lineId, session);
  if (!backendLineId) throw new Error("Line is not synced to server.");
  const response = await apiRequest("/api/logs/runs", {
    method: "POST",
    token: session.backendToken,
    body: {
      ...payload,
      lineId: backendLineId
    }
  });
  return response?.runLog || null;
}

async function patchSupervisorRunLog(session, logId, payload) {
  const response = await apiRequest(`/api/logs/runs/${logId}`, {
    method: "PATCH",
    token: session.backendToken,
    body: payload
  });
  return response?.runLog || null;
}

async function syncSupervisorDowntimeLog(session, payload) {
  const backendLineId = await ensureBackendLineId(payload.lineId, session);
  if (!backendLineId) throw new Error("Line is not synced to server.");
  const backendEquipmentId = payload.equipment
    ? await ensureBackendStageId(payload.lineId, payload.equipment, session)
    : null;
  const response = await apiRequest("/api/logs/downtime", {
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
  return response?.downtimeLog || null;
}

async function patchSupervisorDowntimeLog(session, logId, payload) {
  const backendEquipmentId = payload.equipment
    ? await ensureBackendStageId(payload.lineId, payload.equipment, session)
    : null;
  const response = await apiRequest(`/api/logs/downtime/${logId}`, {
    method: "PATCH",
    token: session.backendToken,
    body: {
      downtimeStart: payload.downtimeStart,
      downtimeFinish: payload.downtimeFinish,
      equipmentStageId: backendEquipmentId || null,
      reason: payload.reason || ""
    }
  });
  return response?.downtimeLog || null;
}

function requiredCrewForLineShift(line, shift) {
  const stages = line?.stages?.length ? line.stages : STAGES;
  return stages.reduce((sum, stage) => sum + Math.max(0, num(line?.crewsByShift?.[shift]?.[stage.id]?.crew)), 0);
}

function nowIso() {
  return new Date().toISOString();
}

function makeLocalLogId(prefix = "log") {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return `${prefix}-${crypto.randomUUID()}`;
  }
  return `${prefix}-${Date.now()}-${Math.floor(Math.random() * 1000000)}`;
}

function ensureManagerLogRowIds(line) {
  if (!line) return;
  const ensureRows = (rows, prefix) => {
    if (!Array.isArray(rows)) return [];
    rows.forEach((row) => {
      if (row && typeof row === "object" && !String(row.id || "").trim()) {
        row.id = makeLocalLogId(prefix);
      }
    });
    return rows;
  };
  line.shiftRows = ensureRows(line.shiftRows, "shift");
  line.runRows = ensureRows(line.runRows, "run");
  line.downtimeRows = ensureRows(line.downtimeRows, "down");
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
  if (isFullDayShift(shift)) {
    const dayCrew = num(state.crewsByShift?.Day?.[stageId]?.crew);
    const nightCrew = num(state.crewsByShift?.Night?.[stageId]?.crew);
    const stage = getStages().find((s) => s.id === stageId);
    const crew = Math.max(dayCrew, nightCrew);
    if (crew > 0) return crew;
    if (stage?.kind === "transfer") return 1;
    return 0;
  }
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

function downtimeDetailOptions(line, category) {
  if (category === "Equipment") {
    const stages = line?.stages?.length ? line.stages : getStages();
    return stages.map((stage, index) => ({ value: stage.id, label: stageDisplayName(stage, index) }));
  }
  const list = DOWNTIME_REASON_PRESETS[category] || [];
  return list.map((value) => ({ value, label: value }));
}

function setDowntimeDetailOptions(selectNode, line, category, selectedValue = "") {
  if (!selectNode) return;
  const options = downtimeDetailOptions(line, category);
  const placeholder = category === "Equipment" ? "Select Stage" : "Select Reason";
  selectNode.innerHTML = [
    `<option value="">${placeholder}</option>`,
    ...options.map((option) => `<option value="${option.value}">${option.label}</option>`)
  ].join("");
  if (selectedValue && options.some((option) => option.value === selectedValue)) {
    selectNode.value = selectedValue;
  }
}

function downtimeDetailLabel(line, category, detail) {
  if (category === "Equipment") return stageNameByIdForLine(line || state, detail) || "Equipment";
  return String(detail || "").trim();
}

function buildDowntimeReasonText(line, category, detail, note) {
  const group = String(category || "").trim();
  const detailLabel = downtimeDetailLabel(line, group, detail);
  const detailText = String(detailLabel || "").trim();
  const noteText = String(note || "").trim();
  if (!group) return noteText;
  if (!detailText && !noteText) return group;
  if (!noteText) return `${group} > ${detailText}`;
  if (!detailText) return `${group} > ${noteText}`;
  return `${group} > ${detailText} > ${noteText}`;
}

function parseDowntimeReasonParts(reasonText, equipment = "") {
  const raw = String(reasonText || "").trim();
  const parts = raw
    .split(">")
    .map((part) => part.trim())
    .filter(Boolean);
  const category = parts[0] || "";
  const detailFromReason = parts[1] || "";
  const note = parts.slice(2).join(" > ");
  const detail = category === "Equipment" ? (equipment || detailFromReason) : detailFromReason;
  return {
    reasonCategory: category,
    reasonDetail: detail,
    reasonNote: note
  };
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

function capSampleRunRatesForUtilisation(sample, line, maxRatio = 0.95) {
  if (!sample || !line) return;
  const stages = line?.stages?.length ? line.stages : STAGES;
  const minStageCapacityByShift = {
    Day: 0,
    Night: 0
  };

  ["Day", "Night"].forEach((shift) => {
    const capacities = stages
      .map((stage) => stageTotalMaxThroughputForLine(line, stage.id, shift))
      .filter((value) => value > 0);
    minStageCapacityByShift[shift] = capacities.length ? Math.min(...capacities) : 0;
  });

  const downtimeByShift = new Map();
  (sample.downtimeRows || []).forEach((row) => {
    const key = shiftKey(row.date, row.shift);
    const mins = Math.max(0, diffMinutes(row.downtimeStart, row.downtimeFinish));
    downtimeByShift.set(key, (downtimeByShift.get(key) || 0) + mins);
  });

  const runGroups = new Map();
  (sample.runRows || []).forEach((row) => {
    const key = shiftKey(row.date, row.shift);
    if (!runGroups.has(key)) runGroups.set(key, []);
    runGroups.get(key).push(row);
  });

  runGroups.forEach((rows, key) => {
    if (!rows.length) return;
    const shift = rows[0]?.shift === "Night" ? "Night" : "Day";
    const capRate = Math.max(0, minStageCapacityByShift[shift] * Math.max(0, num(maxRatio)));
    if (capRate <= 0) return;

    const downtimeMins = Math.max(0, num(downtimeByShift.get(key)));
    let unitsTotal = 0;
    let netTimeTotal = 0;
    rows.forEach((row) => {
      const gross = Math.max(0, diffMinutes(row.productionStartTime, row.finishTime));
      const net = Math.max(0, gross - downtimeMins);
      unitsTotal += Math.max(0, num(row.unitsProduced));
      netTimeTotal += net;
    });
    if (unitsTotal <= 0 || netTimeTotal <= 0) return;

    const currentRate = unitsTotal / netTimeTotal;
    if (currentRate <= capRate) return;

    const scale = capRate / currentRate;
    rows.forEach((row) => {
      row.unitsProduced = Math.max(1, Math.round(Math.max(0, num(row.unitsProduced)) * scale));
    });
  });
}

function sampleDataSet(lineModel = state) {
  const line = lineModel || state || {};
  const shiftRows = [];
  const breakRows = [];
  const runRows = [];
  const downtimeRows = [];
  const start = new Date(2025, 10, 1);
  const days = 100;
  const stageIds = (line?.stages?.length ? line.stages : STAGES).map((stage) => stage.id).filter(Boolean);
  const dayEquipment = stageIds.filter((_, idx) => idx % 2 === 0);
  const nightEquipment = stageIds.filter((_, idx) => idx % 2 === 1);
  const fallbackEquipment = stageIds[0] || "";
  const equipAt = (list, i, offset = 0) => (list.length ? list[(i + offset) % list.length] : fallbackEquipment);
  const dayRequired = requiredCrewForLineShift(line, "Day");
  const nightRequired = requiredCrewForLineShift(line, "Night");

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
        finishTime: "14:00"
      },
      {
        date,
        shift: "Night",
        crewOnShift: Math.max(0, nightRequired - (i % 10 === 0 ? 1 : 0)),
        startTime: "14:00",
        finishTime: "22:00"
      }
    );

    breakRows.push(
      { date, shift: "Day", breakStart: "09:00", breakFinish: "09:15" },
      { date, shift: "Day", breakStart: "12:00", breakFinish: "12:30" },
      { date, shift: "Day", breakStart: "13:40", breakFinish: "13:55" },
      { date, shift: "Night", breakStart: "17:00", breakFinish: "17:15" },
      { date, shift: "Night", breakStart: "20:00", breakFinish: "20:30" },
      { date, shift: "Night", breakStart: "21:40", breakFinish: "21:55" }
    );

    runRows.push(
      {
        date,
        shift: "Day",
        product: "Teriyaki",
        productionStartTime: "06:10",
        finishTime: "10:35",
        unitsProduced: Math.round(2850 * dayTrend)
      },
      {
        date,
        shift: "Day",
        product: "Honey Soy",
        productionStartTime: "10:55",
        finishTime: "15:25",
        unitsProduced: Math.round(2600 * dayTrend)
      },
      {
        date,
        shift: "Night",
        product: "Peri Peri",
        productionStartTime: "14:15",
        finishTime: "18:55",
        unitsProduced: Math.round(2500 * nightTrend)
      },
      {
        date,
        shift: "Night",
        product: "Lemon Herb",
        productionStartTime: "21:30",
        finishTime: "00:25",
        unitsProduced: Math.round(2200 * nightTrend)
      }
    );

    const deq = equipAt(dayEquipment, i);
    const neq = equipAt(nightEquipment, i);
    const dayReasonCategory = i % 4 === 0 ? "People" : "Equipment";
    const dayReasonDetail = dayReasonCategory === "Equipment" ? deq : DOWNTIME_REASON_PRESETS.People[i % DOWNTIME_REASON_PRESETS.People.length];
    const dayReason = buildDowntimeReasonText(line, dayReasonCategory, dayReasonDetail, dayReasonCategory === "Equipment" ? "Planned maintenance" : "");
    const dayReasonCategory2 = i % 5 === 0 ? "Materials" : "Equipment";
    const dayReasonDetail2 =
      dayReasonCategory2 === "Equipment" ? equipAt(dayEquipment, i, 2) : DOWNTIME_REASON_PRESETS.Materials[i % DOWNTIME_REASON_PRESETS.Materials.length];
    const dayReason2 = buildDowntimeReasonText(line, dayReasonCategory2, dayReasonDetail2, dayReasonCategory2 === "Equipment" ? "Minor stoppage" : "");
    const nightReasonCategory = i % 3 === 0 ? "Donor Meat" : "Equipment";
    const nightReasonDetail =
      nightReasonCategory === "Equipment" ? neq : DOWNTIME_REASON_PRESETS["Donor Meat"][i % DOWNTIME_REASON_PRESETS["Donor Meat"].length];
    const nightReason = buildDowntimeReasonText(line, nightReasonCategory, nightReasonDetail, nightReasonCategory === "Equipment" ? "Sensor reset" : "");
    const nightReasonCategory2 = i % 6 === 0 ? "Other" : "Equipment";
    const nightReasonDetail2 =
      nightReasonCategory2 === "Equipment" ? equipAt(nightEquipment, i, 3) : DOWNTIME_REASON_PRESETS.Other[i % DOWNTIME_REASON_PRESETS.Other.length];
    const nightReason2 = buildDowntimeReasonText(line, nightReasonCategory2, nightReasonDetail2, nightReasonCategory2 === "Equipment" ? "Label adjustment" : "");
    downtimeRows.push(
      {
        date,
        shift: "Day",
        downtimeStart: "08:10",
        downtimeFinish: `08:${String(22 + (i % 8)).padStart(2, "0")}`,
        equipment: dayReasonCategory === "Equipment" ? deq : "",
        reasonCategory: dayReasonCategory,
        reasonDetail: dayReasonDetail,
        reasonNote: dayReasonCategory === "Equipment" ? "Planned maintenance" : "",
        reason: dayReason
      },
      {
        date,
        shift: "Day",
        downtimeStart: "11:20",
        downtimeFinish: `11:${String(30 + (i % 10)).padStart(2, "0")}`,
        equipment: dayReasonCategory2 === "Equipment" ? equipAt(dayEquipment, i, 2) : "",
        reasonCategory: dayReasonCategory2,
        reasonDetail: dayReasonDetail2,
        reasonNote: dayReasonCategory2 === "Equipment" ? "Minor stoppage" : "",
        reason: dayReason2
      },
      {
        date,
        shift: "Night",
        downtimeStart: "16:30",
        downtimeFinish: `16:${String(42 + (i % 9)).padStart(2, "0")}`,
        equipment: nightReasonCategory === "Equipment" ? neq : "",
        reasonCategory: nightReasonCategory,
        reasonDetail: nightReasonDetail,
        reasonNote: nightReasonCategory === "Equipment" ? "Sensor reset" : "",
        reason: nightReason
      },
      {
        date,
        shift: "Night",
        downtimeStart: "20:05",
        downtimeFinish: `20:${String(18 + (i % 11)).padStart(2, "0")}`,
        equipment: nightReasonCategory2 === "Equipment" ? equipAt(nightEquipment, i, 3) : "",
        reasonCategory: nightReasonCategory2,
        reasonDetail: nightReasonDetail2,
        reasonNote: nightReasonCategory2 === "Equipment" ? "Label adjustment" : "",
        reason: nightReason2
      }
    );
  }

  const sample = {
    selectedDate: "2026-02-08",
    selectedShift: "Day",
    trendGranularity: "daily",
    trendMonth: "2026-02",
    shiftRows,
    breakRows,
    runRows,
    downtimeRows
  };
  capSampleRunRatesForUtilisation(sample, line, 0.95);
  return sample;
}

function loadSampleDataIntoLine(line) {
  if (!line) return false;
  const sample = sampleDataSet(line);
  line.selectedDate = sample.selectedDate;
  line.selectedShift = sample.selectedShift;
  line.trendGranularity = sample.trendGranularity;
  line.trendMonth = sample.trendMonth;
  line.shiftRows = sample.shiftRows;
  line.breakRows = sample.breakRows;
  line.runRows = sample.runRows;
  line.downtimeRows = sample.downtimeRows;
  ensureManagerLogRowIds(line);
  return true;
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

function computeBreakRow(row) {
  const fallback = num(row.breakMins);
  const calc = diffMinutes(row.breakStart, row.breakFinish);
  return { ...row, breakMins: calc > 0 ? calc : fallback };
}

function computeRunRow(row, downtimeByShift) {
  const grossFallback = num(row.grossProductionTime);
  const grossCalc = diffMinutes(row.productionStartTime, row.finishTime);
  const grossProductionTime = grossCalc > 0 ? grossCalc : grossFallback;

  const associatedFallback = num(row.associatedDownTime);
  const perShiftMinutes = row?.__shiftMinutes;
  const coveredShiftMinutes = LOG_ASSIGNABLE_SHIFTS.reduce((sum, shiftKey) => sum + Math.max(0, num(perShiftMinutes?.[shiftKey])), 0);
  let associatedFromDowntime = 0;
  if (coveredShiftMinutes > 0) {
    LOG_ASSIGNABLE_SHIFTS.forEach((shiftKey) => {
      const mins = Math.max(0, num(perShiftMinutes?.[shiftKey]));
      if (mins <= 0) return;
      const downForShift = Math.max(0, num(downtimeByShift.get(`${row.date}__${shiftKey}`) || 0));
      associatedFromDowntime += (mins / coveredShiftMinutes) * downForShift;
    });
  }
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
  const downtimeRows = state.downtimeRows
    .map(computeDowntimeRow)
    .map((row) => decorateTimedLogShift(row, state, "downtimeStart", "downtimeFinish"));
  const shiftRows = state.shiftRows.map(computeShiftRow);
  const breakRows = (state.breakRows || []).map(computeBreakRow);
  const downtimeByShift = new Map();
  downtimeRows.forEach((row) => {
    LOG_ASSIGNABLE_SHIFTS.forEach((shiftKey) => {
      const mins = Math.max(0, num(row?.__shiftMinutes?.[shiftKey]));
      if (mins <= 0) return;
      const key = `${row.date}__${shiftKey}`;
      downtimeByShift.set(key, (downtimeByShift.get(key) || 0) + mins);
    });
  });
  const runRows = state.runRows
    .map((row) => decorateTimedLogShift(row, state, "productionStartTime", "finishTime"))
    .map((row) => computeRunRow(row, downtimeByShift));
  return { shiftRows, breakRows, runRows, downtimeRows };
}

function clearLineTrackingData(line) {
  line.shiftRows = [];
  line.breakRows = [];
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
    const tabId = btn.dataset.tab;
    if (!tabId) return;
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach((b) => {
        b.classList.remove("active");
        b.setAttribute("aria-selected", "false");
      });
      document.querySelectorAll(".tab-panel").forEach((panel) => panel.classList.remove("active"));
      btn.classList.add("active");
      btn.setAttribute("aria-selected", "true");
      const panel = document.getElementById(tabId);
      if (panel) panel.classList.add("active");
    });
  });
}

function bindHome() {
  const goHomeBtn = document.getElementById("goHome");
  const modeManagerBtn = document.getElementById("modeManager");
  const modeSupervisorBtn = document.getElementById("modeSupervisor");
  const lineModeManagerBtn = document.getElementById("lineModeManager");
  const lineModeSupervisorBtn = document.getElementById("lineModeSupervisor");
  const sidebarToggleBtn = document.getElementById("sidebarToggle");
  const sidebarBackdrop = document.getElementById("sidebarBackdrop");
  const homeSidebar = document.getElementById("homeSidebar");
  const dashboardDate = document.getElementById("dashboardDate");
  const dashboardShiftButtons = Array.from(document.querySelectorAll("[data-dash-shift]"));
  const exportDashboardCsvBtn = document.getElementById("exportDashboardCsv");
  const managerLoginForm = document.getElementById("managerLoginForm");
  const managerUserInput = document.getElementById("managerUser");
  const managerPassInput = document.getElementById("managerPass");
  const managerTestLoginBtn = document.getElementById("managerTestLoginBtn");
  const supervisorLoginForm = document.getElementById("supervisorLoginForm");
  const supervisorUserInput = document.getElementById("supervisorUser");
  const supervisorPassInput = document.getElementById("supervisorPass");
  const supervisorTestLoginBtn = document.getElementById("supervisorTestLoginBtn");
  const supervisorLogoutBtn = document.getElementById("supervisorLogout");
  const supervisorMobileModeBtn = document.getElementById("supervisorMobileMode");
  const supervisorLineSelect = document.getElementById("supervisorLineSelect");
  const svDateInputs = Array.from(document.querySelectorAll("[data-sv-date]"));
  const svShiftButtons = Array.from(document.querySelectorAll("[data-sv-shift]"));
  const svPrevBtns = Array.from(document.querySelectorAll("[data-sv-prev]"));
  const svNextBtns = Array.from(document.querySelectorAll("[data-sv-next]"));
  const supervisorShiftForm = document.getElementById("supervisorShiftForm");
  const supervisorRunForm = document.getElementById("supervisorRunForm");
  const supervisorDownForm = document.getElementById("supervisorDownForm");
  const superShiftLogIdInput = document.getElementById("superShiftLogId");
  const superShiftOpenBreakIdInput = document.getElementById("superShiftOpenBreakId");
  const superShiftBreakTimeInput = document.getElementById("superShiftBreakTime");
  const superRunLogIdInput = document.getElementById("superRunLogId");
  const superRunCrewingPatternInput = document.getElementById("superRunCrewingPattern");
  const superRunCrewingPatternBtn = document.getElementById("superRunCrewingPatternBtn");
  const superRunCrewingPatternSummary = document.getElementById("superRunCrewingPatternSummary");
  const superDownLogIdInput = document.getElementById("superDownLogId");
  const superShiftSaveProgressBtn = document.getElementById("superShiftSaveProgress");
  const superShiftBreakStartBtn = document.getElementById("superShiftBreakStart");
  const superShiftBreakEndBtn = document.getElementById("superShiftBreakEnd");
  const superShiftOpenList = document.getElementById("superShiftOpenList");
  const superRunSaveProgressBtn = document.getElementById("superRunSaveProgress");
  const superRunOpenList = document.getElementById("superRunOpenList");
  const superDownSaveProgressBtn = document.getElementById("superDownSaveProgress");
  const superDownOpenList = document.getElementById("superDownOpenList");
  const supervisorDownReasonCategory = document.getElementById("superDownReasonCategory");
  const supervisorDownReasonDetail = document.getElementById("superDownReasonDetail");
  const manageSupervisorsBtn = document.getElementById("manageSupervisorsBtn");
  const addSupervisorBtn = document.getElementById("addSupervisorBtn");
  const manageSupervisorsModal = document.getElementById("manageSupervisorsModal");
  const closeManageSupervisorsModalBtn = document.getElementById("closeManageSupervisorsModal");
  const supervisorManagerList = document.getElementById("supervisorManagerList");
  const addSupervisorModal = document.getElementById("addSupervisorModal");
  const closeAddSupervisorModalBtn = document.getElementById("closeAddSupervisorModal");
  const addSupervisorForm = document.getElementById("addSupervisorForm");
  const newSupervisorLines = document.getElementById("newSupervisorLines");
  const editSupervisorModal = document.getElementById("editSupervisorModal");
  const closeEditSupervisorModalBtn = document.getElementById("closeEditSupervisorModal");
  const editSupervisorForm = document.getElementById("editSupervisorForm");
  const editSupervisorNameInput = document.getElementById("editSupervisorName");
  const editSupervisorUsernameInput = document.getElementById("editSupervisorUsername");
  const editSupervisorPasswordInput = document.getElementById("editSupervisorPassword");
  const editLineModal = document.getElementById("editLineModal");
  const closeEditLineModalBtn = document.getElementById("closeEditLineModal");
  const editLineForm = document.getElementById("editLineForm");
  const editLineNameInput = document.getElementById("editLineName");
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
  let editingSupervisorId = "";
  let editingLineId = "";
  let suppressSupervisorSelectionReset = false;

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

  const renderSupervisorLineChecklist = (selectedLineShifts = {}) => {
    const lineIds = Object.keys(appState.lines);
    if (!newSupervisorLines) return;
    if (!lineIds.length) {
      newSupervisorLines.innerHTML = `<p class="muted">No production lines available.</p>`;
      return;
    }
    const normalizedSelection = normalizeSupervisorLineShifts(selectedLineShifts, appState.lines, Array.isArray(selectedLineShifts) ? selectedLineShifts : []);
    newSupervisorLines.innerHTML = lineIds
      .map((id, index) => {
        const line = appState.lines[id];
        const lineLabel = htmlEscape(line.name);
        const allowed = normalizeSupervisorShifts(normalizedSelection[id], { fallbackToAll: false });
        const dayId = `newSupervisor-${index}-day`;
        const nightId = `newSupervisor-${index}-night`;
        return `
          <div class="supervisor-access-row">
            <span class="supervisor-access-line">${lineLabel}</span>
            <div class="supervisor-access-shifts">
              <label class="supervisor-shift-pill${allowed.includes("Day") ? " is-selected" : ""}" for="${dayId}">
                <input
                  id="${dayId}"
                  type="checkbox"
                  data-new-supervisor-line-shift
                  data-line-id="${id}"
                  value="Day"
                  ${allowed.includes("Day") ? "checked" : ""}
                />
                <span>Day</span>
              </label>
              <label class="supervisor-shift-pill${allowed.includes("Night") ? " is-selected" : ""}" for="${nightId}">
                <input
                  id="${nightId}"
                  type="checkbox"
                  data-new-supervisor-line-shift
                  data-line-id="${id}"
                  value="Night"
                  ${allowed.includes("Night") ? "checked" : ""}
                />
                <span>Night</span>
              </label>
            </div>
          </div>
        `;
      })
      .join("");
    syncSupervisorShiftPillStyles(newSupervisorLines);
  };

  const syncSupervisorShiftPillStyles = (root) => {
    if (!root || typeof root.querySelectorAll !== "function") return;
    root.querySelectorAll(".supervisor-shift-pill").forEach((pill) => {
      const checkbox = pill.querySelector('input[type="checkbox"]');
      pill.classList.toggle("is-selected", Boolean(checkbox?.checked));
    });
  };

  const refreshSupervisorDowntimeDetailOptions = (line = selectedSupervisorLine()) => {
    const category = String(supervisorDownReasonCategory?.value || "");
    if (!supervisorDownReasonDetail) return;
    setDowntimeDetailOptions(supervisorDownReasonDetail, line, category, supervisorDownReasonDetail.value || "");
  };

  const renderSupervisorManagerList = () => {
    const lineIds = Object.keys(appState.lines);
    const linesById = appState.lines;
    supervisorManagerList.innerHTML = (appState.supervisors || [])
      .map((sup) => {
        const accessMap = normalizeSupervisorLineShifts(sup.assignedLineShifts, appState.lines, sup.assignedLineIds || []);
        const assignedLineCount = Object.keys(accessMap).length;
        const assignmentSummary = `${assignedLineCount} line${assignedLineCount === 1 ? "" : "s"} assigned`;
        const accessRows = lineIds
          .map(
            (lineId, index) => {
              const allowed = normalizeSupervisorShifts(accessMap[lineId], { fallbackToAll: false });
              const lineLabel = htmlEscape(linesById[lineId].name);
              return `
                <div class="supervisor-access-row">
                  <span class="supervisor-access-line">${lineLabel}</span>
                  <div class="supervisor-access-shifts" role="group" aria-label="${lineLabel} shift access">
                    <label class="supervisor-shift-pill${allowed.includes("Day") ? " is-selected" : ""}" for="sup-${sup.id}-${index}-day">
                      <input
                        id="sup-${sup.id}-${index}-day"
                        type="checkbox"
                        data-supervisor-line-shift="${sup.id}"
                        data-line-id="${lineId}"
                        value="Day"
                        ${allowed.includes("Day") ? "checked" : ""}
                      />
                      <span>Day</span>
                    </label>
                    <label class="supervisor-shift-pill${allowed.includes("Night") ? " is-selected" : ""}" for="sup-${sup.id}-${index}-night">
                      <input
                        id="sup-${sup.id}-${index}-night"
                        type="checkbox"
                        data-supervisor-line-shift="${sup.id}"
                        data-line-id="${lineId}"
                        value="Night"
                        ${allowed.includes("Night") ? "checked" : ""}
                      />
                      <span>Night</span>
                    </label>
                  </div>
                </div>
              `;
            }
          )
          .join("");
        return `
          <section class="panel supervisor-manager-row" data-supervisor-id="${sup.id}">
            <div class="supervisor-manager-head">
              <div class="supervisor-manager-identity">
                <h3>${htmlEscape(sup.name)}</h3>
                <span class="supervisor-manager-handle">@${htmlEscape(sup.username)}</span>
              </div>
              <span class="supervisor-count-chip">${assignmentSummary}</span>
            </div>
            <div class="supervisor-assignment-wrap">
              <section class="supervisor-assignment-group">
                <h4 class="supervisor-assignment-title">Line and Shift Access</h4>
                <div class="supervisor-access-grid">
                  ${accessRows || `<p class="muted">No production lines available.</p>`}
                </div>
              </section>
            </div>
            <div class="action-row supervisor-manager-actions">
              <button type="button" class="supervisor-action-save" data-supervisor-save="${sup.id}">Save Assignments</button>
              <button type="button" class="ghost-btn supervisor-action-edit" data-supervisor-edit="${sup.id}">Edit Details</button>
              <button type="button" class="danger supervisor-action-delete" data-supervisor-delete="${sup.id}">Delete Supervisor</button>
            </div>
          </section>
        `;
      })
      .join("");
    if (!appState.supervisors?.length) {
      supervisorManagerList.innerHTML = `
        <div class="supervisor-manager-empty">
          <p class="muted">No supervisors created yet.</p>
        </div>
      `;
    }
    syncSupervisorShiftPillStyles(supervisorManagerList);
  };

  const openManageSupervisorsModal = async () => {
    await refreshHostedState();
    renderSupervisorManagerList();
    manageSupervisorsModal.classList.add("open");
    manageSupervisorsModal.setAttribute("aria-hidden", "false");
  };

  const closeManageSupervisorsModal = () => {
    manageSupervisorsModal.classList.remove("open");
    manageSupervisorsModal.setAttribute("aria-hidden", "true");
  };

  const openAddSupervisorModal = async () => {
    await refreshHostedState();
    addSupervisorForm.reset();
    renderSupervisorLineChecklist(Object.fromEntries(Object.keys(appState.lines).map((lineId) => [lineId, SUPERVISOR_SHIFT_OPTIONS.slice()])));
    addSupervisorModal.classList.add("open");
    addSupervisorModal.setAttribute("aria-hidden", "false");
  };

  const closeAddSupervisorModal = () => {
    addSupervisorModal.classList.remove("open");
    addSupervisorModal.setAttribute("aria-hidden", "true");
  };

  const openEditSupervisorModal = (supervisorId) => {
    const sup = (appState.supervisors || []).find((item) => item.id === supervisorId);
    if (!sup) return;
    editingSupervisorId = supervisorId;
    editSupervisorNameInput.value = sup.name || "";
    editSupervisorUsernameInput.value = sup.username || "";
    editSupervisorPasswordInput.value = "";
    editSupervisorModal.classList.add("open");
    editSupervisorModal.setAttribute("aria-hidden", "false");
    editSupervisorNameInput.focus();
    editSupervisorNameInput.select();
  };

  const closeEditSupervisorModal = () => {
    editingSupervisorId = "";
    editSupervisorModal.classList.remove("open");
    editSupervisorModal.setAttribute("aria-hidden", "true");
    editSupervisorForm.reset();
  };

  const openEditLineModal = (lineId) => {
    const line = appState.lines?.[lineId];
    if (!line) return;
    editingLineId = lineId;
    editLineNameInput.value = line.name || "";
    editLineModal.classList.add("open");
    editLineModal.setAttribute("aria-hidden", "false");
    editLineNameInput.focus();
    editLineNameInput.select();
  };

  const closeEditLineModal = () => {
    editingLineId = "";
    editLineModal.classList.remove("open");
    editLineModal.setAttribute("aria-hidden", "true");
    editLineForm.reset();
  };

  goHomeBtn.addEventListener("click", () => {
    appState.activeView = "home";
    saveState();
    renderAll();
  });

  const switchAppMode = (mode, { goHome = false } = {}) => {
    appState.appMode = mode;
    if (goHome) appState.activeView = "home";
    saveState();
    renderAll();
  };

  modeManagerBtn.addEventListener("click", () => {
    switchAppMode("manager");
  });

  modeSupervisorBtn.addEventListener("click", () => {
    switchAppMode("supervisor");
  });

  if (lineModeManagerBtn) {
    lineModeManagerBtn.addEventListener("click", () => {
      switchAppMode("manager");
    });
  }
  if (lineModeSupervisorBtn) {
    lineModeSupervisorBtn.addEventListener("click", () => {
      switchAppMode("supervisor", { goHome: true });
    });
  }

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
      appState.dashboardShift = SHIFT_OPTIONS.includes(shift) ? shift : "Day";
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

  if (managerTestLoginBtn && managerUserInput && managerPassInput) {
    managerTestLoginBtn.addEventListener("click", () => {
      managerUserInput.value = "manager";
      managerPassInput.value = "manager123";
      if (typeof managerLoginForm.requestSubmit === "function") {
        managerLoginForm.requestSubmit();
      } else {
        managerLoginForm.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
      }
    });
  }

  managerLoginForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const username = String(managerUserInput.value || "").trim().toLowerCase();
    const password = String(managerPassInput.value || "");
    try {
      const loginPayload = await apiRequest("/api/auth/login", {
        method: "POST",
        body: { username, password }
      });
      if (loginPayload?.user?.role !== "manager" || !loginPayload?.token) {
        alert("Invalid manager credentials.");
        return;
      }
      managerBackendSession.backendToken = loginPayload.token;
      managerBackendSession.backendLineMap = {};
      managerBackendSession.backendStageMap = {};
      managerBackendSession.role = "manager";
      persistManagerBackendSession();
      appState.appMode = "manager";
      appState.activeView = "home";
      saveState();
      await refreshHostedState(managerBackendSession);
      renderAll();
    } catch (error) {
      const message = String(error?.message || "");
      if (message.toLowerCase().includes("invalid")) {
        alert("Invalid manager credentials.");
      } else {
        alert(`Could not connect to login service.\n${message || "Please try again."}`);
      }
    }
  });

  if (supervisorTestLoginBtn && supervisorUserInput && supervisorPassInput) {
    supervisorTestLoginBtn.addEventListener("click", () => {
      supervisorUserInput.value = "supervisor";
      supervisorPassInput.value = "supervisor123";
      if (typeof supervisorLoginForm.requestSubmit === "function") {
        supervisorLoginForm.requestSubmit();
      } else {
        supervisorLoginForm.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
      }
    });
  }

  supervisorLoginForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const username = String(supervisorUserInput.value || "").trim().toLowerCase();
    const password = String(supervisorPassInput.value || "");
    try {
      const loginPayload = await apiRequest("/api/auth/login", {
        method: "POST",
        body: { username, password }
      });
      if (loginPayload?.user?.role !== "supervisor" || !loginPayload?.token) {
        alert("Invalid supervisor credentials.");
        return;
      }
      appState.supervisorSession = {
        username: loginPayload.user.username,
        assignedLineIds: [],
        assignedLineShifts: {},
        backendToken: loginPayload.token,
        backendLineMap: {},
        backendStageMap: {},
        role: "supervisor"
      };
      saveState();
      await refreshHostedState(appState.supervisorSession);
      renderAll();
    } catch (error) {
      const message = String(error?.message || "");
      if (message.toLowerCase().includes("invalid")) {
        alert("Invalid supervisor credentials.");
      } else {
        alert(`Could not connect to login service.\n${message || "Please try again."}`);
      }
    }
  });

  supervisorLogoutBtn.addEventListener("click", () => {
    if (appState.appMode === "supervisor") {
      appState.supervisorSession = null;
    } else {
      clearManagerBackendSession();
      appState.activeView = "home";
    }
    saveState();
    renderAll();
  });

  supervisorMobileModeBtn.addEventListener("click", () => {
    appState.supervisorMobileMode = !appState.supervisorMobileMode;
    saveState();
    renderHome();
  });

  document.querySelectorAll("[data-now-target]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const target = btn.getAttribute("data-now-target");
      if (!target) return;
      const input = document.getElementById(target);
      if (!input) return;
      input.value = nowTimeHHMM();
      syncNowInputPrefixState(input);
    });
  });

  document.querySelectorAll(".input-now-wrap[data-time-prefix] input").forEach((input) => {
    const sync = () => syncNowInputPrefixState(input);
    input.addEventListener("input", sync);
    input.addEventListener("change", sync);
    sync();
  });

  ["superShiftDate", "superShiftShift"].forEach((id) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("change", () => {
      if (suppressSupervisorSelectionReset) return;
      if (selectedSupervisorShiftLogId()) {
        updateSupervisorProgressButtonLabels();
        return;
      }
      if (superShiftLogIdInput) superShiftLogIdInput.value = "";
      if (superShiftOpenBreakIdInput) superShiftOpenBreakIdInput.value = "";
      supervisorShiftTileEditId = "";
      updateSupervisorProgressButtonLabels();
    });
  });
  ["superRunDate", "superRunProduct"].forEach((id) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("input", () => {
      if (suppressSupervisorSelectionReset) return;
      updateSupervisorProgressButtonLabels();
    });
    el.addEventListener("change", () => {
      if (suppressSupervisorSelectionReset) return;
      updateSupervisorProgressButtonLabels();
    });
  });
  ["superDownDate", "superDownStart", "superDownReasonCategory", "superDownReasonDetail"].forEach((id) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("input", () => {
      if (suppressSupervisorSelectionReset) return;
      updateSupervisorProgressButtonLabels();
    });
    el.addEventListener("change", () => {
      if (suppressSupervisorSelectionReset) return;
      updateSupervisorProgressButtonLabels();
    });
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
    const visualAllowedShifts = SHIFT_OPTIONS.slice();
    if (!visualAllowedShifts.includes(appState.supervisorSelectedShift)) {
      appState.supervisorSelectedShift = visualAllowedShifts.includes("Full Day") ? "Full Day" : visualAllowedShifts[0] || "Day";
    }
    refreshSupervisorDowntimeDetailOptions(selectedSupervisorLine());
    saveState();
    renderHome();
  });

  svDateInputs.forEach((svDateInput) => {
    svDateInput.addEventListener("change", () => {
      appState.supervisorSelectedDate = svDateInput.value || todayISO();
      saveState();
      renderHome();
    });
  });

  svShiftButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      if (btn.disabled) return;
      const shift = btn.dataset.svShift;
      if (!shift) return;
      appState.supervisorSelectedShift = SHIFT_OPTIONS.includes(shift) ? shift : "Day";
      saveState();
      renderHome();
    });
  });

  if (supervisorDownReasonCategory) {
    supervisorDownReasonCategory.addEventListener("change", () => {
      refreshSupervisorDowntimeDetailOptions(selectedSupervisorLine());
      updateSupervisorProgressButtonLabels();
    });
  }

  svPrevBtns.forEach((svPrevBtn) => {
    svPrevBtn.addEventListener("click", () => {
      const dt = parseDateLocal(appState.supervisorSelectedDate || todayISO());
      dt.setDate(dt.getDate() - 1);
      appState.supervisorSelectedDate = formatDateLocal(dt);
      saveState();
      renderHome();
    });
  });

  svNextBtns.forEach((svNextBtn) => {
    svNextBtn.addEventListener("click", () => {
      const dt = parseDateLocal(appState.supervisorSelectedDate || todayISO());
      dt.setDate(dt.getDate() + 1);
      appState.supervisorSelectedDate = formatDateLocal(dt);
      saveState();
      renderHome();
    });
  });

  document.querySelectorAll("[data-super-main-tab]").forEach((btn) => {
    btn.addEventListener("click", () => {
      appState.supervisorMainTab = btn.dataset.superMainTab || "supervisorDay";
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

  closeEditSupervisorModalBtn.addEventListener("click", closeEditSupervisorModal);
  editSupervisorModal.addEventListener("click", (event) => {
    if (event.target === editSupervisorModal) closeEditSupervisorModal();
  });

  newSupervisorLines.addEventListener("change", (event) => {
    const checkbox = event.target.closest('input[data-new-supervisor-line-shift]');
    if (!checkbox) return;
    syncSupervisorShiftPillStyles(newSupervisorLines);
  });

  supervisorManagerList.addEventListener("change", (event) => {
    const checkbox = event.target.closest('input[data-supervisor-line-shift]');
    if (!checkbox) return;
    syncSupervisorShiftPillStyles(supervisorManagerList);
  });

  supervisorManagerList.addEventListener("click", async (event) => {
    const editBtn = event.target.closest("[data-supervisor-edit]");
    if (editBtn) {
      const supId = editBtn.getAttribute("data-supervisor-edit");
      if (!supId) return;
      openEditSupervisorModal(supId);
      return;
    }

    const saveBtn = event.target.closest("[data-supervisor-save]");
    if (saveBtn) {
      const supId = saveBtn.getAttribute("data-supervisor-save");
      const sup = (appState.supervisors || []).find((item) => item.id === supId);
      if (!sup) return;
      const prevLineShiftMap = normalizeSupervisorLineShifts(sup.assignedLineShifts, appState.lines, sup.assignedLineIds || []);
      const prevAssigned = Object.keys(prevLineShiftMap);
      const nextLineShiftMap = {};
      Array.from(supervisorManagerList.querySelectorAll(`[data-supervisor-line-shift="${supId}"]`)).forEach((input) => {
        const lineId = input.dataset.lineId || "";
        const shift = input.value;
        if (!input.checked || !appState.lines[lineId]) return;
        if (!nextLineShiftMap[lineId]) nextLineShiftMap[lineId] = [];
        if ((shift === "Day" || shift === "Night") && !nextLineShiftMap[lineId].includes(shift)) nextLineShiftMap[lineId].push(shift);
      });
      const nextAssigned = Object.keys(nextLineShiftMap);
      try {
        const session = await ensureManagerBackendSession();
        await apiRequest(`/api/supervisors/${sup.id}/assignments`, {
          method: "PATCH",
          token: session.backendToken,
          body: {
            assignedLineIds: nextAssigned,
            assignedLineShifts: nextLineShiftMap
          }
        });
        sup.assignedLineShifts = nextLineShiftMap;
        sup.assignedLineIds = nextAssigned;
        const added = sup.assignedLineIds.filter((id) => !prevAssigned.includes(id));
        const removed = prevAssigned.filter((id) => !sup.assignedLineIds.includes(id));
        added.forEach((lineId) => addAudit(appState.lines[lineId], "ASSIGN_SUPERVISOR", `${sup.name} assigned to line`));
        removed.forEach((lineId) => addAudit(appState.lines[lineId], "UNASSIGN_SUPERVISOR", `${sup.name} removed from line`));
        if (appState.supervisorSession?.username === sup.username) {
          appState.supervisorSession.assignedLineIds = sup.assignedLineIds.slice();
          appState.supervisorSession.assignedLineShifts = clone(sup.assignedLineShifts || {});
          if (!appState.supervisorSession.assignedLineIds.includes(appState.supervisorSelectedLineId)) {
            appState.supervisorSelectedLineId = appState.supervisorSession.assignedLineIds[0] || "";
          }
          if (!supervisorCanAccessShift(appState.supervisorSession, appState.supervisorSelectedLineId, appState.supervisorSelectedShift)) {
            const fallbackShifts = normalizeSupervisorShifts(appState.supervisorSession.assignedLineShifts?.[appState.supervisorSelectedLineId], { fallbackToAll: false });
            appState.supervisorSelectedShift = fallbackShifts[0] || "Day";
          }
        }
        const allShiftLineIds = Array.from(new Set([...Object.keys(prevLineShiftMap), ...Object.keys(nextLineShiftMap)]))
          .filter((lineId) => appState.lines[lineId]);
        allShiftLineIds.forEach((lineId) => {
          const prevShifts = normalizeSupervisorShifts(prevLineShiftMap[lineId], { fallbackToAll: false });
          const nextShifts = normalizeSupervisorShifts(nextLineShiftMap[lineId], { fallbackToAll: false });
          SUPERVISOR_SHIFT_OPTIONS.forEach((shift) => {
            if (!prevShifts.includes(shift) && nextShifts.includes(shift)) {
              addAudit(appState.lines[lineId], "ASSIGN_SUPERVISOR_SHIFT", `${sup.name} granted ${shift} shift access`);
            }
            if (prevShifts.includes(shift) && !nextShifts.includes(shift)) {
              addAudit(appState.lines[lineId], "UNASSIGN_SUPERVISOR_SHIFT", `${sup.name} removed from ${shift} shift access`);
            }
          });
        });
        saveState();
        renderHome();
        openManageSupervisorsModal();
      } catch (error) {
        console.warn("Supervisor assignment sync failed:", error);
        alert(`Could not save supervisor assignments.\n${error?.message || "Please try again."}`);
      }
      return;
    }

    const delBtn = event.target.closest("[data-supervisor-delete]");
    if (delBtn) {
      const supId = delBtn.getAttribute("data-supervisor-delete");
      const sup = (appState.supervisors || []).find((item) => item.id === supId);
      if (!sup) return;
      if (!window.confirm(`Delete supervisor "${sup.name}"?`)) return;
      try {
        const session = await ensureManagerBackendSession();
        await apiRequest(`/api/supervisors/${supId}`, {
          method: "DELETE",
          token: session.backendToken
        });
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
      } catch (error) {
        console.warn("Supervisor delete sync failed:", error);
        alert(`Could not delete supervisor.\n${error?.message || "Please try again."}`);
      }
    }
  });

  editSupervisorForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const sup = (appState.supervisors || []).find((item) => item.id === editingSupervisorId);
    if (!sup) return;
    const name = String(editSupervisorNameInput.value || "").trim();
    const username = String(editSupervisorUsernameInput.value || "").trim().toLowerCase();
    const password = String(editSupervisorPasswordInput.value || "").trim();
    if (!name || !username) {
      alert("Name and username are required.");
      return;
    }
    if (password && password.length < 6) {
      alert("New password must be at least 6 characters.");
      return;
    }
    try {
      const session = await ensureManagerBackendSession();
      await apiRequest(`/api/supervisors/${sup.id}`, {
        method: "PATCH",
        token: session.backendToken,
        body: {
          name,
          username,
          password
        }
      });
      const previousUsername = sup.username;
      sup.name = name;
      sup.username = username;
      if (appState.supervisorSession?.username === previousUsername) {
        appState.supervisorSession.username = username;
      }
      closeEditSupervisorModal();
      saveState();
      await refreshHostedState();
      renderAll();
      openManageSupervisorsModal();
    } catch (error) {
      console.warn("Supervisor update sync failed:", error);
      alert(`Could not update supervisor.\n${error?.message || "Please try again."}`);
    }
  });

  addSupervisorForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const name = String(document.getElementById("newSupervisorName").value || "").trim();
    const username = String(document.getElementById("newSupervisorUsername").value || "").trim().toLowerCase();
    const password = String(document.getElementById("newSupervisorPassword").value || "").trim();
    const assignedLineShifts = {};
    Array.from(newSupervisorLines.querySelectorAll("[data-new-supervisor-line-shift]")).forEach((input) => {
      const lineId = input.dataset.lineId || "";
      const shift = input.value;
      if (!input.checked || !appState.lines[lineId]) return;
      if (!assignedLineShifts[lineId]) assignedLineShifts[lineId] = [];
      if ((shift === "Day" || shift === "Night") && !assignedLineShifts[lineId].includes(shift)) {
        assignedLineShifts[lineId].push(shift);
      }
    });
    const assignedLineIds = Object.keys(assignedLineShifts);
    if (!name || !username || !password) return;
    if (supervisorByUsername(username)) {
      alert("Username already exists.");
      return;
    }
    try {
      const session = await ensureManagerBackendSession();
      const payload = await apiRequest("/api/supervisors", {
        method: "POST",
        token: session.backendToken,
        body: {
          name,
          username,
          password,
          assignedLineIds,
          assignedLineShifts
        }
      });
      const createdId = payload?.supervisor?.id;
      if (!createdId) throw new Error("Supervisor was not created on server.");
      appState.supervisors = Array.isArray(appState.supervisors) ? appState.supervisors : [];
      appState.supervisors.push({
        id: createdId,
        name,
        username,
        password: "",
        assignedLineIds,
        assignedLineShifts
      });
      assignedLineIds.forEach((lineId) => addAudit(appState.lines[lineId], "CREATE_SUPERVISOR", `Supervisor ${name} created and assigned`));
      saveState();
      closeAddSupervisorModal();
      await refreshHostedState();
      renderAll();
    } catch (error) {
      console.warn("Supervisor create sync failed:", error);
      alert(`Could not create supervisor.\n${error?.message || "Please try again."}`);
    }
  });

  const latestBySubmittedAt = (rows = []) =>
    rows.reduce((latest, row) => {
      if (!latest) return row;
      const a = Date.parse(row?.submittedAt || "") || 0;
      const b = Date.parse(latest?.submittedAt || "") || 0;
      return a >= b ? row : latest;
    }, null);

  const upsertRowById = (rows, nextRow) => {
    if (!Array.isArray(rows) || !nextRow) return;
    if (nextRow.id) {
      const idx = rows.findIndex((row) => row.id === nextRow.id);
      if (idx >= 0) {
        rows[idx] = { ...rows[idx], ...nextRow };
        return;
      }
    }
    rows.push(nextRow);
  };

  const supervisorActorName = (session) => {
    const sup = supervisorByUsername(session?.username || "");
    return sup?.name || session?.username || "supervisor";
  };

  const resolveSupervisorLineContext = (lineIdOverride = "") => {
    const session = appState.supervisorSession;
    if (!session) return null;
    const lineId = String(lineIdOverride || selectedSupervisorLineId() || "");
    const line = appState.lines[lineId];
    if (!lineId || !line || !session.assignedLineIds.includes(lineId)) return null;
    return { session, lineId, line };
  };

  const isOpenRunRow = (row) =>
    strictTimeValid(row?.productionStartTime) &&
    strictTimeValid(row?.finishTime) &&
    row.productionStartTime === row.finishTime;

  const isOpenShiftRow = (row) =>
    strictTimeValid(row?.startTime) &&
    strictTimeValid(row?.finishTime) &&
    row.startTime === row.finishTime;

  const isOpenBreakRow = (row) =>
    strictTimeValid(row?.breakStart) &&
    !strictTimeValid(row?.breakFinish);

  const isOpenDowntimeRow = (row) =>
    strictTimeValid(row?.downtimeStart) &&
    strictTimeValid(row?.downtimeFinish) &&
    row.downtimeStart === row.downtimeFinish;

  supervisorShiftForm.addEventListener("submit", (event) => event.preventDefault());
  supervisorRunForm.addEventListener("submit", (event) => event.preventDefault());
  supervisorDownForm.addEventListener("submit", (event) => event.preventDefault());
  if (superRunCrewingPatternBtn) {
    superRunCrewingPatternBtn.addEventListener("click", () => {
      const line = selectedSupervisorLine();
      if (!line) {
        alert("No assigned line selected.");
        return;
      }
      const runDate = String(document.getElementById("superRunDate")?.value || appState.supervisorSelectedDate || todayISO());
      const runStart = String(document.getElementById("superRunProdStart")?.value || nowTimeHHMM());
      const shiftForPattern = inferShiftForLog(line, runDate, runStart, appState.supervisorSelectedShift || "Day");
      openRunCrewingPatternModal({
        line,
        shift: shiftForPattern,
        inputEl: superRunCrewingPatternInput,
        summaryEl: superRunCrewingPatternSummary
      });
    });
  }

  const submitSupervisorShift = async ({ complete = false } = {}) => {
    const session = appState.supervisorSession;
    if (!session) return;
    const lineId = selectedSupervisorLineId();
    const line = appState.lines[lineId];
    if (!session.assignedLineIds.includes(lineId) || !line) {
      alert("You are not assigned to that line.");
      return;
    }

    const date = document.getElementById("superShiftDate").value || todayISO();
    const shift = document.getElementById("superShiftShift").value || "Day";
    const startInput = document.getElementById("superShiftStart").value || "";
    const finishInput = document.getElementById("superShiftFinish").value || "";
    const crewRaw = document.getElementById("superShiftCrew").value;

    if (!rowIsValidDateShift(date, shift)) {
      alert("Date/shift are invalid.");
      return;
    }
    if (!supervisorCanAccessShift(session, lineId, shift)) {
      alert(`You are not assigned to the ${shift} shift.`);
      return;
    }

    const selectedShiftId = selectedSupervisorShiftLogId({ line });
    const selectedShiftTileNode = selectedShiftId && superShiftOpenList
      ? superShiftOpenList
          .querySelector(`[data-super-shift-edit="${selectedShiftId}"]`)
          ?.closest(".pending-log-item") || null
      : null;
    let existing = null;
    if (selectedShiftId) {
      existing = (line.shiftRows || []).find((row) => row.id === selectedShiftId) || null;
      if (!existing) {
        alert("Selected shift log could not be found. Re-open it from Pending Shift Logs and try again.");
        return;
      }
      if (!isOpenShiftRow(existing)) {
        alert("Selected shift log is no longer open.");
        return;
      }
      if (existing.date !== date || existing.shift !== shift) {
        alert("Date and shift cannot be changed while editing an open shift log.");
        return;
      }
    }

    const startTime = startInput || existing?.startTime || nowTimeHHMM();
    const finishTime = finishInput || (complete ? nowTimeHHMM() : startTime);
    const crewOnShift = crewRaw === ""
      ? Math.max(0, Math.floor(num(existing?.crewOnShift)))
      : Math.max(0, Math.floor(num(crewRaw)));

    if (!strictTimeValid(startTime) || !strictTimeValid(finishTime)) {
      alert("Shift start and finish must be in HH:MM (24h).");
      return;
    }
    if (!complete && finishTime !== startTime) {
      alert("Open shift logs keep finish equal to start. Use the Finalise pill on the log when the shift ends.");
      return;
    }
    if (complete && superShiftOpenBreakIdInput?.value) {
      alert("End the current break before completing the shift.");
      return;
    }

    const payload = { lineId, date, shift, crewOnShift, startTime, finishTime };
    try {
      const saved = existing?.id
        ? await patchSupervisorShiftLog(session, existing.id, payload)
        : await syncSupervisorShiftLog(session, payload);

      const savedRow = {
        ...(existing || {}),
        ...payload,
        id: saved?.id || existing?.id || "",
        submittedBy: supervisorActorName(session),
        submittedAt: saved?.submittedAt || nowIso()
      };
      upsertRowById(line.shiftRows, savedRow);
      if (existing?.id && selectedShiftTileNode) {
        await saveSupervisorShiftBreakEdits(existing.id, selectedShiftTileNode, {
          suppressNoChangesAlert: true,
          skipPersist: true
        });
      }
      if (complete) {
        superShiftLogIdInput.value = "";
        if (superShiftOpenBreakIdInput) superShiftOpenBreakIdInput.value = "";
        supervisorShiftTileEditId = "";
      } else {
        if (superShiftLogIdInput) superShiftLogIdInput.value = "";
        if (superShiftOpenBreakIdInput) superShiftOpenBreakIdInput.value = "";
        supervisorShiftTileEditId = "";
      }
      addAudit(
        line,
        complete ? "SUPERVISOR_SHIFT_COMPLETE" : "SUPERVISOR_SHIFT_PROGRESS",
        `${supervisorActorName(session)} ${complete ? "completed" : "updated"} ${shift} shift for ${date}`
      );
      if (complete) {
        supervisorShiftForm.reset();
        document.getElementById("superShiftDate").value = date;
        document.getElementById("superShiftShift").value = shift;
      }
      saveState();
      renderAll();
      if (!complete) {
        supervisorShiftForm.reset();
        if (superShiftLogIdInput) superShiftLogIdInput.value = "";
        if (superShiftOpenBreakIdInput) superShiftOpenBreakIdInput.value = "";
        if (superShiftBreakTimeInput) superShiftBreakTimeInput.value = "";
        supervisorShiftTileEditId = "";
      }
    } catch (error) {
      alert(`Could not save shift log.\n${error?.message || "Please try again."}`);
    }
  };

  const submitSupervisorBreakStart = async ({ shiftId = "", breakStartOverride = "" } = {}) => {
    const session = appState.supervisorSession;
    if (!session) return;
    const lineId = selectedSupervisorLineId();
    const line = appState.lines[lineId];
    if (!session.assignedLineIds.includes(lineId) || !line) {
      alert("You are not assigned to that line.");
      return;
    }

    let shiftLog = null;
    if (shiftId) {
      shiftLog = (line.shiftRows || []).find((row) => row.id === shiftId) || null;
      if (!shiftLog?.id) {
        alert("Shift log could not be found.");
        return;
      }
      if (!isOpenShiftRow(shiftLog)) {
        alert("This shift is already finalised.");
        return;
      }
    }

    const date = shiftLog?.date || document.getElementById("superShiftDate").value || todayISO();
    const shift = shiftLog?.shift || document.getElementById("superShiftShift").value || "Day";
    const breakStart = String(breakStartOverride || superShiftBreakTimeInput?.value || "").trim() || nowTimeHHMM();
    if (!strictTimeValid(breakStart)) {
      alert("Break start time must be HH:MM.");
      return;
    }
    if (!supervisorCanAccessShift(session, lineId, shift)) {
      alert(`You are not assigned to the ${shift} shift.`);
      return;
    }

    if (!shiftLog && superShiftLogIdInput?.value) {
      shiftLog = (line.shiftRows || []).find((row) => row.id === superShiftLogIdInput.value) || null;
      if (shiftLog && (shiftLog.date !== date || shiftLog.shift !== shift)) shiftLog = null;
    }
    if (!shiftLog) {
      shiftLog = latestBySubmittedAt((line.shiftRows || []).filter((row) => row.date === date && row.shift === shift && isOpenShiftRow(row)));
    }
    if (!shiftLog?.id) {
      alert("Start the shift first, then you can log breaks.");
      return;
    }
    const openBreak = latestBySubmittedAt((line.breakRows || []).filter((row) => row.shiftLogId === shiftLog.id && isOpenBreakRow(row)));
    if (openBreak?.id) {
      alert("There is already an open break. End it before starting a new break.");
      return;
    }
    try {
      const savedBreak = await startSupervisorShiftBreak(session, shiftLog.id, breakStart);
      const savedRow = {
        date,
        shift,
        shiftLogId: shiftLog.id,
        breakStart,
        breakFinish: "",
        id: savedBreak?.id || "",
        submittedBy: supervisorActorName(session),
        submittedAt: savedBreak?.submittedAt || nowIso()
      };
      upsertRowById(line.breakRows, savedRow);
      if (superShiftOpenBreakIdInput && superShiftLogIdInput?.value === shiftLog.id) superShiftOpenBreakIdInput.value = savedRow.id || "";
      if (superShiftBreakTimeInput) superShiftBreakTimeInput.value = "";
      addAudit(line, "SUPERVISOR_BREAK_START", `${supervisorActorName(session)} started break (${shift} ${date})`);
      saveState();
      renderAll();
    } catch (error) {
      alert(`Could not start break.\n${error?.message || "Please try again."}`);
    }
  };

  const submitSupervisorBreakEnd = async ({ shiftId = "", breakId = "", breakFinishOverride = "" } = {}) => {
    const session = appState.supervisorSession;
    if (!session) return;
    const lineId = selectedSupervisorLineId();
    const line = appState.lines[lineId];
    if (!session.assignedLineIds.includes(lineId) || !line) {
      alert("You are not assigned to that line.");
      return;
    }

    let shiftLog = null;
    if (shiftId) {
      shiftLog = (line.shiftRows || []).find((row) => row.id === shiftId) || null;
      if (!shiftLog?.id) {
        alert("Shift log could not be found.");
        return;
      }
    }

    const date = shiftLog?.date || document.getElementById("superShiftDate").value || todayISO();
    const shift = shiftLog?.shift || document.getElementById("superShiftShift").value || "Day";
    const breakFinish = String(breakFinishOverride || superShiftBreakTimeInput?.value || "").trim() || nowTimeHHMM();
    if (!strictTimeValid(breakFinish)) {
      alert("Break finish time must be HH:MM.");
      return;
    }
    if (!supervisorCanAccessShift(session, lineId, shift)) {
      alert(`You are not assigned to the ${shift} shift.`);
      return;
    }

    if (!shiftLog && superShiftLogIdInput?.value) {
      shiftLog = (line.shiftRows || []).find((row) => row.id === superShiftLogIdInput.value) || null;
      if (shiftLog && (shiftLog.date !== date || shiftLog.shift !== shift)) shiftLog = null;
    }
    if (!shiftLog) {
      shiftLog = latestBySubmittedAt((line.shiftRows || []).filter((row) => row.date === date && row.shift === shift && isOpenShiftRow(row)));
    }
    if (!shiftLog?.id) {
      alert("No open shift found for this date/shift.");
      return;
    }
    let openBreak = null;
    if (breakId) {
      openBreak = (line.breakRows || []).find((row) => row.id === breakId && row.shiftLogId === shiftLog.id) || null;
    }
    if (!openBreak && superShiftOpenBreakIdInput?.value) {
      openBreak = (line.breakRows || []).find((row) => row.id === superShiftOpenBreakIdInput.value) || null;
      if (openBreak && openBreak.shiftLogId !== shiftLog.id) openBreak = null;
    }
    if (!openBreak) {
      openBreak = latestBySubmittedAt((line.breakRows || []).filter((row) => row.shiftLogId === shiftLog.id && isOpenBreakRow(row)));
    }
    if (!openBreak?.id) {
      alert("No open break found for this shift.");
      return;
    }
    try {
      const savedBreak = await endSupervisorShiftBreak(session, shiftLog.id, openBreak.id, breakFinish);
      const savedRow = {
        ...openBreak,
        breakFinish,
        submittedBy: supervisorActorName(session),
        submittedAt: savedBreak?.submittedAt || nowIso()
      };
      upsertRowById(line.breakRows, savedRow);
      if (superShiftOpenBreakIdInput && superShiftLogIdInput?.value === shiftLog.id) superShiftOpenBreakIdInput.value = "";
      if (superShiftBreakTimeInput) superShiftBreakTimeInput.value = "";
      addAudit(line, "SUPERVISOR_BREAK_END", `${supervisorActorName(session)} ended break (${shift} ${date})`);
      saveState();
      renderAll();
    } catch (error) {
      alert(`Could not end break.\n${error?.message || "Please try again."}`);
    }
  };

  const saveSupervisorShiftBreakEdits = async (
    shiftId,
    tileNode,
    { suppressNoChangesAlert = false, skipPersist = false } = {}
  ) => {
    const session = appState.supervisorSession;
    if (!session) return;
    const lineId = selectedSupervisorLineId();
    const line = appState.lines[lineId];
    if (!session.assignedLineIds.includes(lineId) || !line) {
      alert("You are not assigned to that line.");
      return;
    }
    const shiftLog = (line.shiftRows || []).find((row) => row.id === shiftId) || null;
    if (!shiftLog?.id) {
      alert("Shift log could not be found.");
      return;
    }
    const editorRows = Array.from(tileNode?.querySelectorAll("[data-break-editor-row]") || []);
    if (!editorRows.length) return;

    try {
      let changed = false;
      const warnings = new Set();
      for (const editorRow of editorRows) {
        const breakId = String(editorRow.getAttribute("data-break-id") || "");
        const breakIndex = Number(editorRow.getAttribute("data-break-index"));
        const startValue = String(editorRow.querySelector('[data-break-field="start"]')?.value || "").trim();
        const finishValue = String(editorRow.querySelector('[data-break-field="finish"]')?.value || "").trim();
        if (startValue && !strictTimeValid(startValue)) {
          alert("Break start time must be HH:MM.");
          return;
        }
        if (finishValue && !strictTimeValid(finishValue)) {
          alert("Break finish time must be HH:MM.");
          return;
        }

        const shiftBreakRows = breakRowsForShift(line, shiftId);
        let existingBreak = null;
        if (breakId) {
          existingBreak = shiftBreakRows.find((row) => String(row.id || "") === breakId) || null;
        }
        if (!existingBreak && Number.isInteger(breakIndex) && breakIndex >= 0) {
          existingBreak = shiftBreakRows[breakIndex] || null;
        }

        if (existingBreak) {
          const originalStart = String(existingBreak.breakStart || "").trim();
          const originalFinish = String(existingBreak.breakFinish || "").trim();
          const nextStart = startValue || originalStart;
          const nextFinish = finishValue;
          if (!nextStart || !strictTimeValid(nextStart)) {
            alert("Each existing break must have a valid start time.");
            return;
          }
          const isBackendBreak = UUID_RE.test(String(existingBreak.id || ""));
          if (isBackendBreak) {
            const breakPatch = {};
            if (nextStart !== originalStart) breakPatch.breakStart = nextStart;
            if (nextFinish !== originalFinish) {
              if (!nextFinish && originalFinish) {
                warnings.add("Break finish cannot be cleared once set.");
              } else if (nextFinish) {
                breakPatch.breakFinish = nextFinish;
              }
            }
            if (Object.keys(breakPatch).length) {
              const savedBreak = await patchSupervisorShiftBreak(session, shiftLog.id, existingBreak.id, breakPatch);
              const savedStart = savedBreak?.breakStart || breakPatch.breakStart || originalStart;
              const savedFinish = savedBreak?.breakFinish !== undefined
                ? String(savedBreak.breakFinish || "")
                : breakPatch.breakFinish !== undefined
                  ? breakPatch.breakFinish
                  : originalFinish;
              Object.assign(existingBreak, {
                breakStart: savedStart,
                breakFinish: savedFinish,
                submittedBy: supervisorActorName(session),
                submittedAt: savedBreak?.submittedAt || nowIso()
              });
              changed = true;
            }
          } else {
            if (nextStart !== originalStart || nextFinish !== originalFinish) {
              Object.assign(existingBreak, {
                breakStart: nextStart,
                breakFinish: nextFinish,
                submittedBy: supervisorActorName(session),
                submittedAt: nowIso()
              });
              changed = true;
            }
          }
          continue;
        }

        if (!startValue && !finishValue) continue;
        if (!startValue) {
          alert("New break rows need a start time.");
          return;
        }

        if (UUID_RE.test(String(shiftLog.id || ""))) {
          const savedStart = await startSupervisorShiftBreak(session, shiftLog.id, startValue);
          const newBreakRow = {
            id: savedStart?.id || "",
            lineId: lineId || "",
            date: shiftLog.date || todayISO(),
            shift: shiftLog.shift || "Day",
            shiftLogId: shiftLog.id,
            breakStart: startValue,
            breakFinish: "",
            submittedBy: supervisorActorName(session),
            submittedAt: savedStart?.submittedAt || nowIso()
          };
          upsertRowById(line.breakRows, newBreakRow);
          changed = true;
          if (finishValue) {
            if (!newBreakRow.id || !UUID_RE.test(String(newBreakRow.id))) {
              warnings.add("New break was created but finish could not be set immediately.");
            } else {
              const savedFinish = await endSupervisorShiftBreak(session, shiftLog.id, newBreakRow.id, finishValue);
              Object.assign(newBreakRow, {
                breakFinish: finishValue,
                submittedBy: supervisorActorName(session),
                submittedAt: savedFinish?.submittedAt || nowIso()
              });
              upsertRowById(line.breakRows, newBreakRow);
            }
          }
        } else {
          upsertRowById(line.breakRows, {
            id: makeLocalLogId("break"),
            lineId: lineId || "",
            date: shiftLog.date || todayISO(),
            shift: shiftLog.shift || "Day",
            shiftLogId: shiftLog.id,
            breakStart: startValue,
            breakFinish: finishValue || "",
            submittedBy: supervisorActorName(session),
            submittedAt: nowIso()
          });
          changed = true;
        }
      }

      if (warnings.size) alert(Array.from(warnings).join("\n"));
      if (!changed) {
        if (!warnings.size && !suppressNoChangesAlert) alert("No break changes to save.");
        return;
      }
      addAudit(
        line,
        "SUPERVISOR_BREAK_EDIT",
        `${supervisorActorName(session)} updated break logs (${shiftLog.shift || "Shift"} ${shiftLog.date || todayISO()})`
      );
      if (!skipPersist) {
        saveState();
        renderAll();
      }
    } catch (error) {
      alert(`Could not save break edits.\n${error?.message || "Please try again."}`);
    }
  };

  const submitSupervisorRun = async ({ complete = false } = {}) => {
    const session = appState.supervisorSession;
    if (!session) return;
    const lineId = selectedSupervisorLineId();
    const line = appState.lines[lineId];
    if (!session.assignedLineIds.includes(lineId) || !line) {
      alert("You are not assigned to that line.");
      return;
    }

    const date = document.getElementById("superRunDate").value || todayISO();
    const productInput = String(document.getElementById("superRunProduct").value || "").trim();
    const prodStartInput = document.getElementById("superRunProdStart").value || "";
    const finishInput = document.getElementById("superRunFinish").value || "";
    const unitsRaw = document.getElementById("superRunUnits").value;

    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date || ""))) {
      alert("Date is invalid.");
      return;
    }

    const selectedRunLogId = String(superRunLogIdInput?.value || "").trim();
    let existing = null;
    if (selectedRunLogId) {
      existing = (line.runRows || []).find((row) => row.id === selectedRunLogId) || null;
      if (!existing) {
        alert("Selected production run could not be found. Re-open it from Pending Production Runs and try again.");
        return;
      }
      if (existing.date !== date) {
        alert("Date cannot be changed while editing an open production run.");
        return;
      }
    }

    const product = productInput || existing?.product || "";
    if (!product) {
      alert("Product is required.");
      return;
    }

    const productionStartTime = prodStartInput || existing?.productionStartTime || nowTimeHHMM();
    const finishTime = finishInput || (complete ? nowTimeHHMM() : productionStartTime);
    const unitsProduced = unitsRaw === ""
      ? Math.max(0, num(existing?.unitsProduced))
      : Math.max(0, num(unitsRaw));
    const backendShift = inferShiftForLog(line, date, productionStartTime, appState.supervisorSelectedShift || "Day");

    if (!strictTimeValid(productionStartTime) || !strictTimeValid(finishTime)) {
      alert("Production start and finish must be HH:MM (24h).");
      return;
    }
    if (!complete && finishTime !== productionStartTime) {
      alert("Open run logs keep finish equal to start. Use the Finalise pill on the log when the run ends.");
      return;
    }
    if (!supervisorCanAccessShift(session, lineId, backendShift)) {
      alert(`You are not assigned to the ${backendShift} shift.`);
      return;
    }

    const inputPattern = runCrewingPatternFromInput(superRunCrewingPatternInput, line, backendShift, { fallbackToIdeal: false });
    const existingPattern = normalizeRunCrewingPattern(existing?.runCrewingPattern, line, backendShift, { fallbackToIdeal: false });
    const runCrewingPattern = Object.keys(inputPattern).length ? inputPattern : existingPattern;
    if (!Object.keys(runCrewingPattern).length) {
      alert("Set crewing pattern for this run before saving.");
      return;
    }
    setRunCrewingPatternField(
      superRunCrewingPatternInput,
      superRunCrewingPatternSummary,
      line,
      backendShift,
      runCrewingPattern,
      { fallbackToIdeal: false }
    );

    const payload = {
      lineId,
      date,
      shift: backendShift,
      setUpStartTime: "",
      product,
      productionStartTime,
      finishTime,
      unitsProduced,
      runCrewingPattern
    };
    try {
      const saved = existing?.id
        ? await patchSupervisorRunLog(session, existing.id, payload)
        : await syncSupervisorRunLog(session, payload);

      const savedRow = {
        ...(existing || {}),
        ...payload,
        shift: "",
        runCrewingPattern: normalizeRunCrewingPattern(saved?.runCrewingPattern || runCrewingPattern, line, backendShift, { fallbackToIdeal: false }),
        id: saved?.id || existing?.id || "",
        submittedBy: supervisorActorName(session),
        submittedAt: saved?.submittedAt || nowIso()
      };
      upsertRowById(line.runRows, savedRow);
      if (complete) {
        superRunLogIdInput.value = "";
      } else {
        if (superRunLogIdInput) superRunLogIdInput.value = "";
      }
      addAudit(
        line,
        complete ? "SUPERVISOR_RUN_COMPLETE" : "SUPERVISOR_RUN_PROGRESS",
        `${supervisorActorName(session)} ${complete ? "completed" : "updated"} run ${product}`
      );
      if (complete) {
        supervisorRunForm.reset();
        document.getElementById("superRunDate").value = date;
        setRunCrewingPatternField(
          superRunCrewingPatternInput,
          superRunCrewingPatternSummary,
          line,
          appState.supervisorSelectedShift || "Day",
          {},
          { fallbackToIdeal: false }
        );
      }
      saveState();
      renderAll();
      if (!complete) {
        supervisorRunForm.reset();
        if (superRunLogIdInput) superRunLogIdInput.value = "";
        setRunCrewingPatternField(
          superRunCrewingPatternInput,
          superRunCrewingPatternSummary,
          line,
          appState.supervisorSelectedShift || "Day",
          {},
          { fallbackToIdeal: false }
        );
      }
    } catch (error) {
      alert(`Could not save production run.\n${error?.message || "Please try again."}`);
    }
  };

  const formatMinutesToHHMM = (minutes) => {
    const minsPerDay = 24 * 60;
    const normalized = ((Math.floor(num(minutes)) % minsPerDay) + minsPerDay) % minsPerDay;
    const hh = String(Math.floor(normalized / 60)).padStart(2, "0");
    const mm = String(normalized % 60).padStart(2, "0");
    return `${hh}:${mm}`;
  };

  const loadSupervisorShiftForEdit = (shiftId) => {
    const session = appState.supervisorSession;
    if (!session) return;
    const lineId = selectedSupervisorLineId();
    const line = appState.lines[lineId];
    if (!line || !session.assignedLineIds.includes(lineId)) return;
    const targetShiftId = String(shiftId || "");
    const row = (line.shiftRows || []).find((item) => String(item?.id || "") === targetShiftId);
    if (!row) {
      renderHome();
      updateSupervisorProgressButtonLabels();
      return;
    }

    const openBreak = latestBySubmittedAt((line.breakRows || []).filter((breakRow) => breakRow.shiftLogId === row.id && isOpenBreakRow(breakRow)));
    suppressSupervisorSelectionReset = true;
    document.getElementById("superShiftDate").value = row.date || todayISO();
    document.getElementById("superShiftShift").value = row.shift || "Day";
    document.getElementById("superShiftCrew").value = String(Math.max(0, Math.floor(num(row.crewOnShift))));
    document.getElementById("superShiftStart").value = row.startTime || "";
    document.getElementById("superShiftFinish").value = isOpenShiftRow(row) ? "" : row.finishTime || "";
    if (superShiftBreakTimeInput) superShiftBreakTimeInput.value = "";
    superShiftLogIdInput.value = row.id || "";
    supervisorShiftTileEditId = row.id || "";
    if (superShiftOpenBreakIdInput) superShiftOpenBreakIdInput.value = openBreak?.id || "";
    if (shiftBreakStartBtn) shiftBreakStartBtn.disabled = !row.id || Boolean(openBreak?.id);
    if (shiftBreakEndBtn) shiftBreakEndBtn.disabled = !row.id || !openBreak?.id;
    suppressSupervisorSelectionReset = false;
    // Refresh the supervisor data entry UI so the selected pending shift shows Save immediately.
    renderHome();
    updateSupervisorProgressButtonLabels();
    const crewInput = document.getElementById("superShiftCrew");
    if (crewInput) crewInput.focus();
  };

  const loadSupervisorRunForEdit = (runId) => {
    const session = appState.supervisorSession;
    if (!session) return;
    const lineId = selectedSupervisorLineId();
    const line = appState.lines[lineId];
    if (!line || !session.assignedLineIds.includes(lineId)) return;
    const row = (line.runRows || []).find((item) => item.id === runId);
    if (!row) return;

    suppressSupervisorSelectionReset = true;
    document.getElementById("superRunDate").value = row.date || todayISO();
    document.getElementById("superRunProduct").value = row.product || "";
    document.getElementById("superRunProdStart").value = row.productionStartTime || "";
    const rowIsOpen = isOpenRunRow(row);
    document.getElementById("superRunFinish").value = rowIsOpen ? "" : row.finishTime || "";
    document.getElementById("superRunUnits").value = num(row.unitsProduced) > 0 ? formatNum(Math.max(0, num(row.unitsProduced)), 0) : "";
    superRunLogIdInput.value = row.id || "";
    const rowShift = row.shift || inferShiftForLog(line, row.date, row.productionStartTime, appState.supervisorSelectedShift || "Day");
    setRunCrewingPatternField(
      superRunCrewingPatternInput,
      superRunCrewingPatternSummary,
      line,
      rowShift,
      row.runCrewingPattern,
      { fallbackToIdeal: true }
    );
    suppressSupervisorSelectionReset = false;
    updateSupervisorProgressButtonLabels();
    const productInput = document.getElementById("superRunProduct");
    if (productInput) productInput.focus();
  };

  const loadSupervisorDowntimeForEdit = (downtimeId) => {
    const session = appState.supervisorSession;
    if (!session) return;
    const lineId = selectedSupervisorLineId();
    const line = appState.lines[lineId];
    if (!line || !session.assignedLineIds.includes(lineId)) return;
    const row = (line.downtimeRows || []).find((item) => item.id === downtimeId);
    if (!row) return;

    const parsedReason = parseDowntimeReasonParts(row.reason, row.equipment);
    const reasonCategory = row.reasonCategory || parsedReason.reasonCategory;
    const reasonDetail = row.reasonDetail || parsedReason.reasonDetail;
    const reasonNote = String((row.reasonNote ?? parsedReason.reasonNote) || "");
    suppressSupervisorSelectionReset = true;
    document.getElementById("superDownDate").value = row.date || todayISO();
    document.getElementById("superDownStart").value = row.downtimeStart || "";
    document.getElementById("superDownFinish").value = isOpenDowntimeRow(row) ? "" : row.downtimeFinish || "";
    if (supervisorDownReasonCategory) supervisorDownReasonCategory.value = reasonCategory || "";
    if (supervisorDownReasonDetail) {
      setDowntimeDetailOptions(supervisorDownReasonDetail, line, reasonCategory || "", reasonDetail || "");
    }
    document.getElementById("superDownReasonNote").value = reasonNote;
    superDownLogIdInput.value = row.id || "";
    suppressSupervisorSelectionReset = false;
    updateSupervisorProgressButtonLabels();
    const startInput = document.getElementById("superDownStart");
    if (startInput) startInput.focus();
  };

  const completeSupervisorShiftById = async (shiftId, { lineId: lineIdOverride } = {}) => {
    const context = resolveSupervisorLineContext(lineIdOverride);
    if (!context) {
      alert("You are not assigned to that line.");
      return;
    }
    const { session, lineId, line } = context;
    const row = (line.shiftRows || []).find((item) => item.id === shiftId);
    if (!row) {
      alert("Shift log could not be found.");
      return;
    }
    if (!isOpenShiftRow(row)) {
      alert("This shift is already finalised.");
      return;
    }
    const startTime = row.startTime || nowTimeHHMM();
    if (!strictTimeValid(startTime)) {
      alert("Shift start time is invalid. Edit the shift and set a valid start time first.");
      return;
    }
    const openBreak = latestBySubmittedAt((line.breakRows || []).filter((breakRow) => breakRow.shiftLogId === row.id && isOpenBreakRow(breakRow)));
    if (openBreak?.id) {
      alert("End the current break before finalising the shift.");
      return;
    }
    const nowCandidate = nowTimeHHMM();
    const startMins = parseTimeToMinutes(startTime);
    const nowMins = parseTimeToMinutes(nowCandidate);
    const defaultFinish = Number.isFinite(startMins) && Number.isFinite(nowMins) && startMins === nowMins
      ? formatMinutesToHHMM(startMins + 1)
      : nowCandidate;
    const finishPrompt = window.prompt("Finish time (HH:MM)", defaultFinish);
    if (finishPrompt === null) return;
    const finishTime = String(finishPrompt || "").trim();
    if (!strictTimeValid(finishTime)) {
      alert("Finish time must be HH:MM (24h).");
      return;
    }
    if (finishTime === startTime) {
      alert("Finish time must be different from start time.");
      return;
    }
    const shift = row.shift || inferShiftForLog(line, row.date, startTime, appState.supervisorSelectedShift || "Day");
    if (!supervisorCanAccessShift(session, lineId, shift)) {
      alert(`You are not assigned to the ${shift} shift.`);
      return;
    }
    const payload = {
      lineId,
      date: row.date,
      shift,
      crewOnShift: Math.max(0, Math.floor(num(row.crewOnShift))),
      startTime,
      finishTime
    };
    try {
      const saved = await patchSupervisorShiftLog(session, row.id, payload);
      const savedRow = {
        ...row,
        ...payload,
        id: saved?.id || row.id || "",
        submittedBy: supervisorActorName(session),
        submittedAt: saved?.submittedAt || nowIso()
      };
      upsertRowById(line.shiftRows, savedRow);
      if (superShiftLogIdInput.value === row.id) superShiftLogIdInput.value = "";
      if (superShiftOpenBreakIdInput?.value) {
        const activeBreak = (line.breakRows || []).find((breakRow) => breakRow.id === superShiftOpenBreakIdInput.value);
        if (activeBreak?.shiftLogId === row.id) superShiftOpenBreakIdInput.value = "";
      }
      addAudit(
        line,
        "SUPERVISOR_SHIFT_COMPLETE",
        `${supervisorActorName(session)} completed ${shift} shift for ${row.date}`
      );
      saveState();
      renderAll();
    } catch (error) {
      alert(`Could not complete shift log.\n${error?.message || "Please try again."}`);
    }
  };

  const completeSupervisorRunById = async (runId, { lineId: lineIdOverride } = {}) => {
    const context = resolveSupervisorLineContext(lineIdOverride);
    if (!context) {
      alert("You are not assigned to that line.");
      return;
    }
    const { session, lineId, line } = context;
    const row = (line.runRows || []).find((item) => item.id === runId);
    if (!row) {
      alert("Run log could not be found.");
      return;
    }
    if (!isOpenRunRow(row)) {
      alert("This production run is already finalised.");
      return;
    }
    const productionStartTime = row.productionStartTime || nowTimeHHMM();
    if (!strictTimeValid(productionStartTime)) {
      alert("Run start time is invalid. Edit the run and set a valid start time first.");
      return;
    }
    const nowCandidate = nowTimeHHMM();
    const startMins = parseTimeToMinutes(productionStartTime);
    const nowMins = parseTimeToMinutes(nowCandidate);
    const defaultFinish = Number.isFinite(startMins) && Number.isFinite(nowMins) && startMins === nowMins
      ? formatMinutesToHHMM(startMins + 1)
      : nowCandidate;
    const finishPrompt = window.prompt("Finish time (HH:MM)", defaultFinish);
    if (finishPrompt === null) return;
    const finishTime = String(finishPrompt || "").trim();
    if (!strictTimeValid(finishTime)) {
      alert("Finish time must be HH:MM (24h).");
      return;
    }
    if (finishTime === productionStartTime) {
      alert("Finish time must be different from start time.");
      return;
    }
    const unitsPrompt = window.prompt("Units produced", String(Math.max(0, num(row.unitsProduced))));
    if (unitsPrompt === null) return;
    const unitsProduced = Math.max(0, num(unitsPrompt));
    const shift = row.shift || inferShiftForLog(line, row.date, productionStartTime, appState.supervisorSelectedShift || "Day");
    if (!supervisorCanAccessShift(session, lineId, shift)) {
      alert(`You are not assigned to the ${shift} shift.`);
      return;
    }
    const payload = {
      lineId,
      date: row.date,
      shift,
      setUpStartTime: "",
      product: row.product || "Run",
      productionStartTime,
      finishTime,
      unitsProduced
    };
    try {
      const saved = await patchSupervisorRunLog(session, row.id, payload);
      const savedRow = {
        ...row,
        ...payload,
        shift: "",
        id: saved?.id || row.id || "",
        submittedBy: supervisorActorName(session),
        submittedAt: saved?.submittedAt || nowIso()
      };
      upsertRowById(line.runRows, savedRow);
      if (superRunLogIdInput.value === row.id) superRunLogIdInput.value = "";
      addAudit(
        line,
        "SUPERVISOR_RUN_COMPLETE",
        `${supervisorActorName(session)} completed run ${payload.product}`
      );
      saveState();
      renderAll();
    } catch (error) {
      alert(`Could not complete production run.\n${error?.message || "Please try again."}`);
    }
  };

  const completeSupervisorDowntimeById = async (downtimeId, { lineId: lineIdOverride } = {}) => {
    const context = resolveSupervisorLineContext(lineIdOverride);
    if (!context) {
      alert("You are not assigned to that line.");
      return;
    }
    const { session, lineId, line } = context;
    const row = (line.downtimeRows || []).find((item) => item.id === downtimeId);
    if (!row) {
      alert("Downtime log could not be found.");
      return;
    }
    if (!isOpenDowntimeRow(row)) {
      alert("This downtime log is already finalised.");
      return;
    }
    const downtimeStart = row.downtimeStart || nowTimeHHMM();
    if (!strictTimeValid(downtimeStart)) {
      alert("Downtime start is invalid. Edit the downtime log and set a valid start time first.");
      return;
    }
    const nowCandidate = nowTimeHHMM();
    const startMins = parseTimeToMinutes(downtimeStart);
    const nowMins = parseTimeToMinutes(nowCandidate);
    const defaultFinish = Number.isFinite(startMins) && Number.isFinite(nowMins) && startMins === nowMins
      ? formatMinutesToHHMM(startMins + 1)
      : nowCandidate;
    const finishPrompt = window.prompt("Downtime finish time (HH:MM)", defaultFinish);
    if (finishPrompt === null) return;
    const downtimeFinish = String(finishPrompt || "").trim();
    if (!strictTimeValid(downtimeFinish)) {
      alert("Finish time must be HH:MM (24h).");
      return;
    }
    if (downtimeFinish === downtimeStart) {
      alert("Finish time must be different from start time.");
      return;
    }
    const parsedReason = parseDowntimeReasonParts(row.reason, row.equipment);
    const reasonCategory = row.reasonCategory || parsedReason.reasonCategory;
    const reasonDetail = row.reasonDetail || parsedReason.reasonDetail;
    const reasonNote = String((row.reasonNote ?? parsedReason.reasonNote) || "").trim();
    if (!reasonCategory || !reasonDetail) {
      alert("Reason group and detail are required before finalising downtime.");
      return;
    }
    const equipment = reasonCategory === "Equipment" ? (row.equipment || reasonDetail) : "";
    if (reasonCategory === "Equipment" && !equipment) {
      alert("Equipment downtime requires an equipment stage.");
      return;
    }
    const shift = row.shift || inferShiftForLog(line, row.date, downtimeStart, appState.supervisorSelectedShift || "Day");
    if (!supervisorCanAccessShift(session, lineId, shift)) {
      alert(`You are not assigned to the ${shift} shift.`);
      return;
    }
    const payload = {
      lineId,
      date: row.date,
      shift,
      downtimeStart,
      downtimeFinish,
      equipment,
      reasonCategory,
      reasonDetail,
      reasonNote,
      reason: buildDowntimeReasonText(line, reasonCategory, reasonDetail, reasonNote)
    };
    try {
      const saved = await patchSupervisorDowntimeLog(session, row.id, payload);
      const savedRow = {
        ...row,
        ...payload,
        shift: "",
        id: saved?.id || row.id || "",
        submittedBy: supervisorActorName(session),
        submittedAt: saved?.submittedAt || nowIso()
      };
      upsertRowById(line.downtimeRows, savedRow);
      if (superDownLogIdInput.value === row.id) superDownLogIdInput.value = "";
      addAudit(
        line,
        "SUPERVISOR_DOWNTIME_COMPLETE",
        `${supervisorActorName(session)} completed downtime on ${stageNameByIdForLine(line, equipment) || reasonCategory}`
      );
      saveState();
      renderAll();
    } catch (error) {
      alert(`Could not complete downtime log.\n${error?.message || "Please try again."}`);
    }
  };

  const submitSupervisorDowntime = async ({ complete = false } = {}) => {
    const session = appState.supervisorSession;
    if (!session) return;
    const lineId = selectedSupervisorLineId();
    const line = appState.lines[lineId];
    if (!session.assignedLineIds.includes(lineId) || !line) {
      alert("You are not assigned to that line.");
      return;
    }

    const date = document.getElementById("superDownDate").value || todayISO();
    const startInput = document.getElementById("superDownStart").value || "";
    const finishInput = document.getElementById("superDownFinish").value || "";
    const reasonCategoryInput = document.getElementById("superDownReasonCategory").value || "";
    const reasonDetailInput = document.getElementById("superDownReasonDetail").value || "";
    const reasonNoteInput = String(document.getElementById("superDownReasonNote").value || "").trim();

    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(date || ""))) {
      alert("Date is invalid.");
      return;
    }

    const selectedDownLogId = String(superDownLogIdInput?.value || "").trim();
    let existing = null;
    if (selectedDownLogId) {
      existing = (line.downtimeRows || []).find((row) => row.id === selectedDownLogId) || null;
      if (!existing) {
        alert("Selected downtime log could not be found. Re-open it from Pending Downtime Logs and try again.");
        return;
      }
      if (existing.date !== date) {
        alert("Date cannot be changed while editing an open downtime log.");
        return;
      }
    }

    const downtimeStart = startInput || existing?.downtimeStart || nowTimeHHMM();
    const downtimeFinish = finishInput || (complete ? nowTimeHHMM() : existing?.downtimeFinish || downtimeStart);
    const backendShift = inferShiftForLog(line, date, downtimeStart, appState.supervisorSelectedShift || "Day");
    const reasonCategory = reasonCategoryInput || existing?.reasonCategory || "";
    const reasonDetail = reasonDetailInput || existing?.reasonDetail || "";
    const reasonNote = reasonNoteInput !== "" ? reasonNoteInput : existing?.reasonNote || "";
    const equipment = reasonCategory === "Equipment" ? (reasonDetail || existing?.equipment || "") : "";
    const reason = buildDowntimeReasonText(line, reasonCategory, reasonDetail, reasonNote);

    if (!strictTimeValid(downtimeStart) || !strictTimeValid(downtimeFinish)) {
      alert("Downtime start and finish must be HH:MM (24h).");
      return;
    }
    if (!reasonCategory) {
      alert("Reason group is required.");
      return;
    }
    if (!reasonDetail) {
      alert("Reason detail is required.");
      return;
    }
    if (reasonCategory === "Equipment" && !equipment) {
      alert("Select an equipment stage.");
      return;
    }
    if (!supervisorCanAccessShift(session, lineId, backendShift)) {
      alert(`You are not assigned to the ${backendShift} shift.`);
      return;
    }

    const payload = {
      lineId,
      date,
      shift: backendShift,
      downtimeStart,
      downtimeFinish,
      equipment,
      reason,
      reasonCategory,
      reasonDetail,
      reasonNote
    };
    try {
      const saved = existing?.id
        ? await patchSupervisorDowntimeLog(session, existing.id, payload)
        : await syncSupervisorDowntimeLog(session, payload);

      const savedRow = {
        ...(existing || {}),
        ...payload,
        shift: "",
        id: saved?.id || existing?.id || "",
        submittedBy: supervisorActorName(session),
        submittedAt: saved?.submittedAt || nowIso()
      };
      upsertRowById(line.downtimeRows, savedRow);
      if (complete) {
        superDownLogIdInput.value = "";
      } else {
        if (superDownLogIdInput) superDownLogIdInput.value = "";
      }
      addAudit(
        line,
        complete ? "SUPERVISOR_DOWNTIME_COMPLETE" : "SUPERVISOR_DOWNTIME_PROGRESS",
        `${supervisorActorName(session)} ${complete ? "completed" : "updated"} downtime on ${stageNameById(equipment)}`
      );
      if (complete) {
        supervisorDownForm.reset();
        document.getElementById("superDownDate").value = date;
        if (supervisorDownReasonCategory) supervisorDownReasonCategory.value = "";
        refreshSupervisorDowntimeDetailOptions(selectedSupervisorLine());
      }
      saveState();
      renderAll();
      if (!complete) {
        supervisorDownForm.reset();
        if (superDownLogIdInput) superDownLogIdInput.value = "";
        if (supervisorDownReasonCategory) supervisorDownReasonCategory.value = "";
        refreshSupervisorDowntimeDetailOptions(selectedSupervisorLine());
      }
    } catch (error) {
      alert(`Could not save downtime log.\n${error?.message || "Please try again."}`);
    }
  };

  superShiftSaveProgressBtn.addEventListener("click", () => submitSupervisorShift({ complete: false }));
  if (superShiftOpenList) {
    superShiftOpenList.addEventListener("click", (event) => {
      const editBtn = event.target.closest("[data-super-shift-edit]");
      if (editBtn) {
        const shiftId = editBtn.getAttribute("data-super-shift-edit");
        if (shiftId) {
          if (superShiftLogIdInput) superShiftLogIdInput.value = shiftId;
          supervisorShiftTileEditId = shiftId;
          appState.supervisorMainTab = "supervisorData";
          appState.supervisorTab = "superShift";
          // Force the tile expansion immediately so Edit behaves deterministically.
          renderHome();
          updateSupervisorProgressButtonLabels();
          loadSupervisorShiftForEdit(shiftId);
        }
        return;
      }
      const breakStartBtn = event.target.closest("[data-super-shift-break-start]");
      if (breakStartBtn) {
        const shiftId = breakStartBtn.getAttribute("data-super-shift-break-start");
        if (shiftId) {
          updateSupervisorProgressButtonLabels();
          submitSupervisorBreakStart({ shiftId });
        }
        return;
      }
      const breakEndBtn = event.target.closest("[data-super-shift-break-end]");
      if (breakEndBtn) {
        const shiftId = breakEndBtn.getAttribute("data-super-shift-break-end");
        const breakId = breakEndBtn.getAttribute("data-super-shift-break-id") || "";
        if (shiftId) {
          updateSupervisorProgressButtonLabels();
          submitSupervisorBreakEnd({ shiftId, breakId });
        }
        return;
      }
      const completeBtn = event.target.closest("[data-super-shift-complete]");
      if (completeBtn) {
        const shiftId = completeBtn.getAttribute("data-super-shift-complete");
        if (shiftId) {
          if (superShiftLogIdInput) superShiftLogIdInput.value = "";
          if (superShiftOpenBreakIdInput) superShiftOpenBreakIdInput.value = "";
          supervisorShiftTileEditId = "";
          updateSupervisorProgressButtonLabels();
          completeSupervisorShiftById(shiftId);
        }
      }
    });
  }
  superRunSaveProgressBtn.addEventListener("click", () => submitSupervisorRun({ complete: false }));
  if (superRunOpenList) {
    superRunOpenList.addEventListener("click", (event) => {
      const editBtn = event.target.closest("[data-super-run-edit]");
      if (editBtn) {
        const runId = editBtn.getAttribute("data-super-run-edit");
        if (runId) loadSupervisorRunForEdit(runId);
        return;
      }
      const completeBtn = event.target.closest("[data-super-run-complete]");
      if (completeBtn) {
        const runId = completeBtn.getAttribute("data-super-run-complete");
        if (runId) completeSupervisorRunById(runId);
      }
    });
  }
  superDownSaveProgressBtn.addEventListener("click", () => submitSupervisorDowntime({ complete: false }));
  if (superDownOpenList) {
    superDownOpenList.addEventListener("click", (event) => {
      const editBtn = event.target.closest("[data-super-down-edit]");
      if (editBtn) {
        const downtimeId = editBtn.getAttribute("data-super-down-edit");
        if (downtimeId) loadSupervisorDowntimeForEdit(downtimeId);
        return;
      }
      const completeBtn = event.target.closest("[data-super-down-complete]");
      if (completeBtn) {
        const downtimeId = completeBtn.getAttribute("data-super-down-complete");
        if (downtimeId) completeSupervisorDowntimeById(downtimeId);
      }
    });
  }

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

  createBuilderLineBtn.addEventListener("click", async () => {
    if (!builderDraft) {
      alert("Builder is not ready. Please close and reopen the builder.");
      return;
    }

    try {
      const lineName = String(builderDraft.lineName || "").trim() || nextAutoLineName();
      const builtStages = makeStagesFromBuilder(builderDraft.stages || []);
      const line = makeDefaultLine(`line-${Date.now()}`, lineName);
      line.stages = builtStages.length ? builtStages : clone(STAGES);
      line.crewsByShift = makeCrewByShiftFromStages(line.stages);
      line.stageSettings = makeSettingsFromStages(line.stages);
      line.selectedStageId = line.stages[0]?.id || "s1";
      addAudit(line, "CREATE_LINE", `Line created with ${line.stages.length} stages`);
      const { lineId, created } = await createLineOnBackend(lineName, line.secretKey, line);
      line.secretKey = created?.line?.secretKey || line.secretKey;
      appState.activeLineId = lineId;
      appState.activeView = "line";
      state = line;
      saveState();
      alert(`Line created.\nDelete key: ${line.secretKey}\nSave this key to delete the line later.`);
      closeBuilderModal();
      await refreshHostedState();
      renderAll();
    } catch (error) {
      console.error(error);
      alert("Could not create line due to an unexpected error. Please try again.");
    }
  });

  cards.addEventListener("click", async (event) => {
    const editBtn = event.target.closest("[data-edit-line]");
    if (editBtn) {
      const id = editBtn.getAttribute("data-edit-line");
      if (!id || !appState.lines[id]) return;
      openEditLineModal(id);
      return;
    }

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
      try {
        const session = await ensureManagerBackendSession();
        const backendLineId = UUID_RE.test(String(id)) ? id : await ensureBackendLineId(id, session);
        if (!backendLineId) throw new Error("Line is not synced to server.");
        await apiRequest(`/api/lines/${backendLineId}`, {
          method: "DELETE",
          token: session.backendToken
        });
        addAudit(line, "DELETE_LINE", "Line deleted");
        delete appState.lines[id];
        appState.supervisors = (appState.supervisors || []).map((sup) => ({
          ...sup,
          assignedLineIds: (sup.assignedLineIds || []).filter((lineId) => lineId !== id),
          assignedLineShifts: Object.fromEntries(
            Object.entries(sup.assignedLineShifts || {}).filter(([lineId]) => lineId !== id)
          )
        }));
        if (appState.supervisorSession) {
          appState.supervisorSession.assignedLineIds = (appState.supervisorSession.assignedLineIds || []).filter((lineId) => lineId !== id);
          appState.supervisorSession.assignedLineShifts = Object.fromEntries(
            Object.entries(appState.supervisorSession.assignedLineShifts || {}).filter(([lineId]) => lineId !== id)
          );
        }
        appState.activeLineId = Object.keys(appState.lines)[0] || "";
        state = appState.lines[appState.activeLineId] || null;
        appState.activeView = "home";
        saveState();
        await refreshHostedState();
        renderAll();
      } catch (error) {
        console.warn("Line delete sync failed:", error);
        alert(`Could not delete line.\n${error?.message || "Please try again."}`);
      }
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

  closeEditLineModalBtn.addEventListener("click", closeEditLineModal);
  editLineModal.addEventListener("click", (event) => {
    if (event.target === editLineModal) closeEditLineModal();
  });

  editLineForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const lineId = editingLineId;
    const line = appState.lines?.[lineId];
    if (!line) return;
    const nextName = String(editLineNameInput.value || "").trim();
    if (nextName.length < 2) {
      alert("Line name must be at least 2 characters.");
      return;
    }
    if (nextName === line.name) {
      closeEditLineModal();
      return;
    }
    try {
      const session = await ensureManagerBackendSession();
      const backendLineId = UUID_RE.test(String(lineId)) ? lineId : await ensureBackendLineId(lineId, session);
      if (!backendLineId) throw new Error("Line is not synced to server.");
      await apiRequest(`/api/lines/${backendLineId}`, {
        method: "PATCH",
        token: session.backendToken,
        body: { name: nextName }
      });
      line.name = nextName;
      addAudit(line, "RENAME_LINE", `Line renamed to ${nextName}`);
      closeEditLineModal();
      saveState();
      await refreshHostedState();
      renderAll();
    } catch (error) {
      console.warn("Line rename sync failed:", error);
      alert(`Could not rename line.\n${error?.message || "Please try again."}`);
    }
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
  const dateInputs = Array.from(document.querySelectorAll("[data-manager-date]"));
  const shiftButtons = Array.from(document.querySelectorAll(".shift-option[data-shift], .shift-option[data-day-shift]"));
  const map = document.getElementById("lineMap");
  const editBtn = document.getElementById("toggleLayoutEdit");
  const addLineBtn = document.getElementById("addFlowLine");
  const addArrowBtn = document.getElementById("addFlowArrow");
  const uploadShapeBtn = document.getElementById("uploadFlowShapeBtn");
  const uploadShapeInput = document.getElementById("uploadFlowShapeInput");
  const prevButtons = [document.getElementById("prevShift"), document.getElementById("dayPrevShift")].filter(Boolean);
  const nextButtons = [document.getElementById("nextShift"), document.getElementById("dayNextShift")].filter(Boolean);
  const lineTrendPrevBtn = document.getElementById("lineTrendPrevPeriod");
  const lineTrendNextBtn = document.getElementById("lineTrendNextPeriod");
  const lineTrendRangeButtons = Array.from(document.querySelectorAll("[data-line-trend-range]"));
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

  dateInputs.forEach((dateInput) => {
    dateInput.value = state.selectedDate;
  });
  setShiftToggleUI();
  setLayoutEditButtonUI();

  dateInputs.forEach((dateInput) => {
    dateInput.addEventListener("change", () => {
      state.selectedDate = dateInput.value || todayISO();
      dateInputs.forEach((input) => {
        input.value = state.selectedDate;
      });
      saveState();
      renderAll();
    });
  });

  shiftButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const nextShift = btn.dataset.shift || btn.dataset.dayShift;
      if (!nextShift || nextShift === state.selectedShift) return;
      state.selectedShift = nextShift;
      setShiftToggleUI();
      saveState();
      renderAll();
    });
  });

  prevButtons.forEach((btn) => btn.addEventListener("click", () => moveShift(-1)));
  nextButtons.forEach((btn) => btn.addEventListener("click", () => moveShift(1)));

  if (lineTrendPrevBtn) lineTrendPrevBtn.addEventListener("click", () => moveLineTrendPeriod(-1));
  if (lineTrendNextBtn) lineTrendNextBtn.addEventListener("click", () => moveLineTrendPeriod(1));
  lineTrendRangeButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const range = String(btn.dataset.lineTrendRange || "").toLowerCase();
      if (!LINE_TREND_RANGES.includes(range) || range === state.lineTrendRange) return;
      state.lineTrendRange = range;
      saveState();
      renderLineTrends();
    });
  });

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
    saveState({ syncModel: true });
    renderVisualiser();
  });

  addArrowBtn.addEventListener("click", () => {
    if (!state.visualEditMode) return;
    state.flowGuides = Array.isArray(state.flowGuides) ? state.flowGuides : [];
    state.flowGuides.push(createGuide("arrow"));
    addAudit(state, "LAYOUT_ADD_ARROW", "Flow arrow guide added");
    saveState({ syncModel: true });
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
      saveState({ syncModel: true });
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
      saveState({ syncModel: true });
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
      saveState({ syncModel: true });
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
    saveState({ syncModel: true });
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

  document.querySelectorAll("[data-manager-date]").forEach((input) => {
    input.value = state.selectedDate;
  });
  saveState();
  renderAll();
}

function moveLineTrendPeriod(direction) {
  const date = parseDateLocal(state.selectedDate || todayISO());
  const range = LINE_TREND_RANGES.includes(String(state.lineTrendRange || "").toLowerCase()) ? String(state.lineTrendRange).toLowerCase() : "day";
  if (range === "quarter") {
    date.setMonth(date.getMonth() + direction * 3);
  } else if (range === "month") {
    date.setMonth(date.getMonth() + direction);
  } else if (range === "week") {
    date.setDate(date.getDate() + direction * 7);
  } else {
    date.setDate(date.getDate() + direction);
  }
  state.selectedDate = formatDateLocal(date);
  document.querySelectorAll("[data-manager-date]").forEach((input) => {
    input.value = state.selectedDate;
  });
  saveState();
  renderAll();
}

function setShiftToggleUI() {
  document.querySelectorAll(".shift-option[data-shift], .shift-option[data-day-shift]").forEach((btn) => {
    const active = (btn.dataset.shift || btn.dataset.dayShift) === state.selectedShift;
    btn.classList.toggle("active", active);
    btn.setAttribute("aria-pressed", String(active));
  });
}

function bindForms() {
  const reasonCategorySelect = document.getElementById("downtimeReasonCategory");
  const reasonDetailSelect = document.getElementById("downtimeReasonDetail");

  const refreshManagerDowntimeDetailOptions = () => {
    const category = String(reasonCategorySelect?.value || "");
    setDowntimeDetailOptions(reasonDetailSelect, state, category, reasonDetailSelect?.value || "");
  };

  if (reasonCategorySelect) {
    reasonCategorySelect.addEventListener("change", refreshManagerDowntimeDetailOptions);
  }
  refreshManagerDowntimeDetailOptions();

  const shiftForm = document.getElementById("shiftForm");
  const runForm = document.getElementById("runForm");
  const runCrewingPatternInput = document.getElementById("runCrewingPattern");
  const runCrewingPatternBtn = document.getElementById("runCrewingPatternBtn");
  const runCrewingPatternSummary = document.getElementById("runCrewingPatternSummary");
  const downtimeForm = document.getElementById("downtimeForm");

  if (runCrewingPatternBtn && runCrewingPatternInput) {
    runCrewingPatternBtn.addEventListener("click", () => {
      if (!state) return;
      const runDate = String(runForm?.querySelector('[name="date"]')?.value || state.selectedDate || todayISO());
      const runStart = String(runForm?.querySelector('[name="productionStartTime"]')?.value || nowTimeHHMM());
      const shiftForPattern = inferShiftForLog(state, runDate, runStart, state.selectedShift || "Day");
      openRunCrewingPatternModal({
        line: state,
        shift: shiftForPattern,
        inputEl: runCrewingPatternInput,
        summaryEl: runCrewingPatternSummary
      });
    });
  }
  setRunCrewingPatternField(
    runCrewingPatternInput,
    runCrewingPatternSummary,
    state,
    state?.selectedShift || "Day",
    runCrewingPatternInput?.value || "",
    { fallbackToIdeal: false }
  );

  shiftForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
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
    if (data.crewOnShift < 0) {
      alert("Crew on shift cannot be negative.");
      return;
    }
    try {
      const payload = {
        lineId: state.id,
        date: data.date,
        shift: data.shift,
        crewOnShift: data.crewOnShift,
        startTime: data.startTime,
        finishTime: data.finishTime
      };
      const saved = await syncManagerShiftLog(payload);
      state.shiftRows.push({
        ...data,
        id: saved?.id || data.id || makeLocalLogId("shift"),
        submittedBy: "manager",
        submittedAt: saved?.submittedAt || nowIso()
      });
      addAudit(state, "MANAGER_SHIFT_LOG", `Manager logged ${data.shift} shift for ${data.date} (crew ${data.crewOnShift})`);
      form.reset();
      saveState();
      renderAll();
    } catch (error) {
      alert(`Could not save shift log.\n${error?.message || "Please try again."}`);
    }
  });

  runForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    const data = Object.fromEntries(new FormData(event.currentTarget).entries());
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(data.date || ""))) {
      alert("Date is required.");
      return;
    }
    if (!String(data.product || "").trim()) {
      alert("Product is required.");
      return;
    }
    if (!strictTimeValid(data.productionStartTime) || !strictTimeValid(data.finishTime)) {
      alert("Production start and finish must be HH:MM (24h).");
      return;
    }
    if (num(data.unitsProduced) < 0) {
      alert("Units produced cannot be negative.");
      return;
    }
    const backendShift = inferShiftForLog(state, data.date, data.productionStartTime, state.selectedShift || "Day");
    const runCrewingPattern = runCrewingPatternFromInput(runCrewingPatternInput, state, backendShift, { fallbackToIdeal: false });
    if (!Object.keys(runCrewingPattern).length) {
      alert("Set crewing pattern for this run before saving.");
      return;
    }
    setRunCrewingPatternField(runCrewingPatternInput, runCrewingPatternSummary, state, backendShift, runCrewingPattern, { fallbackToIdeal: false });
    data.setUpStartTime = "";
    data.unitsProduced = num(data.unitsProduced);
    data.runCrewingPattern = runCrewingPattern;
    try {
      const payload = {
        lineId: state.id,
        date: data.date,
        shift: backendShift,
        product: data.product,
        setUpStartTime: "",
        productionStartTime: data.productionStartTime,
        finishTime: data.finishTime,
        unitsProduced: data.unitsProduced,
        runCrewingPattern
      };
      const saved = await syncManagerRunLog(payload);
      state.runRows.push({
        ...data,
        shift: "",
        runCrewingPattern: normalizeRunCrewingPattern(saved?.runCrewingPattern || data.runCrewingPattern, state, backendShift, { fallbackToIdeal: false }),
        id: saved?.id || data.id || makeLocalLogId("run"),
        submittedBy: "manager",
        submittedAt: saved?.submittedAt || nowIso()
      });
      addAudit(state, "MANAGER_RUN_LOG", `Manager logged run ${data.product} (${data.unitsProduced} units)`);
      form.reset();
      setRunCrewingPatternField(runCrewingPatternInput, runCrewingPatternSummary, state, state.selectedShift || "Day", {}, { fallbackToIdeal: false });
      saveState();
      renderAll();
    } catch (error) {
      alert(`Could not save production run.\n${error?.message || "Please try again."}`);
    }
  });

  downtimeForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    const data = Object.fromEntries(new FormData(event.currentTarget).entries());
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(data.date || ""))) {
      alert("Date is required.");
      return;
    }
    if (!strictTimeValid(data.downtimeStart) || !strictTimeValid(data.downtimeFinish)) {
      alert("Downtime start and finish must be HH:MM (24h).");
      return;
    }
    data.reasonCategory = String(data.reasonCategory || "").trim();
    data.reasonDetail = String(data.reasonDetail || "").trim();
    data.reasonNote = String(data.reasonNote || "").trim();
    if (!data.reasonCategory) {
      alert("Reason group is required.");
      return;
    }
    if (!data.reasonDetail) {
      alert("Reason detail is required.");
      return;
    }
    const backendShift = inferShiftForLog(state, data.date, data.downtimeStart, state.selectedShift || "Day");
    data.equipment = data.reasonCategory === "Equipment" ? data.reasonDetail : "";
    data.reason = buildDowntimeReasonText(state, data.reasonCategory, data.reasonDetail, data.reasonNote);
    try {
      const payload = {
        lineId: state.id,
        date: data.date,
        shift: backendShift,
        downtimeStart: data.downtimeStart,
        downtimeFinish: data.downtimeFinish,
        equipment: data.equipment,
        reason: data.reason || ""
      };
      const saved = await syncManagerDowntimeLog(payload);
      state.downtimeRows.push({
        ...data,
        shift: "",
        id: saved?.id || data.id || makeLocalLogId("down"),
        submittedBy: "manager",
        submittedAt: saved?.submittedAt || nowIso()
      });
      addAudit(state, "MANAGER_DOWNTIME_LOG", `Manager logged downtime on ${stageNameById(data.equipment)}`);
      form.reset();
      refreshManagerDowntimeDetailOptions();
      saveState();
      renderAll();
    } catch (error) {
      alert(`Could not save downtime log.\n${error?.message || "Please try again."}`);
    }
  });

  const inlineValue = (rowNode, field) => String(rowNode?.querySelector(`[data-inline-field="${field}"]`)?.value || "").trim();

  const saveManagerLogInlineEdit = async (type, logId, rowNode) => {
    if (!logId || !rowNode) return;
    if (!isManagerLogInlineEditRow(state.id, type, logId)) return;

    if (type === "shift") {
      const row = (state.shiftRows || []).find((item) => item.id === logId);
      if (!row) return;
      const crewOnShift = Math.max(0, Math.floor(num(inlineValue(rowNode, "crewOnShift"))));
      const startTime = inlineValue(rowNode, "startTime");
      const finishTime = inlineValue(rowNode, "finishTime");
      if (!strictTimeValid(startTime) || !strictTimeValid(finishTime)) {
        alert("Times must be HH:MM (24h).");
        return;
      }
      const payload = { crewOnShift, startTime, finishTime };
      const canPatchServer = UUID_RE.test(String(logId || ""));
      try {
        if (canPatchServer) {
          const saved = await patchManagerShiftLog(logId, payload);
          Object.assign(row, payload, { submittedBy: "manager", submittedAt: saved?.submittedAt || nowIso() });
        } else {
          Object.assign(row, payload, { submittedBy: "manager", submittedAt: nowIso() });
        }
        clearManagerLogInlineEdit();
        addAudit(state, "MANAGER_SHIFT_EDIT", `Manager edited shift row for ${row.date} (${row.shift})`);
        saveState();
        renderAll();
      } catch (error) {
        alert(`Could not update shift row.\n${error?.message || "Please try again."}`);
      }
      return;
    }

    if (type === "run") {
      const row = (state.runRows || []).find((item) => item.id === logId);
      if (!row) return;
      const product = inlineValue(rowNode, "product");
      const productionStartTime = inlineValue(rowNode, "productionStartTime");
      const finishTime = inlineValue(rowNode, "finishTime");
      const unitsProduced = Math.max(0, num(inlineValue(rowNode, "unitsProduced")));
      if (!product) {
        alert("Product is required.");
        return;
      }
      if (!strictTimeValid(productionStartTime) || !strictTimeValid(finishTime)) {
        alert("Times must be HH:MM (24h).");
        return;
      }
      const payload = {
        product,
        setUpStartTime: "",
        productionStartTime,
        finishTime,
        unitsProduced
      };
      const canPatchServer = UUID_RE.test(String(logId || ""));
      try {
        if (canPatchServer) {
          const saved = await patchManagerRunLog(logId, payload);
          Object.assign(row, payload, { shift: "", submittedBy: "manager", submittedAt: saved?.submittedAt || nowIso() });
        } else {
          Object.assign(row, payload, { shift: "", submittedBy: "manager", submittedAt: nowIso() });
        }
        clearManagerLogInlineEdit();
        addAudit(state, "MANAGER_RUN_EDIT", `Manager edited run ${row.product}`);
        saveState();
        renderAll();
      } catch (error) {
        alert(`Could not update run row.\n${error?.message || "Please try again."}`);
      }
      return;
    }

    if (type === "downtime") {
      const row = (state.downtimeRows || []).find((item) => item.id === logId);
      if (!row) return;
      const downtimeStart = inlineValue(rowNode, "downtimeStart");
      const downtimeFinish = inlineValue(rowNode, "downtimeFinish");
      const reason = inlineValue(rowNode, "reason");
      if (!strictTimeValid(downtimeStart) || !strictTimeValid(downtimeFinish)) {
        alert("Times must be HH:MM (24h).");
        return;
      }
      const updatedValues = {
        downtimeStart,
        downtimeFinish,
        equipment: row.equipment || "",
        reason
      };
      const payload = {
        lineId: state.id,
        ...updatedValues
      };
      const canPatchServer = UUID_RE.test(String(logId || ""));
      try {
        if (canPatchServer) {
          const saved = await patchManagerDowntimeLog(logId, payload);
          Object.assign(row, updatedValues, { shift: "", submittedBy: "manager", submittedAt: saved?.submittedAt || nowIso() });
        } else {
          Object.assign(row, updatedValues, { shift: "", submittedBy: "manager", submittedAt: nowIso() });
        }
        clearManagerLogInlineEdit();
        addAudit(
          state,
          "MANAGER_DOWNTIME_EDIT",
          `Manager edited downtime row for ${row.date} (${resolveTimedLogShiftLabel(row, state, "downtimeStart", "downtimeFinish")})`
        );
        saveState();
        renderAll();
      } catch (error) {
        alert(`Could not update downtime row.\n${error?.message || "Please try again."}`);
      }
    }
  };

  const handleLogTableEdit = async (event) => {
    const btn = event.target.closest("[data-log-action]");
    if (!btn) return;
    const type = btn.getAttribute("data-log-type");
    const logId = btn.getAttribute("data-log-id");
    const action = btn.getAttribute("data-log-action");
    if (!type || !logId || !action) return;
    if (action === "save") {
      await saveManagerLogInlineEdit(type, logId, btn.closest("tr"));
      return;
    }
    setManagerLogInlineEdit(state.id, type, logId);
    renderTrackingTables();
  };

  ["shiftTable", "runTable", "downtimeTable"].forEach((tableId) => {
    const table = document.getElementById(tableId);
    if (table) table.addEventListener("click", handleLogTableEdit);
  });
}

function bindDataControls() {
  document.getElementById("loadSampleData").addEventListener("click", () => {
    loadSampleDataIntoLine(state);
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
        Details: `Breaks logged: ${(data.breakRows || []).filter((br) => br.date === row.date && br.shift === row.shift).length}`
      })),
      ...data.runRows.map((row) => ({
        RecordType: "Run",
        Date: row.date,
        Shift: resolveTimedLogShiftLabel(row, state, "productionStartTime", "finishTime"),
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
        Shift: resolveTimedLogShiftLabel(row, state, "downtimeStart", "downtimeFinish"),
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
      state.lineTrendRange = LINE_TREND_RANGES.includes(String(parsed.lineTrendRange || "").toLowerCase())
        ? String(parsed.lineTrendRange).toLowerCase()
        : state.lineTrendRange;
      state.crewsByShift = normalizeCrewByShift(parsed);
      state.stageSettings = normalizeStageSettings(parsed);
      state.shiftRows = parsed.shiftRows || [];
      state.breakRows = parsed.breakRows || [];
      state.runRows = Array.isArray(parsed.runRows)
        ? parsed.runRows.map((row) => ({
            ...row,
            runCrewingPattern: parseRunCrewingPattern(row?.runCrewingPattern),
            shift: ""
          }))
        : [];
      state.downtimeRows = Array.isArray(parsed.downtimeRows)
        ? parsed.downtimeRows.map((row) => ({
            ...row,
            shift: ""
          }))
        : [];
      ensureManagerLogRowIds(state);
      addAudit(state, "IMPORT_JSON", "Line JSON imported");
      saveState();
      renderAll();
    } catch {
      alert("Invalid JSON file.");
    }

    event.target.value = "";
  });

  document.getElementById("clearData").addEventListener("click", async () => {
    const entered = window.prompt(`Enter secret key to clear all data for "${state.name}" (or admin):`) || "";
    if (entered !== "admin" && entered !== state.secretKey) {
      alert("Invalid key/password. Data was not cleared.");
      return;
    }
    if (!window.confirm(`Clear all tracking data for "${state.name}" only?`)) return;
    try {
      const session = await ensureManagerBackendSession();
      const backendLineId = UUID_RE.test(String(state.id)) ? state.id : await ensureBackendLineId(state.id, session);
      if (!backendLineId) throw new Error("Line is not synced to server.");
      await apiRequest(`/api/lines/${backendLineId}/clear-data`, {
        method: "POST",
        token: session.backendToken,
        body: { secretKey: entered }
      });
      clearLineTrackingData(state);
      addAudit(state, "CLEAR_DATA", "All shift/run/downtime rows cleared for this line");
      saveState();
      await refreshHostedState();
      renderAll();
    } catch (error) {
      alert(`Could not clear data.\n${error?.message || "Please try again."}`);
    }
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
      saveState({ syncModel: true });
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
      saveState({ syncModel: true });
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
    .map((row) => {
      const rowClass = String(row?.__rowClass || "").trim();
      const rowClassAttr = rowClass ? ` class="${htmlEscape(rowClass)}"` : "";
      return `<tr${rowClassAttr}>${columns.map((col) => `<td>${formatTableCellValue(row[fieldMap[col]])}</td>`).join("")}</tr>`;
    })
    .join("");

  table.innerHTML = `${header}<tbody>${body}</tbody>`;
}

function selectedRows(rows) {
  return selectedShiftRowsByDate(rows, state.selectedDate, state.selectedShift);
}

function parseTimeToMinutes(value) {
  const fraction = parseTimeToDayFraction(value);
  if (fraction === null) return null;
  return Math.round(fraction * 24 * 60);
}

function formatTime12h(value) {
  const raw = String(value || "").trim();
  const m = raw.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return raw;
  let h = Number(m[1]);
  const mins = m[2];
  const suffix = h >= 12 ? "PM" : "AM";
  h = h % 12;
  if (h === 0) h = 12;
  return `${h}:${mins}${suffix}`;
}

function htmlEscape(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function splitAcrossMidnight(startMins, finishMins) {
  if (startMins === null || finishMins === null || startMins === finishMins) return [];
  if (finishMins > startMins) return [{ start: startMins, end: finishMins }];
  return [
    { start: startMins, end: 24 * 60 },
    { start: 0, end: finishMins }
  ];
}

function stageNameByIdForLine(line, id) {
  const stages = line?.stages?.length ? line.stages : [];
  const idx = stages.findIndex((stage) => stage.id === id);
  if (idx === -1) return id || "";
  return stageDisplayName(stages[idx], idx);
}

function shiftMatchesSelection(rowShift, selectedShift) {
  return !selectedShift || selectedShift === "Full Day" || rowShift === selectedShift;
}

function shortHourLabel(hour) {
  const normalized = ((Number(hour) % 24) + 24) % 24;
  const suffix = normalized >= 12 ? "p" : "a";
  const display = normalized % 12 === 0 ? 12 : normalized % 12;
  return `${display}${suffix}`;
}

function buildDayVisualiserBlocks(data, selectedDate, selectedShift, stageNameResolver) {
  const blocks = { shifts: [], breaks: [], runs: [], downtime: [], source: {} };
  const allShiftRowsForDate = data.shiftRows.filter((row) => row.date === selectedDate);
  const shiftRows = allShiftRowsForDate.filter((row) => shiftMatchesSelection(row.shift, selectedShift));
  const breakRows = (data.breakRows || []).filter((row) => row.date === selectedDate && shiftMatchesSelection(row.shift, selectedShift));
  const runRows = data.runRows.filter((row) => rowMatchesDateShift(row, selectedDate, selectedShift));
  const downRows = data.downtimeRows.filter((row) => rowMatchesDateShift(row, selectedDate, selectedShift));
  const selectedShiftIntervals =
    selectedShift === "Day" || selectedShift === "Night"
      ? shiftIntervalsForDate({ shiftRows: allShiftRowsForDate }, selectedDate)[selectedShift] || []
      : [];
  const clipSegmentsToSelectedShift = (segments) => {
    if (!selectedShiftIntervals.length || isFullDayShift(selectedShift)) return segments;
    return segments.flatMap((segment) =>
      selectedShiftIntervals
        .map((windowSegment) => {
          const start = Math.max(num(segment.start), num(windowSegment.start));
          const end = Math.min(num(segment.end), num(windowSegment.end));
          return end > start ? { start, end } : null;
        })
        .filter(Boolean)
    );
  };

  shiftRows
    .slice()
    .sort((a, b) => String(a.startTime || "").localeCompare(String(b.startTime || "")))
    .forEach((row) => {
      const start = parseTimeToMinutes(row.startTime);
      const end = parseTimeToMinutes(row.finishTime);
      splitAcrossMidnight(start, end).forEach((segment) => {
        blocks.shifts.push({
          ...segment,
          type: "shift",
          title: `${row.shift} Shift`,
          sub: `${formatTime12h(row.startTime)} - ${formatTime12h(row.finishTime)}`
        });
      });
    });

  breakRows
    .slice()
    .sort((a, b) => String(a.breakStart || "").localeCompare(String(b.breakStart || "")))
    .forEach((brk, index) => {
      const breakStart = parseTimeToMinutes(brk.breakStart);
      const breakEnd = parseTimeToMinutes(brk.breakFinish);
      splitAcrossMidnight(breakStart, breakEnd).forEach((segment) => {
        blocks.breaks.push({
          ...segment,
          type: "break",
          title: `Break ${index + 1}`,
          sub: `${formatTime12h(brk.breakStart)} - ${formatTime12h(brk.breakFinish)}`
        });
      });
    });

  runRows
    .slice()
    .sort((a, b) => String(a.productionStartTime || "").localeCompare(String(b.productionStartTime || "")))
    .forEach((row, index) => {
      const prodStart = parseTimeToMinutes(row.productionStartTime);
      const finish = parseTimeToMinutes(row.finishTime);
      const runLabel = `${row.product || "Run"} ${runRows.length > 1 ? `(${index + 1})` : ""}`.trim();
      const unitsLabel = num(row.unitsProduced) > 0 ? ` | ${formatNum(num(row.unitsProduced), 0)} units` : "";

      const segments = clipSegmentsToSelectedShift(splitAcrossMidnight(prodStart, finish));
      segments.forEach((segment) => {
        blocks.runs.push({
          ...segment,
          type: "run-main",
          title: runLabel,
          sub: `${formatTime12h(row.productionStartTime)} - ${formatTime12h(row.finishTime)}${unitsLabel}`
        });
      });
    });

  downRows
    .slice()
    .sort((a, b) => String(a.downtimeStart || "").localeCompare(String(b.downtimeStart || "")))
    .forEach((row) => {
      const start = parseTimeToMinutes(row.downtimeStart);
      const end = parseTimeToMinutes(row.downtimeFinish);
      const equipmentLabel = stageNameResolver(row.equipment || "");
      const reasonText = String(row.reason || "").trim();
      const segments = clipSegmentsToSelectedShift(splitAcrossMidnight(start, end));
      segments.forEach((segment) => {
        blocks.downtime.push({
          ...segment,
          type: "downtime",
          title: equipmentLabel || "Downtime",
          sub: `${formatTime12h(row.downtimeStart)} - ${formatTime12h(row.downtimeFinish)}${reasonText ? ` | ${reasonText}` : ""}`
        });
      });
    });

  blocks.source = { shiftRows, breakRows, runRows, downRows };
  return blocks;
}

function buildDayAtGlance(blocks, stageNameResolver, selectedShift = "Full Day") {
  const mergeIntervals = (rows) =>
    rows
      .map((row) => ({
        start: Math.max(0, num(row.start)),
        end: Math.min(24 * 60, num(row.end))
      }))
      .filter((interval) => interval.end > interval.start)
      .sort((a, b) => a.start - b.start)
      .reduce((merged, current) => {
        const last = merged[merged.length - 1];
        if (!last || current.start > last.end) merged.push({ ...current });
        else if (current.end > last.end) last.end = current.end;
        return merged;
      }, []);
  const overlapMins = (leftIntervals, rightIntervals) => {
    let i = 0;
    let j = 0;
    let total = 0;
    while (i < leftIntervals.length && j < rightIntervals.length) {
      const left = leftIntervals[i];
      const right = rightIntervals[j];
      const start = Math.max(left.start, right.start);
      const end = Math.min(left.end, right.end);
      if (end > start) total += end - start;
      if (left.end <= right.end) i += 1;
      else j += 1;
    }
    return total;
  };
  const minutesFromBlocks = (rows) => rows.reduce((sum, row) => sum + Math.max(0, num(row.end) - num(row.start)), 0);
  const runMins = minutesFromBlocks(blocks.runs);
  const breakMins = minutesFromBlocks(blocks.breaks);
  const downtimeMins = minutesFromBlocks(blocks.downtime);
  const trackedTotal = runMins + breakMins + downtimeMins;
  const shiftIntervals = mergeIntervals(blocks.shifts);
  const activityIntervals = mergeIntervals([...blocks.runs, ...blocks.breaks, ...blocks.downtime]);
  const shiftMins = shiftIntervals.reduce((sum, interval) => sum + (interval.end - interval.start), 0);
  const coveredInShiftMins = overlapMins(shiftIntervals, activityIntervals);
  const unassignedMins = Math.max(0, shiftMins - coveredInShiftMins);
  const allocationTotal = trackedTotal + unassignedMins;

  const allocations = [
    { key: "run", label: "Run", minutes: runMins, className: "run-main" },
    { key: "break", label: "Break", minutes: breakMins, className: "break" },
    { key: "downtime", label: "Downtime", minutes: downtimeMins, className: "downtime" },
    { key: "unassigned", label: "Unassigned", minutes: unassignedMins, className: "unassigned" }
  ].map((entry) => ({
    ...entry,
    pct: allocationTotal > 0 ? (entry.minutes / allocationTotal) * 100 : 0
  }));

  const downtimeByCause = new Map();
  (blocks.source.downRows || []).forEach((row) => {
    const equipmentLabel = stageNameResolver(row.equipment || "");
    const reasonText = String(row.reason || "").trim();
    const label = reasonText ? `${reasonText}${equipmentLabel ? ` (${equipmentLabel})` : ""}` : equipmentLabel || "Unspecified cause";
    const loggedMins = num(row.downtimeMins);
    const weightedLogged = loggedMins * timedLogShiftWeight(row, selectedShift);
    const rawMins = loggedMins > 0 ? weightedLogged : diffMinutes(row.downtimeStart, row.downtimeFinish) * timedLogShiftWeight(row, selectedShift);
    const mins = Math.max(0, rawMins);
    downtimeByCause.set(label, (downtimeByCause.get(label) || 0) + Math.max(0, mins));
  });
  const topDowntime = Array.from(downtimeByCause.entries())
    .map(([label, minutes]) => ({ label, minutes }))
    .sort((a, b) => b.minutes - a.minutes)
    .slice(0, 3);

  return {
    trackedTotal,
    allocations,
    topDowntime,
    runCount: (blocks.source.runRows || []).filter((row) => timedLogShiftWeight(row, selectedShift) > 0).length,
    downtimeCount: (blocks.source.downRows || []).filter((row) => timedLogShiftWeight(row, selectedShift) > 0).length
  };
}

function renderDayVisualiserTo(rootId, data, selectedDate, selectedShift, stageNameResolver) {
  const root = document.getElementById(rootId);
  if (!root) return;
  const blocks = buildDayVisualiserBlocks(data, selectedDate, selectedShift, stageNameResolver);
  const all = [...blocks.shifts, ...blocks.breaks, ...blocks.runs, ...blocks.downtime].filter(
    (item) => Number.isFinite(item.start) && Number.isFinite(item.end)
  );
  if (!all.length) {
    root.innerHTML = `<div class="day-viz-empty">No shift/run/downtime records for ${selectedDate}.</div>`;
    return;
  }

  const dayStartHour = 5;
  const startMins = dayStartHour * 60;
  const endMins = startMins + 24 * 60;
  const rangeMins = endMins - startMins;
  const hourMarks = Array.from({ length: 25 }, (_, offset) => dayStartHour + offset);
  const glance = buildDayAtGlance(blocks, stageNameResolver, selectedShift);
  const selectedShiftLabel = selectedShift === "Day" || selectedShift === "Night" ? selectedShift : "All shifts";

  const toWindowSegment = (start, end) => {
    if (!Number.isFinite(start) || !Number.isFinite(end) || end <= start) return null;
    let mappedStart = start;
    let mappedEnd = end;
    if (mappedEnd <= startMins) {
      mappedStart += 24 * 60;
      mappedEnd += 24 * 60;
    }
    const clippedStart = Math.max(startMins, mappedStart);
    const clippedEnd = Math.min(endMins, mappedEnd);
    if (clippedEnd <= clippedStart) return null;
    return { start: clippedStart, end: clippedEnd };
  };

  const visibleIntervals = (items) =>
    items
      .map((item) => toWindowSegment(num(item.start), num(item.end)))
      .filter(Boolean)
      .sort((a, b) => a.start - b.start)
      .reduce((merged, current) => {
        const last = merged[merged.length - 1];
        if (!last || current.start > last.end) {
          merged.push({ ...current });
        } else if (current.end > last.end) {
          last.end = current.end;
        }
        return merged;
      }, []);

  const intervalsOverlapMinutes = (leftIntervals, rightIntervals) => {
    let i = 0;
    let j = 0;
    let total = 0;
    while (i < leftIntervals.length && j < rightIntervals.length) {
      const a = leftIntervals[i];
      const b = rightIntervals[j];
      const start = Math.max(a.start, b.start);
      const end = Math.min(a.end, b.end);
      if (end > start) total += end - start;
      if (a.end <= b.end) i += 1;
      else j += 1;
    }
    return total;
  };

  const laneMinutes = (items) =>
    items.reduce((sum, item) => {
      const visible = toWindowSegment(num(item.start), num(item.end));
      return sum + (visible ? visible.end - visible.start : 0);
    }, 0);
  const productionUnits = (blocks.source.runRows || []).reduce(
    (sum, row) => sum + Math.max(0, num(row.unitsProduced) * timedLogShiftWeight(row, selectedShift)),
    0
  );
  const productionMinutes = laneMinutes(blocks.runs);
  const productionIntervals = visibleIntervals(blocks.runs);
  const downtimeIntervals = visibleIntervals(blocks.downtime);
  const overlappingDowntimeMins = intervalsOverlapMinutes(productionIntervals, downtimeIntervals);
  const netProductionMinutes = Math.max(0, productionMinutes - overlappingDowntimeMins);
  const productionRate = netProductionMinutes > 0 ? productionUnits / netProductionMinutes : 0;
  const productionTrafficForRate = (rate, netMins) =>
    netMins <= 0
      ? { className: "red", label: "Red", detail: "No production logged" }
      : rate >= 10
        ? { className: "green", label: "Green", detail: "On target" }
        : rate >= 7
          ? { className: "amber", label: "Amber", detail: "Monitor" }
          : { className: "red", label: "Red", detail: "Below target" };
  const laneTotals = {
    shift: `${formatNum(laneMinutes(blocks.shifts), 0)} min`,
    break: `${formatNum(laneMinutes(blocks.breaks), 0)} min`,
    production: `${formatNum(productionRate, 2)} units / min`,
    downtime: `${formatNum(laneMinutes(blocks.downtime), 1)} min`
  };
  const runTileShiftWindows =
    selectedShift === "Day" || selectedShift === "Night"
      ? shiftIntervalsForDate({ shiftRows: (data.shiftRows || []).filter((row) => row.date === selectedDate) }, selectedDate)[selectedShift] || []
      : [];
  const clipRunTileSegmentsToShift = (segments) => {
    if (isFullDayShift(selectedShift) || !runTileShiftWindows.length) return segments;
    return segments.flatMap((segment) =>
      runTileShiftWindows
        .map((windowSegment) => {
          const start = Math.max(num(segment.start), num(windowSegment.start));
          const end = Math.min(num(segment.end), num(windowSegment.end));
          return end > start ? { start, end } : null;
        })
        .filter(Boolean)
    );
  };
  const productionRunTiles = (blocks.source.runRows || [])
    .slice()
    .sort((a, b) => String(a.productionStartTime || "").localeCompare(String(b.productionStartTime || "")))
    .map((row, index) => {
      const productName = String(row.product || "").trim() || `Run ${index + 1}`;
      const units = Math.max(0, num(row.unitsProduced) * timedLogShiftWeight(row, selectedShift));
      const runIntervals = clipRunTileSegmentsToShift(splitAcrossMidnight(parseTimeToMinutes(row.productionStartTime), parseTimeToMinutes(row.finishTime)))
        .map((segment) => toWindowSegment(segment.start, segment.end))
        .filter(Boolean);
      const grossMins = runIntervals.reduce((sum, interval) => sum + (interval.end - interval.start), 0);
      const downInRunMins = intervalsOverlapMinutes(runIntervals, downtimeIntervals);
      const netRunMins = Math.max(0, grossMins - downInRunMins);
      const traysPerMin = netRunMins > 0 ? units / netRunMins : 0;
      const traffic = productionTrafficForRate(traysPerMin, netRunMins);
      return `
        <article class="day-glance-run-tile">
          <div class="day-glance-run-main">
            <span class="day-glance-light ${traffic.className}" aria-hidden="true"></span>
            <div class="day-glance-run-copy">
              <h5 class="day-glance-run-title">${htmlEscape(productName)}</h5>
              <span class="day-glance-run-rate">${formatNum(traysPerMin, 2)} trays / min</span>
            </div>
          </div>
          <div class="day-glance-run-callouts">
            <div class="day-glance-run-callout"><span>Trays</span><strong>${formatNum(units, 0)}</strong></div>
            <div class="day-glance-run-callout"><span>Down</span><strong>${formatNum(downInRunMins, 1)} min</strong></div>
          </div>
        </article>
      `;
    })
    .join("");

  const axisLines = hourMarks
    .map((hour) => {
      const left = ((hour - dayStartHour) / 24) * 100;
      return `<div class="day-viz-axis-line ${hour % 2 === 0 ? "major" : ""}" style="left:${left}%"></div>`;
    })
    .join("");
  const axisTicks = hourMarks
    .map((hour) => {
      const left = ((hour - dayStartHour) / 24) * 100;
      const major = hour % 2 === 0;
      return `<div class="day-viz-axis-tick ${major ? "major" : ""}" style="left:${left}%;">${major ? `<span>${shortHourLabel(hour)}</span>` : ""}</div>`;
    })
    .join("");

  const renderLane = (label, totalLabel, items) => {
    const laneLines = hourMarks
      .map((hour) => {
        const left = ((hour - dayStartHour) / 24) * 100;
        return `<div class="day-viz-hour-line ${hour % 2 === 0 ? "major" : ""}" style="left:${left}%"></div>`;
      })
      .join("");

    const cards = items
      .map((item) => {
        const visible = toWindowSegment(num(item.start), num(item.end));
        if (!visible) return "";
        const left = ((visible.start - startMins) / rangeMins) * 100;
        const width = ((visible.end - visible.start) / rangeMins) * 100;
        const clampedLeft = Math.max(0, Math.min(100, left));
        const clampedWidth = Math.max(0.65, Math.min(100 - clampedLeft, width));
        if (clampedWidth <= 0) return "";
        const tooltip = `${item.title}${item.sub ? ` | ${item.sub}` : ""}`;
        return `
          <article class="day-viz-block ${item.type}" style="left:${clampedLeft}%;width:${clampedWidth}%;" title="${htmlEscape(tooltip)}">
            <span class="day-viz-title">${htmlEscape(item.title)}</span>
            <span class="day-viz-sub">${htmlEscape(item.sub)}</span>
          </article>
        `;
      })
      .join("");

    return `
      <div class="day-viz-swimlane">
        <div class="day-viz-lane-label">
          <span class="day-viz-lane-title">${htmlEscape(label)}</span>
          <span class="day-viz-lane-total">${htmlEscape(totalLabel)}</span>
        </div>
        <div class="day-viz-lane-track">
          ${laneLines}
          ${cards || `<div class="day-viz-lane-empty">No records</div>`}
        </div>
      </div>
    `;
  };

  const allocationSegments = glance.allocations
    .filter((segment) => segment.pct > 0)
    .map((segment) => `<span class="day-glance-seg ${segment.className}" style="width:${segment.pct}%"></span>`)
    .join("");
  const allocationLegend = glance.allocations
    .map(
      (segment) => `
        <div class="day-glance-legend-item">
          <span class="day-glance-dot ${segment.className}"></span>
          <span>${htmlEscape(segment.label)}</span>
          <strong>${formatNum(segment.minutes, 0)} min</strong>
        </div>
      `
    )
    .join("");
  const topDowntime =
    glance.topDowntime.length > 0
      ? `<ul class="day-glance-list">
          ${glance.topDowntime
            .map(
              (entry) => `
                <li>
                  <span>${htmlEscape(entry.label)}</span>
                  <strong>${formatNum(entry.minutes, 1)} min</strong>
                </li>
              `
            )
            .join("")}
        </ul>`
      : `<p class="day-glance-empty">No downtime events in this selection.</p>`;

  root.innerHTML = `
    <div class="day-viz-shell">
      <div class="day-viz-inner">
        <div class="day-viz-axis-row">
          <div class="day-viz-axis-label">Time</div>
          <div class="day-viz-axis-track">
            ${axisLines}
            ${axisTicks}
          </div>
        </div>
        ${renderLane("Shift", laneTotals.shift, blocks.shifts)}
        ${renderLane("Break", laneTotals.break, blocks.breaks)}
        ${renderLane("Production", laneTotals.production, blocks.runs)}
        ${renderLane("Downtime", laneTotals.downtime, blocks.downtime)}
      </div>
    </div>
    <section class="day-glance">
      <div class="day-glance-head">
        <h3>Day at a glance</h3>
        <p>${selectedDate} | ${selectedShiftLabel} | Runs: ${formatNum(glance.runCount, 0)} | Downtime events: ${formatNum(glance.downtimeCount, 0)}</p>
      </div>
      <div class="day-glance-grid">
        <article class="day-glance-card">
          <h4>Time Allocation</h4>
          <p class="day-glance-meta">${formatNum(glance.trackedTotal, 0)} tracked minutes</p>
          <div class="day-glance-stack">
            ${allocationSegments || `<span class="day-glance-seg fallback"></span>`}
          </div>
          <div class="day-glance-legend">${allocationLegend}</div>
        </article>
        <article class="day-glance-card">
          <h4>Production Runs</h4>
          <p class="day-glance-meta">Per-run net trays/min. Downtime overlapping each run is excluded.</p>
          <div class="day-glance-run-list">
            ${productionRunTiles || `<p class="day-glance-empty">No production runs in this selection.</p>`}
          </div>
        </article>
        <article class="day-glance-card">
          <h4>Top Downtime Causes</h4>
          <p class="day-glance-meta">By total logged minutes.</p>
          ${topDowntime}
        </article>
      </div>
    </section>
  `;
}

function renderDayVisualiser() {
  renderDayVisualiserTo("dayVisualiserCanvas", derivedData(), state.selectedDate, state.selectedShift, stageNameById);
}

function lineTrendRangeKey() {
  const value = String(state?.lineTrendRange || "").toLowerCase();
  return LINE_TREND_RANGES.includes(value) ? value : "day";
}

function formatIsoDateLabel(isoDate, options = { month: "short", day: "numeric" }) {
  const date = parseDateLocal(isoDate || todayISO());
  return date.toLocaleDateString(undefined, options);
}

function startOfIsoWeek(date) {
  const base = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const day = (base.getDay() + 6) % 7;
  base.setDate(base.getDate() - day);
  return base;
}

function datesBetweenInclusive(startDate, endDate) {
  const dates = [];
  const cursor = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  const end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  while (cursor <= end) {
    dates.push(formatDateLocal(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }
  return dates;
}

function buildLineTrendBuckets(anchorIsoDate, range = "day") {
  const safeRange = LINE_TREND_RANGES.includes(String(range || "").toLowerCase()) ? String(range).toLowerCase() : "day";
  const anchor = parseDateLocal(anchorIsoDate || todayISO());
  const buckets = [];
  const pushBucket = (start, end, label) => {
    const startIso = formatDateLocal(start);
    const endIso = formatDateLocal(end);
    buckets.push({
      key: `${startIso}__${endIso}`,
      label,
      startIso,
      endIso,
      dates: datesBetweenInclusive(start, end)
    });
  };

  if (safeRange === "day") {
    for (let offset = 13; offset >= 0; offset -= 1) {
      const current = new Date(anchor.getFullYear(), anchor.getMonth(), anchor.getDate() - offset);
      pushBucket(current, current, formatIsoDateLabel(formatDateLocal(current), { month: "short", day: "numeric" }));
    }
    return buckets;
  }

  if (safeRange === "week") {
    const anchorWeek = startOfIsoWeek(anchor);
    for (let offset = 11; offset >= 0; offset -= 1) {
      const start = new Date(anchorWeek.getFullYear(), anchorWeek.getMonth(), anchorWeek.getDate() - offset * 7);
      const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6);
      pushBucket(start, end, `Wk ${formatIsoDateLabel(formatDateLocal(start), { month: "short", day: "numeric" })}`);
    }
    return buckets;
  }

  if (safeRange === "month") {
    const anchorMonth = new Date(anchor.getFullYear(), anchor.getMonth(), 1);
    for (let offset = 11; offset >= 0; offset -= 1) {
      const start = new Date(anchorMonth.getFullYear(), anchorMonth.getMonth() - offset, 1);
      const end = new Date(start.getFullYear(), start.getMonth() + 1, 0);
      pushBucket(start, end, start.toLocaleDateString(undefined, { month: "short", year: "2-digit" }));
    }
    return buckets;
  }

  const anchorQuarter = new Date(anchor.getFullYear(), Math.floor(anchor.getMonth() / 3) * 3, 1);
  for (let offset = 7; offset >= 0; offset -= 1) {
    const start = new Date(anchorQuarter.getFullYear(), anchorQuarter.getMonth() - offset * 3, 1);
    const end = new Date(start.getFullYear(), start.getMonth() + 3, 0);
    const qNum = Math.floor(start.getMonth() / 3) + 1;
    pushBucket(start, end, `Q${qNum} '${String(start.getFullYear()).slice(-2)}`);
  }
  return buckets;
}

function aggregateLineTrendPoints(line, buckets, shift) {
  return buckets.map((bucket) => {
    let units = 0;
    let downtime = 0;
    let utilSum = 0;
    let utilCount = 0;
    let runRateSum = 0;
    let runRateCount = 0;

    bucket.dates.forEach((date) => {
      const daily = computeLineMetrics(line, date, shift);
      units += Math.max(0, num(daily.units));
      downtime += Math.max(0, num(daily.totalDowntime));
      const util = Math.max(0, num(daily.lineUtil));
      if (Number.isFinite(util)) {
        utilSum += util;
        utilCount += 1;
      }
      const runRate = Math.max(0, num(daily.netRunRate));
      if (runRate > 0) {
        runRateSum += runRate;
        runRateCount += 1;
      }
    });

    return {
      ...bucket,
      units,
      downtime,
      lineUtil: utilCount > 0 ? utilSum / utilCount : 0,
      netRunRate: runRateCount > 0 ? runRateSum / runRateCount : 0
    };
  });
}

function lineTrendTopDowntimeReasons(line, buckets, shift, maxItems = 5) {
  const dateSet = new Set(buckets.flatMap((bucket) => bucket.dates));
  const totals = new Map();
  const downRows = derivedDataForLine(line || {}).downtimeRows || [];
  downRows.forEach((row) => {
    if (!dateSet.has(row.date)) return;
    if (!rowMatchesDateShift(row, row.date, shift, { line, startField: "downtimeStart", finishField: "downtimeFinish" })) return;
    const weightedMins = Math.max(0, num(row.downtimeMins) * timedLogShiftWeight(row, shift));
    if (weightedMins <= 0) return;
    const parsed = parseDowntimeReasonParts(row.reason, row.equipment);
    const category = row.reasonCategory || parsed.reasonCategory || "";
    const detail = row.reasonDetail || parsed.reasonDetail || "";
    const detailLabel = detail ? downtimeDetailLabel(line, category, detail) : "";
    const label = category ? `${category}${detailLabel ? ` > ${detailLabel}` : ""}` : String(row.reason || "").trim() || "Unspecified";
    totals.set(label, (totals.get(label) || 0) + weightedMins);
  });
  return Array.from(totals.entries())
    .map(([label, minutes]) => ({ label, minutes }))
    .sort((a, b) => b.minutes - a.minutes)
    .slice(0, maxItems);
}

function renderLineTrendUnitsChart(points) {
  const root = document.getElementById("lineTrendUnitsChart");
  if (!root) return;
  if (!points.length) {
    root.innerHTML = `<div class="empty-chart">No trend data available.</div>`;
    return;
  }

  const width = 1280;
  const height = 430;
  const pad = { top: 34, right: 110, bottom: 78, left: 76 };
  const chartW = width - pad.left - pad.right;
  const chartH = height - pad.top - pad.bottom;
  const maxUnits = Math.max(1, ...points.map((point) => point.units));
  const avgUnits = Math.round(points.reduce((sum, point) => sum + point.units, 0) / points.length);
  const stepX = points.length > 1 ? chartW / (points.length - 1) : chartW;
  const x = (index) => pad.left + stepX * index;
  const yUnits = (value) => pad.top + chartH - (Math.max(0, value) / maxUnits) * chartH;

  const barWidth = Math.max(14, Math.min(36, stepX * 0.58));
  const bars = points
    .map((point, index) => {
      const y = yUnits(point.units);
      const h = pad.top + chartH - y;
      return `<rect x="${x(index) - barWidth / 2}" y="${y}" width="${barWidth}" height="${Math.max(1, h)}" rx="3" class="bar-units"><title>${point.label}: ${formatNum(point.units, 0)} units</title></rect>`;
    })
    .join("");

  const averageY = yUnits(avgUnits);
  const avgUnitsLabel = `Avg ${formatNum(avgUnits, 0)} units`;
  const avgUnitsCalloutWidth = Math.max(128, avgUnitsLabel.length * 8 + 22);
  const avgUnitsCalloutHeight = 28;
  const avgUnitsCalloutX = pad.left + chartW - avgUnitsCalloutWidth - 10;
  const avgUnitsCalloutY = Math.max(
    pad.top + 6,
    Math.min(averageY - avgUnitsCalloutHeight / 2, pad.top + chartH - avgUnitsCalloutHeight - 6)
  );

  const ticks = [0, 0.25, 0.5, 0.75, 1]
    .map((fraction) => Math.round(maxUnits * fraction))
    .map((tick) => {
      const y = yUnits(tick);
      return `<g><line x1="${pad.left}" y1="${y}" x2="${pad.left + chartW}" y2="${y}" class="grid"/><text x="${pad.left - 12}" y="${y + 5}" text-anchor="end" class="axis">${tick}</text></g>`;
    })
    .join("");

  const labelSkip = points.length > 14 ? Math.ceil(points.length / 10) : 1;
  const labels = points
    .map((point, index) => {
      const label = index % labelSkip === 0 || index === points.length - 1 ? point.label : "";
      return `<text x="${x(index)}" y="${height - 24}" text-anchor="middle" class="axis">${htmlEscape(label)}</text>`;
    })
    .join("");

  root.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" class="trend-svg line-trend-svg" role="img" aria-label="Line units trend with average line">
      ${ticks}
      <line x1="${pad.left}" y1="${pad.top + chartH}" x2="${pad.left + chartW}" y2="${pad.top + chartH}" class="axis-line" />
      <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${pad.top + chartH}" class="axis-line" />
      ${bars}
      <line x1="${pad.left}" y1="${averageY}" x2="${pad.left + chartW}" y2="${averageY}" class="line-rate"><title>Average units: ${formatNum(
        avgUnits,
        0
      )}</title></line>
      <g transform="translate(${avgUnitsCalloutX},${avgUnitsCalloutY})">
        <rect width="${avgUnitsCalloutWidth}" height="${avgUnitsCalloutHeight}" rx="${avgUnitsCalloutHeight / 2}" class="trend-avg-callout-bg units" />
        <text x="${avgUnitsCalloutWidth / 2}" y="${avgUnitsCalloutHeight / 2}" text-anchor="middle" dominant-baseline="middle" class="trend-avg-callout-text units">${htmlEscape(
          avgUnitsLabel
        )}</text>
      </g>
      ${labels}
      <text x="${pad.left}" y="${20}" class="legend units">Units (bars)</text>
      <text x="${pad.left + 152}" y="${20}" class="legend rate">Average units (line)</text>
      <text x="${width - 8}" y="${pad.top + 12}" text-anchor="end" class="axis">Average ${formatNum(avgUnits, 0)} units</text>
    </svg>
  `;
}

function renderLineTrendUtilDownChart(points) {
  const root = document.getElementById("lineTrendUtilDownChart");
  if (!root) return;
  if (!points.length) {
    root.innerHTML = `<div class="empty-chart">No trend data available.</div>`;
    return;
  }

  const width = 1280;
  const height = 430;
  const pad = { top: 34, right: 110, bottom: 78, left: 76 };
  const chartW = width - pad.left - pad.right;
  const chartH = height - pad.top - pad.bottom;
  const maxDown = Math.max(1, ...points.map((point) => point.downtime));
  const avgDown = Math.round(points.reduce((sum, point) => sum + point.downtime, 0) / points.length);
  const maxUtil = Math.max(100, ...points.map((point) => point.lineUtil));
  const stepX = points.length > 1 ? chartW / (points.length - 1) : chartW;
  const x = (index) => pad.left + stepX * index;
  const yDown = (value) => pad.top + chartH - (Math.max(0, value) / maxDown) * chartH;
  const yUtil = (value) => pad.top + chartH - (Math.max(0, value) / maxUtil) * chartH;
  const avgDownY = yDown(avgDown);
  const avgDownLabel = `Avg ${formatNum(avgDown, 0)} min`;
  const avgDownCalloutWidth = Math.max(120, avgDownLabel.length * 8 + 22);
  const avgDownCalloutHeight = 28;
  const avgDownCalloutX = pad.left + chartW - avgDownCalloutWidth - 10;
  const avgDownCalloutY = Math.max(
    pad.top + 6,
    Math.min(avgDownY - avgDownCalloutHeight / 2, pad.top + chartH - avgDownCalloutHeight - 6)
  );

  const barWidth = Math.max(14, Math.min(34, stepX * 0.54));
  const bars = points
    .map((point, index) => {
      const y = yDown(point.downtime);
      const h = pad.top + chartH - y;
      return `<rect x="${x(index) - barWidth / 2}" y="${y}" width="${barWidth}" height="${Math.max(1, h)}" rx="3" class="bar-down"><title>${point.label}: ${formatNum(point.downtime, 1)} min downtime</title></rect>`;
    })
    .join("");

  const utilPath = points.map((point, index) => `${index === 0 ? "M" : "L"}${x(index)},${yUtil(point.lineUtil)}`).join(" ");
  const utilDots = points
    .map(
      (point, index) => `<circle cx="${x(index)}" cy="${yUtil(point.lineUtil)}" r="4.4" class="dot-util"><title>${point.label}: ${formatNum(
        point.lineUtil,
        1
      )}% utilisation</title></circle>`
    )
    .join("");

  const ticks = [0, 0.25, 0.5, 0.75, 1]
    .map((fraction) => Math.round(maxUtil * fraction))
    .map((tick) => {
      const y = yUtil(tick);
      return `<g><line x1="${pad.left}" y1="${y}" x2="${pad.left + chartW}" y2="${y}" class="grid"/><text x="${pad.left - 12}" y="${y + 5}" text-anchor="end" class="axis">${tick}%</text></g>`;
    })
    .join("");

  const labelSkip = points.length > 14 ? Math.ceil(points.length / 10) : 1;
  const labels = points
    .map((point, index) => {
      const label = index % labelSkip === 0 || index === points.length - 1 ? point.label : "";
      return `<text x="${x(index)}" y="${height - 24}" text-anchor="middle" class="axis">${htmlEscape(label)}</text>`;
    })
    .join("");

  root.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" class="trend-svg line-trend-svg" role="img" aria-label="Line utilisation and downtime trend">
      ${ticks}
      <line x1="${pad.left}" y1="${pad.top + chartH}" x2="${pad.left + chartW}" y2="${pad.top + chartH}" class="axis-line" />
      <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${pad.top + chartH}" class="axis-line" />
      ${bars}
      <path d="${utilPath}" class="line-util" />
      <line x1="${pad.left}" y1="${avgDownY}" x2="${pad.left + chartW}" y2="${avgDownY}" class="line-down-avg"><title>Average downtime: ${formatNum(
        avgDown,
        0
      )} min</title></line>
      <g transform="translate(${avgDownCalloutX},${avgDownCalloutY})">
        <rect width="${avgDownCalloutWidth}" height="${avgDownCalloutHeight}" rx="${avgDownCalloutHeight / 2}" class="trend-avg-callout-bg down" />
        <text x="${avgDownCalloutWidth / 2}" y="${avgDownCalloutHeight / 2}" text-anchor="middle" dominant-baseline="middle" class="trend-avg-callout-text down">${htmlEscape(
          avgDownLabel
        )}</text>
      </g>
      ${utilDots}
      ${labels}
      <text x="${pad.left}" y="${20}" class="legend util">Utilisation % (line)</text>
      <text x="${pad.left + 178}" y="${20}" class="legend down">Downtime min (bars)</text>
      <text x="${pad.left + 364}" y="${20}" class="legend down-avg">Average downtime (line)</text>
      <text x="${width - 8}" y="${pad.top + 12}" text-anchor="end" class="axis">Max downtime ${formatNum(maxDown, 1)} min</text>
    </svg>
  `;
}

function renderLineTrendReasons(reasons) {
  const root = document.getElementById("lineTrendTopReasons");
  if (!root) return;
  if (!Array.isArray(reasons) || !reasons.length) {
    root.innerHTML = `<p class="muted line-trend-empty">No downtime reasons recorded in this time window.</p>`;
    return;
  }
  root.innerHTML = reasons
    .map(
      (entry) => `
        <article class="line-trend-reason-item">
          <span>${htmlEscape(entry.label)}</span>
          <strong>${formatNum(entry.minutes, 1)} min</strong>
        </article>
      `
    )
    .join("");
}

function renderLineTrends() {
  const range = lineTrendRangeKey();
  if (state.lineTrendRange !== range) state.lineTrendRange = range;
  document.querySelectorAll("[data-line-trend-range]").forEach((btn) => {
    const active = String(btn.dataset.lineTrendRange || "").toLowerCase() === range;
    btn.classList.toggle("active", active);
    btn.setAttribute("aria-pressed", String(active));
  });

  const rangeMeta = document.getElementById("lineTrendRangeMeta");
  const setText = (id, value) => {
    const node = document.getElementById(id);
    if (node) node.textContent = value;
  };

  const buckets = buildLineTrendBuckets(state.selectedDate || todayISO(), range);
  const points = aggregateLineTrendPoints(state, buckets, state.selectedShift || "Day");
  const current = points[points.length - 1] || { units: 0, downtime: 0, lineUtil: 0, netRunRate: 0 };
  const startIso = points[0]?.startIso || state.selectedDate || todayISO();
  const endIso = points[points.length - 1]?.endIso || state.selectedDate || todayISO();
  const rangeLabel = range === "day" ? "Daily" : range === "week" ? "Weekly" : range === "month" ? "Monthly" : "Quarterly";
  if (rangeMeta) {
    rangeMeta.textContent = `${rangeLabel} view | ${state.selectedShift} shift | ${formatIsoDateLabel(startIso)} to ${formatIsoDateLabel(endIso)}`;
  }

  setText("lineTrendKpiUnits", formatNum(current.units, 0));
  setText("lineTrendKpiDowntime", `${formatNum(current.downtime, 1)} min`);
  setText("lineTrendKpiUtilisation", `${formatNum(current.lineUtil, 1)}%`);
  setText("lineTrendKpiRunRate", `${formatNum(current.netRunRate, 2)} u/min`);

  renderLineTrendUnitsChart(points);
  renderLineTrendUtilDownChart(points);
  renderLineTrendReasons(lineTrendTopDowntimeReasons(state, buckets, state.selectedShift || "Day"));
}

function renderSupervisorDayVisualiser(line, selectedDate) {
  if (!line) {
    const root = document.getElementById("supervisorDayVisualiserCanvas");
    if (root) root.innerHTML = `<div class="day-viz-empty">No assigned lines selected.</div>`;
    return;
  }
  renderDayVisualiserTo(
    "supervisorDayVisualiserCanvas",
    derivedDataForLine(line),
    selectedDate,
    appState.supervisorSelectedShift,
    (id) => stageNameByIdForLine(line, id)
  );
}

function renderVisualiser() {
  const data = derivedData();
  const selectedRunRows = selectedRows(data.runRows);
  const selectedDowntimeRows = selectedRows(data.downtimeRows);
  const selectedShiftRows = selectedRows(data.shiftRows);
  const selectedShift = state.selectedShift;

  const shiftMins = num(selectedShiftRows[0]?.totalShiftTime);
  const units = selectedRunRows.reduce((sum, row) => sum + num(row.unitsProduced) * timedLogShiftWeight(row, selectedShift), 0);
  const totalDowntime = selectedDowntimeRows.reduce((sum, row) => sum + num(row.downtimeMins) * timedLogShiftWeight(row, selectedShift), 0);
  const totalNetTime = selectedRunRows.reduce((sum, row) => sum + num(row.netProductionTime) * timedLogShiftWeight(row, selectedShift), 0);
  const netRunRate = totalNetTime > 0 ? units / totalNetTime : 0;
  let utilAccumulator = 0;
  let utilCount = 0;

  const unitsText = formatNum(units, 0);
  const downtimeText = `${formatNum(totalDowntime, 1)} min`;
  const runRateText = `${formatNum(netRunRate, 2)} u/min`;
  const setText = (id, value) => {
    const node = document.getElementById(id);
    if (node) node.textContent = value;
  };

  setText("kpiUnits", unitsText);
  setText("kpiDowntime", downtimeText);
  setText("kpiRunRate", runRateText);
  setText("dayKpiUnits", unitsText);
  setText("dayKpiDowntime", downtimeText);
  setText("dayKpiRunRate", runRateText);

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
      .reduce((sum, row) => sum + num(row.downtimeMins) * timedLogShiftWeight(row, selectedShift), 0);

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
  const utilisationText = `${formatNum(Math.max(lineUtil, 0), 1)}%`;
  setText("kpiUtilisation", utilisationText);
  setText("dayKpiUtilisation", utilisationText);
}

function stageDailyMetrics(stage, date, shift, data) {
  const shiftRows = selectedShiftRowsByDate(data.shiftRows, date, shift, { line: state });
  const shiftMins = shiftRows.reduce((sum, row) => sum + num(row.totalShiftTime), 0);
  const stageDowntime = selectedShiftRowsByDate(data.downtimeRows, date, shift, { line: state })
    .filter((row) => matchesStage(stage, row.equipment))
    .reduce((sum, row) => sum + num(row.downtimeMins) * timedLogShiftWeight(row, shift), 0);

  const runRows = selectedShiftRowsByDate(data.runRows, date, shift, { line: state });
  const units = runRows.reduce((sum, row) => sum + num(row.unitsProduced) * timedLogShiftWeight(row, shift), 0);
  const totalNetTime = runRows.reduce((sum, row) => sum + num(row.netProductionTime) * timedLogShiftWeight(row, shift), 0);
  const netRunRate = totalNetTime > 0 ? units / totalNetTime : 0;
  const uptimeRatio = shiftMins > 0 ? Math.max(0, (shiftMins - stageDowntime) / shiftMins) : 0;
  const stageEtc = netRunRate * uptimeRatio;
  const totalMaxThroughput = stageTotalMaxThroughput(stage.id, shift);
  const utilisation = totalMaxThroughput > 0 ? (stageEtc / totalMaxThroughput) * 100 : 0;

  return { date, shiftMins, stageDowntime, utilisation, stageEtc };
}

function trendDatesForShift(shift, data) {
  const dates = new Set();
  data.shiftRows.forEach((row) => rowMatchesDateShift(row, row.date, shift) && row.date && dates.add(row.date));
  data.runRows.forEach((row) => rowMatchesDateShift(row, row.date, shift) && row.date && dates.add(row.date));
  data.downtimeRows.forEach((row) => rowMatchesDateShift(row, row.date, shift) && row.date && dates.add(row.date));
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
  ensureManagerLogRowIds(state);
  const data = derivedData();
  const activeLineId = String(state?.id || "");
  if (managerLogInlineEdit.lineId && managerLogInlineEdit.lineId !== activeLineId) clearManagerLogInlineEdit();

  const rowExistsForEdit = (() => {
    if (!managerLogInlineEdit.logId) return true;
    const rowList =
      managerLogInlineEdit.type === "shift"
        ? state.shiftRows
        : managerLogInlineEdit.type === "run"
          ? state.runRows
          : managerLogInlineEdit.type === "downtime"
            ? state.downtimeRows
            : [];
    return (rowList || []).some((row) => String(row.id || "") === managerLogInlineEdit.logId);
  })();
  if (!rowExistsForEdit) clearManagerLogInlineEdit();

  const inlineInputHtml = (field, value, { type = "text", min = "", step = "", placeholder = "" } = {}) => {
    const minAttr = min !== "" ? ` min="${htmlEscape(min)}"` : "";
    const stepAttr = step !== "" ? ` step="${htmlEscape(step)}"` : "";
    const placeholderAttr = placeholder ? ` placeholder="${htmlEscape(placeholder)}"` : "";
    return `<input class="table-inline-input" data-inline-field="${htmlEscape(field)}" type="${htmlEscape(type)}" value="${htmlEscape(
      value ?? ""
    )}"${minAttr}${stepAttr}${placeholderAttr} />`;
  };
  const isInlineEditing = (type, rowId) => isManagerLogInlineEditRow(activeLineId, type, rowId);
  const actionHtml = (type, row, { editing = false, pending = false, pendingLabel = "Pending" } = {}) => `
    <div class="table-action-stack">
      ${pending ? `<span class="table-pending-pill">${htmlEscape(pendingLabel)}</span>` : ""}
      <button type="button" class="table-edit-pill${editing ? " is-save" : ""}" data-log-action="${
        editing ? "save" : "edit"
      }" data-log-type="${type}" data-log-id="${row.id || ""}">${editing ? "Save" : "Edit"}</button>
    </div>
  `;

  const sortNewestFirst = (rows, primaryTimeField) =>
    (rows || [])
      .slice()
      .sort((a, b) => {
        const tsDiff = rowNewestSortValue(b, primaryTimeField) - rowNewestSortValue(a, primaryTimeField);
        if (tsDiff !== 0) return tsDiff;
        const submittedCmp = String(b?.submittedAt || "").localeCompare(String(a?.submittedAt || ""));
        if (submittedCmp !== 0) return submittedCmp;
        const dateCmp = String(b?.date || "").localeCompare(String(a?.date || ""));
        if (dateCmp !== 0) return dateCmp;
        return String(b?.id || "").localeCompare(String(a?.id || ""));
      });

  const openBreakByShiftLogId = new Set(
    (data.breakRows || [])
      .filter((row) => isPendingBreakLogRow(row))
      .map((row) => String(row.shiftLogId || ""))
      .filter(Boolean)
  );
  const breakSummaryByShift = new Map();
  (data.breakRows || []).forEach((row) => {
    const key = `${row.date}__${row.shift}`;
    const prev = breakSummaryByShift.get(key) || { count: 0, mins: 0 };
    breakSummaryByShift.set(key, {
      count: prev.count + 1,
      mins: prev.mins + Math.max(0, num(row.breakMins))
    });
  });
  const displayShiftRows = sortNewestFirst(data.shiftRows, "startTime").map((row) => {
    const key = `${row.date}__${row.shift}`;
    const summary = breakSummaryByShift.get(key) || { count: 0, mins: 0 };
    const editing = isInlineEditing("shift", String(row.id || ""));
    const hasOpenBreak = openBreakByShiftLogId.has(String(row.id || ""));
    const pending = isPendingShiftLogRow(row) || hasOpenBreak;
    const pendingLabel = hasOpenBreak ? "Pending | Break Active" : "Pending";
    return {
      ...row,
      __rowClass: pending ? "table-row-pending" : "",
      crewOnShift: editing ? inlineInputHtml("crewOnShift", Math.max(0, num(row.crewOnShift)), { type: "number", min: "0", step: "1" }) : row.crewOnShift,
      startTime: editing ? inlineInputHtml("startTime", String(row.startTime || ""), { placeholder: "HH:MM" }) : row.startTime,
      finishTime: editing ? inlineInputHtml("finishTime", String(row.finishTime || ""), { placeholder: "HH:MM" }) : row.finishTime,
      breakCount: summary.count,
      breakTimeMins: summary.mins,
      submittedBy: managerLogSubmittedByLabel(row),
      action: actionHtml("shift", row, { editing, pending, pendingLabel })
    };
  });

  renderTable("shiftTable", SHIFT_COLUMNS, displayShiftRows, {
    Date: "date",
    Shift: "shift",
    "Crew On Shift": "crewOnShift",
    "Start Time": "startTime",
    "Finish Time": "finishTime",
    "Break Count": "breakCount",
    "Break Time (min)": "breakTimeMins",
    "Total Shift Time": "totalShiftTime",
    "Submitted By": "submittedBy",
    Action: "action"
  });

  const displayRunRows = sortNewestFirst(data.runRows, "productionStartTime").map((row) => {
    const editing = isInlineEditing("run", String(row.id || ""));
    const pending = isPendingRunLogRow(row);
    return {
      ...row,
      __rowClass: pending ? "table-row-pending" : "",
      assignedShift: resolveTimedLogShiftLabel(row, state, "productionStartTime", "finishTime"),
      product: editing ? inlineInputHtml("product", String(row.product || "")) : row.product,
      productionStartTime: editing
        ? inlineInputHtml("productionStartTime", String(row.productionStartTime || ""), { placeholder: "HH:MM" })
        : row.productionStartTime,
      finishTime: editing ? inlineInputHtml("finishTime", String(row.finishTime || ""), { placeholder: "HH:MM" }) : row.finishTime,
      unitsProduced: editing ? inlineInputHtml("unitsProduced", Math.max(0, num(row.unitsProduced)), { type: "number", min: "0", step: "1" }) : row.unitsProduced,
      submittedBy: managerLogSubmittedByLabel(row),
      action: actionHtml("run", row, { editing, pending })
    };
  });
  renderTable("runTable", RUN_COLUMNS, displayRunRows, {
    Date: "date",
    "Assigned Shift": "assignedShift",
    Product: "product",
    "Production Start Time": "productionStartTime",
    "Finish Time": "finishTime",
    "Units Produced": "unitsProduced",
    "Gross Production Time": "grossProductionTime",
    "Associated Down Time": "associatedDownTime",
    "Net Production Time": "netProductionTime",
    "Gross Run Rate": "grossRunRate",
    "Net Run Rate": "netRunRate",
    "Submitted By": "submittedBy",
    Action: "action"
  });

  const displayDowntimeRows = sortNewestFirst(data.downtimeRows, "downtimeStart").map((row) => {
    const editing = isInlineEditing("downtime", String(row.id || ""));
    const pending = isPendingDowntimeLogRow(row);
    return {
      ...row,
      __rowClass: pending ? "table-row-pending" : "",
      assignedShift: resolveTimedLogShiftLabel(row, state, "downtimeStart", "downtimeFinish"),
      downtimeStart: editing ? inlineInputHtml("downtimeStart", String(row.downtimeStart || ""), { placeholder: "HH:MM" }) : row.downtimeStart,
      downtimeFinish: editing ? inlineInputHtml("downtimeFinish", String(row.downtimeFinish || ""), { placeholder: "HH:MM" }) : row.downtimeFinish,
      equipment: row.equipment ? stageNameById(row.equipment) : "-",
      reason: editing ? inlineInputHtml("reason", String(row.reason || "")) : row.reason,
      submittedBy: managerLogSubmittedByLabel(row),
      action: actionHtml("downtime", row, { editing, pending })
    };
  });
  renderTable("downtimeTable", DOWN_COLUMNS, displayDowntimeRows, {
    Date: "date",
    "Assigned Shift": "assignedShift",
    "Downtime Start": "downtimeStart",
    "Downtime Finish": "downtimeFinish",
    "Downtime (mins)": "downtimeMins",
    Equipment: "equipment",
    Reason: "reason",
    "Submitted By": "submittedBy",
    Action: "action"
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
  const selectedRunRows = selectedShiftRowsByDate(data.runRows, selectedDate, selectedShift, { line });
  const selectedDownRows = selectedShiftRowsByDate(data.downtimeRows, selectedDate, selectedShift, { line });
  const selectedShiftRows = selectedShiftRowsByDate(data.shiftRows, selectedDate, selectedShift, { line });
  const shiftMins = selectedShiftRows.reduce((sum, row) => sum + num(row.totalShiftTime), 0);
  const units = selectedRunRows.reduce((sum, row) => sum + num(row.unitsProduced) * timedLogShiftWeight(row, selectedShift), 0);
  const totalDowntime = selectedDownRows.reduce((sum, row) => sum + num(row.downtimeMins) * timedLogShiftWeight(row, selectedShift), 0);
  const totalNetTime = selectedRunRows.reduce((sum, row) => sum + num(row.netProductionTime) * timedLogShiftWeight(row, selectedShift), 0);
  const netRunRate = totalNetTime > 0 ? units / totalNetTime : 0;
  let utilAccumulator = 0;
  let utilCount = 0;
  let bottleneckCard = null;
  let bottleneckUtil = -1;
  const activeCrew = isFullDayShift(selectedShift)
    ? Object.fromEntries(
        stages.map((stage) => [
          stage.id,
          {
            crew: Math.max(
              0,
              num(line?.crewsByShift?.Day?.[stage.id]?.crew),
              num(line?.crewsByShift?.Night?.[stage.id]?.crew)
            )
          }
        ])
      )
    : line?.crewsByShift?.[selectedShift] || defaultStageCrew(stages);

  const svUnitsText = formatNum(units, 0);
  const svDowntimeText = `${formatNum(totalDowntime, 1)} min`;
  const svRunRateText = `${formatNum(netRunRate, 2)} u/min`;
  const setSvText = (id, value) => {
    const node = document.getElementById(id);
    if (node) node.textContent = value;
  };

  setSvText("svKpiUnits", svUnitsText);
  setSvText("svKpiDowntime", svDowntimeText);
  setSvText("svKpiRunRate", svRunRateText);
  setSvText("svDayKpiUnits", svUnitsText);
  setSvText("svDayKpiDowntime", svDowntimeText);
  setSvText("svDayKpiRunRate", svRunRateText);
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
      .reduce((sum, row) => sum + num(row.downtimeMins) * timedLogShiftWeight(row, selectedShift), 0);
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
  const svUtilText = `${formatNum(Math.max(lineUtil, 0), 1)}%`;
  setSvText("svKpiUtilisation", svUtilText);
  setSvText("svDayKpiUtilisation", svUtilText);
}

function renderHome() {
  const homeTitle = document.getElementById("homeTitle");
  const sidebarBackdrop = document.getElementById("sidebarBackdrop");
  const homeSidebar = document.getElementById("homeSidebar");
  const headerLogoutBtn = document.getElementById("supervisorLogout");
  const homeUserChip = document.getElementById("homeUserChip");
  const homeUserAvatar = document.getElementById("homeUserAvatar");
  const homeUserRole = document.getElementById("homeUserRole");
  const homeUserIdentity = document.getElementById("homeUserIdentity");
  const managerHome = document.getElementById("managerHome");
  const supervisorHome = document.getElementById("supervisorHome");
  const modeManagerBtn = document.getElementById("modeManager");
  const modeSupervisorBtn = document.getElementById("modeSupervisor");
  const lineModeManagerBtn = document.getElementById("lineModeManager");
  const lineModeSupervisorBtn = document.getElementById("lineModeSupervisor");
  const managerLoginSection = document.getElementById("managerLoginSection");
  const managerAppSection = document.getElementById("managerAppSection");
  const dashboardDateInput = document.getElementById("dashboardDate");
  const dashboardShiftButtons = Array.from(document.querySelectorAll("[data-dash-shift]"));
  const dashboardTable = document.getElementById("dashboardTable");
  const loginSection = document.getElementById("supervisorLoginSection");
  const appSection = document.getElementById("supervisorAppSection");
  const supervisorMobileModeBtn = document.getElementById("supervisorMobileMode");
  const lineSelect = document.getElementById("supervisorLineSelect");
  const svDateInputs = Array.from(document.querySelectorAll("[data-sv-date]"));
  const svShiftButtons = Array.from(document.querySelectorAll("[data-sv-shift]"));
  const shiftDateInput = document.getElementById("superShiftDate");
  const shiftShiftInput = document.getElementById("superShiftShift");
  const shiftLogIdInput = document.getElementById("superShiftLogId");
  const shiftOpenBreakIdInput = document.getElementById("superShiftOpenBreakId");
  const shiftBreakStartBtn = document.getElementById("superShiftBreakStart");
  const shiftBreakEndBtn = document.getElementById("superShiftBreakEnd");
  const shiftOpenList = document.getElementById("superShiftOpenList");
  const runDateInput = document.getElementById("superRunDate");
  const runLogIdInput = document.getElementById("superRunLogId");
  const runOpenList = document.getElementById("superRunOpenList");
  const downDateInput = document.getElementById("superDownDate");
  const downReasonCategoryInput = document.getElementById("superDownReasonCategory");
  const downReasonDetailInput = document.getElementById("superDownReasonDetail");
  const downLogIdInput = document.getElementById("superDownLogId");
  const downOpenList = document.getElementById("superDownOpenList");
  const entryList = document.getElementById("supervisorEntryList");
  const entryCards = document.getElementById("supervisorEntryCards");
  const superMainTabBtns = Array.from(document.querySelectorAll("[data-super-main-tab]"));
  const superMainPanels = Array.from(document.querySelectorAll(".supervisor-main-panel"));
  const session = normalizeSupervisorSession(appState.supervisorSession, appState.supervisors, appState.lines);
  appState.supervisorSession = session;
  const isSupervisor = appState.appMode === "supervisor";
  const managerSessionActive = Boolean(managerBackendSession.backendToken);
  const activeSupervisor = session?.username ? supervisorByUsername(session.username) : null;

  if (homeTitle) {
    homeTitle.textContent = "Production Line Dashboard";
  }
  if (headerLogoutBtn) {
    const showLogout = isSupervisor ? Boolean(session) : managerSessionActive;
    headerLogoutBtn.classList.toggle("hidden", !showLogout);
    headerLogoutBtn.textContent = "Logout";
  }
  if (homeUserChip) {
    const showSupervisorTile = isSupervisor && Boolean(session);
    const showManagerTile = !isSupervisor && managerSessionActive;
    homeUserChip.classList.toggle("hidden", !(showManagerTile || showSupervisorTile));
    if (showSupervisorTile) {
      const supervisorLabel = String(activeSupervisor?.name || "Supervisor").trim() || "Supervisor";
      if (homeUserRole) homeUserRole.textContent = supervisorLabel;
      if (homeUserIdentity) homeUserIdentity.textContent = String(session?.username || "").trim();
      if (homeUserAvatar) homeUserAvatar.textContent = (supervisorLabel.charAt(0) || "S").toUpperCase();
    } else {
      if (homeUserRole) homeUserRole.textContent = "Manager";
      if (homeUserIdentity) homeUserIdentity.textContent = "production@plant.local";
      if (homeUserAvatar) homeUserAvatar.textContent = "M";
    }
  }

  managerHome.classList.toggle("hidden", isSupervisor);
  supervisorHome.classList.toggle("hidden", !isSupervisor);
  if (managerLoginSection) managerLoginSection.classList.toggle("hidden", isSupervisor || managerSessionActive);
  if (managerAppSection) managerAppSection.classList.toggle("hidden", isSupervisor || !managerSessionActive);
  modeManagerBtn.classList.toggle("active", !isSupervisor);
  modeSupervisorBtn.classList.toggle("active", isSupervisor);
  modeManagerBtn.setAttribute("aria-pressed", String(!isSupervisor));
  modeSupervisorBtn.setAttribute("aria-pressed", String(isSupervisor));
  if (lineModeManagerBtn) lineModeManagerBtn.classList.toggle("active", !isSupervisor);
  if (lineModeSupervisorBtn) lineModeSupervisorBtn.classList.toggle("active", isSupervisor);
  if (lineModeManagerBtn) lineModeManagerBtn.setAttribute("aria-pressed", String(!isSupervisor));
  if (lineModeSupervisorBtn) lineModeSupervisorBtn.setAttribute("aria-pressed", String(isSupervisor));
  supervisorMobileModeBtn.classList.toggle("hidden", true);

  if (!isSupervisor && !managerSessionActive) {
    if (homeSidebar) homeSidebar.classList.remove("open");
    if (sidebarBackdrop) sidebarBackdrop.classList.add("hidden");
    return;
  }

  const cards = document.getElementById("lineCards");
  const lineList = hostedDatabaseAvailable ? Object.values(appState.lines || {}) : [];
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
  if (!hostedDatabaseAvailable) {
    cards.innerHTML = `
      <article class="line-card line-card-unavailable" role="status">
        <h3>Data base not avaiable</h3>
        <p>We could not load production lines because the database is currently unavailable.</p>
      </article>
    `;
  } else if (!lineList.length) {
    cards.innerHTML = `
      <article class="line-card line-card-empty" role="status">
        <h3>No Production Lines</h3>
        <p>No production lines are currently available.</p>
      </article>
    `;
  } else {
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
              <button type="button" class="ghost-btn" data-edit-line="${line.id}">Edit</button>
              <button type="button" class="danger" data-delete-line="${line.id}">Delete</button>
            </div>
          </article>
        `;
      })
      .join("");
  }
  if (statTotalLines) statTotalLines.textContent = formatNum(lineList.length, 0);
  if (statSupervisors) statSupervisors.textContent = formatNum((appState.supervisors || []).length, 0);
  if (statShiftRecords) statShiftRecords.textContent = formatNum(lineList.reduce((sum, line) => sum + (line.shiftRows?.length || 0), 0), 0);
  if (statRunRecords) statRunRecords.textContent = formatNum(lineList.reduce((sum, line) => sum + (line.runRows?.length || 0), 0), 0);

  if (!isSupervisor) {
    if (homeSidebar) homeSidebar.classList.remove("open");
    if (sidebarBackdrop) sidebarBackdrop.classList.add("hidden");
  }

  if (!isSupervisor) return;

  const assignedLineShiftMap = normalizeSupervisorLineShifts(session?.assignedLineShifts, appState.lines, session?.assignedLineIds || []);
  let assignedIds = Object.keys(assignedLineShiftMap);
  if (session) {
    session.assignedLineIds = assignedIds.slice();
    session.assignedLineShifts = clone(assignedLineShiftMap);
  }
  loginSection.classList.toggle("hidden", Boolean(session));
  appSection.classList.toggle("hidden", !session);
  if (!session) return;

  const activeMainTab = ["supervisorDay", "supervisorData"].includes(appState.supervisorMainTab)
    ? appState.supervisorMainTab
    : "supervisorDay";
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
  const entryAllowedShifts = expandedSupervisorShiftAccess(assignedLineShiftMap[appState.supervisorSelectedLineId]);
  const visualAllowedShifts = SHIFT_OPTIONS.slice();
  if (!appState.supervisorSelectedDate) appState.supervisorSelectedDate = todayISO();
  if (!visualAllowedShifts.includes(appState.supervisorSelectedShift)) {
    appState.supervisorSelectedShift = visualAllowedShifts.includes("Full Day") ? "Full Day" : visualAllowedShifts[0] || "Day";
  }

  supervisorMobileModeBtn.classList.toggle("hidden", false);
  appSection.classList.toggle("mobile-mode", Boolean(appState.supervisorMobileMode));
  supervisorMobileModeBtn.classList.toggle("active", Boolean(appState.supervisorMobileMode));
  supervisorMobileModeBtn.textContent = appState.supervisorMobileMode ? "Mobile Mode On" : "Mobile Mode";
  lineSelect.innerHTML = assignedIds.length
    ? assignedIds.map((id) => `<option value="${id}">${appState.lines[id].name}</option>`).join("")
    : `<option value="">No assigned lines</option>`;
  lineSelect.value = appState.supervisorSelectedLineId || assignedIds[0] || "";
  svDateInputs.forEach((svDateInput) => {
    svDateInput.value = appState.supervisorSelectedDate;
  });
  [shiftShiftInput].forEach((shiftInput) => {
    if (!shiftInput) return;
    Array.from(shiftInput.options).forEach((option) => {
      option.disabled = !entryAllowedShifts.includes(option.value);
    });
    if (!entryAllowedShifts.includes(shiftInput.value)) {
      shiftInput.value = entryAllowedShifts[0] || "Day";
    }
    shiftInput.disabled = !entryAllowedShifts.length;
  });
  svShiftButtons.forEach((btn) => {
    const shift = btn.dataset.svShift || "";
    const active = shift === appState.supervisorSelectedShift;
    const enabled = visualAllowedShifts.includes(shift);
    btn.disabled = !enabled;
    btn.classList.toggle("active", active);
    btn.setAttribute("aria-pressed", String(active));
  });
  if (downReasonDetailInput) {
    setDowntimeDetailOptions(
      downReasonDetailInput,
      selectedSupervisorLine(),
      String(downReasonCategoryInput?.value || ""),
      downReasonDetailInput.value || ""
    );
  }
  renderSupervisorVisualiser(selectedSupervisorLine(), appState.supervisorSelectedDate, appState.supervisorSelectedShift);
  renderSupervisorDayVisualiser(selectedSupervisorLine(), appState.supervisorSelectedDate);

  if (!shiftDateInput.value) shiftDateInput.value = todayISO();
  if (!runDateInput.value) runDateInput.value = todayISO();
  if (!downDateInput.value) downDateInput.value = todayISO();

  const pickLatest = (rows = []) =>
    rows.reduce((latest, row) => {
      if (!latest) return row;
      const a = Date.parse(row?.submittedAt || "") || 0;
      const b = Date.parse(latest?.submittedAt || "") || 0;
      return a >= b ? row : latest;
    }, null);
  const isRunOpen = (row) =>
    strictTimeValid(row?.productionStartTime) &&
    strictTimeValid(row?.finishTime) &&
    row.productionStartTime === row.finishTime;
  const isDownOpen = (row) =>
    strictTimeValid(row?.downtimeStart) &&
    strictTimeValid(row?.downtimeFinish) &&
    row.downtimeStart === row.downtimeFinish;
  const isShiftOpen = (row) =>
    strictTimeValid(row?.startTime) &&
    strictTimeValid(row?.finishTime) &&
    row.startTime === row.finishTime;
  const isBreakOpen = (row) =>
    strictTimeValid(row?.breakStart) &&
    !strictTimeValid(row?.breakFinish);

  const activeLine = selectedSupervisorLine();
  if (activeLine) {
    const shiftKeyDate = shiftDateInput.value || appState.supervisorSelectedDate;
    const shiftKeyShift = shiftShiftInput?.value || appState.supervisorSelectedShift;
    const shiftOpenRows = (activeLine.shiftRows || [])
      .filter((row) => rowMatchesDateShift(row, shiftKeyDate, shiftKeyShift) && isShiftOpen(row))
      .slice()
      .sort((a, b) => {
        const submittedCmp = String(b.submittedAt || "").localeCompare(String(a.submittedAt || ""));
        if (submittedCmp !== 0) return submittedCmp;
        return String(a.startTime || "").localeCompare(String(b.startTime || ""));
      });
    const hasShiftRow = (rowId) => shiftOpenRows.some((row) => String(row.id || "") === String(rowId || ""));
    if (supervisorShiftTileEditId && !hasShiftRow(supervisorShiftTileEditId)) {
      supervisorShiftTileEditId = "";
    }
    if (shiftLogIdInput && shiftLogIdInput.value && !hasShiftRow(shiftLogIdInput.value)) {
      shiftLogIdInput.value = "";
    }
    const selectedShiftId = hasShiftRow(shiftLogIdInput?.value)
      ? String(shiftLogIdInput?.value || "")
      : hasShiftRow(supervisorShiftTileEditId)
        ? String(supervisorShiftTileEditId || "")
        : "";
    if (selectedShiftId) {
      if (shiftLogIdInput) shiftLogIdInput.value = selectedShiftId;
      supervisorShiftTileEditId = selectedShiftId;
    }
    const activeShiftId = selectedShiftId;
    const activeShiftLog = shiftOpenRows.find((row) => String(row.id || "") === activeShiftId) || shiftOpenRows[0] || null;
    const openBreak = pickLatest((activeLine.breakRows || []).filter((row) => row.shiftLogId === activeShiftLog?.id && isBreakOpen(row)));
    if (shiftOpenBreakIdInput) shiftOpenBreakIdInput.value = openBreak?.id || "";
    if (shiftBreakStartBtn) shiftBreakStartBtn.disabled = !activeShiftLog?.id || Boolean(openBreak?.id);
    if (shiftBreakEndBtn) shiftBreakEndBtn.disabled = !activeShiftLog?.id || !openBreak?.id;
    if (shiftOpenList) {
      shiftOpenList.innerHTML = shiftOpenRows.length
        ? `
          <div class="pending-log-list">
            ${shiftOpenRows
              .map((row) => {
                const isSelected = Boolean(supervisorShiftTileEditId) && String(supervisorShiftTileEditId) === String(row.id || "");
                const selectedClass = isSelected ? " active" : "";
                const rowBreakRows = breakRowsForShift(activeLine, row.id);
                const rowOpenBreak = pickLatest(rowBreakRows.filter((breakRow) => isBreakOpen(breakRow)));
                const breakEditorRowsHtml = rowBreakRows
                  .map(
                    (breakRow, index) => `
                      <div class="pending-break-editor-row" data-break-editor-row data-break-id="${htmlEscape(breakRow.id || "")}" data-break-index="${index}">
                        <span class="pending-break-editor-label">Break ${index + 1}</span>
                        <input class="pending-break-editor-input" data-break-field="start" value="${htmlEscape(breakRow.breakStart || "")}" placeholder="Start HH:MM" />
                        <input class="pending-break-editor-input" data-break-field="finish" value="${htmlEscape(breakRow.breakFinish || "")}" placeholder="Finish HH:MM" />
                      </div>
                    `
                  )
                  .join("");
                return `
                  <article class="pending-log-item${selectedClass}">
                    <div class="pending-log-meta">
                      <h5>${htmlEscape(row.shift || "Shift")} Shift</h5>
                      <p>
                        ${htmlEscape(row.date || "-")} | Start ${htmlEscape(row.startTime || "-")} | Crew ${formatNum(Math.max(0, num(row.crewOnShift)), 0)}${rowOpenBreak?.id ? ' <span class="pending-break-active">Break active</span>' : ""}
                      </p>
                    </div>
                    <div class="pending-log-actions">
                      <button type="button" class="table-edit-pill ghost-btn" data-super-shift-edit="${row.id}">Edit</button>
                      ${
                        rowOpenBreak?.id
                          ? `<button type="button" class="table-edit-pill ghost-btn" data-super-shift-break-end="${row.id}" data-super-shift-break-id="${rowOpenBreak.id}">End Break</button>`
                          : `<button type="button" class="table-edit-pill ghost-btn" data-super-shift-break-start="${row.id}">Start Break</button>`
                      }
                      <button type="button" class="table-edit-pill finalise-pill" data-super-shift-complete="${row.id}">Finalise</button>
                    </div>
                    ${
                      isSelected
                        ? `
                          <div class="pending-break-editor">
                            <h6>Shift Breaks</h6>
                            <div class="pending-break-editor-list">
                              ${breakEditorRowsHtml || `<p class="pending-break-editor-empty">No breaks logged yet.</p>`}
                              <div class="pending-break-editor-row pending-break-editor-row-new" data-break-editor-row data-break-id="" data-break-index="-1">
                                <span class="pending-break-editor-label">New Break</span>
                                <input class="pending-break-editor-input" data-break-field="start" value="" placeholder="Start HH:MM" />
                                <input class="pending-break-editor-input" data-break-field="finish" value="" placeholder="Finish HH:MM" />
                              </div>
                            </div>
                          </div>
                        `
                        : ""
                    }
                  </article>
                `;
              })
              .join("")}
          </div>
        `
        : `<p class="muted pending-log-empty">No open shift logs for this date/shift.</p>`;
    }

    const runKeyDate = runDateInput.value || appState.supervisorSelectedDate;
    const runOpenRows = (activeLine.runRows || [])
      .filter(
        (row) =>
          row.date === runKeyDate &&
          rowMatchesDateShift(row, runKeyDate, appState.supervisorSelectedShift, {
            line: activeLine,
            startField: "productionStartTime",
            finishField: "finishTime"
          }) &&
          isRunOpen(row)
      )
      .slice()
      .sort((a, b) => {
        const submittedCmp = String(b.submittedAt || "").localeCompare(String(a.submittedAt || ""));
        if (submittedCmp !== 0) return submittedCmp;
        return String(a.productionStartTime || "").localeCompare(String(b.productionStartTime || ""));
      });
    if (runLogIdInput && !runOpenRows.some((row) => row.id === runLogIdInput.value)) {
      runLogIdInput.value = "";
    }
    if (runOpenList) {
      runOpenList.innerHTML = runOpenRows.length
        ? `
          <div class="pending-log-list">
            ${runOpenRows
              .map((row) => {
                const selectedClass = runLogIdInput?.value && runLogIdInput.value === row.id ? " active" : "";
                const shiftLabel = resolveTimedLogShiftLabel(row, activeLine, "productionStartTime", "finishTime");
                return `
                  <article class="pending-log-item${selectedClass}">
                    <div class="pending-log-meta">
                      <h5>${htmlEscape(row.product || "Run")}</h5>
                      <p>
                        ${htmlEscape(row.date || "-")} | ${htmlEscape(shiftLabel)} | Start ${htmlEscape(row.productionStartTime || "-")} | Units ${formatNum(Math.max(0, num(row.unitsProduced)), 0)}
                      </p>
                    </div>
                    <div class="pending-log-actions">
                      <button type="button" class="table-edit-pill ghost-btn" data-super-run-edit="${row.id}">Edit</button>
                      <button type="button" class="table-edit-pill finalise-pill" data-super-run-complete="${row.id}">Finalise</button>
                    </div>
                  </article>
                `;
              })
              .join("")}
          </div>
        `
        : `<p class="muted pending-log-empty">No open production runs for this date/shift.</p>`;
    }

    const downKeyDate = downDateInput.value || appState.supervisorSelectedDate;
    const downOpenRows = (activeLine.downtimeRows || [])
      .filter(
        (row) =>
          row.date === downKeyDate &&
          rowMatchesDateShift(row, downKeyDate, appState.supervisorSelectedShift, {
            line: activeLine,
            startField: "downtimeStart",
            finishField: "downtimeFinish"
          }) &&
          isDownOpen(row)
      )
      .slice()
      .sort((a, b) => {
        const submittedCmp = String(b.submittedAt || "").localeCompare(String(a.submittedAt || ""));
        if (submittedCmp !== 0) return submittedCmp;
        return String(a.downtimeStart || "").localeCompare(String(b.downtimeStart || ""));
      });
    if (downLogIdInput && !downOpenRows.some((row) => row.id === downLogIdInput.value)) {
      downLogIdInput.value = "";
    }
    if (downOpenList) {
      downOpenList.innerHTML = downOpenRows.length
        ? `
          <div class="pending-log-list">
            ${downOpenRows
              .map((row) => {
                const selectedClass = downLogIdInput?.value && downLogIdInput.value === row.id ? " active" : "";
                const parsedReason = parseDowntimeReasonParts(row.reason, row.equipment);
                const rowCategory = row.reasonCategory || parsedReason.reasonCategory || "Downtime";
                const rowDetail = row.reasonDetail || parsedReason.reasonDetail || "";
                const detailLabel = rowCategory === "Equipment"
                  ? stageNameByIdForLine(activeLine, row.equipment || rowDetail) || "Equipment"
                  : rowDetail;
                const shiftLabel = resolveTimedLogShiftLabel(row, activeLine, "downtimeStart", "downtimeFinish");
                return `
                  <article class="pending-log-item${selectedClass}">
                    <div class="pending-log-meta">
                      <h5>${htmlEscape(rowCategory)}${detailLabel ? ` > ${htmlEscape(detailLabel)}` : ""}</h5>
                      <p>
                        ${htmlEscape(row.date || "-")} | ${htmlEscape(shiftLabel)} | Start ${htmlEscape(row.downtimeStart || "-")}
                      </p>
                    </div>
                    <div class="pending-log-actions">
                      <button type="button" class="table-edit-pill ghost-btn" data-super-down-edit="${row.id}">Edit</button>
                      <button type="button" class="table-edit-pill finalise-pill" data-super-down-complete="${row.id}">Finalise</button>
                    </div>
                  </article>
                `;
              })
              .join("")}
          </div>
        `
        : `<p class="muted pending-log-empty">No open downtime logs for this date/shift.</p>`;
    }
  } else {
    supervisorShiftTileEditId = "";
    if (shiftLogIdInput) shiftLogIdInput.value = "";
    if (shiftOpenBreakIdInput) shiftOpenBreakIdInput.value = "";
    if (runLogIdInput) runLogIdInput.value = "";
    if (downLogIdInput) downLogIdInput.value = "";
    if (shiftOpenList) shiftOpenList.innerHTML = `<p class="muted pending-log-empty">No assigned line selected.</p>`;
    if (runOpenList) runOpenList.innerHTML = `<p class="muted pending-log-empty">No assigned line selected.</p>`;
    if (downOpenList) downOpenList.innerHTML = `<p class="muted pending-log-empty">No assigned line selected.</p>`;
    if (shiftBreakStartBtn) shiftBreakStartBtn.disabled = true;
    if (shiftBreakEndBtn) shiftBreakEndBtn.disabled = true;
  }
  updateSupervisorProgressButtonLabels();

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
          shift: resolveTimedLogShiftLabel(row, line, "productionStartTime", "finishTime"),
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
          shift: resolveTimedLogShiftLabel(row, line, "downtimeStart", "downtimeFinish"),
          type: "Downtime",
          summary: `${stageNameByIdForLine(line, row.equipment) || row.reasonCategory || "Downtime"}: ${row.reason || "-"}`,
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
                `<tr>
                  <td>${htmlEscape(log.lineName)}</td>
                  <td>${htmlEscape(log.date)}</td>
                  <td>${htmlEscape(log.shift)}</td>
                  <td>${htmlEscape(log.type)}</td>
                  <td>${htmlEscape(log.summary)}</td>
                  <td>${htmlEscape(log.supervisor)}</td>
                </tr>`
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
            <h4>${htmlEscape(log.type)} | ${htmlEscape(log.lineName)}</h4>
            <p>${htmlEscape(log.date)} ${htmlEscape(log.shift)} | ${htmlEscape(log.summary)}</p>
            <p>By ${htmlEscape(log.supervisor)}</p>
          </article>
        `
        )
        .join("")
    : `<p class="muted">No supervisor submissions yet.</p>`;

  syncAllNowInputPrefixStates();
}

function renderAll() {
  const managerSessionActive = Boolean(managerBackendSession.backendToken);
  if (appState.activeView === "line" && (appState.appMode !== "manager" || !managerSessionActive)) {
    appState.activeView = "home";
  }

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
  refreshRunCrewingPatternSummaries();

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
  const reasonCategorySelect = document.getElementById("downtimeReasonCategory");
  const reasonDetailSelect = document.getElementById("downtimeReasonDetail");
  if (reasonDetailSelect) {
    setDowntimeDetailOptions(
      reasonDetailSelect,
      state,
      String(reasonCategorySelect?.value || ""),
      reasonDetailSelect.value || ""
    );
  }
  document.querySelectorAll("[data-manager-date]").forEach((input) => {
    input.value = state.selectedDate;
  });
  setShiftToggleUI();
  setActiveDataSubtab();
  renderCrewInputs();
  renderThroughputInputs();
  renderTrackingTables();
  renderAuditTrail();
  renderVisualiser();
  renderDayVisualiser();
  renderLineTrends();
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

function armDbLoadingFailsafe() {
  if (dbLoadingFailsafeTimer) {
    clearTimeout(dbLoadingFailsafeTimer);
    dbLoadingFailsafeTimer = null;
  }
  if (dbLoadingRequestCount <= 0) return;
  dbLoadingFailsafeTimer = window.setTimeout(() => {
    if (dbLoadingRequestCount <= 0) {
      dbLoadingFailsafeTimer = null;
      return;
    }
    console.warn("DB loading watchdog reset: forcing loader hide after timeout.");
    dbLoadingRequestCount = 0;
    hideStartupLoading({ force: true });
    dbLoadingFailsafeTimer = null;
  }, API_REQUEST_TIMEOUT_MS + 2000);
}

function hideStartupLoading({ force = false } = {}) {
  const loader = document.getElementById("startupLoading");
  if (!loader) return;
  if (!force && dbLoadingRequestCount > 0) return;
  loader.classList.add("hidden");
  if (dbLoadingFailsafeTimer && dbLoadingRequestCount <= 0) {
    clearTimeout(dbLoadingFailsafeTimer);
    dbLoadingFailsafeTimer = null;
  }
}

function showStartupLoading(message = "") {
  const loader = document.getElementById("startupLoading");
  if (!loader) return;
  if (message) {
    const label = loader.querySelector("p");
    if (label) label.textContent = message;
  }
  loader.classList.remove("hidden");
  armDbLoadingFailsafe();
}

async function bootstrapApp() {
  try {
    bindTabs();
    bindRunCrewingPatternModal();
    bindHome();
    bindDataSubtabs();
    bindVisualiserControls();
    bindTrendModal();
    bindStageSettingsModal();
    bindForms();
    bindDataControls();
    renderAll();
    if (appState.appMode === "supervisor" && appState.supervisorSession?.backendToken) {
      await refreshHostedState(appState.supervisorSession);
    } else if ((appState.appMode === "manager" || appState.activeView === "line") && managerBackendSession.backendToken) {
      await refreshHostedState();
    }
  } finally {
    hideStartupLoading({ force: true });
  }
}

bootstrapApp();

window.addEventListener("hashchange", () => {
  restoreRouteFromHash();
  state = appState.lines[appState.activeLineId] || appState.lines[Object.keys(appState.lines)[0]] || null;
  if (state) appState.activeLineId = state.id;
  renderAll();
});
