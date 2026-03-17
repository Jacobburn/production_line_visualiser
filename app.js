const STORAGE_KEY = "kebab-line-data-v2";
const STORAGE_BACKUP_KEY = "kebab-line-data-v2-backup";
const AUTH_STORAGE_KEY = "kebab-line-auth-v1";
const ROUTE_STORAGE_KEY = "kebab-line-route-v1";
const HOME_UI_STORAGE_KEY = "kebab-line-home-ui-v1";
const APP_VARIANT = (() => {
  const raw = String(window.PRODUCTION_LINE_APP_VARIANT || "").trim().toLowerCase();
  return raw === "supervisor" ? "supervisor" : "manager";
})();
const API_BASE_URL = `${
  window.PRODUCTION_LINE_API_BASE ||
  "http://localhost:4000"
}`.replace(/\/+$/, "");
const API_REQUEST_TIMEOUT_MS = 45000;
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

const SHIFT_COLUMNS = ["Date", "Shift", "Start Time", "Finish Time", "Break Count", "Break Time (min)", "Total Shift Time", "Notes", "Submitted By", "Action"];
const RUN_COLUMNS = ["Date", "Product", "Production Start Time", "Finish Time", "Units Produced", "Gross Production Time", "Associated Down Time", "Net Production Time", "Gross Run Rate", "Net Run Rate", "Notes", "Submitted By", "Action"];
const DOWN_COLUMNS = ["Date", "Downtime Start", "Downtime Finish", "Downtime (mins)", "Equipment", "Reason", "Notes", "Submitted By", "Action"];
const AUDIT_COLUMNS = ["When", "Actor", "Action", "Details"];
const DASHBOARD_COLUMNS = ["Line", "Date", "Shift", "Units", "Downtime (min)", "Utilisation (%)", "Net Run Rate (u/min)", "Bottleneck", "Staffing"];
const SHIFT_OPTIONS = ["Day", "Night", "Full Day"];
const MANAGER_LINE_TABS = Object.freeze(["visualiser", "dayVisualiser", "lineTrends", "data", "settings"]);
const MANAGER_DATA_TABS = Object.freeze(["dataShift", "dataRun", "dataDown", "dataControls"]);
const LINE_TREND_RANGES = ["day", "week", "month", "quarter"];
const LOG_ASSIGNABLE_SHIFTS = ["Day", "Night"];
const LINE_SHIFT_TRACKER_MIN_WEEKS = 1;
const LINE_SHIFT_TRACKER_DEFAULT_WEEKS = 20;
const LINE_SHIFT_TRACKER_MAX_WEEKS = 260;
const LINE_SHIFT_TRACKER_CELL_SIZE = 8;
const LINE_SHIFT_TRACKER_CELL_GAP = 2;
const LINE_TILE_FEEDBACK_REFRESH_MS = 30000;
const LINE_TILE_DOWNTIME_CRITICAL_MINS = 15;
const TREND_DAILY_WINDOW_SIZE = 22;
const SUPERVISOR_SHIFT_OPTIONS = ["Day", "Night"];
const CREW_SETTINGS_SHIFTS = ["Day", "Night"];
const BACKEND_LOG_NOTES_MAX_LENGTH = 2000;
const BACKEND_DOWNTIME_REASON_MAX_LENGTH = 250;
const DOWNTIME_REASON_PRESETS = {
  "Donor Meat": ["Stock Out", "Late Delivery", "Quality Hold", "Temperature Hold"],
  People: ["Understaffed", "Training", "Handover Delay", "Absence"],
  Materials: ["Film Shortage", "Label Shortage", "Tray Shortage", "Marinade Shortage", "Skewer shortage"],
  Other: ["Cleaning", "QA Hold", "Power", "Unplanned Stop"],
  Break: ["Standard Break", "Non-standard Break"]
};
const DEFAULT_SUPERVISORS = [
  { id: "sup-1", name: "Supervisor", username: "supervisor", password: "supervisor", mode: "all", shifts: ["Day", "Night"] },
  { id: "sup-2", name: "Day Lead", username: "daylead", password: "day123", mode: "even", shifts: ["Day"] },
  { id: "sup-3", name: "Night Lead", username: "nightlead", password: "night123", mode: "odd", shifts: ["Night"] }
];
const ACTION_PRIORITY_OPTIONS = ["Low", "Medium", "High", "Critical"];
const ACTION_STATUS_OPTIONS = ["Open", "In Progress", "Blocked", "Completed"];
const ACTION_STATUS_SORT_ORDER = Object.freeze({
  Blocked: 0,
  Open: 1,
  "In Progress": 2,
  Completed: 3
});
const ACTION_REASON_CATEGORIES = Object.freeze(Array.from(new Set(["Equipment", ...Object.keys(DOWNTIME_REASON_PRESETS)])));
const ACTION_SPECIAL_ASSIGNMENTS = Object.freeze([
  { username: "maintenance", label: "Maintenance" },
  { username: "continuous-improvement", label: "Continuous Improvement" }
]);
const ACTION_SPECIAL_ASSIGNMENT_LABELS = Object.freeze(
  ACTION_SPECIAL_ASSIGNMENTS.reduce((acc, assignment) => {
    acc[assignment.username] = assignment.label;
    return acc;
  }, {})
);
const ACTION_PRIORITY_SCORES = Object.freeze({
  Low: 1,
  Medium: 2,
  High: 3,
  Critical: 4
});
const MANAGER_ACTION_URGENCY_GROUPS = Object.freeze([
  { key: "overdue", label: "Overdue", tone: "critical" },
  { key: "urgent", label: "Urgent", tone: "high" },
  { key: "due-soon", label: "Due Soon", tone: "medium" },
  { key: "unscheduled", label: "Unscheduled", tone: "neutral" },
  { key: "planned", label: "Planned", tone: "low" },
  { key: "completed", label: "Completed", tone: "done" }
]);
const MANAGER_ACTION_URGENCY_GROUP_MAP = Object.freeze(
  MANAGER_ACTION_URGENCY_GROUPS.reduce((acc, group, index) => {
    acc[group.key] = {
      ...group,
      order: index
    };
    return acc;
  }, {})
);
const PRODUCT_CATALOG_MANAGER_DATALIST_ID = "productCatalogManagerRunProducts";
const PRODUCT_CATALOG_SUPERVISOR_DATALIST_ID = "productCatalogSupervisorRunProducts";
const PRODUCT_CATALOG_COLUMN_COUNT = 12;
const PRODUCT_CATALOG_ALL_LINES_TOKEN = "*";
const PRODUCT_CATALOG_ID_ATTR = "data-product-id";
const PRODUCT_CATALOG_LINES_ATTR = "data-product-line-ids";
const DATA_SOURCE_PROVIDER_LABELS = {
  sql: "SQL",
  api: "API"
};

let appState = loadState();
let state = appState.lines[appState.activeLineId] || null;
let visualiserDragMoved = false;
let managerLogInlineEdit = { lineId: "", type: "", logId: "" };
let managerActionTicketEditId = "";
let supervisorShiftTileEditId = "";
let runCrewingPatternModalState = null;
let dayVizBlockModalState = null;
let dayVizBlockModalBusy = false;
let dayVizAddRecordModalState = null;
let passwordResetModalBusy = false;
let trendModalContext = { type: "stage", metricKey: "" };
appState.supervisors = normalizeSupervisors(appState.supervisors, appState.lines);
appState.lineGroups = normalizeLineGroups(appState.lineGroups);
appState.dataSources = normalizeDataSources(appState.dataSources);
appState.supervisorActions = normalizeSupervisorActions(appState.supervisorActions);
setProductCatalogEntries(appState.productCatalogEntries);
let lineModelSyncTimer = null;
const deferredLineModelSyncIds = new Set();
const pendingLineSettingsSaveIds = new Set();
const lineSettingsSaveInFlightIds = new Set();
let lineShiftTrackerResizeTimer = null;
let lineTileFeedbackTimer = null;
let hostedRefreshErrorShown = false;
let hostedDatabaseAvailable = true;
let dbLoadingRequestCount = 0;
let dbLoadingFailsafeTimer = null;
let managerBackendSession = {
  backendToken: "",
  backendLineMap: {},
  backendStageMap: {},
  role: "manager",
  name: "",
  username: ""
};
let pendingManagerDataTabRestore = { lineId: "", tabId: "" };
const dataSourceConnectionTestState = new Map();

function enforceAppVariantState() {
  if (APP_VARIANT === "supervisor") {
    appState.appMode = "supervisor";
    appState.activeView = "home";
    return;
  }
  appState.appMode = "manager";
}

clearLegacyStateStorage();
restoreAuthSessionsFromStorage();
restoreHomeUiState();
restoreRouteFromHash();
enforceAppVariantState();
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

function setManagerActionTicketEdit(actionId = "") {
  managerActionTicketEditId = String(actionId || "");
}

function clearManagerActionTicketEdit() {
  setManagerActionTicketEdit("");
}

function isManagerActionTicketEditRow(actionId) {
  return managerActionTicketEditId === String(actionId || "");
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function sanitizeHomeLineGroupExpanded(raw) {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return {};
  const next = {};
  Object.entries(raw).forEach(([key, value]) => {
    const safeKey = String(key || "").trim();
    if (!safeKey) return;
    next[safeKey] = Boolean(value);
  });
  return next;
}

function readHomeUiStorage() {
  try {
    const raw = window.localStorage.getItem(HOME_UI_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : null;
  } catch {
    return null;
  }
}

function writeHomeUiStorage(value) {
  try {
    if (value && typeof value === "object") {
      window.localStorage.setItem(HOME_UI_STORAGE_KEY, JSON.stringify(value));
    } else {
      window.localStorage.removeItem(HOME_UI_STORAGE_KEY);
    }
  } catch {
    // Ignore storage failures (private mode, quota, disabled storage).
  }
}

function restoreHomeUiState() {
  const stored = readHomeUiStorage();
  appState.homeLineGroupExpanded = sanitizeHomeLineGroupExpanded(stored?.homeLineGroupExpanded);
}

function persistHomeUiState() {
  writeHomeUiStorage({
    homeLineGroupExpanded: sanitizeHomeLineGroupExpanded(appState.homeLineGroupExpanded)
  });
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

function parseManagerLineTabId(value) {
  const tabId = String(value || "").trim();
  return MANAGER_LINE_TABS.includes(tabId) ? tabId : "";
}

function parseManagerDataTabId(value) {
  const tabId = String(value || "").trim();
  return MANAGER_DATA_TABS.includes(tabId) ? tabId : "";
}

function activeManagerLineTabId() {
  const activeStateTab = parseManagerLineTabId(appState.managerLineTab);
  if (activeStateTab) return activeStateTab;
  const activeDomTab = parseManagerLineTabId(document.querySelector(".tab-btn[data-tab].active")?.dataset?.tab);
  if (activeDomTab) return activeDomTab;
  return "visualiser";
}

function setActiveManagerLineTab(tabId) {
  const activeId = parseManagerLineTabId(tabId) || "visualiser";
  document.querySelectorAll(".tab-btn[data-tab]").forEach((btn) => {
    const isActive = btn.dataset.tab === activeId;
    btn.classList.toggle("active", isActive);
    btn.setAttribute("aria-selected", String(isActive));
  });
  document.querySelectorAll(".tab-panel").forEach((panel) => {
    panel.classList.toggle("active", panel.id === activeId);
  });
  appState.managerLineTab = activeId;
}

function applyRouteSnapshot(route) {
  if (!route || typeof route !== "object") return;
  const mode = route.mode;
  if (mode === "manager" || mode === "supervisor") appState.appMode = mode;
  const view = route.view;
  if (view === "home" || view === "line") appState.activeView = view;
  const lineId = String(route.lineId || "");
  if (lineId) appState.activeLineId = lineId;

  if (["supervisorDay", "supervisorData", "supervisorActions"].includes(route.supervisorMainTab)) {
    appState.supervisorMainTab = route.supervisorMainTab;
  }
  const supervisorTab = String(route.supervisorTab || "");
  if (["superShift", "superRun", "superDown", "superVisual"].includes(supervisorTab)) {
    appState.supervisorTab = supervisorTab;
  }
  const supervisorSelectedLineId = String(route.supervisorSelectedLineId || "");
  if (supervisorSelectedLineId) appState.supervisorSelectedLineId = supervisorSelectedLineId;
  const supervisorSelectedDate = String(route.supervisorSelectedDate || "");
  if (isIsoDateValue(supervisorSelectedDate)) {
    appState.supervisorSelectedDate = normalizeWeekdayIsoDate(supervisorSelectedDate, { direction: -1 });
  }
  const supervisorSelectedShift = String(route.supervisorSelectedShift || "");
  if (SHIFT_OPTIONS.includes(supervisorSelectedShift)) appState.supervisorSelectedShift = supervisorSelectedShift;
  const dashboardDate = String(route.dashboardDate || "");
  if (isIsoDateValue(dashboardDate)) {
    appState.dashboardDate = normalizeWeekdayIsoDate(dashboardDate, { direction: -1 });
  }
  const dashboardShift = String(route.dashboardShift || "");
  if (SHIFT_OPTIONS.includes(dashboardShift)) appState.dashboardShift = dashboardShift;
  const managerLineTab = parseManagerLineTabId(route.managerLineTab);
  if (managerLineTab) appState.managerLineTab = managerLineTab;
  const managerDataTab = parseManagerDataTabId(route.managerDataTab);
  if (managerDataTab) {
    const targetLineId = String(appState.activeLineId || "");
    const targetLine = appState.lines?.[targetLineId];
    if (targetLine) {
      targetLine.activeDataTab = managerDataTab;
      if (state && state.id === targetLineId) state.activeDataTab = managerDataTab;
      pendingManagerDataTabRestore = { lineId: "", tabId: "" };
    } else {
      pendingManagerDataTabRestore = { lineId: targetLineId, tabId: managerDataTab };
    }
  }
  enforceAppVariantState();
}

function currentRouteSnapshot() {
  const routeMode = APP_VARIANT === "supervisor" ? "supervisor" : "manager";
  const activeLine = appState.lines?.[appState.activeLineId] || state || null;
  const pendingTabId = parseManagerDataTabId(pendingManagerDataTabRestore.tabId);
  const pendingLineId = String(pendingManagerDataTabRestore.lineId || "");
  const activeLineId = String(activeLine?.id || appState.activeLineId || "");
  const safeSupervisorDate = normalizeWeekdayIsoDate(appState.supervisorSelectedDate || todayISO(), { direction: -1 });
  const safeDashboardDate = normalizeWeekdayIsoDate(appState.dashboardDate || todayISO(), { direction: -1 });
  const managerDataTab =
    pendingTabId && (!pendingLineId || pendingLineId === activeLineId)
      ? pendingTabId
      : parseManagerDataTabId(activeLine?.activeDataTab) || "dataShift";
  return {
    mode: routeMode,
    view: appState.activeView === "line" ? "line" : "home",
    lineId: appState.activeView === "line" ? String(appState.activeLineId || "") : "",
    managerLineTab: activeManagerLineTabId(),
    managerDataTab,
    supervisorMainTab: appState.supervisorMainTab || "supervisorDay",
    supervisorTab: appState.supervisorTab || "superShift",
    supervisorSelectedLineId: String(appState.supervisorSelectedLineId || ""),
    supervisorSelectedDate: safeSupervisorDate,
    supervisorSelectedShift: appState.supervisorSelectedShift || "Full Day",
    dashboardDate: safeDashboardDate,
    dashboardShift: appState.dashboardShift || "Day"
  };
}

function restoreRouteFromHash() {
  try {
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
      managerLineTab: params.get("mlt"),
      managerDataTab: params.get("mdt"),
      supervisorMainTab: params.get("smt"),
      supervisorTab: params.get("st"),
      supervisorSelectedLineId: params.get("sl"),
      supervisorSelectedDate: params.get("sd"),
      supervisorSelectedShift: params.get("ss"),
      dashboardDate: params.get("dd"),
      dashboardShift: params.get("ds")
    });
  } catch (error) {
    console.warn("Route restore failed, using stored snapshot fallback:", error);
    applyRouteSnapshot(readRouteStorage());
  }
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
    if (route.view === "line") {
      params.set("mlt", parseManagerLineTabId(route.managerLineTab) || "visualiser");
      params.set("mdt", parseManagerDataTabId(route.managerDataTab) || "dataShift");
    }
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
    const stageCrew = Number(stage?.crew);
    if (!Number.isFinite(stageCrew) || stageCrew < 0) {
      stage.crew = crew;
    }
    const dayCrew = Number(line.crewsByShift.Day?.[stage.id]?.crew);
    if (!Number.isFinite(dayCrew) || dayCrew < 0) {
      line.crewsByShift.Day[stage.id] = { ...(line.crewsByShift.Day[stage.id] || {}), crew };
    }
    const nightCrew = Number(line.crewsByShift.Night?.[stage.id]?.crew);
    if (!Number.isFinite(nightCrew) || nightCrew < 0) {
      line.crewsByShift.Night[stage.id] = { ...(line.crewsByShift.Night[stage.id] || {}), crew };
    }
    const maxValue = Number(line.stageSettings?.[stage.id]?.maxThroughput);
    if (!Number.isFinite(maxValue) || maxValue < 0) {
      line.stageSettings[stage.id] = { ...(line.stageSettings[stage.id] || {}), maxThroughput };
    }
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

function stageFlowRect(stage) {
  const x = num(stage?.x);
  const y = num(stage?.y);
  const w = Math.max(1, num(stage?.w));
  const h = Math.max(1, num(stage?.h));
  return {
    x,
    y,
    w,
    h,
    cx: x + w / 2,
    cy: y + h / 2
  };
}

function pointOnStageEdgeToward(rect, targetX, targetY) {
  const dx = num(targetX) - rect.cx;
  const dy = num(targetY) - rect.cy;
  if (Math.abs(dx) < 0.0001 && Math.abs(dy) < 0.0001) return { x: rect.cx, y: rect.cy };
  const scaleX = Math.abs(dx) > 0.0001 ? rect.w / (2 * Math.abs(dx)) : Number.POSITIVE_INFINITY;
  const scaleY = Math.abs(dy) > 0.0001 ? rect.h / (2 * Math.abs(dy)) : Number.POSITIVE_INFINITY;
  const scale = Math.min(scaleX, scaleY);
  return {
    x: rect.cx + dx * scale,
    y: rect.cy + dy * scale
  };
}

function autoFlowGuideBetweenStages(fromStage, toStage, index) {
  if (!fromStage || !toStage) return null;
  const fromRect = stageFlowRect(fromStage);
  const toRect = stageFlowRect(toStage);
  const start = pointOnStageEdgeToward(fromRect, toRect.cx, toRect.cy);
  const end = pointOnStageEdgeToward(toRect, fromRect.cx, fromRect.cy);
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const distance = Math.sqrt(dx * dx + dy * dy);
  if (!Number.isFinite(distance) || distance < 0.25) return null;
  return {
    id: `auto-flow-${index}`,
    type: "arrow",
    x: start.x,
    y: start.y - 1,
    w: distance,
    h: 2,
    angle: (Math.atan2(dy, dx) * 180) / Math.PI,
    auto: true
  };
}

function autoFlowGuidesFromStages(stages) {
  if (!Array.isArray(stages) || stages.length < 2) return [];
  const guides = [];
  for (let index = 0; index < stages.length - 1; index += 1) {
    const guide = autoFlowGuideBetweenStages(stages[index], stages[index + 1], index);
    if (guide) guides.push(guide);
  }
  return guides;
}

function lineFlowGuidesForMap(stages, guides) {
  const shapeGuides = normalizeFlowGuides(guides).filter((guide) => guide.type === "shape");
  return [...autoFlowGuidesFromStages(stages), ...shapeGuides];
}

function appendFlowGuidesToMap(map, guides, { editable = false } = {}) {
  if (!map || !Array.isArray(guides) || !guides.length) return;
  guides.forEach((guide) => {
    const isAutoGuide = Boolean(guide.auto);
    const node = document.createElement("div");
    node.className = `flow-guide flow-${guide.type}${isAutoGuide ? " auto-flow" : ""}`;
    if (!isAutoGuide) node.setAttribute("data-guide-id", guide.id);
    node.style.left = `${guide.x}%`;
    node.style.top = `${guide.y}%`;
    node.style.width = `${guide.w}%`;
    node.style.height = `${guide.h}%`;
    node.style.transform = `rotate(${guide.angle || 0}deg)`;
    if (isAutoGuide) node.style.transformOrigin = "0 50%";

    if (editable && !isAutoGuide) {
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

function normalizeLineGroups(lineGroups) {
  const source = Array.isArray(lineGroups) ? lineGroups : [];
  const seen = new Set();
  return source
    .map((group, index) => {
      const id = String(group?.id || "").trim();
      const name = String(group?.name || "").trim();
      if (!id || !name || seen.has(id)) return null;
      seen.add(id);
      const orderRaw = Number(group?.displayOrder);
      const displayOrder = Number.isFinite(orderRaw) ? Math.max(0, Math.floor(orderRaw)) : index;
      return { id, name, displayOrder };
    })
    .filter(Boolean)
    .sort((a, b) => (a.displayOrder - b.displayOrder) || a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));
}

function normalizeDataSourceProvider(provider) {
  const value = String(provider || "").trim().toLowerCase();
  if (value === "api") return "api";
  if (value === "sql") return "sql";
  return "api";
}

function dataSourceKeyFromName(name, fallback = "data-source") {
  const slug = String(name || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 160);
  if (slug) return slug;
  return `${fallback}-${Date.now()}`;
}

function normalizeDataSources(dataSources) {
  const source = Array.isArray(dataSources) ? dataSources : [];
  const seen = new Set();
  return source
    .map((dataSource, index) => {
      const id = String(dataSource?.id || "").trim();
      if (!id || seen.has(id)) return null;
      seen.add(id);
      const sourceName = String(dataSource?.sourceName || dataSource?.name || "").trim();
      const sourceKey = String(dataSource?.sourceKey || dataSource?.key || "").trim();
      const machineNo = String(dataSource?.machineNo || "").trim();
      const deviceName = String(dataSource?.deviceName || "").trim();
      const deviceId = String(dataSource?.deviceId || "").trim();
      const scaleNumber = String(dataSource?.scaleNumber || "").trim();
      const provider = normalizeDataSourceProvider(dataSource?.provider);
      const connectionMode = String(dataSource?.connectionMode || provider || "api").trim().toLowerCase() === "sql" ? "sql" : "api";
      const apiBaseUrl = String(dataSource?.apiBaseUrl || "").trim();
      const hasApiKey = Boolean(dataSource?.hasApiKey);
      const hasSqlCredentials = Boolean(dataSource?.hasSqlCredentials);
      const isActive = dataSource?.isActive !== false;
      const displayOrderRaw = Number(dataSource?.displayOrder);
      const displayOrder = Number.isFinite(displayOrderRaw) ? Math.max(0, Math.floor(displayOrderRaw)) : index;
      return {
        id,
        sourceKey,
        sourceName: sourceName || sourceKey || `Data Source ${index + 1}`,
        machineNo,
        deviceName,
        deviceId,
        scaleNumber,
        provider,
        connectionMode,
        apiBaseUrl,
        hasApiKey,
        hasSqlCredentials,
        isActive,
        displayOrder
      };
    })
    .filter(Boolean)
    .sort(
      (a, b) =>
        (a.displayOrder - b.displayOrder)
        || a.sourceName.localeCompare(b.sourceName, undefined, { sensitivity: "base", numeric: true })
    );
}

function dataSourceById(dataSourceId) {
  const targetId = String(dataSourceId || "").trim();
  if (!targetId) return null;
  return (appState.dataSources || []).find((source) => source.id === targetId) || null;
}

function dataSourceDisplayLabel(dataSource) {
  if (!dataSource) return "Unknown source";
  const machineSuffix = dataSource.machineNo ? ` (Machine ${dataSource.machineNo})` : "";
  return `${dataSource.sourceName || "Data Source"}${machineSuffix}`;
}

function dataSourceAssignments(lines = appState.lines) {
  const map = new Map();
  Object.values(lines || {}).forEach((line) => {
    const lineId = String(line?.id || "").trim();
    const lineName = String(line?.name || "Production Line").trim() || "Production Line";
    (line?.stages || []).forEach((stage, index) => {
      const sourceId = String(stage?.dataSourceId || "").trim();
      if (!sourceId) return;
      const entry = {
        sourceId,
        lineId,
        lineName,
        stageId: String(stage?.id || "").trim(),
        stageName: stageDisplayName(stage, index)
      };
      if (!map.has(sourceId)) map.set(sourceId, []);
      map.get(sourceId).push(entry);
    });
  });
  return map;
}

function rememberDataSourceConnectionTest(sourceId, result = null) {
  const safeSourceId = String(sourceId || "").trim();
  if (!safeSourceId) return;
  if (!result || typeof result !== "object") {
    dataSourceConnectionTestState.delete(safeSourceId);
    return;
  }
  dataSourceConnectionTestState.set(safeSourceId, {
    ok: Boolean(result.ok),
    message: String(result.message || "").trim(),
    testedAt: new Date().toISOString()
  });
}

function dataSourceConnectionTestFor(sourceId) {
  const safeSourceId = String(sourceId || "").trim();
  if (!safeSourceId) return null;
  return dataSourceConnectionTestState.get(safeSourceId) || null;
}

function dataSourceConnectionTestLabel(testState) {
  if (!testState) return { text: "Not tested", css: "" };
  const testedAt = new Date(testState.testedAt || "");
  const hasTime = Number.isFinite(testedAt.getTime());
  const timeText = hasTime
    ? testedAt.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
    : "";
  if (testState.ok) {
    return {
      text: timeText ? `Passed ${timeText}` : "Passed",
      css: " is-success"
    };
  }
  return {
    text: timeText ? `Failed ${timeText}` : "Failed",
    css: " is-failed"
  };
}

function conflictingDataSourceAssignment(sourceId, lineId, stageId, assignments = null) {
  const safeSourceId = String(sourceId || "").trim();
  if (!safeSourceId) return null;
  const safeLineId = String(lineId || "").trim();
  const safeStageId = String(stageId || "").trim();
  const assignmentMap = assignments instanceof Map ? assignments : dataSourceAssignments();
  const list = assignmentMap.get(safeSourceId) || [];
  return list.find((item) => !(item.lineId === safeLineId && item.stageId === safeStageId)) || null;
}

function sortLinesByName(lines) {
  return (Array.isArray(lines) ? lines.slice() : []).sort((a, b) =>
    String(a?.name || "").localeCompare(String(b?.name || ""), undefined, { sensitivity: "base", numeric: true })
  );
}

function sortLinesByDisplayOrder(lines) {
  return (Array.isArray(lines) ? lines.slice() : []).sort((a, b) => {
    const orderA = Number(a?.displayOrder);
    const orderB = Number(b?.displayOrder);
    const hasOrderA = Number.isFinite(orderA);
    const hasOrderB = Number.isFinite(orderB);
    if (hasOrderA && hasOrderB && orderA !== orderB) return orderA - orderB;
    if (hasOrderA !== hasOrderB) return hasOrderA ? -1 : 1;
    return String(a?.name || "").localeCompare(String(b?.name || ""), undefined, { sensitivity: "base", numeric: true });
  });
}

function homeLineGroupExpandedState() {
  if (!appState.homeLineGroupExpanded || typeof appState.homeLineGroupExpanded !== "object" || Array.isArray(appState.homeLineGroupExpanded)) {
    appState.homeLineGroupExpanded = {};
  }
  return appState.homeLineGroupExpanded;
}

function isHomeLineGroupExpanded(groupKey) {
  const key = String(groupKey || "").trim();
  if (!key) return true;
  const stateMap = homeLineGroupExpandedState();
  if (!Object.prototype.hasOwnProperty.call(stateMap, key)) return true;
  return Boolean(stateMap[key]);
}

function setHomeLineGroupExpanded(groupKey, expanded) {
  const key = String(groupKey || "").trim();
  if (!key) return;
  const stateMap = homeLineGroupExpandedState();
  stateMap[key] = Boolean(expanded);
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

function normalizeActionPriority(priority) {
  const value = String(priority || "").trim();
  return ACTION_PRIORITY_OPTIONS.includes(value) ? value : "Medium";
}

function normalizeActionStatus(status) {
  const value = String(status || "").trim();
  return ACTION_STATUS_OPTIONS.includes(value) ? value : "Open";
}

function normalizeActionReasonCategory(category) {
  const value = String(category || "").trim();
  return ACTION_REASON_CATEGORIES.includes(value) ? value : "";
}

function actionStatusSortRank(status) {
  const normalized = normalizeActionStatus(status);
  return Object.prototype.hasOwnProperty.call(ACTION_STATUS_SORT_ORDER, normalized)
    ? ACTION_STATUS_SORT_ORDER[normalized]
    : ACTION_STATUS_OPTIONS.length;
}

function setActionEquipmentOptions(selectNode, line, selectedValue = "") {
  if (!selectNode) return;
  const stages = line?.stages?.length ? line.stages : [];
  const options = stages.map((stage, index) => ({
    value: String(stage?.id || "").trim(),
    label: stageDisplayName(stage, index)
  }));
  selectNode.innerHTML = [
    `<option value="">Related Equipment (optional)</option>`,
    ...options.map((option) => `<option value="${htmlEscape(option.value)}">${htmlEscape(option.label)}</option>`)
  ].join("");
  if (selectedValue && options.some((option) => option.value === selectedValue)) {
    selectNode.value = selectedValue;
  } else {
    selectNode.value = "";
  }
}

function setActionReasonCategoryOptions(selectNode, selectedValue = "") {
  if (!selectNode) return;
  const safeSelected = normalizeActionReasonCategory(selectedValue);
  selectNode.innerHTML = [
    `<option value="">Downtime Category (optional)</option>`,
    ...ACTION_REASON_CATEGORIES.map((category) => `<option value="${htmlEscape(category)}">${htmlEscape(category)}</option>`)
  ].join("");
  selectNode.value = safeSelected;
}

function setActionReasonDetailOptions(selectNode, line, category, selectedValue = "") {
  if (!selectNode) return;
  const safeCategory = normalizeActionReasonCategory(category);
  const options = safeCategory ? downtimeDetailOptions(line, safeCategory) : [];
  const placeholder = safeCategory ? (safeCategory === "Equipment" ? "Downtime Reason / Stage" : "Downtime Reason") : "Downtime Reason (optional)";
  selectNode.innerHTML = [
    `<option value="">${htmlEscape(placeholder)}</option>`,
    ...options.map((option) => `<option value="${htmlEscape(option.value)}">${htmlEscape(option.label)}</option>`)
  ].join("");
  selectNode.disabled = !safeCategory;
  if (selectedValue && options.some((option) => option.value === selectedValue)) {
    selectNode.value = selectedValue;
  } else {
    selectNode.value = "";
  }
}

function actionRelationSummary(action, line) {
  const equipmentId = String(action?.relatedEquipmentId || "").trim();
  const reasonCategory = normalizeActionReasonCategory(action?.relatedReasonCategory);
  const reasonDetail = String(action?.relatedReasonDetail || "").trim();
  const equipmentLabel = equipmentId ? stageNameByIdForLine(line, equipmentId) || equipmentId : "";
  const reasonDetailLabel = reasonCategory ? downtimeDetailLabel(line, reasonCategory, reasonDetail) : "";
  const parts = [];
  if (equipmentLabel) parts.push(`Equipment: ${equipmentLabel}`);
  if (reasonCategory) {
    const reasonText = [reasonCategory, reasonDetailLabel].filter(Boolean).join(" > ");
    if (reasonText) parts.push(`Downtime: ${reasonText}`);
  }
  return parts.join(" | ");
}

function normalizeSupervisorAction(action) {
  if (!action || typeof action !== "object") return null;
  const supervisorUsername = String(action.supervisorUsername || "").trim().toLowerCase();
  if (!supervisorUsername) return null;
  const id = String(action.id || "").trim();
  if (!UUID_RE.test(id)) return null;
  const dueDateRaw = String(action.dueDate || "").trim();
  const relatedReasonCategory = normalizeActionReasonCategory(action.relatedReasonCategory);
  const relatedReasonDetailRaw = String(action.relatedReasonDetail || "").trim();
  const relatedReasonDetail = relatedReasonCategory && relatedReasonDetailRaw ? relatedReasonDetailRaw : "";
  return {
    id,
    supervisorUsername,
    supervisorName: String(action.supervisorName || supervisorUsername).trim() || supervisorUsername,
    lineId: String(action.lineId || "").trim(),
    title: String(action.title || "Untitled action").trim() || "Untitled action",
    description: String(action.description || action.notes || "").trim(),
    priority: normalizeActionPriority(action.priority),
    status: normalizeActionStatus(action.status),
    dueDate: /^\d{4}-\d{2}-\d{2}$/.test(dueDateRaw) ? dueDateRaw : "",
    relatedEquipmentId: String(action.relatedEquipmentId || "").trim(),
    relatedReasonCategory,
    relatedReasonDetail,
    createdAt: String(action.createdAt || "").trim() || nowIso(),
    createdBy: String(action.createdBy || "System").trim() || "System"
  };
}

function normalizeSupervisorActions(actions) {
  const source = Array.isArray(actions) ? actions : [];
  const seen = new Set();
  return source
    .map((action) => normalizeSupervisorAction(action))
    .filter((action) => {
      if (!action) return false;
      if (seen.has(action.id)) return false;
      seen.add(action.id);
      return true;
    });
}

function ensureSupervisorActionsState() {
  appState.supervisorActions = normalizeSupervisorActions(appState.supervisorActions);
  return appState.supervisorActions;
}

function supervisorActionsForUsername(username) {
  const key = String(username || "").trim().toLowerCase();
  if (!key) return [];
  return ensureSupervisorActionsState()
    .filter((action) => action.supervisorUsername === key)
    .slice()
    .sort((a, b) => String(b.createdAt || "").localeCompare(String(a.createdAt || "")));
}

function actionTicketToneKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "") || "none";
}

function renderSupervisorActionTickets(actions, { emptyMessage = "No actions yet. Use Lodge Action to add one." } = {}) {
  const rows = Array.isArray(actions) ? actions : [];
  if (!rows.length) return `<p class="muted pending-log-empty">${htmlEscape(emptyMessage)}</p>`;
  return `
    <div class="pending-log-wrap supervisor-action-ticket-wrap">
      <div class="pending-log-list supervisor-action-ticket-list">
        ${rows
          .map((action) => {
            const lineName = action.lineId && appState.lines[action.lineId] ? appState.lines[action.lineId].name : "Unassigned";
            const line = action.lineId && appState.lines[action.lineId] ? appState.lines[action.lineId] : null;
            const createdLabel = String(action.createdAt || "").replace("T", " ").slice(0, 16) || "-";
            const createdBy = String(action.createdBy || action.supervisorName || "System").trim() || "System";
            const dueLabel = String(action.dueDate || "").trim() || "No due date";
            const description = String(action.description || "").trim() || "No description provided.";
            const relationText = actionRelationSummary(action, line);
            const title = String(action.title || "").trim() || "Untitled action";
            const priority = normalizeActionPriority(action.priority);
            const status = normalizeActionStatus(action.status);
            const ticketIdRaw = String(action.id || "").trim();
            const ticketId = ticketIdRaw ? ticketIdRaw.slice(-6).toUpperCase() : "N/A";
            return `
              <article class="pending-log-item supervisor-action-ticket">
                <div class="pending-log-meta">
                  <h5>
                    ${htmlEscape(title)}
                    <span class="supervisor-action-ticket-id">#${htmlEscape(ticketId)}</span>
                  </h5>
                  <p class="supervisor-action-ticket-line">${htmlEscape(lineName)} | Due ${htmlEscape(dueLabel)} | Created ${htmlEscape(createdLabel)} by ${htmlEscape(createdBy)}</p>
                  ${relationText ? `<p class="supervisor-action-ticket-relation">Related: ${htmlEscape(relationText)}</p>` : ""}
                  <p class="supervisor-action-ticket-description">${htmlEscape(description)}</p>
                </div>
                <div class="pending-log-actions supervisor-action-ticket-badges">
                  <span class="supervisor-action-ticket-badge priority-${actionTicketToneKey(priority)}">${htmlEscape(priority)}</span>
                  <span class="supervisor-action-ticket-badge status-${actionTicketToneKey(status)}">${htmlEscape(status)}</span>
                </div>
              </article>
            `;
          })
          .join("")}
      </div>
    </div>
  `;
}

function actionAssignmentLabel(username, fallbackLabel = "") {
  const key = String(username || "").trim().toLowerCase();
  if (!key) return "Unassigned";
  if (ACTION_SPECIAL_ASSIGNMENT_LABELS[key]) return ACTION_SPECIAL_ASSIGNMENT_LABELS[key];
  const supervisor = supervisorByUsername(key);
  if (supervisor?.name) return String(supervisor.name).trim() || key;
  const fallback = String(fallbackLabel || "").trim();
  return fallback || key;
}

function actionDueDayDelta(isoDueDate, referenceIsoDate = todayISO()) {
  const dueDate = String(isoDueDate || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dueDate)) return null;
  const baseRaw = String(referenceIsoDate || "").trim();
  const baseDate = /^\d{4}-\d{2}-\d{2}$/.test(baseRaw) ? baseRaw : todayISO();
  const due = parseDateLocal(dueDate);
  const base = parseDateLocal(baseDate);
  due.setHours(0, 0, 0, 0);
  base.setHours(0, 0, 0, 0);
  return Math.round((due.getTime() - base.getTime()) / 86400000);
}

function managerActionUrgencyMeta(action, referenceIsoDate = todayISO()) {
  const priority = normalizeActionPriority(action?.priority);
  const status = normalizeActionStatus(action?.status);
  const priorityScore = num(ACTION_PRIORITY_SCORES[priority]);
  const dueDays = actionDueDayDelta(action?.dueDate, referenceIsoDate);
  const hasDueDate = Number.isFinite(dueDays);
  const isCompleted = status === "Completed";
  const isBlocked = status === "Blocked";
  let urgencyScore = priorityScore;
  if (isBlocked) urgencyScore += 2;
  if (!isCompleted) {
    if (hasDueDate) {
      if (dueDays < 0) urgencyScore += 3;
      else if (dueDays === 0) urgencyScore += 2;
      else if (dueDays <= 2) urgencyScore += 1;
      else if (dueDays >= 7) urgencyScore -= 1;
    } else if (priority === "Critical") {
      urgencyScore += 1;
    }
  } else {
    urgencyScore = 0;
  }
  urgencyScore = Math.max(0, Math.min(10, urgencyScore));

  let groupKey = "planned";
  if (isCompleted) groupKey = "completed";
  else if (hasDueDate && dueDays < 0) groupKey = "overdue";
  else if (isBlocked || priority === "Critical" || (hasDueDate && dueDays === 0)) groupKey = "urgent";
  else if (hasDueDate && dueDays <= 3) groupKey = "due-soon";
  else if (!hasDueDate) groupKey = "unscheduled";
  const group = MANAGER_ACTION_URGENCY_GROUP_MAP[groupKey] || MANAGER_ACTION_URGENCY_GROUP_MAP.planned;

  let urgencyLabel = "Low";
  let urgencyTone = "low";
  if (isCompleted) {
    urgencyLabel = "Complete";
    urgencyTone = "done";
  } else if (hasDueDate && dueDays < 0) {
    urgencyLabel = `${Math.abs(dueDays)}d Overdue`;
    urgencyTone = "critical";
  } else if (isBlocked) {
    urgencyLabel = "Blocked";
    urgencyTone = "critical";
  } else if (dueDays === 0) {
    urgencyLabel = "Due Today";
    urgencyTone = "high";
  } else if (urgencyScore >= 7) {
    urgencyLabel = "Critical";
    urgencyTone = "critical";
  } else if (urgencyScore >= 5) {
    urgencyLabel = "High";
    urgencyTone = "high";
  } else if (urgencyScore >= 3) {
    urgencyLabel = "Medium";
    urgencyTone = "medium";
  }

  let dueHint = "No due date";
  if (isCompleted) dueHint = "Completed";
  else if (hasDueDate && dueDays < 0) dueHint = `Overdue by ${Math.abs(dueDays)} day${Math.abs(dueDays) === 1 ? "" : "s"}`;
  else if (hasDueDate && dueDays === 0) dueHint = "Due today";
  else if (hasDueDate && dueDays === 1) dueHint = "Due tomorrow";
  else if (hasDueDate) dueHint = `Due in ${dueDays} days`;

  return {
    dueDays,
    urgencyScore,
    urgencyLabel,
    urgencyTone,
    dueHint,
    groupKey: group.key,
    groupLabel: group.label,
    groupTone: group.tone,
    groupOrder: num(group.order)
  };
}

function compareManagerActionItems(left, right) {
  const leftMeta = left?.urgency || {};
  const rightMeta = right?.urgency || {};
  const groupDelta = num(leftMeta.groupOrder) - num(rightMeta.groupOrder);
  if (groupDelta !== 0) return groupDelta;
  const urgencyDelta = num(rightMeta.urgencyScore) - num(leftMeta.urgencyScore);
  if (urgencyDelta !== 0) return urgencyDelta;
  const leftDue = Number.isFinite(leftMeta.dueDays) ? leftMeta.dueDays : Number.POSITIVE_INFINITY;
  const rightDue = Number.isFinite(rightMeta.dueDays) ? rightMeta.dueDays : Number.POSITIVE_INFINITY;
  if (leftDue !== rightDue) return leftDue - rightDue;
  const createdDelta = String(right?.action?.createdAt || "").localeCompare(String(left?.action?.createdAt || ""));
  if (createdDelta !== 0) return createdDelta;
  return String(left?.action?.title || "").localeCompare(String(right?.action?.title || ""), undefined, { sensitivity: "base" });
}

function groupManagerActionsForDisplay(actions, referenceIsoDate = todayISO()) {
  const baseRows = Array.isArray(actions) ? actions : [];
  const items = baseRows
    .map((action) => ({
      action,
      urgency: managerActionUrgencyMeta(action, referenceIsoDate)
    }))
    .sort(compareManagerActionItems);
  const groups = MANAGER_ACTION_URGENCY_GROUPS.map((group) => ({
    ...group,
    actions: []
  }));
  const groupMap = groups.reduce((acc, group) => {
    acc[group.key] = group;
    return acc;
  }, {});
  items.forEach((item) => {
    const key = item?.urgency?.groupKey || "planned";
    const target = groupMap[key] || groupMap.planned;
    if (target) target.actions.push(item);
  });
  return groups.filter((group) => group.actions.length);
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
  const supervisorName = String(sup?.name || session?.name || session?.username || "").trim();
  return {
    name: supervisorName || String(sup?.username || session.username || "").trim().toLowerCase(),
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

function resolveLineBottleneckStage(line, stages, shift, capacityForStage) {
  const stageList = Array.isArray(stages) ? stages : [];
  const candidates = stageList
    .map((stage, index) => ({
      stage,
      index,
      capacity: Math.max(0, num(typeof capacityForStage === "function" ? capacityForStage(stage.id, shift) : 0))
    }))
    .filter((entry) => entry.capacity > 0 && isCrewedStage(entry.stage));

  if (!candidates.length) return null;

  const explicitId = String(line?.bottleneckStageId || "").trim();
  if (explicitId) {
    const explicit = candidates.find((entry) => entry.stage.id === explicitId);
    if (explicit) return explicit;
  }

  const proseal = candidates.find((entry) => {
    const name = String(entry.stage?.name || "").toLowerCase();
    const matches = Array.isArray(entry.stage?.match) ? entry.stage.match.map((token) => String(token || "").toLowerCase()) : [];
    return entry.stage.id === "s10" || name.includes("proseal") || matches.includes("proseal");
  });
  if (proseal) return proseal;

  return candidates.reduce((best, entry) => (entry.capacity < best.capacity ? entry : best), candidates[0]);
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
  const operationalDowntimeRows = (line?.downtimeRows || []).filter((row) => isOperationalDate(String(row?.date || "")));
  const operationalShiftRows = (line?.shiftRows || []).filter((row) => isOperationalDate(String(row?.date || "")));
  const operationalBreakRows = (line?.breakRows || []).filter((row) => isOperationalDate(String(row?.date || "")));
  const operationalRunRows = (line?.runRows || []).filter((row) => isOperationalDate(String(row?.date || "")));
  const computedDowntimeRows = operationalDowntimeRows
    .map(computeDowntimeRow)
    .map((row) => decorateTimedLogShift(row, line, "downtimeStart", "downtimeFinish"));
  const { downtimeRowsLogged, downtimeRows, breakRowsFromDowntime } = splitDowntimeRowsForBreakViews(computedDowntimeRows, line);
  const shiftRows = operationalShiftRows.map(computeShiftRow);
  const breakRows = [...operationalBreakRows.map(computeBreakRow), ...breakRowsFromDowntime.map(computeBreakRow)];
  const nonProductionIntervalsByDate = buildNonProductionIntervalsByDate(breakRows, calculationDowntimeRows(downtimeRows));
  const runRows = operationalRunRows
    .map((row) => decorateTimedLogShift(row, line, "productionStartTime", "finishTime"))
    .map((row) => computeRunRow(row, nonProductionIntervalsByDate));
  return { shiftRows, breakRows, runRows, downtimeRows, downtimeRowsLogged };
}

function computeLineMetricsFromData(line, date, shift, data = derivedDataForLine(line || {})) {
  const stages = line?.stages?.length ? line.stages : STAGES;
  const selectedRunRows = selectedShiftRowsByDate(data.runRows, date, shift, { line });
  const selectedDownRows = selectedShiftRowsByDate(data.downtimeRows, date, shift, { line });
  const selectedCalcDownRows = calculationDowntimeRows(selectedDownRows);
  const selectedShiftRows = selectedShiftRowsByDate(data.shiftRows, date, shift, { line });
  const shiftMins = selectedShiftRows.reduce((sum, row) => sum + num(row.totalShiftTime), 0);

  const staffing = staffingSnapshotForSelection(line, selectedRunRows, shift);
  const units = selectedRunRows.reduce((sum, row) => sum + num(row.unitsProduced) * timedLogShiftWeight(row, shift), 0);
  const totalDowntime = selectedDownRows.reduce((sum, row) => sum + num(row.downtimeMins) * timedLogShiftWeight(row, shift), 0);
  const totalNetTime = selectedRunRows.reduce((sum, row) => sum + num(row.netProductionTime) * timedLogShiftWeight(row, shift), 0);
  const netRunRate = totalNetTime > 0 ? units / totalNetTime : 0;
  const bottleneck = resolveLineBottleneckStage(line, stages, shift, (stageId, selectedShift) =>
    stageTotalMaxThroughputForLine(line, stageId, selectedShift)
  );
  let bottleneckStageName = "-";
  let bottleneckUtil = 0;
  let bottleneckUtilGross = 0;
  if (bottleneck) {
    const stageDowntime = selectedCalcDownRows
      .filter((row) => matchesStage(bottleneck.stage, row.equipment))
      .reduce((sum, row) => sum + downtimeMinutesForCalculations(row, shift), 0);
    const uptimeRatio = shiftMins > 0 ? Math.max(0, (shiftMins - stageDowntime) / shiftMins) : 0;
    const stageRate = netRunRate * uptimeRatio;
    const totalMax = Math.max(0, num(bottleneck.capacity));
    bottleneckUtil = totalMax > 0 ? (stageRate / totalMax) * 100 : 0;
    bottleneckUtilGross = totalMax > 0 ? (netRunRate / totalMax) * 100 : 0;
    bottleneckStageName = stageDisplayName(bottleneck.stage, bottleneck.index);
  }

  return {
    lineName: line?.name || "Line",
    date,
    shift,
    units,
    totalDowntime,
    lineUtil: Math.max(0, bottleneckUtilGross),
    lineUtilGross: Math.max(0, bottleneckUtil),
    netRunRate,
    bottleneckStageName,
    requiredCrew: staffing.requiredCrew,
    actualCrew: staffing.actualCrew,
    understaffedBy: staffing.understaffedBy,
    overstaffedBy: staffing.overstaffedBy,
    staffingCallout: staffing.staffingCallout
  };
}

function computeLineMetrics(line, date, shift) {
  return computeLineMetricsFromData(line, date, shift, derivedDataForLine(line || {}));
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

function productCatalogCellText(row, columnIndex) {
  const cell = row?.cells?.[columnIndex];
  if (!cell) return "";
  const inlineInput = cell.querySelector("input[data-product-inline-input]");
  if (inlineInput) return String(inlineInput.value || "").trim();
  return String(cell.textContent || "").trim();
}

function createProductCatalogRowId() {
  return `prod-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function normalizeProductCatalogLineIds(lineIds) {
  const source = Array.isArray(lineIds) ? lineIds : String(lineIds || "").split(",");
  const normalized = [];
  let allowAll = false;
  source.forEach((rawId) => {
    const lineId = String(rawId || "").trim();
    if (!lineId) return;
    if (lineId === PRODUCT_CATALOG_ALL_LINES_TOKEN) {
      allowAll = true;
      return;
    }
    if (!normalized.includes(lineId)) normalized.push(lineId);
  });
  if (allowAll) return [PRODUCT_CATALOG_ALL_LINES_TOKEN];
  return normalized;
}

function ensureProductCatalogRowId(row) {
  if (!row) return "";
  let rowId = String(row.getAttribute(PRODUCT_CATALOG_ID_ATTR) || "").trim();
  if (!rowId) {
    rowId = createProductCatalogRowId();
    row.setAttribute(PRODUCT_CATALOG_ID_ATTR, rowId);
  }
  return rowId;
}

function getProductCatalogLineIdsFromRow(row) {
  if (!row) return [];
  if (!row.hasAttribute(PRODUCT_CATALOG_LINES_ATTR)) return [PRODUCT_CATALOG_ALL_LINES_TOKEN];
  const raw = String(row.getAttribute(PRODUCT_CATALOG_LINES_ATTR) || "").trim();
  if (!raw) return [];
  return normalizeProductCatalogLineIds(raw);
}

function setProductCatalogLineIdsForRow(row, lineIds) {
  if (!row) return;
  const normalized = normalizeProductCatalogLineIds(lineIds);
  row.setAttribute(PRODUCT_CATALOG_LINES_ATTR, normalized.join(","));
}

function productCatalogMatchesLine(entryLineIds, lineId) {
  const targetLineId = String(lineId || "").trim();
  if (!targetLineId) return true;
  const normalized = normalizeProductCatalogLineIds(entryLineIds);
  if (normalized.includes(PRODUCT_CATALOG_ALL_LINES_TOKEN)) return true;
  return normalized.includes(targetLineId);
}

function listProductCatalogEntries(options = {}) {
  const table = document.getElementById("productCatalogTable");
  const rows = Array.from(table?.tBodies?.[0]?.rows || []);
  const filterLineId = String(options?.lineId || "").trim();
  return rows
    .map((row) => {
      const values = Array.from({ length: PRODUCT_CATALOG_COLUMN_COUNT }, (_, i) => productCatalogCellText(row, i));
      return {
        id: ensureProductCatalogRowId(row),
        row,
        values,
        code: values[0],
        desc1: values[1],
        desc2: values[2],
        lineIds: getProductCatalogLineIdsFromRow(row)
      };
    })
    .filter((entry) => (entry.code || entry.desc1 || entry.desc2) && productCatalogMatchesLine(entry.lineIds, filterLineId));
}

function listProductCatalogNames(options = {}) {
  const seen = new Set();
  const names = [];
  listProductCatalogEntries(options).forEach((entry) => {
    const value = String(entry.desc1 || entry.code || "").trim();
    if (!value) return;
    const key = value.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    names.push(value);
  });
  return names;
}

function catalogProductCanonicalName(value, options = {}) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const target = raw.toLowerCase();
  const entries = listProductCatalogEntries(options);
  let canonicalFromCode = "";
  for (const entry of entries) {
    const code = String(entry.code || "").trim();
    const canonicalName = String(entry.desc1 || entry.code || "").trim();
    if (!canonicalName) continue;
    if (target === canonicalName.toLowerCase()) return canonicalName;
    if (code && target === code.toLowerCase()) canonicalFromCode = canonicalName;
    if (!code) continue;
    if (target === `${code} - ${canonicalName}`.toLowerCase()) return canonicalName;
    if (target === `${code} | ${canonicalName}`.toLowerCase()) return canonicalName;
  }
  if (canonicalFromCode) return canonicalFromCode;
  const composite = raw.match(/^([A-Za-z0-9._/-]+)\s*[-|:]\s*(.+)$/);
  if (composite) {
    const inputCode = String(composite[1] || "").trim().toLowerCase();
    const codeMatch = entries.find((entry) => String(entry.code || "").trim().toLowerCase() === inputCode);
    if (codeMatch) return String(codeMatch.desc1 || codeMatch.code || "").trim();
  }
  return raw;
}

function isCatalogProductName(value, options = {}) {
  const target = catalogProductCanonicalName(value, options).toLowerCase();
  if (!target) return false;
  return listProductCatalogNames(options).some((name) => name.toLowerCase() === target);
}

function normalizeProductCatalogEntry(entry) {
  if (!entry || typeof entry !== "object") return null;
  const rowId = String(entry.id || "").trim() || createProductCatalogRowId();
  const valuesSource = Array.isArray(entry.values) ? entry.values : [];
  const values = Array.from({ length: PRODUCT_CATALOG_COLUMN_COUNT }, (_, index) => {
    const fallback = index === 0 ? entry.code : index === 1 ? entry.desc1 : index === 2 ? entry.desc2 : "";
    return String(valuesSource[index] ?? fallback ?? "").trim();
  });
  if (!values[0] && !values[1] && !values[2]) return null;
  return {
    id: rowId,
    values,
    lineIds: normalizeProductCatalogLineIds(entry.lineIds)
  };
}

function normalizeProductCatalogEntries(entries) {
  const source = Array.isArray(entries) ? entries : [];
  const seen = new Set();
  return source
    .map((entry) => normalizeProductCatalogEntry(entry))
    .filter((entry) => {
      if (!entry) return false;
      const key = String(entry.id || "").trim();
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function setProductCatalogEntries(entries) {
  appState.productCatalogEntries = normalizeProductCatalogEntries(entries);
}

function serializeProductCatalogTable(table) {
  const rows = Array.from(table?.tBodies?.[0]?.rows || []);
  return rows
    .map((row) => {
      const values = Array.from({ length: PRODUCT_CATALOG_COLUMN_COUNT }, (_, i) => productCatalogCellText(row, i));
      if (!values[0] && !values[1] && !values[2]) return null;
      return normalizeProductCatalogEntry({
        id: ensureProductCatalogRowId(row),
        values,
        lineIds: getProductCatalogLineIdsFromRow(row)
      });
    })
    .filter(Boolean);
}

function syncProductCatalogStateFromTable(table) {
  setProductCatalogEntries(serializeProductCatalogTable(table));
}

function renderProductCatalogTableRows(table, entries) {
  if (!table) return;
  const tbody = table.tBodies?.[0] || table.createTBody();
  const safeEntries = normalizeProductCatalogEntries(entries);
  tbody.innerHTML = safeEntries
    .map((entry) => {
      const values = Array.from({ length: PRODUCT_CATALOG_COLUMN_COUNT }, (_, index) => String(entry.values?.[index] || "").trim());
      const lineIds = normalizeProductCatalogLineIds(entry.lineIds);
      return `
        <tr ${PRODUCT_CATALOG_ID_ATTR}="${htmlEscape(entry.id)}" ${PRODUCT_CATALOG_LINES_ATTR}="${htmlEscape(lineIds.join(","))}">
          ${values.map((value) => `<td>${htmlEscape(value)}</td>`).join("")}
        </tr>
      `;
    })
    .join("");
}

function hydrateProductCatalogTableFromState(table) {
  if (!table) return;
  renderProductCatalogTableRows(table, appState.productCatalogEntries);
}

function productCatalogPayload(values, lineIds) {
  return {
    values: Array.from({ length: PRODUCT_CATALOG_COLUMN_COUNT }, (_, index) => String(values?.[index] || "").trim()),
    lineIds: normalizeProductCatalogLineIds(lineIds)
  };
}

async function createProductCatalogEntryOnBackend(values, lineIds = []) {
  const session = await ensureManagerBackendSession();
  const payload = productCatalogPayload(values, lineIds);
  const response = await apiRequest("/api/product-catalog", {
    method: "POST",
    token: session.backendToken,
    body: payload
  });
  const product = normalizeProductCatalogEntry(response?.product);
  if (!product || !UUID_RE.test(String(product.id || "").trim())) {
    throw new Error("Server returned invalid product catalog data.");
  }
  return product;
}

async function updateProductCatalogEntryOnBackend(productId, values, lineIds = []) {
  if (!UUID_RE.test(String(productId || "").trim())) throw new Error("Product is not synced to server.");
  const session = await ensureManagerBackendSession();
  const payload = productCatalogPayload(values, lineIds);
  const response = await apiRequest(`/api/product-catalog/${encodeURIComponent(String(productId || "").trim())}`, {
    method: "PATCH",
    token: session.backendToken,
    body: payload
  });
  const product = normalizeProductCatalogEntry(response?.product);
  if (!product || !UUID_RE.test(String(product.id || "").trim())) {
    throw new Error("Server returned invalid product catalog data.");
  }
  return product;
}

function ensureRunProductDatalistById(datalistId) {
  let datalist = document.getElementById(datalistId);
  if (!datalist) {
    datalist = document.createElement("datalist");
    datalist.id = datalistId;
    document.body.append(datalist);
  }
  return datalist;
}

function listProductCatalogOptions(options = {}) {
  const seen = new Set();
  const deduped = [];
  const includeCodeInValue = Boolean(options?.includeCodeInValue);
  listProductCatalogEntries(options).forEach((entry) => {
    const canonicalValue = String(entry.desc1 || entry.code || "").trim();
    if (!canonicalValue) return;
    const key = canonicalValue.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    const code = String(entry.code || "").trim();
    const value = includeCodeInValue && code && code.toLowerCase() !== canonicalValue.toLowerCase()
      ? `${code} - ${canonicalValue}`
      : canonicalValue;
    const displayLabel = code && code.toLowerCase() !== canonicalValue.toLowerCase()
      ? `${code} - ${canonicalValue}`
      : canonicalValue;
    deduped.push({
      value,
      canonicalValue,
      code,
      label: [entry.code, entry.desc2].filter(Boolean).join(" | "),
      displayLabel
    });
  });
  return deduped;
}

function selectedSupervisorProductCatalogLineId() {
  const lineSelect = document.getElementById("supervisorLineSelect");
  return String(lineSelect?.value || appState.supervisorSelectedLineId || "").trim();
}

function syncRunProductInputsFromCatalog() {
  const managerInput = document.querySelector('#runForm [name="product"]');
  if (managerInput) {
    const managerDatalist = ensureRunProductDatalistById(PRODUCT_CATALOG_MANAGER_DATALIST_ID);
    managerDatalist.innerHTML = listProductCatalogOptions()
      .map((item) => `<option value="${htmlEscape(item.value)}">${htmlEscape(item.label || item.value)}</option>`)
      .join("");
    managerInput.setAttribute("list", PRODUCT_CATALOG_MANAGER_DATALIST_ID);
    managerInput.setAttribute("autocomplete", "off");
  }

  const supervisorInput = document.getElementById("superRunProduct");
  if (supervisorInput) {
    const supervisorLineId = selectedSupervisorProductCatalogLineId();
    const supervisorInputTag = String(supervisorInput.tagName || "").toUpperCase();
    if (supervisorInputTag === "SELECT") {
      const selectedValue = String(supervisorInput.value || "").trim();
      const supervisorOptions = supervisorLineId
        ? listProductCatalogOptions({ lineId: supervisorLineId, includeCodeInValue: false })
        : [];
      const selectedOption = selectedValue
        ? supervisorOptions.find((item) => String(item.canonicalValue || item.value || "").trim().toLowerCase() === selectedValue.toLowerCase())
        : null;
      const selectedOptionHtml =
        selectedValue && !selectedOption
          ? `<option value="${htmlEscape(selectedValue)}">${htmlEscape(selectedValue)}</option>`
          : "";
      supervisorInput.innerHTML = [
        `<option value="">Product Code</option>`,
        selectedOptionHtml,
        ...supervisorOptions.map((item) => {
          const value = String(item.canonicalValue || item.value || "").trim();
          const text = String(item.displayLabel || item.label || value).trim() || value;
          return `<option value="${htmlEscape(value)}">${htmlEscape(text)}</option>`;
        })
      ].join("");
      supervisorInput.value = selectedOption
        ? String(selectedOption.canonicalValue || selectedOption.value || "").trim()
        : selectedOptionHtml
          ? selectedValue
          : "";
      supervisorInput.removeAttribute("list");
      supervisorInput.removeAttribute("autocomplete");
    } else {
      const supervisorDatalist = ensureRunProductDatalistById(PRODUCT_CATALOG_SUPERVISOR_DATALIST_ID);
      const supervisorOptions = supervisorLineId
        ? listProductCatalogOptions({ lineId: supervisorLineId, includeCodeInValue: true })
        : [];
      supervisorDatalist.innerHTML = supervisorOptions
        .map((item) => `<option value="${htmlEscape(item.value)}">${htmlEscape(item.displayLabel || item.label || item.value)}</option>`)
        .join("");
      supervisorInput.setAttribute("list", PRODUCT_CATALOG_SUPERVISOR_DATALIST_ID);
      supervisorInput.setAttribute("autocomplete", "off");
    }
  }
}

function setProductRowAssignmentSummary(rowNode) {
  if (!rowNode) return;
  const summaryNode = rowNode.querySelector("[data-product-row-line-summary]");
  if (!summaryNode) return;
  const assignedLineIds = getProductCatalogLineIdsFromRow(rowNode);
  if (assignedLineIds.includes(PRODUCT_CATALOG_ALL_LINES_TOKEN)) {
    summaryNode.textContent = "All lines";
    summaryNode.title = "Assigned to all lines";
    return;
  }
  if (!assignedLineIds.length) {
    summaryNode.textContent = "No lines";
    summaryNode.title = "No line assignment";
    return;
  }
  const names = assignedLineIds.map((lineId) => String(appState.lines?.[lineId]?.name || "").trim()).filter(Boolean);
  if (names.length === 1) {
    summaryNode.textContent = names[0];
    summaryNode.title = names[0];
    return;
  }
  const count = names.length || assignedLineIds.length;
  summaryNode.textContent = `${count} lines`;
  summaryNode.title = names.length ? names.join(", ") : `${count} line assignments`;
}

function refreshProductRowAssignmentSummaries(table) {
  Array.from(table?.tBodies?.[0]?.rows || []).forEach((rowNode) => {
    setProductRowAssignmentSummary(rowNode);
  });
}

function ensureProductCatalogActionColumn(table) {
  if (!table) return;
  const headerRow = table.tHead?.rows?.[0];
  if (headerRow && !headerRow.querySelector("[data-product-action-head]")) {
    const headerCell = document.createElement("th");
    headerCell.textContent = "Action";
    headerCell.setAttribute("data-product-action-head", "true");
    headerRow.append(headerCell);
  }
  const bodyRows = Array.from(table.tBodies?.[0]?.rows || []);
  bodyRows.forEach((rowNode) => {
    ensureProductCatalogRowId(rowNode);
    setProductCatalogLineIdsForRow(rowNode, getProductCatalogLineIdsFromRow(rowNode));
    if (!rowNode.getAttribute("data-product-editing")) rowNode.setAttribute("data-product-editing", "false");
    let actionCell = rowNode.querySelector("[data-product-action-cell]");
    if (!actionCell) {
      actionCell = document.createElement("td");
      actionCell.setAttribute("data-product-action-cell", "true");
      rowNode.append(actionCell);
    }
    if (!actionCell.querySelector("[data-product-row-edit]")) {
      actionCell.innerHTML = `
        <div class="table-action-stack">
          <button type="button" class="table-edit-pill" data-product-row-edit="true">Edit</button>
          <button type="button" class="table-edit-pill" data-product-row-lines="true">Lines</button>
        </div>
        <div class="product-row-line-summary" data-product-row-line-summary></div>
      `;
    }
    setProductRowAssignmentSummary(rowNode);
  });
}

function loadState() {
  const baselineDate = normalizeWeekdayIsoDate(todayISO(), { direction: -1 });
  return {
    activeView: "home",
    appMode: "manager",
    managerHomeTab: "dashboard",
    dashboardDate: baselineDate,
    dashboardShift: "Day",
    managerLineTab: "visualiser",
    supervisorMobileMode: APP_VARIANT === "supervisor",
    supervisorMainTab: "supervisorDay",
    supervisorTab: "superShift",
    supervisorSelectedLineId: "",
    supervisorSelectedDate: baselineDate,
    supervisorSelectedShift: "Full Day",
    supervisorSession: null,
    supervisors: [],
    dataSources: [],
    supervisorActions: [],
    productCatalogEntries: [],
    lineGroups: [],
    homeLineGroupExpanded: {},
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
  const managerName = String(managerBackendSession?.name || "").trim();
  const managerUsername = String(managerBackendSession?.username || "").trim().toLowerCase();
  const supervisorSession = appState?.supervisorSession;
  const supervisorToken = String(supervisorSession?.backendToken || "").trim();
  const supervisorUsername = String(supervisorSession?.username || "").trim().toLowerCase();
  if (!managerToken && !(supervisorToken && supervisorUsername)) {
    writeAuthStorage(null);
    return;
  }
  writeAuthStorage({
    manager:
      managerToken
        ? {
            backendToken: managerToken,
            name: managerName,
            username: managerUsername
          }
        : null,
    supervisor:
      supervisorToken && supervisorUsername
        ? {
            name: String(supervisorSession?.name || "").trim(),
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
  const managerName = String(stored?.manager?.name || "").trim();
  const managerUsername = String(stored?.manager?.username || "").trim().toLowerCase();
  if (managerToken) {
    managerBackendSession.backendToken = managerToken;
    managerBackendSession.role = "manager";
    managerBackendSession.name = managerName;
    managerBackendSession.username = managerUsername;
  }
  const supervisorToken = String(stored?.supervisor?.backendToken || "").trim();
  const supervisorUsername = String(stored?.supervisor?.username || "").trim().toLowerCase();
  const supervisorName = String(stored?.supervisor?.name || "").trim();
  if (!supervisorToken || !supervisorUsername) return;
  const supervisorSession = normalizeSupervisorSession(
    {
      name: supervisorName,
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
    name: supervisorName || supervisorUsername,
    username: supervisorUsername,
    assignedLineIds: [],
    assignedLineShifts: {},
    backendToken: supervisorToken,
    backendLineMap: {},
    role: "supervisor"
  };
}

function queueLineModelSync(lineId, { delayMs = 300 } = {}) {
  const safeLineId = String(lineId || "").trim();
  if (!safeLineId) return;
  if (lineModelSyncTimer) clearTimeout(lineModelSyncTimer);
  const runSync = async () => {
    lineModelSyncTimer = null;
    try {
      await saveLineModelToBackend(safeLineId);
    } catch (error) {
      console.warn("Backend line model sync failed:", error);
      if (document.visibilityState === "visible") {
        alert(`Could not save layout/settings changes to server.\n${error?.message || "Please retry."}`);
      }
    }
  };
  if (Number(delayMs) <= 0) {
    runSync();
    return;
  }
  lineModelSyncTimer = setTimeout(runSync, Number(delayMs));
}

function requestLineModelSync(lineId, { force = false, immediate = false } = {}) {
  const safeLineId = String(lineId || "");
  if (!safeLineId) return;
  const line = appState.lines?.[safeLineId];
  const editingLayout = Boolean(line?.visualEditMode);
  if (editingLayout && !force) {
    deferredLineModelSyncIds.add(safeLineId);
    return;
  }
  deferredLineModelSyncIds.delete(safeLineId);
  queueLineModelSync(safeLineId, { delayMs: immediate ? 0 : 300 });
}

function saveState(options = {}) {
  enforceAppVariantState();
  if (state && state.id) {
    appState.lines[state.id] = state;
    appState.activeLineId = state.id;
    if (options.syncModel) {
      requestLineModelSync(state.id, {
        force: Boolean(options.forceSyncModel),
        immediate: Boolean(options.immediateSyncModel)
      });
    }
  }
  persistAuthSessions();
  persistHomeUiState();
  syncRouteToHash();
}

function makeDefaultLine(id, name, { seedSample = false } = {}) {
  const baselineDate = normalizeWeekdayIsoDate(todayISO(), { direction: -1 });
  const stages = clone(STAGES);
  const line = {
    id,
    name,
    groupId: "",
    displayOrder: 0,
    secretKey: generateSecretKey(),
    selectedDate: baselineDate,
    selectedShift: "Day",
    crewSettingsShift: "Day",
    visualEditMode: false,
    flowGuides: [],
    dayVisualiserKeyStageId: "",
    selectedStageId: STAGES[0].id,
    activeDataTab: "dataShift",
    trendGranularity: "daily",
    trendMonth: todayISO().slice(0, 7),
    trendDateCursor: baselineDate,
    trendShowZeroDays: true,
    lineTrendRange: "day",
    lineTrendLegendFocusKey: "",
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
  const displayOrderRaw = Number(line?.displayOrder);
  const stages = Array.isArray(line?.stages) && line.stages.length ? line.stages : clone(STAGES);
  const normalized = {
    ...base,
    ...line,
    id,
    name: line?.name || base.name,
    selectedDate: normalizeWeekdayIsoDate(line?.selectedDate || base.selectedDate, { direction: -1 }),
    groupId: String(line?.groupId || "").trim(),
    displayOrder: Number.isFinite(displayOrderRaw) ? Math.max(0, Math.floor(displayOrderRaw)) : base.displayOrder,
    secretKey: line?.secretKey || base.secretKey,
    crewSettingsShift: CREW_SETTINGS_SHIFTS.includes(String(line?.crewSettingsShift || ""))
      ? String(line.crewSettingsShift)
      : fallbackShiftValue(line?.selectedShift || base.selectedShift),
    visualEditMode: Boolean(line?.visualEditMode),
    flowGuides: normalizeFlowGuides(line?.flowGuides),
    dayVisualiserKeyStageId: "",
    trendDateCursor: normalizeWeekdayIsoDate(line?.trendDateCursor || line?.selectedDate || base.selectedDate, { direction: -1 }),
    trendShowZeroDays: line?.trendShowZeroDays !== false,
    lineTrendLegendFocusKey: String(line?.lineTrendLegendFocusKey || "").trim(),
    crewsByShift: normalizeCrewByShift(line || {}, stages),
    stageSettings: normalizeStageSettings(line || {}, stages),
    stages,
    shiftRows: line?.shiftRows || [],
    breakRows: line?.breakRows || [],
    runRows: Array.isArray(line?.runRows)
      ? line.runRows.map((row) => normalizeRunLogRow(row))
      : [],
    downtimeRows: Array.isArray(line?.downtimeRows)
      ? line.downtimeRows.map((row) => normalizeDowntimeLogRow(row))
      : [],
    supervisorLogs: Array.isArray(line?.supervisorLogs) ? line.supervisorLogs : [],
    auditRows: Array.isArray(line?.auditRows) ? line.auditRows : []
  };
  normalized.dayVisualiserKeyStageId = normalizeDayVisualiserKeyStageId(normalized, line?.dayVisualiserKeyStageId);
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
      dataSourceId: "",
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

function isIsoDateValue(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value || ""));
}

function isWeekendDate(date) {
  if (!(date instanceof Date)) return false;
  const day = date.getDay();
  return day === 0 || day === 6;
}

function isWeekendIsoDate(isoDate) {
  if (!isIsoDateValue(isoDate)) return false;
  return isWeekendDate(parseDateLocal(isoDate));
}

function normalizeWeekdayIsoDate(isoDate, { direction = 1, fallbackIso = "" } = {}) {
  const step = direction < 0 ? -1 : 1;
  const fallbackDate = isIsoDateValue(fallbackIso) ? fallbackIso : todayISO();
  const source = isIsoDateValue(isoDate) ? String(isoDate) : fallbackDate;
  const date = parseDateLocal(source);
  while (isWeekendDate(date)) {
    date.setDate(date.getDate() + step);
  }
  return formatDateLocal(date);
}

function shiftWeekdayIsoDate(isoDate, direction = 1, stepCount = 1) {
  const step = direction < 0 ? -1 : 1;
  const count = Math.max(1, Math.floor(num(stepCount)) || 1);
  const date = parseDateLocal(normalizeWeekdayIsoDate(isoDate, { direction: step }));
  let remaining = count;
  while (remaining > 0) {
    date.setDate(date.getDate() + step);
    if (!isWeekendDate(date)) remaining -= 1;
  }
  return formatDateLocal(date);
}

function isOperationalDate(isoDate) {
  return isIsoDateValue(isoDate) && !isWeekendIsoDate(isoDate);
}

function todayISO() {
  return formatDateLocal(new Date());
}

function supervisorAutoEntryDate() {
  return normalizeWeekdayIsoDate(todayISO(), { direction: -1 });
}

function setSupervisorAutoDateValue(inputOrId, isoDate, { fallbackIso = supervisorAutoEntryDate() } = {}) {
  const input = typeof inputOrId === "string" ? document.getElementById(inputOrId) : inputOrId;
  const safeFallback = isIsoDateValue(fallbackIso) ? String(fallbackIso).trim() : supervisorAutoEntryDate();
  const nextValue = isIsoDateValue(isoDate) ? String(isoDate).trim() : safeFallback;
  if (input) input.value = nextValue;
  if (input?.id && typeof document !== "undefined") {
    const display = document.querySelector(`[data-supervisor-auto-date="${input.id}"] .supervisor-auto-date-value`);
    if (display) display.textContent = formatIsoDateLabel(nextValue, { month: "short", day: "numeric", year: "numeric" });
    const wrapper = document.querySelector(`[data-supervisor-auto-date="${input.id}"]`);
    if (wrapper) wrapper.setAttribute("title", nextValue);
  }
  return nextValue;
}

function supervisorAutoDateValue(inputOrId, fallbackIso = supervisorAutoEntryDate()) {
  const input = typeof inputOrId === "string" ? document.getElementById(inputOrId) : inputOrId;
  return setSupervisorAutoDateValue(input, String(input?.value || "").trim(), { fallbackIso });
}

function shiftPresenceByDate(line) {
  const presenceByDate = {};
  (line?.shiftRows || []).forEach((row) => {
    const isoDate = String(row?.date || "");
    if (!isOperationalDate(isoDate)) return;
    const shiftValue = String(row?.shift || "");
    if (!presenceByDate[isoDate]) presenceByDate[isoDate] = { day: false, night: false };
    if (shiftValue === "Day") presenceByDate[isoDate].day = true;
    if (shiftValue === "Night") presenceByDate[isoDate].night = true;
    if (shiftValue === "Full Day") {
      presenceByDate[isoDate].day = true;
      presenceByDate[isoDate].night = true;
    }
  });
  return presenceByDate;
}

function lineShiftTrackerWeeksForWidth(widthPx) {
  const safeWidth = Math.max(0, num(widthPx));
  if (!safeWidth) return LINE_SHIFT_TRACKER_DEFAULT_WEEKS;
  const perWeekWidth = LINE_SHIFT_TRACKER_CELL_SIZE + LINE_SHIFT_TRACKER_CELL_GAP;
  const fitWeeks = Math.floor((safeWidth + LINE_SHIFT_TRACKER_CELL_GAP) / perWeekWidth);
  return Math.max(LINE_SHIFT_TRACKER_MIN_WEEKS, Math.min(LINE_SHIFT_TRACKER_MAX_WEEKS, fitWeeks));
}

function lineShiftTrackerCells(line, { anchorIsoDate = todayISO(), weeks = LINE_SHIFT_TRACKER_DEFAULT_WEEKS } = {}) {
  const safeWeeks = Math.max(LINE_SHIFT_TRACKER_MIN_WEEKS, Math.min(LINE_SHIFT_TRACKER_MAX_WEEKS, Math.floor(num(weeks)) || LINE_SHIFT_TRACKER_DEFAULT_WEEKS));
  const anchorDate = parseDateLocal(normalizeWeekdayIsoDate(anchorIsoDate || todayISO(), { direction: -1 }));
  const anchor = new Date(anchorDate.getFullYear(), anchorDate.getMonth(), anchorDate.getDate());
  const currentWeekStart = new Date(anchor.getFullYear(), anchor.getMonth(), anchor.getDate() - anchor.getDay());
  const start = new Date(currentWeekStart.getFullYear(), currentWeekStart.getMonth(), currentWeekStart.getDate() - (safeWeeks - 1) * 7);
  const presenceByDate = shiftPresenceByDate(line);
  const cells = [];
  for (let index = 0; index < safeWeeks * 7; index += 1) {
    const date = new Date(start.getFullYear(), start.getMonth(), start.getDate() + index);
    if (date > anchor) {
      cells.push({ isPad: true });
      continue;
    }
    if (isWeekendDate(date)) {
      cells.push({ isPad: true });
      continue;
    }
    const isoDate = formatDateLocal(date);
    const presence = presenceByDate[isoDate] || { day: false, night: false };
    let level = "none";
    if (presence.day && presence.night) level = "both";
    else if (presence.day || presence.night) level = "one";
    cells.push({ isoDate, level });
  }
  return { cells, weeks: safeWeeks };
}

function renderLineShiftTrackerGrid(line, { anchorIsoDate = todayISO(), weeks = LINE_SHIFT_TRACKER_DEFAULT_WEEKS } = {}) {
  const { cells, weeks: safeWeeks } = lineShiftTrackerCells(line, { anchorIsoDate, weeks });
  const levelClass = { none: "is-none", one: "is-one", both: "is-both" };
  const levelLabel = {
    none: "No shifts logged",
    one: "One shift logged",
    both: "Day and night shifts logged"
  };
  const cellHtml = cells
    .map((entry) => {
      if (entry?.isPad) return `<span class="line-shift-cell is-pad" aria-hidden="true"></span>`;
      const tone = levelClass[entry.level] || "is-none";
      const label = levelLabel[entry.level] || levelLabel.none;
      return `<span class="line-shift-cell ${tone}" title="${entry.isoDate}: ${label}" aria-label="${entry.isoDate}: ${label}"></span>`;
    })
    .join("");
  return `
    <div class="line-shift-grid" role="img" aria-label="Shift log activity over the last ${safeWeeks} weeks">
      ${cellHtml}
    </div>
  `;
}

function renderLineShiftTracker(line) {
  const lineId = String(line?.id || "");
  return `<section class="line-shift-tracker" data-line-shift-tracker="${lineId}" aria-label="Shift log activity"></section>`;
}

function renderLineShiftTrackersForWidth(anchorIsoDate = todayISO()) {
  if (typeof document === "undefined") return;
  const trackerNodes = Array.from(document.querySelectorAll("[data-line-shift-tracker]"));
  trackerNodes.forEach((tracker) => {
    const lineId = String(tracker.getAttribute("data-line-shift-tracker") || "");
    const line = appState.lines?.[lineId];
    if (!line) {
      tracker.innerHTML = "";
      return;
    }
    const styles = window.getComputedStyle(tracker);
    const paddingX = num(parseFloat(styles.paddingLeft)) + num(parseFloat(styles.paddingRight));
    const contentWidth = Math.max(0, tracker.clientWidth - paddingX);
    const weeks = lineShiftTrackerWeeksForWidth(contentWidth);
    tracker.innerHTML = renderLineShiftTrackerGrid(line, { anchorIsoDate, weeks });
  });
}

function scheduleLineShiftTrackerResizeRender() {
  if (lineShiftTrackerResizeTimer) clearTimeout(lineShiftTrackerResizeTimer);
  lineShiftTrackerResizeTimer = window.setTimeout(() => {
    lineShiftTrackerResizeTimer = null;
    renderLineShiftTrackersForWidth();
  }, 80);
}

function hhmmFromDate(date) {
  return `${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;
}

function minutesSinceRowStart(row, timeField, now = new Date()) {
  if (!row || !strictTimeValid(row?.[timeField])) return 0;
  const rowDate = String(row?.date || "");
  if (/^\d{4}-\d{2}-\d{2}$/.test(rowDate)) {
    const [hours, minutes] = String(row[timeField]).split(":").map(Number);
    const startedAt = parseDateLocal(rowDate);
    startedAt.setHours(hours, minutes, 0, 0);
    const elapsed = (now.getTime() - startedAt.getTime()) / 60000;
    return Number.isFinite(elapsed) ? Math.max(0, elapsed) : 0;
  }
  return Math.max(0, diffMinutes(String(row[timeField]), hhmmFromDate(now)));
}

function lineDowntimeReasonCategory(row) {
  return downtimeReasonCategoryFromRow(row);
}

function lineTileLiveSnapshot(line, now = new Date()) {
  const hasOpenShift = (line?.shiftRows || []).some((row) => isOperationalDate(String(row?.date || "")) && isPendingShiftLogRow(row));
  const hasOpenRun = (line?.runRows || []).some((row) => isOperationalDate(String(row?.date || "")) && isPendingRunLogRow(row));
  const openDowntimeRows = (line?.downtimeRows || []).filter(
    (row) => isOperationalDate(String(row?.date || "")) && isPendingDowntimeLogRow(row) && !isDowntimeBreakRow(row)
  );
  const hasOpenDowntime = openDowntimeRows.length > 0;
  const longestOpenDowntime = openDowntimeRows.reduce(
    (best, row) => {
      const minutes = minutesSinceRowStart(row, "downtimeStart", now);
      if (minutes > best.minutes) {
        return {
          minutes,
          category: lineDowntimeReasonCategory(row) || "Downtime"
        };
      }
      return best;
    },
    { minutes: 0, category: "" }
  );
  const longestOpenDowntimeMins = longestOpenDowntime.minutes;
  const hasSeriousDowntime = hasOpenDowntime && longestOpenDowntimeMins >= LINE_TILE_DOWNTIME_CRITICAL_MINS;

  return {
    hasOpenShift,
    hasOpenRun,
    hasOpenDowntime,
    hasSeriousDowntime,
    longestOpenDowntimeMins,
    liveDowntimeCategory: longestOpenDowntime.category || (hasOpenDowntime ? "Downtime" : "")
  };
}

function lineHasMovingFlow(line) {
  if (!line) return false;
  const hasOpenRun = (line?.runRows || []).some((row) => isOperationalDate(String(row?.date || "")) && isPendingRunLogRow(row));
  if (!hasOpenRun) return false;
  const hasOpenDowntime = (line?.downtimeRows || []).some(
    (row) => isOperationalDate(String(row?.date || "")) && isPendingDowntimeLogRow(row) && !isDowntimeBreakRow(row)
  );
  return !hasOpenDowntime;
}

function latestOpenRunLogRow(line) {
  const openRunRows = (line?.runRows || []).filter((row) => isOperationalDate(String(row?.date || "")) && isPendingRunLogRow(row));
  if (!openRunRows.length) return null;
  return openRunRows
    .slice()
    .sort((a, b) => rowNewestSortValue(b, "productionStartTime") - rowNewestSortValue(a, "productionStartTime"))[0];
}

function liveRunCrewingPatternForLine(line) {
  const openRunRow = latestOpenRunLogRow(line);
  if (!openRunRow) return null;
  const shift = preferredTimedLogShift(
    line,
    openRunRow.date,
    openRunRow.productionStartTime,
    openRunRow.finishTime || openRunRow.productionStartTime,
    "Day"
  );
  return normalizeRunCrewingPattern(openRunRow.runCrewingPattern, line, shift, { fallbackToIdeal: false });
}

function crewMapForLineShift(line, shift, stages = line?.stages || STAGES) {
  if (isFullDayShift(shift)) {
    return Object.fromEntries(
      stages.map((stage) => [
        stage.id,
        {
          crew: Math.max(0, num(line?.crewsByShift?.Day?.[stage.id]?.crew), num(line?.crewsByShift?.Night?.[stage.id]?.crew))
        }
      ])
    );
  }
  return line?.crewsByShift?.[shift] || defaultStageCrew(stages);
}

function stageCrewCountForVisual(line, stage, baseCrewMap, liveRunPattern = null) {
  const baseCrew = Math.max(0, Math.floor(num(baseCrewMap?.[stage.id]?.crew ?? stage?.crew)));
  if (!liveRunPattern || !isCrewedStage(stage)) return baseCrew;
  if (!Object.prototype.hasOwnProperty.call(liveRunPattern, stage.id)) return baseCrew;
  return Math.max(0, Math.floor(num(liveRunPattern[stage.id])));
}

function latestRunCrewingPatternForRows(line, runRows = [], fallbackShift = "Day") {
  const latestRunRow = (runRows || [])
    .slice()
    .sort((a, b) => rowNewestSortValue(b, "productionStartTime") - rowNewestSortValue(a, "productionStartTime"))[0];
  if (!latestRunRow) return null;
  const rowShift = preferredTimedLogShift(
    line,
    latestRunRow.date,
    latestRunRow.productionStartTime,
    latestRunRow.finishTime || latestRunRow.productionStartTime,
    fallbackShift
  );
  const pattern = normalizeRunCrewingPattern(latestRunRow.runCrewingPattern, line, rowShift, { fallbackToIdeal: false });
  return Object.keys(pattern).length ? pattern : null;
}

function latestRunCrewingSnapshotForRows(line, runRows = [], fallbackShift = "Day") {
  const latestRunRow = (runRows || [])
    .slice()
    .sort((a, b) => rowNewestSortValue(b, "productionStartTime") - rowNewestSortValue(a, "productionStartTime"))[0];
  if (!latestRunRow) return null;
  const rowShift = preferredTimedLogShift(
    line,
    latestRunRow.date,
    latestRunRow.productionStartTime,
    latestRunRow.finishTime || latestRunRow.productionStartTime,
    fallbackShift
  );
  const pattern = normalizeRunCrewingPattern(latestRunRow.runCrewingPattern, line, rowShift, { fallbackToIdeal: false });
  if (!Object.keys(pattern).length) return null;
  return {
    shift: rowShift,
    pattern,
    totalCrew: runCrewingPatternTotalCrew(pattern)
  };
}

function staffingSnapshotForSelection(line, runRows = [], shift = "Day") {
  const rows = Array.isArray(runRows) ? runRows : [];
  if (isFullDayShift(shift)) {
    const snapshots = LOG_ASSIGNABLE_SHIFTS.map((shiftValue) => {
      const rowsForShift = rows.filter((row) => {
        const coverage = shiftCoverageForTimedLog(line, row?.date, row?.productionStartTime, row?.finishTime || row?.productionStartTime);
        return coverage.shiftMatches.includes(shiftValue);
      });
      return latestRunCrewingSnapshotForRows(line, rowsForShift, shiftValue);
    }).filter(Boolean);
    const requiredCrew = LOG_ASSIGNABLE_SHIFTS.reduce((sum, shiftValue) => sum + requiredCrewForLineShift(line, shiftValue), 0);
    const actualCrew = snapshots.reduce((sum, snapshot) => sum + snapshot.totalCrew, 0);
    const understaffedBy = snapshots.length ? Math.max(0, requiredCrew - actualCrew) : 0;
    const overstaffedBy = snapshots.length ? Math.max(0, actualCrew - requiredCrew) : 0;
    return {
      actualCrew,
      requiredCrew,
      understaffedBy,
      overstaffedBy,
      hasCrewingData: snapshots.length > 0,
      staffingCallout: !snapshots.length
        ? "No run crewing"
        : understaffedBy > 0
          ? `Understaffed by ${understaffedBy}`
          : overstaffedBy > 0
            ? `Overstaffed by ${overstaffedBy}`
            : "Matches plan"
    };
  }

  const snapshot = latestRunCrewingSnapshotForRows(line, rows, shift);
  const requiredCrew = requiredCrewForLineShift(line, shift);
  const actualCrew = snapshot?.totalCrew || 0;
  const understaffedBy = snapshot ? Math.max(0, requiredCrew - actualCrew) : 0;
  const overstaffedBy = snapshot ? Math.max(0, actualCrew - requiredCrew) : 0;
  return {
    actualCrew,
    requiredCrew,
    understaffedBy,
    overstaffedBy,
    hasCrewingData: Boolean(snapshot),
    staffingCallout: !snapshot
      ? "No run crewing"
      : understaffedBy > 0
        ? `Understaffed by ${understaffedBy}`
        : overstaffedBy > 0
          ? `Overstaffed by ${overstaffedBy}`
          : "Matches plan"
  };
}

function lineTileLiveFeedbackLevel(snapshot) {
  if (!snapshot) return "";
  if (snapshot.hasSeriousDowntime && snapshot.hasOpenShift) return "critical";
  if (snapshot.hasOpenDowntime) return "warning";
  if (snapshot.hasOpenRun) return "active";
  return "";
}

function lineTileCalloutPillsHtml(snapshot) {
  if (!snapshot) return "";
  const pills = [];
  const downCategory = String(snapshot.liveDowntimeCategory || "Downtime").trim();
  const downLabel = `Down - ${htmlEscape(downCategory || "Downtime")}`;
  pills.push(
    snapshot.hasOpenShift
      ? `<span class="line-callout-pill is-green">On Shift</span>`
      : `<span class="line-callout-pill is-neutral">Not on Shift</span>`
  );
  pills.push(
    snapshot.hasOpenRun
      ? `<span class="line-callout-pill is-green">Running Product</span>`
      : `<span class="line-callout-pill is-neutral">Not Running Product</span>`
  );
  if (snapshot.hasSeriousDowntime && snapshot.hasOpenShift) pills.push(`<span class="line-callout-pill is-red">${downLabel} - Urgent</span>`);
  else if (snapshot.hasOpenDowntime) pills.push(`<span class="line-callout-pill is-yellow">${downLabel}</span>`);
  return pills.join("");
}

function renderLineCalloutPills(line, now = new Date()) {
  return lineTileCalloutPillsHtml(lineTileLiveSnapshot(line, now));
}

function renderHomeLineCard(line, { groupKey = "", showDragHandle = false } = {}) {
  const callouts = renderLineCalloutPills(line);
  const tracker = renderLineShiftTracker(line);
  const safeGroupKey = String(groupKey || "").trim() || "ungrouped";
  const dragHandle = showDragHandle
    ? `<button
          type="button"
          class="line-card-drag-handle"
          data-line-card-drag-handle
          draggable="true"
          title="Drag to reorder line"
          aria-label="Reorder ${htmlEscape(line.name)}"
        ><span aria-hidden="true">::</span></button>`
    : "";
  return `
    <article class="line-card" data-line-tile="${line.id}" data-line-card-id="${line.id}" data-line-card-group="${safeGroupKey}" data-open-line="${line.id}">
      <div class="line-card-head">
        <h3>${htmlEscape(line.name)}</h3>
        ${dragHandle}
      </div>
      ${tracker}
      <div class="line-card-footer">
        <div class="line-card-callouts">${callouts}</div>
        <div class="line-card-actions">
          <button type="button" class="table-edit-pill" data-edit-line="${line.id}">Edit</button>
        </div>
      </div>
    </article>
  `;
}

function renderGroupedHomeLineCards(lineList, lineGroups) {
  const sortedLines = sortLinesByDisplayOrder(lineList);
  const groups = normalizeLineGroups(lineGroups);
  if (!groups.length) {
    return {
      grouped: false,
      html: sortedLines.map((line) => renderHomeLineCard(line)).join("")
    };
  }

  const groupsById = new Set(groups.map((group) => group.id));
  const groupedLines = Object.fromEntries(groups.map((group) => [group.id, []]));
  const ungrouped = [];
  sortedLines.forEach((line) => {
    const groupId = String(line?.groupId || "").trim();
    if (groupId && groupsById.has(groupId)) groupedLines[groupId].push(line);
    else ungrouped.push(line);
  });

  const sections = groups
    .map((group) => {
      const lines = groupedLines[group.id] || [];
      if (!lines.length) return "";
      const lineCount = lines.length;
      const runningCount = lines.filter((line) => lineHasMovingFlow(line)).length;
      const groupKey = `group-${group.id}`;
      const expanded = isHomeLineGroupExpanded(groupKey);
      return `
        <section class="line-group-section" data-line-group-section="${group.id}">
          <header class="line-group-head">
            <button type="button" class="line-group-toggle" data-line-group-toggle="${groupKey}" aria-expanded="${expanded ? "true" : "false"}">
              <span class="line-group-caret" aria-hidden="true"></span>
              <h3>${htmlEscape(group.name)}</h3>
            </button>
            <div class="line-group-tools">
              <span class="line-group-count">${runningCount}/${lineCount} lines running</span>
              <button
                type="button"
                class="line-group-drag-handle"
                data-line-group-drag-handle
                draggable="true"
                title="Drag to reorder group"
                aria-label="Reorder ${htmlEscape(group.name)} group"
              ><span aria-hidden="true">::</span></button>
            </div>
          </header>
          <div class="line-group-body"${expanded ? "" : " hidden"}>
            ${expanded
              ? `<div class="line-cards" data-line-card-list="${group.id}">${lines.map((line) => renderHomeLineCard(line, { groupKey: group.id, showDragHandle: true })).join("")}</div>`
              : ""}
          </div>
        </section>
      `;
    })
    .filter(Boolean);

  if (ungrouped.length) {
    const lineCount = ungrouped.length;
    const runningCount = ungrouped.filter((line) => lineHasMovingFlow(line)).length;
    const groupKey = "group-ungrouped";
    const expanded = isHomeLineGroupExpanded(groupKey);
    sections.push(`
      <section class="line-group-section line-group-section-ungrouped" data-line-group-section="ungrouped">
        <header class="line-group-head">
          <button type="button" class="line-group-toggle" data-line-group-toggle="${groupKey}" aria-expanded="${expanded ? "true" : "false"}">
            <span class="line-group-caret" aria-hidden="true"></span>
            <h3>Ungrouped</h3>
          </button>
          <div class="line-group-tools">
            <span class="line-group-count">${runningCount}/${lineCount} lines running</span>
          </div>
        </header>
        <div class="line-group-body"${expanded ? "" : " hidden"}>
          ${expanded
            ? `<div class="line-cards" data-line-card-list="ungrouped">${ungrouped.map((line) => renderHomeLineCard(line, { groupKey: "ungrouped", showDragHandle: true })).join("")}</div>`
            : ""}
        </div>
      </section>
    `);
  }

  return {
    grouped: true,
    html: sections.join("")
  };
}

function renderManagerDataSourcesList(rootNode) {
  if (!rootNode) return;
  const sources = (appState.dataSources || []).filter((source) => source.isActive !== false);
  const assignments = dataSourceAssignments(appState.lines);
  const activeSourceIds = new Set(sources.map((source) => source.id));
  Array.from(dataSourceConnectionTestState.keys()).forEach((sourceId) => {
    if (!activeSourceIds.has(sourceId)) dataSourceConnectionTestState.delete(sourceId);
  });
  if (!sources.length) {
    rootNode.innerHTML = `<p class="data-source-empty">No data sources are configured yet.</p>`;
    return;
  }

  rootNode.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Source</th>
          <th>Machine</th>
          <th>Provider</th>
          <th>Connection</th>
          <th>Device ID</th>
          <th>Status</th>
          <th>Assigned Equipment</th>
          <th>Last Test</th>
          <th>Actions</th>
        </tr>
      </thead>
      <tbody>
        ${sources
          .map((source) => {
            const sourceAssignments = assignments.get(source.id) || [];
            const primaryAssignment = sourceAssignments[0] || null;
            const isConflict = sourceAssignments.length > 1;
            const statusClass = isConflict ? " is-conflict" : primaryAssignment ? " is-used" : "";
            const statusText = isConflict ? `${sourceAssignments.length} assigned` : primaryAssignment ? "Connected" : "Available";
            const assignmentText = isConflict
              ? sourceAssignments.map((entry) => `${entry.lineName} - ${entry.stageName}`).join("; ")
              : primaryAssignment
                ? `${primaryAssignment.lineName} - ${primaryAssignment.stageName}`
                : "Unassigned";
            const provider = DATA_SOURCE_PROVIDER_LABELS[source.provider] || String(source.provider || "SQL").toUpperCase();
            const machineText = source.machineNo || source.scaleNumber || "-";
            const connectionConfigured = source.connectionMode === "sql" ? source.hasSqlCredentials : source.hasApiKey;
            const connectionLabel = source.connectionMode === "sql"
              ? (connectionConfigured ? "SQL Configured" : "SQL Pending")
              : (connectionConfigured ? "API Key Set" : "API Pending");
            const sourceTitleParts = [source.deviceName, source.sourceKey].filter(Boolean);
            const testState = dataSourceConnectionTestFor(source.id);
            const testPill = dataSourceConnectionTestLabel(testState);
            const testTitle = testState?.message ? ` title="${htmlEscape(testState.message)}"` : "";
            return `
              <tr>
                <td title="${htmlEscape(sourceTitleParts.join(" | "))}">${htmlEscape(source.sourceName)}</td>
                <td>${htmlEscape(machineText)}</td>
                <td>${htmlEscape(provider)}</td>
                <td><span class="data-source-connection-pill${connectionConfigured ? " is-configured" : ""}">${htmlEscape(connectionLabel)}</span></td>
                <td>${htmlEscape(source.deviceId || "-")}</td>
                <td><span class="data-source-status-pill${statusClass}">${htmlEscape(statusText)}</span></td>
                <td>${htmlEscape(assignmentText)}</td>
                <td><span class="data-source-test-pill${testPill.css}"${testTitle}>${htmlEscape(testPill.text)}</span></td>
                <td class="data-source-actions-cell"><button type="button" class="ghost-btn data-source-test-btn" data-data-source-test="${htmlEscape(source.id)}">Test</button></td>
              </tr>
            `;
          })
          .join("")}
      </tbody>
    </table>
  `;
}

function isBizerbaIncomingSource(source) {
  const bizerbaPattern = /bizerba|bizberba/;
  const sourceName = String(source?.sourceName || "").trim().toLowerCase();
  const deviceName = String(source?.deviceName || "").trim().toLowerCase();
  return bizerbaPattern.test(sourceName) || bizerbaPattern.test(deviceName);
}

function renderIncomingDataSourcesStatus(rootNode) {
  if (!rootNode) return;
  const sources = (appState.dataSources || []).filter((source) => source.isActive !== false);
  const assignments = dataSourceAssignments(appState.lines);
  const source = sources.find((entry) => isBizerbaIncomingSource(entry)) || null;

  let row = {
    sourceName: "Bizerba Data Feed",
    detailText: "Supabase ingest endpoint | awaiting source mapping",
    tone: "is-attention",
    statusText: "Not Configured"
  };

  if (source) {
    const sourceAssignments = assignments.get(source.id) || [];
    const connectionConfigured = source.connectionMode === "sql" ? source.hasSqlCredentials : source.hasApiKey;
    const testState = dataSourceConnectionTestFor(source.id);
    const testFailed = Boolean(testState && !testState.ok);
    let tone = "is-ok";
    let statusText = "Ready";

    if (!connectionConfigured) {
      tone = "is-attention";
      statusText = "Connection Pending";
    } else if (testFailed) {
      tone = "is-failed";
      statusText = "Test Failed";
    } else if (!testState) {
      tone = "is-attention";
      statusText = "Not Tested";
    }

    const machineText = source.machineNo || source.scaleNumber || source.deviceId || source.deviceName || "-";
    const detailBits = [
      `Machine ${machineText}`,
      source.connectionMode === "sql" ? "SQL Feed" : "API Feed",
      `${sourceAssignments.length} stage${sourceAssignments.length === 1 ? "" : "s"}`
    ];

    row = {
      sourceName: "Bizerba Data Feed",
      detailText: detailBits.join(" | "),
      tone,
      statusText
    };
  }

  const total = 1;
  const readyCount = row.tone === "is-ok" ? 1 : 0;
  const attentionCount = row.tone === "is-attention" ? 1 : 0;
  const failedCount = row.tone === "is-failed" ? 1 : 0;
  const summaryClass = failedCount ? " is-failed" : attentionCount ? " is-attention" : " is-ok";
  const summaryText = failedCount
    ? `${formatNum(failedCount, 0)} source${failedCount === 1 ? "" : "s"} failing checks`
    : attentionCount
      ? `${formatNum(attentionCount, 0)} source${attentionCount === 1 ? "" : "s"} need attention`
      : "All incoming sources ready";

  rootNode.innerHTML = `
    <div class="incoming-source-summary">
      <span class="incoming-source-summary-pill${summaryClass}">${htmlEscape(summaryText)}</span>
      <div class="incoming-source-metrics">
        <span><strong>${formatNum(total, 0)}</strong> total</span>
        <span><strong>${formatNum(readyCount, 0)}</strong> ready</span>
        <span><strong>${formatNum(attentionCount, 0)}</strong> attention</span>
        <span><strong>${formatNum(failedCount, 0)}</strong> failed</span>
      </div>
    </div>
    <ul class="incoming-source-status-list">
      <li class="incoming-source-status-item">
        <div class="incoming-source-status-copy">
          <strong>${htmlEscape(row.sourceName)}</strong>
          <span>${htmlEscape(row.detailText)}</span>
        </div>
        <span class="incoming-source-state-pill ${row.tone}">${htmlEscape(row.statusText)}</span>
      </li>
    </ul>
  `;
}

function stageLiveDowntimePillHtml(liveDown) {
  if (!liveDown) return "";
  const mins = Math.max(1, Math.floor(num(liveDown.minutes)));
  const tone = liveDown.level === "critical" ? "critical" : "warning";
  return `<span class="stage-live-pill ${tone}">Down - ${mins} mins</span>`;
}

function liveEquipmentDowntimeByStage(line, stages = [], now = new Date()) {
  const openEquipmentRows = (line?.downtimeRows || []).filter((row) => {
    if (!isPendingDowntimeLogRow(row)) return false;
    const equipment = String(row?.equipment || "").trim();
    if (!equipment) return false;
    const reasonCategory = downtimeReasonCategoryFromRow(row);
    if (reasonCategory && reasonCategory !== "Equipment") return false;
    return true;
  });
  const byStageId = {};
  stages.forEach((stage) => {
    let maxMinutes = 0;
    openEquipmentRows.forEach((row) => {
      if (!matchesStage(stage, row.equipment)) return;
      maxMinutes = Math.max(maxMinutes, minutesSinceRowStart(row, "downtimeStart", now));
    });
    if (maxMinutes <= 0) return;
    byStageId[stage.id] = {
      minutes: maxMinutes,
      level: maxMinutes > LINE_TILE_DOWNTIME_CRITICAL_MINS ? "critical" : "warning"
    };
  });
  return byStageId;
}

function applyLineTileLiveFeedback(now = new Date()) {
  if (typeof document === "undefined") return;
  const tiles = Array.from(document.querySelectorAll("[data-line-tile]"));
  tiles.forEach((tile) => {
    const lineId = String(tile.getAttribute("data-line-tile") || "");
    const line = appState.lines?.[lineId];
    const snapshot = lineTileLiveSnapshot(line, now);
    const level = lineTileLiveFeedbackLevel(snapshot);
    tile.classList.remove("line-card-live-active", "line-card-live-warning", "line-card-live-critical");
    if (level === "active") tile.classList.add("line-card-live-active");
    if (level === "warning") tile.classList.add("line-card-live-warning");
    if (level === "critical") tile.classList.add("line-card-live-critical");
    const calloutWrap = tile.querySelector(".line-card-callouts");
    if (calloutWrap) calloutWrap.innerHTML = lineTileCalloutPillsHtml(snapshot);
  });
}

function syncLineTileFeedbackLoop({ enabled = false } = {}) {
  if (!enabled) {
    if (lineTileFeedbackTimer) {
      clearInterval(lineTileFeedbackTimer);
      lineTileFeedbackTimer = null;
    }
    return;
  }
  if (lineTileFeedbackTimer) return;
  lineTileFeedbackTimer = window.setInterval(() => {
    if (appState.activeView === "home") applyLineTileLiveFeedback();
    if (appState.activeView === "line" && appState.appMode === "manager" && state && !state.visualEditMode) renderVisualiser();
  }, LINE_TILE_FEEDBACK_REFRESH_MS);
}

function num(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function roundToDecimals(value, decimals = 1) {
  const safeDecimals = Math.max(0, Math.floor(num(decimals)));
  const factor = 10 ** safeDecimals;
  return Math.round(num(value) * factor) / factor;
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
  const startTime = String(row?.productionStartTime || "").trim();
  const finishTime = String(row?.finishTime || "").trim();
  return strictTimeValid(startTime) && (finishTime === "" || (strictTimeValid(finishTime) && startTime === finishTime));
}

function isPendingShiftLogRow(row) {
  return strictTimeValid(row?.startTime) && strictTimeValid(row?.finishTime) && row.startTime === row.finishTime;
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

function supervisorOwnsPendingLogRow(row, session = appState.supervisorSession) {
  if (!row || !session) return false;
  const rowUserId = String(row?.submittedByUserId || "").trim();
  const sessionUserId = String(session?.userId || "").trim();
  if (rowUserId && sessionUserId) return rowUserId === sessionUserId;
  const rowSubmittedBy = String(row?.submittedBy || "").trim().toLowerCase();
  if (!rowSubmittedBy) return false;
  const assignedSupervisor = supervisorByUsername(session?.username || "");
  const identityKeys = new Set(
    [
      session?.username,
      session?.name,
      assignedSupervisor?.username,
      assignedSupervisor?.name
    ]
      .map((value) => String(value || "").trim().toLowerCase())
      .filter(Boolean)
  );
  return identityKeys.has(rowSubmittedBy);
}

function breakRowsForShift(line, shiftLogOrId, fallback = {}) {
  const breakRows = Array.isArray(line?.breakRows) ? line.breakRows : [];
  const hasShiftLinkedBreaks = breakRows.some((row) => Boolean(String(row?.shiftLogId || "").trim()));
  const shiftMeta = (typeof shiftLogOrId === "object" && shiftLogOrId)
    ? shiftLogOrId
    : {
      id: shiftLogOrId,
      date: fallback.date,
      shift: fallback.shift
    };
  const shiftLogId = String(shiftMeta?.id || "").trim();
  const shiftDate = String(shiftMeta?.date || "").trim();
  const shiftName = String(shiftMeta?.shift || "").trim();
  return breakRows
    .filter((row) => {
      if (hasShiftLinkedBreaks) {
        const rowShiftLogId = String(row?.shiftLogId || "").trim();
        if (rowShiftLogId && shiftLogId) return rowShiftLogId === shiftLogId;
        if (!rowShiftLogId && shiftDate && shiftName) {
          return String(row?.date || "") === shiftDate && String(row?.shift || "") === shiftName;
        }
        return false;
      }
      if (shiftDate && shiftName) {
        return String(row?.date || "") === shiftDate && String(row?.shift || "") === shiftName;
      }
      if (!shiftLogId) return false;
      return String(row?.shiftLogId || "") === shiftLogId;
    })
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

function overlapMinutesBetweenIntervalSets(leftIntervals = [], rightIntervals = []) {
  if (!Array.isArray(leftIntervals) || !Array.isArray(rightIntervals)) return 0;
  if (!leftIntervals.length || !rightIntervals.length) return 0;
  return leftIntervals.reduce(
    (sum, left) => sum + rightIntervals.reduce((inner, right) => inner + intervalOverlapMinutes(left, right), 0),
    0
  );
}

function buildNonProductionIntervalsByDate(breakRows = [], downtimeRows = []) {
  const byDate = new Map();
  const addIntervals = (rows, startField, finishField) => {
    (rows || []).forEach((row) => {
      const date = String(row?.date || "").trim();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return;
      const intervals = intervalsFromTimes(row?.[startField], row?.[finishField]);
      if (!intervals.length) return;
      const existing = byDate.get(date) || [];
      existing.push(...intervals);
      byDate.set(date, existing);
    });
  };
  addIntervals(breakRows, "breakStart", "breakFinish");
  addIntervals(downtimeRows, "downtimeStart", "downtimeFinish");
  byDate.forEach((intervals, date) => {
    byDate.set(date, mergeMinuteIntervals(intervals));
  });
  return byDate;
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

function lineHasLoggedShiftRowsForDate(line, date) {
  const safeDate = String(date || "").trim();
  if (!line || !/^\d{4}-\d{2}-\d{2}$/.test(safeDate)) return false;
  return (line.shiftRows || []).some(
    (row) =>
      String(row?.date || "").trim() === safeDate
      && LOG_ASSIGNABLE_SHIFTS.includes(String(row?.shift || "").trim())
      && strictTimeValid(row?.startTime)
      && strictTimeValid(row?.finishTime)
  );
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
  const assignedShift = normalizeLogAssignableShift(row?.assignedShift || row?.shift);
  if (assignedShift) return assignedShift;
  const explicit = String(row?.__shiftLabel || "").trim();
  if (explicit) return explicit;
  const coverage = shiftCoverageForTimedLog(line, row?.date, row?.[startField], row?.[finishField]);
  if (coverage.shiftMatches.length || coverage.unassignedMinutes > 0) return coverage.label;
  return "Unassigned";
}

function rowMatchesDateShift(row, date, shift, { line = null, startField = "", finishField = "" } = {}) {
  if (!row || row.date !== date || !isOperationalDate(date)) return false;
  if (isFullDayShift(shift)) return true;
  const requestedShift = shift === "Night" ? "Night" : "Day";
  const explicitAssignedShift = normalizeLogAssignableShift(row?.assignedShift || row?.shift);
  if (explicitAssignedShift) return explicitAssignedShift === requestedShift;
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

function crewSettingsShiftForLine(line) {
  const explicit = String(line?.crewSettingsShift || "").trim();
  if (CREW_SETTINGS_SHIFTS.includes(explicit)) return explicit;
  return fallbackShiftValue(line?.selectedShift || "Day");
}

function isLineSettingsDirty(lineId = state?.id) {
  const safeLineId = String(lineId || "").trim();
  return Boolean(safeLineId) && pendingLineSettingsSaveIds.has(safeLineId);
}

function isLineSettingsSaveInFlight(lineId = state?.id) {
  const safeLineId = String(lineId || "").trim();
  return Boolean(safeLineId) && lineSettingsSaveInFlightIds.has(safeLineId);
}

function markLineSettingsDirty(lineId = state?.id) {
  const safeLineId = String(lineId || "").trim();
  if (!safeLineId) return;
  pendingLineSettingsSaveIds.add(safeLineId);
  syncLineSettingsSaveUI();
}

function clearLineSettingsDirty(lineId = state?.id) {
  const safeLineId = String(lineId || "").trim();
  if (!safeLineId) return;
  pendingLineSettingsSaveIds.delete(safeLineId);
  syncLineSettingsSaveUI();
}

function syncLineSettingsSaveUI() {
  const saveBtn = document.getElementById("saveStageCrewSettingsBtn");
  const saveStatus = document.getElementById("stageCrewSaveStatus");
  if (!saveBtn && !saveStatus) return;
  const activeLineId = String(state?.id || "").trim();
  const hasLine = Boolean(activeLineId);
  const dirty = hasLine && isLineSettingsDirty(activeLineId);
  const saving = hasLine && isLineSettingsSaveInFlight(activeLineId);
  if (saveBtn) {
    saveBtn.disabled = !hasLine || !dirty || saving;
    saveBtn.textContent = saving ? "Saving..." : "Save Stage Settings";
  }
  if (saveStatus) {
    if (!hasLine) {
      saveStatus.textContent = "";
    } else if (saving) {
      saveStatus.textContent = "Saving stage settings...";
    } else if (dirty) {
      saveStatus.textContent = "Unsaved changes.";
    } else {
      saveStatus.textContent = "All stage settings saved.";
    }
  }
}

async function saveCurrentLineSettings() {
  const activeLineId = String(state?.id || "").trim();
  if (!activeLineId || !isLineSettingsDirty(activeLineId) || isLineSettingsSaveInFlight(activeLineId)) {
    syncLineSettingsSaveUI();
    return;
  }
  lineSettingsSaveInFlightIds.add(activeLineId);
  syncLineSettingsSaveUI();
  try {
    await saveLineModelToBackend(activeLineId);
    clearLineSettingsDirty(activeLineId);
    saveState();
  } catch (error) {
    console.warn("Stage settings save failed:", error);
    alert(`Could not save stage settings.\n${error?.message || "Please retry."}`);
  } finally {
    lineSettingsSaveInFlightIds.delete(activeLineId);
    syncLineSettingsSaveUI();
  }
}

function isWithinShiftWindow(targetMins, startMins, finishMins) {
  if (!Number.isFinite(targetMins) || !Number.isFinite(startMins) || !Number.isFinite(finishMins)) return false;
  if (startMins === finishMins) return targetMins === startMins;
  if (finishMins > startMins) return targetMins >= startMins && targetMins <= finishMins;
  return targetMins >= startMins || targetMins <= finishMins;
}

function inferShiftForLog(line, date, timeValue, fallbackShift = "Day") {
  const fallback = fallbackShiftValue(fallbackShift);
  const shiftRows = (line?.shiftRows || [])
    .filter((row) => row.date === date && SHIFT_OPTIONS.includes(row.shift))
    .slice();
  const targetMins = parseTimeToMinutes(timeValue);

  if (Number.isFinite(targetMins)) {
    const matchedRows = shiftRows.filter((row) => {
      const startMins = parseTimeToMinutes(row.startTime);
      const finishMins = parseTimeToMinutes(row.finishTime);
      return isWithinShiftWindow(targetMins, startMins, finishMins);
    });
    if (matchedRows.length) {
      const fallbackMatch = matchedRows.find((row) => row.shift === fallback);
      if (fallbackMatch?.shift) return fallbackMatch.shift;
      const mostRecentMatch = matchedRows
        .slice()
        .sort((a, b) => rowNewestSortValue(b, "startTime") - rowNewestSortValue(a, "startTime"))[0];
      if (mostRecentMatch?.shift) return mostRecentMatch.shift;
    }
  }

  return fallback;
}

function preferredTimedLogShift(line, date, startValue, finishValue, fallbackShift = "Day") {
  const fallback = fallbackShiftValue(fallbackShift);
  const coverage = shiftCoverageForTimedLog(line, date, startValue, finishValue);
  if (coverage.shiftMatches.includes(fallback)) return fallback;
  const inferred = inferShiftForLog(line, date, startValue, fallback);
  if (coverage.shiftMatches.includes(inferred)) return inferred;
  return coverage.shiftMatches[0] || inferred || fallback;
}

function supervisorCanAccessTimedLog(session, lineId, line, date, startValue, finishValue) {
  const coverage = shiftCoverageForTimedLog(line, date, startValue, finishValue);
  return coverage.shiftMatches.some((shift) => supervisorCanAccessShift(session, lineId, shift));
}

function selectedShiftRowsByDate(rows, date, shift, { line = null } = {}) {
  if (!isOperationalDate(date)) return [];
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
  return isOperationalDate(date) && SHIFT_OPTIONS.includes(String(shift || ""));
}

function normalizeLogAssignableShift(value) {
  const shift = String(value || "").trim();
  return LOG_ASSIGNABLE_SHIFTS.includes(shift) ? shift : "";
}

async function apiRequest(path, { method = "GET", token = "", body, timeoutMs = API_REQUEST_TIMEOUT_MS } = {}) {
  dbLoadingRequestCount += 1;
  showStartupLoading("Loading production data...");
  try {
    const headers = {};
    if (token) headers.Authorization = `Bearer ${token}`;
    if (body !== undefined) headers["Content-Type"] = "application/json";
    const requestTimeoutMs =
      Number.isFinite(Number(timeoutMs)) && Number(timeoutMs) > 0
        ? Number(timeoutMs)
        : API_REQUEST_TIMEOUT_MS;
    const controller = typeof AbortController !== "undefined" ? new AbortController() : null;
    const timeoutId =
      controller && typeof window !== "undefined"
        ? window.setTimeout(() => controller.abort(), requestTimeoutMs)
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
        throw new Error(`Request timed out after ${Math.round(requestTimeoutMs / 1000)}s.`);
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
  managerBackendSession.name = "";
  managerBackendSession.username = "";
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
  line.groupId = String(lineSummary?.groupId || "").trim();
  line.displayOrder = Math.max(0, Math.floor(num(lineSummary?.displayOrder)));
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
          dataSourceId: String(stage.dataSourceId || "").trim(),
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

  line.shiftRows = Array.isArray(logs?.shiftRows)
    ? logs.shiftRows.map((row) => ({
        ...row,
        notes: String(row?.notes || "")
      }))
    : [];
  line.breakRows = Array.isArray(logs?.breakRows) ? logs.breakRows : [];
  line.runRows = Array.isArray(logs?.runRows)
    ? logs.runRows.map((row) => normalizeRunLogRow(row))
    : [];
  line.downtimeRows = Array.isArray(logs?.downtimeRows)
    ? logs.downtimeRows.map((row) => normalizeDowntimeLogRow(row))
    : [];
  line.auditRows = [];
  line.supervisorLogs = [];
  line.selectedStageId = line.stages[0]?.id || "";
  line.dayVisualiserKeyStageId = "";
  enforcePermanentCrewAndThroughput(line);
  return line;
}

async function refreshHostedState(preferredSession = null) {
  try {
    const previousLines = appState.lines || {};
    const activeSession = preferredSession || (await ensureManagerBackendSession());
    if (!activeSession?.backendToken) throw new Error("Missing backend token.");
    const activeSessionName = String(activeSession?.name || "").trim();
    const activeSessionUsername = String(activeSession?.username || "").trim();
    if (!activeSessionName || !activeSessionUsername) {
      try {
        const mePayload = await apiRequest("/api/me", { token: activeSession.backendToken });
        const meName = String(mePayload?.user?.name || "").trim();
        const meUsername = String(mePayload?.user?.username || "").trim().toLowerCase();
        if (activeSession?.role === "manager") {
          if (meName) managerBackendSession.name = meName;
          if (meUsername) managerBackendSession.username = meUsername;
          persistManagerBackendSession();
        } else if (activeSession?.role === "supervisor" && appState.supervisorSession) {
          if (meName) appState.supervisorSession.name = meName;
          if (meUsername) appState.supervisorSession.username = meUsername;
          saveState();
        }
      } catch (profileError) {
        console.warn("Could not load user profile details:", profileError);
      }
    }
    let snapshot = null;
    let snapshotError = null;
    for (let attempt = 1; attempt <= 2; attempt += 1) {
      try {
        snapshot = await apiRequest("/api/state-snapshot", { token: activeSession.backendToken });
        snapshotError = null;
        break;
      } catch (error) {
        snapshotError = error;
        if (attempt >= 2 || !shouldRetryHostedSnapshot(error)) break;
        await delayMs(350 * attempt);
      }
    }
    if (snapshotError || !snapshot) throw snapshotError || new Error("Could not load state snapshot.");
    let snapshotLines = Array.isArray(snapshot?.lines) ? snapshot.lines : [];
    if (!snapshotLines.length && activeSession.role === "manager") {
      try {
        const fallbackLinesPayload = await apiRequest("/api/lines", { token: activeSession.backendToken });
        const fallbackLines = Array.isArray(fallbackLinesPayload?.lines) ? fallbackLinesPayload.lines : [];
        if (fallbackLines.length) {
          const bundles = await Promise.all(
            fallbackLines
              .map((line) => line?.id)
              .filter((lineId) => UUID_RE.test(String(lineId || "")))
              .map(async (lineId) => {
                const [detail, logs] = await Promise.all([
                  apiRequest(`/api/lines/${lineId}`, { token: activeSession.backendToken }),
                  apiRequest(`/api/lines/${lineId}/logs`, { token: activeSession.backendToken })
                ]);
                return {
                  line: fallbackLines.find((item) => item.id === lineId) || { id: lineId, name: "Production Line" },
                  stages: Array.isArray(detail?.stages) ? detail.stages : [],
                  guides: Array.isArray(detail?.guides) ? detail.guides : [],
                  shiftRows: Array.isArray(logs?.shiftRows) ? logs.shiftRows : [],
                  breakRows: Array.isArray(logs?.breakRows) ? logs.breakRows : [],
                  runRows: Array.isArray(logs?.runRows) ? logs.runRows : [],
                  downtimeRows: Array.isArray(logs?.downtimeRows) ? logs.downtimeRows : []
                };
              })
          );
          snapshotLines = bundles;
        }
      } catch (fallbackError) {
        console.warn("State snapshot fallback load failed:", fallbackError);
      }
      if (!snapshotLines.length) {
        try {
          await apiRequest("/api/health");
        } catch (healthError) {
          throw new Error("Database unavailable");
        }
      }
    }
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
      line.dayVisualiserKeyStageId = normalizeDayVisualiserKeyStageId(
        line,
        previousLines?.[line.id]?.dayVisualiserKeyStageId || ""
      );
      hostedLines[line.id] = line;
      activeSession.backendLineMap[line.id] = line.id;
      (line.stages || []).forEach((stage) => {
        activeSession.backendStageMap[`${line.id}::${stage.id}`] = stage.id;
      });
    });

    appState.lines = hostedLines;
    const snapshotLineGroups = normalizeLineGroups(snapshot?.lineGroups);
    const derivedLineGroups = normalizeLineGroups(
      snapshotLines.map((bundle, index) => ({
        id: bundle?.line?.groupId,
        name: bundle?.line?.groupName,
        displayOrder: index
      }))
    );
    appState.lineGroups = snapshotLineGroups.length ? snapshotLineGroups : derivedLineGroups;
    if (Array.isArray(snapshot?.dataSources)) {
      appState.dataSources = normalizeDataSources(snapshot.dataSources);
    }
    const ids = Object.keys(hostedLines);
    pendingLineSettingsSaveIds.clear();
    lineSettingsSaveInFlightIds.clear();
    appState.activeLineId = hostedLines[appState.activeLineId] ? appState.activeLineId : ids[0] || "";
    state = appState.lines[appState.activeLineId] || null;
    if (pendingManagerDataTabRestore.lineId && !hostedLines[pendingManagerDataTabRestore.lineId]) {
      pendingManagerDataTabRestore = { lineId: "", tabId: "" };
    }

    const supervisors = Array.isArray(snapshot?.supervisors)
      ? snapshot.supervisors.filter((sup) => sup && sup.isActive !== false)
      : [];
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
    if (Array.isArray(snapshot?.supervisorActions)) {
      appState.supervisorActions = normalizeSupervisorActions(snapshot.supervisorActions);
    } else {
      appState.supervisorActions = [];
    }
    if (Array.isArray(snapshot?.productCatalog)) {
      setProductCatalogEntries(snapshot.productCatalog);
    } else {
      appState.productCatalogEntries = [];
    }
    const productCatalogTable = document.getElementById("productCatalogTable");
    if (productCatalogTable) {
      hydrateProductCatalogTableFromState(productCatalogTable);
      ensureProductCatalogActionColumn(productCatalogTable);
      refreshProductRowAssignmentSummaries(productCatalogTable);
      syncRunProductInputsFromCatalog();
    }

    if (appState.supervisorSession) {
      const current = appState.supervisors.find((sup) => sup.username === appState.supervisorSession.username);
      if (current) {
        appState.supervisorSession = {
          ...appState.supervisorSession,
          name: String(current.name || current.username || appState.supervisorSession.name || appState.supervisorSession.username || "").trim(),
          assignedLineIds: current.assignedLineIds.slice(),
          assignedLineShifts: clone(current.assignedLineShifts || {})
        };
      } else if (hasSnapshotSupervisorAssignments) {
        appState.supervisorSession = {
          ...appState.supervisorSession,
          name: String(appState.supervisorSession.name || appState.supervisorSession.username || "").trim(),
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
          name: String(appState.supervisorSession.name || appState.supervisorSession.username || "").trim(),
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

function shouldRetryHostedSnapshot(error) {
  const status = Number(error?.status || 0);
  const message = String(error?.message || "").toLowerCase();
  if ([408, 429, 500, 502, 503, 504].includes(status)) return true;
  return (
    message.includes("timed out") ||
    message.includes("timeout") ||
    message.includes("networkerror") ||
    message.includes("failed to fetch") ||
    message.includes("database unavailable")
  );
}

function delayMs(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function ensureBackendStageId(localLineId, localStageId, session) {
  if (!localLineId || !localStageId || !session?.backendToken) return "";
  const key = `${localLineId}::${localStageId}`;
  const cachedStageId = String(session.backendStageMap?.[key] || "").trim();
  if (cachedStageId) {
    if (UUID_RE.test(cachedStageId)) return cachedStageId;
    // Discard stale/non-UUID cache entries so they cannot poison API payloads.
    if (session.backendStageMap && Object.prototype.hasOwnProperty.call(session.backendStageMap, key)) {
      delete session.backendStageMap[key];
      if (session === managerBackendSession) persistManagerBackendSession();
      else saveState();
    }
  }
  const backendLineId = await ensureBackendLineId(localLineId, session);
  if (!backendLineId) return "";
  const payload = await apiRequest(`/api/lines/${backendLineId}`, { token: session.backendToken });
  const backendStages = Array.isArray(payload?.stages) ? payload.stages : [];
  const localStages = appState.lines[localLineId]?.stages || [];
  const localStage = localStages.find((stage) => stage.id === localStageId);
  if (!localStage) return "";
  let matched = null;
  const localStageIdRaw = String(localStage?.id || "").trim();
  if (UUID_RE.test(localStageIdRaw)) {
    const byId = backendStages.find((stage) => String(stage?.id || "").trim() === localStageIdRaw);
    if (byId?.id && UUID_RE.test(String(byId.id).trim())) matched = byId;
  }
  const localIndex = localStages.findIndex((stage) => stage.id === localStageId);
  if (!matched && localIndex >= 0) {
    const expectedOrder = localIndex + 1;
    const byOrder = backendStages.find(
      (stage) =>
        Number(stage?.stageOrder) === expectedOrder
        && UUID_RE.test(String(stage?.id || "").trim())
    );
    if (byOrder) matched = byOrder;
  }
  if (!matched) {
    const localName = stageNameCore(localStage.name);
    const byName = backendStages.filter(
      (stage) =>
        stageNameCore(stage?.stageName) === localName
        && UUID_RE.test(String(stage?.id || "").trim())
    );
    if (byName.length === 1) {
      [matched] = byName;
    } else if (byName.length > 1 && localIndex >= 0) {
      const expectedOrder = localIndex + 1;
      matched = byName
        .slice()
        .sort((left, right) => Math.abs(num(left?.stageOrder) - expectedOrder) - Math.abs(num(right?.stageOrder) - expectedOrder))[0];
    }
  }
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
  const validationError = validateDowntimeBackendPayload(payload);
  if (validationError) throw new Error(validationError);
  const session = await ensureManagerBackendSession();
  const backendLineId = await ensureBackendLineId(payload.lineId, session);
  if (!backendLineId) throw new Error("Line is not synced to server.");
  const backendEquipmentId = await ensureBackendStageId(payload.lineId, payload.equipment, session);
  const safeEquipmentStageId = UUID_RE.test(String(backendEquipmentId || "").trim()) ? backendEquipmentId : null;
  const body = {
    lineId: backendLineId,
    date: payload.date,
    downtimeStart: payload.downtimeStart,
    downtimeFinish: payload.downtimeFinish,
    equipmentStageId: safeEquipmentStageId,
    reason: payload.reason || "",
    notes: String(payload.notes || "")
  };
  if (Object.prototype.hasOwnProperty.call(payload, "excludeFromCalculation")) {
    body.excludeFromCalculation = normalizeDowntimeCalculationFlag(payload.excludeFromCalculation);
  }
  const response = await apiRequest("/api/logs/downtime", {
    method: "POST",
    token: session.backendToken,
    body
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

async function startManagerShiftBreak(shiftLogId, breakStart, breakFinish = null) {
  const session = await ensureManagerBackendSession();
  const payload = {
    breakStart
  };
  if (strictTimeValid(String(breakFinish || ""))) payload.breakFinish = String(breakFinish).trim();
  const response = await apiRequest(`/api/logs/shifts/${shiftLogId}/breaks`, {
    method: "POST",
    token: session.backendToken,
    body: payload
  });
  return response?.breakLog || null;
}

async function deleteManagerShiftBreak(shiftLogId, breakId) {
  const session = await ensureManagerBackendSession();
  await apiRequest(`/api/logs/shifts/${shiftLogId}/breaks/${breakId}`, {
    method: "DELETE",
    token: session.backendToken
  });
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
  const validationError = validateDowntimeBackendPayload(payload);
  if (validationError) throw new Error(validationError);
  const session = await ensureManagerBackendSession();
  const backendEquipmentId = payload.equipment
    ? await ensureBackendStageId(payload.lineId || state.id, payload.equipment, session)
    : null;
  const safeEquipmentStageId = UUID_RE.test(String(backendEquipmentId || "").trim()) ? backendEquipmentId : null;
  const body = {
    date: payload.date,
    downtimeStart: payload.downtimeStart,
    downtimeFinish: payload.downtimeFinish,
    equipmentStageId: safeEquipmentStageId,
    reason: payload.reason || "",
    notes: String(payload.notes || "")
  };
  if (Object.prototype.hasOwnProperty.call(payload, "excludeFromCalculation")) {
    body.excludeFromCalculation = normalizeDowntimeCalculationFlag(payload.excludeFromCalculation);
  }
  const response = await apiRequest(`/api/logs/downtime/${logId}`, {
    method: "PATCH",
    token: session.backendToken,
    body
  });
  return response?.downtimeLog || null;
}

async function deleteManagerShiftLog(logId) {
  const session = await ensureManagerBackendSession();
  await apiRequest(`/api/logs/shifts/${logId}`, {
    method: "DELETE",
    token: session.backendToken
  });
}

async function deleteManagerRunLog(logId) {
  const session = await ensureManagerBackendSession();
  await apiRequest(`/api/logs/runs/${logId}`, {
    method: "DELETE",
    token: session.backendToken
  });
}

async function deleteManagerDowntimeLog(logId) {
  const session = await ensureManagerBackendSession();
  await apiRequest(`/api/logs/downtime/${logId}`, {
    method: "DELETE",
    token: session.backendToken
  });
}

function managerDowntimeAuditLabel(line, {
  equipment = "",
  reasonCategory = "",
  reasonDetail = ""
} = {}) {
  return (
    stageNameByIdForLine(line || state, equipment) ||
    downtimeDetailLabel(line || state, reasonCategory, reasonDetail) ||
    reasonCategory ||
    "line"
  );
}

async function createManagerShiftLogEntry(data = {}, { line = state } = {}) {
  const activeLine = line || state;
  const lineId = String(activeLine?.id || "").trim();
  if (!lineId) throw new Error("No production line selected.");
  const next = {
    date: String(data.date || "").trim(),
    shift: String(data.shift || "").trim(),
    startTime: String(data.startTime || "").trim(),
    finishTime: String(data.finishTime || "").trim(),
    notes: String(data.notes || "").trim()
  };
  if (!rowIsValidDateShift(next.date, next.shift)) {
    throw new Error("Date and shift are required. Weekend dates are excluded.");
  }
  if (!strictTimeValid(next.startTime) || !strictTimeValid(next.finishTime)) {
    throw new Error("Shift start and finish must be HH:MM (24h).");
  }
  const saved = await syncManagerShiftLog({
    lineId,
    date: next.date,
    shift: next.shift,
    startTime: next.startTime,
    finishTime: next.finishTime,
    notes: next.notes
  });
  if (!Array.isArray(activeLine.shiftRows)) activeLine.shiftRows = [];
  const row = {
    ...next,
    id: saved?.id || data.id || makeLocalLogId("shift"),
    submittedBy: "manager",
    submittedAt: saved?.submittedAt || nowIso()
  };
  activeLine.shiftRows.push(row);
  addAudit(activeLine, "MANAGER_SHIFT_LOG", `Manager logged ${next.shift} shift for ${next.date}`);
  saveState();
  return row;
}

function managerBreakShiftChoices(line, date) {
  return (Array.isArray(line?.shiftRows) ? line.shiftRows : [])
    .filter((row) => String(row?.date || "").trim() === String(date || "").trim())
    .filter((row) => SHIFT_OPTIONS.includes(String(row?.shift || "").trim()))
    .filter((row) => strictTimeValid(String(row?.startTime || "").trim()) && strictTimeValid(String(row?.finishTime || "").trim()))
    .filter((row) => UUID_RE.test(String(row?.id || "").trim()))
    .slice()
    .sort((a, b) => {
      const startCmp = String(a?.startTime || "").localeCompare(String(b?.startTime || ""));
      if (startCmp !== 0) return startCmp;
      return rowNewestSortValue(b, "startTime") - rowNewestSortValue(a, "startTime");
    })
    .map((row) => ({
      id: String(row.id || "").trim(),
      shift: String(row.shift || "").trim(),
      label: `${row.shift} | ${formatTime12h(row.startTime)} - ${formatTime12h(row.finishTime)}`
    }));
}

async function createManagerBreakLogEntry(data = {}, { line = state } = {}) {
  const activeLine = line || state;
  const shiftLogId = String(data.shiftLogId || "").trim();
  const breakStart = String(data.breakStart || "").trim();
  const breakFinish = String(data.breakFinish || "").trim();
  if (!UUID_RE.test(shiftLogId)) {
    throw new Error("Select a saved shift record before adding a break.");
  }
  if (!strictTimeValid(breakStart) || !strictTimeValid(breakFinish)) {
    throw new Error("Break start and finish must be HH:MM (24h).");
  }
  const parentShift = (Array.isArray(activeLine?.shiftRows) ? activeLine.shiftRows : [])
    .find((row) => String(row?.id || "").trim() === shiftLogId);
  if (!parentShift) {
    throw new Error("The selected shift record could not be found.");
  }
  const saved = await startManagerShiftBreak(shiftLogId, breakStart, breakFinish);
  if (!Array.isArray(activeLine.breakRows)) activeLine.breakRows = [];
  const row = {
    id: saved?.id || makeLocalLogId("break"),
    shiftLogId: saved?.shiftLogId || shiftLogId,
    date: String(saved?.date || parentShift.date || data.date || "").trim(),
    shift: String(saved?.shift || parentShift.shift || "").trim(),
    breakStart,
    breakFinish,
    submittedBy: "manager",
    submittedAt: saved?.submittedAt || nowIso()
  };
  activeLine.breakRows.push(row);
  addAudit(activeLine, "MANAGER_BREAK_LOG", `Manager logged ${row.shift} break for ${row.date}`);
  saveState();
  return row;
}

async function createManagerRunLogEntry(
  data = {},
  {
    line = state,
    runCrewingPatternInput = null,
    runCrewingPatternSummary = null,
    selectedShift = state?.selectedShift || "Day"
  } = {}
) {
  const activeLine = line || state;
  const lineId = String(activeLine?.id || "").trim();
  if (!lineId) throw new Error("No production line selected.");
  const next = {
    date: String(data.date || "").trim(),
    product: catalogProductCanonicalName(String(data.product || "").trim()),
    productionStartTime: String(data.productionStartTime || "").trim(),
    finishTime: String(data.finishTime || "").trim(),
    unitsProduced: num(data.unitsProduced),
    notes: String(data.notes || "").trim()
  };
  if (!isOperationalDate(next.date)) {
    throw new Error("Date is required. Weekend dates are excluded.");
  }
  if (!next.product) {
    throw new Error("Product is required.");
  }
  if (!isCatalogProductName(next.product)) {
    throw new Error("Select a product from Manage Products before starting a run.");
  }
  if (!strictTimeValid(next.productionStartTime) || !strictTimeValid(next.finishTime)) {
    throw new Error("Production start and finish must be HH:MM (24h).");
  }
  if (next.unitsProduced < 0) {
    throw new Error("Units produced cannot be negative.");
  }
  const patternShift = preferredTimedLogShift(
    activeLine,
    next.date,
    next.productionStartTime,
    next.finishTime,
    fallbackShiftValue(selectedShift)
  );
  const rawPattern = Object.prototype.hasOwnProperty.call(data, "runCrewingPattern")
    ? data.runCrewingPattern
    : runCrewingPatternInput?.value || "";
  const runCrewingPattern = normalizeRunCrewingPattern(rawPattern, activeLine, patternShift, { fallbackToIdeal: false });
  if (!Object.keys(runCrewingPattern).length) {
    throw new Error("Set crewing pattern for this run before saving.");
  }
  if (runCrewingPatternInput || runCrewingPatternSummary) {
    setRunCrewingPatternField(
      runCrewingPatternInput,
      runCrewingPatternSummary,
      activeLine,
      patternShift,
      runCrewingPattern,
      { fallbackToIdeal: false }
    );
  }
  const saved = await syncManagerRunLog({
    lineId,
    date: next.date,
    product: next.product,
    setUpStartTime: "",
    productionStartTime: next.productionStartTime,
    finishTime: next.finishTime,
    unitsProduced: next.unitsProduced,
    notes: next.notes,
    runCrewingPattern
  });
  if (!Array.isArray(activeLine.runRows)) activeLine.runRows = [];
  const row = {
    ...next,
    setUpStartTime: "",
    assignedShift: "",
    shift: "",
    runCrewingPattern: normalizeRunCrewingPattern(saved?.runCrewingPattern || runCrewingPattern, activeLine, patternShift, { fallbackToIdeal: false }),
    id: saved?.id || data.id || makeLocalLogId("run"),
    submittedBy: "manager",
    submittedAt: saved?.submittedAt || nowIso()
  };
  activeLine.runRows.push(row);
  addAudit(activeLine, "MANAGER_RUN_LOG", `Manager logged run ${next.product} (${next.unitsProduced} units)`);
  saveState();
  return row;
}

async function createManagerDowntimeLogEntry(data = {}, { line = state } = {}) {
  const activeLine = line || state;
  const lineId = String(activeLine?.id || "").trim();
  if (!lineId) throw new Error("No production line selected.");
  const next = {
    date: String(data.date || "").trim(),
    downtimeStart: String(data.downtimeStart || "").trim(),
    downtimeFinish: String(data.downtimeFinish || "").trim(),
    reasonCategory: String(data.reasonCategory || "").trim(),
    reasonDetail: String(data.reasonDetail || "").trim(),
    reasonNote: String(data.reasonNote || "").trim(),
    notes: String(data.notes || "").trim()
  };
  if (!isOperationalDate(next.date)) {
    throw new Error("Date is required. Weekend dates are excluded.");
  }
  if (!strictTimeValid(next.downtimeStart) || !strictTimeValid(next.downtimeFinish)) {
    throw new Error("Downtime start and finish must be HH:MM (24h).");
  }
  if (!next.reasonCategory) {
    throw new Error("Reason group is required.");
  }
  if (!next.reasonDetail) {
    throw new Error("Reason detail is required.");
  }
  next.equipment = next.reasonCategory === "Equipment" ? next.reasonDetail : "";
  next.reason = buildDowntimeReasonText(activeLine, next.reasonCategory, next.reasonDetail, next.reasonNote);
  const saved = await syncManagerDowntimeLog({
    lineId,
    date: next.date,
    downtimeStart: next.downtimeStart,
    downtimeFinish: next.downtimeFinish,
    equipment: next.equipment,
    reason: next.reason || "",
    notes: next.notes
  });
  if (!Array.isArray(activeLine.downtimeRows)) activeLine.downtimeRows = [];
  const row = {
    ...next,
    assignedShift: "",
    shift: "",
    excludeFromCalculation: normalizeDowntimeCalculationFlag(saved?.excludeFromCalculation),
    id: saved?.id || data.id || makeLocalLogId("down"),
    submittedBy: "manager",
    submittedAt: saved?.submittedAt || nowIso()
  };
  activeLine.downtimeRows.push(row);
  addAudit(
    activeLine,
    "MANAGER_DOWNTIME_LOG",
    `Manager logged downtime on ${managerDowntimeAuditLabel(activeLine, next)}`
  );
  saveState();
  return row;
}

async function saveLineModelToBackend(lineId) {
  const session = await ensureManagerBackendSession();
  const localLineId = String(lineId || "").trim();
  if (!localLineId) return;
  const line = appState.lines[localLineId];
  if (!line) return;
  const backendLineId = UUID_RE.test(localLineId) ? localLineId : await ensureBackendLineId(localLineId, session);
  if (!backendLineId) throw new Error("Line is not synced to server.");
  const stages = (line.stages || []).map((stage, index) => ({
    stageOrder: index + 1,
    stageName: stageBaseName(stage.name) || "Stage",
    stageType: toBackendStageType(stage),
    dayCrew: Math.max(0, num(line?.crewsByShift?.Day?.[stage.id]?.crew ?? stage.crew)),
    nightCrew: Math.max(0, num(line?.crewsByShift?.Night?.[stage.id]?.crew ?? stage.crew)),
    maxThroughputPerCrew: Math.max(0, num(line?.stageSettings?.[stage.id]?.maxThroughput)),
    dataSourceId: UUID_RE.test(String(stage?.dataSourceId || "").trim()) ? String(stage.dataSourceId).trim() : "",
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
  await apiRequest(`/api/lines/${backendLineId}/model`, {
    method: "PUT",
    token: session.backendToken,
    body: { stages, guides }
  });
}

async function createDataSourceOnBackend(payload) {
  const session = await ensureManagerBackendSession();
  return apiRequest("/api/data-sources", {
    method: "POST",
    token: session.backendToken,
    body: payload
  });
}

async function testDataSourceConnectionOnBackend(payload) {
  const session = await ensureManagerBackendSession();
  return apiRequest("/api/data-sources/test", {
    method: "POST",
    token: session.backendToken,
    body: payload,
    timeoutMs: 20000
  });
}

async function testSavedDataSourceConnectionOnBackend(dataSourceId) {
  const session = await ensureManagerBackendSession();
  const safeDataSourceId = String(dataSourceId || "").trim();
  if (!UUID_RE.test(safeDataSourceId)) {
    throw new Error("Invalid data source id.");
  }
  return apiRequest(`/api/data-sources/${encodeURIComponent(safeDataSourceId)}/test`, {
    method: "POST",
    token: session.backendToken,
    timeoutMs: 20000
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
  lineModel.groupId = String(created?.line?.groupId || "").trim();
  const nextOrder = Number(created?.line?.displayOrder);
  if (Number.isFinite(nextOrder)) lineModel.displayOrder = Math.max(0, Math.floor(nextOrder));
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
  const validationError = validateDowntimeBackendPayload(payload);
  if (validationError) throw new Error(validationError);
  const backendLineId = await ensureBackendLineId(payload.lineId, session);
  if (!backendLineId) throw new Error("Line is not synced to server.");
  const backendEquipmentId = payload.equipment
    ? await ensureBackendStageId(payload.lineId, payload.equipment, session)
    : null;
  const safeEquipmentStageId = UUID_RE.test(String(backendEquipmentId || "").trim()) ? backendEquipmentId : null;
  const body = {
    lineId: backendLineId,
    date: payload.date,
    downtimeStart: payload.downtimeStart,
    downtimeFinish: payload.downtimeFinish,
    equipmentStageId: safeEquipmentStageId,
    reason: payload.reason || "",
    notes: String(payload.notes || "")
  };
  if (Object.prototype.hasOwnProperty.call(payload, "excludeFromCalculation")) {
    body.excludeFromCalculation = normalizeDowntimeCalculationFlag(payload.excludeFromCalculation);
  }
  const response = await apiRequest("/api/logs/downtime", {
    method: "POST",
    token: session.backendToken,
    body
  });
  return response?.downtimeLog || null;
}

async function patchSupervisorDowntimeLog(session, logId, payload) {
  const validationError = validateDowntimeBackendPayload(payload);
  if (validationError) throw new Error(validationError);
  const backendEquipmentId = payload.equipment
    ? await ensureBackendStageId(payload.lineId, payload.equipment, session)
    : null;
  const safeEquipmentStageId = UUID_RE.test(String(backendEquipmentId || "").trim()) ? backendEquipmentId : null;
  const body = {
    downtimeStart: payload.downtimeStart,
    downtimeFinish: payload.downtimeFinish,
    equipmentStageId: safeEquipmentStageId,
    reason: payload.reason || "",
    notes: String(payload.notes || "")
  };
  if (Object.prototype.hasOwnProperty.call(payload, "excludeFromCalculation")) {
    body.excludeFromCalculation = normalizeDowntimeCalculationFlag(payload.excludeFromCalculation);
  }
  const response = await apiRequest(`/api/logs/downtime/${logId}`, {
    method: "PATCH",
    token: session.backendToken,
    body
  });
  return response?.downtimeLog || null;
}

async function syncSupervisorAction(session, payload) {
  const localLineId = String(payload.lineId || "").trim();
  const backendLineId = localLineId ? await ensureBackendLineId(localLineId, session) : "";
  if (localLineId && !backendLineId) throw new Error("Line is not synced to server.");
  const localEquipmentId = String(payload.relatedEquipmentId || "").trim();
  const backendEquipmentId = localEquipmentId
    ? await ensureBackendStageId(localLineId, localEquipmentId, session)
    : "";
  if (localEquipmentId && !backendEquipmentId) throw new Error("Related equipment could not be mapped to the backend line.");
  const response = await apiRequest("/api/supervisor-actions", {
    method: "POST",
    token: session.backendToken,
    body: {
      lineId: backendLineId || "",
      title: String(payload.title || ""),
      description: String(payload.description || ""),
      priority: normalizeActionPriority(payload.priority),
      status: normalizeActionStatus(payload.status),
      dueDate: /^\d{4}-\d{2}-\d{2}$/.test(String(payload.dueDate || "").trim()) ? String(payload.dueDate).trim() : "",
      relatedEquipmentId: backendEquipmentId || "",
      relatedReasonCategory: normalizeActionReasonCategory(payload.relatedReasonCategory),
      relatedReasonDetail: String(payload.relatedReasonDetail || "").trim()
    }
  });
  return response?.action || null;
}

async function patchManagerSupervisorAction(actionId, payload) {
  const session = await ensureManagerBackendSession();
  const localLineId = String(payload.lineId || "").trim();
  const backendLineId = localLineId ? await ensureBackendLineId(localLineId, session) : "";
  if (localLineId && !backendLineId) throw new Error("Line is not synced to server.");
  const localEquipmentId = String(payload.relatedEquipmentId || "").trim();
  const backendEquipmentId = localEquipmentId
    ? await ensureBackendStageId(localLineId, localEquipmentId, session)
    : "";
  if (localEquipmentId && !backendEquipmentId) throw new Error("Related equipment could not be mapped to the backend line.");
  const response = await apiRequest(`/api/supervisor-actions/${encodeURIComponent(String(actionId || ""))}`, {
    method: "PATCH",
    token: session.backendToken,
    body: {
      supervisorUsername: String(payload.supervisorUsername || "").trim().toLowerCase(),
      supervisorName: String(payload.supervisorName || "").trim(),
      lineId: backendLineId || "",
      priority: normalizeActionPriority(payload.priority),
      status: normalizeActionStatus(payload.status),
      dueDate: /^\d{4}-\d{2}-\d{2}$/.test(String(payload.dueDate || "").trim()) ? String(payload.dueDate).trim() : "",
      relatedEquipmentId: backendEquipmentId || "",
      relatedReasonCategory: normalizeActionReasonCategory(payload.relatedReasonCategory),
      relatedReasonDetail: String(payload.relatedReasonDetail || "").trim()
    }
  });
  return response?.action || null;
}

async function deleteManagerSupervisorAction(actionId) {
  const session = await ensureManagerBackendSession();
  await apiRequest(`/api/supervisor-actions/${encodeURIComponent(String(actionId || ""))}`, {
    method: "DELETE",
    token: session.backendToken
  });
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
    const seenIds = new Set();
    rows.forEach((row) => {
      if (!row || typeof row !== "object") return;
      let rowId = String(row.id || "").trim();
      if (!rowId || seenIds.has(rowId)) {
        do {
          rowId = makeLocalLogId(prefix);
        } while (seenIds.has(rowId));
        row.id = rowId;
      }
      seenIds.add(rowId);
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
    `<option value="">${htmlEscape(placeholder)}</option>`,
    ...options.map((option) => `<option value="${htmlEscape(option.value)}">${htmlEscape(option.label)}</option>`)
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

function normalizeRunLogRow(row) {
  return {
    ...(row || {}),
    assignedShift: String(row?.assignedShift || ""),
    shift: String(row?.shift || ""),
    notes: String(row?.notes || ""),
    runCrewingPattern: parseRunCrewingPattern(row?.runCrewingPattern)
  };
}

function validateDowntimeBackendPayload(payload) {
  const date = String(payload?.date || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return "Date is invalid.";
  const downtimeStart = String(payload?.downtimeStart || "").trim();
  const downtimeFinish = String(payload?.downtimeFinish || "").trim();
  if (!strictTimeValid(downtimeStart) || !strictTimeValid(downtimeFinish)) {
    return "Downtime start and finish must be HH:MM (24h).";
  }
  const reason = String(payload?.reason || "").trim();
  if (!reason) return "Reason is required.";
  if (reason.length > BACKEND_DOWNTIME_REASON_MAX_LENGTH) {
    return `Reason is too long (${reason.length}/${BACKEND_DOWNTIME_REASON_MAX_LENGTH}). Shorten Reason Notes.`;
  }
  const notes = String(payload?.notes || "").trim();
  if (notes.length > BACKEND_LOG_NOTES_MAX_LENGTH) {
    return `Notes are too long (${notes.length}/${BACKEND_LOG_NOTES_MAX_LENGTH}).`;
  }
  return "";
}

function normalizeDowntimeCalculationFlag(value) {
  return value === true || value === "true" || value === 1 || value === "1";
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

function normalizeDowntimeLogRow(row) {
  const parsedReason = parseDowntimeReasonParts(row?.reason, row?.equipment);
  return {
    ...(row || {}),
    assignedShift: String(row?.assignedShift || ""),
    shift: String(row?.shift || ""),
    notes: String(row?.notes || ""),
    reasonCategory: row?.reasonCategory || parsedReason.reasonCategory,
    reasonDetail: row?.reasonDetail || parsedReason.reasonDetail,
    reasonNote: row?.reasonNote || parsedReason.reasonNote,
    excludeFromCalculation: normalizeDowntimeCalculationFlag(row?.excludeFromCalculation)
  };
}

function isDowntimeExcludedFromCalculation(row) {
  return normalizeDowntimeCalculationFlag(row?.excludeFromCalculation);
}

function calculationDowntimeRows(rows = []) {
  return (rows || []).filter((row) => !isDowntimeExcludedFromCalculation(row));
}

function downtimeMinutesForCalculations(row, shift) {
  if (isDowntimeExcludedFromCalculation(row)) return 0;
  return num(row?.downtimeMins) * timedLogShiftWeight(row, shift);
}

function downtimeReasonCategoryFromRow(row) {
  const explicitCategory = String(row?.reasonCategory || "").trim();
  if (explicitCategory) return explicitCategory;
  const parsed = parseDowntimeReasonParts(row?.reason, row?.equipment);
  return String(parsed?.reasonCategory || "").trim();
}

function isBreakReasonCategory(category) {
  return String(category || "").trim().toLowerCase() === "break";
}

function isDowntimeBreakRow(row) {
  return isBreakReasonCategory(downtimeReasonCategoryFromRow(row));
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
        startTime: "06:00",
        finishTime: "14:00"
      },
      {
        date,
        shift: "Night",
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
        product: "Teriyaki",
        productionStartTime: "06:10",
        finishTime: "10:35",
        unitsProduced: Math.round(2850 * dayTrend)
      },
      {
        date,
        product: "Honey Soy",
        productionStartTime: "10:55",
        finishTime: "15:25",
        unitsProduced: Math.round(2600 * dayTrend)
      },
      {
        date,
        product: "Peri Peri",
        productionStartTime: "14:15",
        finishTime: "18:55",
        unitsProduced: Math.round(2500 * nightTrend)
      },
      {
        date,
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
  line.selectedDate = normalizeWeekdayIsoDate(sample.selectedDate, { direction: -1 });
  line.selectedShift = sample.selectedShift;
  line.trendGranularity = sample.trendGranularity;
  line.trendMonth = sample.trendMonth;
  line.trendDateCursor = line.selectedDate;
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
  return {
    ...row,
    downtimeMins: calc > 0 ? calc : fallback,
    excludeFromCalculation: normalizeDowntimeCalculationFlag(row?.excludeFromCalculation)
  };
}

function computeShiftRow(row) {
  const fallback = num(row.totalShiftTime);
  const calc = diffMinutes(row.startTime, row.finishTime);
  return { ...row, totalShiftTime: calc > 0 ? calc : fallback };
}

function computeBreakRow(row) {
  const fallback = num(row.breakMins);
  const calc = diffMinutes(row.breakStart, row.breakFinish);
  return { ...row, breakMins: calc > 0 ? calc : fallback };
}

function deriveBreakRowFromDowntimeRow(row, line = state) {
  const rowId = String(row?.id || "").trim();
  const breakStart = String(row?.downtimeStart || "").trim();
  const breakFinishRaw = String(row?.downtimeFinish || "").trim();
  const isOpenDowntime = strictTimeValid(breakStart) && strictTimeValid(breakFinishRaw) && breakStart === breakFinishRaw;
  const breakFinish = isOpenDowntime ? "" : breakFinishRaw;
  const assignedShift = preferredTimedLogShift(
    line,
    String(row?.date || "").trim(),
    breakStart,
    breakFinishRaw || breakStart,
    line?.selectedShift || "Day"
  );
  const fallbackId = `${String(row?.date || "").trim()}-${breakStart || "start"}-${breakFinishRaw || "finish"}`;
  return {
    ...row,
    id: `down-break-${rowId || fallbackId}`,
    shift: assignedShift || String(row?.shift || ""),
    assignedShift: assignedShift || String(row?.assignedShift || ""),
    breakStart,
    breakFinish,
    breakMins: Math.max(0, num(row?.downtimeMins)),
    isDerivedDowntimeBreak: true,
    sourceDowntimeId: rowId
  };
}

function splitDowntimeRowsForBreakViews(sourceRows = [], line = state) {
  const downtimeRowsLogged = [];
  const downtimeRows = [];
  const breakRowsFromDowntime = [];
  (sourceRows || []).forEach((row) => {
    if (!row) return;
    downtimeRowsLogged.push(row);
    if (isDowntimeBreakRow(row)) {
      breakRowsFromDowntime.push(deriveBreakRowFromDowntimeRow(row, line));
      return;
    }
    downtimeRows.push(row);
  });
  return { downtimeRowsLogged, downtimeRows, breakRowsFromDowntime };
}

function breakMigrationReasonDetail(row) {
  const reasonText = [
    row?.notes,
    row?.reason,
    row?.reasonDetail,
    row?.reasonNote
  ]
    .map((value) => String(value || "").trim().toLowerCase())
    .filter(Boolean)
    .join(" ");
  return /non[\s-]?standard/.test(reasonText) ? "Non-standard Break" : "Standard Break";
}

function buildDowntimeBreakMigrationPayload(line, breakRow) {
  const date = String(breakRow?.date || "").trim();
  const breakStart = String(breakRow?.breakStart || "").trim();
  if (!isOperationalDate(date) || !strictTimeValid(breakStart)) return null;
  const breakFinishRaw = String(breakRow?.breakFinish || "").trim();
  const downtimeFinish = strictTimeValid(breakFinishRaw) ? breakFinishRaw : breakStart;
  const reasonCategory = "Break";
  const reasonDetail = breakMigrationReasonDetail(breakRow);
  const reasonNote = "";
  const reason = buildDowntimeReasonText(line, reasonCategory, reasonDetail, reasonNote);
  const notes = String(breakRow?.notes || "").trim();
  return {
    lineId: String(line?.id || "").trim(),
    date,
    downtimeStart: breakStart,
    downtimeFinish,
    equipment: "",
    reason,
    notes,
    reasonCategory,
    reasonDetail,
    reasonNote
  };
}

function downtimeBreakRowFromMigration(line, breakRow, payload, savedDowntime = null) {
  const equipment = String(savedDowntime?.equipment || "").trim();
  const reason = String(savedDowntime?.reason || payload?.reason || "").trim();
  const parsedReason = parseDowntimeReasonParts(reason, equipment);
  const reasonCategory = String(parsedReason.reasonCategory || payload?.reasonCategory || "Break").trim() || "Break";
  const reasonDetail = String(parsedReason.reasonDetail || payload?.reasonDetail || "Standard Break").trim() || "Standard Break";
  const reasonNote = String(parsedReason.reasonNote || payload?.reasonNote || "").trim();
  return {
    ...(breakRow || {}),
    id: String(savedDowntime?.id || "").trim() || makeLocalLogId("down"),
    lineId: String(line?.id || "").trim() || String(payload?.lineId || "").trim(),
    date: String(savedDowntime?.date || payload?.date || "").trim(),
    assignedShift: "",
    shift: "",
    downtimeStart: String(savedDowntime?.downtimeStart || payload?.downtimeStart || "").trim(),
    downtimeFinish: String(savedDowntime?.downtimeFinish || payload?.downtimeFinish || "").trim(),
    downtimeMins: Math.max(
      0,
      num(savedDowntime?.downtimeMins) || diffMinutes(savedDowntime?.downtimeStart || payload?.downtimeStart, savedDowntime?.downtimeFinish || payload?.downtimeFinish)
    ),
    equipment,
    reason,
    reasonCategory,
    reasonDetail,
    reasonNote,
    excludeFromCalculation: normalizeDowntimeCalculationFlag(savedDowntime?.excludeFromCalculation),
    notes: String(savedDowntime?.notes ?? payload?.notes ?? "").trim(),
    submittedBy: String(savedDowntime?.submittedBy || breakRow?.submittedBy || "manager").trim() || "manager",
    submittedByUserId: String(savedDowntime?.submittedByUserId || breakRow?.submittedByUserId || "").trim(),
    submittedAt: savedDowntime?.submittedAt || nowIso()
  };
}

function computeRunRow(row, nonProductionIntervalsByDate) {
  const grossFallback = num(row.grossProductionTime);
  const grossCalc = diffMinutes(row.productionStartTime, row.finishTime);
  const grossProductionTime = grossCalc > 0 ? grossCalc : grossFallback;

  const associatedFallback = num(row.associatedDownTime);
  const runIntervals = intervalsFromTimes(row.productionStartTime, row.finishTime);
  const runDate = String(row?.date || "").trim();
  const nonProductionIntervals = nonProductionIntervalsByDate?.get(runDate) || [];
  const hasComputedOverlap = runIntervals.length > 0;
  const associatedDownTime = hasComputedOverlap
    ? overlapMinutesBetweenIntervalSets(runIntervals, nonProductionIntervals)
    : associatedFallback;

  const netFallback = num(row.netProductionTime);
  const netCalc = Math.max(0, grossProductionTime - associatedDownTime);
  const netProductionTime = hasComputedOverlap ? netCalc : netFallback;

  const unitsProduced = num(row.unitsProduced);
  const grossRunRate = grossProductionTime > 0 ? unitsProduced / grossProductionTime : num(row.grossRunRate);
  const netRunRate = netProductionTime > 0 ? unitsProduced / netProductionTime : hasComputedOverlap ? 0 : num(row.netRunRate);

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
  const operationalDowntimeRows = state.downtimeRows.filter((row) => isOperationalDate(String(row?.date || "")));
  const operationalShiftRows = state.shiftRows.filter((row) => isOperationalDate(String(row?.date || "")));
  const operationalBreakRows = (state.breakRows || []).filter((row) => isOperationalDate(String(row?.date || "")));
  const operationalRunRows = state.runRows.filter((row) => isOperationalDate(String(row?.date || "")));
  const computedDowntimeRows = operationalDowntimeRows
    .map(computeDowntimeRow)
    .map((row) => decorateTimedLogShift(row, state, "downtimeStart", "downtimeFinish"));
  const { downtimeRowsLogged, downtimeRows, breakRowsFromDowntime } = splitDowntimeRowsForBreakViews(computedDowntimeRows, state);
  const shiftRows = operationalShiftRows.map(computeShiftRow);
  const breakRows = [...operationalBreakRows.map(computeBreakRow), ...breakRowsFromDowntime.map(computeBreakRow)];
  const nonProductionIntervalsByDate = buildNonProductionIntervalsByDate(breakRows, calculationDowntimeRows(downtimeRows));
  const runRows = operationalRunRows
    .map((row) => decorateTimedLogShift(row, state, "productionStartTime", "finishTime"))
    .map((row) => computeRunRow(row, nonProductionIntervalsByDate));
  return { shiftRows, breakRows, runRows, downtimeRows, downtimeRowsLogged };
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
  document.querySelectorAll(".tab-btn[data-tab]").forEach((btn) => {
    const tabId = parseManagerLineTabId(btn.dataset.tab);
    if (!tabId) return;
    btn.addEventListener("click", () => {
      setActiveManagerLineTab(tabId);
      saveState();
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
  const supervisorLoginForm = document.getElementById("supervisorLoginForm");
  const supervisorUserInput = document.getElementById("supervisorUser");
  const supervisorPassInput = document.getElementById("supervisorPass");
  const supervisorTestLoginBtn = document.getElementById("supervisorTestLoginBtn");
  const supervisorLogoutBtn = document.getElementById("supervisorLogout");
  const managerSettingsTabBtn = document.getElementById("managerSettingsTabBtn");
  const supervisorMobileModeBtn = document.getElementById("supervisorMobileMode");
  const supervisorLineSelect = document.getElementById("supervisorLineSelect");
  const svDateInputs = Array.from(document.querySelectorAll("[data-sv-date]"));
  const svShiftButtons = Array.from(document.querySelectorAll("[data-sv-shift]"));
  const svPrevBtns = Array.from(document.querySelectorAll("[data-sv-prev]"));
  const svNextBtns = Array.from(document.querySelectorAll("[data-sv-next]"));
  const supervisorShiftForm = document.getElementById("supervisorShiftForm");
  const supervisorRunForm = document.getElementById("supervisorRunForm");
  const supervisorDownForm = document.getElementById("supervisorDownForm");
  const supervisorActionForm = document.getElementById("supervisorActionForm");
  const supervisorActionLineInput = document.getElementById("supervisorActionLine");
  const supervisorActionPriorityInput = document.getElementById("supervisorActionPriority");
  const supervisorActionStatusInput = document.getElementById("supervisorActionStatus");
  const supervisorActionDueDateInput = document.getElementById("supervisorActionDueDate");
  const supervisorActionEquipmentInput = document.getElementById("supervisorActionEquipment");
  const supervisorActionReasonCategoryInput = document.getElementById("supervisorActionReasonCategory");
  const supervisorActionReasonDetailInput = document.getElementById("supervisorActionReasonDetail");
  const supervisorActionTitleInput = document.getElementById("supervisorActionTitle");
  const supervisorActionDescriptionInput = document.getElementById("supervisorActionDescription");
  const superShiftLogIdInput = document.getElementById("superShiftLogId");
  const superRunLogIdInput = document.getElementById("superRunLogId");
  const superRunCrewingPatternInput = document.getElementById("superRunCrewingPattern");
  const superRunCrewingPatternBtn = document.getElementById("superRunCrewingPatternBtn");
  const superRunCrewingPatternSummary = document.getElementById("superRunCrewingPatternSummary");
  const superDownLogIdInput = document.getElementById("superDownLogId");
  const superShiftSaveProgressBtn = document.getElementById("superShiftSaveProgress");
  const superRunSaveProgressBtn = document.getElementById("superRunSaveProgress");
  const superDownSaveProgressBtn = document.getElementById("superDownSaveProgress");
  const supervisorDownReasonCategory = document.getElementById("superDownReasonCategory");
  const supervisorDownReasonDetail = document.getElementById("superDownReasonDetail");
  const supervisorEntryList = document.getElementById("supervisorEntryList");
  const supervisorEntryCards = document.getElementById("supervisorEntryCards");
  const managerActionList = document.getElementById("managerActionList");
  const manageSupervisorsBtn = document.getElementById("manageSupervisorsBtn");
  const manageLineGroupsBtn = document.getElementById("manageLineGroupsBtn");
  const manageProductCatalogBtn = document.getElementById("manageProductCatalogBtn");
  const connectDataSourceBtn = document.getElementById("connectDataSourceBtn");
  const dataSourcesList = document.getElementById("dataSourcesList");
  const incomingDataSourcesStatus = document.getElementById("incomingDataSourcesStatus");
  const addManagerBtn = document.getElementById("addManagerBtn");
  const addSupervisorBtn = document.getElementById("addSupervisorBtn");
  const manageSupervisorsModal = document.getElementById("manageSupervisorsModal");
  const closeManageSupervisorsModalBtn = document.getElementById("closeManageSupervisorsModal");
  const supervisorManagerList = document.getElementById("supervisorManagerList");
  const manageLineGroupsModal = document.getElementById("manageLineGroupsModal");
  const closeManageLineGroupsModalBtn = document.getElementById("closeManageLineGroupsModal");
  const manageProductCatalogModal = document.getElementById("manageProductCatalogModal");
  const closeManageProductCatalogModalBtn = document.getElementById("closeManageProductCatalogModal");
  const manageProductCatalogContent = document.getElementById("manageProductCatalogContent");
  const lineGroupManagerList = document.getElementById("lineGroupManagerList");
  const lineGroupCreateForm = document.getElementById("lineGroupCreateForm");
  const newLineGroupNameInput = document.getElementById("newLineGroupName");
  const addManagerModal = document.getElementById("addManagerModal");
  const closeAddManagerModalBtn = document.getElementById("closeAddManagerModal");
  const addManagerForm = document.getElementById("addManagerForm");
  const newManagerNameInput = document.getElementById("newManagerName");
  const newManagerUsernameInput = document.getElementById("newManagerUsername");
  const newManagerPasswordInput = document.getElementById("newManagerPassword");
  const addSupervisorModal = document.getElementById("addSupervisorModal");
  const closeAddSupervisorModalBtn = document.getElementById("closeAddSupervisorModal");
  const addSupervisorForm = document.getElementById("addSupervisorForm");
  const newSupervisorLines = document.getElementById("newSupervisorLines");
  const editSupervisorModal = document.getElementById("editSupervisorModal");
  const closeEditSupervisorModalBtn = document.getElementById("closeEditSupervisorModal");
  const editSupervisorForm = document.getElementById("editSupervisorForm");
  const editSupervisorNameInput = document.getElementById("editSupervisorName");
  const editSupervisorUsernameInput = document.getElementById("editSupervisorUsername");
  const editLineModal = document.getElementById("editLineModal");
  const closeEditLineModalBtn = document.getElementById("closeEditLineModal");
  const editLineForm = document.getElementById("editLineForm");
  const editLineNameInput = document.getElementById("editLineName");
  const editLineDeleteBtn = document.getElementById("editLineDeleteBtn");
  const connectDataSourceModal = document.getElementById("connectDataSourceModal");
  const closeConnectDataSourceModalBtn = document.getElementById("closeConnectDataSourceModal");
  const connectDataSourceForm = document.getElementById("connectDataSourceForm");
  const connectDataSourceNameInput = document.getElementById("connectDataSourceName");
  const connectDataSourceKeyInput = document.getElementById("connectDataSourceKey");
  const connectDataSourceProviderInput = document.getElementById("connectDataSourceProvider");
  const connectDataSourceModeInput = document.getElementById("connectDataSourceMode");
  const connectDataSourceMachineNoInput = document.getElementById("connectDataSourceMachineNo");
  const connectDataSourceDeviceNameInput = document.getElementById("connectDataSourceDeviceName");
  const connectDataSourceDeviceIdInput = document.getElementById("connectDataSourceDeviceId");
  const connectDataSourceScaleNumberInput = document.getElementById("connectDataSourceScaleNumber");
  const connectDataSourceApiFields = document.getElementById("connectDataSourceApiFields");
  const connectDataSourceApiBaseUrlInput = document.getElementById("connectDataSourceApiBaseUrl");
  const connectDataSourceApiKeyInput = document.getElementById("connectDataSourceApiKey");
  const connectDataSourceApiSecretInput = document.getElementById("connectDataSourceApiSecret");
  const connectDataSourceSqlFields = document.getElementById("connectDataSourceSqlFields");
  const connectDataSourceSqlHostInput = document.getElementById("connectDataSourceSqlHost");
  const connectDataSourceSqlPortInput = document.getElementById("connectDataSourceSqlPort");
  const connectDataSourceSqlDatabaseInput = document.getElementById("connectDataSourceSqlDatabase");
  const connectDataSourceSqlUsernameInput = document.getElementById("connectDataSourceSqlUsername");
  const connectDataSourceSqlPasswordInput = document.getElementById("connectDataSourceSqlPassword");
  const connectDataSourceHelp = document.getElementById("connectDataSourceHelp");
  const connectDataSourceTestResult = document.getElementById("connectDataSourceTestResult");
  const connectDataSourceTestBtn = document.getElementById("connectDataSourceTestBtn");
  const productCatalogTable = document.getElementById("productCatalogTable");
  const addProductBtn = document.getElementById("addProductBtn");
  const productLineAssignModal = document.getElementById("productLineAssignModal");
  const closeProductLineAssignModalBtn = document.getElementById("closeProductLineAssignModal");
  const productLineAssignProductLabel = document.getElementById("productLineAssignProduct");
  const productLineAssignList = document.getElementById("productLineAssignList");
  const productLineAssignSelectAllBtn = document.getElementById("productLineAssignSelectAll");
  const productLineAssignClearAllBtn = document.getElementById("productLineAssignClearAll");
  const productLineAssignSaveBtn = document.getElementById("productLineAssignSave");
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
  let assigningProductRowId = "";
  let suppressSupervisorSelectionReset = false;
  const PRODUCT_CATALOG_EDITABLE_COLUMN_COUNT = 12;
  const PRODUCT_CATALOG_NUMERIC_COLUMNS = new Set([6, 7, 9, 10, 11, 12]);

  const resetSupervisorShiftFormSelection = () => {
    if (superShiftLogIdInput) superShiftLogIdInput.value = "";
    supervisorShiftTileEditId = "";
    supervisorShiftForm?.reset();
    setSupervisorAutoDateValue("superShiftDate", supervisorAutoEntryDate(), { fallbackIso: supervisorAutoEntryDate() });
    const lineShiftMap = normalizeSupervisorLineShifts(
      appState.supervisorSession?.assignedLineShifts,
      appState.lines,
      appState.supervisorSession?.assignedLineIds || []
    );
    const allowedShifts = expandedSupervisorShiftAccess(lineShiftMap[selectedSupervisorLineId()]);
    const shiftInput = document.getElementById("superShiftShift");
    if (shiftInput) shiftInput.value = allowedShifts[0] || "Day";
    syncAllNowInputPrefixStates(supervisorShiftForm);
  };

  const resetSupervisorRunFormSelection = () => {
    if (superRunLogIdInput) superRunLogIdInput.value = "";
    supervisorRunForm?.reset();
    setSupervisorAutoDateValue("superRunDate", supervisorAutoEntryDate(), { fallbackIso: supervisorAutoEntryDate() });
    setRunCrewingPatternField(
      superRunCrewingPatternInput,
      superRunCrewingPatternSummary,
      selectedSupervisorLine(),
      appState.supervisorSelectedShift || "Day",
      {},
      { fallbackToIdeal: false }
    );
    syncAllNowInputPrefixStates(supervisorRunForm);
  };

  const resetSupervisorDowntimeFormSelection = () => {
    if (superDownLogIdInput) superDownLogIdInput.value = "";
    supervisorDownForm?.reset();
    if (supervisorDownReasonCategory) supervisorDownReasonCategory.value = "";
    setSupervisorAutoDateValue("superDownDate", supervisorAutoEntryDate(), { fallbackIso: supervisorAutoEntryDate() });
    refreshSupervisorDowntimeDetailOptions(selectedSupervisorLine());
    syncAllNowInputPrefixStates(supervisorDownForm);
  };

  const clearSupervisorEditingSelection = ({ resetForms = false } = {}) => {
    if (resetForms) {
      resetSupervisorShiftFormSelection();
      resetSupervisorRunFormSelection();
      resetSupervisorDowntimeFormSelection();
    } else {
      if (superShiftLogIdInput) superShiftLogIdInput.value = "";
      supervisorShiftTileEditId = "";
      if (superRunLogIdInput) superRunLogIdInput.value = "";
      if (superDownLogIdInput) superDownLogIdInput.value = "";
    }
    updateSupervisorProgressButtonLabels();
  };

  const setProductRowEditingState = (rowNode, editing) => {
    if (!rowNode) return;
    rowNode.classList.toggle("product-row-editing", editing);
    rowNode.setAttribute("data-product-editing", editing ? "true" : "false");
    const editBtn = rowNode.querySelector("[data-product-row-edit]");
    if (editBtn) {
      editBtn.textContent = editing ? "Save" : "Edit";
      editBtn.classList.toggle("is-save", editing);
    }
  };

  const beginProductRowInlineEdit = (rowNode) => {
    if (!rowNode) return;
    const cells = Array.from(rowNode.cells || []).slice(0, PRODUCT_CATALOG_EDITABLE_COLUMN_COUNT);
    cells.forEach((cell, index) => {
      if (cell.querySelector("input[data-product-inline-input]")) return;
      const currentValue = String(cell.textContent || "").trim();
      const input = document.createElement("input");
      input.className = "table-inline-input";
      input.type = "text";
      input.value = currentValue;
      input.setAttribute("data-product-inline-input", String(index + 1));
      if (PRODUCT_CATALOG_NUMERIC_COLUMNS.has(index + 1)) input.setAttribute("inputmode", "decimal");
      cell.innerHTML = "";
      cell.append(input);
    });
    setProductRowEditingState(rowNode, true);
  };

  const finalizeProductRowInlineEdit = async (rowNode) => {
    if (!rowNode) return false;
    const cells = Array.from(rowNode.cells || []).slice(0, PRODUCT_CATALOG_EDITABLE_COLUMN_COUNT);
    const values = cells.map((cell) => {
      const input = cell.querySelector("input[data-product-inline-input]");
      return String(input?.value ?? cell.textContent ?? "").trim();
    });
    const descValue = String(values[1] || "").trim();
    if (!descValue) {
      const descInput = rowNode.querySelector('input[data-product-inline-input="2"]');
      alert("Desc 1 is required.");
      descInput?.focus();
      return false;
    }
    const rowId = ensureProductCatalogRowId(rowNode);
    let savedProduct = null;
    try {
      savedProduct = UUID_RE.test(rowId)
        ? await updateProductCatalogEntryOnBackend(rowId, values, getProductCatalogLineIdsFromRow(rowNode))
        : await createProductCatalogEntryOnBackend(values, getProductCatalogLineIdsFromRow(rowNode));
    } catch (error) {
      alert(`Could not save product.\n${error?.message || "Please try again."}`);
      return false;
    }
    const safeValues = Array.from({ length: PRODUCT_CATALOG_EDITABLE_COLUMN_COUNT }, (_, index) => String(savedProduct.values?.[index] || "").trim());
    cells.forEach((cell, index) => {
      cell.textContent = safeValues[index];
    });
    rowNode.setAttribute(PRODUCT_CATALOG_ID_ATTR, savedProduct.id);
    setProductCatalogLineIdsForRow(rowNode, savedProduct.lineIds);
    setProductRowEditingState(rowNode, false);
    syncProductCatalogStateFromTable(productCatalogTable);
    refreshProductRowAssignmentSummaries(productCatalogTable);
    syncRunProductInputsFromCatalog();
    return true;
  };

  const findProductCatalogRowById = (rowId) =>
    Array.from(productCatalogTable?.tBodies?.[0]?.rows || []).find(
      (rowNode) => String(rowNode.getAttribute(PRODUCT_CATALOG_ID_ATTR) || "") === String(rowId || "")
    ) || null;

  const buildProductCatalogRow = (values = [], lineIds = []) => {
    const rowNode = document.createElement("tr");
    const safeValues = Array.from({ length: PRODUCT_CATALOG_EDITABLE_COLUMN_COUNT }, (_, index) => String(values[index] || "").trim());
    safeValues.forEach((value) => {
      const cell = document.createElement("td");
      cell.textContent = value;
      rowNode.append(cell);
    });
    rowNode.setAttribute(PRODUCT_CATALOG_ID_ATTR, createProductCatalogRowId());
    setProductCatalogLineIdsForRow(rowNode, lineIds);
    return rowNode;
  };

  const closeProductLineAssignModal = () => {
    if (!productLineAssignModal) return;
    productLineAssignModal.classList.remove("open");
    productLineAssignModal.setAttribute("aria-hidden", "true");
    assigningProductRowId = "";
  };

  const openProductLineAssignModal = (rowNode) => {
    if (!rowNode || !productLineAssignModal || !productLineAssignList) return;
    const rowId = ensureProductCatalogRowId(rowNode);
    assigningProductRowId = rowId;
    const productLabel = String(productCatalogCellText(rowNode, 1) || productCatalogCellText(rowNode, 0) || "Unnamed Product").trim();
    if (productLineAssignProductLabel) {
      productLineAssignProductLabel.textContent = productLabel;
      productLineAssignProductLabel.title = productLabel;
    }
    const lines = sortLinesByDisplayOrder(Object.values(appState.lines || {}));
    const assignedLineIds = new Set(getProductCatalogLineIdsFromRow(rowNode));
    const isAllLines = assignedLineIds.has(PRODUCT_CATALOG_ALL_LINES_TOKEN);
    productLineAssignList.innerHTML = lines.length
      ? lines
          .map(
            (line) => `
              <label class="product-line-assign-row">
                <input type="checkbox" value="${htmlEscape(line.id)}" ${isAllLines || assignedLineIds.has(line.id) ? "checked" : ""} />
                <span>${htmlEscape(line.name)}</span>
              </label>
            `
          )
          .join("")
      : `<p class="muted">No production lines available.</p>`;
    productLineAssignModal.classList.add("open");
    productLineAssignModal.setAttribute("aria-hidden", "false");
  };

  const saveProductLineAssignments = async () => {
    if (!productCatalogTable || !productLineAssignList) return;
    const rowNode = findProductCatalogRowById(assigningProductRowId);
    if (!rowNode) {
      closeProductLineAssignModal();
      return;
    }
    const lines = sortLinesByDisplayOrder(Object.values(appState.lines || {}));
    const allLineIds = lines.map((line) => String(line.id || ""));
    const selectedLineIds = Array.from(productLineAssignList.querySelectorAll('input[type="checkbox"]:checked'))
      .map((input) => String(input.value || "").trim())
      .filter(Boolean);
    const normalizedSelected = normalizeProductCatalogLineIds(selectedLineIds);
    const nextLineIds =
      allLineIds.length && normalizedSelected.length === allLineIds.length
        ? [PRODUCT_CATALOG_ALL_LINES_TOKEN]
        : normalizedSelected;
    const rowId = ensureProductCatalogRowId(rowNode);
    if (!UUID_RE.test(rowId)) {
      alert("Save the product row before assigning lines.");
      return;
    }
    const values = Array.from({ length: PRODUCT_CATALOG_EDITABLE_COLUMN_COUNT }, (_, index) => productCatalogCellText(rowNode, index));
    try {
      const savedProduct = await updateProductCatalogEntryOnBackend(rowId, values, nextLineIds);
      const safeValues = Array.from({ length: PRODUCT_CATALOG_EDITABLE_COLUMN_COUNT }, (_, index) => String(savedProduct.values?.[index] || "").trim());
      const cells = Array.from(rowNode.cells || []).slice(0, PRODUCT_CATALOG_EDITABLE_COLUMN_COUNT);
      cells.forEach((cell, index) => {
        cell.textContent = safeValues[index];
      });
      rowNode.setAttribute(PRODUCT_CATALOG_ID_ATTR, savedProduct.id);
      setProductCatalogLineIdsForRow(rowNode, savedProduct.lineIds);
      syncProductCatalogStateFromTable(productCatalogTable);
      refreshProductRowAssignmentSummaries(productCatalogTable);
      syncRunProductInputsFromCatalog();
      closeProductLineAssignModal();
    } catch (error) {
      alert(`Could not save line assignments.\n${error?.message || "Please try again."}`);
    }
  };

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

  const selectedSupervisorActionLine = () => {
    const lineId = String(supervisorActionLineInput?.value || "").trim();
    return lineId && appState.lines[lineId] ? appState.lines[lineId] : null;
  };

  const refreshSupervisorActionRelationOptions = () => {
    const line = selectedSupervisorActionLine();
    if (supervisorActionEquipmentInput) {
      setActionEquipmentOptions(supervisorActionEquipmentInput, line, supervisorActionEquipmentInput.value || "");
    }
    if (supervisorActionReasonCategoryInput) {
      setActionReasonCategoryOptions(supervisorActionReasonCategoryInput, supervisorActionReasonCategoryInput.value || "");
    }
    if (supervisorActionReasonDetailInput) {
      setActionReasonDetailOptions(
        supervisorActionReasonDetailInput,
        line,
        String(supervisorActionReasonCategoryInput?.value || ""),
        supervisorActionReasonDetailInput.value || ""
      );
    }
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

  const renderLineGroupManagerList = () => {
    if (!lineGroupManagerList) return;
    const lineGroups = normalizeLineGroups(appState.lineGroups);
    const lines = sortLinesByDisplayOrder(Object.values(appState.lines || {}));
    const lineCountByGroup = {};
    lines.forEach((line) => {
      const groupId = String(line?.groupId || "").trim();
      if (!groupId) return;
      lineCountByGroup[groupId] = (lineCountByGroup[groupId] || 0) + 1;
    });

    const groupRows = lineGroups
      .map((group) => {
        const lineCount = lineCountByGroup[group.id] || 0;
        return `
          <div class="line-group-row" data-line-group-id="${group.id}">
            <input
              data-line-group-name-input="${group.id}"
              value="${htmlEscape(group.name)}"
              placeholder="Group name"
              aria-label="Line group name"
            />
            <span class="line-group-count">${lineCount} line${lineCount === 1 ? "" : "s"}</span>
            <button type="button" class="ghost-btn" data-line-group-rename="${group.id}">Save</button>
            <button type="button" class="danger" data-line-group-delete="${group.id}">Delete</button>
          </div>
        `;
      })
      .join("");

    const groupOptions = [
      `<option value="">Ungrouped</option>`,
      ...lineGroups.map((group) => `<option value="${group.id}">${htmlEscape(group.name)}</option>`)
    ].join("");
    const assignmentRows = lines
      .map((line, index) => {
        const selectedGroupId = String(line?.groupId || "").trim();
        return `
          <label class="line-group-line-row" for="line-group-assignment-${index}">
            <span>${htmlEscape(line.name)}</span>
            <select id="line-group-assignment-${index}" data-line-group-assignment="${line.id}">
              ${groupOptions}
            </select>
          </label>
        `;
      })
      .join("");

    lineGroupManagerList.innerHTML = `
      <section class="panel line-group-panel">
        <h3>Groups</h3>
        <div class="line-group-list">
          ${groupRows || `<p class="muted">No groups yet. Create one above.</p>`}
        </div>
      </section>
      <section class="panel line-group-panel">
        <h3>Line Assignments</h3>
        <div class="line-group-line-list">
          ${assignmentRows || `<p class="muted">No production lines available.</p>`}
        </div>
      </section>
    `;

    Array.from(lineGroupManagerList.querySelectorAll("[data-line-group-assignment]")).forEach((select) => {
      const lineId = String(select.getAttribute("data-line-group-assignment") || "");
      const line = appState.lines?.[lineId];
      if (!line) return;
      const selectedGroupId = String(line?.groupId || "").trim();
      select.value = selectedGroupId;
    });
  };

  const openManageLineGroupsModal = async () => {
    if (!manageLineGroupsModal) return;
    await refreshHostedState();
    if (lineGroupCreateForm) lineGroupCreateForm.reset();
    renderLineGroupManagerList();
    manageLineGroupsModal.classList.add("open");
    manageLineGroupsModal.setAttribute("aria-hidden", "false");
    if (newLineGroupNameInput) newLineGroupNameInput.focus();
  };

  const closeManageLineGroupsModal = () => {
    if (!manageLineGroupsModal) return;
    manageLineGroupsModal.classList.remove("open");
    manageLineGroupsModal.setAttribute("aria-hidden", "true");
  };

  const openManageProductCatalogModal = () => {
    if (!manageProductCatalogModal) return;
    if (productCatalogTable) {
      ensureProductCatalogActionColumn(productCatalogTable);
      refreshProductRowAssignmentSummaries(productCatalogTable);
      syncRunProductInputsFromCatalog();
    }
    manageProductCatalogModal.classList.add("open");
    manageProductCatalogModal.setAttribute("aria-hidden", "false");
  };

  const closeManageProductCatalogModal = () => {
    if (!manageProductCatalogModal) return;
    manageProductCatalogModal.classList.remove("open");
    manageProductCatalogModal.setAttribute("aria-hidden", "true");
    closeProductLineAssignModal();
  };

  const openAddManagerModal = () => {
    if (!addManagerModal || !addManagerForm) return;
    addManagerForm.reset();
    addManagerModal.classList.add("open");
    addManagerModal.setAttribute("aria-hidden", "false");
    newManagerNameInput?.focus();
  };

  const closeAddManagerModal = () => {
    if (!addManagerModal) return;
    addManagerModal.classList.remove("open");
    addManagerModal.setAttribute("aria-hidden", "true");
  };

  const setConnectDataSourceTestFeedback = (message = "", tone = "") => {
    if (!connectDataSourceTestResult) return;
    connectDataSourceTestResult.textContent = String(message || "").trim();
    connectDataSourceTestResult.classList.toggle("is-success", tone === "success");
    connectDataSourceTestResult.classList.toggle("is-error", tone === "error");
    if (!message) connectDataSourceTestResult.classList.remove("is-success", "is-error");
  };

  const connectDataSourceFormPayload = () => {
    const sourceName = String(connectDataSourceNameInput?.value || "").trim();
    const sourceKeyInput = String(connectDataSourceKeyInput?.value || "").trim();
    const sourceKey = sourceKeyInput || dataSourceKeyFromName(sourceName, "source");
    const provider = normalizeDataSourceProvider(connectDataSourceProviderInput?.value || "api");
    const connectionMode = connectDataSourceModeInput?.value === "sql" ? "sql" : "api";
    const machineNo = String(connectDataSourceMachineNoInput?.value || "").trim();
    const deviceName = String(connectDataSourceDeviceNameInput?.value || "").trim();
    const deviceId = String(connectDataSourceDeviceIdInput?.value || "").trim();
    const scaleNumber = String(connectDataSourceScaleNumberInput?.value || "").trim();
    const apiBaseUrl = String(connectDataSourceApiBaseUrlInput?.value || "").trim();
    const apiKey = String(connectDataSourceApiKeyInput?.value || "").trim();
    const apiSecret = String(connectDataSourceApiSecretInput?.value || "").trim();
    const sqlHost = String(connectDataSourceSqlHostInput?.value || "").trim();
    const sqlPortRaw = Number(connectDataSourceSqlPortInput?.value);
    const sqlPort = Number.isFinite(sqlPortRaw) && sqlPortRaw > 0 ? Math.floor(sqlPortRaw) : null;
    const sqlDatabase = String(connectDataSourceSqlDatabaseInput?.value || "").trim();
    const sqlUsername = String(connectDataSourceSqlUsernameInput?.value || "").trim();
    const sqlPassword = String(connectDataSourceSqlPasswordInput?.value || "").trim();
    return {
      sourceName,
      sourceKey,
      provider,
      connectionMode,
      machineNo,
      deviceName,
      deviceId,
      scaleNumber,
      apiBaseUrl,
      apiKey,
      apiSecret,
      sqlHost,
      sqlPort,
      sqlDatabase,
      sqlUsername,
      sqlPassword
    };
  };

  const validateConnectDataSourcePayload = (
    payload,
    {
      requireName = true,
      requireApiBaseUrl = false
    } = {}
  ) => {
    if (requireName && !payload.sourceName) {
      return {
        ok: false,
        message: "Source name is required.",
        focusNode: connectDataSourceNameInput
      };
    }
    if (payload.connectionMode === "api") {
      if (requireApiBaseUrl && !payload.apiBaseUrl) {
        return {
          ok: false,
          message: "API Base URL is required to test API connections.",
          focusNode: connectDataSourceApiBaseUrlInput
        };
      }
      if (!payload.apiKey) {
        return {
          ok: false,
          message: "API Key is required for API connection mode.",
          focusNode: connectDataSourceApiKeyInput
        };
      }
      return { ok: true };
    }

    if (!payload.sqlHost || !payload.sqlDatabase || !payload.sqlUsername || !payload.sqlPassword) {
      return {
        ok: false,
        message: "SQL connection mode requires host, database, username and password.",
        focusNode: connectDataSourceSqlHostInput
      };
    }
    return { ok: true };
  };

  const syncConnectDataSourceModeState = () => {
    const mode = connectDataSourceModeInput?.value === "sql" ? "sql" : "api";
    if (connectDataSourceApiFields) connectDataSourceApiFields.classList.toggle("hidden", mode !== "api");
    if (connectDataSourceSqlFields) connectDataSourceSqlFields.classList.toggle("hidden", mode !== "sql");
    if (connectDataSourceApiKeyInput) connectDataSourceApiKeyInput.required = mode === "api";
    if (connectDataSourceSqlHostInput) connectDataSourceSqlHostInput.required = mode === "sql";
    if (connectDataSourceSqlDatabaseInput) connectDataSourceSqlDatabaseInput.required = mode === "sql";
    if (connectDataSourceSqlUsernameInput) connectDataSourceSqlUsernameInput.required = mode === "sql";
    if (connectDataSourceSqlPasswordInput) connectDataSourceSqlPasswordInput.required = mode === "sql";
    if (connectDataSourceHelp) {
      connectDataSourceHelp.textContent =
        mode === "sql"
          ? "Enter SQL host, database, username and password to configure this source."
          : "Enter API key details to configure this source.";
    }
    setConnectDataSourceTestFeedback();
  };

  const resetConnectDataSourceForm = () => {
    if (!connectDataSourceForm) return;
    connectDataSourceForm.reset();
    if (connectDataSourceProviderInput) connectDataSourceProviderInput.value = "api";
    if (connectDataSourceModeInput) connectDataSourceModeInput.value = "api";
    if (connectDataSourceSqlPortInput) connectDataSourceSqlPortInput.value = "5432";
    if (connectDataSourceTestBtn) {
      connectDataSourceTestBtn.disabled = false;
      connectDataSourceTestBtn.textContent = "Test Connection";
    }
    setConnectDataSourceTestFeedback();
    syncConnectDataSourceModeState();
  };

  const openConnectDataSourceModal = () => {
    if (!connectDataSourceModal) return;
    resetConnectDataSourceForm();
    connectDataSourceModal.classList.add("open");
    connectDataSourceModal.setAttribute("aria-hidden", "false");
    connectDataSourceNameInput?.focus();
  };

  const closeConnectDataSourceModal = () => {
    if (!connectDataSourceModal) return;
    connectDataSourceModal.classList.remove("open");
    connectDataSourceModal.setAttribute("aria-hidden", "true");
  };

  const saveLineGroupAssignment = async (lineId, nextGroupId) => {
    const line = appState.lines?.[lineId];
    if (!line) return;
    const validGroupIds = new Set(normalizeLineGroups(appState.lineGroups).map((group) => group.id));
    const safeGroupId = nextGroupId && validGroupIds.has(nextGroupId) ? nextGroupId : "";
    const session = await ensureManagerBackendSession();
    const backendLineId = UUID_RE.test(String(lineId)) ? lineId : await ensureBackendLineId(lineId, session);
    if (!backendLineId) throw new Error("Line is not synced to server.");
    const response = await apiRequest(`/api/lines/${backendLineId}`, {
      method: "PATCH",
      token: session.backendToken,
      body: { groupId: safeGroupId || null }
    });
    line.groupId = safeGroupId;
    const nextOrder = Number(response?.line?.displayOrder);
    if (Number.isFinite(nextOrder)) line.displayOrder = Math.max(0, Math.floor(nextOrder));
    saveState();
    renderHome();
    renderLineGroupManagerList();
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

  const deleteLineById = async (lineId) => {
    const line = appState.lines?.[lineId];
    if (!line) return;
    const entered = window.prompt(`Enter delete key for "${line.name}" (or admin password):`) || "";
    if (entered !== "admin" && entered !== line.secretKey) {
      alert("Invalid key/password. Line was not deleted.");
      return;
    }
    try {
      const session = await ensureManagerBackendSession();
      const backendLineId = UUID_RE.test(String(lineId)) ? lineId : await ensureBackendLineId(lineId, session);
      if (!backendLineId) throw new Error("Line is not synced to server.");
      await apiRequest(`/api/lines/${backendLineId}`, {
        method: "DELETE",
        token: session.backendToken
      });
      addAudit(line, "DELETE_LINE", "Line deleted");
      delete appState.lines[lineId];
      appState.supervisors = (appState.supervisors || []).map((sup) => ({
        ...sup,
        assignedLineIds: (sup.assignedLineIds || []).filter((assignedId) => assignedId !== lineId),
        assignedLineShifts: Object.fromEntries(
          Object.entries(sup.assignedLineShifts || {}).filter(([assignedId]) => assignedId !== lineId)
        )
      }));
      if (appState.supervisorSession) {
        appState.supervisorSession.assignedLineIds = (appState.supervisorSession.assignedLineIds || []).filter(
          (assignedId) => assignedId !== lineId
        );
        appState.supervisorSession.assignedLineShifts = Object.fromEntries(
          Object.entries(appState.supervisorSession.assignedLineShifts || {}).filter(([assignedId]) => assignedId !== lineId)
        );
      }
      appState.activeLineId = Object.keys(appState.lines)[0] || "";
      state = appState.lines[appState.activeLineId] || null;
      appState.activeView = "home";
      closeEditLineModal();
      saveState();
      await refreshHostedState();
      renderAll();
    } catch (error) {
      console.warn("Line delete sync failed:", error);
      alert(`Could not delete line.\n${error?.message || "Please try again."}`);
    }
  };

  goHomeBtn.addEventListener("click", () => {
    appState.activeView = "home";
    saveState();
    renderAll();
  });

  const switchAppMode = (mode, { goHome = false } = {}) => {
    const nextMode = mode === "supervisor" ? "supervisor" : "manager";
    if (APP_VARIANT === "manager" && nextMode !== "manager") return;
    if (APP_VARIANT === "supervisor" && nextMode !== "supervisor") return;
    appState.appMode = nextMode;
    if (goHome) appState.activeView = "home";
    saveState();
    renderAll();
  };

  const setManagerHomeTab = (tab) => {
    appState.managerHomeTab = tab === "settings" ? "settings" : "dashboard";
    saveState();
    renderHome();
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

  if (managerSettingsTabBtn) {
    managerSettingsTabBtn.addEventListener("click", () => {
      if (appState.appMode !== "manager" || !managerBackendSession.backendToken) return;
      setManagerHomeTab(appState.managerHomeTab === "settings" ? "dashboard" : "settings");
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
  window.addEventListener("resize", scheduleLineShiftTrackerResizeRender);

  dashboardDate.addEventListener("change", () => {
    appState.dashboardDate = normalizeWeekdayIsoDate(dashboardDate.value || appState.dashboardDate || todayISO(), { direction: -1 });
    dashboardDate.value = appState.dashboardDate;
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
      managerBackendSession.name = String(loginPayload?.user?.name || loginPayload?.user?.username || username).trim();
      managerBackendSession.username = String(loginPayload?.user?.username || username).trim().toLowerCase();
      persistManagerBackendSession();
      appState.appMode = "manager";
      appState.managerHomeTab = "dashboard";
      appState.activeView = "home";
      clearManagerActionTicketEdit();
      saveState();
      await refreshHostedState(managerBackendSession);
      renderAll();
    } catch (error) {
      const message = String(error?.message || "");
      if (message.toLowerCase().includes("invalid")) {
        alert("Invalid manager credentials.");
      } else {
        alert(`Could not connect to login service at ${API_BASE_URL}.\n${message || "Please try again."}`);
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
        name: String(loginPayload?.user?.name || loginPayload?.user?.username || username).trim(),
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
        alert(`Could not connect to login service at ${API_BASE_URL}.\n${message || "Please try again."}`);
      }
    }
  });

  supervisorLogoutBtn.addEventListener("click", () => {
    if (appState.appMode === "supervisor") {
      appState.supervisorSession = null;
    } else {
      clearManagerBackendSession();
      appState.managerHomeTab = "dashboard";
      appState.activeView = "home";
      clearManagerActionTicketEdit();
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
        if (superShiftLogIdInput) superShiftLogIdInput.value = "";
        supervisorShiftTileEditId = "";
      }
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
  ["superShiftDate", "superRunDate", "superDownDate"].forEach((id) => {
    const dateInput = document.getElementById(id);
    if (!dateInput) return;
    dateInput.addEventListener("change", () => {
      if (!dateInput.value) return;
      dateInput.value = normalizeWeekdayIsoDate(dateInput.value, { direction: -1 });
    });
  });

  document.querySelectorAll("[data-supervisor-tab]").forEach((btn) => {
    btn.addEventListener("click", () => {
      appState.supervisorTab = btn.dataset.supervisorTab || "superShift";
      clearSupervisorEditingSelection({ resetForms: true });
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
    clearSupervisorEditingSelection({ resetForms: true });
    syncRunProductInputsFromCatalog();
    refreshSupervisorDowntimeDetailOptions(selectedSupervisorLine());
    saveState();
    renderHome();
  });

  if (supervisorActionLineInput) {
    supervisorActionLineInput.addEventListener("change", () => {
      refreshSupervisorActionRelationOptions();
    });
  }

  if (supervisorActionReasonCategoryInput) {
    supervisorActionReasonCategoryInput.addEventListener("change", () => {
      refreshSupervisorActionRelationOptions();
    });
  }

  svDateInputs.forEach((svDateInput) => {
    svDateInput.addEventListener("change", () => {
      appState.supervisorSelectedDate = normalizeWeekdayIsoDate(
        svDateInput.value || appState.supervisorSelectedDate || todayISO(),
        { direction: -1 }
      );
      svDateInputs.forEach((input) => {
        input.value = appState.supervisorSelectedDate;
      });
      clearSupervisorEditingSelection({ resetForms: true });
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
      clearSupervisorEditingSelection({ resetForms: true });
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
      appState.supervisorSelectedDate = shiftWeekdayIsoDate(appState.supervisorSelectedDate || todayISO(), -1);
      clearSupervisorEditingSelection({ resetForms: true });
      saveState();
      renderHome();
    });
  });

  svNextBtns.forEach((svNextBtn) => {
    svNextBtn.addEventListener("click", () => {
      appState.supervisorSelectedDate = shiftWeekdayIsoDate(appState.supervisorSelectedDate || todayISO(), 1);
      clearSupervisorEditingSelection({ resetForms: true });
      saveState();
      renderHome();
    });
  });

  document.querySelectorAll("[data-super-main-tab]").forEach((btn) => {
    btn.addEventListener("click", () => {
      appState.supervisorMainTab = btn.dataset.superMainTab || "supervisorDay";
      clearSupervisorEditingSelection({ resetForms: true });
      saveState();
      renderHome();
    });
  });

  if (managerActionList) {
    managerActionList.addEventListener("change", (event) => {
      const categorySelect = event.target.closest("[data-manager-action-reason-category]");
      if (!categorySelect) return;
      const row = categorySelect.closest("[data-manager-action-row]") || categorySelect.closest("tr");
      const reasonDetailSelect = row?.querySelector("[data-manager-action-reason-detail]");
      if (!reasonDetailSelect) return;
      const lineId = String(row?.getAttribute("data-manager-line-id") || "").trim();
      const line = lineId && appState.lines[lineId] ? appState.lines[lineId] : null;
      setActionReasonDetailOptions(reasonDetailSelect, line, String(categorySelect.value || ""), reasonDetailSelect.value || "");
    });

    managerActionList.addEventListener("click", async (event) => {
      const editBtn = event.target.closest("[data-manager-action-edit]");
      if (editBtn) {
        if (appState.appMode !== "manager" || !managerBackendSession.backendToken) return;
        const actionId = String(editBtn.getAttribute("data-manager-action-edit") || "").trim();
        if (!actionId) return;
        setManagerActionTicketEdit(actionId);
        renderHome();
        return;
      }

      const deleteBtn = event.target.closest("[data-manager-action-delete]");
      if (deleteBtn) {
        if (appState.appMode !== "manager" || !managerBackendSession.backendToken) return;
        const actionId = String(deleteBtn.getAttribute("data-manager-action-delete") || "").trim();
        if (!actionId) return;
        if (!UUID_RE.test(actionId)) {
          alert("Action is not synced to the server.");
          return;
        }
        const actions = ensureSupervisorActionsState();
        const actionIndex = actions.findIndex((item) => String(item?.id || "") === actionId);
        if (actionIndex < 0) {
          alert("Action could not be found.");
          return;
        }
        const action = actions[actionIndex];
        const actionTitle = String(action?.title || "this action").trim() || "this action";
        if (!window.confirm(`Delete "${actionTitle}"?`)) return;
        try {
          await deleteManagerSupervisorAction(actionId);
          actions.splice(actionIndex, 1);
          if (isManagerActionTicketEditRow(actionId)) clearManagerActionTicketEdit();
          appState.supervisorActions = normalizeSupervisorActions(actions);
          saveState();
          renderHome();
        } catch (error) {
          alert(`Could not delete action.\n${error?.message || "Please try again."}`);
        }
        return;
      }

      const saveBtn = event.target.closest("[data-manager-action-reassign]");
      if (!saveBtn) return;
      if (appState.appMode !== "manager" || !managerBackendSession.backendToken) return;
      const actionId = String(saveBtn.getAttribute("data-manager-action-reassign") || "").trim();
      if (!actionId) return;
      if (!UUID_RE.test(actionId)) {
        alert("Action is not synced to the server.");
        return;
      }
      const row = saveBtn.closest("[data-manager-action-row]") || saveBtn.closest("tr");
      const supervisorSelect = row?.querySelector("[data-manager-action-supervisor]");
      if (!supervisorSelect) return;
      const prioritySelect = row?.querySelector("[data-manager-action-priority]");
      const statusSelect = row?.querySelector("[data-manager-action-status]");
      const dueDateInput = row?.querySelector("[data-manager-action-due-date]");
      const relatedEquipmentSelect = row?.querySelector("[data-manager-action-related-equipment]");
      const relatedReasonCategorySelect = row?.querySelector("[data-manager-action-reason-category]");
      const relatedReasonDetailSelect = row?.querySelector("[data-manager-action-reason-detail]");
      const nextUsername = String(supervisorSelect.value || "").trim().toLowerCase();
      if (!nextUsername) {
        alert("Assigned owner is required.");
        return;
      }
      const actions = ensureSupervisorActionsState();
      const action = actions.find((item) => String(item?.id || "") === actionId);
      if (!action) {
        alert("Action could not be found.");
        return;
      }
      const previousUsername = String(action.supervisorUsername || "").trim().toLowerCase();
      const previousPriority = normalizeActionPriority(action.priority);
      const previousStatus = normalizeActionStatus(action.status);
      const previousDueDateRaw = String(action.dueDate || "").trim();
      const previousDueDate = /^\d{4}-\d{2}-\d{2}$/.test(previousDueDateRaw) ? previousDueDateRaw : "";
      const previousRelatedEquipmentId = String(action.relatedEquipmentId || "").trim();
      const previousRelatedReasonCategory = normalizeActionReasonCategory(action.relatedReasonCategory);
      const previousRelatedReasonDetail = previousRelatedReasonCategory ? String(action.relatedReasonDetail || "").trim() : "";
      const nextPriority = normalizeActionPriority(prioritySelect?.value);
      const nextStatus = normalizeActionStatus(statusSelect?.value);
      const nextDueDateRaw = String(dueDateInput?.value || "").trim();
      const nextDueDate = /^\d{4}-\d{2}-\d{2}$/.test(nextDueDateRaw) ? nextDueDateRaw : "";
      const line = action.lineId && appState.lines[action.lineId] ? appState.lines[action.lineId] : null;
      const nextRelatedEquipmentId = String(relatedEquipmentSelect?.value || "").trim();
      const validEquipmentIds = new Set((line?.stages || []).map((stage) => String(stage?.id || "").trim()).filter(Boolean));
      if (nextRelatedEquipmentId && !validEquipmentIds.has(nextRelatedEquipmentId)) {
        alert("Select valid related equipment.");
        return;
      }
      const nextRelatedReasonCategory = normalizeActionReasonCategory(relatedReasonCategorySelect?.value);
      const nextRelatedReasonDetailRaw = String(relatedReasonDetailSelect?.value || "").trim();
      let nextRelatedReasonDetail = "";
      if (nextRelatedReasonCategory) {
        const validReasonDetails = new Set(
          downtimeDetailOptions(line, nextRelatedReasonCategory).map((option) => String(option?.value || "").trim()).filter(Boolean)
        );
        if (nextRelatedReasonDetailRaw && !validReasonDetails.has(nextRelatedReasonDetailRaw)) {
          alert("Select valid downtime reason.");
          return;
        }
        nextRelatedReasonDetail = nextRelatedReasonDetailRaw;
      }
      if (
        previousUsername === nextUsername &&
        previousPriority === nextPriority &&
        previousStatus === nextStatus &&
        previousDueDate === nextDueDate &&
        previousRelatedEquipmentId === nextRelatedEquipmentId &&
        previousRelatedReasonCategory === nextRelatedReasonCategory &&
        previousRelatedReasonDetail === nextRelatedReasonDetail
      ) {
        clearManagerActionTicketEdit();
        renderHome();
        return;
      }
      const nextSupervisorName = actionAssignmentLabel(nextUsername, action.supervisorName);
      try {
        const savedAction = await patchManagerSupervisorAction(actionId, {
          supervisorUsername: nextUsername,
          supervisorName: nextSupervisorName,
          lineId: action.lineId,
          priority: nextPriority,
          status: nextStatus,
          dueDate: nextDueDate,
          relatedEquipmentId: nextRelatedEquipmentId,
          relatedReasonCategory: nextRelatedReasonCategory,
          relatedReasonDetail: nextRelatedReasonDetail
        });
        const updatedAction = normalizeSupervisorAction(savedAction);
        if (!updatedAction) throw new Error("Action update returned invalid data.");
        appState.supervisorActions = normalizeSupervisorActions(
          actions.map((item) => (String(item?.id || "") === actionId ? updatedAction : item))
        );
        clearManagerActionTicketEdit();
        saveState();
        renderHome();
      } catch (error) {
        alert(`Could not update action.\n${error?.message || "Please try again."}`);
      }
    });
  }

  manageSupervisorsBtn.addEventListener("click", openManageSupervisorsModal);
  closeManageSupervisorsModalBtn.addEventListener("click", closeManageSupervisorsModal);
  manageSupervisorsModal.addEventListener("click", (event) => {
    if (event.target === manageSupervisorsModal) closeManageSupervisorsModal();
  });

  if (manageLineGroupsBtn) {
    manageLineGroupsBtn.addEventListener("click", openManageLineGroupsModal);
  }
  if (closeManageLineGroupsModalBtn) {
    closeManageLineGroupsModalBtn.addEventListener("click", closeManageLineGroupsModal);
  }
  if (manageLineGroupsModal) {
    manageLineGroupsModal.addEventListener("click", (event) => {
      if (event.target === manageLineGroupsModal) closeManageLineGroupsModal();
    });
  }
  if (manageProductCatalogBtn) {
    manageProductCatalogBtn.addEventListener("click", openManageProductCatalogModal);
  }
  if (closeManageProductCatalogModalBtn) {
    closeManageProductCatalogModalBtn.addEventListener("click", closeManageProductCatalogModal);
  }
  if (manageProductCatalogModal) {
    manageProductCatalogModal.addEventListener("click", (event) => {
      if (event.target === manageProductCatalogModal) closeManageProductCatalogModal();
    });
  }
  if (connectDataSourceBtn) {
    connectDataSourceBtn.addEventListener("click", openConnectDataSourceModal);
  }
  if (closeConnectDataSourceModalBtn) {
    closeConnectDataSourceModalBtn.addEventListener("click", closeConnectDataSourceModal);
  }
  if (connectDataSourceModal) {
    connectDataSourceModal.addEventListener("click", (event) => {
      if (event.target === connectDataSourceModal) closeConnectDataSourceModal();
    });
  }
  if (connectDataSourceModeInput) {
    connectDataSourceModeInput.addEventListener("change", syncConnectDataSourceModeState);
  }
  if (connectDataSourceProviderInput && connectDataSourceModeInput) {
    connectDataSourceProviderInput.addEventListener("change", () => {
      const provider = normalizeDataSourceProvider(connectDataSourceProviderInput.value);
      if (provider === "sql") connectDataSourceModeInput.value = "sql";
      if (provider === "api") connectDataSourceModeInput.value = "api";
      syncConnectDataSourceModeState();
    });
  }
  if (connectDataSourceNameInput && connectDataSourceKeyInput) {
    connectDataSourceNameInput.addEventListener("blur", () => {
      if (String(connectDataSourceKeyInput.value || "").trim()) return;
      const name = String(connectDataSourceNameInput.value || "").trim();
      if (!name) return;
      connectDataSourceKeyInput.value = dataSourceKeyFromName(name, "source");
    });
  }

  const runModalDataSourceConnectionTest = async () => {
    const payload = connectDataSourceFormPayload();
    const validation = validateConnectDataSourcePayload(payload, {
      requireName: false,
      requireApiBaseUrl: payload.connectionMode === "api"
    });
    if (!validation.ok) {
      setConnectDataSourceTestFeedback(validation.message || "Connection test validation failed.", "error");
      validation.focusNode?.focus();
      return;
    }
    if (connectDataSourceTestBtn) {
      connectDataSourceTestBtn.disabled = true;
      connectDataSourceTestBtn.textContent = "Testing...";
    }
    setConnectDataSourceTestFeedback("Testing connection...");
    try {
      const response = await testDataSourceConnectionOnBackend(payload);
      const test = response?.test || null;
      if (test?.ok) {
        setConnectDataSourceTestFeedback(`Connection successful. ${String(test.message || "").trim()}`.trim(), "success");
      } else {
        setConnectDataSourceTestFeedback(
          `Connection failed. ${String(test?.message || "Please review your settings and retry.").trim()}`.trim(),
          "error"
        );
      }
    } catch (error) {
      setConnectDataSourceTestFeedback(`Connection failed. ${error?.message || "Please try again."}`, "error");
    } finally {
      if (connectDataSourceTestBtn) {
        connectDataSourceTestBtn.disabled = false;
        connectDataSourceTestBtn.textContent = "Test Connection";
      }
    }
  };

  const runSavedDataSourceConnectionTest = async (dataSourceId) => {
    const safeDataSourceId = String(dataSourceId || "").trim();
    if (!UUID_RE.test(safeDataSourceId)) {
      alert("Invalid data source id.");
      return;
    }
    const source = dataSourceById(safeDataSourceId);
    const sourceLabel = source?.sourceName || "Data source";
    try {
      const response = await testSavedDataSourceConnectionOnBackend(safeDataSourceId);
      const test = response?.test || null;
      if (test) rememberDataSourceConnectionTest(safeDataSourceId, test);
      renderManagerDataSourcesList(dataSourcesList);
      renderIncomingDataSourcesStatus(incomingDataSourcesStatus);
      if (test?.ok) {
        alert(`${sourceLabel} connection passed.\n${String(test.message || "Connection successful.")}`);
      } else {
        alert(`${sourceLabel} connection failed.\n${String(test?.message || "Please review the data source credentials.")}`);
      }
    } catch (error) {
      console.warn("Data source connection test failed:", error);
      rememberDataSourceConnectionTest(safeDataSourceId, {
        ok: false,
        message: String(error?.message || "Connection test failed.")
      });
      renderManagerDataSourcesList(dataSourcesList);
      renderIncomingDataSourcesStatus(incomingDataSourcesStatus);
      alert(`Could not test data source connection.\n${error?.message || "Please try again."}`);
    }
  };

  if (connectDataSourceTestBtn) {
    connectDataSourceTestBtn.addEventListener("click", runModalDataSourceConnectionTest);
  }

  if (connectDataSourceForm) {
    connectDataSourceForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const payload = connectDataSourceFormPayload();
      const validation = validateConnectDataSourcePayload(payload, {
        requireName: true,
        requireApiBaseUrl: false
      });
      if (!validation.ok) {
        alert(validation.message || "Data source details are incomplete.");
        validation.focusNode?.focus();
        return;
      }
      try {
        await createDataSourceOnBackend(payload);
        closeConnectDataSourceModal();
        await refreshHostedState();
        renderHome();
      } catch (error) {
        console.warn("Data source create failed:", error);
        alert(`Could not connect data source.\n${error?.message || "Please try again."}`);
      }
    });
    syncConnectDataSourceModeState();
  }
  if (dataSourcesList) {
    dataSourcesList.addEventListener("click", async (event) => {
      const testBtn = event.target.closest("[data-data-source-test]");
      if (!testBtn) return;
      const dataSourceId = String(testBtn.getAttribute("data-data-source-test") || "").trim();
      if (!dataSourceId) return;
      await runSavedDataSourceConnectionTest(dataSourceId);
    });
  }
  if (closeProductLineAssignModalBtn) {
    closeProductLineAssignModalBtn.addEventListener("click", closeProductLineAssignModal);
  }
  if (productLineAssignModal) {
    productLineAssignModal.addEventListener("click", (event) => {
      if (event.target === productLineAssignModal) closeProductLineAssignModal();
    });
  }
  if (productLineAssignSelectAllBtn) {
    productLineAssignSelectAllBtn.addEventListener("click", () => {
      productLineAssignList?.querySelectorAll('input[type="checkbox"]')?.forEach((input) => {
        input.checked = true;
      });
    });
  }
  if (productLineAssignClearAllBtn) {
    productLineAssignClearAllBtn.addEventListener("click", () => {
      productLineAssignList?.querySelectorAll('input[type="checkbox"]')?.forEach((input) => {
        input.checked = false;
      });
    });
  }
  if (productLineAssignSaveBtn) {
    productLineAssignSaveBtn.addEventListener("click", async () => {
      await saveProductLineAssignments();
    });
  }
  if (lineGroupCreateForm) {
    lineGroupCreateForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const name = String(newLineGroupNameInput?.value || "").trim();
      if (!name) {
        alert("Group name is required.");
        return;
      }
      try {
        const session = await ensureManagerBackendSession();
        await apiRequest("/api/line-groups", {
          method: "POST",
          token: session.backendToken,
          body: { name }
        });
        await refreshHostedState();
        renderLineGroupManagerList();
        if (lineGroupCreateForm) lineGroupCreateForm.reset();
      } catch (error) {
        console.warn("Line group create failed:", error);
        alert(`Could not create line group.\n${error?.message || "Please try again."}`);
      }
    });
  }
  if (lineGroupManagerList) {
    lineGroupManagerList.addEventListener("change", async (event) => {
      const assignmentSelect = event.target.closest("[data-line-group-assignment]");
      if (!assignmentSelect) return;
      const lineId = String(assignmentSelect.getAttribute("data-line-group-assignment") || "");
      const groupId = String(assignmentSelect.value || "").trim();
      try {
        await saveLineGroupAssignment(lineId, groupId);
      } catch (error) {
        console.warn("Line group assignment failed:", error);
        alert(`Could not update line group assignment.\n${error?.message || "Please try again."}`);
        renderLineGroupManagerList();
      }
    });
    lineGroupManagerList.addEventListener("click", async (event) => {
      const renameBtn = event.target.closest("[data-line-group-rename]");
      if (renameBtn) {
        const groupId = String(renameBtn.getAttribute("data-line-group-rename") || "");
        if (!groupId) return;
        const input = lineGroupManagerList.querySelector(`[data-line-group-name-input="${groupId}"]`);
        const nextName = String(input?.value || "").trim();
        if (!nextName) {
          alert("Group name is required.");
          return;
        }
        try {
          const session = await ensureManagerBackendSession();
          await apiRequest(`/api/line-groups/${groupId}`, {
            method: "PATCH",
            token: session.backendToken,
            body: { name: nextName }
          });
          await refreshHostedState();
          renderLineGroupManagerList();
        } catch (error) {
          console.warn("Line group rename failed:", error);
          alert(`Could not rename line group.\n${error?.message || "Please try again."}`);
        }
        return;
      }

      const deleteBtn = event.target.closest("[data-line-group-delete]");
      if (!deleteBtn) return;
      const groupId = String(deleteBtn.getAttribute("data-line-group-delete") || "");
      if (!groupId) return;
      const group = normalizeLineGroups(appState.lineGroups).find((item) => item.id === groupId);
      const groupName = group?.name || "this group";
      if (!window.confirm(`Delete "${groupName}"? Lines in this group will become ungrouped.`)) return;
      try {
        const session = await ensureManagerBackendSession();
        await apiRequest(`/api/line-groups/${groupId}`, {
          method: "DELETE",
          token: session.backendToken
        });
        await refreshHostedState();
        renderLineGroupManagerList();
      } catch (error) {
        console.warn("Line group delete failed:", error);
        alert(`Could not delete line group.\n${error?.message || "Please try again."}`);
      }
    });
  }

  if (productCatalogTable) {
    hydrateProductCatalogTableFromState(productCatalogTable);
    syncProductCatalogStateFromTable(productCatalogTable);
    const productCatalogWrap = productCatalogTable.closest(".products-table-wrap");
    if (manageProductCatalogContent && productCatalogWrap && !manageProductCatalogContent.contains(productCatalogWrap)) {
      manageProductCatalogContent.append(productCatalogWrap);
    }
    ensureProductCatalogActionColumn(productCatalogTable);
    refreshProductRowAssignmentSummaries(productCatalogTable);
    syncRunProductInputsFromCatalog();
    productCatalogTable.addEventListener("click", async (event) => {
      const editBtn = event.target.closest("[data-product-row-edit]");
      const linesBtn = event.target.closest("[data-product-row-lines]");
      if (linesBtn) {
        if (!managerBackendSession.backendToken || appState.appMode !== "manager") return;
        const rowNode = linesBtn.closest("tr");
        if (!rowNode) return;
        openProductLineAssignModal(rowNode);
        return;
      }
      if (!editBtn) return;
      if (!managerBackendSession.backendToken || appState.appMode !== "manager") return;
      const rowNode = editBtn.closest("tr");
      if (!rowNode) return;
      const isEditing = rowNode.getAttribute("data-product-editing") === "true";
      if (isEditing) {
        await finalizeProductRowInlineEdit(rowNode);
        return;
      }
      const openRows = Array.from(productCatalogTable.querySelectorAll('tbody tr[data-product-editing="true"]')).filter(
        (openRow) => openRow !== rowNode
      );
      for (const openRow of openRows) {
        const saved = await finalizeProductRowInlineEdit(openRow);
        if (!saved) return;
      }
      beginProductRowInlineEdit(rowNode);
    });
  }
  if (addProductBtn && productCatalogTable) {
    addProductBtn.disabled = false;
    addProductBtn.textContent = "Add Product";
    addProductBtn.addEventListener("click", async () => {
      if (!managerBackendSession.backendToken || appState.appMode !== "manager") return;
      const openRows = Array.from(productCatalogTable.querySelectorAll('tbody tr[data-product-editing="true"]'));
      for (const openRow of openRows) {
        const saved = await finalizeProductRowInlineEdit(openRow);
        if (!saved) return;
      }
      const rowNode = buildProductCatalogRow(Array(PRODUCT_CATALOG_EDITABLE_COLUMN_COUNT).fill(""), []);
      const tbody = productCatalogTable.tBodies?.[0] || productCatalogTable.createTBody();
      tbody.prepend(rowNode);
      ensureProductCatalogActionColumn(productCatalogTable);
      setProductRowEditingState(rowNode, false);
      setProductRowAssignmentSummary(rowNode);
      beginProductRowInlineEdit(rowNode);
      const firstInput = rowNode.querySelector('input[data-product-inline-input="2"]') || rowNode.querySelector('input[data-product-inline-input="1"]');
      firstInput?.focus();
    });
  }

  if (addManagerBtn) addManagerBtn.addEventListener("click", openAddManagerModal);
  if (closeAddManagerModalBtn) closeAddManagerModalBtn.addEventListener("click", closeAddManagerModal);
  if (addManagerModal) {
    addManagerModal.addEventListener("click", (event) => {
      if (event.target === addManagerModal) closeAddManagerModal();
    });
  }

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
    if (!name || !username) {
      alert("Name and username are required.");
      return;
    }
    try {
      const session = await ensureManagerBackendSession();
      await apiRequest(`/api/supervisors/${sup.id}`, {
        method: "PATCH",
        token: session.backendToken,
        body: {
          name,
          username
        }
      });
      const previousUsername = sup.username;
      sup.name = name;
      sup.username = username;
      if (appState.supervisorSession?.username === previousUsername) {
        appState.supervisorSession.name = name;
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

  if (addManagerForm) {
    addManagerForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const name = String(newManagerNameInput?.value || "").trim();
      const username = String(newManagerUsernameInput?.value || "").trim().toLowerCase();
      const password = String(newManagerPasswordInput?.value || "").trim();
      if (!name || !username || !password) {
        alert("Name, username and password are required.");
        return;
      }
      if (password.length < 6) {
        alert("Password must be at least 6 characters.");
        return;
      }
      try {
        const session = await ensureManagerBackendSession();
        await apiRequest("/api/managers", {
          method: "POST",
          token: session.backendToken,
          body: {
            name,
            username,
            password
          }
        });
        closeAddManagerModal();
      } catch (error) {
        console.warn("Manager create sync failed:", error);
        alert(`Could not create manager.\n${error?.message || "Please try again."}`);
      }
    });
  }

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
    return sup?.name || session?.name || session?.username || "supervisor";
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
    strictTimeValid(String(row?.productionStartTime || "").trim()) &&
    (
      String(row?.finishTime || "").trim() === ""
      || (
        strictTimeValid(String(row?.finishTime || "").trim())
        && String(row?.productionStartTime || "").trim() === String(row?.finishTime || "").trim()
      )
    );

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
  if (supervisorActionForm) {
    supervisorActionForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const session = appState.supervisorSession;
      if (!session?.username) {
        alert("Log in as a supervisor to lodge actions.");
        return;
      }
      if (!session.backendToken) {
        alert("Supervisor session is not connected to the server. Please log in again.");
        return;
      }
      const assignedLineIds = Array.isArray(session.assignedLineIds) ? session.assignedLineIds : [];
      const selectedLineId = String(supervisorActionLineInput?.value || "").trim();
      if (selectedLineId && !assignedLineIds.includes(selectedLineId)) {
        alert("Select one of your assigned lines.");
        return;
      }
      const selectedLine = selectedLineId && appState.lines[selectedLineId] ? appState.lines[selectedLineId] : null;
      const title = String(supervisorActionTitleInput?.value || "").trim();
      if (!title) {
        alert("Action title is required.");
        return;
      }
      const description = String(supervisorActionDescriptionInput?.value || "").trim();
      if (!description) {
        alert("Description is required.");
        return;
      }
      const dueDate = supervisorAutoDateValue(supervisorActionDueDateInput, todayISO());
      if (dueDate && !/^\d{4}-\d{2}-\d{2}$/.test(dueDate)) {
        alert("Due date is invalid.");
        return;
      }
      const relatedEquipmentRaw = String(supervisorActionEquipmentInput?.value || "").trim();
      const validEquipmentIds = new Set((selectedLine?.stages || []).map((stage) => String(stage?.id || "").trim()).filter(Boolean));
      if (relatedEquipmentRaw && !validEquipmentIds.has(relatedEquipmentRaw)) {
        alert("Select valid equipment for the selected line.");
        return;
      }
      const relatedReasonCategory = normalizeActionReasonCategory(supervisorActionReasonCategoryInput?.value);
      const relatedReasonDetailRaw = String(supervisorActionReasonDetailInput?.value || "").trim();
      let relatedReasonDetail = "";
      if (relatedReasonCategory) {
        const validReasonDetails = new Set(
          downtimeDetailOptions(selectedLine, relatedReasonCategory).map((option) => String(option?.value || "").trim()).filter(Boolean)
        );
        if (relatedReasonDetailRaw && !validReasonDetails.has(relatedReasonDetailRaw)) {
          alert("Select a valid downtime reason.");
          return;
        }
        relatedReasonDetail = relatedReasonDetailRaw;
      }
      let nextAction = null;
      try {
        const savedAction = await syncSupervisorAction(session, {
          lineId: selectedLineId,
          title,
          description,
          priority: normalizeActionPriority(supervisorActionPriorityInput?.value),
          status: normalizeActionStatus(supervisorActionStatusInput?.value),
          dueDate,
          relatedEquipmentId: relatedEquipmentRaw,
          relatedReasonCategory,
          relatedReasonDetail
        });
        const normalizedSaved = normalizeSupervisorAction(savedAction);
        if (!normalizedSaved) throw new Error("Server returned invalid action data.");
        nextAction = normalizedSaved;
      } catch (error) {
        alert(`Could not lodge action.\n${error?.message || "Please try again."}`);
        return;
      }
      const actions = ensureSupervisorActionsState();
      actions.unshift(nextAction);
      appState.supervisorActions = normalizeSupervisorActions(actions);
      supervisorActionForm.reset();
      if (supervisorActionLineInput && selectedLineId && assignedLineIds.includes(selectedLineId)) {
        supervisorActionLineInput.value = selectedLineId;
      }
      if (supervisorActionPriorityInput) supervisorActionPriorityInput.value = "Medium";
      if (supervisorActionStatusInput) supervisorActionStatusInput.value = "Open";
      setSupervisorAutoDateValue(supervisorActionDueDateInput, todayISO(), { fallbackIso: todayISO() });
      if (supervisorActionEquipmentInput) supervisorActionEquipmentInput.value = "";
      if (supervisorActionReasonCategoryInput) supervisorActionReasonCategoryInput.value = "";
      if (supervisorActionReasonDetailInput) supervisorActionReasonDetailInput.value = "";
      refreshSupervisorActionRelationOptions();
      saveState();
      renderHome();
    });
  }
  if (superRunCrewingPatternBtn) {
    superRunCrewingPatternBtn.addEventListener("click", () => {
      const line = selectedSupervisorLine();
      if (!line) {
        alert("No assigned line selected.");
        return;
      }
      const runDate = supervisorAutoDateValue("superRunDate", supervisorAutoEntryDate());
      const runStart = String(document.getElementById("superRunProdStart")?.value || nowTimeHHMM());
      const runFinish = String(document.getElementById("superRunFinish")?.value || runStart);
      const shiftForPattern = preferredTimedLogShift(line, runDate, runStart, runFinish, appState.supervisorSelectedShift || "Day");
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

    const date = supervisorAutoDateValue("superShiftDate", supervisorAutoEntryDate());
    const shift = document.getElementById("superShiftShift").value || "Day";
    const startInput = document.getElementById("superShiftStart").value || "";
    const finishInput = document.getElementById("superShiftFinish").value || "";
    const notesInput = String(document.getElementById("superShiftNotes")?.value || "").trim();

    if (!rowIsValidDateShift(date, shift)) {
      alert("Date/shift are invalid. Weekend dates are excluded.");
      return;
    }
    if (!supervisorCanAccessShift(session, lineId, shift)) {
      alert(`You are not assigned to the ${shift} shift.`);
      return;
    }

    const selectedShiftId = selectedSupervisorShiftLogId({ line });
    let existing = null;
    if (selectedShiftId) {
      existing = (line.shiftRows || []).find((row) => row.id === selectedShiftId) || null;
      if (!existing) {
        alert("Selected shift log could not be found. Re-open it from Shift Logs and try again.");
        return;
      }
      if (existing.date !== date || existing.shift !== shift) {
        alert("Date and shift cannot be changed while editing a saved shift log.");
        return;
      }
    }

    const existingIsOpen = Boolean(existing && isOpenShiftRow(existing));
    const startTime = startInput || existing?.startTime || nowTimeHHMM();
    const finishTime = finishInput || (complete ? nowTimeHHMM() : (existing?.finishTime || startTime));
    const notes = notesInput !== "" ? notesInput : String(existing?.notes || "");

    if (!strictTimeValid(startTime) || !strictTimeValid(finishTime)) {
      alert("Shift start and finish must be in HH:MM (24h).");
      return;
    }
    const savingOpenShift = !complete && (!existing || existingIsOpen);
    if (savingOpenShift && finishTime !== startTime) {
      alert("Open shift logs keep finish equal to start. Use the Finalise pill on the log when the shift ends.");
      return;
    }
    if (!complete && existing && !existingIsOpen && finishTime === startTime) {
      alert("Finalised shift logs require a finish time different from start.");
      return;
    }
    const payload = { lineId, date, shift, startTime, finishTime, notes };
    try {
      const saved = existing?.id
        ? await patchSupervisorShiftLog(session, existing.id, payload)
        : await syncSupervisorShiftLog(session, payload);

      const savedRow = {
        ...(existing || {}),
        ...payload,
        id: saved?.id || existing?.id || "",
        submittedBy: supervisorActorName(session),
        submittedByUserId: String(saved?.submittedByUserId || session?.userId || existing?.submittedByUserId || "").trim(),
        submittedAt: saved?.submittedAt || nowIso()
      };
      upsertRowById(line.shiftRows, savedRow);
      if (complete) {
        superShiftLogIdInput.value = "";
        supervisorShiftTileEditId = "";
      } else {
        if (superShiftLogIdInput) superShiftLogIdInput.value = "";
        supervisorShiftTileEditId = "";
      }
      addAudit(
        line,
        complete ? "SUPERVISOR_SHIFT_COMPLETE" : "SUPERVISOR_SHIFT_PROGRESS",
        `${supervisorActorName(session)} ${complete ? "completed" : "updated"} ${shift} shift for ${date}`
      );
      if (complete) {
        supervisorShiftForm.reset();
        setSupervisorAutoDateValue("superShiftDate", supervisorAutoEntryDate(), { fallbackIso: supervisorAutoEntryDate() });
        document.getElementById("superShiftShift").value = shift;
      }
      if (!complete) {
        supervisorShiftForm.reset();
        if (superShiftLogIdInput) superShiftLogIdInput.value = "";
        supervisorShiftTileEditId = "";
        setSupervisorAutoDateValue("superShiftDate", supervisorAutoEntryDate(), { fallbackIso: supervisorAutoEntryDate() });
      }
      saveState();
      renderAll();
    } catch (error) {
      alert(`Could not save shift log.\n${error?.message || "Please try again."}`);
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

    const date = supervisorAutoDateValue("superRunDate", supervisorAutoEntryDate());
    const productInput = String(document.getElementById("superRunProduct").value || "").trim();
    const prodStartInput = String(document.getElementById("superRunProdStart").value || "").trim();
    const finishInput = String(document.getElementById("superRunFinish").value || "").trim();
    const unitsRaw = document.getElementById("superRunUnits").value;
    const notesInput = String(document.getElementById("superRunNotes")?.value || "").trim();

    if (!isOperationalDate(date)) {
      alert("Date is invalid. Weekend dates are excluded.");
      return;
    }

    const selectedRunLogId = String(superRunLogIdInput?.value || "").trim();
    let existing = null;
    if (selectedRunLogId) {
      existing = (line.runRows || []).find((row) => row.id === selectedRunLogId) || null;
      if (!existing) {
        alert("Selected production run could not be found. Re-open it from Production Run Logs and try again.");
        return;
      }
      if (existing.date !== date) {
        alert("Date cannot be changed while editing a saved production run.");
        return;
      }
    }

    const catalogOptions = { lineId };
    const product = catalogProductCanonicalName(productInput || existing?.product || "", catalogOptions);
    if (!product) {
      alert("Product is required.");
      return;
    }
    const existingProduct = catalogProductCanonicalName(existing?.product || "", catalogOptions);
    const productChanged = Boolean(productInput) && product !== existingProduct;
    const requiresCatalogValidation = !existing || productChanged;
    if (requiresCatalogValidation && !isCatalogProductName(product, catalogOptions)) {
      alert("Select a product from Manage Products before starting a run.");
      return;
    }

    const existingIsOpen = Boolean(existing && isOpenRunRow(existing));
    const editingFinalisedRun = Boolean(existing && !existingIsOpen);
    const productionStartTime = prodStartInput || existing?.productionStartTime || nowTimeHHMM();
    const finishTime = complete
      ? (finishInput || nowTimeHHMM())
      : (finishInput || (editingFinalisedRun ? String(existing?.finishTime || "").trim() : ""));
    const unitsProduced = unitsRaw === ""
      ? Math.max(0, num(existing?.unitsProduced))
      : Math.max(0, num(unitsRaw));
    const notes = notesInput !== "" ? notesInput : String(existing?.notes || "");
    const timingFinish = finishTime || productionStartTime;
    const patternShift = preferredTimedLogShift(
      line,
      date,
      productionStartTime,
      timingFinish,
      appState.supervisorSelectedShift || "Day"
    );

    if (!strictTimeValid(productionStartTime)) {
      alert("Production start must be HH:MM (24h).");
      return;
    }
    if (complete) {
      if (!strictTimeValid(finishTime)) {
        alert("Run finish must be HH:MM (24h).");
        return;
      }
      if (finishTime === productionStartTime) {
        alert("Finish time must be different from start time.");
        return;
      }
    } else if (!editingFinalisedRun && finishTime !== "") {
      alert("Leave finish empty while the run is in progress. Use Finalise when the run ends.");
      return;
    } else if (editingFinalisedRun) {
      if (!strictTimeValid(finishTime)) {
        alert("Finalised runs require a valid finish time in HH:MM (24h).");
        return;
      }
      if (finishTime === productionStartTime) {
        alert("Finalised runs require finish time to differ from start time.");
        return;
      }
    }

    const inputPattern = runCrewingPatternFromInput(superRunCrewingPatternInput, line, patternShift, { fallbackToIdeal: false });
    const existingPattern = normalizeRunCrewingPattern(existing?.runCrewingPattern, line, patternShift, { fallbackToIdeal: false });
    const runCrewingPattern = Object.keys(inputPattern).length ? inputPattern : existingPattern;
    if (!Object.keys(runCrewingPattern).length) {
      alert("Set crewing pattern for this run before saving.");
      return;
    }
    setRunCrewingPatternField(
      superRunCrewingPatternInput,
      superRunCrewingPatternSummary,
      line,
      patternShift,
      runCrewingPattern,
      { fallbackToIdeal: false }
    );

    const payload = {
      lineId,
      date,
      setUpStartTime: "",
      product,
      productionStartTime,
      finishTime,
      unitsProduced,
      notes,
      runCrewingPattern
    };
    try {
      const saved = existing?.id
        ? await patchSupervisorRunLog(session, existing.id, payload)
        : await syncSupervisorRunLog(session, payload);

      const savedRow = {
        ...(existing || {}),
        ...payload,
        assignedShift: patternShift || String(existing?.assignedShift || ""),
        shift: "",
        runCrewingPattern: normalizeRunCrewingPattern(saved?.runCrewingPattern || runCrewingPattern, line, patternShift, { fallbackToIdeal: false }),
        id: saved?.id || existing?.id || "",
        submittedBy: supervisorActorName(session),
        submittedByUserId: String(saved?.submittedByUserId || session?.userId || existing?.submittedByUserId || "").trim(),
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
        setSupervisorAutoDateValue("superRunDate", supervisorAutoEntryDate(), { fallbackIso: supervisorAutoEntryDate() });
        setRunCrewingPatternField(
          superRunCrewingPatternInput,
          superRunCrewingPatternSummary,
          line,
          appState.supervisorSelectedShift || "Day",
          {},
          { fallbackToIdeal: false }
        );
      }
      if (!complete) {
        supervisorRunForm.reset();
        if (superRunLogIdInput) superRunLogIdInput.value = "";
        setSupervisorAutoDateValue("superRunDate", supervisorAutoEntryDate(), { fallbackIso: supervisorAutoEntryDate() });
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

  const loadSupervisorShiftForEdit = (shiftId, { lineId: lineIdOverride = "" } = {}) => {
    const session = appState.supervisorSession;
    if (!session) return;
    const lineId = String(lineIdOverride || selectedSupervisorLineId() || "");
    if (lineId && lineId !== appState.supervisorSelectedLineId) appState.supervisorSelectedLineId = lineId;
    const line = appState.lines[lineId];
    if (!line || !session.assignedLineIds.includes(lineId)) return;
    const targetShiftId = String(shiftId || "");
    const row = (line.shiftRows || []).find((item) => String(item?.id || "") === targetShiftId);
    if (!row) {
      renderHome();
      updateSupervisorProgressButtonLabels();
      return;
    }

    suppressSupervisorSelectionReset = true;
    setSupervisorAutoDateValue("superShiftDate", row.date || supervisorAutoEntryDate(), { fallbackIso: supervisorAutoEntryDate() });
    document.getElementById("superShiftShift").value = row.shift || "Day";
    document.getElementById("superShiftStart").value = row.startTime || "";
    document.getElementById("superShiftFinish").value = isOpenShiftRow(row) ? "" : row.finishTime || "";
    document.getElementById("superShiftNotes").value = String(row.notes || "");
    superShiftLogIdInput.value = row.id || "";
    supervisorShiftTileEditId = row.id || "";
    suppressSupervisorSelectionReset = false;
    // Refresh the supervisor data entry UI so the selected pending shift shows Save immediately.
    renderHome();
    updateSupervisorProgressButtonLabels();
    const startInput = document.getElementById("superShiftStart");
    if (startInput) startInput.focus();
  };

  const loadSupervisorRunForEdit = (runId, { lineId: lineIdOverride = "" } = {}) => {
    const session = appState.supervisorSession;
    if (!session) return;
    const lineId = String(lineIdOverride || selectedSupervisorLineId() || "");
    if (lineId && lineId !== appState.supervisorSelectedLineId) appState.supervisorSelectedLineId = lineId;
    const line = appState.lines[lineId];
    if (!line || !session.assignedLineIds.includes(lineId)) return;
    const row = (line.runRows || []).find((item) => item.id === runId);
    if (!row) return;

    suppressSupervisorSelectionReset = true;
    setSupervisorAutoDateValue("superRunDate", row.date || supervisorAutoEntryDate(), { fallbackIso: supervisorAutoEntryDate() });
    document.getElementById("superRunProduct").value = row.product || "";
    document.getElementById("superRunProdStart").value = row.productionStartTime || "";
    const rowIsOpen = isOpenRunRow(row);
    document.getElementById("superRunFinish").value = rowIsOpen ? "" : row.finishTime || "";
    document.getElementById("superRunUnits").value = num(row.unitsProduced) > 0 ? formatNum(Math.max(0, num(row.unitsProduced)), 0) : "";
    document.getElementById("superRunNotes").value = String(row.notes || "");
    superRunLogIdInput.value = row.id || "";
    const rowShift = preferredTimedLogShift(
      line,
      row.date,
      row.productionStartTime,
      row.finishTime || row.productionStartTime,
      appState.supervisorSelectedShift || "Day"
    );
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

  const loadSupervisorDowntimeForEdit = (downtimeId, { lineId: lineIdOverride = "" } = {}) => {
    const session = appState.supervisorSession;
    if (!session) return;
    const lineId = String(lineIdOverride || selectedSupervisorLineId() || "");
    if (lineId && lineId !== appState.supervisorSelectedLineId) appState.supervisorSelectedLineId = lineId;
    const line = appState.lines[lineId];
    if (!line || !session.assignedLineIds.includes(lineId)) return;
    const row = (line.downtimeRows || []).find((item) => item.id === downtimeId);
    if (!row) return;

    const parsedReason = parseDowntimeReasonParts(row.reason, row.equipment);
    const reasonCategory = row.reasonCategory || parsedReason.reasonCategory;
    const reasonDetail = row.reasonDetail || parsedReason.reasonDetail;
    const reasonNote = String((row.reasonNote ?? parsedReason.reasonNote) || "");
    suppressSupervisorSelectionReset = true;
    setSupervisorAutoDateValue("superDownDate", row.date || supervisorAutoEntryDate(), { fallbackIso: supervisorAutoEntryDate() });
    document.getElementById("superDownStart").value = row.downtimeStart || "";
    document.getElementById("superDownFinish").value = isOpenDowntimeRow(row) ? "" : row.downtimeFinish || "";
    if (supervisorDownReasonCategory) supervisorDownReasonCategory.value = reasonCategory || "";
    if (supervisorDownReasonDetail) {
      setDowntimeDetailOptions(supervisorDownReasonDetail, line, reasonCategory || "", reasonDetail || "");
    }
    document.getElementById("superDownReasonNote").value = reasonNote;
    document.getElementById("superDownNotes").value = String(row.notes || "");
    superDownLogIdInput.value = row.id || "";
    suppressSupervisorSelectionReset = false;
    updateSupervisorProgressButtonLabels();
    const startInput = document.getElementById("superDownStart");
    if (startInput) startInput.focus();
  };

  const openSupervisorSubmittedLogForEdit = (type, logId, lineId = "") => {
    const safeType = type === "run" || type === "downtime" ? type : "shift";
    const safeLineId = String(lineId || selectedSupervisorLineId() || "").trim();
    if (safeLineId) appState.supervisorSelectedLineId = safeLineId;
    appState.supervisorMainTab = "supervisorData";
    appState.supervisorTab =
      safeType === "run"
        ? "superRun"
        : safeType === "downtime"
          ? "superDown"
          : "superShift";
    saveState();
    renderHome();
    if (safeType === "run") {
      loadSupervisorRunForEdit(logId, { lineId: safeLineId });
      return;
    }
    if (safeType === "downtime") {
      loadSupervisorDowntimeForEdit(logId, { lineId: safeLineId });
      return;
    }
    loadSupervisorShiftForEdit(logId, { lineId: safeLineId });
  };

  const handleSupervisorSubmittedLogAction = async (action, type, logId, lineId = "") => {
    const safeAction = String(action || "").trim().toLowerCase();
    const safeType = type === "run" || type === "downtime" ? type : "shift";
    const safeLogId = String(logId || "").trim();
    const safeLineId = String(lineId || "").trim();
    if (!safeAction || !safeLogId) return;
    if (safeAction === "edit") {
      openSupervisorSubmittedLogForEdit(safeType, safeLogId, safeLineId);
      return;
    }
    if (safeAction !== "finalise") return;
    if (safeType === "run") {
      await completeSupervisorRunById(safeLogId, { lineId: safeLineId });
      return;
    }
    if (safeType === "downtime") {
      await completeSupervisorDowntimeById(safeLogId, { lineId: safeLineId });
      return;
    }
    await completeSupervisorShiftById(safeLogId, { lineId: safeLineId });
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
      startTime,
      finishTime,
      notes: String(row.notes || "")
    };
    try {
      const saved = await patchSupervisorShiftLog(session, row.id, payload);
      const savedRow = {
        ...row,
        ...payload,
        id: saved?.id || row.id || "",
        submittedBy: supervisorActorName(session),
        submittedByUserId: String(saved?.submittedByUserId || session?.userId || row?.submittedByUserId || "").trim(),
        submittedAt: saved?.submittedAt || nowIso()
      };
      upsertRowById(line.shiftRows, savedRow);
      if (superShiftLogIdInput.value === row.id) superShiftLogIdInput.value = "";
      supervisorShiftTileEditId = "";
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
    const assignedShift = preferredTimedLogShift(
      line,
      row.date,
      productionStartTime,
      finishTime,
      appState.supervisorSelectedShift || "Day"
    );
    const payload = {
      lineId,
      date: row.date,
      setUpStartTime: "",
      product: row.product || "Run",
      productionStartTime,
      finishTime,
      unitsProduced,
      notes: String(row.notes || "")
    };
    try {
      const saved = await patchSupervisorRunLog(session, row.id, payload);
      const savedRow = {
        ...row,
        ...payload,
        assignedShift: assignedShift || String(row?.assignedShift || ""),
        shift: "",
        id: saved?.id || row.id || "",
        submittedBy: supervisorActorName(session),
        submittedByUserId: String(saved?.submittedByUserId || session?.userId || row?.submittedByUserId || "").trim(),
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
    const payload = {
      lineId,
      date: row.date,
      downtimeStart,
      downtimeFinish,
      equipment,
      reasonCategory,
      reasonDetail,
      reasonNote,
      notes: String(row.notes || ""),
      reason: buildDowntimeReasonText(line, reasonCategory, reasonDetail, reasonNote)
    };
    const assignedShift = preferredTimedLogShift(
      line,
      row.date,
      downtimeStart,
      downtimeFinish,
      appState.supervisorSelectedShift || "Day"
    );
    try {
      const saved = await patchSupervisorDowntimeLog(session, row.id, payload);
      const savedRow = {
        ...row,
        ...payload,
        assignedShift: assignedShift || String(row?.assignedShift || ""),
        shift: "",
        id: saved?.id || row.id || "",
        submittedBy: supervisorActorName(session),
        submittedByUserId: String(saved?.submittedByUserId || session?.userId || row?.submittedByUserId || "").trim(),
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

    const date = supervisorAutoDateValue("superDownDate", supervisorAutoEntryDate());
    const startInput = document.getElementById("superDownStart").value || "";
    const finishInput = document.getElementById("superDownFinish").value || "";
    const reasonCategoryInput = document.getElementById("superDownReasonCategory").value || "";
    const reasonDetailInput = document.getElementById("superDownReasonDetail").value || "";
    const reasonNoteInput = String(document.getElementById("superDownReasonNote").value || "").trim();
    const notesInput = String(document.getElementById("superDownNotes")?.value || "").trim();

    if (!isOperationalDate(date)) {
      alert("Date is invalid. Weekend dates are excluded.");
      return;
    }

    const selectedDownLogId = String(superDownLogIdInput?.value || "").trim();
    let existing = null;
    if (selectedDownLogId) {
      existing = (line.downtimeRows || []).find((row) => row.id === selectedDownLogId) || null;
      if (!existing) {
        alert("Selected downtime log could not be found. Re-open it from Downtime Logs and try again.");
        return;
      }
      if (existing.date !== date) {
        alert("Date cannot be changed while editing a saved downtime log.");
        return;
      }
    }

    const downtimeStart = startInput || existing?.downtimeStart || nowTimeHHMM();
    const downtimeFinish = finishInput || (complete ? nowTimeHHMM() : existing?.downtimeFinish || downtimeStart);
    const reasonCategory = reasonCategoryInput || existing?.reasonCategory || "";
    const reasonDetail = reasonDetailInput || existing?.reasonDetail || "";
    const reasonNote = reasonNoteInput !== "" ? reasonNoteInput : existing?.reasonNote || "";
    const notes = notesInput !== "" ? notesInput : String(existing?.notes || "");
    const equipment = reasonCategory === "Equipment" ? (reasonDetail || existing?.equipment || "") : "";
    const reason = buildDowntimeReasonText(line, reasonCategory, reasonDetail, reasonNote);
    const assignedShift = preferredTimedLogShift(
      line,
      date,
      downtimeStart,
      downtimeFinish,
      appState.supervisorSelectedShift || "Day"
    );

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
    const payload = {
      lineId,
      date,
      downtimeStart,
      downtimeFinish,
      equipment,
      reason,
      reasonCategory,
      reasonDetail,
      reasonNote,
      notes
    };
    try {
      const saved = existing?.id
        ? await patchSupervisorDowntimeLog(session, existing.id, payload)
        : await syncSupervisorDowntimeLog(session, payload);

      const savedRow = {
        ...(existing || {}),
        ...payload,
        assignedShift: assignedShift || String(existing?.assignedShift || ""),
        shift: "",
        excludeFromCalculation: normalizeDowntimeCalculationFlag(
          saved?.excludeFromCalculation ?? existing?.excludeFromCalculation
        ),
        id: saved?.id || existing?.id || "",
        submittedBy: supervisorActorName(session),
        submittedByUserId: String(saved?.submittedByUserId || session?.userId || existing?.submittedByUserId || "").trim(),
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
        setSupervisorAutoDateValue("superDownDate", supervisorAutoEntryDate(), { fallbackIso: supervisorAutoEntryDate() });
        if (supervisorDownReasonCategory) supervisorDownReasonCategory.value = "";
        refreshSupervisorDowntimeDetailOptions(selectedSupervisorLine());
      }
      if (!complete) {
        supervisorDownForm.reset();
        if (superDownLogIdInput) superDownLogIdInput.value = "";
        if (supervisorDownReasonCategory) supervisorDownReasonCategory.value = "";
        setSupervisorAutoDateValue("superDownDate", supervisorAutoEntryDate(), { fallbackIso: supervisorAutoEntryDate() });
        refreshSupervisorDowntimeDetailOptions(selectedSupervisorLine());
      }
      saveState();
      renderAll();
    } catch (error) {
      alert(`Could not save downtime log.\n${error?.message || "Please try again."}`);
    }
  };

  superShiftSaveProgressBtn.addEventListener("click", () => submitSupervisorShift({ complete: false }));
  superRunSaveProgressBtn.addEventListener("click", () => submitSupervisorRun({ complete: false }));
  superDownSaveProgressBtn.addEventListener("click", () => submitSupervisorDowntime({ complete: false }));
  [supervisorEntryList, supervisorEntryCards].filter(Boolean).forEach((container) => {
    container.addEventListener("click", async (event) => {
      const actionBtn = event.target.closest("[data-super-entry-action]");
      if (!actionBtn) return;
      const action = actionBtn.getAttribute("data-super-entry-action");
      const type = actionBtn.getAttribute("data-super-entry-type");
      const logId = actionBtn.getAttribute("data-super-entry-id");
      const lineId = actionBtn.getAttribute("data-super-entry-line-id");
      await handleSupervisorSubmittedLogAction(action, type, logId, lineId);
    });
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

  const homeDnDState = {
    type: "",
    draggedGroupId: "",
    draggedLineId: "",
    sourceGroupKey: ""
  };

  const resetHomeDnDState = () => {
    homeDnDState.type = "";
    homeDnDState.draggedGroupId = "";
    homeDnDState.draggedLineId = "";
    homeDnDState.sourceGroupKey = "";
    cards.classList.remove("home-dnd-group-active", "home-dnd-line-active");
    cards.querySelectorAll(".is-dragging").forEach((node) => node.classList.remove("is-dragging"));
  };

  const isManagerHomeDnDEnabled = () =>
    appState.appMode === "manager" &&
    Boolean(managerBackendSession.backendToken) &&
    cards.classList.contains("line-cards-grouped");

  const findGroupSectionById = (groupId) =>
    Array.from(cards.querySelectorAll("[data-line-group-section]")).find(
      (section) => String(section.getAttribute("data-line-group-section") || "") === String(groupId || "")
    ) || null;

  const findLineCardById = (lineId) =>
    Array.from(cards.querySelectorAll("[data-line-card-id]")).find(
      (card) => String(card.getAttribute("data-line-card-id") || "") === String(lineId || "")
    ) || null;

  const persistGroupOrderFromDom = async () => {
    const orderedGroupIds = Array.from(cards.querySelectorAll("[data-line-group-section]"))
      .map((section) => String(section.getAttribute("data-line-group-section") || "").trim())
      .filter((groupId) => groupId && groupId !== "ungrouped");
    if (orderedGroupIds.length < 2) return;
    const allGroupIds = normalizeLineGroups(appState.lineGroups).map((group) => group.id);
    const fullGroupOrder = [
      ...orderedGroupIds,
      ...allGroupIds.filter((groupId) => !orderedGroupIds.includes(groupId))
    ];
    const session = await ensureManagerBackendSession();
    const response = await apiRequest("/api/line-groups/reorder", {
      method: "PATCH",
      token: session.backendToken,
      body: { groupIds: fullGroupOrder }
    });
    const backendGroups = normalizeLineGroups(response?.lineGroups);
    if (backendGroups.length) {
      appState.lineGroups = backendGroups;
    } else {
      const displayOrderById = Object.fromEntries(fullGroupOrder.map((groupId, index) => [groupId, index]));
      appState.lineGroups = normalizeLineGroups(
        normalizeLineGroups(appState.lineGroups).map((group) => ({
          ...group,
          displayOrder: Number.isFinite(displayOrderById[group.id]) ? displayOrderById[group.id] : group.displayOrder
        }))
      );
    }
    saveState();
  };

  const persistLineOrderFromDom = async (groupKey) => {
    const safeGroupKey = String(groupKey || "").trim();
    if (!safeGroupKey) return;
    const targetList = Array.from(cards.querySelectorAll("[data-line-card-list]")).find(
      (list) => String(list.getAttribute("data-line-card-list") || "") === safeGroupKey
    );
    if (!targetList) return;
    const lineIds = Array.from(targetList.querySelectorAll("[data-line-card-id]"))
      .map((card) => String(card.getAttribute("data-line-card-id") || "").trim())
      .filter(Boolean);
    if (lineIds.length < 2) return;
    const nextGroupId = safeGroupKey === "ungrouped" ? null : safeGroupKey;
    const session = await ensureManagerBackendSession();
    const response = await apiRequest("/api/lines/reorder", {
      method: "PATCH",
      token: session.backendToken,
      body: {
        lineIds,
        groupId: nextGroupId
      }
    });
    const syncedLines = Array.isArray(response?.lines) ? response.lines : [];
    if (syncedLines.length) {
      syncedLines.forEach((lineRow) => {
        const line = appState.lines?.[lineRow.id];
        if (!line) return;
        line.groupId = String(lineRow?.groupId || "").trim();
        const nextOrder = Number(lineRow?.displayOrder);
        if (Number.isFinite(nextOrder)) line.displayOrder = Math.max(0, Math.floor(nextOrder));
      });
    } else {
      lineIds.forEach((lineId, index) => {
        const line = appState.lines?.[lineId];
        if (!line) return;
        line.groupId = nextGroupId || "";
        line.displayOrder = index;
      });
    }
    saveState();
  };

  cards.addEventListener("dragstart", (event) => {
    const groupHandle = event.target.closest("[data-line-group-drag-handle]");
    if (groupHandle) {
      if (!isManagerHomeDnDEnabled()) {
        event.preventDefault();
        return;
      }
      const section = groupHandle.closest("[data-line-group-section]");
      const groupId = String(section?.getAttribute("data-line-group-section") || "").trim();
      if (!groupId || groupId === "ungrouped") {
        event.preventDefault();
        return;
      }
      homeDnDState.type = "group";
      homeDnDState.draggedGroupId = groupId;
      cards.classList.add("home-dnd-group-active");
      if (section) section.classList.add("is-dragging");
      if (event.dataTransfer) {
        event.dataTransfer.effectAllowed = "move";
        event.dataTransfer.setData("text/plain", `group:${groupId}`);
      }
      return;
    }

    const lineHandle = event.target.closest("[data-line-card-drag-handle]");
    if (!lineHandle) return;
    if (!isManagerHomeDnDEnabled()) {
      event.preventDefault();
      return;
    }
    const lineCard = lineHandle.closest("[data-line-card-id]");
    const lineId = String(lineCard?.getAttribute("data-line-card-id") || "").trim();
    const groupKey = String(lineCard?.getAttribute("data-line-card-group") || "").trim();
    if (!lineId || !groupKey) {
      event.preventDefault();
      return;
    }
    homeDnDState.type = "line";
    homeDnDState.draggedLineId = lineId;
    homeDnDState.sourceGroupKey = groupKey;
    cards.classList.add("home-dnd-line-active");
    if (lineCard) lineCard.classList.add("is-dragging");
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", `line:${lineId}`);
    }
  });

  cards.addEventListener("dragover", (event) => {
    if (!homeDnDState.type || !isManagerHomeDnDEnabled()) return;
    if (homeDnDState.type === "group") {
      const targetSection = event.target.closest("[data-line-group-section]");
      if (!targetSection) return;
      const targetGroupId = String(targetSection.getAttribute("data-line-group-section") || "").trim();
      if (!targetGroupId || targetGroupId === "ungrouped" || targetGroupId === homeDnDState.draggedGroupId) return;
      const draggedSection = findGroupSectionById(homeDnDState.draggedGroupId);
      if (!draggedSection) return;
      event.preventDefault();
      const rect = targetSection.getBoundingClientRect();
      const beforeTarget = event.clientY < rect.top + rect.height / 2;
      const referenceNode = beforeTarget ? targetSection : targetSection.nextElementSibling;
      if (referenceNode !== draggedSection) {
        cards.insertBefore(draggedSection, referenceNode);
      }
      return;
    }

    if (homeDnDState.type === "line") {
      const targetList = event.target.closest("[data-line-card-list]");
      if (!targetList) return;
      const targetGroupKey = String(targetList.getAttribute("data-line-card-list") || "").trim();
      if (targetGroupKey !== homeDnDState.sourceGroupKey) return;
      const draggedCard = findLineCardById(homeDnDState.draggedLineId);
      if (!draggedCard) return;
      event.preventDefault();
      const targetCard = event.target.closest("[data-line-card-id]");
      if (!targetCard) {
        if (targetList.lastElementChild !== draggedCard) targetList.append(draggedCard);
        return;
      }
      if (targetCard === draggedCard) return;
      const rect = targetCard.getBoundingClientRect();
      const beforeTarget = event.clientY < rect.top + rect.height / 2;
      const referenceNode = beforeTarget ? targetCard : targetCard.nextElementSibling;
      if (referenceNode !== draggedCard) {
        targetList.insertBefore(draggedCard, referenceNode);
      }
    }
  });

  cards.addEventListener("drop", async (event) => {
    if (!homeDnDState.type) return;
    event.preventDefault();
    const dropType = homeDnDState.type;
    const sourceGroupKey = homeDnDState.sourceGroupKey;
    try {
      if (dropType === "group") {
        await persistGroupOrderFromDom();
      } else if (dropType === "line" && sourceGroupKey) {
        await persistLineOrderFromDom(sourceGroupKey);
      }
      renderHome();
    } catch (error) {
      console.warn("Home reorder sync failed:", error);
      alert(`Could not save the new order.\n${error?.message || "Please try again."}`);
      await refreshHostedState();
    } finally {
      resetHomeDnDState();
    }
  });

  cards.addEventListener("dragend", () => {
    resetHomeDnDState();
  });

  cards.addEventListener("click", async (event) => {
    const groupToggleBtn = event.target.closest("[data-line-group-toggle]");
    if (groupToggleBtn) {
      const groupKey = String(groupToggleBtn.getAttribute("data-line-group-toggle") || "");
      if (!groupKey) return;
      const expanded = isHomeLineGroupExpanded(groupKey);
      setHomeLineGroupExpanded(groupKey, !expanded);
      saveState();
      renderHome();
      return;
    }

    const editBtn = event.target.closest("[data-edit-line]");
    if (editBtn) {
      const id = editBtn.getAttribute("data-edit-line");
      if (!id || !appState.lines[id]) return;
      openEditLineModal(id);
      return;
    }

    if (event.target.closest("[data-line-card-drag-handle], [data-line-group-drag-handle]")) return;
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

  if (editLineDeleteBtn) {
    editLineDeleteBtn.addEventListener("click", async () => {
      const lineId = editingLineId;
      if (!lineId || !appState.lines?.[lineId]) return;
      editLineDeleteBtn.disabled = true;
      try {
        await deleteLineById(lineId);
      } finally {
        editLineDeleteBtn.disabled = false;
      }
    });
  }

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
      if (!state) return;
      const target = parseManagerDataTabId(btn.dataset.dataTab);
      if (!target) return;
      state.activeDataTab = target;
      pendingManagerDataTabRestore = { lineId: "", tabId: "" };
      saveState();
      setActiveDataSubtab();
    });
  });
}

function bindVisualiserControls() {
  const dateInputs = Array.from(document.querySelectorAll("[data-manager-date]"));
  const shiftButtons = Array.from(document.querySelectorAll(".shift-option[data-shift], .shift-option[data-day-shift]"));
  const crewShiftButtons = Array.from(document.querySelectorAll(".shift-option[data-crew-shift]"));
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
  const dayKpiTriggers = Array.from(document.querySelectorAll("[data-day-kpi-metric]"));
  const saveStageSettingsBtn = document.getElementById("saveStageCrewSettingsBtn");
  const dayKeyStageSelect = document.getElementById("dayKeyStageSelect");
  let layoutDragState = null;

  const clamp = (value, min, max) => Math.max(min, Math.min(max, value));
  const setLayoutEditButtonUI = () => {
    const active = Boolean(state?.visualEditMode);
    if (editBtn) {
      editBtn.classList.toggle("active", active);
      editBtn.textContent = active ? "Done Editing" : "Edit Layout";
    }
    if (addLineBtn) {
      addLineBtn.disabled = true;
      addLineBtn.classList.add("hidden");
    }
    if (addArrowBtn) {
      addArrowBtn.disabled = true;
      addArrowBtn.classList.add("hidden");
    }
    if (uploadShapeBtn) {
      uploadShapeBtn.disabled = !active;
      uploadShapeBtn.classList.toggle("hidden", !active);
    }
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

  const normalizedSelectedDate = normalizeWeekdayIsoDate(state?.selectedDate || todayISO(), { direction: -1 });
  if (state) state.selectedDate = normalizedSelectedDate;
  dateInputs.forEach((dateInput) => {
    dateInput.value = normalizedSelectedDate;
  });
  setShiftToggleUI();
  setLayoutEditButtonUI();

  dateInputs.forEach((dateInput) => {
    dateInput.addEventListener("change", () => {
      if (!state) return;
      state.selectedDate = normalizeWeekdayIsoDate(dateInput.value || state.selectedDate || todayISO(), { direction: -1 });
      dateInputs.forEach((input) => {
        input.value = state.selectedDate;
      });
      saveState();
      renderAll();
    });
  });

  shiftButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      if (!state) return;
      const nextShift = btn.dataset.shift || btn.dataset.dayShift;
      if (!nextShift || nextShift === state.selectedShift) return;
      state.selectedShift = nextShift;
      setShiftToggleUI();
      saveState();
      renderAll();
    });
  });
  if (dayKeyStageSelect) {
    dayKeyStageSelect.addEventListener("change", () => {
      if (!state) return;
      state.dayVisualiserKeyStageId = normalizeDayVisualiserKeyStageId(state, dayKeyStageSelect.value);
      saveState();
      renderVisualiser();
    });
  }
  dayKpiTriggers.forEach((trigger) => {
    const activate = () => {
      const metricKey = String(trigger.dataset.dayKpiMetric || "").trim();
      if (!metricKey) return;
      openDayKpiTrend(metricKey);
    };
    trigger.addEventListener("click", activate);
    trigger.addEventListener("keydown", (event) => {
      if (event.key !== "Enter" && event.key !== " ") return;
      event.preventDefault();
      activate();
    });
  });
  crewShiftButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      if (!state) return;
      const nextCrewShift = String(btn.dataset.crewShift || "");
      if (!CREW_SETTINGS_SHIFTS.includes(nextCrewShift)) return;
      if (crewSettingsShiftForLine(state) === nextCrewShift) return;
      state.crewSettingsShift = nextCrewShift;
      setShiftToggleUI();
      saveState();
      renderCrewInputs();
    });
  });

  if (saveStageSettingsBtn) {
    saveStageSettingsBtn.addEventListener("click", async () => {
      await saveCurrentLineSettings();
    });
  }
  syncLineSettingsSaveUI();

  prevButtons.forEach((btn) => btn.addEventListener("click", () => moveShift(-1)));
  nextButtons.forEach((btn) => btn.addEventListener("click", () => moveShift(1)));

  if (lineTrendPrevBtn) lineTrendPrevBtn.addEventListener("click", () => moveLineTrendPeriod(-1));
  if (lineTrendNextBtn) lineTrendNextBtn.addEventListener("click", () => moveLineTrendPeriod(1));
  lineTrendRangeButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      if (!state) return;
      const range = String(btn.dataset.lineTrendRange || "").toLowerCase();
      if (!LINE_TREND_RANGES.includes(range) || range === state.lineTrendRange) return;
      state.lineTrendRange = range;
      state.lineTrendLegendFocusKey = "";
      saveState();
      renderLineTrends();
    });
  });

  if (editBtn) {
    editBtn.addEventListener("click", () => {
      if (!state) return;
      const wasEditing = Boolean(state.visualEditMode);
      state.visualEditMode = !state.visualEditMode;
      setLayoutEditButtonUI();
      if (wasEditing && !state.visualEditMode && deferredLineModelSyncIds.has(String(state.id || ""))) {
        saveState({ syncModel: true, forceSyncModel: true });
      } else {
        saveState();
      }
      renderVisualiser();
    });
  }

  if (addLineBtn) {
    addLineBtn.addEventListener("click", () => {
      if (!state || !state.visualEditMode) return;
      state.flowGuides = Array.isArray(state.flowGuides) ? state.flowGuides : [];
      state.flowGuides.push(createGuide("line"));
      addAudit(state, "LAYOUT_ADD_LINE", "Flow line guide added");
      saveState({ syncModel: true });
      renderVisualiser();
    });
  }

  if (addArrowBtn) {
    addArrowBtn.addEventListener("click", () => {
      if (!state || !state.visualEditMode) return;
      state.flowGuides = Array.isArray(state.flowGuides) ? state.flowGuides : [];
      state.flowGuides.push(createGuide("arrow"));
      addAudit(state, "LAYOUT_ADD_ARROW", "Flow arrow guide added");
      saveState({ syncModel: true });
      renderVisualiser();
    });
  }

  if (uploadShapeBtn && uploadShapeInput) {
    uploadShapeBtn.addEventListener("click", () => {
      if (!state || !state.visualEditMode) return;
      uploadShapeInput.click();
    });
  }

  if (uploadShapeInput) {
    uploadShapeInput.addEventListener("change", async (event) => {
      if (!state || !state.visualEditMode) return;
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
  }

  if (map) {
    map.addEventListener("mousedown", (event) => {
      if (!state || !state.visualEditMode) return;
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
  }

  window.addEventListener("mousemove", (event) => {
    if (!state || !map || !state.visualEditMode || !layoutDragState) return;
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
    if (state && layoutDragState) {
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

function setStageSettingsDataSourceHint(stageId, dataSourceId, assignments = null) {
  const hint = document.getElementById("stageSettingsDataSourceHint");
  if (!hint) return;
  hint.classList.remove("is-warning");

  const safeStageId = String(stageId || "").trim();
  const safeLineId = String(state?.id || "").trim();
  const safeSourceId = String(dataSourceId || "").trim();

  if (!safeSourceId) {
    hint.textContent = "No data source connected.";
    return;
  }

  const source = dataSourceById(safeSourceId);
  const conflict = conflictingDataSourceAssignment(safeSourceId, safeLineId, safeStageId, assignments);
  if (conflict) {
    hint.textContent = `This source is already connected to ${conflict.lineName} - ${conflict.stageName}.`;
    hint.classList.add("is-warning");
    return;
  }
  hint.textContent = `Connected to ${dataSourceDisplayLabel(source)}.`;
}

function renderStageSettingsDataSourceOptions(stageId, dataSourceId = "") {
  const select = document.getElementById("stageSettingsDataSource");
  const hint = document.getElementById("stageSettingsDataSourceHint");
  if (!select) return;

  const safeStageId = String(stageId || "").trim();
  const safeLineId = String(state?.id || "").trim();
  const safeSelectedId = String(dataSourceId || "").trim();
  const sources = (appState.dataSources || []).filter((source) => source.isActive !== false);
  const assignments = dataSourceAssignments(appState.lines);
  const sourceById = new Map(sources.map((source) => [source.id, source]));
  const options = [`<option value="">Not connected</option>`];

  sources.forEach((source) => {
    const conflict = conflictingDataSourceAssignment(source.id, safeLineId, safeStageId, assignments);
    const selected = source.id === safeSelectedId;
    const disabled = Boolean(conflict) && !selected;
    const optionLabel = conflict
      ? `${dataSourceDisplayLabel(source)} - In use by ${conflict.lineName} - ${conflict.stageName}`
      : dataSourceDisplayLabel(source);
    options.push(
      `<option value="${htmlEscape(source.id)}"${selected ? " selected" : ""}${disabled ? " disabled" : ""}>${htmlEscape(optionLabel)}</option>`
    );
  });

  if (safeSelectedId && !sourceById.has(safeSelectedId)) {
    options.push(`<option value="${htmlEscape(safeSelectedId)}" selected>Current source (unavailable)</option>`);
  }

  select.innerHTML = options.join("");
  if (safeSelectedId) select.value = safeSelectedId;
  setStageSettingsDataSourceHint(safeStageId, select.value, assignments);
  if (!sources.length && hint) {
    hint.textContent = "No data sources are available.";
    hint.classList.add("is-warning");
  }
}

function openStageSettingsModal(stageId) {
  const stage = getStages().find((item) => item.id === stageId);
  if (!stage) return;
  const defaults = stageDefaultSize(stage);
  const safeWidth = Math.max(2, num(stage.w) || defaults.w);
  const safeHeight = Math.max(1, num(stage.h) || defaults.h);
  document.getElementById("stageSettingsId").value = stage.id;
  document.getElementById("stageSettingsName").value = stageBaseName(stage.name) || "Stage";
  document.getElementById("stageSettingsType").value = stage.kind === "transfer" ? "transfer" : stage.group || "main";
  document.getElementById("stageSettingsCrewDay").value = num(state.crewsByShift?.Day?.[stage.id]?.crew ?? stage.crew);
  document.getElementById("stageSettingsCrewNight").value = num(state.crewsByShift?.Night?.[stage.id]?.crew ?? stage.crew);
  document.getElementById("stageSettingsMaxThroughput").value = stageMaxThroughput(stage.id);
  document.getElementById("stageSettingsWidth").value = roundToDecimals(safeWidth, 1).toFixed(1);
  document.getElementById("stageSettingsHeight").value = roundToDecimals(safeHeight, 1).toFixed(1);
  renderStageSettingsDataSourceOptions(stage.id, String(stage.dataSourceId || "").trim());

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
  const dataSourceSelect = document.getElementById("stageSettingsDataSource");

  closeBtn.addEventListener("click", closeStageSettingsModal);
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) closeStageSettingsModal();
  });
  if (dataSourceSelect) {
    dataSourceSelect.addEventListener("change", () => {
      const stageId = document.getElementById("stageSettingsId").value;
      setStageSettingsDataSourceHint(stageId, dataSourceSelect.value);
    });
  }

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
    const width = roundToDecimals(Math.max(2, num(document.getElementById("stageSettingsWidth").value)), 1);
    const height = roundToDecimals(Math.max(1, num(document.getElementById("stageSettingsHeight").value)), 1);
    const nextDataSourceIdRaw = String(document.getElementById("stageSettingsDataSource")?.value || "").trim();
    const nextDataSourceId = nextDataSourceIdRaw
      && (dataSourceById(nextDataSourceIdRaw) || nextDataSourceIdRaw === String(stage.dataSourceId || "").trim())
      ? nextDataSourceIdRaw
      : "";
    const dataSourceConflict = conflictingDataSourceAssignment(nextDataSourceId, state?.id, stage.id);
    if (dataSourceConflict) {
      alert(`This data source is already connected to ${dataSourceConflict.lineName} - ${dataSourceConflict.stageName}.`);
      renderStageSettingsDataSourceOptions(stage.id, String(stage.dataSourceId || "").trim());
      return;
    }
    const defaults = stageDefaultSize(stage);

    stage.name = name;
    stage.group = type === "transfer" ? "prep" : type;
    stage.kind = type === "transfer" ? "transfer" : undefined;
    stage.w = width > 0 ? width : roundToDecimals(defaults.w, 1);
    stage.h = height > 0 ? height : roundToDecimals(defaults.h, 1);
    stage.match = name
      .toLowerCase()
      .split(/[^a-z0-9]+/)
      .filter(Boolean);
    stage.dataSourceId = nextDataSourceId;

    if (!state.crewsByShift.Day[stage.id]) state.crewsByShift.Day[stage.id] = {};
    if (!state.crewsByShift.Night[stage.id]) state.crewsByShift.Night[stage.id] = {};
    state.crewsByShift.Day[stage.id].crew = crewDay;
    state.crewsByShift.Night[stage.id].crew = crewNight;
    stage.crew = num(state.crewsByShift[state.selectedShift]?.[stage.id]?.crew ?? crewDay);

    if (!state.stageSettings[stage.id]) state.stageSettings[stage.id] = {};
    state.stageSettings[stage.id].maxThroughput = maxThroughput;

    const stageIndex = getStages().findIndex((s) => s.id === stage.id);
    const sourceLabel = nextDataSourceId ? dataSourceDisplayLabel(dataSourceById(nextDataSourceId)) : "Not connected";
    addAudit(
      state,
      "EDIT_STAGE_SETTINGS",
      `Stage updated: ${stageDisplayName(stage, stageIndex)} | Data source: ${sourceLabel}`
    );
    saveState({ syncModel: true });
    closeStageSettingsModal();
    renderAll();
  });
}

function bindTrendModal() {
  const overlay = document.getElementById("trendModal");
  const closeBtn = document.getElementById("closeTrendModal");
  const exportBtn = document.getElementById("trendExportCsv");
  const dailyBtn = document.getElementById("trendDaily");
  const monthlyBtn = document.getElementById("trendMonthly");
  const zeroDaysToggleBtn = document.getElementById("trendZeroDaysToggle");
  const prevBtn = document.getElementById("trendPrevMonth");
  const nextBtn = document.getElementById("trendNextMonth");

  closeBtn.addEventListener("click", closeTrendModal);
  if (exportBtn) exportBtn.addEventListener("click", exportTrendModalCsv);

  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) closeTrendModal();
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeTrendModal();
  });

  dailyBtn.addEventListener("click", () => {
    if (!state) return;
    state.trendGranularity = "daily";
    saveState();
    renderTrendModalContent();
  });

  monthlyBtn.addEventListener("click", () => {
    if (!state) return;
    state.trendGranularity = "monthly";
    saveState();
    renderTrendModalContent();
  });

  if (zeroDaysToggleBtn) {
    zeroDaysToggleBtn.addEventListener("click", () => {
      if (!state) return;
      state.trendShowZeroDays = state.trendShowZeroDays === false;
      saveState();
      renderTrendModalContent();
    });
  }

  prevBtn.addEventListener("click", () => {
    if (!state) return;
    if (state.trendGranularity === "monthly") {
      state.trendMonth = addMonths(state.trendMonth || monthKey(state.selectedDate), -1);
    } else {
      const allDates = trendDatesForShift(state.selectedShift, derivedData());
      const cursorDate = syncTrendDateCursorToAvailableDates(allDates);
      const currentIndex = allDates.indexOf(cursorDate);
      if (currentIndex > 0) state.trendDateCursor = allDates[currentIndex - 1];
    }
    saveState();
    renderTrendModalContent();
  });

  nextBtn.addEventListener("click", () => {
    if (!state) return;
    if (state.trendGranularity === "monthly") {
      state.trendMonth = addMonths(state.trendMonth || monthKey(state.selectedDate), 1);
    } else {
      const allDates = trendDatesForShift(state.selectedShift, derivedData());
      const cursorDate = syncTrendDateCursorToAvailableDates(allDates);
      const currentIndex = allDates.indexOf(cursorDate);
      if (currentIndex >= 0 && currentIndex < allDates.length - 1) state.trendDateCursor = allDates[currentIndex + 1];
    }
    saveState();
    renderTrendModalContent();
  });
}

function openDayKpiTrend(metricKey) {
  const safeMetricKey = String(metricKey || "").trim();
  if (!state || !safeMetricKey) return;
  const requiresKeyStage = safeMetricKey === "keyStageCrew" || safeMetricKey === "keyStageRatePerCrew";
  if (requiresKeyStage && !normalizeDayVisualiserKeyStageId(state, state.dayVisualiserKeyStageId)) {
    alert("Select a key stage first.");
    return;
  }
  state.trendDateCursor = normalizeWeekdayIsoDate(state.selectedDate || todayISO(), { direction: -1 });
  trendModalContext = { type: "dayKpi", metricKey: safeMetricKey };
  renderTrendModalContent();
  openTrendModal();
}

function trendPointTargetDate(point) {
  const targetDate = String(point?.targetDate || point?.date || "").trim();
  if (!isIsoDateValue(targetDate)) return "";
  return normalizeWeekdayIsoDate(targetDate, { direction: -1 });
}

function openDayVisualiserForTrendPoint(point) {
  if (!state) return;
  const targetDate = trendPointTargetDate(point);
  if (!targetDate) return;
  state.selectedDate = targetDate;
  state.trendDateCursor = targetDate;
  appState.managerLineTab = "dayVisualiser";
  closeTrendModal();
  saveState();
  renderAll();
}

function bindLineTrendDetailModal() {
  const overlay = document.getElementById("lineTrendDetailModal");
  const closeBtn = document.getElementById("closeLineTrendDetailModal");
  if (!overlay || !closeBtn) return;
  closeBtn.addEventListener("click", closeLineTrendDetailModal);
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) closeLineTrendDetailModal();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeLineTrendDetailModal();
  });
}

function moveShift(direction) {
  if (!state) return;
  state.selectedDate = shiftWeekdayIsoDate(state.selectedDate || todayISO(), direction < 0 ? -1 : 1);

  document.querySelectorAll("[data-manager-date]").forEach((input) => {
    input.value = state.selectedDate;
  });
  saveState();
  renderAll();
}

function moveLineTrendPeriod(direction) {
  if (!state) return;
  const date = parseDateLocal(normalizeWeekdayIsoDate(state.selectedDate || todayISO(), { direction: direction < 0 ? -1 : 1 }));
  const range = LINE_TREND_RANGES.includes(String(state.lineTrendRange || "").toLowerCase()) ? String(state.lineTrendRange).toLowerCase() : "day";
  if (range === "quarter") {
    date.setMonth(date.getMonth() + direction * 3);
  } else if (range === "month") {
    date.setMonth(date.getMonth() + direction);
  } else if (range === "week") {
    date.setDate(date.getDate() + direction * 7);
  } else {
    state.selectedDate = shiftWeekdayIsoDate(formatDateLocal(date), direction < 0 ? -1 : 1);
    document.querySelectorAll("[data-manager-date]").forEach((input) => {
      input.value = state.selectedDate;
    });
    saveState();
    renderAll();
    return;
  }
  state.selectedDate = normalizeWeekdayIsoDate(formatDateLocal(date), { direction: direction < 0 ? -1 : 1 });
  document.querySelectorAll("[data-manager-date]").forEach((input) => {
    input.value = state.selectedDate;
  });
  saveState();
  renderAll();
}

function setShiftToggleUI() {
  const selectedShift = state?.selectedShift || "Day";
  const crewShift = crewSettingsShiftForLine(state);
  document.querySelectorAll(".shift-option[data-shift], .shift-option[data-day-shift]").forEach((btn) => {
    const active = Boolean(state) && (btn.dataset.shift || btn.dataset.dayShift) === selectedShift;
    btn.classList.toggle("active", active);
    btn.setAttribute("aria-pressed", String(active));
  });
  document.querySelectorAll(".shift-option[data-crew-shift]").forEach((btn) => {
    const active = Boolean(state) && btn.dataset.crewShift === crewShift;
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
  const managerLogDateInputs = [
    shiftForm?.querySelector('[name="date"]'),
    runForm?.querySelector('[name="date"]'),
    downtimeForm?.querySelector('[name="date"]')
  ].filter(Boolean);

  managerLogDateInputs.forEach((input) => {
    input.addEventListener("change", () => {
      if (!input.value) return;
      input.value = normalizeWeekdayIsoDate(input.value, { direction: -1 });
    });
  });

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
    try {
      const data = Object.fromEntries(new FormData(form).entries());
      await createManagerShiftLogEntry(data);
      form.reset();
      renderAll();
    } catch (error) {
      alert(`Could not save shift log.\n${error?.message || "Please try again."}`);
    }
  });

  runForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    try {
      const data = Object.fromEntries(new FormData(form).entries());
      await createManagerRunLogEntry(data, {
        runCrewingPatternInput,
        runCrewingPatternSummary,
        selectedShift: state.selectedShift || "Day"
      });
      form.reset();
      setRunCrewingPatternField(runCrewingPatternInput, runCrewingPatternSummary, state, state.selectedShift || "Day", {}, { fallbackToIdeal: false });
      renderAll();
    } catch (error) {
      alert(`Could not save production run.\n${error?.message || "Please try again."}`);
    }
  });

  downtimeForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    try {
      const data = Object.fromEntries(new FormData(form).entries());
      await createManagerDowntimeLogEntry(data);
      form.reset();
      refreshManagerDowntimeDetailOptions();
      renderAll();
    } catch (error) {
      alert(`Could not save downtime log.\n${error?.message || "Please try again."}`);
    }
  });

  const inlineValue = (rowNode, field) => String(rowNode?.querySelector(`[data-inline-field="${field}"]`)?.value || "").trim();

  const saveManagerLogInlineEdit = async (type, logId, rowNode) => {
    if (!state?.id || !logId || !rowNode) return;
    const isInlineEditing = isManagerLogInlineEditRow(state.id, type, logId);
    const hasInlineFields = Boolean(rowNode.querySelector("[data-inline-field]"));
    if (!isInlineEditing && !hasInlineFields) return;

    if (type === "shift") {
      const row = (state.shiftRows || []).find((item) => String(item?.id || "") === String(logId || ""));
      if (!row) {
        alert("Shift row could not be found. Refresh and try again.");
        return;
      }
      const previousDate = String(row.date || "").trim();
      const previousShift = String(row.shift || "").trim();
      const date = inlineValue(rowNode, "date");
      const shift = inlineValue(rowNode, "shift");
      const startTime = inlineValue(rowNode, "startTime");
      const finishTime = inlineValue(rowNode, "finishTime");
      const notes = inlineValue(rowNode, "notes");
      if (!rowIsValidDateShift(date, shift)) {
        alert("Date/shift are invalid. Weekend dates are excluded.");
        return;
      }
      if (!strictTimeValid(startTime) || !strictTimeValid(finishTime)) {
        alert("Times must be HH:MM (24h).");
        return;
      }
      const payload = { date, shift, startTime, finishTime, notes };
      const canPatchServer = UUID_RE.test(String(logId || ""));
      try {
        if (canPatchServer) {
          const saved = await patchManagerShiftLog(logId, payload);
          Object.assign(row, payload, {
            submittedAt: saved?.submittedAt || nowIso()
          });
        } else {
          Object.assign(row, payload, {
            submittedAt: nowIso()
          });
        }
        const linkedBreakRows = breakRowsForShift(state, { id: row.id, date: previousDate, shift: previousShift });
        linkedBreakRows.forEach((breakRow) => {
          breakRow.date = date;
          breakRow.shift = shift;
          if (!breakRow.shiftLogId) breakRow.shiftLogId = row.id;
        });
        clearManagerLogInlineEdit();
        addAudit(state, "MANAGER_SHIFT_EDIT", `Manager edited shift row for ${row.date} (${row.shift})`);
        saveState();
        renderAll();
      } catch (error) {
        saveState();
        renderAll();
        alert(`Could not update shift row.\n${error?.message || "Please try again."}`);
      }
      return;
    }

    if (type === "run") {
      const row = (state.runRows || []).find((item) => String(item?.id || "") === String(logId || ""));
      if (!row) {
        alert("Run row could not be found. Refresh and try again.");
        return;
      }
      const date = inlineValue(rowNode, "date");
      const product = catalogProductCanonicalName(String(inlineValue(rowNode, "product") || "").trim());
      const productionStartTime = inlineValue(rowNode, "productionStartTime");
      const finishTime = inlineValue(rowNode, "finishTime");
      const unitsProduced = Math.max(0, num(inlineValue(rowNode, "unitsProduced")));
      const notes = inlineValue(rowNode, "notes");
      if (!isOperationalDate(date)) {
        alert("Date is invalid. Weekend dates are excluded.");
        return;
      }
      if (!product) {
        alert("Product is required.");
        return;
      }
      const productChanged = product !== catalogProductCanonicalName(String(row.product || "").trim());
      if (productChanged && !isCatalogProductName(product)) {
        alert("Product must be selected from Manage Products.");
        return;
      }
      if (!strictTimeValid(productionStartTime) || !optionalStrictTimeValid(finishTime)) {
        alert("Start must be HH:MM (24h). Finish must be HH:MM (24h) or empty.");
        return;
      }
      const runPatternRaw = inlineValue(rowNode, "runCrewingPattern");
      const patternShift = preferredTimedLogShift(state, date, productionStartTime, finishTime, state.selectedShift || "Day");
      const runPatternFromEdit = normalizeRunCrewingPattern(runPatternRaw, state, patternShift, { fallbackToIdeal: false });
      const existingRunPattern = normalizeRunCrewingPattern(row.runCrewingPattern, state, patternShift, { fallbackToIdeal: false });
      const runCrewingPattern = Object.keys(runPatternFromEdit).length ? runPatternFromEdit : existingRunPattern;
      const payload = {
        date,
        product,
        setUpStartTime: "",
        productionStartTime,
        finishTime,
        unitsProduced,
        notes,
        runCrewingPattern
      };
      const canPatchServer = UUID_RE.test(String(logId || ""));
      try {
        if (canPatchServer) {
          const saved = await patchManagerRunLog(logId, payload);
          Object.assign(row, {
            date,
            assignedShift: "",
            product,
            setUpStartTime: "",
            productionStartTime,
            finishTime,
            unitsProduced,
            notes,
            shift: "",
            runCrewingPattern: normalizeRunCrewingPattern(saved?.runCrewingPattern || runCrewingPattern, state, patternShift, { fallbackToIdeal: false }),
            submittedAt: saved?.submittedAt || nowIso()
          });
        } else {
          Object.assign(row, {
            date,
            assignedShift: "",
            product,
            setUpStartTime: "",
            productionStartTime,
            finishTime,
            unitsProduced,
            notes,
            shift: "",
            runCrewingPattern,
            submittedAt: nowIso()
          });
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
      const row = (state.downtimeRows || []).find((item) => String(item?.id || "") === String(logId || ""));
      if (!row) {
        alert("Downtime row could not be found. Refresh and try again.");
        return;
      }
      const date = inlineValue(rowNode, "date");
      const downtimeStart = inlineValue(rowNode, "downtimeStart");
      const downtimeFinish = inlineValue(rowNode, "downtimeFinish");
      const reasonCategory = inlineValue(rowNode, "reasonCategory");
      let reasonDetail = inlineValue(rowNode, "reasonDetail");
      const reasonNote = inlineValue(rowNode, "reasonNote");
      const notes = inlineValue(rowNode, "notes");
      let equipment = inlineValue(rowNode, "equipment");
      if (!isOperationalDate(date)) {
        alert("Date is invalid. Weekend dates are excluded.");
        return;
      }
      if (!strictTimeValid(downtimeStart) || !strictTimeValid(downtimeFinish)) {
        alert("Times must be HH:MM (24h).");
        return;
      }
      if (!reasonCategory) {
        alert("Reason group is required.");
        return;
      }
      if (reasonCategory === "Equipment") {
        reasonDetail = reasonDetail || equipment;
        equipment = reasonDetail || equipment;
        if (!equipment) {
          alert("Select an equipment stage.");
          return;
        }
      } else {
        equipment = "";
      }
      if (!reasonDetail) {
        alert("Reason detail is required.");
        return;
      }
      const reason = buildDowntimeReasonText(state, reasonCategory, reasonDetail, reasonNote);
      const updatedValues = {
        date,
        assignedShift: "",
        downtimeStart,
        downtimeFinish,
        equipment,
        reason,
        reasonCategory,
        reasonDetail,
        reasonNote,
        notes
      };
      const payload = {
        lineId: state.id,
        date,
        ...updatedValues
      };
      const canPatchServer = UUID_RE.test(String(logId || ""));
      try {
        if (canPatchServer) {
          const saved = await patchManagerDowntimeLog(logId, payload);
          Object.assign(row, updatedValues, {
            assignedShift: "",
            shift: "",
            submittedAt: saved?.submittedAt || nowIso()
          });
        } else {
          Object.assign(row, updatedValues, {
            assignedShift: "",
            shift: "",
            submittedAt: nowIso()
          });
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

  const managerDefaultFinishTime = (startTime) => {
    const nowCandidate = nowTimeHHMM();
    const startMins = parseTimeToMinutes(startTime);
    const nowMins = parseTimeToMinutes(nowCandidate);
    if (Number.isFinite(startMins) && Number.isFinite(nowMins) && startMins === nowMins) {
      const minsPerDay = 24 * 60;
      const normalized = ((Math.floor(num(startMins + 1)) % minsPerDay) + minsPerDay) % minsPerDay;
      const hh = String(Math.floor(normalized / 60)).padStart(2, "0");
      const mm = String(normalized % 60).padStart(2, "0");
      return `${hh}:${mm}`;
    }
    return nowCandidate;
  };

  const clearManagerInlineEditIfTarget = (type, logId) => {
    if (isManagerLogInlineEditRow(state.id, type, logId)) clearManagerLogInlineEdit();
  };

  const persistManagerFinalisedShift = async (row, logId, payload) => {
    const safeLogId = String(logId || "").trim();
    if (UUID_RE.test(safeLogId)) {
      const saved = await patchManagerShiftLog(safeLogId, payload);
      return { saved, persistedLogId: safeLogId };
    }
    const saved = await syncManagerShiftLog(payload);
    const persistedLogId = String(saved?.id || "").trim();
    if (!UUID_RE.test(persistedLogId)) {
      throw new Error("Shift finalise did not return a backend log id.");
    }
    const previousLogId = String(row?.id || "");
    if (previousLogId && previousLogId !== persistedLogId) {
      (state.breakRows || []).forEach((breakRow) => {
        if (String(breakRow?.shiftLogId || "") === previousLogId) {
          breakRow.shiftLogId = persistedLogId;
        }
      });
    }
    return { saved, persistedLogId };
  };

  const finaliseManagerShiftLogById = async (logId) => {
    const row = (state.shiftRows || []).find((item) => String(item?.id || "") === String(logId || ""));
    if (!row) {
      alert("Shift log could not be found.");
      return;
    }
    if (!isPendingShiftLogRow(row)) {
      alert("This shift is already finalised.");
      return;
    }
    const startTime = String(row.startTime || "").trim();
    if (!strictTimeValid(startTime)) {
      alert("Shift start time is invalid. Edit the row and set a valid start time first.");
      return;
    }
    const finishPrompt = window.prompt("Finish time (HH:MM)", managerDefaultFinishTime(startTime));
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
    const payload = {
      lineId: state.id,
      date: row.date,
      shift: row.shift || inferShiftForLog(state, row.date, startTime, state.selectedShift || "Day"),
      startTime,
      finishTime,
      notes: String(row.notes || "")
    };
    try {
      const { lineId: _ignoredLineId, ...rowPayload } = payload;
      const { saved, persistedLogId } = await persistManagerFinalisedShift(row, logId, payload);
      Object.assign(row, rowPayload, {
        id: persistedLogId || row.id,
        submittedAt: saved?.submittedAt || nowIso()
      });
      clearManagerInlineEditIfTarget("shift", logId);
      addAudit(state, "MANAGER_SHIFT_COMPLETE", `Manager finalised shift row for ${row.date} (${row.shift})`);
      saveState();
      renderAll();
    } catch (error) {
      alert(`Could not finalise shift row.\n${error?.message || "Please try again."}`);
    }
  };

  const finaliseManagerRunLogById = async (logId) => {
    const row = (state.runRows || []).find((item) => String(item?.id || "") === String(logId || ""));
    if (!row) {
      alert("Run log could not be found.");
      return;
    }
    if (!isPendingRunLogRow(row)) {
      alert("This production run is already finalised.");
      return;
    }
    const productionStartTime = String(row.productionStartTime || "").trim();
    if (!strictTimeValid(productionStartTime)) {
      alert("Run start time is invalid. Edit the row and set a valid start time first.");
      return;
    }
    const finishPrompt = window.prompt("Finish time (HH:MM)", managerDefaultFinishTime(productionStartTime));
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
    const patternShift = preferredTimedLogShift(state, row.date, productionStartTime, finishTime, state.selectedShift || "Day");
    const runCrewingPattern = normalizeRunCrewingPattern(row.runCrewingPattern, state, patternShift, { fallbackToIdeal: false });
    const payload = {
      lineId: state.id,
      date: row.date,
      product: row.product || "Run",
      setUpStartTime: "",
      productionStartTime,
      finishTime,
      unitsProduced,
      notes: String(row.notes || ""),
      runCrewingPattern
    };
    const canPatchServer = UUID_RE.test(String(logId || ""));
    try {
      if (canPatchServer) {
        const saved = await patchManagerRunLog(logId, payload);
        Object.assign(row, payload, {
          assignedShift: "",
          shift: "",
          runCrewingPattern: normalizeRunCrewingPattern(saved?.runCrewingPattern || runCrewingPattern, state, patternShift, { fallbackToIdeal: false }),
          submittedAt: saved?.submittedAt || nowIso()
        });
      } else {
        Object.assign(row, payload, {
          assignedShift: "",
          shift: "",
          runCrewingPattern,
          submittedAt: nowIso()
        });
      }
      clearManagerInlineEditIfTarget("run", logId);
      addAudit(state, "MANAGER_RUN_COMPLETE", `Manager finalised run ${row.product || "Run"}`);
      saveState();
      renderAll();
    } catch (error) {
      alert(`Could not finalise production run.\n${error?.message || "Please try again."}`);
    }
  };

  const finaliseManagerDowntimeLogById = async (logId) => {
    const row = (state.downtimeRows || []).find((item) => String(item?.id || "") === String(logId || ""));
    if (!row) {
      alert("Downtime log could not be found.");
      return;
    }
    if (!isPendingDowntimeLogRow(row)) {
      alert("This downtime log is already finalised.");
      return;
    }
    const downtimeStart = String(row.downtimeStart || "").trim();
    if (!strictTimeValid(downtimeStart)) {
      alert("Downtime start time is invalid. Edit the row and set a valid start time first.");
      return;
    }
    const finishPrompt = window.prompt("Downtime finish time (HH:MM)", managerDefaultFinishTime(downtimeStart));
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
    const reasonCategory = String(row.reasonCategory || parsedReason.reasonCategory || "").trim();
    const reasonDetail = String(row.reasonDetail || parsedReason.reasonDetail || "").trim();
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
    const updatedValues = {
      date: row.date,
      assignedShift: "",
      downtimeStart,
      downtimeFinish,
      equipment,
      reasonCategory,
      reasonDetail,
      reasonNote,
      notes: String(row.notes || ""),
      reason: buildDowntimeReasonText(state, reasonCategory, reasonDetail, reasonNote)
    };
    const payload = {
      lineId: state.id,
      date: row.date,
      ...updatedValues
    };
    const canPatchServer = UUID_RE.test(String(logId || ""));
    try {
      if (canPatchServer) {
        const saved = await patchManagerDowntimeLog(logId, payload);
        Object.assign(row, updatedValues, {
          shift: "",
          submittedAt: saved?.submittedAt || nowIso()
        });
      } else {
        Object.assign(row, updatedValues, {
          shift: "",
          submittedAt: nowIso()
        });
      }
      clearManagerInlineEditIfTarget("downtime", logId);
      addAudit(
        state,
        "MANAGER_DOWNTIME_COMPLETE",
        `Manager finalised downtime row for ${row.date} (${stageNameById(equipment) || reasonCategory})`
      );
      saveState();
      renderAll();
    } catch (error) {
      alert(`Could not finalise downtime row.\n${error?.message || "Please try again."}`);
    }
  };

  const deleteManagerShiftLogById = async (logId) => {
    const row = (state.shiftRows || []).find((item) => String(item?.id || "") === String(logId || ""));
    if (!row) {
      alert("Shift log could not be found.");
      return;
    }
    if (!window.confirm("Delete this shift log entry? This cannot be undone.")) return;
    const canDeleteServer = UUID_RE.test(String(logId || ""));
    try {
      if (canDeleteServer) await deleteManagerShiftLog(logId);
      const existingBreakRows = Array.isArray(state.breakRows) ? state.breakRows : [];
      const hasShiftLinkedBreaks = existingBreakRows.some((breakRow) => String(breakRow?.shiftLogId || "") === String(row.id || ""));
      state.shiftRows = (state.shiftRows || []).filter((item) => String(item?.id || "") !== String(logId || ""));
      state.breakRows = existingBreakRows.filter((breakRow) => {
        const breakShiftLogId = String(breakRow?.shiftLogId || "");
        if (breakShiftLogId) return breakShiftLogId !== String(logId || "");
        if (hasShiftLinkedBreaks) return true;
        return !(String(breakRow?.date || "") === String(row.date || "") && String(breakRow?.shift || "") === String(row.shift || ""));
      });
      clearManagerInlineEditIfTarget("shift", logId);
      addAudit(state, "MANAGER_SHIFT_DELETE", `Manager deleted shift row for ${row.date} (${row.shift})`);
      saveState();
      renderAll();
    } catch (error) {
      alert(`Could not delete shift row.\n${error?.message || "Please try again."}`);
    }
  };

  const deleteManagerRunLogById = async (logId) => {
    const row = (state.runRows || []).find((item) => String(item?.id || "") === String(logId || ""));
    if (!row) {
      alert("Run log could not be found.");
      return;
    }
    if (!window.confirm("Delete this run log entry? This cannot be undone.")) return;
    const canDeleteServer = UUID_RE.test(String(logId || ""));
    try {
      if (canDeleteServer) await deleteManagerRunLog(logId);
      state.runRows = (state.runRows || []).filter((item) => String(item?.id || "") !== String(logId || ""));
      clearManagerInlineEditIfTarget("run", logId);
      addAudit(state, "MANAGER_RUN_DELETE", `Manager deleted run ${row.product || "Run"}`);
      saveState();
      renderAll();
    } catch (error) {
      alert(`Could not delete production run.\n${error?.message || "Please try again."}`);
    }
  };

  const deleteManagerDowntimeLogById = async (logId) => {
    const row = (state.downtimeRows || []).find((item) => String(item?.id || "") === String(logId || ""));
    if (!row) {
      alert("Downtime log could not be found.");
      return;
    }
    if (!window.confirm("Delete this downtime log entry? This cannot be undone.")) return;
    const canDeleteServer = UUID_RE.test(String(logId || ""));
    try {
      if (canDeleteServer) await deleteManagerDowntimeLog(logId);
      state.downtimeRows = (state.downtimeRows || []).filter((item) => String(item?.id || "") !== String(logId || ""));
      clearManagerInlineEditIfTarget("downtime", logId);
      addAudit(
        state,
        "MANAGER_DOWNTIME_DELETE",
        `Manager deleted downtime row for ${row.date} (${stageNameById(row.equipment) || lineDowntimeReasonCategory(row) || "Downtime"})`
      );
      saveState();
      renderAll();
    } catch (error) {
      alert(`Could not delete downtime row.\n${error?.message || "Please try again."}`);
    }
  };

  const handleLogTableEdit = async (event) => {
    const patternBtn = event.target.closest("[data-inline-run-pattern]");
    if (patternBtn) {
      const rowNode = patternBtn.closest("tr");
      if (!rowNode || !state) return;
      const patternInput = rowNode.querySelector('[data-inline-field="runCrewingPattern"]');
      const patternSummary = rowNode.querySelector("[data-inline-run-pattern-summary]");
      if (!patternInput) return;
      const date = inlineValue(rowNode, "date") || state.selectedDate || todayISO();
      const startTime = inlineValue(rowNode, "productionStartTime");
      const finishTime = inlineValue(rowNode, "finishTime") || startTime;
      const fallbackShift = String(patternBtn.getAttribute("data-pattern-shift") || state.selectedShift || "Day");
      const shift = preferredTimedLogShift(state, date, startTime, finishTime, fallbackShift);
      openRunCrewingPatternModal({
        line: state,
        shift,
        inputEl: patternInput,
        summaryEl: patternSummary
      });
      return;
    }
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
    if (action === "edit") {
      setManagerLogInlineEdit(state.id, type, logId);
      renderTrackingTables();
      return;
    }
    if (action === "finalise") {
      if (type === "shift") await finaliseManagerShiftLogById(logId);
      if (type === "run") await finaliseManagerRunLogById(logId);
      if (type === "downtime") await finaliseManagerDowntimeLogById(logId);
      return;
    }
    if (action === "delete") {
      if (type === "shift") await deleteManagerShiftLogById(logId);
      if (type === "run") await deleteManagerRunLogById(logId);
      if (type === "downtime") await deleteManagerDowntimeLogById(logId);
    }
  };

  const handleDowntimeInlineFieldChange = (event) => {
    const target = event.target;
    const field = String(target?.getAttribute?.("data-inline-field") || "");
    if (!["reasonCategory", "reasonDetail", "equipment"].includes(field)) return;
    const rowNode = target.closest("tr");
    if (!rowNode) return;
    const categoryNode = rowNode.querySelector('[data-inline-field="reasonCategory"]');
    const detailNode = rowNode.querySelector('[data-inline-field="reasonDetail"]');
    const equipmentNode = rowNode.querySelector('[data-inline-field="equipment"]');
    if (!categoryNode || !detailNode) return;
    const category = String(categoryNode.value || "");
    const detail = String(detailNode.value || "");
    setDowntimeDetailOptions(detailNode, state, category, detail);
    if (!equipmentNode) return;
    const isEquipment = category === "Equipment";
    equipmentNode.disabled = !isEquipment;
    if (!isEquipment) {
      equipmentNode.value = "";
      return;
    }
    if (field === "equipment") {
      if (equipmentNode.value) detailNode.value = equipmentNode.value;
    } else if (detailNode.value) {
      equipmentNode.value = detailNode.value;
    }
  };

  ["shiftTable", "runTable", "downtimeTable"].forEach((tableId) => {
    const table = document.getElementById(tableId);
    if (table) table.addEventListener("click", handleLogTableEdit);
  });
  const downtimeTable = document.getElementById("downtimeTable");
  if (downtimeTable) downtimeTable.addEventListener("change", handleDowntimeInlineFieldChange);
}

function bindDataControls() {
  document.getElementById("loadSampleData").addEventListener("click", () => {
    loadSampleDataIntoLine(state);
    addAudit(state, "LOAD_SAMPLE_DATA", "Sample shift/run/downtime rows loaded");
    saveState();
    renderAll();
  });

  const loadPermanentSampleDataBtn = document.getElementById("loadPermanentSampleData");
  if (loadPermanentSampleDataBtn) {
    loadPermanentSampleDataBtn.addEventListener("click", async () => {
      if (!state) return;
      const confirmed = window.confirm(
        `Are you sure you want to load permanent sample data for "${state.name}"?\n\nThis will replace all existing shift, break, run and downtime logs for this line.`
      );
      if (!confirmed) return;
      try {
        const session = await ensureManagerBackendSession();
        const backendLineId = UUID_RE.test(String(state.id)) ? state.id : await ensureBackendLineId(state.id, session);
        if (!backendLineId) throw new Error("Line is not synced to server.");
        await apiRequest(`/api/lines/${backendLineId}/load-sample-data`, {
          method: "POST",
          token: session.backendToken,
          body: { replaceExisting: true },
          timeoutMs: 180000
        });
        await refreshHostedState();
        renderAll();
      } catch (error) {
        alert(`Could not load permanent sample data.\n${error?.message || "Please try again."}`);
      }
    });
  }

  const portBreakLogsBtn = document.getElementById("portBreakLogsToDowntime");
  if (portBreakLogsBtn) {
    portBreakLogsBtn.addEventListener("click", async () => {
      if (appState.appMode !== "manager") {
        alert("Only managers can port shift break logs.");
        return;
      }
      if (!state) return;
      const sourceBreakRows = Array.isArray(state.breakRows) ? state.breakRows.slice() : [];
      if (!sourceBreakRows.length) {
        alert("No shift break logs found to port.");
        return;
      }

      const migrationRows = sourceBreakRows
        .map((breakRow) => ({ breakRow, payload: buildDowntimeBreakMigrationPayload(state, breakRow) }))
        .filter((entry) => Boolean(entry.payload));
      const skippedInvalid = Math.max(0, sourceBreakRows.length - migrationRows.length);

      if (!migrationRows.length) {
        alert("No valid shift break logs found to port.");
        return;
      }

      const confirmed = window.confirm(
        `Port ${migrationRows.length} shift break log${migrationRows.length === 1 ? "" : "s"} to downtime logs?\n\nThis will remove migrated shift break logs to avoid double-counting.`
      );
      if (!confirmed) return;

      portBreakLogsBtn.disabled = true;
      const createdDowntimeRows = [];
      const migratedBreakIds = new Set();
      const migratedBreakRefs = new Set();
      let migrated = 0;
      let failed = 0;
      let deletedHostedBreaks = 0;
      let rolledBackDowntimeCreates = 0;
      let backendReady = false;

      try {
        try {
          const session = await ensureManagerBackendSession();
          const backendLineId = UUID_RE.test(String(state.id || "").trim()) ? String(state.id || "").trim() : await ensureBackendLineId(state.id, session);
          backendReady = UUID_RE.test(String(backendLineId || "").trim());
        } catch (_error) {
          backendReady = false;
        }

        for (const entry of migrationRows) {
          const breakRow = entry.breakRow;
          const payload = entry.payload;
          if (!payload) {
            failed += 1;
            continue;
          }
          try {
            let savedDowntime = null;
            if (backendReady) {
              savedDowntime = await syncManagerDowntimeLog(payload);
              const breakId = String(breakRow?.id || "").trim();
              const shiftLogId = String(breakRow?.shiftLogId || "").trim();
              const canDeleteHostedBreak = UUID_RE.test(breakId) && UUID_RE.test(shiftLogId);
              if (canDeleteHostedBreak) {
                try {
                  await deleteManagerShiftBreak(shiftLogId, breakId);
                  deletedHostedBreaks += 1;
                } catch (deleteError) {
                  const savedDowntimeId = String(savedDowntime?.id || "").trim();
                  if (UUID_RE.test(savedDowntimeId)) {
                    try {
                      await deleteManagerDowntimeLog(savedDowntimeId);
                      rolledBackDowntimeCreates += 1;
                    } catch (_rollbackError) {
                      // Best effort rollback; keep original error for row failure reporting.
                    }
                  }
                  throw deleteError;
                }
              }
            }

            const nextDowntimeRow = downtimeBreakRowFromMigration(state, breakRow, payload, savedDowntime);
            if (nextDowntimeRow) {
              createdDowntimeRows.push(nextDowntimeRow);
              migrated += 1;
              migratedBreakRefs.add(breakRow);
              const breakId = String(breakRow?.id || "").trim();
              if (breakId) migratedBreakIds.add(breakId);
            } else {
              failed += 1;
            }
          } catch (error) {
            console.warn("Could not port break log to downtime:", error);
            failed += 1;
          }
        }

        if (createdDowntimeRows.length) {
          if (!Array.isArray(state.downtimeRows)) state.downtimeRows = [];
          createdDowntimeRows.forEach((row) => {
            const rowId = String(row?.id || "").trim();
            if (rowId) {
              const existingIndex = state.downtimeRows.findIndex((existing) => String(existing?.id || "").trim() === rowId);
              if (existingIndex >= 0) {
                state.downtimeRows[existingIndex] = { ...state.downtimeRows[existingIndex], ...row };
                return;
              }
            }
            state.downtimeRows.push(row);
          });

          state.breakRows = (state.breakRows || []).filter((breakRow) => {
            const breakId = String(breakRow?.id || "").trim();
            if (breakId && migratedBreakIds.has(breakId)) return false;
            return !migratedBreakRefs.has(breakRow);
          });

          addAudit(
            state,
            "PORT_BREAKS_TO_DOWNTIME",
            `Ported ${migrated} break log${migrated === 1 ? "" : "s"} to downtime Break logs${backendReady ? ` (${deletedHostedBreaks} hosted break row${deletedHostedBreaks === 1 ? "" : "s"} removed)` : ""}`
          );
          saveState();
          if (backendReady) {
            try {
              await refreshHostedState();
            } catch (refreshError) {
              console.warn("Could not refresh hosted state after break port:", refreshError);
            }
          }
          renderAll();
        }

        alert(
          `Break port complete.\nMigrated: ${migrated}\nFailed: ${failed}\nSkipped invalid: ${skippedInvalid}\n${backendReady ? `Hosted break rows removed: ${deletedHostedBreaks}\n` : ""}${rolledBackDowntimeCreates > 0 ? `Rolled back downtime creates: ${rolledBackDowntimeCreates}\n` : ""}Remaining shift break logs: ${Math.max(0, (state.breakRows || []).length)}`
        );
      } finally {
        portBreakLogsBtn.disabled = false;
      }
    });
  }

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
        Details: `Breaks logged: ${(data.breakRows || []).filter((br) => br.date === row.date && br.shift === row.shift).length}`,
        Notes: String(row.notes || "")
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
        Details: `Net rate ${Number(row.netRunRate || 0).toFixed(2)} u/min`,
        Notes: String(row.notes || "")
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
        Details: row.reason || "",
        Notes: String(row.notes || "")
      }))
    ];
    const columns = ["RecordType", "Date", "Shift", "Product", "Equipment", "Units", "DowntimeMins", "Start", "Finish", "Details", "Notes"];
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
      const importedShiftRows = Array.isArray(parsed.shiftRows) ? parsed.shiftRows : [];
      const importedBreakRows = Array.isArray(parsed.breakRows) ? parsed.breakRows : [];
      const importedRunRows = Array.isArray(parsed.runRows)
        ? parsed.runRows.map((row) => normalizeRunLogRow(row))
        : [];
      const importedDowntimeRows = Array.isArray(parsed.downtimeRows)
        ? parsed.downtimeRows.map((row) => normalizeDowntimeLogRow(row))
        : [];

      if (!importedShiftRows.length && !importedBreakRows.length && !importedRunRows.length && !importedDowntimeRows.length) {
        alert("No importable rows found in this JSON file.");
        return;
      }

      const trimLogNotes = (value) => String(value || "").trim().slice(0, BACKEND_LOG_NOTES_MAX_LENGTH);
      const clipDowntimeReason = (value) => String(value || "").trim().slice(0, BACKEND_DOWNTIME_REASON_MAX_LENGTH);
      const isIsoDate = (value) => /^\d{4}-\d{2}-\d{2}$/.test(String(value || "").trim());

      await ensureManagerBackendSession();
      const backendLineId = await ensureBackendLineId(state.id, managerBackendSession);
      if (!backendLineId) throw new Error("Line is not synced to server.");

      const createdShiftIdBySourceId = new Map();
      const createdShiftIdsByKey = new Map();
      const breakShiftIndexByKey = new Map();
      const counts = { shift: 0, break: 0, run: 0, downtime: 0 };
      const skipped = { shift: 0, break: 0, run: 0, downtime: 0 };

      for (const row of importedShiftRows) {
        try {
          const date = String(row?.date || "").trim();
          const shift = String(row?.shift || "").trim();
          const startTime = String(row?.startTime || "").trim();
          const finishTime = String(row?.finishTime || "").trim();
          if (!isIsoDate(date) || !SHIFT_OPTIONS.includes(shift) || !strictTimeValid(startTime) || !strictTimeValid(finishTime)) {
            skipped.shift += 1;
            continue;
          }
          const payload = {
            lineId: state.id,
            date,
            shift,
            startTime,
            finishTime,
            notes: trimLogNotes(row?.notes)
          };
          const saved = await syncManagerShiftLog(payload);
          const savedId = String(saved?.id || "").trim();
          if (!savedId || !UUID_RE.test(savedId)) {
            skipped.shift += 1;
            continue;
          }
          const sourceId = String(row?.id || "").trim();
          if (sourceId) createdShiftIdBySourceId.set(sourceId, savedId);
          const key = shiftKey(date, shift);
          if (!createdShiftIdsByKey.has(key)) createdShiftIdsByKey.set(key, []);
          createdShiftIdsByKey.get(key).push(savedId);
          counts.shift += 1;
        } catch (error) {
          console.warn("Skipping imported shift row due to error:", error);
          skipped.shift += 1;
        }
      }

      for (const row of importedBreakRows) {
        try {
          const breakStart = String(row?.breakStart || "").trim();
          if (!strictTimeValid(breakStart)) {
            skipped.break += 1;
            continue;
          }

          let targetShiftId = "";
          const sourceShiftLogId = String(row?.shiftLogId || "").trim();
          if (sourceShiftLogId && createdShiftIdBySourceId.has(sourceShiftLogId)) {
            targetShiftId = createdShiftIdBySourceId.get(sourceShiftLogId) || "";
          } else if (UUID_RE.test(sourceShiftLogId)) {
            targetShiftId = sourceShiftLogId;
          } else {
            const date = String(row?.date || "").trim();
            const shift = String(row?.shift || "").trim();
            const key = shiftKey(date, shift);
            const candidates = createdShiftIdsByKey.get(key) || [];
            const nextIndex = Math.max(0, Math.floor(num(breakShiftIndexByKey.get(key))));
            if (candidates[nextIndex]) {
              targetShiftId = candidates[nextIndex];
              breakShiftIndexByKey.set(key, nextIndex + 1);
            } else if (candidates.length) {
              targetShiftId = candidates[candidates.length - 1];
            }
          }

          if (!targetShiftId) {
            skipped.break += 1;
            continue;
          }

          const breakFinish = strictTimeValid(String(row?.breakFinish || "").trim())
            ? String(row.breakFinish).trim()
            : null;
          await startManagerShiftBreak(targetShiftId, breakStart, breakFinish);
          counts.break += 1;
        } catch (error) {
          console.warn("Skipping imported break row due to error:", error);
          skipped.break += 1;
        }
      }

      for (const row of importedRunRows) {
        try {
          const date = String(row?.date || "").trim();
          const product = String(row?.product || "").trim();
          const productionStartTime = String(row?.productionStartTime || "").trim();
          if (!isIsoDate(date) || !product || !strictTimeValid(productionStartTime)) {
            skipped.run += 1;
            continue;
          }
          const finishTime = strictTimeValid(String(row?.finishTime || "").trim()) ? String(row.finishTime).trim() : "";
          const setUpStartTime = strictTimeValid(String(row?.setUpStartTime || "").trim()) ? String(row.setUpStartTime).trim() : "";
          const explicitShift = String(row?.assignedShift || row?.shift || "").trim();
          const patternShift = preferredTimedLogShift(
            state,
            date,
            productionStartTime,
            finishTime || productionStartTime,
            normalizeLogAssignableShift(explicitShift) || state.selectedShift || "Day"
          );
          const runCrewingPattern = normalizeRunCrewingPattern(row?.runCrewingPattern, state, patternShift, { fallbackToIdeal: false });
          const payload = {
            lineId: state.id,
            date,
            product,
            setUpStartTime,
            productionStartTime,
            finishTime,
            unitsProduced: Math.max(0, num(row?.unitsProduced)),
            notes: trimLogNotes(row?.notes),
            runCrewingPattern
          };
          await syncManagerRunLog(payload);
          counts.run += 1;
        } catch (error) {
          console.warn("Skipping imported run row due to error:", error);
          skipped.run += 1;
        }
      }

      for (const row of importedDowntimeRows) {
        try {
          const date = String(row?.date || "").trim();
          const downtimeStart = String(row?.downtimeStart || "").trim();
          const downtimeFinish = String(row?.downtimeFinish || "").trim();
          if (!isIsoDate(date) || !strictTimeValid(downtimeStart) || !strictTimeValid(downtimeFinish)) {
            skipped.downtime += 1;
            continue;
          }
          const parsedReason = parseDowntimeReasonParts(row?.reason, row?.equipment);
          const reasonCategory = String(row?.reasonCategory || parsedReason.reasonCategory || "").trim();
          const reasonDetail = String(row?.reasonDetail || parsedReason.reasonDetail || "").trim();
          const reasonNote = String((row?.reasonNote ?? parsedReason.reasonNote) || "").trim();
          const reasonText = clipDowntimeReason(
            String(row?.reason || "").trim() || buildDowntimeReasonText(state, reasonCategory, reasonDetail, reasonNote) || "Imported downtime"
          );
          if (!reasonText) {
            skipped.downtime += 1;
            continue;
          }
          const equipment = String(row?.equipment || "").trim() || (reasonCategory === "Equipment" ? reasonDetail : "");
          const payload = {
            lineId: state.id,
            date,
            downtimeStart,
            downtimeFinish,
            equipment,
            reason: reasonText,
            notes: trimLogNotes(row?.notes),
            excludeFromCalculation: normalizeDowntimeCalculationFlag(row?.excludeFromCalculation)
          };
          await syncManagerDowntimeLog(payload);
          counts.downtime += 1;
        } catch (error) {
          console.warn("Skipping imported downtime row due to error:", error);
          skipped.downtime += 1;
        }
      }

      await refreshHostedState();
      addAudit(
        state,
        "IMPORT_JSON",
        `Line JSON imported to DB (${counts.shift} shifts, ${counts.break} breaks, ${counts.run} runs, ${counts.downtime} downtime; skipped ${skipped.shift + skipped.break + skipped.run + skipped.downtime})`
      );
      saveState();
      renderAll();
      alert(
        `Import complete.\nSaved to DB: ${counts.shift} shifts, ${counts.break} breaks, ${counts.run} runs, ${counts.downtime} downtime.\nSkipped: ${skipped.shift + skipped.break + skipped.run + skipped.downtime}.`
      );
    } catch (error) {
      alert(`Could not import JSON.\n${error?.message || "Invalid JSON file."}`);
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
  const crewShift = crewSettingsShiftForLine(state);
  if (state.crewSettingsShift !== crewShift) state.crewSettingsShift = crewShift;
  shiftNote.textContent = `Crew values for ${crewShift} shift.`;
  const activeCrew = state.crewsByShift[crewShift] || defaultStageCrew(getStages());

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
      if (!state.crewsByShift[crewShift]) state.crewsByShift[crewShift] = defaultStageCrew(getStages());
      if (!state.crewsByShift[crewShift][stage.id]) state.crewsByShift[crewShift][stage.id] = {};
      state.crewsByShift[crewShift][stage.id].crew = num(input.value);
      markLineSettingsDirty(state.id);
      saveState();
      renderVisualiser();
      renderTrendModalContent();
    });

    row.append(name, input);
    form.append(row);
  });
  syncLineSettingsSaveUI();
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
      markLineSettingsDirty(state.id);
      saveState();
      renderVisualiser();
      renderTrendModalContent();
    });

    row.append(name, input);
    form.append(row);
  });
  syncLineSettingsSaveUI();
}

function renderTable(tableId, columns, rows, fieldMap, options = {}) {
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
  const groupByField = String(options?.groupByField || "").trim();
  const hideGroupedFieldCells = Boolean(options?.hideGroupedFieldCells);
  const groupLabelFormatter = typeof options?.groupLabelFormatter === "function"
    ? options.groupLabelFormatter
    : (value) => String(value ?? "");

  const table = document.getElementById(tableId);
  const header = `<thead><tr>${columns.map((col) => `<th>${col}</th>`).join("")}</tr></thead>`;
  const renderRow = (row) => {
    const rowClass = String(row?.__rowClass || "").trim();
    const rowClassAttr = rowClass ? ` class="${htmlEscape(rowClass)}"` : "";
    const htmlFields = new Set(Array.isArray(row?.__htmlFields) ? row.__htmlFields.map((field) => String(field || "")) : []);
    const cells = columns
      .map((col) => {
        const field = String(fieldMap[col] || "");
        const value = row?.[field];
        if (field && htmlFields.has(field)) return `<td>${String(value ?? "")}</td>`;
        if (groupByField && hideGroupedFieldCells && field === groupByField) return `<td></td>`;
        return `<td>${htmlEscape(formatTableCellValue(value))}</td>`;
      })
      .join("");
    return `<tr${rowClassAttr}>${cells}</tr>`;
  };
  let lastGroupValue = "";
  const body = rows
    .map((row, index) => {
      const rowHtml = renderRow(row);
      if (!groupByField) return rowHtml;
      const groupValue = String(row?.[groupByField] ?? "").trim();
      const needsDivider = index === 0 || groupValue !== lastGroupValue;
      lastGroupValue = groupValue;
      if (!needsDivider) return rowHtml;
      const groupLabel = htmlEscape(groupLabelFormatter(groupValue));
      return `<tr class="entry-date-divider"><td colspan="${columns.length}">${groupLabel}</td></tr>${rowHtml}`;
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

function normalizeDayVisualiserKeyStageId(line, stageId) {
  const safeStageId = String(stageId || "").trim();
  if (!safeStageId) return "";
  const stages = line?.stages?.length ? line.stages : [];
  return stages.some((stage) => String(stage?.id || "") === safeStageId) ? safeStageId : "";
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

function formatMinutesToClockLabel(totalMinutes) {
  const minutesInDay = 24 * 60;
  const safe = ((Math.floor(num(totalMinutes)) % minutesInDay) + minutesInDay) % minutesInDay;
  const hour = Math.floor(safe / 60);
  const minute = safe % 60;
  return formatTime12h(`${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`);
}

function dayVizBlockTypeLabel(type) {
  const key = String(type || "").toLowerCase();
  if (key === "shift") return "Shift";
  if (key === "break") return "Break";
  if (key === "run-main") return "Production";
  if (key === "run-setup") return "Setup";
  if (key === "downtime") return "Downtime";
  return "Time block";
}

function dayVizBlockDetailsFromElement(blockEl) {
  if (!(blockEl instanceof Element)) return null;
  const detailId = String(blockEl.getAttribute("data-day-viz-block-id") || "");
  if (!detailId) return null;
  const canvas = blockEl.closest(".day-viz-canvas");
  if (!canvas || typeof canvas !== "object") return null;
  const detailsById = canvas.__dayVizBlockDetails;
  if (!detailsById || typeof detailsById !== "object") return null;
  return detailsById[detailId] || null;
}

function dayVizBlockDetailRows(detail = {}) {
  const rows = [
    { label: "Date", value: detail.date || "-" },
    { label: "Shift Filter", value: detail.shiftLabel || "All shifts" },
    { label: "Lane", value: detail.laneLabel || "-" },
    { label: "Type", value: detail.typeLabel || "Time block" },
    { label: "Time", value: detail.timeRange || "-" },
    { label: "Duration", value: `${formatNum(Math.max(0, num(detail.durationMins)), 1)} min` }
  ];
  if (String(detail.typeKey || "").toLowerCase() === "downtime") {
    rows.push({
      label: "Calculation Status",
      value: detail.excludeFromCalculation ? "Excluded from run rate and utilisation" : "Included in run rate and utilisation"
    });
  }
  const subtitle = String(detail.subtitle || "").trim();
  if (subtitle) rows.push({ label: "Logged Details", value: subtitle });
  return rows
    .map(
      (row) => `
        <div class="day-viz-block-detail-row">
          <span>${htmlEscape(row.label)}</span>
          <strong>${htmlEscape(row.value)}</strong>
        </div>
      `
    )
    .join("");
}

function dayVizBlockModalCanToggleDowntime(detail = dayVizBlockModalState) {
  return Boolean(
    detail
    && String(detail.typeKey || "").toLowerCase() === "downtime"
    && String(detail.rootId || "") === "dayVisualiserCanvas"
    && appState.appMode === "manager"
    && state
    && String(detail.lineId || "") === String(state.id || "")
    && String(detail.logId || "").trim()
  );
}

function currentDayVizDowntimeRow(detail = dayVizBlockModalState) {
  if (!dayVizBlockModalCanToggleDowntime(detail)) return null;
  const logId = String(detail?.logId || "").trim();
  return (state?.downtimeRows || []).find((row) => String(row?.id || "").trim() === logId) || null;
}

function renderDayVizBlockModal(detail) {
  const overlay = document.getElementById("dayVizBlockModal");
  const titleNode = document.getElementById("dayVizBlockModalTitle");
  const metaNode = document.getElementById("dayVizBlockModalMeta");
  const bodyNode = document.getElementById("dayVizBlockModalBody");
  const actionsNode = document.getElementById("dayVizBlockModalActions");
  if (!overlay || !titleNode || !metaNode || !bodyNode || !detail) return;
  dayVizBlockModalState = { ...detail };
  titleNode.textContent = detail.title || "Time Block";
  metaNode.textContent = `${detail.typeLabel || "Time block"} | ${detail.laneLabel || "Schedule"}`;
  bodyNode.innerHTML = dayVizBlockDetailRows(detail);
  if (actionsNode) {
    actionsNode.innerHTML = dayVizBlockModalCanToggleDowntime(detail)
      ? `
        <p class="day-viz-block-modal-action-note">Keeps the downtime visible, but removes it from run rate and utilisation calculations.</p>
        <button id="toggleDayVizExcludeFromCalculation" type="button"${dayVizBlockModalBusy ? " disabled" : ""}>
          ${detail.excludeFromCalculation ? "Include in calculation" : "Exclude from calculation"}
        </button>
      `
      : "";
  }
  overlay.classList.add("open");
  overlay.setAttribute("aria-hidden", "false");
}

function openDayVizBlockModal(detail) {
  renderDayVizBlockModal(detail);
}

function closeDayVizBlockModal() {
  const overlay = document.getElementById("dayVizBlockModal");
  const actionsNode = document.getElementById("dayVizBlockModalActions");
  if (!overlay) return;
  dayVizBlockModalState = null;
  dayVizBlockModalBusy = false;
  if (actionsNode) actionsNode.innerHTML = "";
  overlay.classList.remove("open");
  overlay.setAttribute("aria-hidden", "true");
}

async function toggleDayVizDowntimeCalculationExclusion() {
  const detail = dayVizBlockModalState;
  const row = currentDayVizDowntimeRow(detail);
  if (!detail || !row) return;
  const nextExcludeFromCalculation = !isDowntimeExcludedFromCalculation(row);
  const parsedReason = parseDowntimeReasonParts(row.reason, row.equipment);
  const reasonCategory = String(row.reasonCategory || parsedReason.reasonCategory || "").trim();
  const reasonDetail = String(row.reasonDetail || parsedReason.reasonDetail || "").trim();
  const reasonNote = String((row.reasonNote ?? parsedReason.reasonNote) || "").trim();
  const payload = {
    lineId: state.id,
    date: row.date,
    downtimeStart: row.downtimeStart,
    downtimeFinish: row.downtimeFinish,
    equipment: String(row.equipment || "").trim(),
    reason: String(row.reason || buildDowntimeReasonText(state, reasonCategory, reasonDetail, reasonNote) || "Downtime").trim(),
    notes: String(row.notes || ""),
    excludeFromCalculation: nextExcludeFromCalculation
  };
  const activeLogId = String(row.id || "").trim();
  dayVizBlockModalBusy = true;
  renderDayVizBlockModal({
    ...detail,
    excludeFromCalculation: isDowntimeExcludedFromCalculation(row)
  });
  try {
    if (UUID_RE.test(activeLogId)) {
      const saved = await patchManagerDowntimeLog(activeLogId, payload);
      row.assignedShift = "";
      row.submittedAt = saved?.submittedAt || nowIso();
      row.excludeFromCalculation = normalizeDowntimeCalculationFlag(saved?.excludeFromCalculation ?? nextExcludeFromCalculation);
    } else {
      row.assignedShift = "";
      row.submittedAt = nowIso();
      row.excludeFromCalculation = nextExcludeFromCalculation;
    }
    addAudit(
      state,
      nextExcludeFromCalculation ? "MANAGER_DOWNTIME_EXCLUDE_CALC" : "MANAGER_DOWNTIME_INCLUDE_CALC",
      `Manager ${nextExcludeFromCalculation ? "excluded" : "included"} downtime from calculations for ${row.date} (${stageNameById(row.equipment) || lineDowntimeReasonCategory(row) || "Downtime"})`
    );
    saveState();
    renderAll();
  } catch (error) {
    alert(`Could not update downtime calculation status.\n${error?.message || "Please try again."}`);
  } finally {
    const isSameModal =
      dayVizBlockModalState
      && String(dayVizBlockModalState.logId || "").trim() === activeLogId;
    dayVizBlockModalBusy = false;
    if (isSameModal) {
      renderDayVizBlockModal({
        ...dayVizBlockModalState,
        excludeFromCalculation: isDowntimeExcludedFromCalculation(row)
      });
    }
  }
}

function bindDayVizBlockModal() {
  const overlay = document.getElementById("dayVizBlockModal");
  const closeBtn = document.getElementById("closeDayVizBlockModal");
  if (!overlay || !closeBtn) return;
  const tryOpenFromElement = (blockEl) => {
    const detail = dayVizBlockDetailsFromElement(blockEl);
    if (!detail) return false;
    openDayVizBlockModal(detail);
    return true;
  };

  closeBtn.addEventListener("click", closeDayVizBlockModal);
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) closeDayVizBlockModal();
  });

  document.addEventListener("click", (event) => {
    if (!(event.target instanceof Element)) return;
    if (event.target.closest("#toggleDayVizExcludeFromCalculation")) {
      toggleDayVizDowntimeCalculationExclusion();
      return;
    }
    const blockEl = event.target.closest(".day-viz-block[data-day-viz-block-id]");
    if (!blockEl) return;
    tryOpenFromElement(blockEl);
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && overlay.classList.contains("open")) {
      closeDayVizBlockModal();
      return;
    }
    if (event.key !== "Enter" && event.key !== " ") return;
    if (!(event.target instanceof Element)) return;
    const blockEl = event.target.closest(".day-viz-block[data-day-viz-block-id]");
    if (!blockEl) return;
    event.preventDefault();
    tryOpenFromElement(blockEl);
  });
}

function dayVizAddRecordTypeConfig(recordType) {
  const key = String(recordType || "").trim().toLowerCase();
  if (key === "shift") {
    return {
      key,
      laneLabel: "Shift",
      modalTitle: "Add Shift Record",
      submitLabel: "Add Shift Record",
      failureLabel: "shift log"
    };
  }
  if (key === "break") {
    return {
      key,
      laneLabel: "Break",
      modalTitle: "Add Break Record",
      submitLabel: "Add Break Record",
      failureLabel: "break log"
    };
  }
  if (key === "run") {
    return {
      key,
      laneLabel: "Production",
      modalTitle: "Add Production Record",
      submitLabel: "Add Production Record",
      failureLabel: "production run"
    };
  }
  if (key === "downtime") {
    return {
      key,
      laneLabel: "Downtime",
      modalTitle: "Add Downtime Record",
      submitLabel: "Add Downtime Record",
      failureLabel: "downtime log"
    };
  }
  return null;
}

function dayVizAddRecordModalLine() {
  const lineId = String(dayVizAddRecordModalState?.lineId || "").trim();
  if (!lineId) return state || null;
  return appState.lines?.[lineId] || (state?.id === lineId ? state : null);
}

function dayVizAddRecordDateLabel(isoDate) {
  return parseDateLocal(isoDate || todayISO()).toLocaleDateString(undefined, {
    weekday: "short",
    month: "short",
    day: "numeric",
    year: "numeric"
  });
}

function dayVizAddRecordSummaryHtml(line, selectedDate, selectedShift) {
  const viewLabel = selectedShift === "Full Day" ? "Full Day" : `${selectedShift} Shift`;
  return `
    <div class="day-viz-add-record-summary">
      <div class="day-viz-add-record-context-card">
        <span>Line</span>
        <strong>${htmlEscape(String(line?.name || "Production Line"))}</strong>
      </div>
      <div class="day-viz-add-record-context-card">
        <span>Date</span>
        <strong>${htmlEscape(dayVizAddRecordDateLabel(selectedDate))}</strong>
      </div>
      <div class="day-viz-add-record-context-card">
        <span>View</span>
        <strong>${htmlEscape(viewLabel)}</strong>
      </div>
    </div>
  `;
}

function dayVizAddRecordDefaultShift(selectedShift, { allowFullDay = false } = {}) {
  const safeShift = String(selectedShift || "").trim();
  if (allowFullDay && SHIFT_OPTIONS.includes(safeShift)) return safeShift;
  return fallbackShiftValue(safeShift);
}

function dayVizAddRecordPreferredBreakShiftId(choices, selectedShift = "") {
  const safeShift = String(selectedShift || "").trim();
  const exact = choices.find((choice) => choice.shift === safeShift);
  if (exact?.id) return exact.id;
  const fallback = choices.find((choice) => choice.shift === fallbackShiftValue(safeShift));
  if (fallback?.id) return fallback.id;
  return choices[0]?.id || "";
}

function closeDayVizAddRecordModal() {
  const overlay = document.getElementById("dayVizAddRecordModal");
  const fieldsNode = document.getElementById("dayVizAddRecordFields");
  const statusNode = document.getElementById("dayVizAddRecordStatus");
  if (fieldsNode) fieldsNode.innerHTML = "";
  if (statusNode) statusNode.textContent = "";
  if (overlay) {
    overlay.classList.remove("open");
    overlay.setAttribute("aria-hidden", "true");
  }
  dayVizAddRecordModalState = null;
}

function refreshDayVizAddRecordDowntimeDetailOptions() {
  const line = dayVizAddRecordModalLine();
  const categorySelect = document.getElementById("dayVizAddRecordReasonCategory");
  const detailSelect = document.getElementById("dayVizAddRecordReasonDetail");
  if (!line || !categorySelect || !detailSelect) return;
  setDowntimeDetailOptions(detailSelect, line, String(categorySelect.value || ""), detailSelect.value || "");
}

function openDayVizAddRecordRunCrewingModal() {
  const line = dayVizAddRecordModalLine();
  const form = document.getElementById("dayVizAddRecordForm");
  const patternInput = document.getElementById("dayVizAddRecordRunCrewingPattern");
  const patternSummary = document.getElementById("dayVizAddRecordRunCrewingSummary");
  if (!line || !form || !patternInput || !patternSummary) return;
  const date = String(form.querySelector('[name="date"]')?.value || dayVizAddRecordModalState?.date || line.selectedDate || todayISO());
  const startTime = String(form.querySelector('[name="productionStartTime"]')?.value || nowTimeHHMM());
  const finishTime = String(form.querySelector('[name="finishTime"]')?.value || startTime);
  const shift = preferredTimedLogShift(
    line,
    date,
    startTime,
    finishTime,
    fallbackShiftValue(dayVizAddRecordModalState?.selectedShift)
  );
  openRunCrewingPatternModal({
    line,
    shift,
    inputEl: patternInput,
    summaryEl: patternSummary
  });
}

function renderDayVizAddRecordModal() {
  const overlay = document.getElementById("dayVizAddRecordModal");
  const titleNode = document.getElementById("dayVizAddRecordTitle");
  const metaNode = document.getElementById("dayVizAddRecordMeta");
  const fieldsNode = document.getElementById("dayVizAddRecordFields");
  const statusNode = document.getElementById("dayVizAddRecordStatus");
  const submitBtn = document.getElementById("dayVizAddRecordSubmit");
  const form = document.getElementById("dayVizAddRecordForm");
  if (!overlay || !titleNode || !metaNode || !fieldsNode || !statusNode || !submitBtn || !form) return;

  const config = dayVizAddRecordTypeConfig(dayVizAddRecordModalState?.type);
  const line = dayVizAddRecordModalLine();
  if (!config || !line || appState.appMode !== "manager") {
    closeDayVizAddRecordModal();
    return;
  }

  const selectedDate = normalizeWeekdayIsoDate(dayVizAddRecordModalState?.date || line.selectedDate || todayISO(), { direction: -1 });
  const requestedShift = String(dayVizAddRecordModalState?.selectedShift || line.selectedShift || "Day").trim();
  const selectedShift = SHIFT_OPTIONS.includes(requestedShift) ? requestedShift : fallbackShiftValue(requestedShift);
  const breakChoices = config.key === "break" ? managerBreakShiftChoices(line, selectedDate) : [];
  const defaultBreakShiftId = dayVizAddRecordPreferredBreakShiftId(breakChoices, selectedShift);
  const productOptionsHtml = listProductCatalogOptions()
    .map((item) => `<option value="${htmlEscape(item.value)}">${htmlEscape(item.label || item.value)}</option>`)
    .join("");
  const reasonCategoryOptionsHtml = [
    `<option value="">Reason Group</option>`,
    ...ACTION_REASON_CATEGORIES.map((category) => `<option value="${htmlEscape(category)}">${htmlEscape(category)}</option>`)
  ].join("");
  const summaryHtml = dayVizAddRecordSummaryHtml(line, selectedDate, selectedShift);
  let fieldsHtml = "";
  let statusText = "";
  let disableSubmit = false;

  if (config.key === "shift") {
    const defaultShift = dayVizAddRecordDefaultShift(selectedShift, { allowFullDay: true });
    fieldsHtml = `
      <input type="hidden" name="date" value="${htmlEscape(selectedDate)}" />
      ${summaryHtml}
      <label>
        Shift
        <select name="shift" required>
          ${SHIFT_OPTIONS.map((shift) => `<option value="${htmlEscape(shift)}"${shift === defaultShift ? " selected" : ""}>${htmlEscape(shift)}</option>`).join("")}
        </select>
      </label>
      <label>
        Start Time
        <input name="startTime" type="time" step="60" required />
      </label>
      <label>
        Finish Time
        <input name="finishTime" type="time" step="60" required />
      </label>
      <label class="day-viz-add-record-span-2">
        Notes
        <input name="notes" maxlength="${BACKEND_LOG_NOTES_MAX_LENGTH}" placeholder="Notes (optional)" />
      </label>
    `;
  } else if (config.key === "break") {
    disableSubmit = !breakChoices.length;
    statusText = breakChoices.length ? "Breaks are attached to a saved shift record for this day." : `Add a shift record for ${selectedDate} before logging a break.`;
    fieldsHtml = `
      <input type="hidden" name="date" value="${htmlEscape(selectedDate)}" />
      ${summaryHtml}
      <p class="day-viz-add-record-note">Breaks are stored against an existing shift record, so the shift must already be logged for this date.</p>
      ${
        breakChoices.length
          ? `
            <label class="day-viz-add-record-span-2">
              Shift Record
              <select name="shiftLogId" required>
                ${breakChoices.map((choice) => `<option value="${htmlEscape(choice.id)}"${choice.id === defaultBreakShiftId ? " selected" : ""}>${htmlEscape(choice.label)}</option>`).join("")}
              </select>
            </label>
            <label>
              Break Start
              <input name="breakStart" type="time" step="60" required />
            </label>
            <label>
              Break Finish
              <input name="breakFinish" type="time" step="60" required />
            </label>
          `
          : `
            <div class="day-viz-add-record-empty">
              <strong>No saved shift records available for this day.</strong>
              <p>Log the shift first, then come back to add its break.</p>
            </div>
          `
      }
    `;
  } else if (config.key === "run") {
    fieldsHtml = `
      <input type="hidden" name="date" value="${htmlEscape(selectedDate)}" />
      <input id="dayVizAddRecordRunCrewingPattern" name="runCrewingPattern" type="hidden" />
      ${summaryHtml}
      <label class="day-viz-add-record-span-2">
        Product
        <input name="product" list="dayVizAddRecordProductList" autocomplete="off" placeholder="Select Product" required />
      </label>
      <datalist id="dayVizAddRecordProductList">${productOptionsHtml}</datalist>
      <label>
        Production Start
        <input name="productionStartTime" type="time" step="60" required />
      </label>
      <label>
        Finish Time
        <input name="finishTime" type="time" step="60" required />
      </label>
      <label>
        Units Produced
        <input name="unitsProduced" type="number" min="0" step="1" placeholder="0" />
      </label>
      <button id="dayVizAddRecordRunCrewingBtn" class="ghost-btn day-viz-add-record-crewing-btn" type="button">Set Crewing Pattern</button>
      <p id="dayVizAddRecordRunCrewingSummary" class="run-crewing-summary day-viz-add-record-crewing-summary">No crewing pattern set.</p>
      <label class="day-viz-add-record-span-2">
        Notes
        <input name="notes" maxlength="${BACKEND_LOG_NOTES_MAX_LENGTH}" placeholder="Notes (optional)" />
      </label>
    `;
  } else if (config.key === "downtime") {
    fieldsHtml = `
      <input type="hidden" name="date" value="${htmlEscape(selectedDate)}" />
      ${summaryHtml}
      <label>
        Downtime Start
        <input name="downtimeStart" type="time" step="60" required />
      </label>
      <label>
        Downtime Finish
        <input name="downtimeFinish" type="time" step="60" required />
      </label>
      <label>
        Reason Group
        <select id="dayVizAddRecordReasonCategory" name="reasonCategory" required>
          ${reasonCategoryOptionsHtml}
        </select>
      </label>
      <label>
        Reason Detail
        <select id="dayVizAddRecordReasonDetail" name="reasonDetail" required>
          <option value="">Select Reason</option>
        </select>
      </label>
      <label>
        Reason Notes
        <input name="reasonNote" maxlength="160" placeholder="Reason Notes (optional)" />
      </label>
      <label class="day-viz-add-record-span-2">
        Notes
        <input name="notes" maxlength="${BACKEND_LOG_NOTES_MAX_LENGTH}" placeholder="Notes (optional)" />
      </label>
    `;
  }

  dayVizAddRecordModalState = {
    ...dayVizAddRecordModalState,
    type: config.key,
    lineId: String(line.id || ""),
    date: selectedDate,
    selectedShift
  };
  titleNode.textContent = config.modalTitle;
  metaNode.textContent = `${line.name || "Production Line"} | ${selectedDate} | ${config.laneLabel}`;
  fieldsNode.innerHTML = fieldsHtml;
  statusNode.textContent = statusText;
  submitBtn.textContent = config.submitLabel;
  submitBtn.disabled = disableSubmit;
  overlay.classList.add("open");
  overlay.setAttribute("aria-hidden", "false");

  if (config.key === "downtime") {
    refreshDayVizAddRecordDowntimeDetailOptions();
  }

  if (config.key === "run") {
    const patternInput = document.getElementById("dayVizAddRecordRunCrewingPattern");
    const patternSummary = document.getElementById("dayVizAddRecordRunCrewingSummary");
    if (patternInput && patternSummary) {
      const startTime = String(form.querySelector('[name="productionStartTime"]')?.value || nowTimeHHMM());
      const finishTime = String(form.querySelector('[name="finishTime"]')?.value || startTime);
      const shift = preferredTimedLogShift(line, selectedDate, startTime, finishTime, fallbackShiftValue(selectedShift));
      setRunCrewingPatternField(patternInput, patternSummary, line, shift, patternInput.value || "", { fallbackToIdeal: false });
    }
  }

  const focusTarget = fieldsNode.querySelector('select:not(:disabled), input:not([type="hidden"]):not(:disabled), button:not(:disabled)');
  if (focusTarget instanceof HTMLElement) {
    window.setTimeout(() => focusTarget.focus(), 0);
  }
}

function openDayVizAddRecordModal(recordType, { lineId = state?.id, date = state?.selectedDate || todayISO(), selectedShift = state?.selectedShift || "Day" } = {}) {
  const config = dayVizAddRecordTypeConfig(recordType);
  const safeLineId = String(lineId || "").trim();
  if (!config || !safeLineId || appState.appMode !== "manager") return;
  dayVizAddRecordModalState = {
    type: config.key,
    lineId: safeLineId,
    date: normalizeWeekdayIsoDate(date || todayISO(), { direction: -1 }),
    selectedShift: SHIFT_OPTIONS.includes(String(selectedShift || "").trim()) ? String(selectedShift).trim() : fallbackShiftValue(selectedShift)
  };
  renderDayVizAddRecordModal();
}

async function submitDayVizAddRecordModal(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const config = dayVizAddRecordTypeConfig(dayVizAddRecordModalState?.type);
  const line = dayVizAddRecordModalLine();
  const submitBtn = document.getElementById("dayVizAddRecordSubmit");
  const statusNode = document.getElementById("dayVizAddRecordStatus");
  if (!form || !config || !line || !submitBtn || !statusNode) return;
  const data = Object.fromEntries(new FormData(form).entries());
  submitBtn.disabled = true;
  statusNode.textContent = "Saving...";
  try {
    if (config.key === "shift") {
      await createManagerShiftLogEntry(data, { line });
    } else if (config.key === "break") {
      await createManagerBreakLogEntry(data, { line });
    } else if (config.key === "run") {
      const patternInput = document.getElementById("dayVizAddRecordRunCrewingPattern");
      const patternSummary = document.getElementById("dayVizAddRecordRunCrewingSummary");
      await createManagerRunLogEntry(data, {
        line,
        runCrewingPatternInput: patternInput,
        runCrewingPatternSummary: patternSummary,
        selectedShift: dayVizAddRecordModalState?.selectedShift || line.selectedShift || "Day"
      });
    } else if (config.key === "downtime") {
      await createManagerDowntimeLogEntry(data, { line });
    }
    closeDayVizAddRecordModal();
    renderAll();
  } catch (error) {
    statusNode.textContent = String(error?.message || "").trim();
    submitBtn.disabled = false;
    alert(`Could not save ${config.failureLabel}.\n${error?.message || "Please try again."}`);
  }
}

function bindDayVizAddRecordModal() {
  const overlay = document.getElementById("dayVizAddRecordModal");
  const closeBtn = document.getElementById("closeDayVizAddRecordModal");
  const form = document.getElementById("dayVizAddRecordForm");
  if (!overlay || !closeBtn || !form) return;

  closeBtn.addEventListener("click", closeDayVizAddRecordModal);
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) closeDayVizAddRecordModal();
  });
  form.addEventListener("submit", submitDayVizAddRecordModal);
  form.addEventListener("change", (event) => {
    if (!(event.target instanceof Element)) return;
    if (event.target.id === "dayVizAddRecordReasonCategory") {
      refreshDayVizAddRecordDowntimeDetailOptions();
    }
  });

  document.addEventListener("click", (event) => {
    if (!(event.target instanceof Element)) return;
    const addBtn = event.target.closest("[data-day-viz-add-record]");
    if (addBtn) {
      event.preventDefault();
      openDayVizAddRecordModal(String(addBtn.getAttribute("data-day-viz-add-record") || ""), {
        lineId: addBtn.getAttribute("data-day-viz-line-id") || state?.id,
        date: addBtn.getAttribute("data-day-viz-date") || state?.selectedDate || todayISO(),
        selectedShift: addBtn.getAttribute("data-day-viz-shift") || state?.selectedShift || "Day"
      });
      return;
    }
    if (event.target.closest("#dayVizAddRecordRunCrewingBtn")) {
      event.preventDefault();
      openDayVizAddRecordRunCrewingModal();
    }
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && overlay.classList.contains("open")) {
      closeDayVizAddRecordModal();
    }
  });
}

function activePasswordResetSession() {
  if (appState.appMode === "supervisor") {
    return appState.supervisorSession?.backendToken ? appState.supervisorSession : null;
  }
  return managerBackendSession?.backendToken ? managerBackendSession : null;
}

function setPasswordResetStatus(message = "", tone = "") {
  const statusNode = document.getElementById("passwordResetStatus");
  if (!statusNode) return;
  statusNode.textContent = String(message || "").trim();
  statusNode.classList.toggle("is-success", tone === "success");
  statusNode.classList.toggle("is-error", tone === "error");
  if (!message) statusNode.classList.remove("is-success", "is-error");
}

function closePasswordResetModal() {
  const overlay = document.getElementById("passwordResetModal");
  const form = document.getElementById("passwordResetForm");
  const submitBtn = document.getElementById("passwordResetSubmit");
  passwordResetModalBusy = false;
  if (form) form.reset();
  if (submitBtn) {
    submitBtn.disabled = false;
    submitBtn.textContent = "Update Password";
  }
  setPasswordResetStatus();
  if (!overlay) return;
  overlay.classList.remove("open");
  overlay.setAttribute("aria-hidden", "true");
}

function openPasswordResetModal() {
  const overlay = document.getElementById("passwordResetModal");
  const form = document.getElementById("passwordResetForm");
  const metaNode = document.getElementById("passwordResetMeta");
  const currentPasswordInput = document.getElementById("passwordResetCurrentPassword");
  const session = activePasswordResetSession();
  if (!overlay || !form || !session?.backendToken) return;

  const roleLabel = session.role === "supervisor" ? "Supervisor" : "Manager";
  const username = String(session.username || "").trim().toLowerCase();
  if (metaNode) {
    metaNode.textContent = username
      ? `Update the password for your signed-in ${roleLabel.toLowerCase()} account (@${username}).`
      : `Update the password for your signed-in ${roleLabel.toLowerCase()} account.`;
  }

  passwordResetModalBusy = false;
  form.reset();
  setPasswordResetStatus();
  overlay.classList.add("open");
  overlay.setAttribute("aria-hidden", "false");
  currentPasswordInput?.focus();
}

async function submitPasswordResetModal(event) {
  event.preventDefault();
  const session = activePasswordResetSession();
  if (!session?.backendToken) {
    setPasswordResetStatus("Login required.", "error");
    return;
  }

  const currentPasswordInput = document.getElementById("passwordResetCurrentPassword");
  const newPasswordInput = document.getElementById("passwordResetNewPassword");
  const confirmPasswordInput = document.getElementById("passwordResetConfirmPassword");
  const submitBtn = document.getElementById("passwordResetSubmit");
  const currentPassword = String(currentPasswordInput?.value || "");
  const newPassword = String(newPasswordInput?.value || "");
  const confirmPassword = String(confirmPasswordInput?.value || "");

  if (!currentPassword || !newPassword || !confirmPassword) {
    setPasswordResetStatus("Current password, new password and confirmation are required.", "error");
    return;
  }
  if (newPassword.length < 6) {
    setPasswordResetStatus("New password must be at least 6 characters.", "error");
    newPasswordInput?.focus();
    return;
  }
  if (newPassword !== confirmPassword) {
    setPasswordResetStatus("New password and confirmation must match.", "error");
    confirmPasswordInput?.focus();
    return;
  }

  passwordResetModalBusy = true;
  setPasswordResetStatus("Updating password...");
  if (submitBtn) {
    submitBtn.disabled = true;
    submitBtn.textContent = "Updating...";
  }

  try {
    await apiRequest("/api/me/password", {
      method: "POST",
      token: session.backendToken,
      body: {
        currentPassword,
        newPassword
      }
    });
    setPasswordResetStatus("Password updated.", "success");
    const form = document.getElementById("passwordResetForm");
    form?.reset();
  } catch (error) {
    setPasswordResetStatus(String(error?.message || "Could not update password.").trim(), "error");
    if (currentPasswordInput) currentPasswordInput.value = "";
    currentPasswordInput?.focus();
  } finally {
    passwordResetModalBusy = false;
    if (submitBtn) {
      submitBtn.disabled = false;
      submitBtn.textContent = "Update Password";
    }
  }
}

function bindPasswordResetModal() {
  const overlay = document.getElementById("passwordResetModal");
  const closeBtn = document.getElementById("closePasswordResetModal");
  const form = document.getElementById("passwordResetForm");
  const openBtns = [
    document.getElementById("homeResetPasswordBtn"),
    document.getElementById("lineWorkspaceResetPasswordBtn")
  ].filter(Boolean);

  openBtns.forEach((btn) => {
    btn.addEventListener("click", openPasswordResetModal);
  });
  if (!overlay || !closeBtn || !form) return;

  closeBtn.addEventListener("click", closePasswordResetModal);
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) closePasswordResetModal();
  });
  form.addEventListener("submit", submitPasswordResetModal);
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && overlay.classList.contains("open")) {
      closePasswordResetModal();
    }
  });
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
          sourceRow: row,
          excludeFromCalculation: isDowntimeExcludedFromCalculation(row),
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

function renderDayVisualiserTo(rootId, data, selectedDate, selectedShift, stageNameResolver, lineContext = null) {
  const root = document.getElementById(rootId);
  if (!root) return;
  const blocks = buildDayVisualiserBlocks(data, selectedDate, selectedShift, stageNameResolver);

  const dayStartHour = 5;
  const startMins = dayStartHour * 60;
  const endMins = startMins + 24 * 60;
  const rangeMins = endMins - startMins;
  const hourMarks = Array.from({ length: 25 }, (_, offset) => dayStartHour + offset);
  const glance = buildDayAtGlance(blocks, stageNameResolver, selectedShift);
  const selectedShiftLabel = selectedShift === "Day" || selectedShift === "Night" ? selectedShift : "All shifts";
  const canAddRecords = rootId === "dayVisualiserCanvas" && appState.appMode === "manager" && Boolean(lineContext?.id);

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
  const calculationDowntimeIntervals = visibleIntervals(blocks.downtime.filter((item) => !item.excludeFromCalculation));
  const overlappingDowntimeMins = intervalsOverlapMinutes(productionIntervals, calculationDowntimeIntervals);
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
      const downInRunMins = intervalsOverlapMinutes(runIntervals, calculationDowntimeIntervals);
      const computedNetRunMins = Math.max(0, num(row.netProductionTime) * timedLogShiftWeight(row, selectedShift));
      const fallbackNetRunMins = Math.max(0, grossMins - downInRunMins);
      const netRunMins = computedNetRunMins > 0 ? computedNetRunMins : fallbackNetRunMins;
      const netTraysPerMin = netRunMins > 0 ? units / netRunMins : 0;
      const traffic = productionTrafficForRate(netTraysPerMin, netRunMins);
      return `
        <article class="day-glance-run-tile">
          <div class="day-glance-run-main">
            <span class="day-glance-light ${traffic.className}" aria-hidden="true"></span>
            <div class="day-glance-run-copy">
              <h5 class="day-glance-run-title">${htmlEscape(productName)}</h5>
              <span class="day-glance-run-rate">${formatNum(netTraysPerMin, 2)} net trays / min</span>
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
  const blockDetailsById = {};
  let blockDetailCounter = 0;

  const renderLane = (label, totalLabel, items, recordType = "") => {
    const laneLines = hourMarks
      .map((hour) => {
        const left = ((hour - dayStartHour) / 24) * 100;
        return `<div class="day-viz-hour-line ${hour % 2 === 0 ? "major" : ""}" style="left:${left}%"></div>`;
      })
      .join("");
    const addButton = canAddRecords && recordType
      ? `
        <button
          type="button"
          class="ghost-btn day-viz-add-record-btn"
          data-day-viz-add-record="${htmlEscape(recordType)}"
          data-day-viz-line-id="${htmlEscape(String(lineContext?.id || ""))}"
          data-day-viz-date="${htmlEscape(selectedDate)}"
          data-day-viz-shift="${htmlEscape(selectedShift)}"
          aria-label="${htmlEscape(`Add ${label.toLowerCase()} record for ${selectedDate}`)}"
        >
          Add record
        </button>
      `
      : "";

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
        blockDetailCounter += 1;
        const detailId = `${rootId}-${blockDetailCounter}`;
        const durationMins = Math.max(0, num(item.end) - num(item.start));
        const timeRange = `${formatMinutesToClockLabel(item.start)} - ${formatMinutesToClockLabel(item.end)}`;
        const typeLabel = dayVizBlockTypeLabel(item.type);
        blockDetailsById[detailId] = {
          rootId,
          lineId: String(lineContext?.id || ""),
          logId: String(item?.sourceRow?.id || ""),
          title: String(item.title || ""),
          subtitle: String(item.sub || ""),
          date: selectedDate,
          shiftLabel: selectedShiftLabel,
          laneLabel: label,
          typeKey: String(item.type || ""),
          typeLabel,
          timeRange,
          durationMins,
          excludeFromCalculation: Boolean(item.excludeFromCalculation)
        };
        const ariaLabel = `${label}: ${item.title}. ${timeRange}. ${formatNum(durationMins, 1)} minutes.`;
        return `
          <article class="day-viz-block ${item.type}" style="left:${clampedLeft}%;width:${clampedWidth}%;" title="${htmlEscape(tooltip)}" data-day-viz-block-id="${htmlEscape(detailId)}" role="button" tabindex="0" aria-label="${htmlEscape(ariaLabel)}">
            <span class="day-viz-title">${htmlEscape(item.title)}</span>
            <span class="day-viz-sub">${htmlEscape(item.sub)}</span>
          </article>
        `;
      })
      .join("");

    return `
      <div class="day-viz-swimlane">
        <div class="day-viz-lane-label">
          <div class="day-viz-lane-head">
            <span class="day-viz-lane-title">${htmlEscape(label)}</span>
            ${addButton}
          </div>
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
        ${renderLane("Shift", laneTotals.shift, blocks.shifts, "shift")}
        ${renderLane("Break", laneTotals.break, blocks.breaks, "break")}
        ${renderLane("Production", laneTotals.production, blocks.runs, "run")}
        ${renderLane("Downtime", laneTotals.downtime, blocks.downtime, "downtime")}
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
          <p class="day-glance-meta">Per-run net run rate (net trays/min). Uses net production time for each run.</p>
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
  root.__dayVizBlockDetails = blockDetailsById;
}

function renderDayVisualiser() {
  renderDayVisualiserTo("dayVisualiserCanvas", derivedData(), state.selectedDate, state.selectedShift, stageNameById, state);
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
    if (!isWeekendDate(cursor)) dates.push(formatDateLocal(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }
  return dates;
}

function buildLineTrendBuckets(anchorIsoDate, range = "day") {
  const safeRange = LINE_TREND_RANGES.includes(String(range || "").toLowerCase()) ? String(range).toLowerCase() : "day";
  const anchor = parseDateLocal(normalizeWeekdayIsoDate(anchorIsoDate || todayISO(), { direction: -1 }));
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
    const weekdays = [];
    const cursor = new Date(anchor.getFullYear(), anchor.getMonth(), anchor.getDate());
    while (weekdays.length < 14) {
      if (!isWeekendDate(cursor)) {
        weekdays.push(new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate()));
      }
      cursor.setDate(cursor.getDate() - 1);
    }
    weekdays.reverse().forEach((current) => {
      pushBucket(current, current, formatIsoDateLabel(formatDateLocal(current), { month: "short", day: "numeric" }));
    });
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
    let utilGrossSum = 0;
    let utilGrossCount = 0;
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
      const utilGross = Math.max(0, num(daily.lineUtilGross));
      if (Number.isFinite(utilGross)) {
        utilGrossSum += utilGross;
        utilGrossCount += 1;
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
      lineUtilGross: utilGrossCount > 0 ? utilGrossSum / utilGrossCount : 0,
      netRunRate: runRateCount > 0 ? runRateSum / runRateCount : 0
    };
  });
}

function aggregateLineTrendProductSeries(line, buckets, shift) {
  const data = derivedDataForLine(line || {});
  const runRows = Array.isArray(data.runRows) ? data.runRows : [];
  const bucketTotals = buckets.map(() => new Map());
  const totalUnitsByProduct = new Map();
  const totalUnitsByBucket = buckets.map(() => 0);
  const totalDetailsByBucket = buckets.map(() => []);

  buckets.forEach((bucket, bucketIndex) => {
    const productTotals = bucketTotals[bucketIndex];
    bucket.dates.forEach((date) => {
      selectedShiftRowsByDate(runRows, date, shift, { line }).forEach((row) => {
        const weight = Math.max(0, timedLogShiftWeight(row, shift));
        if (weight <= 0) return;
        const rawProduct = String(row?.product || "").trim();
        const productName = catalogProductCanonicalName(rawProduct, { lineId: line?.id }) || rawProduct || "Unspecified";
        const entry = productTotals.get(productName) || { units: 0, netMins: 0, details: [] };
        const weightedUnits = Math.max(0, num(row?.unitsProduced)) * weight;
        const weightedNetMins = Math.max(0, num(row?.netProductionTime)) * weight;
        const detail = {
          runId: String(row?.id || ""),
          date: String(row?.date || ""),
          shiftLabel: resolveTimedLogShiftLabel(row, line, "productionStartTime", "finishTime"),
          product: productName,
          productionStartTime: String(row?.productionStartTime || ""),
          finishTime: String(row?.finishTime || ""),
          unitsProduced: Math.max(0, num(row?.unitsProduced)),
          netProductionTime: Math.max(0, num(row?.netProductionTime)),
          shiftWeight: weight,
          weightedUnits,
          weightedNetMins,
          weightedRate: weightedNetMins > 0 ? weightedUnits / weightedNetMins : 0
        };
        entry.units += weightedUnits;
        entry.netMins += weightedNetMins;
        entry.details.push(detail);
        productTotals.set(productName, entry);
        totalUnitsByBucket[bucketIndex] += weightedUnits;
        totalDetailsByBucket[bucketIndex].push(detail);
        if (!totalUnitsByProduct.has(productName)) totalUnitsByProduct.set(productName, 0);
        totalUnitsByProduct.set(productName, totalUnitsByProduct.get(productName) + weightedUnits);
      });
    });
  });

  const productNames = Array.from(totalUnitsByProduct.entries())
    .sort((a, b) => {
      const unitsDiff = num(b[1]) - num(a[1]);
      if (Math.abs(unitsDiff) > 0.0001) return unitsDiff;
      return a[0].localeCompare(b[0]);
    })
    .map(([name]) => name);

  const series = productNames.map((name) => ({
    name,
    units: bucketTotals.map((totals) => Math.max(0, num(totals.get(name)?.units))),
    values: bucketTotals.map((totals) => {
      const entry = totals.get(name);
      if (!entry || num(entry.netMins) <= 0) return null;
      return Math.max(0, num(entry.units) / num(entry.netMins));
    }),
    details: bucketTotals.map((totals) => {
      const entry = totals.get(name);
      return Array.isArray(entry?.details) ? entry.details : [];
    })
  }));

  const maxRate = series.reduce(
    (max, item) => Math.max(max, ...item.values.map((value) => (Number.isFinite(value) ? value : 0))),
    0
  );
  return {
    series,
    maxRate: Math.max(1, maxRate),
    totalUnitsByBucket,
    totalDetailsByBucket
  };
}

function lineTrendSeriesColor(index) {
  const palette = ["#1f4f8a", "#0f766e", "#a73838", "#8f4b08", "#6d28d9", "#0c4a6e", "#9f1239", "#166534", "#7c2d12", "#1e3a8a"];
  return palette[index % palette.length];
}

function lineTrendBucketDateRangeLabel(point) {
  if (!point) return "";
  const startIso = String(point.startIso || "").trim();
  const endIso = String(point.endIso || "").trim();
  if (!startIso) return "";
  if (!endIso || endIso === startIso) {
    return formatIsoDateLabel(startIso, { month: "short", day: "numeric", year: "numeric" });
  }
  return `${formatIsoDateLabel(startIso, { month: "short", day: "numeric" })} to ${formatIsoDateLabel(endIso, {
    month: "short",
    day: "numeric",
    year: "numeric"
  })}`;
}

function sortedLineTrendDetailRows(rows = []) {
  return (Array.isArray(rows) ? rows : []).slice().sort((a, b) => {
    const dateCmp = String(a?.date || "").localeCompare(String(b?.date || ""));
    if (dateCmp !== 0) return dateCmp;
    const startA = parseTimeToMinutes(a?.productionStartTime);
    const startB = parseTimeToMinutes(b?.productionStartTime);
    if (Number.isFinite(startA) && Number.isFinite(startB) && startA !== startB) return startA - startB;
    return String(a?.product || "").localeCompare(String(b?.product || ""));
  });
}

function lineTrendDetailRowsTable(rows = []) {
  const sortedRows = sortedLineTrendDetailRows(rows);
  if (!sortedRows.length) return `<p class="muted">No production runs contributed to this datapoint.</p>`;
  return `
    <div class="table-wrap">
      <table class="line-trend-detail-table">
        <thead>
          <tr>
            <th>Date</th>
            <th>Shift</th>
            <th>Product</th>
            <th>Start</th>
            <th>Finish</th>
            <th>Units</th>
            <th>Net Min</th>
            <th>Weight</th>
            <th>Weighted Units</th>
            <th>Weighted Net Min</th>
            <th>Weighted Tray/Min</th>
          </tr>
        </thead>
        <tbody>
          ${sortedRows
            .map(
              (row) => `
              <tr>
                <td>${htmlEscape(String(row.date || ""))}</td>
                <td>${htmlEscape(String(row.shiftLabel || ""))}</td>
                <td>${htmlEscape(String(row.product || ""))}</td>
                <td>${htmlEscape(formatTime12h(row.productionStartTime))}</td>
                <td>${htmlEscape(formatTime12h(row.finishTime))}</td>
                <td>${formatNum(row.unitsProduced, 0)}</td>
                <td>${formatNum(row.netProductionTime, 1)}</td>
                <td>${formatNum(Math.max(0, num(row.shiftWeight)) * 100, 0)}%</td>
                <td>${formatNum(row.weightedUnits, 1)}</td>
                <td>${formatNum(row.weightedNetMins, 2)}</td>
                <td>${formatNum(row.weightedRate, 2)}</td>
              </tr>
            `
            )
            .join("")}
        </tbody>
      </table>
    </div>
  `;
}

function closeLineTrendDetailModal() {
  const overlay = document.getElementById("lineTrendDetailModal");
  if (!overlay) return;
  overlay.classList.remove("open");
  overlay.setAttribute("aria-hidden", "true");
}

function openLineTrendDetailModal({ title = "", meta = "", detailRows = [], primaryLabel = "", primaryValue = "" } = {}) {
  const overlay = document.getElementById("lineTrendDetailModal");
  const titleNode = document.getElementById("lineTrendDetailTitle");
  const metaNode = document.getElementById("lineTrendDetailMeta");
  const summaryNode = document.getElementById("lineTrendDetailSummary");
  const bodyNode = document.getElementById("lineTrendDetailBody");
  if (!overlay || !titleNode || !metaNode || !summaryNode || !bodyNode) return;

  const totalWeightedUnits = detailRows.reduce((sum, row) => sum + Math.max(0, num(row?.weightedUnits)), 0);
  const totalWeightedNetMins = detailRows.reduce((sum, row) => sum + Math.max(0, num(row?.weightedNetMins)), 0);
  const blendedRate = totalWeightedNetMins > 0 ? totalWeightedUnits / totalWeightedNetMins : 0;
  titleNode.textContent = title;
  metaNode.textContent = meta;
  summaryNode.innerHTML = `
    <span class="line-trend-detail-pill"><strong>${htmlEscape(primaryLabel)}</strong><span>${htmlEscape(primaryValue)}</span></span>
    <span class="line-trend-detail-pill"><strong>Rows</strong><span>${formatNum(detailRows.length, 0)}</span></span>
    <span class="line-trend-detail-pill"><strong>Weighted Units</strong><span>${formatNum(totalWeightedUnits, 1)}</span></span>
    <span class="line-trend-detail-pill"><strong>Weighted Net Min</strong><span>${formatNum(totalWeightedNetMins, 2)}</span></span>
    <span class="line-trend-detail-pill"><strong>Blended Tray/Min</strong><span>${formatNum(blendedRate, 2)}</span></span>
  `;
  bodyNode.innerHTML = lineTrendDetailRowsTable(detailRows);
  overlay.classList.add("open");
  overlay.setAttribute("aria-hidden", "false");
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

function renderLineTrendUnitsChart(points, productTrend = { series: [], maxRate: 1, totalUnitsByBucket: [], totalDetailsByBucket: [] }) {
  const root = document.getElementById("lineTrendUnitsChart");
  if (!root) return;
  if (!points.length) {
    root.innerHTML = `<div class="empty-chart">No trend data available.</div>`;
    return;
  }

  const series = Array.isArray(productTrend?.series) ? productTrend.series : [];
  const totalDetailsByBucket = Array.isArray(productTrend?.totalDetailsByBucket) ? productTrend.totalDetailsByBucket : [];
  const width = 1280;
  const height = 430;
  const pad = { top: 34, right: 104, bottom: 78, left: 76 };
  const chartW = width - pad.left - pad.right;
  const chartH = height - pad.top - pad.bottom;
  const focusKey = String(state?.lineTrendLegendFocusKey || "").trim();
  const focusedProductName = focusKey.startsWith("product:") ? focusKey.slice("product:".length) : "";
  const focusedSeries = focusedProductName ? series.find((item) => item.name === focusedProductName) || null : null;
  const visibleSeries = !focusKey ? series : focusKey === "__units__" ? [] : focusedSeries ? [focusedSeries] : [];

  const barValues = focusedSeries
    ? (Array.isArray(focusedSeries.units) ? focusedSeries.units : []).map((value) => Math.max(0, num(value)))
    : points.map((point) => Math.max(0, num(point.units)));
  const maxUnits = Math.max(1, ...barValues.map((value) => Math.max(0, num(value))));
  const maxRateFromVisible = visibleSeries.reduce(
    (max, item) => Math.max(max, ...item.values.map((value) => (Number.isFinite(value) ? value : 0))),
    0
  );
  const maxRate = Math.max(1, maxRateFromVisible || num(productTrend?.maxRate));
  const stepX = points.length > 1 ? chartW / (points.length - 1) : chartW;
  const x = (index) => pad.left + stepX * index;
  const yUnits = (value) => pad.top + chartH - (Math.max(0, value) / maxUnits) * chartH;
  const yRate = (value) => pad.top + chartH - (Math.max(0, value) / maxRate) * chartH;
  const barKey = focusedSeries ? `product:${focusedSeries.name}` : "__units__";
  const barLegendLabel = focusedSeries ? `${focusedSeries.name} units (bars, left axis)` : "Total units (bars, left axis)";

  const barWidth = Math.max(14, Math.min(36, stepX * 0.58));
  const bars = points
    .map((point, index) => {
      const unitsValue = Math.max(0, num(barValues[index]));
      const y = yUnits(unitsValue);
      const h = pad.top + chartH - y;
      const title = focusedSeries
        ? `${focusedSeries.name} ${point.label}: ${formatNum(unitsValue, 1)} units`
        : `${point.label}: ${formatNum(unitsValue, 0)} units produced`;
      return `<rect x="${x(index) - barWidth / 2}" y="${y}" width="${barWidth}" height="${Math.max(
        1,
        h
      )}" rx="3" class="bar-units line-trend-clickable" data-line-trend-bar-index="${index}" data-line-trend-bar-key="${encodeURIComponent(
        barKey
      )}"><title>${htmlEscape(title)}</title></rect>`;
    })
    .join("");

  const ticks = [0, 0.25, 0.5, 0.75, 1]
    .map((fraction) => {
      const y = pad.top + chartH - fraction * chartH;
      const unitsTick = Math.round(maxUnits * fraction);
      const rateTick = roundToDecimals(maxRate * fraction, maxRate < 10 ? 2 : 1);
      const rightTick = visibleSeries.length
        ? `<text x="${pad.left + chartW + 12}" y="${y + 5}" text-anchor="start" class="axis">${formatNum(
          rateTick,
          maxRate < 10 ? 2 : 1
        )}</text>`
        : "";
      return `<g>
        <line x1="${pad.left}" y1="${y}" x2="${pad.left + chartW}" y2="${y}" class="grid"/>
        <text x="${pad.left - 12}" y="${y + 5}" text-anchor="end" class="axis">${formatNum(unitsTick, 0)}</text>
        ${rightTick}
      </g>`;
    })
    .join("");

  const productLines = visibleSeries
    .map((item) => {
      const productIndex = series.findIndex((entry) => entry.name === item.name);
      const color = lineTrendSeriesColor(Math.max(0, productIndex));
      let drawing = false;
      let path = "";
      item.values.forEach((value, pointIndex) => {
        if (!Number.isFinite(value)) {
          drawing = false;
          return;
        }
        path += `${drawing ? " L" : " M"}${x(pointIndex)},${yRate(value)}`;
        drawing = true;
      });
      if (!path.trim()) return "";
      return `<path d="${path.trim()}" class="line-product-rate" style="stroke:${color};"><title>${htmlEscape(item.name)}</title></path>`;
    })
    .join("");

  const productAverageLines = visibleSeries
    .map((item) => {
      const productIndex = series.findIndex((entry) => entry.name === item.name);
      const color = lineTrendSeriesColor(Math.max(0, productIndex));
      const values = item.values.filter((value) => Number.isFinite(value));
      if (!values.length) return "";
      const avg = values.reduce((sum, value) => sum + num(value), 0) / values.length;
      const y = yRate(avg);
      return `<line x1="${pad.left}" y1="${y}" x2="${pad.left + chartW}" y2="${y}" class="line-product-rate-avg" style="stroke:${color};">
        <title>${htmlEscape(item.name)} average: ${formatNum(avg, 2)} trays/min</title>
      </line>`;
    })
    .join("");

  const productDots = visibleSeries
    .map((item) => {
      const productIndex = series.findIndex((entry) => entry.name === item.name);
      const color = lineTrendSeriesColor(Math.max(0, productIndex));
      return item.values
        .map((value, pointIndex) => {
          if (!Number.isFinite(value)) return "";
          return `<circle cx="${x(pointIndex)}" cy="${yRate(value)}" r="4" class="dot-product-rate line-trend-clickable" style="fill:${color};" data-line-trend-point-index="${pointIndex}" data-line-trend-point-product="${encodeURIComponent(
            item.name
          )}">
            <title>${htmlEscape(item.name)} | ${points[pointIndex]?.label || ""}: ${formatNum(value, 2)} trays/min</title>
          </circle>`;
        })
        .join("");
    })
    .join("");

  const legendItem = (key, color, name, valueText) => {
    const active = focusKey === key;
    const dimmed = Boolean(focusKey) && !active;
    return `<button type="button" class="line-trend-product-legend-item${active ? " is-active" : ""}${dimmed ? " is-dimmed" : ""}" data-line-trend-legend-key="${encodeURIComponent(
      key
    )}" aria-pressed="${String(active)}">
      <span class="line-trend-product-legend-swatch" style="background:${color};"></span>
      <span class="line-trend-product-legend-name">${htmlEscape(name)}</span>
      <span class="line-trend-product-legend-value">${htmlEscape(valueText)}</span>
    </button>`;
  };

  const legend = [
    legendItem("__units__", "rgba(16, 24, 32, 0.3)", "Total units", `${formatNum(points[points.length - 1]?.units || 0, 0)} units`),
    ...series.map((item, seriesIndex) => {
      const color = lineTrendSeriesColor(seriesIndex);
      const values = item.values.filter((value) => Number.isFinite(value));
      const average = values.length ? values.reduce((sum, value) => sum + num(value), 0) / values.length : null;
      const averageLabel = Number.isFinite(average) ? `${formatNum(average, 2)} trays/min avg` : "No runs";
      return legendItem(`product:${item.name}`, color, item.name, averageLabel);
    })
  ].join("");

  const labelSkip = points.length > 14 ? Math.ceil(points.length / 10) : 1;
  const labels = points
    .map((point, index) => {
      const label = index % labelSkip === 0 || index === points.length - 1 ? point.label : "";
      return `<text x="${x(index)}" y="${height - 24}" text-anchor="middle" class="axis">${htmlEscape(label)}</text>`;
    })
    .join("");

  root.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" class="trend-svg line-trend-svg" role="img" aria-label="Total production and product tray per minute trend by period">
      ${ticks}
      <line x1="${pad.left}" y1="${pad.top + chartH}" x2="${pad.left + chartW}" y2="${pad.top + chartH}" class="axis-line" />
      <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${pad.top + chartH}" class="axis-line" />
      ${visibleSeries.length ? `<line x1="${pad.left + chartW}" y1="${pad.top}" x2="${pad.left + chartW}" y2="${pad.top + chartH}" class="axis-line" />` : ""}
      ${bars}
      ${productAverageLines}
      ${productLines}
      ${productDots}
      ${labels}
      <text x="${pad.left}" y="${20}" class="legend units">${htmlEscape(barLegendLabel)}</text>
      ${visibleSeries.length ? `<text x="${pad.left + 292}" y="${20}" class="legend rate">Tray/min by product (solid) + product avg (dashed)</text>` : ""}
    </svg>
    <div class="line-trend-product-legend">${legend}</div>
  `;

  root.querySelectorAll("[data-line-trend-legend-key]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const encoded = String(btn.getAttribute("data-line-trend-legend-key") || "");
      const key = decodeURIComponent(encoded);
      const current = String(state?.lineTrendLegendFocusKey || "").trim();
      state.lineTrendLegendFocusKey = current === key ? "" : key;
      renderLineTrends();
      saveState();
    });
  });

  root.querySelectorAll("[data-line-trend-bar-index]").forEach((bar) => {
    bar.addEventListener("click", () => {
      const bucketIndex = Math.max(0, Math.floor(num(bar.getAttribute("data-line-trend-bar-index"))));
      const point = points[bucketIndex];
      if (!point) return;
      const encodedKey = String(bar.getAttribute("data-line-trend-bar-key") || "");
      const key = decodeURIComponent(encodedKey);
      const isProduct = key.startsWith("product:");
      const productName = isProduct ? key.slice("product:".length) : "";
      const selectedSeries = isProduct ? series.find((item) => item.name === productName) || null : null;
      const detailRows = isProduct
        ? (selectedSeries?.details?.[bucketIndex] || [])
        : (Array.isArray(totalDetailsByBucket[bucketIndex]) ? totalDetailsByBucket[bucketIndex] : []);
      const unitsValue = isProduct ? Math.max(0, num(selectedSeries?.units?.[bucketIndex])) : Math.max(0, num(barValues[bucketIndex]));
      const bucketRange = lineTrendBucketDateRangeLabel(point);
      openLineTrendDetailModal({
        title: isProduct ? `${productName} Units Breakdown` : "Total Units Breakdown",
        meta: `${point.label}${bucketRange ? ` | ${bucketRange}` : ""} | ${state.selectedShift} shift`,
        detailRows,
        primaryLabel: "Bar Value",
        primaryValue: `${formatNum(unitsValue, isProduct ? 1 : 0)} units`
      });
    });
  });

  root.querySelectorAll("[data-line-trend-point-index][data-line-trend-point-product]").forEach((dot) => {
    dot.addEventListener("click", () => {
      const bucketIndex = Math.max(0, Math.floor(num(dot.getAttribute("data-line-trend-point-index"))));
      const point = points[bucketIndex];
      if (!point) return;
      const encodedProduct = String(dot.getAttribute("data-line-trend-point-product") || "");
      const productName = decodeURIComponent(encodedProduct);
      const selectedSeries = series.find((item) => item.name === productName) || null;
      if (!selectedSeries) return;
      const pointValue = Number.isFinite(selectedSeries.values?.[bucketIndex]) ? selectedSeries.values[bucketIndex] : 0;
      const detailRows = selectedSeries.details?.[bucketIndex] || [];
      const bucketRange = lineTrendBucketDateRangeLabel(point);
      openLineTrendDetailModal({
        title: `${productName} Tray/Min Breakdown`,
        meta: `${point.label}${bucketRange ? ` | ${bucketRange}` : ""} | ${state.selectedShift} shift`,
        detailRows,
        primaryLabel: "Point Value",
        primaryValue: `${formatNum(pointValue, 2)} trays/min`
      });
    });
  });
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
  const productTrend = aggregateLineTrendProductSeries(state, buckets, state.selectedShift || "Day");
  const validLegendKeys = new Set(["__units__", ...(productTrend.series || []).map((item) => `product:${String(item?.name || "")}`)]);
  if (!validLegendKeys.has(String(state.lineTrendLegendFocusKey || ""))) {
    state.lineTrendLegendFocusKey = "";
  }
  const current = points[points.length - 1] || { units: 0, downtime: 0, lineUtil: 0, lineUtilGross: 0, netRunRate: 0 };
  const startIso = points[0]?.startIso || state.selectedDate || todayISO();
  const endIso = points[points.length - 1]?.endIso || state.selectedDate || todayISO();
  const rangeLabel = range === "day" ? "Daily" : range === "week" ? "Weekly" : range === "month" ? "Monthly" : "Quarterly";
  if (rangeMeta) {
    rangeMeta.textContent = `${rangeLabel} view | ${state.selectedShift} shift | ${formatIsoDateLabel(startIso)} to ${formatIsoDateLabel(endIso)}`;
  }

  setText("lineTrendKpiUnits", formatNum(current.units, 0));
  setText("lineTrendKpiDowntime", `${formatNum(current.downtime, 1)} min`);
  setText("lineTrendKpiUtilisation", `${formatNum(current.lineUtil, 1)}%`);
  setText("lineTrendKpiUtilisationGross", `${formatNum(current.lineUtilGross, 1)}%`);
  setText("lineTrendKpiRunRate", `${formatNum(current.netRunRate, 2)} u/min`);

  renderLineTrendUnitsChart(points, productTrend);
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
    (id) => stageNameByIdForLine(line, id),
    line
  );
}

function computeDayVisualiserKeyStageMetrics(
  line,
  stageId,
  {
    stages = [],
    selectedRunRows = [],
    selectedCalculationDowntimeRows = [],
    selectedShift = "Day",
    shiftMins = 0,
    netRunRate = 0
  } = {}
) {
  const stageList = Array.isArray(stages) && stages.length ? stages : line?.stages || [];
  const normalizedStageId = normalizeDayVisualiserKeyStageId(line, stageId);
  const selectedStageIndex = stageList.findIndex((stage) => stage.id === normalizedStageId);
  const selectedStage = selectedStageIndex >= 0 ? stageList[selectedStageIndex] : null;
  if (!selectedStage) {
    return {
      stage: null,
      stageIndex: -1,
      stageLabel: "Select key stage",
      crewCount: 0,
      stageDowntime: 0,
      stageNetRunRate: 0,
      perCrewRate: 0
    };
  }

  const baseCrewMap = crewMapForLineShift(line, selectedShift, stageList);
  const selectedRunPattern = latestRunCrewingPatternForRows(line, selectedRunRows, selectedShift);
  const crewCount = stageCrewCountForVisual(line, selectedStage, baseCrewMap, selectedRunPattern);
  const stageDowntime = selectedCalculationDowntimeRows
    .filter((row) => matchesStage(selectedStage, row.equipment))
    .reduce((sum, row) => sum + downtimeMinutesForCalculations(row, selectedShift), 0);
  const uptimeRatio = shiftMins > 0 ? Math.max(0, (shiftMins - stageDowntime) / shiftMins) : 0;
  const stageNetRunRate = netRunRate * uptimeRatio;
  const perCrewRate = crewCount > 0 ? stageNetRunRate / crewCount : 0;

  return {
    stage: selectedStage,
    stageIndex: selectedStageIndex,
    stageLabel: stageDisplayName(selectedStage, selectedStageIndex),
    crewCount,
    stageDowntime,
    stageNetRunRate,
    perCrewRate
  };
}

function computeDayVisualiserKeyStageMetricsForDate(line, stageId, date, shift, data = derivedDataForLine(line || {})) {
  const stages = line?.stages?.length ? line.stages : STAGES;
  const selectedRunRows = selectedShiftRowsByDate(data.runRows, date, shift, { line });
  const selectedCalculationDowntimeRows = calculationDowntimeRows(selectedShiftRowsByDate(data.downtimeRows, date, shift, { line }));
  const selectedShiftRows = selectedShiftRowsByDate(data.shiftRows, date, shift, { line });
  const shiftMins = selectedShiftRows.reduce((sum, row) => sum + num(row.totalShiftTime), 0);
  const units = selectedRunRows.reduce((sum, row) => sum + num(row.unitsProduced) * timedLogShiftWeight(row, shift), 0);
  const totalNetTime = selectedRunRows.reduce((sum, row) => sum + num(row.netProductionTime) * timedLogShiftWeight(row, shift), 0);
  const netRunRate = totalNetTime > 0 ? units / totalNetTime : 0;

  return computeDayVisualiserKeyStageMetrics(line, stageId, {
    stages,
    selectedRunRows,
    selectedCalculationDowntimeRows,
    selectedShift: shift,
    shiftMins,
    netRunRate
  });
}

function renderDayVisualiserKeyStageKpi(
  line,
  {
    stages = [],
    selectedRunRows = [],
    selectedCalculationDowntimeRows = [],
    selectedShift = "Day",
    shiftMins = 0,
    netRunRate = 0
  } = {}
) {
  const selectNode = document.getElementById("dayKeyStageSelect");
  const nameNode = document.getElementById("dayKpiKeyStageName");
  const crewNode = document.getElementById("dayKpiKeyStageCrew");
  const rateNode = document.getElementById("dayKpiKeyStageRatePerCrew");
  if (!selectNode || !nameNode || !crewNode || !rateNode) return;
  const stageList = Array.isArray(stages) && stages.length ? stages : line?.stages || [];
  const normalizedKeyStageId = normalizeDayVisualiserKeyStageId(line, line?.dayVisualiserKeyStageId);
  if (line && normalizedKeyStageId !== String(line?.dayVisualiserKeyStageId || "")) {
    line.dayVisualiserKeyStageId = normalizedKeyStageId;
  }
  selectNode.innerHTML = [
    `<option value="">Select Key Stage</option>`,
    ...stageList.map((stage, index) => `<option value="${htmlEscape(stage.id)}">${htmlEscape(stageDisplayName(stage, index))}</option>`)
  ].join("");
  selectNode.value = normalizedKeyStageId;

  const selectedStageIndex = stageList.findIndex((stage) => stage.id === normalizedKeyStageId);
  const selectedStage = selectedStageIndex >= 0 ? stageList[selectedStageIndex] : null;
  if (!selectedStage) {
    nameNode.textContent = "Select key stage";
    crewNode.textContent = "Not set";
    rateNode.textContent = "-";
    return;
  }

  const keyStageMetrics = computeDayVisualiserKeyStageMetrics(line, normalizedKeyStageId, {
    stages: stageList,
    selectedRunRows,
    selectedCalculationDowntimeRows,
    selectedShift,
    shiftMins,
    netRunRate
  });
  nameNode.textContent = keyStageMetrics.stageLabel;
  crewNode.textContent = `${formatNum(keyStageMetrics.crewCount, 0)} crew`;
  rateNode.textContent = keyStageMetrics.crewCount > 0 ? `${formatNum(keyStageMetrics.perCrewRate, 2)} u/min/crew` : "No crew set";
}

function renderVisualiser() {
  const data = derivedData();
  const stages = getStages();
  const liveDowntimeByStage = liveEquipmentDowntimeByStage(state, stages);
  const selectedRunRows = selectedRows(data.runRows);
  const selectedDowntimeRows = selectedRows(data.downtimeRows);
  const selectedCalculationDowntimeRows = calculationDowntimeRows(selectedDowntimeRows);
  const selectedShiftRows = selectedRows(data.shiftRows);
  const selectedShift = state.selectedShift;

  const shiftMins = num(selectedShiftRows[0]?.totalShiftTime);
  const units = selectedRunRows.reduce((sum, row) => sum + num(row.unitsProduced) * timedLogShiftWeight(row, selectedShift), 0);
  const totalDowntime = selectedDowntimeRows.reduce((sum, row) => sum + num(row.downtimeMins) * timedLogShiftWeight(row, selectedShift), 0);
  const totalNetTime = selectedRunRows.reduce((sum, row) => sum + num(row.netProductionTime) * timedLogShiftWeight(row, selectedShift), 0);
  const netRunRate = totalNetTime > 0 ? units / totalNetTime : 0;
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
  const isFlowMoving = lineHasMovingFlow(state);
  map.classList.toggle("flow-running", isFlowMoving);
  map.classList.toggle("flow-static", !isFlowMoving);
  const activeCrew = crewMapForLineShift(state, state.selectedShift, stages);
  const liveRunPattern = liveRunCrewingPatternForLine(state);
  renderDayVisualiserKeyStageKpi(state, {
    stages,
    selectedRunRows,
    selectedCalculationDowntimeRows,
    selectedShift,
    shiftMins,
    netRunRate
  });
  const selectedBottleneck = resolveLineBottleneckStage(state, stages, selectedShift, (stageId, selectedShiftValue) =>
    stageTotalMaxThroughput(stageId, selectedShiftValue)
  );
  let bottleneckCard = null;
  let bottleneckUtilisation = 0;
  let bottleneckGrossUtilisation = 0;
  const guides = lineFlowGuidesForMap(stages, state.flowGuides);
  appendFlowGuidesToMap(map, guides, { editable: Boolean(state.visualEditMode) });

  stages.forEach((stage, index) => {
    const stageDowntime = selectedCalculationDowntimeRows
      .filter((row) => matchesStage(stage, row.equipment))
      .reduce((sum, row) => sum + downtimeMinutesForCalculations(row, selectedShift), 0);

    const uptimeRatio = shiftMins > 0 ? Math.max(0, (shiftMins - stageDowntime) / shiftMins) : 0;
    const stageRate = netRunRate * uptimeRatio;
    const totalMaxThroughput = stageTotalMaxThroughput(stage.id, state.selectedShift);
    const utilisation = totalMaxThroughput > 0 ? (stageRate / totalMaxThroughput) * 100 : 0;
    const grossUtilisation = totalMaxThroughput > 0 ? (netRunRate / totalMaxThroughput) * 100 : 0;
    const stageCrew = stageCrewCountForVisual(state, stage, activeCrew, liveRunPattern);
    const compact = stage.w * stage.h < 140;
    const status = statusClass(utilisation);

    const card = document.createElement("article");
    card.className = `stage-card group-${stage.group}${stage.kind ? ` kind-${stage.kind}` : ""} status-${status}`;
    const liveDown = liveDowntimeByStage[stage.id];
    const liveDownPill = stageLiveDowntimePillHtml(liveDown);
    if (liveDown?.level === "critical") card.classList.add("live-downtime-critical");
    if (liveDown?.level === "warning") card.classList.add("live-downtime-warning");
    card.setAttribute("data-stage-id", stage.id);
    card.style.left = `${stage.x}%`;
    card.style.top = `${stage.y}%`;
    card.style.width = `${stage.w}%`;
    card.style.height = `${stage.h}%`;
    card.classList.toggle("compact", compact);
    card.classList.toggle("selected", stage.id === state.selectedStageId);
    const isSelectedBottleneck = Boolean(selectedBottleneck) && selectedBottleneck.stage.id === stage.id;
    if (isSelectedBottleneck) {
      bottleneckUtilisation = Math.max(0, utilisation);
      bottleneckGrossUtilisation = Math.max(0, grossUtilisation);
      bottleneckCard = card;
    }
    if (!state.visualEditMode) {
      card.addEventListener("click", () => {
        state.selectedStageId = stage.id;
        state.trendDateCursor = normalizeWeekdayIsoDate(state.selectedDate || todayISO(), { direction: -1 });
        trendModalContext = { type: "stage", metricKey: "" };
        saveState();
        renderVisualiser();
        renderTrendModalContent();
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

    card.innerHTML = `
      <h3 class="stage-title">${stageDisplayName(stage, index)}</h3>
      <div class="stage-pill-row">
        ${liveDownPill}
        <span class="stage-crew-pill">${formatNum(stageCrew, 0)} Crew</span>
      </div>
      ${state.visualEditMode ? `<span class="stage-resize-handle" data-stage-resize="${stage.id}"></span>` : ""}
    `;
    if (liveDown) {
      card.setAttribute(
        "title",
        liveDown.level === "critical"
          ? `Urgent downtime: ${Math.floor(liveDown.minutes)} min`
          : `Active downtime: ${Math.floor(liveDown.minutes)} min`
      );
    }

    map.append(card);
  });

  if (bottleneckCard) {
    bottleneckCard.classList.add("bottleneck");
  }

  const lineUtil = Math.max(0, bottleneckGrossUtilisation);
  const lineUtilGross = Math.max(0, bottleneckUtilisation);
  const utilisationText = `${formatNum(Math.max(lineUtil, 0), 1)}%`;
  const utilisationGrossText = `${formatNum(Math.max(lineUtilGross, 0), 1)}%`;
  setText("kpiUtilisation", utilisationText);
  setText("dayKpiUtilisation", utilisationText);
  setText("kpiUtilisationGross", utilisationGrossText);
  setText("dayKpiUtilisationGross", utilisationGrossText);
}

function stageDailyMetrics(line, stage, date, shift, data) {
  const shiftRows = selectedShiftRowsByDate(data.shiftRows, date, shift, { line });
  const shiftMins = shiftRows.reduce((sum, row) => sum + num(row.totalShiftTime), 0);
  const stageDowntime = calculationDowntimeRows(selectedShiftRowsByDate(data.downtimeRows, date, shift, { line }))
    .filter((row) => matchesStage(stage, row.equipment))
    .reduce((sum, row) => sum + downtimeMinutesForCalculations(row, shift), 0);

  const runRows = selectedShiftRowsByDate(data.runRows, date, shift, { line });
  const units = runRows.reduce((sum, row) => sum + num(row.unitsProduced) * timedLogShiftWeight(row, shift), 0);
  const totalNetTime = runRows.reduce((sum, row) => sum + num(row.netProductionTime) * timedLogShiftWeight(row, shift), 0);
  const netRunRate = totalNetTime > 0 ? units / totalNetTime : 0;
  const uptimeRatio = shiftMins > 0 ? Math.max(0, (shiftMins - stageDowntime) / shiftMins) : 0;
  const stageEtc = netRunRate * uptimeRatio;
  const totalMaxThroughput = stageTotalMaxThroughputForLine(line, stage.id, shift);
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

function trendMonthPointLabel(month) {
  const [year, monthNumber] = String(month || "").split("-").map(Number);
  if (!year || !monthNumber) return String(month || "");
  return new Date(year, monthNumber - 1, 1).toLocaleDateString(undefined, { month: "short", year: "2-digit" });
}

function isTrendZeroValue(value) {
  return Math.abs(num(value)) < 0.0001;
}

function nonZeroAverageTrendValues(values = []) {
  return (Array.isArray(values) ? values : []).filter((value) => Number.isFinite(value) && !isTrendZeroValue(value));
}

function averageTrendMetricValues(values = []) {
  const safeValues = nonZeroAverageTrendValues(values);
  if (!safeValues.length) return 0;
  return safeValues.reduce((sum, value) => sum + value, 0) / safeValues.length;
}

function stageTrendPointIsZeroDay(point = {}) {
  return isTrendZeroValue(point.utilisation) && isTrendZeroValue(point.stageDowntime) && isTrendZeroValue(point.stageEtc);
}

function filterTrendWindowPoints(points = [], windowDates = [], { hideZeroDays = false, zeroDayPredicate = () => false } = {}) {
  const visibleDates = new Set(Array.isArray(windowDates) ? windowDates : []);
  return (Array.isArray(points) ? points : []).filter((point) => {
    if (!visibleDates.has(point.date)) return false;
    return !hideZeroDays || !zeroDayPredicate(point);
  });
}

function syncTrendDateCursorToAvailableDates(allDates = []) {
  if (!state) return "";
  const fallbackCursor = normalizeWeekdayIsoDate(state.trendDateCursor || state.selectedDate || todayISO(), { direction: -1 });
  if (!allDates.length) {
    state.trendDateCursor = fallbackCursor;
    return fallbackCursor;
  }
  const requestedCursor = isIsoDateValue(state.trendDateCursor)
    ? normalizeWeekdayIsoDate(state.trendDateCursor, { direction: -1, fallbackIso: fallbackCursor })
    : fallbackCursor;
  if (allDates.includes(requestedCursor)) {
    state.trendDateCursor = requestedCursor;
    return requestedCursor;
  }
  const previousDates = allDates.filter((date) => date <= requestedCursor);
  const nextCursor = previousDates.length ? previousDates[previousDates.length - 1] : allDates[0];
  state.trendDateCursor = nextCursor;
  return nextCursor;
}

function trendDailyWindowDates(allDates = [], cursorDate = "", windowSize = TREND_DAILY_WINDOW_SIZE) {
  if (!Array.isArray(allDates) || !allDates.length) return [];
  const idx = allDates.indexOf(cursorDate);
  const safeIdx = idx >= 0 ? idx : allDates.length - 1;
  const startIdx = Math.max(0, safeIdx - Math.max(1, Math.floor(num(windowSize)) || 1) + 1);
  return allDates.slice(startIdx, safeIdx + 1);
}

function syncTrendMonthToAvailableMonths(allMonths = []) {
  if (!state) return "";
  const selectedMonthFromDate = monthKey(state.selectedDate);
  if (!state.trendMonth) state.trendMonth = selectedMonthFromDate;
  if (allMonths.length && !allMonths.includes(state.trendMonth)) {
    state.trendMonth = allMonths.includes(selectedMonthFromDate) ? selectedMonthFromDate : allMonths[allMonths.length - 1];
  }
  return state.trendMonth || selectedMonthFromDate;
}

function buildStageTrendSeries() {
  const data = derivedData();
  const stages = getStages();
  const stageIdx = stages.findIndex((item) => item.id === state.selectedStageId);
  const stage = stageIdx >= 0 ? stages[stageIdx] : stages[0];
  const allDates = trendDatesForShift(state.selectedShift, data);
  const allMonths = Array.from(new Set(allDates.map(monthKey))).sort();
  const cursorDate = syncTrendDateCursorToAvailableDates(allDates);
  const windowDates = trendDailyWindowDates(allDates, cursorDate);
  syncTrendMonthToAvailableMonths(allMonths);

  const dailyPoints = allDates.map((date) => {
    const point = stageDailyMetrics(state, stage, date, state.selectedShift, data);
    return { ...point, label: date.slice(5) };
  });
  const monthlyPoints = allMonths.map((month) => {
    const monthDates = allDates.filter((date) => monthKey(date) === month);
    const dayPoints = monthDates.map((date) => stageDailyMetrics(state, stage, date, state.selectedShift, data));
    const averageDayPoints = dayPoints.filter((point) => !stageTrendPointIsZeroDay(point));
    const utilisation = averageDayPoints.length
      ? averageDayPoints.reduce((sum, point) => sum + point.utilisation, 0) / averageDayPoints.length
      : 0;
    const stageDowntime = dayPoints.reduce((sum, point) => sum + point.stageDowntime, 0);
    const stageEtc = averageDayPoints.length
      ? averageDayPoints.reduce((sum, point) => sum + point.stageEtc, 0) / averageDayPoints.length
      : 0;
    return {
      date: month,
      label: trendMonthPointLabel(month),
      utilisation,
      stageDowntime,
      stageEtc
    };
  });

  return {
    stage,
    stageIndex: stageIdx >= 0 ? stageIdx : 0,
    allDates,
    allMonths,
    cursorDate,
    windowDates,
    dailyPoints,
    monthlyPoints
  };
}

function dayKpiTrendConfig(metricKey) {
  const keyStageId = normalizeDayVisualiserKeyStageId(state, state?.dayVisualiserKeyStageId);
  const keyStageMetrics = computeDayVisualiserKeyStageMetrics(state, keyStageId, { stages: getStages(), selectedShift: state?.selectedShift || "Day" });
  const keyStageLabel = keyStageMetrics.stageLabel || "Key Stage";
  const crewFormat = (value) => `${formatNum(value, Math.abs(value - Math.round(value)) >= 0.05 ? 1 : 0)} crew`;

  const configs = {
    units: {
      title: "Total Units Trend",
      description: "Production units by selected day visualiser period.",
      legend: "Total units",
      aggregate: "sum",
      barClass: "bar-units",
      lineClass: "line-rate",
      dotClass: "dot-rate",
      lineColor: "#1f4f8a",
      avgTheme: "units",
      scaleFloor: 1,
      formatValue: (value) => `${formatNum(value, 0)} units`,
      formatTick: (value) => formatNum(value, 0)
    },
    downtime: {
      title: "Total Downtime Trend",
      description: "Logged downtime minutes in the selected shift.",
      legend: "Total downtime",
      aggregate: "sum",
      barClass: "bar-down",
      lineClass: "line-rate",
      dotClass: "dot-rate",
      lineColor: "#a73838",
      avgTheme: "down",
      scaleFloor: 1,
      formatValue: (value) => `${formatNum(value, value < 10 ? 1 : 0)} min`,
      formatTick: (value) => formatNum(value, value < 10 ? 1 : 0)
    },
    utilisation: {
      title: "Uptime Utilisation Trend",
      description: "Utilisation based on bottleneck stage ETC against max throughput.",
      legend: "Uptime utilisation",
      aggregate: "avg",
      barClass: "bar-units",
      lineClass: "line-util",
      dotClass: "dot-util",
      lineColor: "",
      avgTheme: "units",
      scaleFloor: 100,
      formatValue: (value) => `${formatNum(value, 1)}%`,
      formatTick: (value) => `${formatNum(value, 0)}%`
    },
    utilisationGross: {
      title: "Gross Utilisation Trend",
      description: "Gross utilisation before stage downtime is applied.",
      legend: "Gross utilisation",
      aggregate: "avg",
      barClass: "bar-units",
      lineClass: "line-util",
      dotClass: "dot-util",
      lineColor: "",
      avgTheme: "units",
      scaleFloor: 100,
      formatValue: (value) => `${formatNum(value, 1)}%`,
      formatTick: (value) => `${formatNum(value, 0)}%`
    },
    netRunRate: {
      title: "Net Run Rate Trend",
      description: "Units produced divided by net production minutes.",
      legend: "Net run rate",
      aggregate: "avg",
      barClass: "bar-units",
      lineClass: "line-rate",
      dotClass: "dot-rate",
      lineColor: "#1f4f8a",
      avgTheme: "units",
      scaleFloor: 1,
      formatValue: (value) => `${formatNum(value, 2)} u/min`,
      formatTick: (value) => formatNum(value, value < 10 ? 2 : 1)
    },
    keyStageCrew: {
      title: `${keyStageLabel} Crew Trend`,
      description: "Crew allocated to the selected key stage.",
      legend: `${keyStageLabel} crew`,
      aggregate: "avg",
      barClass: "bar-units",
      lineClass: "line-rate",
      dotClass: "dot-rate",
      lineColor: "#0f766e",
      avgTheme: "units",
      scaleFloor: 1,
      formatValue: crewFormat,
      formatTick: (value) => formatNum(value, Math.abs(value - Math.round(value)) >= 0.05 ? 1 : 0)
    },
    keyStageRatePerCrew: {
      title: `${keyStageLabel} Net Rate / Crew Trend`,
      description: "Key stage net run rate adjusted for stage downtime, divided by crew.",
      legend: `${keyStageLabel} net rate / crew`,
      aggregate: "avg",
      barClass: "bar-units",
      lineClass: "line-rate",
      dotClass: "dot-rate",
      lineColor: "#1f4f8a",
      avgTheme: "units",
      scaleFloor: 1,
      formatValue: (value) => `${formatNum(value, 2)} u/min/crew`,
      formatTick: (value) => formatNum(value, value < 10 ? 2 : 1)
    }
  };

  return configs[String(metricKey || "").trim()] || null;
}

function dayKpiTrendValue(metricKey, date, shift, data) {
  const safeMetricKey = String(metricKey || "").trim();
  const metrics = computeLineMetricsFromData(state, date, shift, data);
  if (safeMetricKey === "units") return Math.max(0, num(metrics.units));
  if (safeMetricKey === "downtime") return Math.max(0, num(metrics.totalDowntime));
  if (safeMetricKey === "utilisation") return Math.max(0, num(metrics.lineUtil));
  if (safeMetricKey === "utilisationGross") return Math.max(0, num(metrics.lineUtilGross));
  if (safeMetricKey === "netRunRate") return Math.max(0, num(metrics.netRunRate));
  if (safeMetricKey === "keyStageCrew" || safeMetricKey === "keyStageRatePerCrew") {
    const keyStageMetrics = computeDayVisualiserKeyStageMetricsForDate(
      state,
      state?.dayVisualiserKeyStageId,
      date,
      shift,
      data
    );
    return safeMetricKey === "keyStageCrew" ? Math.max(0, num(keyStageMetrics.crewCount)) : Math.max(0, num(keyStageMetrics.perCrewRate));
  }
  return 0;
}

function aggregateTrendMetricValues(values, mode = "avg") {
  const safeValues = values.filter((value) => Number.isFinite(value));
  if (!safeValues.length) return 0;
  if (mode === "sum") return safeValues.reduce((sum, value) => sum + value, 0);
  return averageTrendMetricValues(safeValues);
}

function buildDayKpiTrendSeries(metricKey) {
  const config = dayKpiTrendConfig(metricKey);
  if (!config || !state) {
    return { config: null, allDates: [], allMonths: [], dailyPoints: [], monthlyPoints: [] };
  }

  const data = derivedData();
  const allDates = trendDatesForShift(state.selectedShift, data);
  const allMonths = Array.from(new Set(allDates.map(monthKey))).sort();
  const cursorDate = syncTrendDateCursorToAvailableDates(allDates);
  const windowDates = trendDailyWindowDates(allDates, cursorDate);
  syncTrendMonthToAvailableMonths(allMonths);

  const dailyPoints = allDates.map((date) => ({
    date,
    targetDate: date,
    label: date.slice(5),
    value: dayKpiTrendValue(metricKey, date, state.selectedShift, data)
  }));
  const monthlyPoints = allMonths.map((month) => {
    const monthDates = allDates.filter((date) => monthKey(date) === month);
    const values = monthDates.map((date) => dayKpiTrendValue(metricKey, date, state.selectedShift, data));
    return {
      date: month,
      targetDate: monthDates[monthDates.length - 1] || "",
      label: trendMonthPointLabel(month),
      value: aggregateTrendMetricValues(values, config.aggregate)
    };
  });

  return {
    config,
    allDates,
    allMonths,
    cursorDate,
    windowDates,
    dailyPoints,
    monthlyPoints
  };
}

function renderDayKpiTrendChart(container, points, config, { scalePoints = points } = {}) {
  if (!container || !Array.isArray(points) || !points.length || !config) return;
  const width = 1180;
  const height = 340;
  const pad = { top: 26, right: 88, bottom: 58, left: 72 };
  const chartW = width - pad.left - pad.right;
  const chartH = height - pad.top - pad.bottom;
  const values = points.map((point) => Math.max(0, num(point.value)));
  const scaleValues = (Array.isArray(scalePoints) && scalePoints.length ? scalePoints : points).map((point) => Math.max(0, num(point.value)));
  const maxValue = Math.max(Math.max(0, num(config.scaleFloor)), ...scaleValues);
  const avgValue = averageTrendMetricValues(values);
  const stepX = points.length > 1 ? chartW / (points.length - 1) : chartW;
  const x = (index) => pad.left + stepX * index;
  const yValue = (value) => pad.top + chartH - (Math.max(0, value) / maxValue) * chartH;
  const avgY = yValue(avgValue);
  const avgLabel = `Avg ${config.formatValue(avgValue)}`;
  const avgCalloutWidth = Math.max(132, avgLabel.length * 8 + 24);
  const avgCalloutHeight = 28;
  const avgCalloutX = pad.left + chartW - avgCalloutWidth - 10;
  const avgCalloutY = Math.max(pad.top + 6, Math.min(avgY - avgCalloutHeight / 2, pad.top + chartH - avgCalloutHeight - 6));
  const barWidth = Math.max(12, Math.min(28, stepX * 0.46));
  const bars = points
    .map((point, index) => {
      const y = yValue(point.value);
      const h = pad.top + chartH - y;
      const targetDate = trendPointTargetDate(point);
      const targetDateLabel = targetDate ? formatIsoDateLabel(targetDate, { month: "short", day: "numeric", year: "numeric" }) : "";
      const clickableClass = targetDate ? " line-trend-clickable" : "";
      const interactiveAttrs = targetDate
        ? ` data-day-kpi-point-index="${index}" tabindex="0" role="button" aria-label="${htmlEscape(
            `${point.label}: ${config.formatValue(point.value)}. Open ${targetDateLabel} in Day visualiser.`
          )}"`
        : "";
      return `<rect x="${x(index) - barWidth / 2}" y="${y}" width="${barWidth}" height="${Math.max(
        1,
        h
      )}" rx="3" class="${htmlEscape(config.barClass)}${clickableClass}"${interactiveAttrs}><title>${htmlEscape(
        `${point.label}: ${config.formatValue(point.value)}${targetDateLabel ? ` | Open ${targetDateLabel}` : ""}`
      )}</title></rect>`;
    })
    .join("");
  const ticks = [0, 0.25, 0.5, 0.75, 1]
    .map((fraction) => maxValue * fraction)
    .map((tickValue) => {
      const y = yValue(tickValue);
      return `<g>
        <line x1="${pad.left}" y1="${y}" x2="${pad.left + chartW}" y2="${y}" class="grid"/>
        <text x="${pad.left - 12}" y="${y + 5}" text-anchor="end" class="axis">${htmlEscape(config.formatTick(tickValue))}</text>
      </g>`;
    })
    .join("");
  const labelSkip = points.length > 16 ? Math.ceil(points.length / 12) : 1;
  const labels = points
    .map((point, index) => {
      const label = index % labelSkip === 0 || index === points.length - 1 ? point.label : "";
      return `<text x="${x(index)}" y="${height - 16}" text-anchor="middle" class="axis">${htmlEscape(label)}</text>`;
    })
    .join("");

  container.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" class="trend-svg" role="img" aria-label="${htmlEscape(config.title)}">
      ${ticks}
      <line x1="${pad.left}" y1="${pad.top + chartH}" x2="${pad.left + chartW}" y2="${pad.top + chartH}" class="axis-line" />
      <line x1="${pad.left}" y1="${pad.top}" x2="${pad.left}" y2="${pad.top + chartH}" class="axis-line" />
      ${bars}
      <line x1="${pad.left}" y1="${avgY}" x2="${pad.left + chartW}" y2="${avgY}" class="trend-avg-line ${htmlEscape(
        config.avgTheme
      )}"><title>${htmlEscape(avgLabel)}</title></line>
      <g transform="translate(${avgCalloutX},${avgCalloutY})">
        <rect width="${avgCalloutWidth}" height="${avgCalloutHeight}" rx="${avgCalloutHeight / 2}" class="trend-avg-callout-bg ${htmlEscape(
          config.avgTheme
        )}" />
        <text x="${avgCalloutWidth / 2}" y="${avgCalloutHeight / 2}" text-anchor="middle" dominant-baseline="middle" class="trend-avg-callout-text ${htmlEscape(
          config.avgTheme
        )}">${htmlEscape(avgLabel)}</text>
      </g>
      ${labels}
      <text x="${pad.left}" y="${16}" class="legend units">${htmlEscape(config.legend)} (bars)</text>
      <text x="${width - 8}" y="${pad.top + 12}" text-anchor="end" class="axis">Scale max ${htmlEscape(config.formatValue(maxValue))}</text>
    </svg>
  `;

  container.querySelectorAll("[data-day-kpi-point-index]").forEach((bar) => {
    const activate = () => {
      const pointIndex = Math.max(0, Math.floor(num(bar.getAttribute("data-day-kpi-point-index"))));
      const point = points[pointIndex];
      if (!point) return;
      openDayVisualiserForTrendPoint(point);
    };
    bar.addEventListener("click", activate);
    bar.addEventListener("keydown", (event) => {
      if (event.key !== "Enter" && event.key !== " ") return;
      event.preventDefault();
      activate();
    });
  });
}

function renderDayKpiTrend() {
  const container = document.getElementById("stageTrendChart");
  const title = document.getElementById("trendTitle");
  const meta = document.getElementById("trendMeta");
  const metricKey = String(trendModalContext.metricKey || "").trim();
  const { config, allDates, allMonths, cursorDate, windowDates, dailyPoints, monthlyPoints } = buildDayKpiTrendSeries(metricKey);
  if (!container || !title || !meta || !config || !state) return;
  const isMonthly = state.trendGranularity === "monthly";
  const hideZeroDays = state.trendShowZeroDays === false;
  const points = isMonthly
    ? monthlyPoints
    : filterTrendWindowPoints(dailyPoints, windowDates, {
        hideZeroDays,
        zeroDayPredicate: (point) => isTrendZeroValue(point.value)
      });
  const scalePoints = isMonthly ? monthlyPoints : dailyPoints;

  title.textContent = config.title;
  meta.textContent = `Shift: ${state.selectedShift} | ${
    state.trendGranularity === "monthly"
      ? "Monthly aggregated"
      : `Rolling ${TREND_DAILY_WINDOW_SIZE}-day view ending ${formatIsoDateLabel(cursorDate, { month: "short", day: "numeric", year: "numeric" })}${hideZeroDays ? " | 0 days hidden" : ""}`
  } | ${config.description}`;
  setTrendControlsUI({ allDates, allMonths, cursorDate, windowDates });

  if (points.length < 2) {
    container.innerHTML = `<div class="empty-chart">Need at least 2 ${
      isMonthly ? "months" : hideZeroDays ? "visible non-zero dates" : "visible dates"
    } to draw a trend.</div>`;
    return;
  }

  renderDayKpiTrendChart(container, points, config, { scalePoints });
}

function renderStageTrend() {
  const container = document.getElementById("stageTrendChart");
  const title = document.getElementById("trendTitle");
  const meta = document.getElementById("trendMeta");
  const { stage, stageIndex, allDates, allMonths, cursorDate, windowDates, dailyPoints, monthlyPoints } = buildStageTrendSeries();
  const isMonthly = state.trendGranularity === "monthly";
  const hideZeroDays = state.trendShowZeroDays === false;
  const points = isMonthly
    ? monthlyPoints
    : filterTrendWindowPoints(dailyPoints, windowDates, {
        hideZeroDays,
        zeroDayPredicate: stageTrendPointIsZeroDay
      });
  const scalePoints = isMonthly ? monthlyPoints : dailyPoints;

  title.textContent = `${stageDisplayName(stage, stageIndex)} Trend`;
  meta.textContent = `Shift: ${state.selectedShift} | ${
    state.trendGranularity === "monthly"
      ? "Monthly aggregated"
      : `Rolling ${TREND_DAILY_WINDOW_SIZE}-day view ending ${formatIsoDateLabel(cursorDate, { month: "short", day: "numeric", year: "numeric" })}${hideZeroDays ? " | 0 days hidden" : ""}`
  } | Utilisation is based on ETC vs max throughput.`;
  setTrendControlsUI({ allDates, allMonths, cursorDate, windowDates });

  if (points.length < 2) {
    container.innerHTML = `<div class="empty-chart">Need at least 2 ${
      isMonthly ? "months" : hideZeroDays ? "visible non-zero dates" : "visible dates"
    } to draw a trend.</div>`;
    return;
  }

  const width = 1180;
  const height = 340;
  const pad = { top: 22, right: 88, bottom: 58, left: 62 };
  const chartW = width - pad.left - pad.right;
  const chartH = height - pad.top - pad.bottom;
  const maxUtil = Math.max(100, ...scalePoints.map((p) => p.utilisation));
  const maxDown = Math.max(1, ...scalePoints.map((p) => p.stageDowntime));
  const maxEtc = Math.max(1, ...scalePoints.map((p) => p.stageEtc));
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

function renderTrendModalContent() {
  if (String(trendModalContext.type || "").trim() === "dayKpi") {
    renderDayKpiTrend();
    return;
  }
  renderStageTrend();
}

function exportTrendModalCsv() {
  if (!state) return;
  const slug = (value) =>
    String(value || "")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "")
      || "trend";

  if (String(trendModalContext.type || "").trim() === "dayKpi") {
    const metricKey = String(trendModalContext.metricKey || "").trim();
    const { config, dailyPoints, monthlyPoints } = buildDayKpiTrendSeries(metricKey);
    if (!config || (!dailyPoints.length && !monthlyPoints.length)) {
      alert("No trend data available to export.");
      return;
    }
    const rows = [
      ...dailyPoints.map((point) => ({
        Line: state.name || "Production Line",
        Shift: state.selectedShift || "Day",
        TrendType: "Day KPI",
        Metric: config.legend,
        Granularity: "Daily",
        PeriodKey: point.date,
        PeriodLabel: formatIsoDateLabel(point.date, { month: "short", day: "numeric", year: "numeric" }),
        Value: roundToDecimals(num(point.value), 4),
        DisplayValue: config.formatValue(point.value)
      })),
      ...monthlyPoints.map((point) => ({
        Line: state.name || "Production Line",
        Shift: state.selectedShift || "Day",
        TrendType: "Day KPI",
        Metric: config.legend,
        Granularity: "Monthly",
        PeriodKey: point.date,
        PeriodLabel: formatMonthLabel(point.date),
        Value: roundToDecimals(num(point.value), 4),
        DisplayValue: config.formatValue(point.value)
      }))
    ];
    const columns = ["Line", "Shift", "TrendType", "Metric", "Granularity", "PeriodKey", "PeriodLabel", "Value", "DisplayValue"];
    downloadTextFile(
      `${slug(state.name)}-${slug(config.title)}-trend.csv`,
      toCsv(rows, columns),
      "text/csv;charset=utf-8"
    );
    addAudit(state, "EXPORT_TREND_CSV", `${config.title} trend CSV exported`);
    saveState();
    return;
  }

  const { stage, stageIndex, dailyPoints, monthlyPoints } = buildStageTrendSeries();
  if (!stage || (!dailyPoints.length && !monthlyPoints.length)) {
    alert("No trend data available to export.");
    return;
  }
  const stageLabel = stageDisplayName(stage, stageIndex);
  const rows = [
    ...dailyPoints.map((point) => ({
      Line: state.name || "Production Line",
      Shift: state.selectedShift || "Day",
      TrendType: "Stage",
      Stage: stageLabel,
      Granularity: "Daily",
      PeriodKey: point.date,
      PeriodLabel: formatIsoDateLabel(point.date, { month: "short", day: "numeric", year: "numeric" }),
      UtilisationPct: roundToDecimals(num(point.utilisation), 4),
      DowntimeMin: roundToDecimals(num(point.stageDowntime), 4),
      EtcUnitsPerMin: roundToDecimals(num(point.stageEtc), 4)
    })),
    ...monthlyPoints.map((point) => ({
      Line: state.name || "Production Line",
      Shift: state.selectedShift || "Day",
      TrendType: "Stage",
      Stage: stageLabel,
      Granularity: "Monthly",
      PeriodKey: point.date,
      PeriodLabel: formatMonthLabel(point.date),
      UtilisationPct: roundToDecimals(num(point.utilisation), 4),
      DowntimeMin: roundToDecimals(num(point.stageDowntime), 4),
      EtcUnitsPerMin: roundToDecimals(num(point.stageEtc), 4)
    }))
  ];
  const columns = ["Line", "Shift", "TrendType", "Stage", "Granularity", "PeriodKey", "PeriodLabel", "UtilisationPct", "DowntimeMin", "EtcUnitsPerMin"];
  downloadTextFile(
    `${slug(state.name)}-${slug(stageLabel)}-stage-trend.csv`,
    toCsv(rows, columns),
    "text/csv;charset=utf-8"
  );
  addAudit(state, "EXPORT_TREND_CSV", `${stageLabel} trend CSV exported`);
  saveState();
}

function setTrendControlsUI({ allDates = [], allMonths = [], cursorDate = "", windowDates = [] } = {}) {
  const dailyBtn = document.getElementById("trendDaily");
  const monthlyBtn = document.getElementById("trendMonthly");
  const zeroDaysToggleBtn = document.getElementById("trendZeroDaysToggle");
  const label = document.getElementById("trendMonthLabel");
  const prevBtn = document.getElementById("trendPrevMonth");
  const nextBtn = document.getElementById("trendNextMonth");
  const activeMonth = state.trendMonth || monthKey(state.selectedDate);
  const zeroDaysHidden = state.trendShowZeroDays === false;

  dailyBtn.classList.toggle("active", state.trendGranularity === "daily");
  monthlyBtn.classList.toggle("active", state.trendGranularity === "monthly");
  if (zeroDaysToggleBtn) {
    zeroDaysToggleBtn.textContent = `${zeroDaysHidden ? "Show" : "Hide"} 0 Days`;
    zeroDaysToggleBtn.classList.toggle("active", zeroDaysHidden);
    zeroDaysToggleBtn.setAttribute("aria-pressed", String(zeroDaysHidden));
  }

  if (state.trendGranularity === "daily") {
    const safeCursor = cursorDate || syncTrendDateCursorToAvailableDates(allDates);
    const visibleDates = windowDates.length ? windowDates : trendDailyWindowDates(allDates, safeCursor);
    label.textContent = visibleDates.length
      ? lineTrendBucketDateRangeLabel({ startIso: visibleDates[0], endIso: visibleDates[visibleDates.length - 1] })
      : formatIsoDateLabel(safeCursor, { month: "short", day: "numeric", year: "numeric" });
    if (!allDates.length) {
      prevBtn.disabled = true;
      nextBtn.disabled = true;
      return;
    }
    const currentIndex = allDates.indexOf(safeCursor);
    prevBtn.disabled = currentIndex <= 0;
    nextBtn.disabled = currentIndex < 0 || currentIndex >= allDates.length - 1;
    return;
  }

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
  const inlineSelectHtml = (field, value, options = [], { placeholder = "", disabled = false } = {}) => {
    const normalized = Array.isArray(options)
      ? options.map((option) => {
          if (typeof option === "string") return { value: option, label: option };
          return {
            value: String(option?.value || ""),
            label: String(option?.label ?? option?.value ?? "")
          };
        })
      : [];
    const selected = String(value ?? "");
    if (selected && !normalized.some((option) => option.value === selected)) {
      normalized.unshift({ value: selected, label: selected });
    }
    const placeholderHtml = placeholder ? `<option value="">${htmlEscape(placeholder)}</option>` : "";
    const optionsHtml = normalized
      .map(
        (option) =>
          `<option value="${htmlEscape(option.value)}"${option.value === selected ? " selected" : ""}>${htmlEscape(option.label)}</option>`
      )
      .join("");
    return `<select class="table-inline-input" data-inline-field="${htmlEscape(field)}"${disabled ? " disabled" : ""}>${placeholderHtml}${optionsHtml}</select>`;
  };
  const isInlineEditing = (type, rowId) => isManagerLogInlineEditRow(activeLineId, type, rowId);
  const actionHtml = (type, row, { editing = false, pending = false } = {}) => {
    const id = htmlEscape(row.id || "");
    const typeAttr = htmlEscape(type);
    return `
      <div class="table-action-stack">
        <button type="button" class="table-edit-pill${editing ? " is-save" : ""}" data-log-action="${
          editing ? "save" : "edit"
        }" data-log-type="${typeAttr}" data-log-id="${id}">${editing ? "Save" : "Edit"}</button>
        <button type="button" class="table-edit-pill table-delete-pill" data-log-action="delete" data-log-type="${typeAttr}" data-log-id="${id}">Delete</button>
        ${!editing && pending ? `<button type="button" class="table-edit-pill finalise-pill" data-log-action="finalise" data-log-type="${typeAttr}" data-log-id="${id}">Finalise</button>` : ""}
      </div>
    `;
  };

  const sortNewestFirst = (rows, primaryTimeField) =>
    (rows || [])
      .slice()
      .sort((a, b) => {
        const dateCmp = String(b?.date || "").localeCompare(String(a?.date || ""));
        if (dateCmp !== 0) return dateCmp;
        const tsDiff = rowNewestSortValue(b, primaryTimeField) - rowNewestSortValue(a, primaryTimeField);
        if (tsDiff !== 0) return tsDiff;
        const submittedCmp = String(b?.submittedAt || "").localeCompare(String(a?.submittedAt || ""));
        if (submittedCmp !== 0) return submittedCmp;
        return String(b?.id || "").localeCompare(String(a?.id || ""));
      });
  const groupedLogTableOptions = {
    groupByField: "date",
    hideGroupedFieldCells: true,
    groupLabelFormatter: (value) => (
      isIsoDateValue(value)
        ? formatIsoDateLabel(value, { weekday: "long", day: "numeric", month: "long", year: "numeric" })
        : (String(value || "").trim() || "Unknown Date")
    )
  };

  const derivedBreakContext = { breakRows: data.breakRows };
  const displayShiftRows = sortNewestFirst(data.shiftRows, "startTime").map((row) => {
    const rowBreakRows = breakRowsForShift(derivedBreakContext, row);
    const breakCount = Math.max(0, Math.floor(num(rowBreakRows.length)));
    const breakTimeMins = rowBreakRows.reduce((sum, breakRow) => sum + Math.max(0, num(breakRow.breakMins)), 0);
    const editing = isInlineEditing("shift", String(row.id || ""));
    const shiftPending = isPendingShiftLogRow(row);
    const pending = shiftPending;
    const htmlFields = ["action"];
    if (editing) htmlFields.push("date", "shift", "startTime", "finishTime", "notes");
    return {
      ...row,
      __rowClass: pending ? "table-row-pending" : "",
      __htmlFields: htmlFields,
      date: editing ? inlineInputHtml("date", String(row.date || ""), { type: "date" }) : row.date,
      shift: editing
        ? inlineSelectHtml(
          "shift",
          String(row.shift || ""),
          SHIFT_OPTIONS.map((shiftOption) => ({ value: shiftOption, label: shiftOption }))
        )
        : row.shift,
      startTime: editing ? inlineInputHtml("startTime", String(row.startTime || ""), { placeholder: "HH:MM" }) : row.startTime,
      finishTime: editing ? inlineInputHtml("finishTime", String(row.finishTime || ""), { placeholder: "HH:MM" }) : row.finishTime,
      breakCount,
      breakTimeMins,
      notes: editing ? inlineInputHtml("notes", String(row.notes || ""), { placeholder: "Notes" }) : String(row.notes || ""),
      submittedBy: managerLogSubmittedByLabel(row),
      action: actionHtml("shift", row, { editing, pending: shiftPending })
    };
  });

  renderTable("shiftTable", SHIFT_COLUMNS, displayShiftRows, {
    Date: "date",
    Shift: "shift",
    "Start Time": "startTime",
    "Finish Time": "finishTime",
    "Break Count": "breakCount",
    "Break Time (min)": "breakTimeMins",
    "Total Shift Time": "totalShiftTime",
    Notes: "notes",
    "Submitted By": "submittedBy",
    Action: "action"
  }, groupedLogTableOptions);

  const displayRunRows = sortNewestFirst(data.runRows, "productionStartTime").map((row) => {
    const editing = isInlineEditing("run", String(row.id || ""));
    const pending = isPendingRunLogRow(row);
    const htmlFields = ["action"];
    if (editing) htmlFields.push("date", "product", "productionStartTime", "finishTime", "unitsProduced", "notes");
    const patternShift = preferredTimedLogShift(
      state,
      String(row.date || todayISO()),
      String(row.productionStartTime || ""),
      String(row.finishTime || ""),
      state.selectedShift || "Day"
    );
    const runPattern = normalizeRunCrewingPattern(row.runCrewingPattern, state, patternShift, { fallbackToIdeal: false });
    const runPatternRaw = Object.keys(runPattern).length ? JSON.stringify(runPattern) : "";
    const runPatternSummary = runCrewingPatternSummaryText(runPattern, state);
    return {
      ...row,
      __rowClass: pending ? "table-row-pending" : "",
      __htmlFields: htmlFields,
      date: editing ? inlineInputHtml("date", String(row.date || ""), { type: "date" }) : row.date,
      product: editing
        ? `
          <div class="table-inline-stack">
            ${inlineInputHtml("product", String(row.product || ""))}
            <input type="hidden" data-inline-field="runCrewingPattern" value="${htmlEscape(runPatternRaw)}" />
            <div class="table-inline-row">
              <span class="table-inline-note" data-inline-run-pattern-summary>${htmlEscape(runPatternSummary)}</span>
              <button type="button" class="table-edit-pill table-inline-mini" data-inline-run-pattern data-pattern-shift="${htmlEscape(
                patternShift
              )}">Crewing Pattern</button>
            </div>
          </div>
        `
        : row.product,
      productionStartTime: editing
        ? inlineInputHtml("productionStartTime", String(row.productionStartTime || ""), { placeholder: "HH:MM" })
        : row.productionStartTime,
      finishTime: editing ? inlineInputHtml("finishTime", String(row.finishTime || ""), { placeholder: "HH:MM" }) : row.finishTime,
      unitsProduced: editing ? inlineInputHtml("unitsProduced", Math.max(0, num(row.unitsProduced)), { type: "number", min: "0", step: "1" }) : row.unitsProduced,
      notes: editing ? inlineInputHtml("notes", String(row.notes || ""), { placeholder: "Notes" }) : String(row.notes || ""),
      submittedBy: managerLogSubmittedByLabel(row),
      action: actionHtml("run", row, { editing, pending })
    };
  });
  renderTable("runTable", RUN_COLUMNS, displayRunRows, {
    Date: "date",
    Product: "product",
    "Production Start Time": "productionStartTime",
    "Finish Time": "finishTime",
    "Units Produced": "unitsProduced",
    "Gross Production Time": "grossProductionTime",
    "Associated Down Time": "associatedDownTime",
    "Net Production Time": "netProductionTime",
    "Gross Run Rate": "grossRunRate",
    "Net Run Rate": "netRunRate",
    Notes: "notes",
    "Submitted By": "submittedBy",
    Action: "action"
  }, groupedLogTableOptions);

  const downtimeRowsForTable = Array.isArray(data.downtimeRowsLogged) ? data.downtimeRowsLogged : data.downtimeRows;
  const displayDowntimeRows = sortNewestFirst(downtimeRowsForTable, "downtimeStart").map((row) => {
    const editing = isInlineEditing("downtime", String(row.id || ""));
    const pending = isPendingDowntimeLogRow(row);
    const htmlFields = ["action"];
    if (editing) htmlFields.push("date", "downtimeStart", "downtimeFinish", "equipment", "reason", "notes");
    const parsedReason = parseDowntimeReasonParts(row.reason, row.equipment);
    const reasonCategory = String(row.reasonCategory || parsedReason.reasonCategory || "");
    const reasonDetail = String(row.reasonDetail || parsedReason.reasonDetail || "");
    const reasonNote = String((row.reasonNote ?? parsedReason.reasonNote) || "");
    const downtimeCategories = Array.from(new Set(["Equipment", ...Object.keys(DOWNTIME_REASON_PRESETS), reasonCategory])).filter(Boolean);
    const reasonDetailChoices = downtimeDetailOptions(state, reasonCategory);
    const equipmentChoices = downtimeDetailOptions(state, "Equipment");
    const equipmentValue = String(row.equipment || (reasonCategory === "Equipment" ? reasonDetail : ""));
    return {
      ...row,
      __rowClass: pending ? "table-row-pending" : "",
      __htmlFields: htmlFields,
      date: editing ? inlineInputHtml("date", String(row.date || ""), { type: "date" }) : row.date,
      downtimeStart: editing ? inlineInputHtml("downtimeStart", String(row.downtimeStart || ""), { placeholder: "HH:MM" }) : row.downtimeStart,
      downtimeFinish: editing ? inlineInputHtml("downtimeFinish", String(row.downtimeFinish || ""), { placeholder: "HH:MM" }) : row.downtimeFinish,
      equipment: editing
        ? inlineSelectHtml("equipment", equipmentValue, equipmentChoices, { placeholder: "Select Stage", disabled: reasonCategory !== "Equipment" })
        : row.equipment
          ? stageNameById(row.equipment)
          : "-",
      reason: editing
        ? `
          <div class="table-inline-stack">
            ${inlineSelectHtml(
              "reasonCategory",
              reasonCategory,
              downtimeCategories.map((category) => ({ value: category, label: category })),
              { placeholder: "Reason Group" }
            )}
            ${inlineSelectHtml(
              "reasonDetail",
              reasonDetail,
              reasonDetailChoices.map((detailOption) => ({ value: detailOption.value, label: detailOption.label })),
              { placeholder: reasonCategory === "Equipment" ? "Select Stage" : "Select Reason" }
            )}
            ${inlineInputHtml("reasonNote", reasonNote, { placeholder: "Optional note" })}
          </div>
        `
        : row.reason,
      notes: editing ? inlineInputHtml("notes", String(row.notes || ""), { placeholder: "Notes" }) : String(row.notes || ""),
      submittedBy: managerLogSubmittedByLabel(row),
      action: actionHtml("downtime", row, { editing, pending })
    };
  });
  renderTable("downtimeTable", DOWN_COLUMNS, displayDowntimeRows, {
    Date: "date",
    "Downtime Start": "downtimeStart",
    "Downtime Finish": "downtimeFinish",
    "Downtime (mins)": "downtimeMins",
    Equipment: "equipment",
    Reason: "reason",
    Notes: "notes",
    "Submitted By": "submittedBy",
    Action: "action"
  }, groupedLogTableOptions);
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
  const liveDowntimeByStage = liveEquipmentDowntimeByStage(line, stages);
  const data = derivedDataForLine(line || {});
  const selectedRunRows = selectedShiftRowsByDate(data.runRows, selectedDate, selectedShift, { line });
  const selectedDownRows = selectedShiftRowsByDate(data.downtimeRows, selectedDate, selectedShift, { line });
  const selectedCalcDownRows = calculationDowntimeRows(selectedDownRows);
  const selectedShiftRows = selectedShiftRowsByDate(data.shiftRows, selectedDate, selectedShift, { line });
  const shiftMins = selectedShiftRows.reduce((sum, row) => sum + num(row.totalShiftTime), 0);
  const units = selectedRunRows.reduce((sum, row) => sum + num(row.unitsProduced) * timedLogShiftWeight(row, selectedShift), 0);
  const totalDowntime = selectedDownRows.reduce((sum, row) => sum + num(row.downtimeMins) * timedLogShiftWeight(row, selectedShift), 0);
  const totalNetTime = selectedRunRows.reduce((sum, row) => sum + num(row.netProductionTime) * timedLogShiftWeight(row, selectedShift), 0);
  const netRunRate = totalNetTime > 0 ? units / totalNetTime : 0;
  const selectedBottleneck = resolveLineBottleneckStage(line, stages, selectedShift, (stageId, selectedShiftValue) =>
    stageTotalMaxThroughputForLine(line, stageId, selectedShiftValue)
  );
  let bottleneckCard = null;
  let bottleneckUtil = 0;
  let bottleneckUtilGross = 0;
  const activeCrew = crewMapForLineShift(line, selectedShift, stages);
  const liveRunPattern = liveRunCrewingPatternForLine(line);

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
  const isFlowMoving = lineHasMovingFlow(line);
  map.classList.toggle("flow-running", isFlowMoving);
  map.classList.toggle("flow-static", !isFlowMoving);
  const guides = lineFlowGuidesForMap(stages, line?.flowGuides);
  appendFlowGuidesToMap(map, guides, { editable: false });

  stages.forEach((stage, index) => {
    const stageDowntime = selectedCalcDownRows
      .filter((row) => matchesStage(stage, row.equipment))
      .reduce((sum, row) => sum + downtimeMinutesForCalculations(row, selectedShift), 0);
    const uptimeRatio = shiftMins > 0 ? Math.max(0, (shiftMins - stageDowntime) / shiftMins) : 0;
    const stageRate = netRunRate * uptimeRatio;
    const totalMax = stageTotalMaxThroughputForLine(line, stage.id, selectedShift);
    const utilisation = totalMax > 0 ? (stageRate / totalMax) * 100 : 0;
    const grossUtilisation = totalMax > 0 ? (netRunRate / totalMax) * 100 : 0;
    const stageCrew = stageCrewCountForVisual(line, stage, activeCrew, liveRunPattern);
    const compact = stage.w * stage.h < 140;
    const status = statusClass(utilisation);
    const card = document.createElement("article");
    card.className = `stage-card group-${stage.group}${stage.kind ? ` kind-${stage.kind}` : ""} status-${status}`;
    const liveDown = liveDowntimeByStage[stage.id];
    const liveDownPill = stageLiveDowntimePillHtml(liveDown);
    if (liveDown?.level === "critical") card.classList.add("live-downtime-critical");
    if (liveDown?.level === "warning") card.classList.add("live-downtime-warning");
    card.style.left = `${stage.x}%`;
    card.style.top = `${stage.y}%`;
    card.style.width = `${stage.w}%`;
    card.style.height = `${stage.h}%`;
    card.classList.toggle("compact", compact);

    card.innerHTML = `
      <h3 class="stage-title">${stageDisplayName(stage, index)}</h3>
      <div class="stage-pill-row">
        ${liveDownPill}
        <span class="stage-crew-pill">${formatNum(stageCrew, 0)} Crew</span>
      </div>
    `;
    if (liveDown) {
      card.setAttribute(
        "title",
        liveDown.level === "critical"
          ? `Urgent downtime: ${Math.floor(liveDown.minutes)} min`
          : `Active downtime: ${Math.floor(liveDown.minutes)} min`
      );
    }

    const isSelectedBottleneck = Boolean(selectedBottleneck) && selectedBottleneck.stage.id === stage.id;
    if (isSelectedBottleneck) {
      bottleneckUtil = Math.max(0, utilisation);
      bottleneckUtilGross = Math.max(0, grossUtilisation);
      bottleneckCard = card;
    }
    map.append(card);
  });

  if (bottleneckCard) {
    bottleneckCard.classList.add("bottleneck");
  }
  const lineUtil = Math.max(0, bottleneckUtilGross);
  const lineUtilGross = Math.max(0, bottleneckUtil);
  const svUtilText = `${formatNum(Math.max(lineUtil, 0), 1)}%`;
  const svUtilGrossText = `${formatNum(Math.max(lineUtilGross, 0), 1)}%`;
  setSvText("svKpiUtilisation", svUtilText);
  setSvText("svDayKpiUtilisation", svUtilText);
  setSvText("svKpiUtilisationGross", svUtilGrossText);
  setSvText("svDayKpiUtilisationGross", svUtilGrossText);
}

function renderHome() {
  const homeTitle = document.getElementById("homeTitle");
  const sidebarBackdrop = document.getElementById("sidebarBackdrop");
  const homeSidebar = document.getElementById("homeSidebar");
  const homeResetPasswordBtn = document.getElementById("homeResetPasswordBtn");
  const headerLogoutBtn = document.getElementById("supervisorLogout");
  const managerSettingsTabBtn = document.getElementById("managerSettingsTabBtn");
  const homeUserChip = document.getElementById("homeUserChip");
  const homeUserAvatar = document.getElementById("homeUserAvatar");
  const homeUserRole = document.getElementById("homeUserRole");
  const homeUserIdentity = document.getElementById("homeUserIdentity");
  const lineWorkspaceUserChip = document.getElementById("lineWorkspaceUserChip");
  const lineWorkspaceUserAvatar = document.getElementById("lineWorkspaceUserAvatar");
  const lineWorkspaceUserRole = document.getElementById("lineWorkspaceUserRole");
  const lineWorkspaceUserIdentity = document.getElementById("lineWorkspaceUserIdentity");
  const lineWorkspaceResetPasswordBtn = document.getElementById("lineWorkspaceResetPasswordBtn");
  const managerHome = document.getElementById("managerHome");
  const supervisorHome = document.getElementById("supervisorHome");
  const modeManagerBtn = document.getElementById("modeManager");
  const modeSupervisorBtn = document.getElementById("modeSupervisor");
  const lineModeManagerBtn = document.getElementById("lineModeManager");
  const lineModeSupervisorBtn = document.getElementById("lineModeSupervisor");
  const homeModeToggle = modeManagerBtn?.closest(".shift-toggle");
  const lineModeToggle = lineModeManagerBtn?.closest(".shift-toggle");
  const managerLoginSection = document.getElementById("managerLoginSection");
  const managerAppSection = document.getElementById("managerAppSection");
  const dashboardDateInput = document.getElementById("dashboardDate");
  const dashboardShiftButtons = Array.from(document.querySelectorAll("[data-dash-shift]"));
  const dashboardTable = document.getElementById("dashboardTable");
  const dataSourcesList = document.getElementById("dataSourcesList");
  const incomingDataSourcesStatus = document.getElementById("incomingDataSourcesStatus");
  const managerActionList = document.getElementById("managerActionList");
  const loginSection = document.getElementById("supervisorLoginSection");
  const appSection = document.getElementById("supervisorAppSection");
  const supervisorMobileModeBtn = document.getElementById("supervisorMobileMode");
  const lineSelect = document.getElementById("supervisorLineSelect");
  const productCatalogTable = document.getElementById("productCatalogTable");
  const svDateInputs = Array.from(document.querySelectorAll("[data-sv-date]"));
  const svShiftButtons = Array.from(document.querySelectorAll("[data-sv-shift]"));
  const shiftDateInput = document.getElementById("superShiftDate");
  const shiftShiftInput = document.getElementById("superShiftShift");
  const shiftLogIdInput = document.getElementById("superShiftLogId");
  const runDateInput = document.getElementById("superRunDate");
  const runLogIdInput = document.getElementById("superRunLogId");
  const downDateInput = document.getElementById("superDownDate");
  const downReasonCategoryInput = document.getElementById("superDownReasonCategory");
  const downReasonDetailInput = document.getElementById("superDownReasonDetail");
  const downLogIdInput = document.getElementById("superDownLogId");
  const entryList = document.getElementById("supervisorEntryList");
  const entryCards = document.getElementById("supervisorEntryCards");
  const actionLineInput = document.getElementById("supervisorActionLine");
  const actionPriorityInput = document.getElementById("supervisorActionPriority");
  const actionStatusInput = document.getElementById("supervisorActionStatus");
  const actionDueDateInput = document.getElementById("supervisorActionDueDate");
  const actionEquipmentInput = document.getElementById("supervisorActionEquipment");
  const actionReasonCategoryInput = document.getElementById("supervisorActionReasonCategory");
  const actionReasonDetailInput = document.getElementById("supervisorActionReasonDetail");
  const actionList = document.getElementById("supervisorActionList");
  const actionCards = document.getElementById("supervisorActionCards");
  const superMainTabBtns = Array.from(document.querySelectorAll("[data-super-main-tab]"));
  const superMainPanels = Array.from(document.querySelectorAll(".supervisor-main-panel"));
  const session = normalizeSupervisorSession(appState.supervisorSession, appState.supervisors, appState.lines);
  appState.supervisorSession = session;
  const isSupervisor = appState.appMode === "supervisor";
  const managerSessionActive = Boolean(managerBackendSession.backendToken);
  syncLineTileFeedbackLoop({ enabled: !isSupervisor && managerSessionActive });
  const activeSupervisor = session?.username ? supervisorByUsername(session.username) : null;

  if (homeTitle) {
    homeTitle.textContent = "Production Line Dashboard";
  }
  if (headerLogoutBtn) {
    const showLogout = isSupervisor ? Boolean(session) : managerSessionActive;
    headerLogoutBtn.classList.toggle("hidden", !showLogout);
    headerLogoutBtn.textContent = "Logout";
  }
  if (managerSettingsTabBtn) {
    const showManagerSettingsTab = !isSupervisor && managerSessionActive;
    const managerSettingsActive = showManagerSettingsTab && appState.managerHomeTab === "settings";
    managerSettingsTabBtn.classList.toggle("hidden", !showManagerSettingsTab);
    managerSettingsTabBtn.classList.toggle("active", managerSettingsActive);
    managerSettingsTabBtn.setAttribute("aria-pressed", String(managerSettingsActive));
  }
  const showSupervisorTile = isSupervisor && Boolean(session);
  const showManagerTile = !isSupervisor && managerSessionActive;
  if (homeResetPasswordBtn) {
    homeResetPasswordBtn.classList.toggle("hidden", !(showManagerTile || showSupervisorTile));
  }
  const setUserChip = (roleEl, identityEl, avatarEl, label, username, fallbackAvatar = "M") => {
    const safeLabel = String(label || "").trim() || "Manager";
    const safeUsername = String(username || "").trim() || "manager";
    if (roleEl) roleEl.textContent = safeLabel;
    if (identityEl) identityEl.textContent = safeUsername;
    if (avatarEl) avatarEl.textContent = (safeLabel.charAt(0) || fallbackAvatar).toUpperCase();
  };
  if (homeUserChip) {
    homeUserChip.classList.toggle("hidden", !(showManagerTile || showSupervisorTile));
  }
  if (lineWorkspaceUserChip) {
    lineWorkspaceUserChip.classList.toggle("hidden", !showManagerTile);
  }
  if (lineWorkspaceResetPasswordBtn) {
    lineWorkspaceResetPasswordBtn.classList.toggle("hidden", !showManagerTile);
  }
  if (showSupervisorTile) {
    const supervisorLabel = String(activeSupervisor?.name || session?.name || session?.username || "Supervisor").trim() || "Supervisor";
    const supervisorUsername = String(session?.username || "").trim();
    setUserChip(homeUserRole, homeUserIdentity, homeUserAvatar, supervisorLabel, supervisorUsername, "S");
  } else {
    const managerLabel = String(managerBackendSession?.name || managerBackendSession?.username || "Manager").trim() || "Manager";
    const managerUsername = String(managerBackendSession?.username || "").trim().toLowerCase() || "manager";
    setUserChip(homeUserRole, homeUserIdentity, homeUserAvatar, managerLabel, managerUsername, "M");
    setUserChip(lineWorkspaceUserRole, lineWorkspaceUserIdentity, lineWorkspaceUserAvatar, managerLabel, managerUsername, "M");
  }

  managerHome.classList.toggle("hidden", isSupervisor);
  supervisorHome.classList.toggle("hidden", !isSupervisor);
  if (managerLoginSection) managerLoginSection.classList.toggle("hidden", isSupervisor || managerSessionActive);
  if (managerAppSection) managerAppSection.classList.toggle("hidden", isSupervisor || !managerSessionActive);
  if (managerAppSection) managerAppSection.classList.toggle("is-settings-view", !isSupervisor && managerSessionActive && appState.managerHomeTab === "settings");
  modeManagerBtn.classList.toggle("active", !isSupervisor);
  modeSupervisorBtn.classList.toggle("active", isSupervisor);
  modeManagerBtn.setAttribute("aria-pressed", String(!isSupervisor));
  modeSupervisorBtn.setAttribute("aria-pressed", String(isSupervisor));
  if (lineModeManagerBtn) lineModeManagerBtn.classList.toggle("active", !isSupervisor);
  if (lineModeSupervisorBtn) lineModeSupervisorBtn.classList.toggle("active", isSupervisor);
  if (lineModeManagerBtn) lineModeManagerBtn.setAttribute("aria-pressed", String(!isSupervisor));
  if (lineModeSupervisorBtn) lineModeSupervisorBtn.setAttribute("aria-pressed", String(isSupervisor));
  if (homeModeToggle) homeModeToggle.classList.add("hidden");
  if (lineModeToggle) lineModeToggle.classList.add("hidden");
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
  appState.dashboardDate = normalizeWeekdayIsoDate(appState.dashboardDate || todayISO(), { direction: -1 });
  dashboardDateInput.value = appState.dashboardDate;
  dashboardShiftButtons.forEach((btn) => {
    const active = btn.dataset.dashShift === (appState.dashboardShift || "Day");
    btn.classList.toggle("active", active);
    btn.setAttribute("aria-pressed", String(active));
  });
  const dashboardRows = lineList.map((line) => computeLineMetrics(line, appState.dashboardDate, appState.dashboardShift || "Day"));
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
    cards.classList.remove("line-cards-grouped");
    cards.innerHTML = `
      <article class="line-card line-card-unavailable" role="status">
        <h3>Database not available</h3>
        <p>We could not load production lines because the database is currently unavailable.</p>
      </article>
    `;
  } else if (!lineList.length) {
    cards.classList.remove("line-cards-grouped");
    cards.innerHTML = `
      <article class="line-card line-card-empty" role="status">
        <h3>No Production Lines</h3>
        <p>No production lines are currently available.</p>
      </article>
    `;
  } else {
    const groupedCards = renderGroupedHomeLineCards(lineList, appState.lineGroups);
    cards.classList.toggle("line-cards-grouped", groupedCards.grouped);
    cards.innerHTML = groupedCards.html;
  }
  renderLineShiftTrackersForWidth();
  applyLineTileLiveFeedback();
  if (statTotalLines) statTotalLines.textContent = formatNum(lineList.length, 0);
  if (statSupervisors) statSupervisors.textContent = formatNum((appState.supervisors || []).length, 0);
  if (statShiftRecords) {
    statShiftRecords.textContent = formatNum(
      lineList.reduce((sum, line) => sum + (line.shiftRows || []).filter((row) => isOperationalDate(String(row?.date || ""))).length, 0),
      0
    );
  }
  if (statRunRecords) {
    statRunRecords.textContent = formatNum(
      lineList.reduce((sum, line) => sum + (line.runRows || []).filter((row) => isOperationalDate(String(row?.date || ""))).length, 0),
      0
    );
  }

  if (!isSupervisor) {
    if (homeSidebar) homeSidebar.classList.remove("open");
    if (sidebarBackdrop) sidebarBackdrop.classList.add("hidden");
  }

  if (productCatalogTable) {
    refreshProductRowAssignmentSummaries(productCatalogTable);
  }
  syncRunProductInputsFromCatalog();

  if (!isSupervisor) {
    renderManagerDataSourcesList(dataSourcesList);
    renderIncomingDataSourcesStatus(incomingDataSourcesStatus);
    if (managerActionList) {
      const managerActions = ensureSupervisorActionsState()
        .slice()
        .sort((a, b) => {
          const statusDelta = actionStatusSortRank(a?.status) - actionStatusSortRank(b?.status);
          if (statusDelta !== 0) return statusDelta;
          const dueA = /^\d{4}-\d{2}-\d{2}$/.test(String(a?.dueDate || "").trim()) ? String(a.dueDate).trim() : "9999-12-31";
          const dueB = /^\d{4}-\d{2}-\d{2}$/.test(String(b?.dueDate || "").trim()) ? String(b.dueDate).trim() : "9999-12-31";
          const dueDelta = dueA.localeCompare(dueB);
          if (dueDelta !== 0) return dueDelta;
          const createdDelta = String(b?.createdAt || "").localeCompare(String(a?.createdAt || ""));
          if (createdDelta !== 0) return createdDelta;
          return String(a?.title || "").localeCompare(String(b?.title || ""), undefined, { sensitivity: "base" });
        })
        .slice(0, 400);
      if (managerActionTicketEditId && !managerActions.some((action) => String(action?.id || "") === managerActionTicketEditId)) {
        clearManagerActionTicketEdit();
      }
      const supervisors = (appState.supervisors || [])
        .slice()
        .sort((a, b) =>
          String(a?.name || a?.username || "").localeCompare(String(b?.name || b?.username || ""), undefined, { sensitivity: "base" })
        );
      const assignmentOptions = [];
      const assignmentSeen = new Set();
      const pushAssignmentOption = (value, label) => {
        const key = String(value || "").trim().toLowerCase();
        if (assignmentSeen.has(key)) return;
        assignmentSeen.add(key);
        assignmentOptions.push({
          value: key,
          label: String(label || key || "Unassigned").trim() || "Unassigned"
        });
      };
      pushAssignmentOption("", "Unassigned");
      ACTION_SPECIAL_ASSIGNMENTS.forEach((assignment) => {
        pushAssignmentOption(assignment.username, assignment.label);
      });
      supervisors.forEach((sup) => {
        pushAssignmentOption(sup.username, sup.name || sup.username);
      });
      const supervisorOptionsHtml = (selectedUsername, selectedLabel = "") => {
        const safeSelected = String(selectedUsername || "").trim().toLowerCase();
        const hasSelected = assignmentOptions.some((option) => option.value === safeSelected);
        const fallbackLabel = actionAssignmentLabel(safeSelected, selectedLabel);
        const fallbackOption = safeSelected && !hasSelected
          ? [`<option value="${htmlEscape(safeSelected)}" selected>${htmlEscape(fallbackLabel)}</option>`]
          : [];
        return [
          ...assignmentOptions.map(
            (option) =>
              `<option value="${htmlEscape(option.value)}"${option.value === safeSelected ? " selected" : ""}>${htmlEscape(option.label)}</option>`
          ),
          ...fallbackOption,
        ].join("");
      };
      const actionPriorityOptionsHtml = (selectedPriority) => {
        const safeSelected = normalizeActionPriority(selectedPriority);
        return ACTION_PRIORITY_OPTIONS.map(
          (priority) => `<option value="${htmlEscape(priority)}"${priority === safeSelected ? " selected" : ""}>${htmlEscape(priority)}</option>`
        ).join("");
      };
      const actionStatusOptionsHtml = (selectedStatus) => {
        const safeSelected = normalizeActionStatus(selectedStatus);
        return ACTION_STATUS_OPTIONS.map(
          (status) => `<option value="${htmlEscape(status)}"${status === safeSelected ? " selected" : ""}>${htmlEscape(status)}</option>`
        ).join("");
      };
      const actionEquipmentOptionsHtml = (line, selectedEquipmentId = "") => {
        const safeSelected = String(selectedEquipmentId || "").trim();
        const options = (line?.stages || []).map((stage, index) => ({
          value: String(stage?.id || "").trim(),
          label: stageDisplayName(stage, index)
        }));
        const hasSelected = options.some((option) => option.value === safeSelected);
        const fallbackLabel = safeSelected ? stageNameByIdForLine(line, safeSelected) || safeSelected : "";
        const fallbackOption = safeSelected && !hasSelected
          ? [`<option value="${htmlEscape(safeSelected)}" selected>${htmlEscape(fallbackLabel)}</option>`]
          : [];
        return [
          `<option value="">Related Equipment (optional)</option>`,
          ...fallbackOption,
          ...options.map(
            (option) =>
              `<option value="${htmlEscape(option.value)}"${option.value === safeSelected ? " selected" : ""}>${htmlEscape(option.label)}</option>`
          )
        ].join("");
      };
      const actionReasonCategoryOptionsHtml = (selectedCategory = "") => {
        const safeSelected = normalizeActionReasonCategory(selectedCategory);
        return [
          `<option value="">Downtime Category (optional)</option>`,
          ...ACTION_REASON_CATEGORIES.map(
            (category) =>
              `<option value="${htmlEscape(category)}"${category === safeSelected ? " selected" : ""}>${htmlEscape(category)}</option>`
          )
        ].join("");
      };
      const actionReasonDetailOptionsHtml = (line, selectedCategory = "", selectedDetail = "") => {
        const safeCategory = normalizeActionReasonCategory(selectedCategory);
        const safeDetail = String(selectedDetail || "").trim();
        const options = safeCategory ? downtimeDetailOptions(line, safeCategory) : [];
        const hasSelected = options.some((option) => option.value === safeDetail);
        const fallbackLabel = safeCategory ? downtimeDetailLabel(line, safeCategory, safeDetail) || safeDetail : safeDetail;
        const fallbackOption = safeDetail && !hasSelected
          ? [`<option value="${htmlEscape(safeDetail)}" selected>${htmlEscape(fallbackLabel)}</option>`]
          : [];
        const placeholder = safeCategory
          ? (safeCategory === "Equipment" ? "Downtime Reason / Stage" : "Downtime Reason")
          : "Downtime Reason (optional)";
        return [
          `<option value="">${htmlEscape(placeholder)}</option>`,
          ...fallbackOption,
          ...options.map(
            (option) =>
              `<option value="${htmlEscape(option.value)}"${option.value === safeDetail ? " selected" : ""}>${htmlEscape(option.label)}</option>`
          )
        ].join("");
      };

      managerActionList.innerHTML = managerActions.length
        ? `
          <div class="pending-log-wrap supervisor-action-ticket-wrap manager-action-ticket-wrap">
            <div class="pending-log-list supervisor-action-ticket-list manager-action-ticket-list">
              ${managerActions
                .map((action) => {
                  const line = action.lineId && appState.lines[action.lineId] ? appState.lines[action.lineId] : null;
                  const lineName = line?.name || "Unassigned";
                  const relatedReasonCategory = normalizeActionReasonCategory(action.relatedReasonCategory);
                  const relatedReasonDetail = relatedReasonCategory ? String(action.relatedReasonDetail || "").trim() : "";
                  const createdBy = String(action.createdBy || action.supervisorName || "System").trim() || "System";
                  const createdLabel = String(action.createdAt || "").replace("T", " ").slice(0, 16) || "-";
                  const dueDateValue = /^\d{4}-\d{2}-\d{2}$/.test(String(action.dueDate || "").trim()) ? String(action.dueDate).trim() : "";
                  const dueLabel = dueDateValue || "No due date";
                  const relationSummary = actionRelationSummary(action, line);
                  const description = String(action.description || "").trim() || "No description provided.";
                  const title = String(action.title || "").trim() || "Untitled action";
                  const priority = normalizeActionPriority(action.priority);
                  const status = normalizeActionStatus(action.status);
                  const editing = isManagerActionTicketEditRow(action.id);
                  const ticketIdRaw = String(action.id || "").trim();
                  const ticketId = ticketIdRaw ? ticketIdRaw.slice(-6).toUpperCase() : "N/A";
                  return `
                    <article class="pending-log-item supervisor-action-ticket manager-action-ticket" data-manager-action-row data-manager-line-id="${htmlEscape(
                      action.lineId || ""
                    )}">
                      <div class="pending-log-meta">
                        <h5>
                          ${htmlEscape(title)}
                          <span class="supervisor-action-ticket-id">#${htmlEscape(ticketId)}</span>
                        </h5>
                        <p class="supervisor-action-ticket-line">${htmlEscape(lineName)} | Due ${htmlEscape(dueLabel)} | Created ${htmlEscape(
                          createdLabel
                        )} by ${htmlEscape(createdBy)}</p>
                        ${relationSummary ? `<p class="supervisor-action-ticket-relation">Related: ${htmlEscape(relationSummary)}</p>` : ""}
                        <p class="supervisor-action-ticket-description">${htmlEscape(description)}</p>
                        ${
                          editing
                            ? `
                        <div class="manager-action-ticket-edit-grid">
                          <label class="manager-action-ticket-field">
                            <span>Urgency</span>
                            <select class="manager-action-priority-select" data-manager-action-priority>
                              ${actionPriorityOptionsHtml(priority)}
                            </select>
                          </label>
                          <label class="manager-action-ticket-field">
                            <span>Status</span>
                            <select class="manager-action-status-select" data-manager-action-status>
                              ${actionStatusOptionsHtml(status)}
                            </select>
                          </label>
                          <label class="manager-action-ticket-field">
                            <span>Due Date</span>
                            <input type="date" class="manager-action-due-date-input" data-manager-action-due-date value="${htmlEscape(
                              dueDateValue
                            )}" />
                          </label>
                          <label class="manager-action-ticket-field">
                            <span>Assigned</span>
                            <select class="manager-action-assign-select" data-manager-action-supervisor>
                              ${supervisorOptionsHtml(action.supervisorUsername, action.supervisorName)}
                            </select>
                          </label>
                        </div>
                        <div class="manager-action-related-cell manager-action-ticket-related">
                          <select class="manager-action-related-equipment-select" data-manager-action-related-equipment>
                            ${actionEquipmentOptionsHtml(line, action.relatedEquipmentId)}
                          </select>
                          <select class="manager-action-reason-category-select" data-manager-action-reason-category>
                            ${actionReasonCategoryOptionsHtml(relatedReasonCategory)}
                          </select>
                          <select class="manager-action-reason-detail-select" data-manager-action-reason-detail${
                            relatedReasonCategory ? "" : " disabled"
                          }>
                            ${actionReasonDetailOptionsHtml(line, relatedReasonCategory, relatedReasonDetail)}
                          </select>
                        </div>
                      `
                            : ""
                        }
                      </div>
                      <div class="pending-log-actions supervisor-action-ticket-badges manager-action-ticket-controls">
                        <span class="supervisor-action-ticket-badge priority-${actionTicketToneKey(priority)}">${htmlEscape(priority)}</span>
                        <span class="supervisor-action-ticket-badge status-${actionTicketToneKey(status)}">${htmlEscape(status)}</span>
                        <div class="table-action-stack manager-action-actions-cell">
                          ${
                            editing
                              ? `<button type="button" class="table-edit-pill is-save" data-manager-action-reassign="${htmlEscape(
                                  action.id
                                )}">Save</button>`
                              : `<button type="button" class="table-edit-pill" data-manager-action-edit="${htmlEscape(action.id)}">Edit</button>`
                          }
                          <button type="button" class="table-edit-pill table-delete-pill" data-manager-action-delete="${htmlEscape(
                            action.id
                          )}">Delete</button>
                        </div>
                      </div>
                    </article>
                  `;
                })
                .join("")}
            </div>
          </div>
        `
        : `<p class="muted">No actions logged yet.</p>`;
    }
    return;
  }

  const assignedLineShiftMap = normalizeSupervisorLineShifts(session?.assignedLineShifts, appState.lines, session?.assignedLineIds || []);
  let assignedIds = Object.keys(assignedLineShiftMap);
  if (session) {
    session.assignedLineIds = assignedIds.slice();
    session.assignedLineShifts = clone(assignedLineShiftMap);
  }
  loginSection.classList.toggle("hidden", Boolean(session));
  appSection.classList.toggle("hidden", !session);
  if (!session) return;

  const activeMainTab = ["supervisorDay", "supervisorData", "supervisorActions"].includes(appState.supervisorMainTab)
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
  appState.supervisorSelectedDate = normalizeWeekdayIsoDate(appState.supervisorSelectedDate || todayISO(), { direction: -1 });
  if (!visualAllowedShifts.includes(appState.supervisorSelectedShift)) {
    appState.supervisorSelectedShift = visualAllowedShifts.includes("Full Day") ? "Full Day" : visualAllowedShifts[0] || "Day";
  }

  supervisorMobileModeBtn.classList.toggle("hidden", false);
  appSection.classList.toggle("mobile-mode", Boolean(appState.supervisorMobileMode));
  supervisorMobileModeBtn.classList.toggle("active", Boolean(appState.supervisorMobileMode));
  supervisorMobileModeBtn.textContent = appState.supervisorMobileMode ? "Mobile Mode On" : "Mobile Mode";
  lineSelect.innerHTML = assignedIds.length
    ? assignedIds.map((id) => `<option value="${htmlEscape(id)}">${htmlEscape(appState.lines[id]?.name || id)}</option>`).join("")
    : `<option value="">No assigned lines</option>`;
  lineSelect.value = appState.supervisorSelectedLineId || assignedIds[0] || "";
  if (actionLineInput) {
    actionLineInput.innerHTML = assignedIds.length
      ? assignedIds.map((id) => `<option value="${htmlEscape(id)}">${htmlEscape(appState.lines[id]?.name || id)}</option>`).join("")
      : `<option value="">No assigned lines</option>`;
    const preferredActionLineId = assignedIds.includes(actionLineInput.value)
      ? actionLineInput.value
      : appState.supervisorSelectedLineId || assignedIds[0] || "";
    actionLineInput.value = preferredActionLineId;
    actionLineInput.disabled = !assignedIds.length;
  }
  if (actionReasonCategoryInput) {
    setActionReasonCategoryOptions(actionReasonCategoryInput, actionReasonCategoryInput.value || "");
  }
  if (actionEquipmentInput || actionReasonDetailInput) {
    const actionLine = actionLineInput?.value && appState.lines[actionLineInput.value] ? appState.lines[actionLineInput.value] : null;
    if (actionEquipmentInput) {
      setActionEquipmentOptions(actionEquipmentInput, actionLine, actionEquipmentInput.value || "");
    }
    if (actionReasonDetailInput) {
      setActionReasonDetailOptions(
        actionReasonDetailInput,
        actionLine,
        String(actionReasonCategoryInput?.value || ""),
        actionReasonDetailInput.value || ""
      );
    }
  }
  if (actionPriorityInput && !ACTION_PRIORITY_OPTIONS.includes(actionPriorityInput.value)) {
    actionPriorityInput.value = "Medium";
  }
  if (actionStatusInput && !ACTION_STATUS_OPTIONS.includes(actionStatusInput.value)) {
    actionStatusInput.value = "Open";
  }
  supervisorAutoDateValue(actionDueDateInput, todayISO());
  syncRunProductInputsFromCatalog();
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

  supervisorAutoDateValue(shiftDateInput, supervisorAutoEntryDate());
  supervisorAutoDateValue(runDateInput, supervisorAutoEntryDate());
  supervisorAutoDateValue(downDateInput, supervisorAutoEntryDate());
  const activeTab = ["superShift", "superRun", "superDown"].includes(appState.supervisorTab) ? appState.supervisorTab : "superShift";
  document.querySelectorAll("[data-supervisor-tab]").forEach((btn) => {
    const active = btn.dataset.supervisorTab === activeTab;
    btn.classList.toggle("active", active);
    btn.setAttribute("aria-pressed", String(active));
  });
  document.querySelectorAll("#supervisorAppSection .data-section").forEach((section) => {
    section.classList.toggle("active", section.id === activeTab);
  });
  assignedIds.forEach((id) => ensureManagerLogRowIds(appState.lines[id]));
  const rowExistsOnAssignedLines = (type, logId) => {
    const safeLogId = String(logId || "").trim();
    if (!safeLogId) return false;
    const rowsKey = type === "run" ? "runRows" : type === "downtime" ? "downtimeRows" : "shiftRows";
    return assignedIds.some((id) =>
      (appState.lines[id]?.[rowsKey] || []).some((row) => String(row?.id || "").trim() === safeLogId)
    );
  };
  if (supervisorShiftTileEditId && !rowExistsOnAssignedLines("shift", supervisorShiftTileEditId)) {
    supervisorShiftTileEditId = "";
  }
  if (shiftLogIdInput?.value && !rowExistsOnAssignedLines("shift", shiftLogIdInput.value)) {
    shiftLogIdInput.value = "";
  }
  if (runLogIdInput?.value && !rowExistsOnAssignedLines("run", runLogIdInput.value)) {
    runLogIdInput.value = "";
  }
  if (downLogIdInput?.value && !rowExistsOnAssignedLines("downtime", downLogIdInput.value)) {
    downLogIdInput.value = "";
  }
  updateSupervisorProgressButtonLabels();

  const currentSelectedLogId =
    activeTab === "superRun"
      ? String(runLogIdInput?.value || "").trim()
      : activeTab === "superDown"
        ? String(downLogIdInput?.value || "").trim()
        : String(selectedSupervisorShiftLogId() || "").trim();
  const isShiftOpen = (row) => isPendingShiftLogRow(row);
  const isRunOpen = (row) => isPendingRunLogRow(row);
  const isDownOpen = (row) => isPendingDowntimeLogRow(row);
  const entryStatusHtml = (open) =>
    `<span class="entry-status-pill ${open ? "is-open" : "is-finalised"}">${open ? "In Progress" : "Finalised"}</span>`;
  const entryActionHtml = (log) => `
    <div class="table-action-stack">
      <button
        type="button"
        class="table-edit-pill"
        data-super-entry-action="edit"
        data-super-entry-type="${htmlEscape(log.entryType)}"
        data-super-entry-id="${htmlEscape(log.logId)}"
        data-super-entry-line-id="${htmlEscape(log.lineId)}"
      >Edit</button>
      ${log.isOpen
        ? `
          <button
            type="button"
            class="table-edit-pill finalise-pill"
            data-super-entry-action="finalise"
            data-super-entry-type="${htmlEscape(log.entryType)}"
            data-super-entry-id="${htmlEscape(log.logId)}"
            data-super-entry-line-id="${htmlEscape(log.lineId)}"
          >Finalise</button>
        `
        : ""}
    </div>
  `;
  const emptyLogLabel =
    activeTab === "superRun"
      ? "No production run submissions yet."
      : activeTab === "superDown"
        ? "No downtime submissions yet."
        : "No shift submissions yet.";
  const logs = assignedIds
    .flatMap((id) => {
      const line = appState.lines[id];
      if (!line) return [];
      if (activeTab === "superRun") {
        return (line.runRows || [])
          .filter((row) => row.submittedAt && isOperationalDate(String(row?.date || "")) && supervisorOwnsPendingLogRow(row, session))
          .map((row) => ({
            entryType: "run",
            logId: String(row.id || ""),
            lineId: id,
            lineName: line.name,
            date: row.date,
            shift: resolveTimedLogShiftLabel(row, line, "productionStartTime", "finishTime"),
            type: "Run",
            summary: `${row.product || "-"} (${formatNum(row.unitsProduced, 0)} units)`,
            supervisor: row.submittedBy || session?.name || session?.username || "-",
            createdAt: row.submittedAt,
            isOpen: isRunOpen(row),
            sortValue: rowNewestSortValue(row, "productionStartTime")
          }));
      }
      if (activeTab === "superDown") {
        return (line.downtimeRows || [])
          .filter((row) => row.submittedAt && isOperationalDate(String(row?.date || "")) && supervisorOwnsPendingLogRow(row, session))
          .map((row) => {
            const parsedReason = parseDowntimeReasonParts(row.reason, row.equipment);
            const reasonCategory = row.reasonCategory || parsedReason.reasonCategory || "Downtime";
            const reasonDetail = row.reasonDetail || parsedReason.reasonDetail || "";
            const detailLabel = reasonCategory === "Equipment"
              ? stageNameByIdForLine(line, row.equipment || reasonDetail) || "Equipment"
              : reasonDetail;
            return {
              entryType: "downtime",
              logId: String(row.id || ""),
              lineId: id,
              lineName: line.name,
              date: row.date,
              shift: resolveTimedLogShiftLabel(row, line, "downtimeStart", "downtimeFinish"),
              type: "Downtime",
              summary: `${reasonCategory}${detailLabel ? ` > ${detailLabel}` : ""}${row.reason ? `: ${row.reason}` : ""}`,
              supervisor: row.submittedBy || session?.name || session?.username || "-",
              createdAt: row.submittedAt,
              isOpen: isDownOpen(row),
              sortValue: rowNewestSortValue(row, "downtimeStart")
            };
          });
      }
      return (line.shiftRows || [])
        .filter((row) => row.submittedAt && isOperationalDate(String(row?.date || "")) && supervisorOwnsPendingLogRow(row, session))
        .map((row) => ({
          entryType: "shift",
          logId: String(row.id || ""),
          lineId: id,
          lineName: line.name,
          date: row.date,
          shift: row.shift,
          type: "Shift",
          summary: `${row.startTime || "-"} to ${row.finishTime || "-"}`,
          supervisor: row.submittedBy || session?.name || session?.username || "-",
          createdAt: row.submittedAt,
          isOpen: isShiftOpen(row),
          sortValue: rowNewestSortValue(row, "startTime")
        }));
    })
    .sort((a, b) => {
      const dateCmp = String(b.date || "").localeCompare(String(a.date || ""));
      if (dateCmp !== 0) return dateCmp;
      const sortDiff = num(b.sortValue) - num(a.sortValue);
      if (sortDiff !== 0) return sortDiff;
      const createdCmp = String(b.createdAt || "").localeCompare(String(a.createdAt || ""));
      if (createdCmp !== 0) return createdCmp;
      const lineCmp = String(a.lineName || "").localeCompare(String(b.lineName || ""), undefined, { sensitivity: "base" });
      if (lineCmp !== 0) return lineCmp;
      return String(b.logId || "").localeCompare(String(a.logId || ""));
    });

  const groupedLogs = logs.reduce((groups, log) => {
    const logDate = String(log?.date || "").trim();
    const label = isIsoDateValue(logDate)
      ? formatIsoDateLabel(logDate, { weekday: "long", day: "numeric", month: "long", year: "numeric" })
      : (logDate || "Unknown Date");
    const currentGroup = groups[groups.length - 1];
    if (currentGroup && currentGroup.date === logDate) {
      currentGroup.logs.push(log);
      return groups;
    }
    groups.push({
      date: logDate,
      label,
      logs: [log]
    });
    return groups;
  }, []);

  entryList.innerHTML = logs.length
    ? `
      <table>
        <thead><tr><th>Line</th><th>Shift</th><th>Type</th><th>Details</th><th>Status</th><th>By</th><th></th></tr></thead>
        <tbody>
          ${groupedLogs
            .map(
              (group) => `
                <tr class="entry-date-divider">
                  <td colspan="7">${htmlEscape(group.label)}</td>
                </tr>
                ${group.logs
                  .map(
                    (log) => `
                      <tr${currentSelectedLogId && currentSelectedLogId === log.logId ? ' class="entry-log-row-active"' : ""}>
                        <td>${htmlEscape(log.lineName)}</td>
                        <td>${htmlEscape(log.shift)}</td>
                        <td>${htmlEscape(log.type)}</td>
                        <td>${htmlEscape(log.summary)}</td>
                        <td>${entryStatusHtml(log.isOpen)}</td>
                        <td>${htmlEscape(log.supervisor)}</td>
                        <td class="entry-actions-cell">${entryActionHtml(log)}</td>
                      </tr>
                    `
                  )
                  .join("")}
              `
            )
            .join("")}
        </tbody>
      </table>
    `
    : `<p class="muted">${htmlEscape(emptyLogLabel)}</p>`;

  entryCards.innerHTML = logs.length
    ? groupedLogs
        .map(
          (group) => `
          <section class="entry-card-group">
            <h4 class="entry-group-title">${htmlEscape(group.label)}</h4>
            <div class="entry-card-group-list">
              ${group.logs
                .map(
                  (log) => `
                    <article class="entry-card${currentSelectedLogId && currentSelectedLogId === log.logId ? " entry-log-row-active" : ""}">
                      <div class="entry-card-meta">
                        <h4>${htmlEscape(log.type)} | ${htmlEscape(log.lineName)}</h4>
                        <p>${htmlEscape(log.shift)} | ${htmlEscape(log.summary)}</p>
                        <p>${entryStatusHtml(log.isOpen)} By ${htmlEscape(log.supervisor)}</p>
                      </div>
                      <div class="entry-card-actions">
                        ${entryActionHtml(log)}
                      </div>
                    </article>
                  `
                )
                .join("")}
            </div>
          </section>
        `
        )
        .join("")
    : `<p class="muted">${htmlEscape(emptyLogLabel)}</p>`;

  const supervisorActions = supervisorActionsForUsername(session?.username).slice(0, 50);
  const supervisorActionTicketsHtml = renderSupervisorActionTickets(supervisorActions);
  if (actionList) {
    actionList.innerHTML = supervisorActionTicketsHtml;
  }
  if (actionCards) {
    actionCards.innerHTML = supervisorActionTicketsHtml;
  }

  syncAllNowInputPrefixStates();
}

function renderAll() {
  enforceAppVariantState();
  const baselineDate = normalizeWeekdayIsoDate(todayISO(), { direction: -1 });
  appState.dashboardDate = normalizeWeekdayIsoDate(appState.dashboardDate || baselineDate, { direction: -1 });
  appState.supervisorSelectedDate = normalizeWeekdayIsoDate(appState.supervisorSelectedDate || baselineDate, { direction: -1 });
  Object.values(appState.lines || {}).forEach((line) => {
    if (!line) return;
    line.selectedDate = normalizeWeekdayIsoDate(line.selectedDate || baselineDate, { direction: -1 });
  });
  const managerSessionActive = Boolean(managerBackendSession.backendToken);
  if (!activePasswordResetSession()?.backendToken) {
    closePasswordResetModal();
  }
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

  if (appState.activeView !== "line" || !state) {
    closeDayVizAddRecordModal();
    return;
  }
  if (dayVizAddRecordModalState && String(dayVizAddRecordModalState.lineId || "") !== String(state.id || "")) {
    closeDayVizAddRecordModal();
  }
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
      addLineBtn.disabled = true;
      addLineBtn.classList.add("hidden");
    }
    if (addArrowBtn) {
      addArrowBtn.disabled = true;
      addArrowBtn.classList.add("hidden");
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
  setActiveManagerLineTab(activeManagerLineTabId());
  const pendingTabId = parseManagerDataTabId(pendingManagerDataTabRestore.tabId);
  const pendingLineId = String(pendingManagerDataTabRestore.lineId || "");
  if (pendingTabId && (!pendingLineId || pendingLineId === String(state.id || ""))) {
    state.activeDataTab = pendingTabId;
    pendingManagerDataTabRestore = { lineId: "", tabId: "" };
  }
  setShiftToggleUI();
  setActiveDataSubtab();
  renderCrewInputs();
  renderThroughputInputs();
  renderTrackingTables();
  renderAuditTrail();
  renderVisualiser();
  renderDayVisualiser();
  renderLineTrends();
  syncLineSettingsSaveUI();
}

function setActiveDataSubtab() {
  const activeId = parseManagerDataTabId(state?.activeDataTab) || "dataShift";
  if (state) state.activeDataTab = activeId;
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
  enforceAppVariantState();
  const shouldRefreshSupervisor = appState.appMode === "supervisor" && Boolean(appState.supervisorSession?.backendToken);
  const shouldRefreshManager =
    !shouldRefreshSupervisor && (appState.appMode === "manager" || appState.activeView === "line") && Boolean(managerBackendSession.backendToken);
  const bindStep = (name, fn) => {
    try {
      fn();
    } catch (error) {
      console.error(`Bootstrap step failed: ${name}`, error);
    }
  };
  try {
    bindStep("bindTabs", bindTabs);
    bindStep("bindRunCrewingPatternModal", bindRunCrewingPatternModal);
    bindStep("bindDayVizBlockModal", bindDayVizBlockModal);
    bindStep("bindDayVizAddRecordModal", bindDayVizAddRecordModal);
    bindStep("bindPasswordResetModal", bindPasswordResetModal);
    bindStep("bindHome", bindHome);
    bindStep("bindDataSubtabs", bindDataSubtabs);
    bindStep("bindVisualiserControls", bindVisualiserControls);
    bindStep("bindTrendModal", bindTrendModal);
    bindStep("bindLineTrendDetailModal", bindLineTrendDetailModal);
    bindStep("bindStageSettingsModal", bindStageSettingsModal);
    bindStep("bindForms", bindForms);
    bindStep("bindDataControls", bindDataControls);
    renderAll();
    if (shouldRefreshSupervisor) {
      await refreshHostedState(appState.supervisorSession);
    } else if (shouldRefreshManager) {
      await refreshHostedState();
    }
  } catch (error) {
    console.error("Bootstrap failed:", error);
    renderAll();
  } finally {
    hideStartupLoading({ force: true });
  }
}

bootstrapApp();

window.addEventListener("hashchange", () => {
  try {
    restoreRouteFromHash();
    enforceAppVariantState();
    state = appState.lines[appState.activeLineId] || appState.lines[Object.keys(appState.lines)[0]] || null;
    if (state) appState.activeLineId = state.id;
    renderAll();
  } catch (error) {
    console.error("Hash navigation restore failed:", error);
    appState.activeView = "home";
    renderAll();
  }
});
