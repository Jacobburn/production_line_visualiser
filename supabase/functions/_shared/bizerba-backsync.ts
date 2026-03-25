import { createClient } from "npm:@supabase/supabase-js@2";

export type BizerbaBacksyncJobRow = {
  id: string;
  status: string;
  trigger: string;
  full_rescan: boolean;
  requested_by_user_id: string;
  requested_by_name: string;
  summary: Record<string, unknown> | null;
  state: Record<string, unknown> | null;
  error: string;
  worker_id: string;
  lease_expires_at: string | null;
  started_at: string | null;
  finished_at: string | null;
  created_at: string;
  updated_at: string;
};

type JsonRecord = Record<string, unknown>;

const SUPABASE_URL = Deno.env.get("SUPABASE_URL") ?? "";
const SUPABASE_SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ??
  "";
const BACKSYNC_FUNCTION_KEY = Deno.env.get("BIZERBA_BACKSYNC_FUNCTION_KEY") ??
  "";

if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY || !BACKSYNC_FUNCTION_KEY) {
  throw new Error(
    "Missing required edge function environment variables for Bizerba backsync.",
  );
}

export const bizerbaAutoMinRunSamples = 10;
export const bizerbaAutoDowntimeGapMins = 2;
export const bizerbaAutoRunBreakGapMins = 30;
export const bizerbaAutoLogNotePrefix = "[AUTO_BIZERBA]";
export const bizerbaAutoDowntimeReason = "Bizerba Auto Gap";

export const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, {
  auth: { persistSession: false },
});

const corsHeaders = {
  "Content-Type": "application/json",
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type, x-backsync-key",
  "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
};

export function jsonResponse(payload: unknown, status = 200): Response {
  return new Response(JSON.stringify(payload), { status, headers: corsHeaders });
}

export function optionsResponse(): Response {
  return new Response("ok", { status: 200, headers: corsHeaders });
}

export function unauthorizedResponse(): Response {
  return jsonResponse({ error: "Unauthorized" }, 401);
}

export function isAuthorized(req: Request): boolean {
  return req.headers.get("x-backsync-key") === BACKSYNC_FUNCTION_KEY;
}

export function normalizeIsoDate(value: unknown): string {
  const raw = String(value ?? "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return "";
  return parsed.toISOString().slice(0, 10);
}

export function parseTimestamp(value: unknown): Date | null {
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

export function clampNonNegativeInteger(value: unknown, fallback = 0): number {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(0, Math.floor(parsed));
}

export function clampNonNegativeNumber(value: unknown, fallback = 0): number {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(0, parsed);
}

export function bizerbaArticleLookupKeys(value: unknown): string[] {
  const raw = String(value ?? "").trim();
  if (!raw) return [];
  const lowered = raw.toLowerCase();
  const compact = lowered.replace(/\s+/g, "");
  const keys: string[] = [];
  if (compact) keys.push(compact);
  if (lowered && lowered !== compact) keys.push(lowered);
  const numericCompact = normalizeBizerbaNumericArticleKey(compact);
  if (numericCompact && !keys.includes(numericCompact)) keys.push(numericCompact);
  const numericLowered = normalizeBizerbaNumericArticleKey(lowered);
  if (numericLowered && !keys.includes(numericLowered)) keys.push(numericLowered);
  return keys;
}

function normalizeBizerbaNumericArticleKey(value: string): string {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const match = raw.match(/^([+-]?)(\d+)(?:\.(\d+))?$/);
  if (!match) return "";
  const sign = match[1] === "-" ? "-" : "";
  const intPart = String(match[2] || "").replace(/^0+/, "") || "0";
  const fraction = String(match[3] || "").replace(/0+$/, "");
  if (!fraction) return `${sign}${intPart}`;
  return `${sign}${intPart}.${fraction}`;
}

export function bizerbaArticleKey(value: unknown): string {
  const [first] = bizerbaArticleLookupKeys(value);
  return first || "";
}

export function normalizeBigIntString(value: unknown, fallback = "0"): string {
  const raw = String(value ?? "").trim();
  if (!/^-?\d+$/.test(raw)) return fallback;
  try {
    return BigInt(raw).toString();
  } catch {
    return fallback;
  }
}

export function normalizeBizerbaDate(
  value: unknown,
  timestamp: Date | null,
): string {
  const direct = normalizeIsoDate(value);
  if (direct) return direct;
  if (!(timestamp instanceof Date) || Number.isNaN(timestamp.getTime())) return "";
  return timestamp.toISOString().slice(0, 10);
}

export function buildBizerbaAutoRunNote(
  deviceName: string,
  articleNumber: string,
  productionStartTimestamp: Date | null,
): string {
  const safeDevice = String(deviceName || "").trim();
  const safeArticle = String(articleNumber || "").trim();
  const safeStart = productionStartTimestamp &&
      !Number.isNaN(productionStartTimestamp.getTime())
    ? productionStartTimestamp.toISOString()
    : "";
  return `${bizerbaAutoLogNotePrefix} run device=${safeDevice} article=${safeArticle}${
    safeStart ? ` start=${safeStart}` : ""
  }`;
}

export function buildBizerbaAutoDowntimeNote(
  deviceName: string,
  articleNumber: string,
  gapMinutes: number,
  downtimeStartTimestamp: Date | null,
  downtimeFinishTimestamp: Date | null,
): string {
  const safeDevice = String(deviceName || "").trim();
  const safeArticle = String(articleNumber || "").trim();
  const safeGap = Number.isFinite(gapMinutes) ? gapMinutes.toFixed(2) : "0.00";
  const safeStart = downtimeStartTimestamp &&
      !Number.isNaN(downtimeStartTimestamp.getTime())
    ? downtimeStartTimestamp.toISOString()
    : "";
  const safeFinish = downtimeFinishTimestamp &&
      !Number.isNaN(downtimeFinishTimestamp.getTime())
    ? downtimeFinishTimestamp.toISOString()
    : "";
  const parts = [
    `${bizerbaAutoLogNotePrefix} downtime`,
    `device=${safeDevice}`,
    `article=${safeArticle}`,
    `gap=${safeGap}m`,
  ];
  if (safeStart) parts.push(`start=${safeStart}`);
  if (safeFinish) parts.push(`finish=${safeFinish}`);
  return parts.join(" ");
}

export function mapJobForApi(row: Partial<BizerbaBacksyncJobRow> = {}): JsonRecord {
  return {
    id: String(row.id || "").trim(),
    status: String(row.status || "").trim() || "unknown",
    trigger: String(row.trigger || "").trim() || "manual",
    fullRescan: Boolean(row.full_rescan),
    createdAt: String(row.created_at || ""),
    startedAt: String(row.started_at || ""),
    finishedAt: String(row.finished_at || ""),
    requestedByUserId: String(row.requested_by_user_id || ""),
    requestedByName: String(row.requested_by_name || ""),
    error: String(row.error || ""),
    summary: row.summary && typeof row.summary === "object" ? row.summary : null,
    state: row.state && typeof row.state === "object" ? row.state : null,
  };
}

export function parseBodyAsObject(payload: unknown): JsonRecord {
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) return {};
  return payload as JsonRecord;
}

export async function parseJsonBody(req: Request): Promise<JsonRecord> {
  try {
    const payload = await req.json();
    return parseBodyAsObject(payload);
  } catch {
    return {};
  }
}

export function sanitizeSummary(summary: unknown): JsonRecord {
  if (!summary || typeof summary !== "object" || Array.isArray(summary)) return {};
  return summary as JsonRecord;
}

export function sanitizeState(state: unknown): JsonRecord {
  if (!state || typeof state !== "object" || Array.isArray(state)) return {};
  return state as JsonRecord;
}
