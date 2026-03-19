import { createClient } from "npm:@supabase/supabase-js@2";

type BizerbaPayload = {
  id?: unknown;
  ActualNetWeightValue: unknown;
  DeviceName: unknown;
  Date: unknown;
  Timestamp: unknown;
  ArticleNumber: unknown;
};

type BizerbaInsertRow = {
  source_id?: string;
  ActualNetWeightValue: number;
  DeviceName: string;
  Date: string;
  Timestamp: string;
  ArticleNumber: string;
};

function requiredString(value: unknown, field: string): string {
  if (typeof value !== "string" || !value.trim()) {
    throw new Error(`Missing or invalid field: ${field}`);
  }
  return value.trim();
}

function requiredNumeric(value: unknown, field: string): number {
  const parsed = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(parsed)) {
    throw new Error(`Missing or invalid field: ${field}`);
  }
  return parsed;
}

function requiredDate(value: unknown, field: string): string {
  const parsed = requiredString(value, field);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(parsed)) {
    throw new Error(`Invalid format for ${field}; expected YYYY-MM-DD`);
  }
  return parsed;
}

function requiredTimestamp(value: unknown, field: string): string {
  const raw = requiredString(value, field);
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) {
    throw new Error(`Invalid format for ${field}; expected ISO timestamp`);
  }
  return parsed.toISOString();
}

function optionalSourceId(value: unknown, field: string): string | undefined {
  if (value === undefined || value === null) {
    return undefined;
  }
  const raw = typeof value === "string" ? value.trim() : String(value).trim();
  if (!raw) {
    throw new Error(`Missing or invalid field: ${field}`);
  }
  return raw;
}

function requireObject(value: unknown, field: string): BizerbaPayload {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error(`Missing or invalid row: ${field}`);
  }
  return value as BizerbaPayload;
}

function parseBatchSize(raw: string | undefined): number {
  const parsed = Number(raw ?? "100000");
  if (!Number.isInteger(parsed) || parsed <= 0) {
    return 100000;
  }
  return parsed;
}

function extractPayloads(body: unknown): BizerbaPayload[] {
  if (Array.isArray(body)) {
    return body.map((entry, index) => requireObject(entry, `rows[${index}]`));
  }

  const payload = requireObject(body, "body");
  const maybeRows = (payload as { rows?: unknown }).rows;
  if (Array.isArray(maybeRows)) {
    return maybeRows.map((entry, index) => requireObject(entry, `rows[${index}]`));
  }

  return [payload];
}

function buildRow(payload: BizerbaPayload, index: number): BizerbaInsertRow {
  const prefix = `rows[${index}]`;
  const sourceId = optionalSourceId(payload.id, `${prefix}.id`);

  return {
    // The vendor's row id is used for idempotency, not as the table PK.
    ...(sourceId ? { source_id: sourceId } : {}),
    ActualNetWeightValue: requiredNumeric(
      payload.ActualNetWeightValue,
      `${prefix}.ActualNetWeightValue`,
    ),
    DeviceName: requiredString(payload.DeviceName, `${prefix}.DeviceName`),
    Date: requiredDate(payload.Date, `${prefix}.Date`),
    Timestamp: requiredTimestamp(payload.Timestamp, `${prefix}.Timestamp`),
    ArticleNumber: requiredString(
      payload.ArticleNumber,
      `${prefix}.ArticleNumber`,
    ),
  };
}

function hasSourceId(
  row: BizerbaInsertRow,
): row is BizerbaInsertRow & { source_id: string } {
  return typeof row.source_id === "string" && row.source_id.length > 0;
}

async function persistRows(rows: BizerbaInsertRow[]) {
  const rowsWithSourceId = rows.filter(hasSourceId);
  const rowsWithoutSourceId = rows.filter((row) => !hasSourceId(row));

  if (rowsWithSourceId.length > 0) {
    const { error } = await supabase.from("Bizerba_ID").upsert(rowsWithSourceId, {
      onConflict: "source_id",
      defaultToNull: false,
      ignoreDuplicates: true,
    });
    if (error) {
      return error;
    }
  }

  if (rowsWithoutSourceId.length > 0) {
    const { error } = await supabase.from("Bizerba_ID").insert(rowsWithoutSourceId, {
      defaultToNull: false,
    });
    if (error) {
      return error;
    }
  }

  return null;
}

const supabaseUrl = Deno.env.get("SUPABASE_URL") ?? "";
const serviceRoleKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? "";
const deviceKey = Deno.env.get("BIZERBA_DEVICE_KEY") ?? "";
const maxBatchSize = parseBatchSize(Deno.env.get("BIZERBA_MAX_BATCH_SIZE"));

if (!supabaseUrl || !serviceRoleKey || !deviceKey) {
  throw new Error("Missing required function environment variables.");
}

const supabase = createClient(supabaseUrl, serviceRoleKey, {
  auth: { persistSession: false },
});

Deno.serve(async (req) => {
  if (req.method !== "POST") {
    return new Response("Method not allowed", { status: 405 });
  }

  if (req.headers.get("x-device-key") !== deviceKey) {
    return new Response("Unauthorized", { status: 401 });
  }

  let requestBody: unknown;
  try {
    requestBody = await req.json();
  } catch {
    return new Response("Invalid JSON body", { status: 400 });
  }

  try {
    const payloads = extractPayloads(requestBody);
    if (payloads.length === 0) {
      throw new Error("Payload must contain at least one row.");
    }
    if (payloads.length > maxBatchSize) {
      throw new Error(`Payload exceeds max batch size of ${maxBatchSize}.`);
    }

    const rows = payloads.map((payload, index) => buildRow(payload, index));
    const error = await persistRows(rows);
    if (error) {
      return new Response(JSON.stringify({ error: error.message }), {
        status: 400,
        headers: { "Content-Type": "application/json" },
      });
    }

    return new Response(JSON.stringify({ ok: true, received: rows.length }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Bad request";
    return new Response(JSON.stringify({ error: message }), {
      status: 400,
      headers: { "Content-Type": "application/json" },
    });
  }
});
