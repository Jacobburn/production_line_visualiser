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
  id: string;
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

function requiredId(value: unknown, field: string): string {
  if (value === undefined || value === null) {
    throw new Error(`Missing or invalid field: ${field}`);
  }
  const raw = typeof value === "string" ? value.trim() : String(value).trim();
  if (!/^\d+$/.test(raw)) {
    throw new Error(`Missing or invalid field: ${field}; expected integer id`);
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
  return {
    id: requiredId(payload.id, `${prefix}.id`),
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
    const { error } = await supabase.from("Bizerba_ID").upsert(rows, {
      onConflict: "id",
      ignoreDuplicates: true,
    });
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
