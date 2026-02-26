import { createClient } from "npm:@supabase/supabase-js@2";

type BizerbaPayload = {
  ArticleName: unknown;
  ActualNetWeightValue: unknown;
  DeviceName: unknown;
  Date: unknown;
  Timestamp: unknown;
  ArticleNumber: unknown;
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

const supabaseUrl = Deno.env.get("SUPABASE_URL") ?? "";
const serviceRoleKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? "";
const deviceKey = Deno.env.get("BIZERBA_DEVICE_KEY") ?? "";

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

  let payload: BizerbaPayload;
  try {
    payload = await req.json();
  } catch {
    return new Response("Invalid JSON body", { status: 400 });
  }

  try {
    const row = {
      ArticleName: requiredString(payload.ArticleName, "ArticleName"),
      ActualNetWeightValue: requiredNumeric(
        payload.ActualNetWeightValue,
        "ActualNetWeightValue",
      ),
      DeviceName: requiredString(payload.DeviceName, "DeviceName"),
      Date: requiredDate(payload.Date, "Date"),
      Timestamp: requiredTimestamp(payload.Timestamp, "Timestamp"),
      ArticleNumber: requiredString(payload.ArticleNumber, "ArticleNumber"),
    };

    const { error } = await supabase.from("Bizerba_ID").insert(row);
    if (error) {
      return new Response(JSON.stringify({ error: error.message }), {
        status: 400,
        headers: { "Content-Type": "application/json" },
      });
    }

    return new Response(JSON.stringify({ ok: true }), {
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
