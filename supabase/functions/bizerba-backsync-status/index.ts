import {
  isAuthorized,
  jsonResponse,
  mapJobForApi,
  optionsResponse,
  parseJsonBody,
  supabase,
  unauthorizedResponse,
} from "../_shared/bizerba-backsync.ts";

function readJobIdFromRequest(req: Request, body: Record<string, unknown>): string {
  const fromBody = String(body.jobId || "").trim();
  if (fromBody) return fromBody;
  try {
    const url = new URL(req.url);
    return String(url.searchParams.get("jobId") || "").trim();
  } catch {
    return "";
  }
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return optionsResponse();
  if (req.method !== "GET" && req.method !== "POST") {
    return jsonResponse({ error: "Method not allowed" }, 405);
  }
  if (!isAuthorized(req)) return unauthorizedResponse();

  const body = req.method === "POST" ? await parseJsonBody(req) : {};
  const jobId = readJobIdFromRequest(req, body);
  if (!jobId) return jsonResponse({ error: "Missing jobId" }, 400);

  const { data: job, error } = await supabase
    .from("bizerba_backsync_jobs")
    .select("*")
    .eq("id", jobId)
    .maybeSingle();

  if (error) {
    console.error("bizerba-backsync-status failed", error);
    return jsonResponse({ error: String(error.message || error) }, 500);
  }
  if (!job) return jsonResponse({ error: "Bizerba backsync job not found" }, 404);

  return jsonResponse({ job: mapJobForApi(job) }, 200);
});
