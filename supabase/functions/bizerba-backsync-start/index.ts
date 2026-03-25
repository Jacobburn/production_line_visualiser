import {
  isAuthorized,
  jsonResponse,
  mapJobForApi,
  optionsResponse,
  parseJsonBody,
  supabase,
  unauthorizedResponse,
} from "../_shared/bizerba-backsync.ts";

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return optionsResponse();
  if (req.method !== "POST") return jsonResponse({ error: "Method not allowed" }, 405);
  if (!isAuthorized(req)) return unauthorizedResponse();

  const body = await parseJsonBody(req);
  const trigger = String(body.trigger || "manual").trim() || "manual";
  const fullRescan = Boolean(body.fullRescan);
  const requestedByUserId = String(body.requestedByUserId || "").trim();
  const requestedByName = String(body.requestedByName || "").trim();

  const { data: runningJobs, error: runningError } = await supabase
    .from("bizerba_backsync_jobs")
    .select("*")
    .in("status", ["queued", "running"])
    .order("created_at", { ascending: false })
    .limit(1);

  if (runningError) {
    console.error("bizerba-backsync-start list running jobs failed", runningError);
    return jsonResponse({ error: String(runningError.message || runningError) }, 500);
  }

  if (Array.isArray(runningJobs) && runningJobs.length > 0) {
    return jsonResponse({ job: mapJobForApi(runningJobs[0]) }, 202);
  }

  const { data: insertedJob, error: insertError } = await supabase
    .from("bizerba_backsync_jobs")
    .insert({
      status: "queued",
      trigger,
      full_rescan: fullRescan,
      requested_by_user_id: requestedByUserId,
      requested_by_name: requestedByName,
      summary: {
        trigger,
        fullRescan,
        startedAt: "",
        finishedAt: "",
        lineCount: 0,
        processedRows: 0,
        mappedRows: 0,
        unmappedRows: 0,
        createdRuns: 0,
        closedRuns: 0,
        createdDowntime: 0,
        skippedLines: 0,
        erroredLines: 0,
        lines: [],
      },
      state: {},
      error: "",
      worker_id: "",
      lease_expires_at: null,
      started_at: null,
      finished_at: null,
    })
    .select("*")
    .single();

  if (insertError || !insertedJob) {
    console.error("bizerba-backsync-start create job failed", insertError);
    return jsonResponse({ error: String(insertError?.message || "Could not create backsync job.") }, 500);
  }

  return jsonResponse({ job: mapJobForApi(insertedJob) }, 202);
});
