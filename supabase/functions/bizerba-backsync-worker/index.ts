import {
  bizerbaArticleKey,
  bizerbaArticleLookupKeys,
  bizerbaAutoDowntimeGapMins,
  bizerbaAutoDowntimeReason,
  bizerbaAutoLogNotePrefix,
  bizerbaAutoMinRunSamples,
  bizerbaAutoRunBreakGapMins,
  buildBizerbaAutoDowntimeNote,
  buildBizerbaAutoRunNote,
  clampNonNegativeInteger,
  clampNonNegativeNumber,
  isAuthorized,
  jsonResponse,
  mapJobForApi,
  normalizeBigIntString,
  normalizeBizerbaDate,
  optionsResponse,
  parseJsonBody,
  parseTimestamp,
  sanitizeState,
  sanitizeSummary,
  supabase,
  unauthorizedResponse,
} from "../_shared/bizerba-backsync.ts";

type JsonObject = Record<string, unknown>;

type ActiveLineState = {
  runLogId: string;
  articleNumber: string;
  articleKey: string;
  product: string;
  runDate: string;
  startTimestamp: Date | null;
  lastDate: string;
  lastTimestamp: Date | null;
  unitsProduced: number;
};

type LineProcessSummary = {
  lineId: string;
  deviceName: string;
  fullRescan: boolean;
  processedRows: number;
  mappedRows: number;
  unmappedRows: number;
  createdRuns: number;
  closedRuns: number;
  createdDowntime: number;
  reachedRowLimit: boolean;
  skipped: boolean;
  reason: string;
  error: string;
};

type ProductCatalogEntry = {
  id: string;
  createdAt: string;
  allLines: boolean;
  lineIds: string[];
  articleCode: string;
  productName: string;
};

function emptyActiveLineState(): ActiveLineState {
  return {
    runLogId: "",
    articleNumber: "",
    articleKey: "",
    product: "",
    runDate: "",
    startTimestamp: null,
    lastDate: "",
    lastTimestamp: null,
    unitsProduced: 0,
  };
}

function mapStateRowToActiveState(row: JsonObject): ActiveLineState {
  const runLogId = String(row.active_run_log_id || "").trim();
  const articleNumber = String(row.active_article_number || "").trim();
  const articleKey = bizerbaArticleKey(articleNumber);
  const product = String(row.active_product || "").trim();
  const runDate = String(row.active_run_date || "").trim();
  const startTimestamp = parseTimestamp(row.active_start_timestamp);
  const lastDate = String(row.active_last_date || "").trim();
  const lastTimestamp = parseTimestamp(row.active_last_timestamp);
  const unitsProduced = clampNonNegativeNumber(row.active_units_produced, 0);
  if (!runLogId || !articleKey || !runDate || !startTimestamp || !lastTimestamp) {
    return emptyActiveLineState();
  }
  return {
    runLogId,
    articleNumber,
    articleKey,
    product,
    runDate,
    startTimestamp,
    lastDate: lastDate || runDate,
    lastTimestamp,
    unitsProduced,
  };
}

function parseCatalogValues(values: unknown): string[] {
  if (!Array.isArray(values)) return [];
  return values.map((value) => String(value ?? "").trim());
}

function mapCatalogEntries(rows: JsonObject[]): ProductCatalogEntry[] {
  return rows
    .map((row) => {
      const values = parseCatalogValues(row.catalog_values);
      const articleCode = String(values[0] || "").trim();
      const productName = String(values[1] || values[0] || "").trim();
      if (!articleCode || !productName) return null;
      const rawLineIds = Array.isArray(row.line_ids) ? row.line_ids : [];
      const lineIds = rawLineIds.map((lineId) => String(lineId || "").trim()).filter(Boolean);
      return {
        id: String(row.id || "").trim(),
        createdAt: String(row.created_at || ""),
        allLines: Boolean(row.all_lines),
        lineIds,
        articleCode,
        productName,
      };
    })
    .filter((entry): entry is ProductCatalogEntry => Boolean(entry));
}

function buildProductLookupForLine(
  lineId: string,
  entries: ProductCatalogEntry[],
): Map<string, string> {
  const ordered = entries.slice().sort((a, b) => {
    const aRank = a.lineIds.includes(lineId) ? 0 : a.allLines ? 1 : 2;
    const bRank = b.lineIds.includes(lineId) ? 0 : b.allLines ? 1 : 2;
    if (aRank !== bRank) return aRank - bRank;
    if (a.createdAt !== b.createdAt) return String(a.createdAt).localeCompare(String(b.createdAt));
    return a.id.localeCompare(b.id);
  });
  const lookup = new Map<string, string>();
  ordered.forEach((entry) => {
    const keys = bizerbaArticleLookupKeys(entry.articleCode);
    keys.forEach((key) => {
      if (!key || lookup.has(key)) return;
      lookup.set(key, entry.productName);
    });
  });
  return lookup;
}

function resolveMappedProduct(lookup: Map<string, string>, articleNumber: string): string {
  const keys = bizerbaArticleLookupKeys(articleNumber);
  for (const key of keys) {
    const product = String(lookup.get(key) || "").trim();
    if (product) return product;
  }
  return "";
}

function isLikelyTimeoutError(error: unknown): boolean {
  const message = String((error as { message?: unknown })?.message || "").toLowerCase();
  return message.includes("timeout") || message.includes("timed out") ||
    message.includes("statement timeout");
}

function formatTimeComponent(timestamp: Date): string {
  return timestamp.toISOString().slice(11, 19);
}

async function closeRunLog(runLogId: string, finishTimestamp: Date, unitsProduced: number): Promise<boolean> {
  if (!runLogId) return false;
  const { data, error } = await supabase
    .from("run_logs")
    .update({
      finish_time: formatTimeComponent(finishTimestamp),
      units_produced: clampNonNegativeNumber(unitsProduced, 0),
      submitted_at: new Date().toISOString(),
    })
    .eq("id", runLogId)
    .select("id")
    .limit(1);
  if (error) throw error;
  return Array.isArray(data) && data.length > 0;
}

async function upsertRunLog(
  lineId: string,
  runDate: string,
  product: string,
  productionStartTimestamp: Date,
  unitsProduced: number,
  note: string,
): Promise<{ runLogId: string; created: boolean }> {
  const startTime = formatTimeComponent(productionStartTimestamp);
  const { data: existing, error: existingError } = await supabase
    .from("run_logs")
    .select("id")
    .eq("line_id", lineId)
    .eq("date", runDate)
    .eq("production_start_time", startTime)
    .ilike("product", product)
    .ilike("notes", `${bizerbaAutoLogNotePrefix}%`)
    .order("created_at", { ascending: true })
    .limit(1);
  if (existingError) throw existingError;
  if (Array.isArray(existing) && existing.length > 0) {
    return { runLogId: String(existing[0]?.id || "").trim(), created: false };
  }

  const { data: inserted, error: insertError } = await supabase
    .from("run_logs")
    .insert({
      line_id: lineId,
      date: runDate,
      product,
      setup_start_time: null,
      production_start_time: startTime,
      finish_time: null,
      units_produced: clampNonNegativeNumber(unitsProduced, 0),
      run_crewing_pattern: {},
      notes: String(note || "").trim(),
      submitted_by_user_id: null,
    })
    .select("id")
    .single();
  if (insertError) throw insertError;
  return {
    runLogId: String(inserted?.id || "").trim(),
    created: true,
  };
}

async function insertDowntimeLog(
  lineId: string,
  downtimeDate: string,
  downtimeStartTimestamp: Date,
  downtimeFinishTimestamp: Date,
  articleNumber: string,
  deviceName: string,
  gapMinutes: number,
): Promise<boolean> {
  const downtimeStart = formatTimeComponent(downtimeStartTimestamp);
  const downtimeFinish = formatTimeComponent(downtimeFinishTimestamp);

  const { data: existing, error: existingError } = await supabase
    .from("downtime_logs")
    .select("id")
    .eq("line_id", lineId)
    .eq("date", downtimeDate)
    .eq("downtime_start", downtimeStart)
    .eq("downtime_finish", downtimeFinish)
    .eq("reason", bizerbaAutoDowntimeReason)
    .ilike("notes", `${bizerbaAutoLogNotePrefix}%`)
    .order("created_at", { ascending: true })
    .limit(1);
  if (existingError) throw existingError;
  if (Array.isArray(existing) && existing.length > 0) return false;

  const note = buildBizerbaAutoDowntimeNote(
    deviceName,
    articleNumber,
    gapMinutes,
    downtimeStartTimestamp,
    downtimeFinishTimestamp,
  );
  const { error: insertError } = await supabase
    .from("downtime_logs")
    .insert({
      line_id: lineId,
      date: downtimeDate,
      downtime_start: downtimeStart,
      downtime_finish: downtimeFinish,
      equipment_stage_id: null,
      reason: bizerbaAutoDowntimeReason,
      notes: note,
      exclude_from_calculation: false,
      submitted_by_user_id: null,
    });
  if (insertError) throw insertError;
  return true;
}

async function fetchBizerbaRowsChunk(
  deviceName: string,
  afterId: string,
  limit: number,
): Promise<JsonObject[]> {
  const safeLimit = Math.max(1, Math.min(5000, Math.floor(limit)));
  const { data, error } = await supabase
    .from("Bizerba_ID")
    .select("id,ArticleNumber,Date,Timestamp")
    .eq("DeviceName", deviceName)
    .gt("id", normalizeBigIntString(afterId, "0"))
    .order("id", { ascending: true })
    .limit(safeLimit);
  if (error) throw error;
  return Array.isArray(data) ? (data as JsonObject[]) : [];
}

async function ensureLineStateRow(lineId: string, deviceName: string): Promise<void> {
  const { error } = await supabase
    .from("bizerba_auto_line_states")
    .upsert(
      {
        line_id: lineId,
        device_name: deviceName,
      },
      { onConflict: "line_id" },
    );
  if (error) throw error;
}

async function loadLineStateRow(lineId: string): Promise<JsonObject | null> {
  const { data, error } = await supabase
    .from("bizerba_auto_line_states")
    .select("*")
    .eq("line_id", lineId)
    .maybeSingle();
  if (error) throw error;
  return data as JsonObject | null;
}

async function saveLineStateRow(lineId: string, payload: JsonObject): Promise<void> {
  const { error } = await supabase
    .from("bizerba_auto_line_states")
    .update(payload)
    .eq("line_id", lineId);
  if (error) throw error;
}

async function processLineChunk(
  line: { lineId: string; deviceName: string },
  opts: {
    fullRescan: boolean;
    resetState: boolean;
    batchSize: number;
    maxRows: number;
    productLookup: Map<string, string>;
  },
): Promise<LineProcessSummary> {
  const lineId = String(line.lineId || "").trim();
  const deviceName = String(line.deviceName || "").trim();
  const summary: LineProcessSummary = {
    lineId,
    deviceName,
    fullRescan: Boolean(opts.fullRescan),
    processedRows: 0,
    mappedRows: 0,
    unmappedRows: 0,
    createdRuns: 0,
    closedRuns: 0,
    createdDowntime: 0,
    reachedRowLimit: false,
    skipped: false,
    reason: "",
    error: "",
  };

  if (!lineId || !deviceName) {
    summary.skipped = true;
    summary.reason = "line_not_configured";
    return summary;
  }

  await ensureLineStateRow(lineId, deviceName);
  const stateRow = await loadLineStateRow(lineId);
  if (!stateRow) {
    summary.skipped = true;
    summary.reason = "line_state_missing";
    return summary;
  }

  const priorDeviceName = String(stateRow.device_name || "").trim();
  let lastProcessedId = normalizeBigIntString(stateRow.last_processed_bizerba_id, "0");
  let lastProcessedTimestamp = parseTimestamp(stateRow.last_processed_timestamp);
  let lastProcessedIdText = String(stateRow.last_processed_bizerba_id_text || "").trim();
  let pendingArticleNumber = String(stateRow.pending_article_number || "").trim();
  let pendingCount = clampNonNegativeInteger(stateRow.pending_count, 0);
  let activeState = mapStateRowToActiveState(stateRow);

  if (opts.resetState || (priorDeviceName && priorDeviceName.toLowerCase() !== deviceName.toLowerCase())) {
    if (activeState.runLogId && activeState.lastTimestamp) {
      const closed = await closeRunLog(
        activeState.runLogId,
        activeState.lastTimestamp,
        activeState.unitsProduced,
      );
      if (closed) summary.closedRuns += 1;
    }
    lastProcessedId = "0";
    lastProcessedTimestamp = null;
    lastProcessedIdText = "";
    pendingArticleNumber = "";
    pendingCount = 0;
    activeState = emptyActiveLineState();
  }

  let remainingRows = Math.max(1, Math.floor(opts.maxRows));
  let tableMissing = false;
  while (remainingRows > 0) {
    const fetchLimit = Math.max(1, Math.min(Math.max(1, opts.batchSize), remainingRows));
    let rows: JsonObject[] = [];
    try {
      rows = await fetchBizerbaRowsChunk(deviceName, lastProcessedId, fetchLimit);
    } catch (error) {
      const message = String((error as { message?: unknown })?.message || "");
      if (message.toLowerCase().includes("relation") && message.includes("Bizerba_ID")) {
        tableMissing = true;
        break;
      }
      if (isLikelyTimeoutError(error)) {
        summary.skipped = true;
        summary.reason = "device_query_timeout";
        summary.error = "Query timeout while scanning Bizerba rows for this device.";
        break;
      }
      throw error;
    }

    if (!rows.length) break;
    remainingRows -= rows.length;
    if (remainingRows <= 0) summary.reachedRowLimit = true;

    for (const row of rows) {
      summary.processedRows += 1;
      const rowIdText = normalizeBigIntString(row.id, lastProcessedId);
      try {
        if (BigInt(rowIdText) > BigInt(lastProcessedId)) lastProcessedId = rowIdText;
      } catch {
        lastProcessedId = rowIdText;
      }

      const rowTimestamp = parseTimestamp(row.Timestamp);
      const rowArticleNumber = String(row.ArticleNumber || "").trim();
      const rowArticleKey = bizerbaArticleKey(rowArticleNumber);
      const rowDate = normalizeBizerbaDate(row.Date, rowTimestamp);
      if (!rowTimestamp || !rowDate) continue;
      if (
        !lastProcessedTimestamp ||
        rowTimestamp.getTime() > lastProcessedTimestamp.getTime() ||
        (rowTimestamp.getTime() === lastProcessedTimestamp.getTime() &&
          rowIdText > lastProcessedIdText)
      ) {
        lastProcessedTimestamp = rowTimestamp;
        lastProcessedIdText = rowIdText;
      }

      if (activeState.runLogId) {
        let gapMinutes = Number.NaN;
        if (activeState.lastTimestamp) {
          gapMinutes =
            (rowTimestamp.getTime() - activeState.lastTimestamp.getTime()) / 60000;
        }
        if (Number.isFinite(gapMinutes) && gapMinutes > bizerbaAutoDowntimeGapMins) {
          const created = await insertDowntimeLog(
            lineId,
            normalizeBizerbaDate(activeState.lastDate, activeState.lastTimestamp),
            activeState.lastTimestamp as Date,
            rowTimestamp,
            activeState.articleNumber,
            deviceName,
            gapMinutes,
          );
          if (created) summary.createdDowntime += 1;
        }

        const isSameArticle = rowArticleKey &&
          rowArticleKey === activeState.articleKey;
        const exceedsRunBreak = Number.isFinite(gapMinutes) &&
          gapMinutes > bizerbaAutoRunBreakGapMins;
        const outOfOrderTimestamp = Number.isFinite(gapMinutes) && gapMinutes < 0;

        if (isSameArticle && !exceedsRunBreak) {
          if (!outOfOrderTimestamp) {
            activeState.lastTimestamp = rowTimestamp;
            activeState.lastDate = rowDate;
          }
          activeState.unitsProduced += 1;
          continue;
        }

        const closed = await closeRunLog(
          activeState.runLogId,
          activeState.lastTimestamp as Date,
          activeState.unitsProduced,
        );
        if (closed) summary.closedRuns += 1;
        activeState = emptyActiveLineState();
      }

      const mappedProduct = rowArticleKey
        ? resolveMappedProduct(opts.productLookup, rowArticleNumber)
        : "";
      if (rowArticleKey) {
        if (mappedProduct) summary.mappedRows += 1;
        else summary.unmappedRows += 1;
      }

      if (!rowArticleKey || !mappedProduct) {
        pendingArticleNumber = "";
        pendingCount = 0;
        continue;
      }

      if (bizerbaArticleKey(pendingArticleNumber) === rowArticleKey) {
        pendingCount += 1;
      } else {
        pendingArticleNumber = rowArticleNumber;
        pendingCount = 1;
      }

      if (pendingCount < bizerbaAutoMinRunSamples) continue;

      const runInsert = await upsertRunLog(
        lineId,
        rowDate,
        mappedProduct,
        rowTimestamp,
        1,
        buildBizerbaAutoRunNote(deviceName, rowArticleNumber, rowTimestamp),
      );
      const runLogId = String(runInsert.runLogId || "").trim();
      if (runLogId) {
        activeState = {
          runLogId,
          articleNumber: rowArticleNumber,
          articleKey: rowArticleKey,
          product: mappedProduct,
          runDate: rowDate,
          startTimestamp: rowTimestamp,
          lastDate: rowDate,
          lastTimestamp: rowTimestamp,
          unitsProduced: 1,
        };
        if (runInsert.created) summary.createdRuns += 1;
      } else {
        activeState = emptyActiveLineState();
      }
      pendingArticleNumber = "";
      pendingCount = 0;
    }

    if (rows.length < fetchLimit) break;
  }

  if (tableMissing) {
    summary.skipped = true;
    summary.reason = "bizerba_table_missing";
  } else if (!summary.reason && summary.processedRows === 0) {
    summary.reason = "no_new_rows";
  } else if (!summary.reason && summary.createdRuns === 0) {
    if (summary.mappedRows === 0 && summary.unmappedRows > 0) {
      summary.reason = "no_mapped_articles";
    } else if (summary.mappedRows > 0) {
      summary.reason = "run_start_threshold_not_reached";
    }
  }

  if (activeState.runLogId && activeState.lastTimestamp) {
    const idleMinutes = (Date.now() - activeState.lastTimestamp.getTime()) / 60000;
    if (Number.isFinite(idleMinutes) && idleMinutes > bizerbaAutoRunBreakGapMins) {
      const closed = await closeRunLog(
        activeState.runLogId,
        activeState.lastTimestamp,
        activeState.unitsProduced,
      );
      if (closed) summary.closedRuns += 1;
      activeState = emptyActiveLineState();
    }
  }

  await saveLineStateRow(lineId, {
    device_name: deviceName,
    last_processed_bizerba_id: normalizeBigIntString(lastProcessedId, "0"),
    last_processed_timestamp: lastProcessedTimestamp
      ? lastProcessedTimestamp.toISOString()
      : null,
    last_processed_bizerba_id_text: lastProcessedIdText,
    pending_article_number: pendingArticleNumber,
    pending_count: Math.max(0, pendingCount),
    active_run_log_id: activeState.runLogId || null,
    active_article_number: activeState.articleNumber || "",
    active_product: activeState.product || "",
    active_run_date: activeState.runDate || null,
    active_start_timestamp: activeState.startTimestamp
      ? activeState.startTimestamp.toISOString()
      : null,
    active_last_date: activeState.lastDate || null,
    active_last_timestamp: activeState.lastTimestamp
      ? activeState.lastTimestamp.toISOString()
      : null,
    active_units_produced: clampNonNegativeNumber(activeState.unitsProduced, 0),
    updated_at: new Date().toISOString(),
  });

  return summary;
}

async function claimJob(
  workerId: string,
  leaseSeconds: number,
  targetJobId = "",
): Promise<JsonObject | null> {
  const safeTargetJobId = String(targetJobId || "").trim();
  if (safeTargetJobId) {
    const { data: existing, error: fetchError } = await supabase
      .from("bizerba_backsync_jobs")
      .select("*")
      .eq("id", safeTargetJobId)
      .maybeSingle();
    if (fetchError) throw fetchError;
    if (!existing) return null;
    const status = String(existing.status || "").trim().toLowerCase();
    if (status === "completed" || status === "failed") return existing as JsonObject;
    const nowIso = new Date().toISOString();
    const { data: claimed, error: claimError } = await supabase
      .from("bizerba_backsync_jobs")
      .update({
        status: "running",
        worker_id: workerId,
        lease_expires_at: new Date(Date.now() + (leaseSeconds * 1000)).toISOString(),
        started_at: existing.started_at || nowIso,
        updated_at: nowIso,
      })
      .eq("id", safeTargetJobId)
      .in("status", ["queued", "running"])
      .select("*")
      .maybeSingle();
    if (claimError) throw claimError;
    return (claimed as JsonObject | null) || (existing as JsonObject);
  }

  const { data, error } = await supabase.rpc("claim_bizerba_backsync_job", {
    p_worker_id: workerId,
    p_lease_seconds: leaseSeconds,
  });
  if (error) throw error;
  if (!Array.isArray(data) || !data.length) return null;
  return (data[0] as JsonObject) || null;
}

async function fetchActiveLines(): Promise<Array<{ lineId: string; deviceName: string }>> {
  const { data, error } = await supabase
    .from("production_lines")
    .select("id,device_name,created_at")
    .eq("is_active", true)
    .not("device_name", "is", null)
    .order("created_at", { ascending: true })
    .order("id", { ascending: true });
  if (error) throw error;
  const rows = Array.isArray(data) ? (data as JsonObject[]) : [];
  return rows
    .map((row) => ({
      lineId: String(row.id || "").trim(),
      deviceName: String(row.device_name || "").trim(),
    }))
    .filter((row) => row.lineId && row.deviceName);
}

async function fetchCatalogEntries(): Promise<ProductCatalogEntry[]> {
  const { data, error } = await supabase
    .from("product_catalog_entries")
    .select("id,created_at,all_lines,line_ids,catalog_values")
    .order("created_at", { ascending: true })
    .order("id", { ascending: true });
  if (error) throw error;
  const rows = Array.isArray(data) ? (data as JsonObject[]) : [];
  return mapCatalogEntries(rows);
}

function ensureSummaryShape(summary: JsonObject, job: JsonObject): JsonObject {
  const next = { ...summary };
  next.trigger = String(next.trigger || job.trigger || "manual").trim() || "manual";
  next.fullRescan = Boolean(
    next.fullRescan === undefined ? job.full_rescan : next.fullRescan,
  );
  next.startedAt = String(next.startedAt || job.started_at || "");
  next.finishedAt = String(next.finishedAt || "");
  next.lineCount = clampNonNegativeInteger(next.lineCount, 0);
  next.processedRows = clampNonNegativeInteger(next.processedRows, 0);
  next.mappedRows = clampNonNegativeInteger(next.mappedRows, 0);
  next.unmappedRows = clampNonNegativeInteger(next.unmappedRows, 0);
  next.createdRuns = clampNonNegativeInteger(next.createdRuns, 0);
  next.closedRuns = clampNonNegativeInteger(next.closedRuns, 0);
  next.createdDowntime = clampNonNegativeInteger(next.createdDowntime, 0);
  next.skippedLines = clampNonNegativeInteger(next.skippedLines, 0);
  next.erroredLines = clampNonNegativeInteger(next.erroredLines, 0);
  if (!Array.isArray(next.lines)) next.lines = [];
  return next;
}

function updateLineSummaryAggregate(
  summary: JsonObject,
  lineSummary: LineProcessSummary,
): void {
  const existingLines = Array.isArray(summary.lines) ? summary.lines : [];
  const lineMap = new Map<string, JsonObject>();
  existingLines.forEach((item) => {
    if (!item || typeof item !== "object" || Array.isArray(item)) return;
    const safe = item as JsonObject;
    const lineId = String(safe.lineId || "").trim();
    if (!lineId) return;
    lineMap.set(lineId, { ...safe });
  });

  const current = lineMap.get(lineSummary.lineId) || {
    lineId: lineSummary.lineId,
    deviceName: lineSummary.deviceName,
    processedRows: 0,
    mappedRows: 0,
    unmappedRows: 0,
    createdRuns: 0,
    closedRuns: 0,
    createdDowntime: 0,
    skipped: false,
    reason: "",
    error: "",
  };

  current.deviceName = lineSummary.deviceName;
  current.processedRows = clampNonNegativeInteger(current.processedRows, 0) +
    clampNonNegativeInteger(lineSummary.processedRows, 0);
  current.mappedRows = clampNonNegativeInteger(current.mappedRows, 0) +
    clampNonNegativeInteger(lineSummary.mappedRows, 0);
  current.unmappedRows = clampNonNegativeInteger(current.unmappedRows, 0) +
    clampNonNegativeInteger(lineSummary.unmappedRows, 0);
  current.createdRuns = clampNonNegativeInteger(current.createdRuns, 0) +
    clampNonNegativeInteger(lineSummary.createdRuns, 0);
  current.closedRuns = clampNonNegativeInteger(current.closedRuns, 0) +
    clampNonNegativeInteger(lineSummary.closedRuns, 0);
  current.createdDowntime = clampNonNegativeInteger(current.createdDowntime, 0) +
    clampNonNegativeInteger(lineSummary.createdDowntime, 0);
  current.skipped = Boolean(current.skipped) || Boolean(lineSummary.skipped);
  if (lineSummary.reason) current.reason = lineSummary.reason;
  if (lineSummary.error) current.error = lineSummary.error;
  lineMap.set(lineSummary.lineId, current);

  summary.lines = Array.from(lineMap.values());
}

async function persistJob(
  jobId: string,
  patch: JsonObject,
): Promise<JsonObject | null> {
  const { data, error } = await supabase
    .from("bizerba_backsync_jobs")
    .update(patch)
    .eq("id", jobId)
    .select("*")
    .maybeSingle();
  if (error) throw error;
  return data as JsonObject | null;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return optionsResponse();
  if (req.method !== "POST") return jsonResponse({ error: "Method not allowed" }, 405);
  if (!isAuthorized(req)) return unauthorizedResponse();

  const body = await parseJsonBody(req);
  const targetJobId = String(body.jobId || "").trim();
  const maxSeconds = Math.max(5, Math.min(120, clampNonNegativeInteger(body.maxSeconds, 20)));
  const batchSize = Math.max(100, Math.min(5000, clampNonNegativeInteger(body.batchSize, 1000)));
  const maxRowsPerInvocation = Math.max(
    batchSize,
    Math.min(100000, clampNonNegativeInteger(body.maxRowsPerInvocation, batchSize * 20)),
  );
  const leaseSeconds = Math.max(30, Math.min(600, clampNonNegativeInteger(body.leaseSeconds, 120)));
  const workerId = `edge-${crypto.randomUUID()}`;
  const startedMs = Date.now();

  let claimedJob: JsonObject | null = null;
  try {
    claimedJob = await claimJob(workerId, leaseSeconds, targetJobId);
    if (!claimedJob) {
      return jsonResponse({ ok: true, idle: true }, 200);
    }

    const status = String(claimedJob.status || "").trim().toLowerCase();
    if (status === "completed" || status === "failed") {
      return jsonResponse({ ok: true, idle: true, job: mapJobForApi(claimedJob as any) }, 200);
    }

    const lines = await fetchActiveLines();
    const catalogEntries = await fetchCatalogEntries();

    const summary = ensureSummaryShape(
      sanitizeSummary(claimedJob.summary),
      claimedJob,
    );
    const state = sanitizeState(claimedJob.state);
    summary.startedAt = String(summary.startedAt || claimedJob.started_at || new Date().toISOString());
    summary.lineCount = lines.length;

    if (!lines.length) {
      summary.finishedAt = new Date().toISOString();
      const completed = await persistJob(String(claimedJob.id || ""), {
        status: "completed",
        summary,
        state,
        error: "",
        lease_expires_at: null,
        finished_at: new Date().toISOString(),
        updated_at: new Date().toISOString(),
      });
      return jsonResponse({ ok: true, idle: false, completed: true, job: mapJobForApi(completed as any) }, 200);
    }

    let lineCursor = clampNonNegativeInteger(state.lineCursor, 0);
    let passVisited = clampNonNegativeInteger(state.passVisited, 0);
    let passNoRows = clampNonNegativeInteger(state.passNoRows, 0);
    let passErrors = clampNonNegativeInteger(state.passErrors, 0);
    const resetLineIds = new Set<string>(
      Array.isArray(state.resetLineIds)
        ? state.resetLineIds.map((value) => String(value || "").trim()).filter(Boolean)
        : [],
    );

    const lookupCache = new Map<string, Map<string, string>>();
    let remainingRows = maxRowsPerInvocation;
    let completed = false;

    while (Date.now() - startedMs < maxSeconds * 1000 && remainingRows > 0) {
      if (!lines.length) {
        completed = true;
        break;
      }
      const index = lineCursor % lines.length;
      const line = lines[index];
      const lookup = lookupCache.get(line.lineId) ||
        buildProductLookupForLine(line.lineId, catalogEntries);
      lookupCache.set(line.lineId, lookup);

      const shouldReset = Boolean(claimedJob.full_rescan) && !resetLineIds.has(line.lineId);
      const lineSummary = await processLineChunk(
        line,
        {
          fullRescan: Boolean(claimedJob.full_rescan),
          resetState: shouldReset,
          batchSize,
          maxRows: Math.min(remainingRows, Math.max(batchSize, batchSize * 5)),
          productLookup: lookup,
        },
      );

      if (shouldReset) resetLineIds.add(line.lineId);
      summary.processedRows = clampNonNegativeInteger(summary.processedRows, 0) + lineSummary.processedRows;
      summary.mappedRows = clampNonNegativeInteger(summary.mappedRows, 0) + lineSummary.mappedRows;
      summary.unmappedRows = clampNonNegativeInteger(summary.unmappedRows, 0) + lineSummary.unmappedRows;
      summary.createdRuns = clampNonNegativeInteger(summary.createdRuns, 0) + lineSummary.createdRuns;
      summary.closedRuns = clampNonNegativeInteger(summary.closedRuns, 0) + lineSummary.closedRuns;
      summary.createdDowntime = clampNonNegativeInteger(summary.createdDowntime, 0) + lineSummary.createdDowntime;
      if (lineSummary.skipped) summary.skippedLines = clampNonNegativeInteger(summary.skippedLines, 0) + 1;
      if (lineSummary.error) summary.erroredLines = clampNonNegativeInteger(summary.erroredLines, 0) + 1;
      updateLineSummaryAggregate(summary, lineSummary);

      remainingRows -= Math.max(0, lineSummary.processedRows);
      passVisited += 1;
      if (lineSummary.processedRows === 0) passNoRows += 1;
      if (lineSummary.error) passErrors += 1;
      lineCursor = (index + 1) % lines.length;

      if (passVisited >= lines.length) {
        if (passNoRows >= lines.length && passErrors === 0) {
          completed = true;
          break;
        }
        passVisited = 0;
        passNoRows = 0;
        passErrors = 0;
      }
    }

    state.lineCursor = lineCursor;
    state.passVisited = passVisited;
    state.passNoRows = passNoRows;
    state.passErrors = passErrors;
    state.resetLineIds = Array.from(resetLineIds.values());

    const patch: JsonObject = {
      status: completed ? "completed" : "running",
      summary,
      state,
      worker_id: completed ? "" : workerId,
      error: "",
      lease_expires_at: completed
        ? null
        : new Date(Date.now() + (leaseSeconds * 1000)).toISOString(),
      updated_at: new Date().toISOString(),
    };
    if (completed) {
      summary.finishedAt = new Date().toISOString();
      patch.finished_at = summary.finishedAt;
    }

    const persisted = await persistJob(String(claimedJob.id || ""), patch);
    return jsonResponse({
      ok: true,
      idle: false,
      completed,
      job: mapJobForApi((persisted || claimedJob) as any),
    }, 200);
  } catch (error) {
    console.error("bizerba-backsync-worker failed", error);
    const message = String((error as { message?: unknown })?.message || error || "Backsync worker failed");
    if (claimedJob?.id) {
      try {
        await persistJob(String(claimedJob.id || ""), {
          status: "failed",
          error: message,
          worker_id: "",
          lease_expires_at: null,
          finished_at: new Date().toISOString(),
          updated_at: new Date().toISOString(),
        });
      } catch (persistError) {
        console.error("bizerba-backsync-worker could not persist failure state", persistError);
      }
    }
    return jsonResponse({ error: message }, 500);
  }
});
