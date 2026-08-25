import webpush from "web-push";

interface Env {
  DB: D1Database;
  ABL_API_KEY: string;
  VAPID_PUBLIC_KEY: string;
  VAPID_PRIVATE_KEY: string;
  VAPID_SUBJECT?: string;
}

type PeriodType = "current" | "prevMonth" | "prevYear";
type SalesRow = {
  id: string;
  period: PeriodType;
  refMonth: string;
  saleDate: string;
  storeCode: string;
  storeName: string;
  channel: string;
  manager: string;
  storeType: string;
  brand: string;
  itemCode: string;
  itemName: string;
  quantity: number;
  salesAmount: number;
  costAmount: number;
  profitAmount: number;
  profitRate: number;
};

type ReplacePayload = {
  period: PeriodType;
  refMonth: string;
  fileName: string;
  uploadedDates: string[];
  rows: SalesRow[];
};

const jsonHeaders = { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" };
const json = (data: unknown, status = 200) => new Response(JSON.stringify(data), { status, headers: jsonHeaders });

function authorized(request: Request, env: Env) {
  const supplied = request.headers.get("X-ABL-API-Key") || "";
  return Boolean(env.ABL_API_KEY) && supplied === env.ABL_API_KEY;
}

function batchId() {
  return `batch_${Date.now()}_${crypto.randomUUID()}`;
}

function chunks<T>(items: T[], size: number) {
  const result: T[][] = [];
  for (let i = 0; i < items.length; i += size) result.push(items.slice(i, i + size));
  return result;
}

async function getSetting(env: Env, key: string) {
  const row = await env.DB.prepare("SELECT key, data, updated_at FROM app_settings WHERE key = ? LIMIT 1").bind(key).first<{ key: string; data: string; updated_at: string }>();
  if (!row) return json(null, 404);
  let data: unknown = null;
  try { data = JSON.parse(row.data); } catch { data = row.data; }
  return json({ id: row.key, data, updated_at: row.updated_at });
}

async function putSetting(request: Request, env: Env, key: string) {
  const payload = await request.json<{ data: unknown }>();
  const updatedAt = new Date().toISOString();
  await env.DB.prepare(
    `INSERT INTO app_settings (key, data, updated_at) VALUES (?, ?, ?)
     ON CONFLICT(key) DO UPDATE SET data = excluded.data, updated_at = excluded.updated_at`,
  ).bind(key, JSON.stringify(payload.data ?? null), updatedAt).run();
  return json({ id: key, updated_at: updatedAt });
}

async function getSales(url: URL, env: Env) {
  const baseMonth = url.searchParams.get("baseMonth") || "";
  if (!/^\d{4}-\d{2}$/.test(baseMonth)) return json({ error: "Invalid baseMonth" }, 400);
  const requestedPeriod = url.searchParams.get("period") || "";
  if (requestedPeriod && !["current", "prevMonth", "prevYear"].includes(requestedPeriod)) {
    return json({ error: "Invalid period" }, 400);
  }

  const recordsSql =
    `SELECT r.id AS row_key, r.period_type AS period, r.base_month AS ref_month,
      r.sales_date AS sale_date, r.customer_code AS store_code, r.customer_name AS store_name,
      r.manager, r.store_type, r.brand, r.item_code, r.item_name, r.category,
      r.quantity, r.sales_amount, r.purchase_amount AS cost_amount,
      r.profit_amount,
      CASE WHEN r.sales_amount != 0 THEN (r.profit_amount / r.sales_amount) * 100 ELSE 0 END AS profit_rate,
      '' AS channel
     FROM sales_records r
     JOIN sales_upload_batches b ON b.id = r.batch_id
     WHERE r.base_month = ? AND b.status = 'success'
       ${requestedPeriod ? "AND r.period_type = ?" : ""}
     ORDER BY r.sales_date, r.id`;
  const recordsStatement = env.DB.prepare(recordsSql);
  const result = requestedPeriod
    ? await recordsStatement.bind(baseMonth, requestedPeriod).all()
    : await recordsStatement.bind(baseMonth).all();

  const batchesSql =
    `SELECT period_type AS period, MAX(completed_at) AS updated_at
     FROM sales_upload_batches
     WHERE base_month = ? AND status = 'success'
       ${requestedPeriod ? "AND period_type = ?" : ""}
     GROUP BY period_type`;
  const batchesStatement = env.DB.prepare(batchesSql);
  const batches = requestedPeriod
    ? await batchesStatement.bind(baseMonth, requestedPeriod).all()
    : await batchesStatement.bind(baseMonth).all();

  return json({ available: true, records: result.results || [], batches: batches.results || [] });
}

async function getPriorYearStoreHistory(url: URL, env: Env) {
  const baseMonth = url.searchParams.get("baseMonth") || "";
  if (!/^\d{4}-\d{2}$/.test(baseMonth)) return json({ error: "Invalid baseMonth" }, 400);

  const [yearText, monthText] = baseMonth.split("-");
  const monthNumber = Number(monthText);
  if (monthNumber < 1 || monthNumber > 12) return json({ error: "Invalid baseMonth" }, 400);

  const rangeStart = `${yearText}-01-01`;
  const rangeEnd = `${baseMonth}-01`;
  if (monthNumber === 1) {
    return json({
      available: true,
      baseMonth,
      rangeStart,
      rangeEnd,
      stores: [],
    });
  }

  const result = await env.DB.prepare(
    `SELECT r.customer_code AS store_code,
            MAX(r.customer_name) AS store_name,
            MIN(r.sales_date) AS first_sale_date
       FROM sales_records r
       JOIN sales_upload_batches b ON b.id = r.batch_id
      WHERE r.period_type = 'current'
        AND b.status = 'success'
        AND r.sales_date >= ?
        AND r.sales_date < ?
      GROUP BY r.customer_code`,
  ).bind(rangeStart, rangeEnd).all();

  return json({
    available: true,
    baseMonth,
    rangeStart,
    rangeEnd,
    stores: result.results || [],
  });
}

async function replaceSales(request: Request, env: Env) {
  const payload = await request.json<ReplacePayload>();
  if (!payload || !["current", "prevMonth", "prevYear"].includes(payload.period) || !/^\d{4}-\d{2}$/.test(payload.refMonth) || !Array.isArray(payload.rows)) {
    return json({ error: "Invalid upload payload" }, 400);
  }

  const newBatchId = batchId();
  const createdAt = new Date().toISOString();
  await env.DB.prepare(
    `INSERT INTO sales_upload_batches
     (id, period_type, base_month, file_name, row_count, status, created_at)
     VALUES (?, ?, ?, ?, ?, 'uploading', ?)`,
  ).bind(newBatchId, payload.period, payload.refMonth, payload.fileName || "", payload.rows.length, createdAt).run();

  try {
    const insertSql = `INSERT INTO sales_records
      (batch_id, period_type, base_month, sales_date, customer_code, customer_name,
       manager, store_type, brand, item_code, item_name, category, quantity,
       sales_amount, purchase_unit_price, purchase_amount, profit_amount, created_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;

    for (const group of chunks(payload.rows, 75)) {
      const statements = group.map((row) => {
        const quantity = Number(row.quantity || 0);
        const costAmount = Number(row.costAmount || 0);
        return env.DB.prepare(insertSql).bind(
          newBatchId, payload.period, payload.refMonth, row.saleDate || "",
          row.storeCode || "", row.storeName || "", row.manager || "",
          row.storeType || "", row.brand || "", row.itemCode || "",
          row.itemName || "", "", quantity, Number(row.salesAmount || 0),
          quantity ? costAmount / quantity : 0, costAmount,
          Number(row.profitAmount || 0), createdAt,
        );
      });
      if (statements.length) await env.DB.batch(statements);
    }

    const completedAt = new Date().toISOString();
    const finalStatements: D1PreparedStatement[] = [
      env.DB.prepare("UPDATE sales_upload_batches SET status = 'success', completed_at = ? WHERE id = ?").bind(completedAt, newBatchId),
    ];

    if (payload.period === "current") {
      // 당월 데이터는 매일 누적됩니다. 이번 파일에 포함된 날짜만 이전 데이터에서 교체합니다.
      const uploadedDates = Array.from(new Set((payload.uploadedDates || []).filter(Boolean)));
      for (const saleDate of uploadedDates) {
        finalStatements.push(
          env.DB.prepare(
            `DELETE FROM sales_records
             WHERE period_type = 'current' AND base_month = ? AND sales_date = ? AND batch_id != ?`,
          ).bind(payload.refMonth, saleDate, newBatchId),
        );
      }
    } else {
      // 전월/전년동월 비교 자료는 기준월 단위로 전체 교체합니다.
      const oldBatchRows = await env.DB.prepare(
        `SELECT id FROM sales_upload_batches
         WHERE period_type = ? AND base_month = ? AND status = 'success' AND id != ?`,
      ).bind(payload.period, payload.refMonth, newBatchId).all<{ id: string }>();
      for (const row of oldBatchRows.results || []) {
        finalStatements.push(env.DB.prepare("DELETE FROM sales_records WHERE batch_id = ?").bind(row.id));
        finalStatements.push(env.DB.prepare("UPDATE sales_upload_batches SET status = 'replaced' WHERE id = ?").bind(row.id));
      }
    }
    await env.DB.batch(finalStatements);
    return json({ ok: true, batchId: newBatchId, rowCount: payload.rows.length, completedAt });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    await env.DB.prepare("UPDATE sales_upload_batches SET status = 'failed', error_message = ? WHERE id = ?").bind(message.slice(0, 1000), newBatchId).run();
    return json({ error: "Upload failed", detail: message }, 500);
  }
}

async function deleteSalesDate(request: Request, env: Env) {
  const payload = await request.json<{ refMonth: string; saleDate: string }>();
  if (!payload?.refMonth || !payload?.saleDate) return json({ error: "Invalid payload" }, 400);
  const batchRows = await env.DB.prepare(
    `SELECT id FROM sales_upload_batches WHERE period_type = 'current' AND base_month = ? AND status = 'success'`,
  ).bind(payload.refMonth).all<{ id: string }>();
  const statements: D1PreparedStatement[] = [];
  for (const row of batchRows.results || []) {
    statements.push(env.DB.prepare("DELETE FROM sales_records WHERE batch_id = ? AND sales_date = ?").bind(row.id, payload.saleDate));
  }
  if (statements.length) await env.DB.batch(statements);
  return json({ ok: true, deletedDate: payload.saleDate });
}

async function ensurePushTables(env: Env) {
  await env.DB.batch([
    env.DB.prepare(
      `CREATE TABLE IF NOT EXISTS push_subscriptions (
        endpoint TEXT PRIMARY KEY,
        p256dh TEXT NOT NULL,
        auth TEXT NOT NULL,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
      )`,
    ),
    env.DB.prepare(
      `CREATE TABLE IF NOT EXISTS push_notifications (
        report_date TEXT PRIMARY KEY,
        base_month TEXT NOT NULL,
        sent_count INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL
      )`,
    ),
    env.DB.prepare(
      `CREATE TABLE IF NOT EXISTS push_notifications_v2 (
        notification_key TEXT PRIMARY KEY,
        report_date TEXT NOT NULL,
        base_month TEXT NOT NULL,
        sync_slot TEXT NOT NULL,
        sent_count INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL
      )`,
    ),
  ]);
}

async function savePushSubscription(request: Request, env: Env) {
  const payload = await request.json<{
    subscription?: { endpoint?: string; keys?: { p256dh?: string; auth?: string } };
  }>();
  const endpoint = payload.subscription?.endpoint || "";
  const p256dh = payload.subscription?.keys?.p256dh || "";
  const auth = payload.subscription?.keys?.auth || "";
  if (!endpoint || !p256dh || !auth) return json({ error: "Invalid push subscription" }, 400);
  await ensurePushTables(env);
  const now = new Date().toISOString();
  await env.DB.prepare(
    `INSERT INTO push_subscriptions (endpoint, p256dh, auth, created_at, updated_at)
     VALUES (?, ?, ?, ?, ?)
     ON CONFLICT(endpoint) DO UPDATE SET
       p256dh = excluded.p256dh,
       auth = excluded.auth,
       updated_at = excluded.updated_at`,
  ).bind(endpoint, p256dh, auth, now, now).run();
  return json({ ok: true });
}

async function deletePushSubscription(request: Request, env: Env) {
  const payload = await request.json<{ endpoint?: string }>();
  if (!payload.endpoint) return json({ error: "Missing endpoint" }, 400);
  await ensurePushTables(env);
  await env.DB.prepare("DELETE FROM push_subscriptions WHERE endpoint = ?").bind(payload.endpoint).run();
  return json({ ok: true });
}

async function sendUpdatePush(request: Request, env: Env) {
  const payload = await request.json<{
    baseMonth?: string;
    reportDate?: string;
    syncSlot?: string;
    changedRanges?: string[];
  }>();

  const baseMonth = payload.baseMonth || "";
  const reportDate = payload.reportDate || "";
  const syncSlot = String(payload.syncSlot || "manual").trim();
  const changedRanges = Array.isArray(payload.changedRanges)
    ? payload.changedRanges.map((value) => String(value || "").trim()).filter(Boolean)
    : [];

  if (
    !/^\d{4}-\d{2}$/.test(baseMonth) ||
    !/^\d{4}-\d{2}-\d{2}$/.test(reportDate) ||
    !/^[A-Za-z0-9_-]{1,32}$/.test(syncSlot)
  ) {
    return json({ error: "Invalid update notification payload" }, 400);
  }

  if (!env.VAPID_PUBLIC_KEY || !env.VAPID_PRIVATE_KEY) {
    return json({ error: "VAPID secrets are not configured" }, 503);
  }

  await ensurePushTables(env);

  // 같은 날짜라도 08:00 / 16:30 / manual은 각각 별개의 알림입니다.
  const notificationKey = `${reportDate}:${syncSlot}`;
  const existing = await env.DB.prepare(
    "SELECT notification_key FROM push_notifications_v2 WHERE notification_key = ? LIMIT 1",
  ).bind(notificationKey).first();

  if (existing) {
    return json({ ok: true, duplicate: true, sent: 0, notificationKey });
  }

  const subscriptions = await env.DB.prepare(
    "SELECT endpoint, p256dh, auth FROM push_subscriptions ORDER BY created_at",
  ).all<{ endpoint: string; p256dh: string; auth: string }>();

  webpush.setVapidDetails(
    env.VAPID_SUBJECT || "mailto:admin@ablab.co.kr",
    env.VAPID_PUBLIC_KEY,
    env.VAPID_PRIVATE_KEY,
  );

  const slotLabel =
    syncSlot === "0800"
      ? "오전 8시"
      : syncSlot === "1630"
        ? "오후 4시 30분"
        : "수동";

  const rangeText = changedRanges.length
    ? ` 변경 범위: ${changedRanges.join(", ")}`
    : "";

  const notification = JSON.stringify({
    title: "Sales Report 업데이트 완료",
    body: `${slotLabel} 매출 데이터가 최신 정보로 반영되었습니다.${rangeText}`,
    tag: `sales-report-${reportDate}-${syncSlot}`,
    url: "/",
    syncSlot,
    reportDate,
  });

  let sent = 0;
  const expired: string[] = [];

  for (const row of subscriptions.results || []) {
    try {
      await webpush.sendNotification(
        { endpoint: row.endpoint, keys: { p256dh: row.p256dh, auth: row.auth } },
        notification,
        { TTL: 60 * 60 * 24 },
      );
      sent += 1;
    } catch (error) {
      const statusCode = Number((error as { statusCode?: number }).statusCode || 0);
      if (statusCode === 404 || statusCode === 410) expired.push(row.endpoint);
    }
  }

  if (expired.length) {
    await env.DB.batch(
      expired.map((endpoint) =>
        env.DB.prepare("DELETE FROM push_subscriptions WHERE endpoint = ?").bind(endpoint),
      ),
    );
  }

  await env.DB.prepare(
    `INSERT INTO push_notifications_v2
     (notification_key, report_date, base_month, sync_slot, sent_count, created_at)
     VALUES (?, ?, ?, ?, ?, ?)`,
  ).bind(
    notificationKey,
    reportDate,
    baseMonth,
    syncSlot,
    sent,
    new Date().toISOString(),
  ).run();

  return json({
    ok: true,
    duplicate: false,
    sent,
    removed: expired.length,
    notificationKey,
  });
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    if (!authorized(request, env)) return json({ error: "Unauthorized" }, 401);
    const url = new URL(request.url);
    const path = url.pathname.replace(/\/+$/, "") || "/";
    try {
      if (request.method === "GET" && path === "/health") {
        const row = await env.DB.prepare("SELECT 1 AS ok").first();
        return json({ ok: true, database: "ablab-sales-db", result: row });
      }
      if (path.startsWith("/settings/")) {
        const key = decodeURIComponent(path.slice("/settings/".length));
        if (!key) return json({ error: "Missing key" }, 400);
        if (request.method === "GET") return getSetting(env, key);
        if (request.method === "PUT") return putSetting(request, env, key);
      }
      if (request.method === "GET" && path === "/sales") return getSales(url, env);
      if (request.method === "GET" && path === "/sales/prior-year-store-history") return getPriorYearStoreHistory(url, env);
      if (request.method === "POST" && path === "/sales/replace") return replaceSales(request, env);
      if (request.method === "POST" && path === "/sales/delete-date") return deleteSalesDate(request, env);
      if (request.method === "GET" && path === "/push/public-key") {
        if (!env.VAPID_PUBLIC_KEY) return json({ error: "VAPID public key is not configured" }, 503);
        return json({ publicKey: env.VAPID_PUBLIC_KEY });
      }
      if (request.method === "POST" && path === "/push/subscribe") return savePushSubscription(request, env);
      if (request.method === "DELETE" && path === "/push/subscribe") return deletePushSubscription(request, env);
      if (request.method === "POST" && path === "/push/notify") return sendUpdatePush(request, env);
      return json({ error: "Not found" }, 404);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return json({ error: "Internal error", detail: message }, 500);
    }
  },
};
