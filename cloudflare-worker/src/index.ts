interface Env {
  DB: D1Database;
  ABL_API_KEY: string;
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

  const result = await env.DB.prepare(
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
     ORDER BY r.sales_date, r.id`,
  ).bind(baseMonth).all();

  const batches = await env.DB.prepare(
    `SELECT period_type AS period, MAX(completed_at) AS updated_at
     FROM sales_upload_batches
     WHERE base_month = ? AND status = 'success'
     GROUP BY period_type`,
  ).bind(baseMonth).all();

  return json({ available: true, records: result.results || [], batches: batches.results || [] });
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
      if (request.method === "POST" && path === "/sales/replace") return replaceSales(request, env);
      if (request.method === "POST" && path === "/sales/delete-date") return deleteSalesDate(request, env);
      return json({ error: "Not found" }, 404);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      return json({ error: "Internal error", detail: message }, 500);
    }
  },
};
