import { NextRequest, NextResponse } from "next/server";

export const dynamic = "force-dynamic";
export const runtime = "nodejs";

type EcountRequest = {
  comCode?: string;
  userId?: string;
  apiCertKey?: string;
  prodCode?: string;
  syncToDailySales?: boolean;
};

type JsonRecord = Record<string, unknown>;

type ItemMasterRecord = {
  itemCode: string;
  itemName: string;
  category: string;
  supplier?: string;
  specification?: string;
  packQuantity?: string;
  stockUnit?: string;
  storageMethod?: string;
  source?: "initial" | "sales";
  firstSeenMonth?: string;
  active?: boolean;
  memo?: string;
};

const TEST_API_BASE = "https://sboapi.ecount.com";
const TEST_ZONE_URL = `${TEST_API_BASE}/OAPI/V2/Zone`;
const ITEM_MASTER_KEY = "ablab_item_masters_v1";

function asRecord(value: unknown): JsonRecord | null {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as JsonRecord)
    : null;
}

function getNestedRecord(source: unknown, ...keys: string[]): JsonRecord | null {
  let current: unknown = source;
  for (const key of keys) {
    const record = asRecord(current);
    if (!record) return null;
    current = record[key];
  }
  return asRecord(current);
}

async function readJsonResponse(response: Response): Promise<unknown> {
  const text = await response.text();
  try {
    return JSON.parse(text);
  } catch {
    return { rawText: text };
  }
}

function maskSessionId(sessionId: string): string {
  return sessionId.length > 12
    ? `${sessionId.slice(0, 6)}...${sessionId.slice(-4)}`
    : "발급됨";
}

function safeLoginResult(value: unknown): unknown {
  const root = asRecord(value);
  if (!root) return value;

  const cloned = structuredClone(root) as JsonRecord;
  const data = getNestedRecord(cloned, "Data");
  const datas = data ? asRecord(data.Datas) : null;
  if (datas && typeof datas.SESSION_ID === "string") {
    datas.SESSION_ID = maskSessionId(datas.SESSION_ID);
  }
  return cloned;
}

function getEcountMessage(value: unknown): string {
  const data = getNestedRecord(value, "Data");
  if (typeof data?.Message === "string" && data.Message.trim()) return data.Message;

  const error = getNestedRecord(value, "Error");
  if (typeof error?.Message === "string" && error.Message.trim()) return error.Message;

  return "";
}

function textValue(record: JsonRecord, keys: string[]): string {
  for (const key of keys) {
    const value = record[key];
    if (value !== null && value !== undefined && String(value).trim()) {
      return String(value).trim();
    }
  }
  return "";
}

function normalizeItem(value: unknown): ItemMasterRecord | null {
  const row = asRecord(value);
  if (!row) return null;

  const itemCode = textValue(row, ["PROD_CD", "PROD_CODE", "ITEM_CD", "ITEM_CODE", "prodCd", "itemCode"]);
  if (!itemCode) return null;

  return {
    itemCode,
    itemName: textValue(row, ["PROD_DES", "PROD_NM", "PROD_NAME", "ITEM_NM", "ITEM_NAME", "prodDes", "itemName"]) || itemCode,
    category: textValue(row, ["PROD_GROUP_DES", "PROD_GROUP_NAME", "GROUP_DES", "GROUP_NAME", "CATEGORY", "CATEGORY_NAME"]) || "미지정",
    supplier: textValue(row, ["CUST_DES", "CUST_NAME", "SUPPLIER_DES", "SUPPLIER_NAME", "PUR_CUST_DES", "PUR_CUST_NAME", "VENDOR_NAME"]) || "미지정",
    specification: textValue(row, ["SIZE_DES", "SPEC", "SPECIFICATION", "PROD_SIZE", "SIZE"]),
    packQuantity: textValue(row, ["IN_QTY", "PACK_QTY", "BOX_QTY", "UNIT_QTY"]),
    stockUnit: textValue(row, ["UNIT", "UNIT_NM", "STOCK_UNIT", "INVEN_UNIT"]),
    storageMethod: textValue(row, ["WH_DES", "STORAGE_METHOD", "STORAGE_METHOD_NAME", "KEEP_TYPE_DES"]),
    source: "initial",
    active: textValue(row, ["USE_YN", "ACTIVE_YN", "IS_ACTIVE"]).toUpperCase() !== "N",
    memo: textValue(row, ["REMARKS", "REMARK", "MEMO", "CONTENTS"]),
  };
}

function extractItemArray(itemJson: unknown): unknown[] {
  const data = getNestedRecord(itemJson, "Data");
  if (!data) return [];

  const candidates = [data.Result, data.Datas, data.RESULT, data.DATA];
  for (const candidate of candidates) {
    if (Array.isArray(candidate)) return candidate;
    const record = asRecord(candidate);
    if (record) {
      for (const key of ["Result", "Datas", "List", "Items", "ITEMS"]) {
        if (Array.isArray(record[key])) return record[key] as unknown[];
      }
    }
  }
  return [];
}

function workerConfig() {
  const url = process.env.D1_WORKER_URL?.trim().replace(/\/$/, "");
  const apiKey = process.env.D1_API_KEY?.trim();
  if (!url || !apiKey) return null;
  return { url, apiKey };
}

async function loadExistingItemMasters(): Promise<ItemMasterRecord[]> {
  const config = workerConfig();
  if (!config) throw new Error("D1_WORKER_URL 또는 D1_API_KEY 환경변수가 없습니다.");

  const response = await fetch(`${config.url}/settings/${encodeURIComponent(ITEM_MASTER_KEY)}`, {
    method: "GET",
    headers: { "X-ABL-API-Key": config.apiKey },
    cache: "no-store",
  });

  if (response.status === 404) return [];
  if (!response.ok) throw new Error(`기존 품목기준정보 조회 실패 (${response.status})`);

  const json = (await response.json()) as { data?: unknown };
  return Array.isArray(json.data) ? (json.data as ItemMasterRecord[]) : [];
}

async function saveItemMasters(items: ItemMasterRecord[]) {
  const config = workerConfig();
  if (!config) throw new Error("D1_WORKER_URL 또는 D1_API_KEY 환경변수가 없습니다.");

  const response = await fetch(`${config.url}/settings/${encodeURIComponent(ITEM_MASTER_KEY)}`, {
    method: "PUT",
    headers: {
      "X-ABL-API-Key": config.apiKey,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ data: items }),
    cache: "no-store",
  });

  const body = await response.text();
  if (!response.ok) throw new Error(`품목기준정보 저장 실패 (${response.status}) ${body.slice(0, 200)}`);
  try {
    return JSON.parse(body) as { updated_at?: string };
  } catch {
    return {};
  }
}

function mergeItemMasters(existing: ItemMasterRecord[], incoming: ItemMasterRecord[]) {
  const map = new Map(existing.map((item) => [item.itemCode, item]));
  let added = 0;
  let updated = 0;

  for (const item of incoming) {
    const previous = map.get(item.itemCode);
    if (!previous) {
      map.set(item.itemCode, item);
      added += 1;
      continue;
    }

    map.set(item.itemCode, {
      ...previous,
      ...item,
      itemName: item.itemName || previous.itemName,
      category: item.category !== "미지정" ? item.category : previous.category || "미지정",
      supplier: item.supplier !== "미지정" ? item.supplier : previous.supplier || "미지정",
      specification: item.specification || previous.specification,
      packQuantity: item.packQuantity || previous.packQuantity,
      stockUnit: item.stockUnit || previous.stockUnit,
      storageMethod: item.storageMethod || previous.storageMethod,
      memo: item.memo || previous.memo,
      source: previous.source || "initial",
    });
    updated += 1;
  }

  return { merged: Array.from(map.values()), added, updated };
}

export async function POST(request: NextRequest) {
  try {
    const body = (await request.json()) as EcountRequest;
    const comCode = body.comCode?.trim();
    const userId = body.userId?.trim();
    const apiCertKey = body.apiCertKey?.trim();
    const prodCode = body.prodCode?.trim() ?? "";
    const syncToDailySales = body.syncToDailySales === true;

    if (!comCode || !userId || !apiCertKey) {
      return NextResponse.json(
        { ok: false, message: "회사코드, 사용자 ID, API 인증키를 모두 입력해 주세요." },
        { status: 400 },
      );
    }

    const zoneResponse = await fetch(TEST_ZONE_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ COM_CODE: comCode }),
      cache: "no-store",
    });
    const zoneJson = await readJsonResponse(zoneResponse);
    const zoneData = getNestedRecord(zoneJson, "Data");
    const rawZone = typeof zoneData?.ZONE === "string" ? zoneData.ZONE.trim() : "";
    const rawDomain = typeof zoneData?.DOMAIN === "string" ? zoneData.DOMAIN.trim() : "ecount.com";
    const zone = rawZone.replace(/^\.+|\.+$/g, "");
    const domain = rawDomain.replace(/^\.+|\.+$/g, "") || "ecount.com";

    if (!zoneResponse.ok || !zone) {
      return NextResponse.json(
        {
          ok: false,
          step: "zone",
          message: getEcountMessage(zoneJson) || "테스트 Zone 확인에 실패했습니다.",
          environment: "test",
          zoneHttpStatus: zoneResponse.status,
          zoneResponse: zoneJson,
        },
        { status: 502 },
      );
    }

    const loginUrl = `https://sboapi${zone}.${domain}/OAPI/V2/OAPILogin`;
    const loginResponse = await fetch(loginUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        COM_CODE: comCode,
        USER_ID: userId,
        API_CERT_KEY: apiCertKey,
        LAN_TYPE: "ko-KR",
        ZONE: zone,
      }),
      cache: "no-store",
    });
    const loginJson = await readJsonResponse(loginResponse);
    const loginDatas = getNestedRecord(loginJson, "Data", "Datas");
    const sessionId = typeof loginDatas?.SESSION_ID === "string" ? loginDatas.SESSION_ID : "";

    if (!loginResponse.ok || !sessionId) {
      return NextResponse.json(
        {
          ok: false,
          step: "login",
          message: getEcountMessage(loginJson) || "테스트 로그인 또는 SESSION_ID 발급에 실패했습니다.",
          environment: "test",
          zone: { zone, domain },
          loginHttpStatus: loginResponse.status,
          loginResponse: safeLoginResult(loginJson),
        },
        { status: 502 },
      );
    }

    const itemUrl = `${TEST_API_BASE}/OAPI/V2/InventoryBasic/GetBasicProductList?SESSION_ID=${encodeURIComponent(sessionId)}`;
    const itemResponse = await fetch(itemUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ PROD_CD: prodCode, PROD_TYPE: "0" }),
      cache: "no-store",
    });
    const itemJson = await readJsonResponse(itemResponse);
    const rawItems = extractItemArray(itemJson);
    const normalizedItems = rawItems.map(normalizeItem).filter((item): item is ItemMasterRecord => Boolean(item));
    const itemData = getNestedRecord(itemJson, "Data");
    const totalCountRaw = itemData?.TotalCnt ?? itemData?.TOTAL_CNT ?? itemData?.TotalCount;
    const totalCount = Number.isFinite(Number(totalCountRaw)) ? Number(totalCountRaw) : normalizedItems.length;

    if (!itemResponse.ok) {
      return NextResponse.json(
        {
          ok: false,
          step: "items",
          message: getEcountMessage(itemJson) || "SESSION_ID는 발급됐지만 품목조회에 실패했습니다.",
          environment: "test",
          zone: { zone, domain },
          login: { sessionIssued: true, maskedSessionId: maskSessionId(sessionId) },
          itemUrl,
          itemHttpStatus: itemResponse.status,
          itemResponse: itemJson,
        },
        { status: 502 },
      );
    }

    let sync: Record<string, unknown> | undefined;
    if (syncToDailySales) {
      if (normalizedItems.length === 0) {
        return NextResponse.json(
          {
            ok: false,
            step: "sync",
            message: "품목조회 응답은 성공했지만 저장할 품목이 없습니다. 원본 응답을 확인해 주세요.",
            zone: { zone, domain },
            login: { sessionIssued: true, maskedSessionId: maskSessionId(sessionId) },
            items: { totalCount, returnedCount: 0, preview: [] },
            raw: { items: itemJson },
          },
          { status: 502 },
        );
      }

      const existing = await loadExistingItemMasters();
      const mergedResult = mergeItemMasters(existing, normalizedItems);
      const saved = await saveItemMasters(mergedResult.merged);
      sync = {
        saved: true,
        key: ITEM_MASTER_KEY,
        addedCount: mergedResult.added,
        updatedCount: mergedResult.updated,
        totalSavedCount: mergedResult.merged.length,
        updatedAt: saved.updated_at || new Date().toISOString(),
      };
    }

    return NextResponse.json({
      ok: true,
      message: syncToDailySales
        ? "이카운트 품목을 Daily Sales 품목기준정보에 반영했습니다."
        : "테스트 로그인과 품목조회에 성공했습니다.",
      environment: "test",
      zone: { zone, domain },
      login: { sessionIssued: true, maskedSessionId: maskSessionId(sessionId) },
      items: {
        requestedProdCode: prodCode,
        totalCount,
        returnedCount: normalizedItems.length,
        preview: normalizedItems.slice(0, 50),
      },
      sync,
      raw: {
        zone: zoneJson,
        login: safeLoginResult(loginJson),
        items: itemJson,
      },
    });
  } catch (error) {
    console.error("ECOUNT API test failed:", error);
    return NextResponse.json(
      {
        ok: false,
        step: "server",
        message: error instanceof Error ? error.message : "알 수 없는 서버 오류가 발생했습니다.",
      },
      { status: 500 },
    );
  }
}
