import { NextRequest, NextResponse } from "next/server";

export const dynamic = "force-dynamic";
export const runtime = "nodejs";

type EcountRequest = {
  comCode?: string;
  userId?: string;
  apiCertKey?: string;
  prodCode?: string;
};

type JsonRecord = Record<string, unknown>;

// 테스트 인증키는 운영용 oapi가 아니라 사전테스트용 sboapi 주소를 사용해야 합니다.
const TEST_ZONE_URL = "https://sboapi.ecount.com/OAPI/V2/Zone";

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

export async function POST(request: NextRequest) {
  try {
    const body = (await request.json()) as EcountRequest;
    const comCode = body.comCode?.trim();
    const userId = body.userId?.trim();
    const apiCertKey = body.apiCertKey?.trim();
    const prodCode = body.prodCode?.trim() ?? "";

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
          zoneUrl: TEST_ZONE_URL,
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
          zoneUrl: TEST_ZONE_URL,
          loginUrl,
          loginHttpStatus: loginResponse.status,
          loginResponse: safeLoginResult(loginJson),
        },
        { status: 502 },
      );
    }

    const itemUrl = `https://sboapi${zone}.${domain}/OAPI/V2/InventoryBasic/GetBasicProductsList?SESSION_ID=${encodeURIComponent(sessionId)}`;
    const itemResponse = await fetch(itemUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ PROD_CD: prodCode, PROD_TYPE: "0" }),
      cache: "no-store",
    });
    const itemJson = await readJsonResponse(itemResponse);
    const itemData = getNestedRecord(itemJson, "Data");
    const resultValue = itemData?.Result;
    const items = Array.isArray(resultValue) ? resultValue : [];
    const totalCount = typeof itemData?.TotalCnt === "number" ? itemData.TotalCnt : items.length;

    if (!itemResponse.ok) {
      return NextResponse.json(
        {
          ok: false,
          step: "items",
          message: getEcountMessage(itemJson) || "SESSION_ID는 발급됐지만 품목조회에 실패했습니다.",
          environment: "test",
          zone: { zone, domain },
          login: { sessionIssued: true, maskedSessionId: maskSessionId(sessionId) },
          itemHttpStatus: itemResponse.status,
          itemResponse: itemJson,
        },
        { status: 502 },
      );
    }

    return NextResponse.json({
      ok: true,
      message: "테스트 로그인과 품목조회에 성공했습니다.",
      environment: "test",
      zone: { zone, domain },
      login: {
        sessionIssued: true,
        maskedSessionId: maskSessionId(sessionId),
      },
      items: {
        requestedProdCode: prodCode,
        totalCount,
        returnedCount: items.length,
        preview: items.slice(0, 20),
      },
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
