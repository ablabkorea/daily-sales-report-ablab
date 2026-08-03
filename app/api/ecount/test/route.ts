import { NextRequest, NextResponse } from "next/server";

export const dynamic = "force-dynamic";
export const runtime = "nodejs";

type JsonRecord = Record<string, unknown>;

type EcountRequest = {
  comCode?: string;
  userId?: string;
  apiCertKey?: string;
  prodCode?: string;
  syncToDailySales?: boolean;
};

type ItemMasterRecord = {
  itemCode: string;
  itemName: string;
  category: string;
  supplier: string;
  specification: string;
  packQuantity: string;
  stockUnit: string;
  storageMethod: string;
  source: "initial" | "sales";
  firstSeenMonth?: string;
  active: boolean;
  memo: string;
  ecountUpdatedAt: string;
};

type D1SettingResponse = {
  id?: string;
  data?: unknown;
  updated_at?: string;
};

const TEST_ZONE_URL = "https://sboapi.ecount.com/OAPI/V2/Zone";

const ITEM_MASTER_KEY = "ablab_item_masters_v1";
const SYNC_STATUS_KEY = "ecount_sync_status_v1";

function asRecord(value: unknown): JsonRecord | null {
  if (
    value === null ||
    typeof value !== "object" ||
    Array.isArray(value)
  ) {
    return null;
  }

  return value as JsonRecord;
}

function getNestedRecord(
  source: unknown,
  ...keys: string[]
): JsonRecord | null {
  let current: unknown = source;

  for (const key of keys) {
    const record = asRecord(current);

    if (!record) {
      return null;
    }

    current = record[key];
  }

  return asRecord(current);
}

function textValue(
  record: JsonRecord,
  keys: string[]
): string {
  for (const key of keys) {
    const value = record[key];

    if (
      value !== undefined &&
      value !== null &&
      String(value).trim()
    ) {
      return String(value).trim();
    }
  }

  return "";
}

async function readJsonResponse(
  response: Response
): Promise<unknown> {
  const text = await response.text();

  if (!text) {
    return {};
  }

  try {
    return JSON.parse(text);
  } catch {
    return {
      rawText: text,
    };
  }
}

function getEcountMessage(value: unknown): string {
  const root = asRecord(value);
  const data = root ? asRecord(root.Data) : null;
  const error = root ? asRecord(root.Error) : null;

  const candidates = [
    data?.Message,
    data?.MESSAGE,
    error?.Message,
    error?.MESSAGE,
    root?.Message,
    root?.message,
  ];

  for (const candidate of candidates) {
    if (
      typeof candidate === "string" &&
      candidate.trim()
    ) {
      return candidate.trim();
    }
  }

  return "";
}

function maskSessionId(sessionId: string): string {
  if (!sessionId) {
    return "";
  }

  if (sessionId.length <= 12) {
    return "발급됨";
  }

  return `${sessionId.slice(0, 6)}...${sessionId.slice(-4)}`;
}

function safeLoginResult(value: unknown): unknown {
  const root = asRecord(value);

  if (!root) {
    return value;
  }

  const cloned = structuredClone(root) as JsonRecord;
  const data = asRecord(cloned.Data);
  const datas = data ? asRecord(data.Datas) : null;

  if (
    datas &&
    typeof datas.SESSION_ID === "string"
  ) {
    datas.SESSION_ID = maskSessionId(
      datas.SESSION_ID
    );
  }

  return cloned;
}

function extractItemArray(value: unknown): unknown[] {
  const root = asRecord(value);

  if (!root) {
    return [];
  }

  const data = asRecord(root.Data);

  if (!data) {
    return [];
  }

  const directCandidates = [
    data.Result,
    data.RESULT,
    data.Datas,
    data.DATAS,
    data.Data,
    data.DATA,
  ];

  for (const candidate of directCandidates) {
    if (Array.isArray(candidate)) {
      return candidate;
    }

    const candidateRecord = asRecord(candidate);

    if (!candidateRecord) {
      continue;
    }

    for (const key of [
      "Result",
      "RESULT",
      "Datas",
      "DATAS",
      "List",
      "LIST",
      "Items",
      "ITEMS",
    ]) {
      const nested = candidateRecord[key];

      if (Array.isArray(nested)) {
        return nested;
      }
    }
  }

  return [];
}

function normalizeItem(
  value: unknown
): ItemMasterRecord | null {
  const row = asRecord(value);

  if (!row) {
    return null;
  }

  const itemCode = textValue(row, [
    "PROD_CD",
    "PROD_CODE",
    "ITEM_CD",
    "ITEM_CODE",
    "prodCd",
    "prodCode",
    "itemCode",
  ]);

  if (!itemCode) {
    return null;
  }

  const itemName =
    textValue(row, [
      "PROD_DES",
      "PROD_NM",
      "PROD_NAME",
      "ITEM_NM",
      "ITEM_NAME",
      "prodDes",
      "prodName",
      "itemName",
    ]) || itemCode;

  const category =
    textValue(row, [
      "PROD_GROUP_DES",
      "PROD_GROUP_NAME",
      "GROUP_DES",
      "GROUP_NAME",
      "CLASS_DES",
      "CLASS_NAME",
      "CATEGORY",
      "CATEGORY_NAME",
    ]) || "미지정";

  const supplier =
    textValue(row, [
      "CUST_DES",
      "CUST_NAME",
      "SUPPLIER_DES",
      "SUPPLIER_NAME",
      "PUR_CUST_DES",
      "PUR_CUST_NAME",
      "VENDOR_NAME",
    ]) || "미지정";

  const specification = textValue(row, [
    "SIZE_DES",
    "SPEC",
    "SPECIFICATION",
    "PROD_SIZE",
    "SIZE",
  ]);

  const packQuantity = textValue(row, [
    "IN_QTY",
    "PACK_QTY",
    "BOX_QTY",
    "UNIT_QTY",
  ]);

  const stockUnit = textValue(row, [
    "UNIT",
    "UNIT_NM",
    "STOCK_UNIT",
    "INVEN_UNIT",
  ]);

  const storageMethod = textValue(row, [
    "WH_DES",
    "STORAGE_METHOD",
    "STORAGE_METHOD_NAME",
    "KEEP_TYPE_DES",
  ]);

  const useYn = textValue(row, [
    "USE_YN",
    "ACTIVE_YN",
    "IS_ACTIVE",
  ]).toUpperCase();

  return {
    itemCode,
    itemName,
    category,
    supplier,
    specification,
    packQuantity,
    stockUnit,
    storageMethod,
    source: "initial",
    active: useYn !== "N",
    memo: textValue(row, [
      "REMARKS",
      "REMARK",
      "MEMO",
      "CONTENTS",
    ]),
    ecountUpdatedAt: new Date().toISOString(),
  };
}

async function getD1Setting(
  requestUrl: string,
  key: string
): Promise<D1SettingResponse | null> {
  const url = new URL(
    `/api/d1/settings/${encodeURIComponent(key)}`,
    requestUrl
  );

  const response = await fetch(url, {
    method: "GET",
    headers: {
      "Content-Type": "application/json",
    },
    cache: "no-store",
  });

  if (response.status === 404) {
    return null;
  }

  const responseText = await response.text();

  if (!response.ok) {
    throw new Error(
      `기존 D1 API 조회 실패 (${response.status}): ${responseText.slice(0, 300)}`
    );
  }

  if (!responseText) {
    return null;
  }

  try {
    return JSON.parse(responseText) as D1SettingResponse;
  } catch {
    throw new Error(
      "기존 D1 API 조회 결과를 JSON으로 해석할 수 없습니다."
    );
  }
}

async function putD1Setting(
  requestUrl: string,
  key: string,
  data: unknown
): Promise<D1SettingResponse> {
  const url = new URL(
    `/api/d1/settings/${encodeURIComponent(key)}`,
    requestUrl
  );

  const response = await fetch(url, {
    method: "PUT",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      data,
    }),
    cache: "no-store",
  });

  const responseText = await response.text();

  if (!response.ok) {
    throw new Error(
      `기존 D1 API 저장 실패 (${response.status}): ${responseText.slice(0, 300)}`
    );
  }

  if (!responseText) {
    return {};
  }

  try {
    return JSON.parse(responseText) as D1SettingResponse;
  } catch {
    return {};
  }
}

function mergeItemMasters(
  existingItems: ItemMasterRecord[],
  incomingItems: ItemMasterRecord[]
) {
  const itemMap = new Map<string, ItemMasterRecord>();

  for (const item of existingItems) {
    if (!item?.itemCode) {
      continue;
    }

    itemMap.set(item.itemCode, item);
  }

  let addedCount = 0;
  let updatedCount = 0;

  for (const incoming of incomingItems) {
    const previous = itemMap.get(incoming.itemCode);

    if (!previous) {
      itemMap.set(incoming.itemCode, incoming);
      addedCount += 1;
      continue;
    }

    const mergedItem: ItemMasterRecord = {
      ...previous,
      ...incoming,

      itemCode: incoming.itemCode,

      itemName:
        incoming.itemName ||
        previous.itemName ||
        incoming.itemCode,

      category:
        incoming.category &&
        incoming.category !== "미지정"
          ? incoming.category
          : previous.category || "미지정",

      supplier:
        incoming.supplier &&
        incoming.supplier !== "미지정"
          ? incoming.supplier
          : previous.supplier || "미지정",

      specification:
        incoming.specification ||
        previous.specification ||
        "",

      packQuantity:
        incoming.packQuantity ||
        previous.packQuantity ||
        "",

      stockUnit:
        incoming.stockUnit ||
        previous.stockUnit ||
        "",

      storageMethod:
        incoming.storageMethod ||
        previous.storageMethod ||
        "",

      memo:
        incoming.memo ||
        previous.memo ||
        "",

      source:
        previous.source ||
        incoming.source ||
        "initial",

      firstSeenMonth:
        previous.firstSeenMonth,

      active: incoming.active,

      ecountUpdatedAt:
        incoming.ecountUpdatedAt,
    };

    itemMap.set(incoming.itemCode, mergedItem);
    updatedCount += 1;
  }

  const merged = Array.from(itemMap.values()).sort(
    (a, b) =>
      a.itemCode.localeCompare(
        b.itemCode,
        "ko",
        {
          numeric: true,
        }
      )
  );

  return {
    merged,
    addedCount,
    updatedCount,
  };
}

export async function POST(
  request: NextRequest
) {
  try {
    const body =
      (await request.json()) as EcountRequest;

    const comCode = body.comCode?.trim();
    const userId = body.userId?.trim();
    const apiCertKey =
      body.apiCertKey?.trim();
    const prodCode =
      body.prodCode?.trim() || "";
    const syncToDailySales =
      body.syncToDailySales === true;

    if (!comCode || !userId || !apiCertKey) {
      return NextResponse.json(
        {
          ok: false,
          step: "input",
          message:
            "회사코드, 사용자 ID, API 인증키를 모두 입력해 주세요.",
        },
        {
          status: 400,
        }
      );
    }

    /*
     * 1. 회사 Zone 확인
     */
    const zoneResponse = await fetch(
      TEST_ZONE_URL,
      {
        method: "POST",
        headers: {
          "Content-Type":
            "application/json",
        },
        body: JSON.stringify({
          COM_CODE: comCode,
        }),
        cache: "no-store",
      }
    );

    const zoneJson =
      await readJsonResponse(
        zoneResponse
      );

    const zoneData =
      getNestedRecord(
        zoneJson,
        "Data"
      );

    const rawZone =
      typeof zoneData?.ZONE ===
      "string"
        ? zoneData.ZONE.trim()
        : "";

    const rawDomain =
      typeof zoneData?.DOMAIN ===
      "string"
        ? zoneData.DOMAIN.trim()
        : "ecount.com";

    const zone = rawZone.replace(
      /^\.+|\.+$/g,
      ""
    );

    const domain =
      rawDomain.replace(
        /^\.+|\.+$/g,
        ""
      ) || "ecount.com";

    if (!zoneResponse.ok || !zone) {
      return NextResponse.json(
        {
          ok: false,
          step: "zone",
          message:
            getEcountMessage(
              zoneJson
            ) ||
            "이카운트 Zone 확인에 실패했습니다.",
          zoneHttpStatus:
            zoneResponse.status,
          zoneResponse: zoneJson,
        },
        {
          status: 502,
        }
      );
    }

    /*
     * 2. ECOUNT 로그인
     */
    const loginUrl =
      `https://sboapi${zone}.${domain}` +
      "/OAPI/V2/OAPILogin";

    const loginResponse =
      await fetch(loginUrl, {
        method: "POST",
        headers: {
          "Content-Type":
            "application/json",
        },
        body: JSON.stringify({
          COM_CODE: comCode,
          USER_ID: userId,
          API_CERT_KEY:
            apiCertKey,
          LAN_TYPE: "ko-KR",
          ZONE: zone,
        }),
        cache: "no-store",
      });

    const loginJson =
      await readJsonResponse(
        loginResponse
      );

    const loginDatas =
      getNestedRecord(
        loginJson,
        "Data",
        "Datas"
      );

    const sessionId =
      typeof loginDatas?.SESSION_ID ===
      "string"
        ? loginDatas.SESSION_ID
        : "";

    if (
      !loginResponse.ok ||
      !sessionId
    ) {
      return NextResponse.json(
        {
          ok: false,
          step: "login",
          message:
            getEcountMessage(
              loginJson
            ) ||
            "이카운트 로그인 또는 SESSION_ID 발급에 실패했습니다.",
          zone: {
            zone,
            domain,
          },
          loginHttpStatus:
            loginResponse.status,
          loginResponse:
            safeLoginResult(
              loginJson
            ),
        },
        {
          status: 502,
        }
      );
    }

    /*
     * 3. 품목 조회
     *
     * 중요:
     * GetBasicProductList가 아니라
     * GetBasicProductsList입니다.
     */
    const itemUrl =
      `https://sboapi${zone}.${domain}` +
      "/OAPI/V2/InventoryBasic/GetBasicProductsList" +
      `?SESSION_ID=${encodeURIComponent(sessionId)}`;

    const itemResponse =
      await fetch(itemUrl, {
        method: "POST",
        headers: {
          "Content-Type":
            "application/json",
        },
       body: JSON.stringify({
  PROD_CD: prodCode,
  PROD_DES: "",
  SIZE_DES: "",
  UNIT: "",
  PROD_TYPE: "",
  CLASS_CD: "",
  PAGE_NO: 1,
  PAGE_SIZE: 5000,
}),
        cache: "no-store",
      });

    const itemJson =
      await readJsonResponse(
        itemResponse
      );

    if (!itemResponse.ok) {
      return NextResponse.json(
        {
          ok: false,
          step: "items",
          message:
            getEcountMessage(
              itemJson
            ) ||
            "SESSION_ID는 발급됐지만 품목조회에 실패했습니다.",
          zone: {
            zone,
            domain,
          },
          login: {
            sessionIssued: true,
            maskedSessionId:
              maskSessionId(
                sessionId
              ),
          },
          itemUrl,
          itemHttpStatus:
            itemResponse.status,
          itemResponse: itemJson,
        },
        {
          status: 502,
        }
      );
    }

    const rawItems =
      extractItemArray(itemJson);

    const normalizedItems =
      rawItems
        .map(normalizeItem)
        .filter(
          (
            item
          ): item is ItemMasterRecord =>
            item !== null
        );

    const itemData =
      getNestedRecord(
        itemJson,
        "Data"
      );

    const totalCountCandidate =
      itemData?.TotalCnt ??
      itemData?.TOTAL_CNT ??
      itemData?.TotalCount ??
      normalizedItems.length;

    const totalCount =
      Number.isFinite(
        Number(totalCountCandidate)
      )
        ? Number(totalCountCandidate)
        : normalizedItems.length;

    let sync:
      | {
          saved: boolean;
          key: string;
          statusKey: string;
          addedCount: number;
          updatedCount: number;
          totalSavedCount: number;
          updatedAt: string;
        }
      | undefined;

    /*
     * 4. D1 저장
     */
    if (syncToDailySales) {
      if (
        normalizedItems.length === 0
      ) {
        return NextResponse.json(
          {
            ok: false,
            step: "sync",
            message:
              "품목조회는 성공했지만 저장할 품목이 없습니다. 원본 응답을 확인해 주세요.",
            zone: {
              zone,
              domain,
            },
            login: {
              sessionIssued: true,
              maskedSessionId:
                maskSessionId(
                  sessionId
                ),
            },
            items: {
              totalCount,
              returnedCount: 0,
              preview: [],
            },
            raw: {
              items: itemJson,
            },
          },
          {
            status: 502,
          }
        );
      }

      const existingSetting =
        await getD1Setting(
          request.url,
          ITEM_MASTER_KEY
        );

      const existingItems =
        Array.isArray(
          existingSetting?.data
        )
          ? (existingSetting
              ?.data as ItemMasterRecord[])
          : [];

      const mergeResult =
        mergeItemMasters(
          existingItems,
          normalizedItems
        );

      const savedResult =
        await putD1Setting(
          request.url,
          ITEM_MASTER_KEY,
          mergeResult.merged
        );

      const updatedAt =
        savedResult.updated_at ||
        new Date().toISOString();

      const syncStatus = {
        ok: true,
        type: "items",
        environment: "test",
        lastSyncedAt: updatedAt,
        itemCount:
          mergeResult.merged.length,
        receivedCount:
          normalizedItems.length,
        addedCount:
          mergeResult.addedCount,
        updatedCount:
          mergeResult.updatedCount,
        zone,
        domain,
        executedFrom:
          request.headers.get(
            "host"
          ) || "unknown",
      };

      await putD1Setting(
        request.url,
        SYNC_STATUS_KEY,
        syncStatus
      );

      sync = {
        saved: true,
        key: ITEM_MASTER_KEY,
        statusKey:
          SYNC_STATUS_KEY,
        addedCount:
          mergeResult.addedCount,
        updatedCount:
          mergeResult.updatedCount,
        totalSavedCount:
          mergeResult.merged.length,
        updatedAt,
      };
    }

    return NextResponse.json({
      ok: true,

      message: syncToDailySales
        ? "이카운트 품목을 Daily Sales 품목기준정보에 반영했습니다."
        : "이카운트 로그인과 품목조회에 성공했습니다.",

      environment: "test",

      zone: {
        zone,
        domain,
      },

      login: {
        sessionIssued: true,
        maskedSessionId:
          maskSessionId(
            sessionId
          ),
      },

      items: {
        requestedProdCode:
          prodCode,
        totalCount,
        returnedCount:
          normalizedItems.length,
        preview:
          normalizedItems.slice(
            0,
            50
          ),
      },

      sync,

      raw: {
        zone: zoneJson,
        login:
          safeLoginResult(
            loginJson
          ),
        items: itemJson,
      },
    });
  } catch (error) {
    console.error(
      "ECOUNT sync failed:",
      error
    );

    return NextResponse.json(
      {
        ok: false,
        step: "server",
        message:
          error instanceof Error
            ? error.message
            : "알 수 없는 서버 오류가 발생했습니다.",
      },
      {
        status: 500,
      }
    );
  }
}