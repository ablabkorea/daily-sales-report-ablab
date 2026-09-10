import { env } from "cloudflare:workers";
import { NextRequest, NextResponse } from "next/server";

export const dynamic = "force-dynamic";
export const runtime = "nodejs";

type ServiceFetcher = {
  fetch(input: RequestInfo | URL, init?: RequestInit): Promise<Response>;
};

function workerConfig() {
  const apiKey = process.env.D1_API_KEY?.trim();
  const salesApi = (env as unknown as { SALES_API?: ServiceFetcher }).SALES_API;

  if (!apiKey || !salesApi) return null;
  return { apiKey, salesApi };
}

async function proxy(
  request: NextRequest,
  context: { params: Promise<{ path: string[] }> },
) {
  const config = workerConfig();

  if (!config) {
    return NextResponse.json(
      { error: "D1 Worker service binding or API key is missing" },
      { status: 503 },
    );
  }

  const { path } = await context.params;

  const target = new URL(
    `https://ablab-sales-api.internal/${path.map(encodeURIComponent).join("/")}`,
  );

  request.nextUrl.searchParams.forEach((value, key) => {
    target.searchParams.append(key, value);
  });

  const body =
    request.method === "GET" || request.method === "HEAD"
      ? undefined
      : await request.text();

  const response = await config.salesApi.fetch(target, {
    method: request.method,
    headers: {
      "X-ABL-API-Key": config.apiKey,
      "Content-Type":
        request.headers.get("content-type") || "application/json",
    },
    body,
  });

  const headers = new Headers();
  headers.set(
    "Content-Type",
    response.headers.get("content-type") ||
      "application/json; charset=utf-8",
  );
  headers.set("Cache-Control", "no-store");

  const cacheSource = response.headers.get("X-ABL-Cache");
  if (cacheSource) {
    headers.set("X-ABL-Cache", cacheSource);
  }

  // 내부 ablab-sales-api Worker의 실제 Cloudflare Workers Cache 상태를
  // 브라우저에서 확인할 수 있도록 별도 헤더로 전달합니다.
  const edgeStatus = response.headers.get("CF-Cache-Status");
  if (edgeStatus) {
    headers.set("X-ABL-Edge-Status", edgeStatus);
  }

  return new NextResponse(response.body, {
    status: response.status,
    headers,
  });
}

export const GET = proxy;
export const POST = proxy;
export const PUT = proxy;
export const DELETE = proxy;
