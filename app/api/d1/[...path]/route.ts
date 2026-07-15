import { NextRequest, NextResponse } from "next/server";

export const dynamic = "force-dynamic";
export const runtime = "nodejs";

function workerConfig() {
  const url = process.env.D1_WORKER_URL?.trim().replace(/\/$/, "");
  const apiKey = process.env.D1_API_KEY?.trim();
  if (!url || !apiKey) return null;
  return { url, apiKey };
}

async function proxy(request: NextRequest, context: { params: Promise<{ path: string[] }> }) {
  const config = workerConfig();
  if (!config) {
    return NextResponse.json(
      { error: "D1 Worker environment variables are missing" },
      { status: 503 },
    );
  }

  const { path } = await context.params;
  const target = new URL(`${config.url}/${path.map(encodeURIComponent).join("/")}`);
  request.nextUrl.searchParams.forEach((value, key) => target.searchParams.append(key, value));

  const body = request.method === "GET" || request.method === "HEAD"
    ? undefined
    : await request.text();

  const response = await fetch(target, {
    method: request.method,
    headers: {
      "X-ABL-API-Key": config.apiKey,
      "Content-Type": request.headers.get("content-type") || "application/json",
    },
    body,
    cache: "no-store",
  });

  return new NextResponse(response.body, {
    status: response.status,
    headers: {
      "Content-Type": response.headers.get("content-type") || "application/json; charset=utf-8",
      "Cache-Control": "no-store",
    },
  });
}

export const GET = proxy;
export const POST = proxy;
export const PUT = proxy;
export const DELETE = proxy;
