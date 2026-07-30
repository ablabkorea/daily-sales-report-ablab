import { NextResponse } from "next/server";

const GOOGLE_SCRIPT_URL =
  "https://script.google.com/macros/s/AKfycbz6JIOT5Cq_K6oHFEPa5fE6SLYO8nDE7V2SfD94GyjTOFQv1wu_s8rq07Deb0mmpUvs/exec";

export async function GET() {
  try {
    const response = await fetch(GOOGLE_SCRIPT_URL, {
      method: "GET",
      cache: "no-store",
      redirect: "follow",
    });

    const text = await response.text();

    return new NextResponse(text, {
      status: 200,
      headers: {
        "Content-Type": "application/json; charset=utf-8",
      },
    });
  } catch (error) {
    console.error("월초 데이터 불러오기 실패:", error);

    return NextResponse.json(
      { error: "월초 데이터 불러오기 실패" },
      { status: 500 }
    );
  }
}