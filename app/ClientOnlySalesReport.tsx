"use client";

import dynamic from "next/dynamic";

const SalesReportClient = dynamic(() => import("./SalesReportClient"), {
  ssr: false,
  loading: () => (
    <main className="flex min-h-screen items-center justify-center bg-orange-50 text-slate-900">
      <div className="rounded-2xl border border-orange-200 bg-white/90 px-6 py-5 text-sm font-semibold shadow-sm">
        Sales Report를 불러오는 중입니다...
      </div>
    </main>
  ),
});

export default function ClientOnlySalesReport() {
  return <SalesReportClient />;
}
