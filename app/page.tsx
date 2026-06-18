"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

type SalesRow = {
  date: string;
  product: string;
  amount: number;
  quantity: number;
};

export default function Home() {
  const [rows, setRows] = useState<SalesRow[]>([]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json<any>(sheet);

    const parsed: SalesRow[] = json.map((row) => ({
      date: String(row["날짜"] ?? row["date"] ?? ""),
      product: String(row["상품명"] ?? row["product"] ?? ""),
      amount: Number(row["매출"] ?? row["amount"] ?? 0),
      quantity: Number(row["수량"] ?? row["quantity"] ?? 0),
    }));

    setRows(parsed);
  };

  const totalSales = useMemo(
    () => rows.reduce((sum, row) => sum + row.amount, 0),
    [rows]
  );

  const totalQuantity = useMemo(
    () => rows.reduce((sum, row) => sum + row.quantity, 0),
    [rows]
  );

  const averageSales = rows.length > 0 ? Math.round(totalSales / rows.length) : 0;

  return (
    <main className="min-h-screen bg-gray-100 p-8">
      <div className="max-w-6xl mx-auto">
        <div className="mb-8">
          <h1 className="text-3xl font-bold">ABLAB 매출 대시보드</h1>
          <p className="text-gray-600 mt-2">
            엑셀 매출 보고 파일을 업로드하면 자동으로 요약됩니다.
          </p>
        </div>

        <div className="bg-white rounded-xl shadow p-6 mb-6">
          <h2 className="text-xl font-bold mb-4">매출 파일 업로드</h2>
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handleFileUpload}
            className="block w-full border rounded-lg p-3"
          />
          <p className="text-sm text-gray-500 mt-3">
            엑셀 컬럼명 예시: 날짜, 상품명, 매출, 수량
          </p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
          <div className="bg-white rounded-xl shadow p-6">
            <p className="text-gray-500">총 매출</p>
            <h2 className="text-2xl font-bold mt-2">
              ₩{totalSales.toLocaleString()}
            </h2>
          </div>

          <div className="bg-white rounded-xl shadow p-6">
            <p className="text-gray-500">총 수량</p>
            <h2 className="text-2xl font-bold mt-2">
              {totalQuantity.toLocaleString()}개
            </h2>
          </div>

          <div className="bg-white rounded-xl shadow p-6">
            <p className="text-gray-500">평균 매출</p>
            <h2 className="text-2xl font-bold mt-2">
              ₩{averageSales.toLocaleString()}
            </h2>
          </div>
        </div>

        <div className="bg-white rounded-xl shadow p-6">
          <h2 className="text-xl font-bold mb-4">매출 내역</h2>

          {rows.length === 0 ? (
            <p className="text-gray-500">아직 업로드된 파일이 없습니다.</p>
          ) : (
            <table className="w-full border-collapse">
              <thead>
                <tr className="border-b bg-gray-50">
                  <th className="p-3 text-left">날짜</th>
                  <th className="p-3 text-left">상품명</th>
                  <th className="p-3 text-right">매출</th>
                  <th className="p-3 text-right">수량</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((row, index) => (
                  <tr key={index} className="border-b">
                    <td className="p-3">{row.date}</td>
                    <td className="p-3">{row.product}</td>
                    <td className="p-3 text-right">
                      ₩{row.amount.toLocaleString()}
                    </td>
                    <td className="p-3 text-right">
                      {row.quantity.toLocaleString()}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>
    </main>
  );
}
