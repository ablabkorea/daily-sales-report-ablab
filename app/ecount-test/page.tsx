"use client";

import { FormEvent, useState } from "react";

type ItemPreview = {
  itemCode?: string;
  itemName?: string;
  category?: string;
  supplier?: string;
  specification?: string;
  stockUnit?: string;
  active?: boolean;
};

type TestResult = {
  ok?: boolean;
  message?: string;
  step?: string;
  zone?: { zone?: string; domain?: string };
  login?: { sessionIssued?: boolean; maskedSessionId?: string };
  items?: { requestedProdCode?: string; totalCount?: number; returnedCount?: number; preview?: ItemPreview[] };
  sync?: {
    saved?: boolean;
    addedCount?: number;
    updatedCount?: number;
    totalSavedCount?: number;
    updatedAt?: string;
  };
  [key: string]: unknown;
};

export default function EcountTestPage() {
  const [comCode, setComCode] = useState("");
  const [userId, setUserId] = useState("");
  const [apiCertKey, setApiCertKey] = useState("");
  const [prodCode, setProdCode] = useState("");
  const [loadingMode, setLoadingMode] = useState<"test" | "sync" | null>(null);
  const [result, setResult] = useState<TestResult | null>(null);

  async function runRequest(syncToDailySales: boolean) {
    setLoadingMode(syncToDailySales ? "sync" : "test");
    setResult(null);

    try {
      const response = await fetch("/api/ecount/test", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ comCode, userId, apiCertKey, prodCode, syncToDailySales }),
      });
      const data = (await response.json()) as TestResult;
      setResult(data);
    } catch (error) {
      setResult({
        ok: false,
        step: "browser",
        message: error instanceof Error ? error.message : "요청 중 오류가 발생했습니다.",
      });
    } finally {
      setLoadingMode(null);
    }
  }

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    await runRequest(false);
  }

  const success = result?.ok === true;
  const preview = result?.items?.preview ?? [];

  return (
    <main style={styles.page}>
      <section style={styles.card}>
        <div style={styles.header}>
          <div>
            <p style={styles.eyebrow}>ABLAB Daily Sales</p>
            <h1 style={styles.title}>ECOUNT 품목 연동</h1>
            <p style={styles.description}>
              이카운트 품목을 조회한 뒤 Daily Sales의 품목기준정보에 저장합니다.
              품목코드·품목명·카테고리·매입처·규격·단위를 자동 반영하고 기존 EST, Target, 담당자 데이터는 건드리지 않습니다.
            </p>
          </div>
          <a href="/" style={styles.homeLink}>Daily Sales로 돌아가기</a>
        </div>

        <form onSubmit={handleSubmit} style={styles.form} autoComplete="off">
          <label style={styles.label}>
            회사코드
            <input value={comCode} onChange={(event) => setComCode(event.target.value)} placeholder="이카운트 회사코드" style={styles.input} required />
          </label>

          <label style={styles.label}>
            사용자 ID
            <input value={userId} onChange={(event) => setUserId(event.target.value)} placeholder="이카운트 로그인 ID" style={styles.input} required />
          </label>

          <label style={styles.label}>
            테스트 API 인증키
            <input type="password" value={apiCertKey} onChange={(event) => setApiCertKey(event.target.value)} placeholder="테스트 인증키 입력" style={styles.input} required />
          </label>

          <label style={styles.label}>
            품목코드 (선택)
            <input value={prodCode} onChange={(event) => setProdCode(event.target.value)} placeholder="비워두면 전체 품목 조회" style={styles.input} />
          </label>

          <p style={styles.notice}>
            인증키는 데이터베이스나 브라우저 저장소에 저장하지 않습니다. 동기화 시 기존 품목기준정보와 병합하며, 이카운트에서 값이 없는 카테고리·매입처는 기존 수기값을 유지합니다.
          </p>

          <div style={styles.buttonRow}>
            <button type="submit" disabled={loadingMode !== null} style={{ ...styles.secondaryButton, opacity: loadingMode ? 0.65 : 1 }}>
              {loadingMode === "test" ? "조회 중..." : "연결 및 품목조회 테스트"}
            </button>
            <button
              type="button"
              disabled={loadingMode !== null || !comCode || !userId || !apiCertKey}
              onClick={() => void runRequest(true)}
              style={{ ...styles.button, opacity: loadingMode || !comCode || !userId || !apiCertKey ? 0.65 : 1 }}
            >
              {loadingMode === "sync" ? "Daily Sales 반영 중..." : "품목을 Daily Sales에 반영"}
            </button>
          </div>
        </form>

        <section style={styles.statusSection}>
          <h2 style={styles.sectionTitle}>연동 결과</h2>
          {!result && <div style={styles.empty}>아직 조회하거나 반영하지 않았습니다.</div>}

          {result && (
            <>
              <div style={{ ...styles.resultBanner, ...(success ? styles.successBanner : styles.errorBanner) }}>
                <strong>{success ? "처리 성공" : "처리 실패"}</strong>
                <span>{result.message ?? "응답 메시지가 없습니다."}</span>
              </div>

              <div style={styles.stepGrid}>
                <div style={styles.stepCard}>
                  <span style={styles.stepNumber}>1</span>
                  <div><strong>Zone 확인</strong><p style={styles.stepText}>{result.zone?.zone ? `${result.zone.zone} / ${result.zone.domain ?? "ecount.com"}` : result.step === "zone" ? "실패" : "확인되지 않음"}</p></div>
                </div>
                <div style={styles.stepCard}>
                  <span style={styles.stepNumber}>2</span>
                  <div><strong>SESSION_ID</strong><p style={styles.stepText}>{result.login?.sessionIssued ? `발급 성공 (${result.login.maskedSessionId ?? "마스킹됨"})` : result.step === "login" ? "실패" : "대기 또는 미확인"}</p></div>
                </div>
                <div style={styles.stepCard}>
                  <span style={styles.stepNumber}>3</span>
                  <div><strong>품목조회</strong><p style={styles.stepText}>{result.items ? `${result.items.returnedCount ?? 0}건 반환 / 전체 ${result.items.totalCount ?? 0}건` : result.step === "items" ? "실패" : "대기 또는 미확인"}</p></div>
                </div>
                <div style={styles.stepCard}>
                  <span style={styles.stepNumber}>4</span>
                  <div><strong>Daily Sales 반영</strong><p style={styles.stepText}>{result.sync?.saved ? `신규 ${result.sync.addedCount ?? 0}건 · 갱신 ${result.sync.updatedCount ?? 0}건 · 총 ${result.sync.totalSavedCount ?? 0}건` : result.step === "sync" ? "실패" : "조회만 실행됨"}</p></div>
                </div>
              </div>

              {preview.length > 0 && (
                <div style={styles.tableWrap}>
                  <div style={styles.tableTitle}>품목 미리보기 ({preview.length}건)</div>
                  <table style={styles.table}>
                    <thead>
                      <tr>
                        <th style={styles.th}>품목코드</th>
                        <th style={styles.th}>품목명</th>
                        <th style={styles.th}>카테고리</th>
                        <th style={styles.th}>매입처</th>
                        <th style={styles.th}>규격</th>
                        <th style={styles.th}>단위</th>
                      </tr>
                    </thead>
                    <tbody>
                      {preview.map((item, index) => (
                        <tr key={`${item.itemCode ?? "item"}-${index}`}>
                          <td style={styles.td}>{item.itemCode || "-"}</td>
                          <td style={styles.td}>{item.itemName || "-"}</td>
                          <td style={styles.td}>{item.category || "미지정"}</td>
                          <td style={styles.td}>{item.supplier || "미지정"}</td>
                          <td style={styles.td}>{item.specification || "-"}</td>
                          <td style={styles.td}>{item.stockUnit || "-"}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              <details style={styles.details}>
                <summary style={styles.summary}>원본 응답 보기</summary>
                <pre style={styles.pre}>{JSON.stringify(result, null, 2)}</pre>
              </details>
            </>
          )}
        </section>
      </section>
    </main>
  );
}

const styles: Record<string, React.CSSProperties> = {
  page: { minHeight: "100vh", background: "#f4f7fb", padding: "48px 20px", color: "#172033", fontFamily: "Arial, 'Noto Sans KR', sans-serif" },
  card: { width: "min(1100px, 100%)", margin: "0 auto", background: "#ffffff", border: "1px solid #dbe4f0", borderRadius: 18, boxShadow: "0 18px 50px rgba(31, 49, 77, 0.08)", padding: 32 },
  header: { display: "flex", justifyContent: "space-between", gap: 24, alignItems: "flex-start", flexWrap: "wrap" },
  eyebrow: { margin: 0, color: "#315efb", fontWeight: 700, fontSize: 13 },
  title: { margin: "8px 0 10px", fontSize: 30 },
  description: { margin: 0, maxWidth: 760, lineHeight: 1.65, color: "#5d687a" },
  homeLink: { color: "#315efb", textDecoration: "none", fontWeight: 700, fontSize: 14 },
  form: { marginTop: 30, display: "grid", gap: 16, padding: 24, background: "#f8faff", borderRadius: 14, border: "1px solid #e3eaf4" },
  label: { display: "grid", gap: 8, fontWeight: 700, fontSize: 14 },
  input: { height: 46, border: "1px solid #cfd9e7", borderRadius: 10, padding: "0 14px", fontSize: 15, outline: "none", background: "white" },
  notice: { margin: 0, color: "#677386", fontSize: 13, lineHeight: 1.5 },
  buttonRow: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 12 },
  button: { height: 48, border: 0, borderRadius: 10, background: "#315efb", color: "white", fontWeight: 800, fontSize: 15, cursor: "pointer" },
  secondaryButton: { height: 48, border: "1px solid #9fb2d0", borderRadius: 10, background: "#ffffff", color: "#31506f", fontWeight: 800, fontSize: 15, cursor: "pointer" },
  statusSection: { marginTop: 30 },
  sectionTitle: { margin: "0 0 14px", fontSize: 20 },
  empty: { padding: 24, border: "1px dashed #cfd9e7", borderRadius: 12, color: "#7a8699", textAlign: "center" },
  resultBanner: { display: "flex", flexDirection: "column", gap: 6, padding: 18, borderRadius: 12, border: "1px solid" },
  successBanner: { background: "#effaf3", borderColor: "#a6ddb9", color: "#176b38" },
  errorBanner: { background: "#fff3f3", borderColor: "#efb5b5", color: "#a52626" },
  stepGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: 14, marginTop: 16 },
  stepCard: { display: "flex", gap: 12, alignItems: "center", border: "1px solid #e1e8f1", borderRadius: 12, padding: 16 },
  stepNumber: { width: 30, height: 30, borderRadius: 999, display: "grid", placeItems: "center", background: "#e7edff", color: "#315efb", fontWeight: 800, flex: "0 0 auto" },
  stepText: { margin: "5px 0 0", color: "#667287", fontSize: 13 },
  tableWrap: { marginTop: 18, overflowX: "auto", border: "1px solid #e1e8f1", borderRadius: 12 },
  tableTitle: { padding: 14, fontWeight: 800, background: "#f8faff", borderBottom: "1px solid #e1e8f1" },
  table: { width: "100%", borderCollapse: "collapse", minWidth: 850, fontSize: 13 },
  th: { padding: "10px 12px", background: "#eef3fb", borderBottom: "1px solid #dce5f0", textAlign: "left", whiteSpace: "nowrap" },
  td: { padding: "10px 12px", borderBottom: "1px solid #edf1f6", whiteSpace: "nowrap" },
  details: { marginTop: 18, border: "1px solid #e1e8f1", borderRadius: 12, overflow: "hidden" },
  summary: { padding: 14, cursor: "pointer", fontWeight: 700, background: "#f8faff" },
  pre: { margin: 0, padding: 16, background: "#111827", color: "#e5e7eb", overflowX: "auto", fontSize: 12, lineHeight: 1.55 },
};
