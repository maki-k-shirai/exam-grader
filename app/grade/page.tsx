"use client"

import { useMemo, useState } from "react"

type Result = {
  total_score: number
  max_score: number
  details: Array<{
    id: number
    report: string
    cell: string
    check_label: string
    expected: any
    actual: any
    ok: boolean
    score: number
    points: number
    reason?: string
  }>
  csv: { summary: string; ng: string }
}

function downloadText(filename: string, text: string) {
  // Excel対策：UTF-8 BOM を先頭につける
  const bom = "\uFEFF"
  const blob = new Blob([bom + text], { type: "text/csv;charset=utf-8" })
  const url = URL.createObjectURL(blob)
  const a = document.createElement("a")
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

export default function GradePage() {
  const [files, setFiles] = useState<FileList | null>(null)
  const [loading, setLoading] = useState(false)
  const [result, setResult] = useState<Result | null>(null)
  const [error, setError] = useState<string | null>(null)

  const ngOnly = useMemo(() => {
    if (!result) return []
    return result.details.filter((d) => !d.ok)
  }, [result])

  async function onGrade() {
    setError(null)
    setResult(null)
    if (!files || files.length === 0) {
      setError("Excelファイル（提出5つ）を選択してください。")
      return
    }

    setLoading(true)
    try {
      const fd = new FormData()
      Array.from(files).forEach((f) => fd.append("files", f))
      const res = await fetch("/api/grade", { method: "POST", body: fd })
      if (!res.ok) throw new Error(`採点に失敗しました（${res.status}）`)
      const json = (await res.json()) as Result
      setResult(json)
    } catch (e: any) {
      setError(e?.message ?? "エラーが発生しました。")
    } finally {
      setLoading(false)
    }
  }

  return (
    <div style={{ maxWidth: 1100, margin: "0 auto", padding: 20 }}>
      <h1 style={{ fontSize: 22, fontWeight: 700 }}>社内テスト 自動採点</h1>
      <p style={{ marginTop: 8, color: "#444" }}>
        提出された5つのExcelを選択 → 採点 → NGのみ一覧＆CSV出力
      </p>

      <div style={{ marginTop: 16, padding: 16, border: "1px solid #ddd", borderRadius: 10 }}>
        <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap" }}>
          <input type="file" multiple accept=".xlsx" onChange={(e) => setFiles(e.target.files)} />
          <button
            onClick={onGrade}
            disabled={loading}
            style={{
              padding: "10px 14px",
              borderRadius: 8,
              border: "1px solid #333",
              background: loading ? "#eee" : "#fff",
              cursor: loading ? "not-allowed" : "pointer",
              fontWeight: 700,
            }}
          >
            {loading ? "採点中..." : "採点する"}
          </button>

{result && (
  <div
    style={{
      marginTop: 16,
      padding: 20,
      borderRadius: 14,
      background: "#f7f9ff",
      border: "2px solid #4f6ef7",
      display: "flex",
      alignItems: "center",
      gap: 24,
    }}
  >
    <div>
      <div style={{ fontSize: 14, color: "#4f6ef7", fontWeight: 700 }}>
        得点
      </div>
      <div style={{ fontSize: 40, fontWeight: 900, lineHeight: 1.1 }}>
        {result.total_score}
        <span style={{ fontSize: 18, fontWeight: 600 }}>
          {" "} / {result.max_score}
        </span>
      </div>
    </div>

    <div style={{ fontSize: 14, color: "#555" }}>
      NG件数：{ngOnly.length} 件
    </div>
  </div>
)}
        </div>

        {error && <div style={{ marginTop: 12, color: "crimson" }}>{error}</div>}

        {result && (
          <div style={{ marginTop: 12, display: "flex", gap: 10, flexWrap: "wrap" }}>
<button
  onClick={() => {
    if (!result?.csv?.summary) return alert("CSVがまだ生成されていません（APIの返却を確認してください）")
    downloadText("採点結果_集計.csv", result.csv.summary)
  }}
  style={{ padding: "8px 12px", borderRadius: 8, border: "1px solid #777", background: "#fff" }}
>
  集計CSVをDL
</button>

<button
  onClick={() => {
    if (!result?.csv?.ng) return alert("CSVがまだ生成されていません（APIの返却を確認してください）")
    downloadText("採点結果_NG一覧.csv", result.csv.ng)
  }}
  style={{ padding: "8px 12px", borderRadius: 8, border: "1px solid #777", background: "#fff" }}
>
  NG一覧CSVをDL
</button>

          </div>
        )}
      </div>

      {result && (
        <div style={{ marginTop: 18 }}>
          <h2 style={{ fontSize: 18, fontWeight: 800 }}>NGのみ一覧（{ngOnly.length}件）</h2>

          <div style={{ marginTop: 10, overflowX: "auto", border: "1px solid #ddd", borderRadius: 10 }}>
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 900 }}>
              <thead>
                <tr style={{ background: "#f6f6f6" }}>
                  {["帳票", "セル", "CK箇所", "期待値", "実値", "配点"].map((h) => (
                    <th key={h} style={{ textAlign: "left", padding: 10, borderBottom: "1px solid #ddd" }}>
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {ngOnly.map((d) => (
                  <tr key={d.id}>
                    <td style={{ padding: 10, borderBottom: "1px solid #eee" }}>{d.report}</td>
                    <td style={{ padding: 10, borderBottom: "1px solid #eee" }}>{d.cell}</td>
                    <td style={{ padding: 10, borderBottom: "1px solid #eee" }}>{d.check_label}</td>
                    <td style={{ padding: 10, borderBottom: "1px solid #eee" }}>{String(d.expected ?? "")}</td>
                    <td style={{ padding: 10, borderBottom: "1px solid #eee" }}>{String(d.actual ?? "")}</td>
                    <td style={{ padding: 10, borderBottom: "1px solid #eee" }}>{d.points}</td>
                  </tr>
                ))}
                {ngOnly.length === 0 && (
                  <tr>
                    <td colSpan={6} style={{ padding: 14, color: "#666" }}>
                      NGはありません（満点の可能性）
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  )
}
