import { NextResponse } from "next/server"
import ExcelJS from "exceljs"
import path from "path"
import fs from "fs/promises"

type Rule = {
  compare_type?: "auto" | "number" | "string"
  tolerance?: { mode: "abs"; value: number }
  normalize?: string[]
  sheet_policy?: "first_sheet"
}

type Item = {
  id: number
  report: string
  cell: string
  expected: any
  points: number
  rule?: Rule
  check_label?: string
}

type Master = {
  answer_key: { default_rules: Rule }
  items: Item[]
}

function normalizeValue(v: any, rules: string[] = []) {
  if (v === null || v === undefined) return ""
  let s = typeof v === "string" ? v : String(v)

  for (const r of rules) {
    if (r === "trim") s = s.trim()
    if (r === "comma_remove") s = s.replace(/,/g, "")
    if (r === "blank_to_zero") if (s.trim() === "") s = "0"
    if (r === "zen2han") {
      s = s.replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0))
    }
  }
  return s
}

function asNumber(v: any): number | null {
  if (typeof v === "number") return v
  const s = String(v ?? "").trim()
  if (s === "") return null
  const n = Number(s.replace(/,/g, ""))
  return Number.isFinite(n) ? n : null
}

function compare(expectedRaw: any, actualRaw: any, rule: Rule) {
  const normalize = rule.normalize ?? []
  const expectedN = normalizeValue(expectedRaw, normalize)
  const actualN = normalizeValue(actualRaw, normalize)

  const type = rule.compare_type ?? "auto"
  const tol = rule.tolerance?.value ?? 0

  if (type === "string") return expectedN === actualN

  const eNum = asNumber(expectedN)
  const aNum = asNumber(actualN)
  if (eNum === null || aNum === null) return expectedN === actualN

  return Math.abs(aNum - eNum) <= tol
}

async function loadWorkbookFromFile(file: File) {
  const arrayBuffer = await file.arrayBuffer()
  const wb = new ExcelJS.Workbook()
  // exceljsの型定義都合でanyを噛ませる（実行はOK）
  await wb.xlsx.load(arrayBuffer as any)
  return wb
}

function readCell(wb: ExcelJS.Workbook, cell: string) {
  const ws = wb.worksheets[0]
  const c = ws.getCell(cell)
  const v: any = c.value
  if (v && typeof v === "object" && "result" in v) return (v as any).result
  return v ?? ""
}

function toCSV(rows: Record<string, any>[]) {
  if (rows.length === 0) return ""
  const header = Object.keys(rows[0]).join(",")
  const body = rows
    .map((r) =>
      Object.values(r)
        .map((v) => `"${String(v ?? "").replace(/"/g, '""')}"`)
        .join(",")
    )
    .join("\n")
  return `${header}\n${body}`
}

export async function POST(req: Request) {
  const form = await req.formData()
  const files = form.getAll("files") as File[]

  const masterPath = path.join(process.cwd(), "data", "採点マスタ_自動採点用.json")
  const master: Master = JSON.parse(await fs.readFile(masterPath, "utf-8"))

  const reportKeys = Array.from(new Set(master.items.map((i) => i.report)))
  const submitted: Record<string, ExcelJS.Workbook> = {}

  for (const f of files) {
    const name = f.name
    const hit = reportKeys.find((k) => name.includes(k))
    if (!hit) continue
    submitted[hit] = await loadWorkbookFromFile(f)
  }

  const details: any[] = []
  let total = 0
  const max = master.items.reduce((a, i) => a + i.points, 0)

  for (const item of master.items) {
    const wb = submitted[item.report]
    if (!wb) {
      details.push({
        id: item.id,
        report: item.report,
        cell: item.cell,
        check_label: item.check_label ?? "",
        expected: item.expected,
        actual: "",
        ok: false,
        score: 0,
        points: item.points,
        reason: "missing_file",
      })
      continue
    }

    const rule = { ...master.answer_key.default_rules, ...(item.rule ?? {}) }
    const actual = readCell(wb, item.cell)
    const ok = compare(item.expected, actual, rule)
    const score = ok ? item.points : 0
    total += score

    details.push({
      id: item.id,
      report: item.report,
      cell: item.cell,
      check_label: item.check_label ?? "",
      expected: item.expected,
      actual,
      ok,
      score,
      points: item.points,
    })
  }

  const ngRows = details
    .filter((d) => !d.ok)
    .map((d) => ({
      report: d.report,
      cell: d.cell,
      check_label: d.check_label,
      expected: d.expected,
      actual: d.actual,
      points: d.points,
      reason: d.reason ?? "",
    }))

  const summaryRows = [
    {
      total_score: total,
      max_score: max,
      ng_count: ngRows.length,
      missing_file: details.some((d) => d.reason === "missing_file"),
    },
  ]

  return NextResponse.json({
    total_score: total,
    max_score: max,
    details,
    csv: {
      summary: toCSV(summaryRows),
      ng: toCSV(ngRows),
    },
  })
}
