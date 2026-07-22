// app/api/controllo-todo-ispettori/trimble/route.ts
//
// Import diretto dei ToDo da Trimble (via backend verifica-elaborati) con i
// markup 2D, e costruzione delle righe di controllo per il report ispettori.
//
// Il backend viene chiamato con with_comments=true e with_snapshots=true, quindi
// ogni ToDo arriva completo di commenti e immagini dei markup. Le eventuali
// Report_Completo.xlsx ed ELENCO_ELABORATI.xlsx (facoltative) servono solo per i
// controlli incrociati su codice elaborato e disciplina.

import { NextRequest, NextResponse } from "next/server";
import * as XLSX from "xlsx";
import { buildChecksFromTrimbleTodos } from "@/lib/controllo-todo-checks";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const TRIMBLE_API_URL = (
  process.env.TRIMBLE_API_URL ||
  "https://verifica-elaborati-production.up.railway.app"
).replace(/\/+$/, "");

async function readXlsxRows(file: File, preferredSheetName?: string): Promise<any[][]> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetName =
    preferredSheetName && workbook.SheetNames.includes(preferredSheetName)
      ? preferredSheetName
      : workbook.SheetNames[0];
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    defval: "",
  }) as any[][];
}

function isUploadedFile(value: FormDataEntryValue | null): value is File {
  return (
    !!value &&
    typeof value === "object" &&
    "arrayBuffer" in value &&
    "size" in value &&
    Number((value as File).size) > 0
  );
}

async function fetchTodosWithSnapshots(projectId: string) {
  const url =
    `${TRIMBLE_API_URL}/tc/todos?project_id=${encodeURIComponent(projectId)}` +
    `&with_comments=true&with_snapshots=true`;

  const res = await fetch(url, {
    headers: { Accept: "application/json" },
    cache: "no-store",
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Trimble ${res.status}: ${text.slice(0, 300)}`);
  }

  const data = await res.json();
  return Array.isArray(data?.todos) ? data.todos : [];
}

async function run(projectId: string, reportFile: File | null, elencoFile: File | null) {
  const [todos, reportRows, elencoRows] = await Promise.all([
    fetchTodosWithSnapshots(projectId),
    reportFile ? readXlsxRows(reportFile, "Verifica Elaborati") : Promise.resolve([] as any[][]),
    elencoFile ? readXlsxRows(elencoFile) : Promise.resolve([] as any[][]),
  ]);

  if (!Array.isArray(todos) || todos.length === 0) {
    return NextResponse.json(
      { ok: false, error: "Nessun ToDo trovato per questo progetto." },
      { status: 404 }
    );
  }

  const { checks, summary } = buildChecksFromTrimbleTodos(todos, reportRows, elencoRows);

  return NextResponse.json({
    ok: true,
    checks,
    summary,
    bcfTopicsCount: todos.length,
    source: "trimble",
    project_id: projectId,
  });
}

// GET semplice: solo project_id (senza controlli incrociati su report/elenco).
export async function GET(req: NextRequest) {
  const projectId = (req.nextUrl.searchParams.get("project_id") || "").trim();
  if (!projectId) {
    return NextResponse.json({ ok: false, error: "Parametro project_id mancante" }, { status: 400 });
  }
  try {
    return await run(projectId, null, null);
  } catch (e: any) {
    return NextResponse.json({ ok: false, error: e?.message || "Errore import ToDo da Trimble" }, { status: 502 });
  }
}

// POST: project_id + Report_Completo.xlsx / ELENCO_ELABORATI.xlsx facoltativi.
export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const projectId = String(formData.get("project_id") || "").trim();
    if (!projectId) {
      return NextResponse.json({ ok: false, error: "Parametro project_id mancante" }, { status: 400 });
    }

    const reportRaw = formData.get("report");
    const elencoRaw = formData.get("elenco");
    const reportFile = isUploadedFile(reportRaw) ? reportRaw : null;
    const elencoFile = isUploadedFile(elencoRaw) ? elencoRaw : null;

    return await run(projectId, reportFile, elencoFile);
  } catch (e: any) {
    return NextResponse.json({ ok: false, error: e?.message || "Errore import ToDo da Trimble" }, { status: 502 });
  }
}
