import { NextRequest, NextResponse } from "next/server";
import * as XLSX from "xlsx";
import JSZip from "jszip";

/**
 * Importa i ToDo (verifiche elaborati) di un progetto Trimble e li restituisce
 * come DUE file, riproducendo esattamente l'export manuale:
 *
 *  1) un ToDo .xlsx (foglio "Todos", stesse 14 colonne dell'export Trimble):
 *     dà Label, Title (elaborato), Description, Tags (NC/OSS), Assignee (disciplina),
 *     Status, autori/date.
 *  2) un ToDo .bcfzip con i commenti (Title/Description/Labels + Comment): dà la
 *     conversazione progettista/ispettore.
 *
 * Il parser della dashboard incrocia i due per Title+Description (flusso già
 * collaudato), quindi la disciplina arriva dall'Assignee e i commenti dal BCF.
 *
 * I ToDo sono una collezione DIVERSA dai BCF Topics (Solibri): questi ultimi si
 * importano con /api/tc/bcfzip.
 *
 * Ritorna JSON con i due file in base64 (una sola chiamata all'app Trimble):
 *   { ok, count, comments_total, base_name, xlsx_base64, bcfzip_base64 }
 *
 * Env (Railway dashboard): TRIMBLE_API_URL (default sotto).
 */

const TRIMBLE_API_URL = (
  process.env.TRIMBLE_API_URL ||
  "https://verifica-elaborati-production.up.railway.app"
).replace(/\/+$/, "");

export const dynamic = "force-dynamic";

// Intestazioni ESATTE del foglio "Todos" dell'export Trimble (attenzione allo
// spazio finale in "Assignee(s) ").
const TODO_HEADERS = [
  "Label",
  "Title",
  "Description",
  "Type",
  "Priority",
  "Status",
  "Completion",
  "Due date",
  "Assignee(s) ",
  "Tags",
  "Created by",
  "Created on",
  "Last modified by",
  "Last modified on",
];

function xmlEscape(v: any): string {
  return String(v ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function buildTodoMarkup(todo: any): string {
  const guid = String(todo?.id ?? "").trim();
  const tags = Array.isArray(todo?.tags) ? todo.tags : [];

  const labels = tags
    .map((l: any) => `    <Labels>${xmlEscape(l)}</Labels>`)
    .join("\n");

  const comments = (Array.isArray(todo?.comments) ? todo.comments : [])
    .map(
      (c: any) => `  <Comment Guid="${xmlEscape(c?.id || "")}">
    <Date>${xmlEscape(c?.date || "")}</Date>
    <Author>${xmlEscape(c?.author || "")}</Author>
    <Comment>${xmlEscape(c?.text || "")}</Comment>
  </Comment>`
    )
    .join("\n");

  return `<?xml version="1.0" encoding="UTF-8"?>
<Markup>
  <Topic Guid="${xmlEscape(guid)}">
    <Title>${xmlEscape(todo?.title || "")}</Title>
    <Description>${xmlEscape(todo?.description || "")}</Description>
    <TopicStatus>${xmlEscape(todo?.status || "")}</TopicStatus>
    <Priority>${xmlEscape(todo?.priority || "")}</Priority>
    <CreationDate>${xmlEscape(todo?.created_on || "")}</CreationDate>
    <CreationAuthor>${xmlEscape(todo?.created_by || "")}</CreationAuthor>
    <ModifiedDate>${xmlEscape(todo?.modified_on || "")}</ModifiedDate>
    <ModifiedAuthor>${xmlEscape(todo?.modified_by || "")}</ModifiedAuthor>
    <AssignedTo>${xmlEscape(todo?.assignees || "")}</AssignedTo>
${labels}
  </Topic>
${comments}
</Markup>`;
}

export async function GET(req: NextRequest) {
  const projectId = (req.nextUrl.searchParams.get("project_id") || "").trim();
  if (!projectId) {
    return NextResponse.json(
      { ok: false, error: "Parametro project_id mancante" },
      { status: 400 }
    );
  }

  try {
    const url =
      `${TRIMBLE_API_URL}/tc/todos?project_id=${encodeURIComponent(projectId)}` +
      `&with_comments=true`;

    const res = await fetch(url, {
      headers: { Accept: "application/json" },
      cache: "no-store",
    });

    if (!res.ok) {
      const text = await res.text().catch(() => "");
      return NextResponse.json(
        { ok: false, error: `Trimble ${res.status}: ${text.slice(0, 300)}` },
        { status: 502 }
      );
    }

    const data = await res.json();
    const todos = Array.isArray(data?.todos) ? data.todos : [];

    if (todos.length === 0) {
      return NextResponse.json(
        { ok: false, error: "Nessun ToDo trovato per questo progetto." },
        { status: 404 }
      );
    }

    // 1) Foglio "Todos" (XLSX)
    const aoa: any[][] = [TODO_HEADERS];
    for (const t of todos) {
      const tags = Array.isArray(t?.tags) ? t.tags.join(", ") : "";
      aoa.push([
        t?.label || "",
        t?.title || "",
        t?.description || "",
        t?.type || "",
        t?.priority || "",
        t?.status || "",
        t?.completion || "",
        t?.due_date || "",
        t?.assignees || "",
        tags,
        t?.created_by || "",
        t?.created_on || "",
        t?.modified_by || "",
        t?.modified_on || "",
      ]);
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Todos");
    const xlsxBuf: Buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    // 2) BCFZIP con i commenti
    const zip = new JSZip();
    let commentsTotal = 0;
    for (const t of todos) {
      const guid = String(t?.id ?? "").trim();
      if (!guid) continue;
      zip.file(`${guid}/markup.bcf`, buildTodoMarkup(t));
      commentsTotal += Array.isArray(t?.comments) ? t.comments.length : 0;
    }
    const bcfBuf: Buffer = await zip.generateAsync({ type: "nodebuffer" });

    return NextResponse.json({
      ok: true,
      count: todos.length,
      comments_total: commentsTotal,
      base_name: `trimble_todo_${projectId}`,
      xlsx_base64: Buffer.from(xlsxBuf).toString("base64"),
      bcfzip_base64: Buffer.from(bcfBuf).toString("base64"),
    });
  } catch (e: any) {
    return NextResponse.json(
      { ok: false, error: e?.message || "Errore import ToDo da Trimble" },
      { status: 502 }
    );
  }
}
