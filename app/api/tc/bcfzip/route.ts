import { NextRequest, NextResponse } from "next/server";
import JSZip from "jszip";

/**
 * Scarica i ToDo (Topics) di un progetto Trimble dall'app verifica-elaborati e
 * li impacchetta in un .bcfzip in memoria, con la STESSA struttura che il
 * parser della dashboard già legge dagli export BCF reali:
 *
 *   {guid}/markup.bcf   ->  <Markup>
 *                             <Topic Guid="...">
 *                               <Title/><Description/><TopicStatus/>...
 *                               <Labels>NC</Labels>
 *                             </Topic>
 *                             <Comment Guid="..."><Date/><Author/><Comment/></Comment>
 *                           </Markup>
 *   {guid}/snapshot.jpg  ->  miniatura del viewpoint (se presente)
 *
 * In questo modo selezionare un progetto Trimble equivale a caricare un vero
 * BCFZIP: il flusso "Analizza" e tutta la logica NC/OSS restano invariati.
 *
 * Env (Railway dashboard): TRIMBLE_API_URL (default sotto).
 */

const TRIMBLE_API_URL = (
  process.env.TRIMBLE_API_URL ||
  "https://verifica-elaborati-production.up.railway.app"
).replace(/\/+$/, "");

export const dynamic = "force-dynamic";

function xmlEscape(v: any): string {
  return String(v ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function buildMarkup(topic: any): string {
  const guid = String(topic?.guid ?? "").trim();

  const labels = (Array.isArray(topic?.labels) ? topic.labels : [])
    .map((l: any) => `    <Labels>${xmlEscape(l)}</Labels>`)
    .join("\n");

  const comments = (Array.isArray(topic?.comments) ? topic.comments : [])
    .map(
      (c: any) => `  <Comment Guid="${xmlEscape(c?.guid || "")}">
    <Date>${xmlEscape(c?.date || "")}</Date>
    <Author>${xmlEscape(c?.author || "")}</Author>
    <Comment>${xmlEscape(c?.text || "")}</Comment>
  </Comment>`
    )
    .join("\n");

  return `<?xml version="1.0" encoding="UTF-8"?>
<Markup>
  <Topic Guid="${xmlEscape(guid)}">
    <Title>${xmlEscape(topic?.title || "")}</Title>
    <Description>${xmlEscape(topic?.description || "")}</Description>
    <TopicStatus>${xmlEscape(topic?.status || "")}</TopicStatus>
    <Priority>${xmlEscape(topic?.priority || "")}</Priority>
    <CreationDate>${xmlEscape(topic?.creation_date || "")}</CreationDate>
    <CreationAuthor>${xmlEscape(topic?.creation_author || "")}</CreationAuthor>
    <ModifiedDate>${xmlEscape(topic?.modified_date || "")}</ModifiedDate>
    <ModifiedAuthor>${xmlEscape(topic?.modified_author || "")}</ModifiedAuthor>
    <AssignedTo>${xmlEscape(topic?.assigned_to || "")}</AssignedTo>
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
      `${TRIMBLE_API_URL}/tc/topics?project_id=${encodeURIComponent(projectId)}` +
      `&with_comments=true&with_snapshots=true`;

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
    const topics = Array.isArray(data?.topics) ? data.topics : [];

    const zip = new JSZip();
    let added = 0;

    for (const t of topics) {
      const guid = String(t?.guid ?? "").trim();
      if (!guid) continue;

      zip.file(`${guid}/markup.bcf`, buildMarkup(t));

      // Snapshot dei viewpoint (immagini BCF sincronizzate con Solibri), scaricati
      // dal backend con with_snapshots: entrano nella cartella del topic così il
      // parser li associa al rilievo e finiscono nell'Export PDF (IMMAGINI NC/OSS).
      const snaps = Array.isArray(t?.snapshots) ? t.snapshots : [];
      snaps.forEach((s: any, i: number) => {
        const b64 = String(s?.image_base64 || "");
        if (!b64) return;
        const ext = String(s?.mime || "").toLowerCase().includes("jpeg") ? "jpg" : "png";
        zip.file(`${guid}/snapshot_${i}.${ext}`, b64, { base64: true });
      });

      // Fallback storico: miniatura inline del viewpoint, se presente.
      const thumb = t?._raw?.viewpoint?.snapshot_thumb;
      if (snaps.length === 0 && typeof thumb === "string" && thumb.length > 0) {
        zip.file(`${guid}/snapshot.jpg`, thumb, { base64: true });
      }

      added += 1;
    }

    if (added === 0) {
      return NextResponse.json(
        { ok: false, error: "Nessun ToDo trovato per questo progetto." },
        { status: 404 }
      );
    }

    const buffer = await zip.generateAsync({ type: "nodebuffer" });

    return new NextResponse(new Uint8Array(buffer), {
      status: 200,
      headers: {
        "Content-Type": "application/zip",
        "Content-Disposition": `attachment; filename="trimble_${projectId}.bcfzip"`,
        "X-Topic-Count": String(added),
      },
    });
  } catch (e: any) {
    return NextResponse.json(
      { ok: false, error: e?.message || "Errore generazione BCFZIP da Trimble" },
      { status: 502 }
    );
  }
}
