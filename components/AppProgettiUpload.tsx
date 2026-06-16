"use client";

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { supabase } from "../lib/supabase";

function getElaboratoKey(r: any) {
  return (
    r.elaborato ||
    r.codiceElaborato ||
    r.codice_elaborato ||
    r.codice ||
    r.titolo ||
    r.id ||
    ""
  );
}

// Chiave tecnica per conteggiare gli elaborati una sola volta.
// Esempio: "T1266 0000 PE OS GTA..." e "T1266-0000-PE-OS-GTA..." diventano uguali.
function normalizeElaboratoCode(v = "") {
  return String(v || "")
    .toUpperCase()
    .replace(/\.PDF$/i, "")
    .replace(/[^A-Z0-9]/g, "");
}

function getElaboratoNormalizedKey(r: any) {
  const raw = getElaboratoKey(r);
  const normalized = normalizeElaboratoCode(raw);
  return normalized || cleanPdfText(raw);
}

function getElaboratoDisplay(r: any) {
  return cleanPdfText(getElaboratoKey(r)) || "Elaborato non identificato";
}

function translateStatus(status = "") {
  const value = String(status || "").trim();

  if (value === "Closed") return "Chiusa";
  if (value === "New") return "Aperta";
  if (value === "Waiting") return "In attesa";
  if (value === "Unknown") return "Non definito";

  return value || "Non definito";
}

function commentsToText(comments: any[]) {
  if (!Array.isArray(comments) || comments.length === 0) return "";

  return comments
    .map((c: any) => {
      const role = c.role ? `[${c.role}] ` : "";
      const author = c.author || "Autore non indicato";
      const date = c.date ? ` - ${c.date}` : "";
      const comment = c.comment || "";
      return `${role}${author}${date}\n${comment}`;
    })
    .join("\n\n");
}

function exportExcel(nomeFile: string, dati: any[]) {
  if (!dati || dati.length === 0) return;

  const ws = XLSX.utils.json_to_sheet(dati);
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, ws, "Export");
  XLSX.writeFile(wb, `${nomeFile}.xlsx`);
}

function cleanPdfText(value: any) {
  return String(value ?? "")
    .replace(/[\u0000-\u001f\u007f]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}


function buildElaboratiDisciplinaSummary(rows: any[]) {
  const grouped: Record<string, { elaborato: string; nc: number; oss: number }> = {};

  rows.forEach((r: any) => {
    const elaborato = cleanPdfText(getElaboratoKey(r)) || "Elaborato non identificato";
    const tipo = cleanPdfText(r.tipo).toUpperCase();

    if (!grouped[elaborato]) {
      grouped[elaborato] = { elaborato, nc: 0, oss: 0 };
    }

    if (tipo === "NC") {
      grouped[elaborato].nc += 1;
    }

    // Le righe "Da NC a OSS" vanno conteggiate come OSS.
    if (tipo === "OSS" || tipo === "DA NC A OSS") {
      grouped[elaborato].oss += 1;
    }
  });

  return Object.values(grouped)
    .sort((a, b) => a.elaborato.localeCompare(b.elaborato, "it"))
    .map((item, index) => {
      const esito = item.nc === 0 && item.oss === 0
        ? "Nessuna NC/OSS"
        : `NC: ${item.nc} - OSS: ${item.oss}`;

      return [
        String(index + 1),
        item.elaborato,
        String(item.nc),
        String(item.oss),
        esito,
      ];
    });
}

type PdfHeaderData = {
  committente?: string;
  oggetto?: string;
  codiceCommessa?: string;
  ispettore?: string;
  firma?: string;
  notaRicezione?: string;
  dataRicezione?: string;
};

const ITS_BLUE = [30, 117, 192] as [number, number, number];
const ITS_DARK_BLUE = [26, 92, 150] as [number, number, number];
const ITS_LIGHT_BLUE = [217, 232, 245] as [number, number, number];
const ITS_TABLE_LIGHT = [234, 243, 251] as [number, number, number];
const ITS_GRAY = [245, 245, 245] as [number, number, number];
const ITS_LOGO_BASE64 = "data:image/png;base64,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";

function detectDisciplinaForPdf(rows: any[], title = "") {
  const disciplines = Array.from(
    new Set(
      (rows || [])
        .map((r: any) => cleanPdfText(r?.disciplina))
        .filter(Boolean)
    )
  );

  if (disciplines.length === 1) return disciplines[0];
  if (disciplines.length > 1) return "Discipline multiple";

  const match = cleanPdfText(title).match(/Disciplina:\s*(.+)$/i);
  return match ? match[1].trim() : "";
}

function headerValue(value: any) {
  const cleaned = cleanPdfText(value);
  return cleaned || "____________________________";
}

function getImageFormatFromDataUrl(dataUrl = "") {
  const lower = String(dataUrl || "").toLowerCase();

  if (lower.startsWith("data:image/jpeg") || lower.startsWith("data:image/jpg")) return "JPEG";
  if (lower.startsWith("data:image/webp")) return "WEBP";
  return "PNG";
}

function addPdfPageNumber(doc: jsPDF) {
  const pageCount = doc.getNumberOfPages();
  const pageWidthCurrent = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();

  doc.setFontSize(8);
  doc.setTextColor(100);
  doc.text(`Pagina ${pageCount}`, pageWidthCurrent - 10, pageHeight - 6, { align: "right" });
}

function exportDetailPdf(rows: any[], title = "", headerData: PdfHeaderData = {}) {
  if (!rows || rows.length === 0) return;

  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  const generatedAt = new Date().toLocaleString("it-IT");
  const disciplina = detectDisciplinaForPdf(rows, title);
  const pageWidth = doc.internal.pageSize.getWidth();
  const marginX = 10;
  const headerY = 8;
  const headerW = pageWidth - marginX * 2;

  // Testata solo sulla prima pagina, con colori e logo ricavati dal template Excel ITS.
  doc.setFillColor(ITS_BLUE[0], ITS_BLUE[1], ITS_BLUE[2]);
  doc.rect(marginX, headerY, headerW, 3, "F");

  doc.setDrawColor(210, 210, 210);
  doc.setLineWidth(0.25);
  doc.rect(marginX, headerY + 3, headerW, 47);

  doc.setFillColor(255, 255, 255);
  doc.rect(marginX, headerY + 3, 66, 17, "F");
  try {
    doc.addImage(ITS_LOGO_BASE64, "PNG", marginX + 4, headerY + 5, 35, 14);
  } catch (error) {
    doc.setTextColor(30, 117, 192);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text("ITS", marginX + 5, headerY + 13);
  }

  doc.setFillColor(255, 255, 255);
  doc.rect(marginX + 66, headerY + 3, 142, 17, "F");
  doc.setTextColor(ITS_BLUE[0], ITS_BLUE[1], ITS_BLUE[2]);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(17);
  doc.text("SCHEDA DI ISPEZIONE", marginX + 137, headerY + 12, { align: "center" });

  doc.setFontSize(10);
  doc.text((disciplina || "").slice(0, 90), marginX + 137, headerY + 17, { align: "center" });

  doc.setFillColor(ITS_LIGHT_BLUE[0], ITS_LIGHT_BLUE[1], ITS_LIGHT_BLUE[2]);
  doc.rect(marginX + 208, headerY + 3, headerW - 208, 17, "F");
  doc.setTextColor(80, 80, 80);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(8);
  doc.text("Export dettaglio NC / OSS", pageWidth - marginX - 4, headerY + 10, { align: "right" });
  doc.setFont("helvetica", "normal");
  doc.text("Nexcommon QC", pageWidth - marginX - 4, headerY + 15, { align: "right" });

  // Sezione 1 - Dati commessa
  doc.setFillColor(ITS_BLUE[0], ITS_BLUE[1], ITS_BLUE[2]);
  doc.rect(marginX, headerY + 20, headerW, 7, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(9);
  doc.text("1. DATI COMMESSA", marginX + 1, headerY + 25);

  const leftX = marginX;
  const labelW = 68;
  const midX = marginX + 168;
  const row1Y = headerY + 27;
  const rowH = 7;

  function labelCell(label: string, x: number, y: number, w: number) {
    doc.setFillColor(ITS_GRAY[0], ITS_GRAY[1], ITS_GRAY[2]);
    doc.rect(x, y, w, rowH, "FD");
    doc.setTextColor(80, 80, 80);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(7.5);
    doc.text(label, x + w - 2, y + 4.7, { align: "right" });
  }

  function valueCell(value: string, x: number, y: number, w: number) {
    doc.setFillColor(255, 255, 255);
    doc.rect(x, y, w, rowH, "D");
    doc.setTextColor(35, 35, 35);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.5);
    doc.text(headerValue(value).slice(0, Math.max(15, Math.floor(w / 2.1))), x + 2, y + 4.7);
  }

  labelCell("Committente / Stazione Appaltante:", leftX, row1Y, labelW);
  valueCell(headerData.committente || "", leftX + labelW, row1Y, headerW - labelW);

  labelCell("Oggetto del Contratto / Accordo Quadro:", leftX, row1Y + rowH, labelW);
  valueCell(headerData.oggetto || "", leftX + labelW, row1Y + rowH, headerW - labelW);

  labelCell("Codice commessa ITS:", leftX, row1Y + rowH * 2, labelW);
  valueCell(headerData.codiceCommessa || "", leftX + labelW, row1Y + rowH * 2, 100);
  labelCell("Firma:", midX, row1Y + rowH * 2, 36);
  valueCell(headerData.firma || "", midX + 36, row1Y + rowH * 2, headerW - 204);

  labelCell("Nome ispettore redattore:", leftX, row1Y + rowH * 3, labelW);
  valueCell(headerData.ispettore || "", leftX + labelW, row1Y + rowH * 3, 100);
  labelCell("Data Ricezione:", midX, row1Y + rowH * 3, 36);
  valueCell(headerData.dataRicezione || "", midX + 36, row1Y + rowH * 3, headerW - 204);

  labelCell("Nota di Ricezione Elaborati:", leftX, row1Y + rowH * 4, labelW);
  valueCell(headerData.notaRicezione || "", leftX + labelW, row1Y + rowH * 4, 100);
  labelCell("Data esportazione:", midX, row1Y + rowH * 4, 36);
  valueCell(generatedAt, midX + 36, row1Y + rowH * 4, headerW - 204);

  // Sezione 4 - Rilievi, coerente con il template.
  const sectionY = headerY + 70;
  doc.setFillColor(ITS_BLUE[0], ITS_BLUE[1], ITS_BLUE[2]);
  doc.rect(marginX, sectionY, headerW, 7, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(9);
  doc.text("4. RILIEVI ITS CONTROLLI TECNICI   (*) NC = Non conformità   |   OSS = Osservazione", marginX + 1, sectionY + 5);

  doc.setTextColor(80, 80, 80);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(7.5);
  doc.text(`Numero righe: ${rows.length}`, pageWidth - marginX, sectionY - 2, { align: "right" });

  const tableRows = rows.map((r: any, i: number) => {
    const allComments = Array.isArray(r.comments) ? r.comments : [];
    return [
      String(i + 1),
      cleanPdfText(r.id),
      cleanPdfText(r.tipo),
      cleanPdfText(r.disciplina),
      cleanPdfText(getElaboratoKey(r)),
      cleanPdfText(r.descrizione),
      cleanPdfText(commentsToText(allComments)),
      cleanPdfText(translateStatus(r.stato)),
    ];
  });

  autoTable(doc, {
    startY: sectionY + 7,
    head: [["N.", "ID rilievo", "Tipologia", "Disciplina", "Elaborato", "Descrizione", "Gestione rilievo", "Stato"]],
    body: tableRows,
    theme: "grid",
    styles: {
      font: "helvetica",
      fontSize: 7,
      cellPadding: 1.8,
      overflow: "linebreak",
      valign: "top",
      lineColor: [200, 200, 200],
      lineWidth: 0.1,
    },
    headStyles: {
      fillColor: ITS_DARK_BLUE,
      textColor: [255, 255, 255],
      fontStyle: "bold",
    },
    alternateRowStyles: { fillColor: ITS_TABLE_LIGHT },
    columnStyles: {
      0: { cellWidth: 9, halign: "center" },
      1: { cellWidth: 22 },
      2: { cellWidth: 22 },
      3: { cellWidth: 25 },
      4: { cellWidth: 38 },
      5: { cellWidth: 55 },
      6: { cellWidth: 82 },
      7: { cellWidth: 20 },
    },
    margin: { left: 10, right: 10, top: 10, bottom: 12 },
    didDrawPage: () => addPdfPageNumber(doc),
  });


  const summaryRows = buildElaboratiDisciplinaSummary(rows);
  const lastAutoTable = (doc as any).lastAutoTable;
  let summaryStartY = (lastAutoTable?.finalY || 20) + 10;

  if (summaryStartY > doc.internal.pageSize.getHeight() - 50) {
    doc.addPage();
    summaryStartY = 18;
  }

  doc.setTextColor(ITS_BLUE[0], ITS_BLUE[1], ITS_BLUE[2]);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("ELENCO ELABORATI DELLA DISCIPLINA E CONTEGGIO NC/OSS", 10, summaryStartY);

  autoTable(doc, {
    startY: summaryStartY + 4,
    head: [["N.", "Elaborato", "NC", "OSS", "Esito"]],
    body: summaryRows,
    theme: "grid",
    styles: {
      font: "helvetica",
      fontSize: 7.5,
      cellPadding: 2,
      overflow: "linebreak",
      valign: "top",
      lineColor: [200, 200, 200],
      lineWidth: 0.1,
    },
    headStyles: {
      fillColor: ITS_DARK_BLUE,
      textColor: [255, 255, 255],
      fontStyle: "bold",
    },
    alternateRowStyles: { fillColor: ITS_TABLE_LIGHT },
    columnStyles: {
      0: { cellWidth: 10, halign: "center" },
      1: { cellWidth: 95 },
      2: { cellWidth: 16, halign: "center" },
      3: { cellWidth: 16, halign: "center" },
      4: { cellWidth: 55 },
    },
    margin: { left: 10, right: 10, top: 10, bottom: 12 },
    didDrawPage: () => addPdfPageNumber(doc),
  });

  const rowsWithImages = rows.filter((r: any) => Boolean(r.snapshotDataUrl));

  if (rowsWithImages.length > 0) {
    const lastSummaryTable = (doc as any).lastAutoTable;
    let imagesStartY = (lastSummaryTable?.finalY || summaryStartY) + 12;
    const pageHeight = doc.internal.pageSize.getHeight();
    const imageMaxW = 135;
    const imageMaxH = 78;
    const imageX = 10;
    const textX = imageX + imageMaxW + 8;
    const contentW = pageWidth - textX - 10;

    if (imagesStartY > pageHeight - 45) {
      doc.addPage();
      imagesStartY = 18;
    }

    doc.setTextColor(ITS_BLUE[0], ITS_BLUE[1], ITS_BLUE[2]);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("IMMAGINI NC/OSS", 10, imagesStartY);

    let currentY = imagesStartY + 8;

    rowsWithImages.forEach((r: any, index: number) => {
      const requiredHeight = imageMaxH + 22;

      if (currentY > pageHeight - requiredHeight - 12) {
        doc.addPage();
        currentY = 18;
      }

      const idRilievo = cleanPdfText(r.id || r.idTodo || `Rilievo ${index + 1}`);
      const tipologia = cleanPdfText(r.tipo || "");
      const disciplina = cleanPdfText(r.disciplina || "");
      const elaborato = cleanPdfText(getElaboratoKey(r));
      const descrizione = cleanPdfText(r.descrizione || "");

      doc.setFillColor(ITS_DARK_BLUE[0], ITS_DARK_BLUE[1], ITS_DARK_BLUE[2]);
      doc.rect(10, currentY, pageWidth - 20, 7, "F");
      doc.setTextColor(255, 255, 255);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(8);
      doc.text(`Immagine ${index + 1} - ${tipologia || "Rilievo"} - ${idRilievo}`, 12, currentY + 4.8);

      currentY += 10;

      try {
        const props = doc.getImageProperties(r.snapshotDataUrl);
        const ratio = props.width && props.height ? props.width / props.height : 1.6;
        let imageW = imageMaxW;
        let imageH = imageW / ratio;

        if (imageH > imageMaxH) {
          imageH = imageMaxH;
          imageW = imageH * ratio;
        }

        doc.addImage(
          r.snapshotDataUrl,
          getImageFormatFromDataUrl(r.snapshotDataUrl),
          imageX,
          currentY,
          imageW,
          imageH
        );

        doc.setDrawColor(200, 200, 200);
        doc.rect(imageX, currentY, imageW, imageH);
      } catch (error) {
        doc.setDrawColor(200, 200, 200);
        doc.rect(imageX, currentY, imageMaxW, 22);
        doc.setTextColor(120, 120, 120);
        doc.setFont("helvetica", "normal");
        doc.setFontSize(8);
        doc.text("Immagine BCF non visualizzabile nel PDF", imageX + 3, currentY + 12);
      }

      doc.setTextColor(35, 35, 35);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(8);
      doc.text("ID rilievo:", textX, currentY + 4);
      doc.text("Tipologia:", textX, currentY + 12);
      doc.text("Disciplina:", textX, currentY + 20);
      doc.text("Elaborato:", textX, currentY + 28);
      doc.text("Descrizione:", textX, currentY + 40);

      doc.setFont("helvetica", "normal");
      doc.text(idRilievo.slice(0, 75), textX + 24, currentY + 4);
      doc.text(tipologia.slice(0, 40), textX + 24, currentY + 12);
      doc.text(disciplina.slice(0, 40), textX + 24, currentY + 20);

      const elaboratoLines = doc.splitTextToSize(elaborato || "Elaborato non identificato", contentW);
      doc.text(elaboratoLines.slice(0, 2), textX + 24, currentY + 28);

      const descrizioneLines = doc.splitTextToSize(descrizione || "", contentW);
      doc.text(descrizioneLines.slice(0, 5), textX, currentY + 46);

      currentY += imageMaxH + 22;
    });

    addPdfPageNumber(doc);
  }

  doc.save("Dettaglio_selezione.pdf");
}
function toDashboardExportRows(rows: any[]) {
  return rows.map((r: any) => ({
    ID_Rilievo: r.id || "",
    "Tipologia rilievo": r.tipo || "",
    Disciplina: r.disciplina || "",
    Elaborato: getElaboratoKey(r),
    Descrizione: r.descrizione || "",
    "Gestione rilievo": commentsToText(Array.isArray(r.comments) ? r.comments : []),
    "Immagine BCF": r.snapshotDataUrl ? "Presente" : "",
    Stato: translateStatus(r.stato),
  }));
}

function toChartExportRows(rows: any[]) {
  return rows.map((r: any) => ({
    Voce: r.label,
    Totale: r.value,
    NC: r.nc ?? "",
    OSS: r.oss ?? "",
  }));
}

function ExportButton({ children, onClick }: any) {
  return (
    <button
      onClick={onClick}
      style={{
        background: "#0f172a",
        color: "white",
        border: "none",
        borderRadius: 10,
        padding: "8px 12px",
        cursor: "pointer",
        fontWeight: 700,
        fontSize: 12,
      }}
    >
      {children}
    </button>
  );
}

function Card({ children, onClick, active = false }: any) {
  return (
    <div
      onClick={onClick}
      style={{
        background: active ? "#e0f2fe" : "white",
        borderRadius: 16,
        padding: 16,
        border: active ? "2px solid #0284c7" : "1px solid #e2e8f0",
        cursor: onClick ? "pointer" : "default",
        boxShadow: "0 6px 18px rgba(15,23,42,.06)",
      }}
    >
      {children}
    </div>
  );
}

function KPI({ title, value, subtitle, onClick, active, colorValue }: any) {
  return (
    <Card onClick={onClick} active={active}>
      <div style={{ fontSize: 13, color: "#64748b" }}>{title}</div>
      <div style={{ fontSize: 34, fontWeight: 800, color: colorValue || "#0f172a" }}>
        {value}
      </div>
      {subtitle && <div style={{ fontSize: 12, color: "#64748b" }}>{subtitle}</div>}
    </Card>
  );
}

function BarList({ title, data, onClick, activeKey, onExport }: any) {
  const max = Math.max(...data.map((d: any) => d.value), 1);

  return (
    <Card>
      <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center" }}>
        <h3 style={{ marginTop: 0 }}>{title}</h3>
        {onExport && <ExportButton onClick={onExport}>Export Excel</ExportButton>}
      </div>

      {data.length === 0 && (
        <div style={{ color: "#64748b", fontSize: 13 }}>Nessun dato disponibile</div>
      )}

      {data.map((d: any) => {
        const itemKey = d.key || d.label;

        return (
        <div
          key={itemKey}
          onClick={() => onClick(itemKey, d.label)}
          style={{
            marginBottom: 12,
            cursor: "pointer",
            background: activeKey === itemKey ? "#e0f2fe" : "transparent",
            borderRadius: 10,
            padding: 6,
          }}
        >
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
            <span>{d.label}</span>
            <b>{d.value}</b>
          </div>

          {(d.nc !== undefined || d.oss !== undefined) && (
            <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>
              NC: {d.nc || 0} · OSS: {d.oss || 0}
            </div>
          )}

          <div
            style={{
              height: 12,
              background: "#e2e8f0",
              borderRadius: 99,
              overflow: "hidden",
              marginTop: 4,
            }}
          >
            <div
              style={{
                width: `${(d.value / max) * 100}%`,
                height: 12,
                background: "#0f172a",
              }}
            />
          </div>
        </div>
        );
      })}
    </Card>
  );
}

function ImportSummary({ importedFiles }: any) {
  if (!importedFiles?.length) return null;

  return (
    <Card>
      <h3 style={{ marginTop: 0 }}>File importati</h3>
      <div style={{ display: "grid", gap: 8 }}>
        {importedFiles.map((f: any, i: number) => (
          <div key={`${f.fileName}-${i}`} style={{ fontSize: 13 }}>
            <b>{f.fileName}</b>{" "}
            <span style={{ color: "#64748b" }}>
              {f.type === "xlsx" && `- Excel letto: ${f.rows || 0} righe`}
              {(f.type === "bcf" || f.type === "bcfzip" || f.type === "zip") &&
                `- BCF letto: ${f.markupCount || 0} topic, ${f.comments || 0} commenti${f.snapshots ? `, ${f.snapshots} immagini` : ""}${f.docx ? `, ${f.docx} schede Word` : ""}`}
              {f.type === "docx" &&
                `- Scheda Word letta: ${f.rilievi || 0} rilievi, ${f.elaboratiSenzaRilievi || 0} elaborati senza rilievi${f.disciplina ? ` - ${f.disciplina}` : ""}`}
            </span>
          </div>
        ))}
      </div>
    </Card>
  );
}

function CommentList({ comments, emptyText = "" }: any) {
  if (!comments || comments.length === 0) {
    return <span style={{ color: "#94a3b8" }}>{emptyText}</span>;
  }

  return (
    <div style={{ minWidth: 320 }}>
      {comments.map((c: any, idx: number) => (
        <div
          key={`${c.date || "data"}-${c.author || "autore"}-${idx}`}
          style={{
            marginBottom: 10,
            paddingBottom: 10,
            borderBottom: idx < comments.length - 1 ? "1px solid #e2e8f0" : "none",
            whiteSpace: "pre-wrap",
          }}
        >
          <div style={{ fontWeight: 700 }}>
            {c.role ? `[${c.role}] ` : ""}
            {c.author || "Autore non indicato"}
          </div>
          {c.date && <div style={{ color: "#64748b", fontSize: 12 }}>{c.date}</div>}
          <div style={{ marginTop: 4 }}>{c.comment || ""}</div>
        </div>
      ))}
    </div>
  );
}

function DetailPanel({ rows, title, onReset }: any) {
  const [pdfHeader, setPdfHeader] = useState<PdfHeaderData>({});

  if (!title) return null;

  function updatePdfHeader(field: keyof PdfHeaderData, value: string) {
    setPdfHeader((prev) => ({ ...prev, [field]: value }));
  }

  const inputStyle = {
    width: "100%",
    border: "1px solid #cbd5e1",
    borderRadius: 8,
    padding: "8px 10px",
    fontSize: 12,
    background: "white",
  };

  const labelStyle = {
    display: "block",
    fontSize: 11,
    color: "#475569",
    fontWeight: 700,
    marginBottom: 4,
  };

  return (
    <Card>
      <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "center" }}>
        <h2 style={{ marginTop: 0 }}>Dettaglio selezione: {title}</h2>
        <div style={{ display: "flex", gap: 10 }}>
          <ExportButton
            onClick={() =>
              exportExcel(
                "Dettaglio_selezione",
                toDashboardExportRows(rows)
              )
            }
          >
            Export Excel
          </ExportButton>
          <ExportButton onClick={() => exportDetailPdf(rows, title, pdfHeader)}>
            Export PDF
          </ExportButton>
          <button
            onClick={onReset}
            style={{
              border: "1px solid #cbd5e1",
              background: "white",
              borderRadius: 10,
              padding: "8px 12px",
              cursor: "pointer",
              height: 38,
            }}
          >
            Reset selezione
          </button>
        </div>
      </div>

      <div
        style={{
          marginTop: 12,
          marginBottom: 14,
          padding: 12,
          border: "1px solid #e2e8f0",
          borderRadius: 12,
          background: "#f8fafc",
        }}
      >
        <div style={{ fontWeight: 800, marginBottom: 8, color: "#0f172a" }}>
          Dati testata PDF - compilazione opzionale
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10 }}>
          <div>
            <label style={labelStyle}>Committente / Stazione Appaltante</label>
            <input
              style={inputStyle}
              value={pdfHeader.committente || ""}
              onChange={(e) => updatePdfHeader("committente", e.target.value)}
              placeholder="es. RPR - Risorse per Roma S.p.A."
            />
          </div>
          <div>
            <label style={labelStyle}>Codice commessa ITS</label>
            <input
              style={inputStyle}
              value={pdfHeader.codiceCommessa || ""}
              onChange={(e) => updatePdfHeader("codiceCommessa", e.target.value)}
              placeholder="es. IT25063"
            />
          </div>
          <div>
            <label style={labelStyle}>Nota di Ricezione Elaborati</label>
            <input
              style={inputStyle}
              value={pdfHeader.notaRicezione || ""}
              onChange={(e) => updatePdfHeader("notaRicezione", e.target.value)}
              placeholder="es. 25063AR-PE1-RE-0001"
            />
          </div>
          <div>
            <label style={labelStyle}>Oggetto del Contratto / Accordo Quadro</label>
            <input
              style={inputStyle}
              value={pdfHeader.oggetto || ""}
              onChange={(e) => updatePdfHeader("oggetto", e.target.value)}
              placeholder="Descrizione sintetica della commessa"
            />
          </div>
          <div>
            <label style={labelStyle}>Nome ispettore redattore</label>
            <input
              style={inputStyle}
              value={pdfHeader.ispettore || ""}
              onChange={(e) => updatePdfHeader("ispettore", e.target.value)}
              placeholder="es. Ing. Mario Rossi"
            />
          </div>
          <div>
            <label style={labelStyle}>Data Ricezione / Firma</label>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
              <input
                style={inputStyle}
                value={pdfHeader.dataRicezione || ""}
                onChange={(e) => updatePdfHeader("dataRicezione", e.target.value)}
                placeholder="es. 15/06/2026"
              />
              <input
                style={inputStyle}
                value={pdfHeader.firma || ""}
                onChange={(e) => updatePdfHeader("firma", e.target.value)}
                placeholder="Firma"
              />
            </div>
          </div>
        </div>
      </div>

      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#f1f5f9" }}>
              <th style={th}>N.</th>
              <th style={th}>ID_Rilievo</th>
              <th style={th}>Tipologia rilievo</th>
              <th style={th}>Disciplina</th>
              <th style={th}>Elaborato</th>
              <th style={th}>Descrizione</th>
              <th style={th}>Gestione rilievo</th>
              <th style={th}>Stato</th>
            </tr>
          </thead>

          <tbody>
            {rows.map((r: any, i: number) => {
              const allComments = Array.isArray(r.comments) ? r.comments : [];

              return (
                <tr key={`${r.id}-${i}`}>
                  <td style={{ ...td, textAlign: "center", fontWeight: 700 }}>{i + 1}</td>
                  <td style={td}>{r.id}</td>
                  <td style={td}>{r.tipo}</td>
                  <td style={td}>{r.disciplina}</td>
                  <td style={td}>{getElaboratoKey(r)}</td>
                  <td style={td}>{r.descrizione}</td>
                  <td style={td}>
                    <CommentList comments={allComments} emptyText="Nessun commento" />
                  </td>
                  <td style={td}>{translateStatus(r.stato)}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </Card>
  );
}

const th = {
  border: "1px solid #e2e8f0",
  padding: 8,
  textAlign: "left" as const,
};

const td = {
  border: "1px solid #e2e8f0",
  padding: 8,
  verticalAlign: "top" as const,
};

const defaultDashboardModules = [
  {
    title: "Nota di Ricezione Elaborati",
    subtitle: "Modulo operativo attivo",
    url: "https://verifica-elaborati-production.up.railway.app",
    active: true,
    visible: true,
    external: true,
    sort_order: 1,
  },
];

export default function AppProgettiUpload() {
  React.useEffect(() => {
    const isAuthenticated = localStorage.getItem("nexcommon_verify_auth");

    if (isAuthenticated !== "true") {
      window.location.href = "/login";
    }
  }, []);

  const [files, setFiles] = useState<File[]>([]);
  const [rows, setRows] = useState<any[]>([]);
  const [selection, setSelection] = useState<any>(null);
  const [importedFiles, setImportedFiles] = useState<any[]>([]);
  const [dashboardModules, setDashboardModules] = useState<any[]>(defaultDashboardModules);
  const [error, setError] = useState("");

  React.useEffect(() => {
    async function loadDashboardModules() {
      const { data, error } = await supabase
        .from("dashboard_modules")
        .select("*")
        .eq("visible", true)
        .order("sort_order", { ascending: true });

      if (error) {
        console.error("Errore caricamento moduli dashboard:", error);
        return;
      }

      setDashboardModules(data && data.length > 0 ? data : defaultDashboardModules);
    }

    loadDashboardModules();
  }, []);

  async function generaDashboard() {
    setError("");

    const fd = new FormData();
    files.forEach((f) => fd.append("files", f));

    const res = await fetch("/api/parse", {
      method: "POST",
      body: fd,
    });

    const data = await res.json();

    if (!res.ok || data.ok === false) {
      setError(data.error || "Errore durante la lettura dei file");
      setRows([]);
      setImportedFiles([]);
      setSelection(null);
      return;
    }

    setRows(data.rows || []);
    setImportedFiles(data.importedFiles || []);
    setSelection(null);
  }

  const enrichedRows = useMemo(() => {
    return rows.map((r) => ({
      ...r,
      tipologiaNcOss: r.tipologiaNcOss || r.tipologiaDocumento || "",
      tipologia: r.tipologiaNcOss || r.tipologiaDocumento || "",
      elaboratoKey: getElaboratoNormalizedKey(r),
      elaboratoDisplay: getElaboratoDisplay(r),
      stato: translateStatus(r.stato),
    }));
  }, [rows]);

  const elaboratiTot = new Set(enrichedRows.map((r) => r.elaboratoKey).filter(Boolean)).size;
  const elaboratiNC = new Set(enrichedRows.filter((r) => r.tipo === "NC").map((r) => r.elaboratoKey).filter(Boolean)).size;
  const elaboratiOSS = new Set(enrichedRows.filter((r) => r.tipo === "OSS").map((r) => r.elaboratoKey).filter(Boolean)).size;
  const elaboratiOK = new Set(enrichedRows.filter((r) => r.tipo === "Nessun rilievo").map((r) => r.elaboratoKey).filter(Boolean)).size;

  const totaleNC = enrichedRows.filter((r) => r.tipo === "NC").length;
  const totaleOSS = enrichedRows.filter((r) => r.tipo === "OSS").length;
  const totaleNCOSS = totaleNC + totaleOSS;

  const rilieviNCOSS = enrichedRows.filter(
    (r) => r.tipo === "NC" || r.tipo === "OSS" || r.tipo === "Da NC a OSS"
  );

  const daVerificareISP = enrichedRows.filter((r) => r.chiDeveAgire === "ISP").length;
  const daRisponderePRG = enrichedRows.filter((r) => r.chiDeveAgire === "PRG").length;

  const filteredRows = useMemo(() => {
    if (!selection) return [];

    if (selection.type === "kpi") {
      if (selection.value === "totali") return enrichedRows;
      if (selection.value === "nc") return enrichedRows.filter((r) => r.tipo === "NC");
      if (selection.value === "oss") return enrichedRows.filter((r) => r.tipo === "OSS");
      if (selection.value === "nessun") return enrichedRows.filter((r) => r.tipo === "Nessun rilievo");
      if (selection.value === "risoluzione-prg") return rilieviNCOSS.filter((r) => r.hasPrgComment);
      if (selection.value === "da-verificare-isp") return enrichedRows.filter((r) => r.chiDeveAgire === "ISP");
      if (selection.value === "da-rispondere-prg") return enrichedRows.filter((r) => r.chiDeveAgire === "PRG");
    }

    if (selection.type === "tipo") return enrichedRows.filter((r) => r.tipo === selection.value);
    if (selection.type === "disciplina") return enrichedRows.filter((r) => r.disciplina === selection.value);
    if (selection.type === "elaborato") return enrichedRows.filter((r) => r.elaboratoKey === selection.value);

    return [];
  }, [enrichedRows, selection, rilieviNCOSS]);

  const discipline: any = {};
  const esiti: any = {};
  const rilieviPerElaborato: any = {};

  enrichedRows.forEach((r) => {
    const d = r.disciplina || "Non assegnata";
    discipline[d] = (discipline[d] || 0) + 1;

    const e = r.tipo || "Rilievo mancante";
    esiti[e] = (esiti[e] || 0) + 1;

    const elaboratoKey = r.elaboratoKey || "Elaborato non identificato";
    const elaboratoLabel = r.elaboratoDisplay || getElaboratoDisplay(r);

    if (!rilieviPerElaborato[elaboratoKey]) {
      rilieviPerElaborato[elaboratoKey] = {
        key: elaboratoKey,
        label: elaboratoLabel,
        value: 0,
        nc: 0,
        oss: 0,
      };
    }

    if (r.tipo === "NC") {
      rilieviPerElaborato[elaboratoKey].value += 1;
      rilieviPerElaborato[elaboratoKey].nc += 1;
    }

    if (r.tipo === "OSS") {
      rilieviPerElaborato[elaboratoKey].value += 1;
      rilieviPerElaborato[elaboratoKey].oss += 1;
    }
  });

  const disciplineData = Object.entries(discipline)
    .map(([label, value]) => ({ label, value: Number(value) }))
    .sort((a: any, b: any) => b.value - a.value);

  const esitiData = Object.entries(esiti)
    .map(([label, value]) => ({ label, value: Number(value) }))
    .sort((a: any, b: any) => b.value - a.value);

  const rilieviPerElaboratoData = Object.values(rilieviPerElaborato)
    .filter((d: any) => d.value > 0)
    .sort((a: any, b: any) => b.value - a.value);

  const selectionTitle = selection
    ? `${selection.label}: ${selection.valueLabel || selection.value}`
    : "";

  return (
    <main
      style={{
        padding: 30,
        background: "#f1f5f9",
        minHeight: "100vh",
        fontFamily: "Arial, sans-serif",
      }}
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          gap: 20,
          alignItems: "flex-start",
        }}
      >
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <img src="/logo_nexcommon.png" alt="Nexcommon" style={{ height: 34, objectFit: "contain" }} />
            <div style={{ fontSize: 13, color: "#64748b" }}>
              Piattaforma creata da Nexcommon S.r.l.
            </div>
          </div>

          <div style={{ marginTop: 22, display: "flex", alignItems: "center", gap: 18 }}>
            <img
              src="/logo_its.png"
              alt="ITS Controlli Tecnici S.p.A."
              style={{
                height: 58,
                objectFit: "contain",
                background: "#0f172a",
                padding: 8,
                borderRadius: 8,
              }}
            />

            <div>
              <h1 style={{ margin: 0, fontSize: 30 }}>
                ITS Controlli Tecnici S.p.A.
              </h1>
              <div style={{ color: "#64748b", fontSize: 14 }}>
                Dashboard verifiche elaborati / NC / OSS
              </div>
            </div>
          </div>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          <button
            onClick={() => {
              localStorage.removeItem("nexcommon_verify_auth");
              window.location.href = "/login";
            }}
            style={{
              alignSelf: "flex-end",
              background: "#dc2626",
              color: "white",
              border: "none",
              borderRadius: 10,
              padding: "10px 14px",
              fontWeight: 700,
              cursor: "pointer",
            }}
          >
            Logout
          </button>

          <Card>
            <input
              type="file"
              multiple
              accept=".xlsx,.xls,.bcf,.bcfzip,.zip,.docx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              onChange={(e) => setFiles(Array.from(e.target.files || []))}
            />
            <button
              onClick={generaDashboard}
              disabled={!files.length}
              style={{
                marginTop: 10,
                width: "100%",
                padding: 10,
                background: "#0f172a",
                color: "white",
                borderRadius: 10,
                border: "none",
                cursor: files.length ? "pointer" : "not-allowed",
              }}
            >
              Analizza ToDo / BCF / BCFZIP / SCHEDE WORD
            </button>

            {files.length > 0 && (
              <div style={{ marginTop: 10, color: "#64748b", fontSize: 12 }}>
                File selezionati: {files.map((f) => f.name).join(", ")}
              </div>
            )}
          </Card>
        </div>
      </div>

      {error && (
        <div
          style={{
            marginTop: 16,
            background: "#fee2e2",
            color: "#991b1b",
            border: "1px solid #fecaca",
            borderRadius: 12,
            padding: 12,
          }}
        >
          {error}
        </div>
      )}

      <div style={{ marginTop: 16 }}>
        <ImportSummary importedFiles={importedFiles} />
      </div>

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))",
          gap: 12,
          marginTop: 24,
          marginBottom: 24,
        }}
      >
        {dashboardModules
          .filter((m) => m.visible)
          .filter((m) => m.code !== "dashboard-nc-oss" && m.url !== "/dashboard-pm")
          .map((m) => (
            <Card
              key={m.code || m.title}
              active={m.active}
              onClick={() => {
                if (!m.active || !m.url) return;

                if (m.external) {
                  window.open(m.url, "_blank", "noopener,noreferrer");
                  return;
                }

                window.location.href = m.url;
              }}
            >
              <b>{m.title}</b>
              <div style={{ marginTop: 6, color: "#64748b", fontSize: 13 }}>
                {m.subtitle}
              </div>

              {m.active ? (
                <div style={{ marginTop: 12, fontSize: 12, fontWeight: 700, color: "#0284c7" }}>
                  Apri modulo →
                </div>
              ) : (
                <div style={{ marginTop: 12, fontSize: 12, fontWeight: 700, color: "#94a3b8" }}>
                  Modulo in standby
                </div>
              )}
            </Card>
          ))}
      </div>

      {rows.length === 0 && (
        <Card>
          <div style={{ fontWeight: 700 }}>Nessun dato caricato</div>
          <div style={{ marginTop: 6, color: "#64748b", fontSize: 13 }}>
            Carica uno o più file ToDo .xlsx, .bcf, .bcfzip o schede ispettive .docx
            per generare la sintesi NC/OSS.
          </div>
        </Card>
      )}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(6, 1fr)", gap: 12, marginTop: 24, marginBottom: 12 }}>
        <KPI title="Elaborati totali" value={elaboratiTot} onClick={() => setSelection({ type: "kpi", value: "totali", label: "KPI", valueLabel: "Elaborati totali" })} />
        <KPI title="Elaborati con NC" value={elaboratiNC} onClick={() => setSelection({ type: "kpi", value: "nc", label: "KPI", valueLabel: "Elaborati con NC" })} />
        <KPI title="Elaborati con OSS" value={elaboratiOSS} onClick={() => setSelection({ type: "kpi", value: "oss", label: "KPI", valueLabel: "Elaborati con OSS" })} />
        <KPI title="Elaborati senza rilievi" value={elaboratiOK} onClick={() => setSelection({ type: "kpi", value: "nessun", label: "KPI", valueLabel: "Elaborati senza rilievi" })} />
        <KPI title="Totale NC" value={totaleNC} onClick={() => setSelection({ type: "tipo", value: "NC", label: "Rilievi" })} />
        <KPI title="Totale OSS" value={totaleOSS} onClick={() => setSelection({ type: "tipo", value: "OSS", label: "Rilievi" })} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 24 }}>
        <KPI title="Totale Rilievi" value={totaleNCOSS} onClick={() => setSelection({ type: "kpi", value: "totali", label: "KPI", valueLabel: "Tutti i rilievi" })} />
        <KPI title="In attesa di riscontro dell'ispettore" value={daVerificareISP} subtitle="Ultimo commento PRG" onClick={() => setSelection({ type: "kpi", value: "da-verificare-isp", label: "KPI", valueLabel: "Da verificare ISP" })} />
        <KPI title="In attesa di risposta del progettista" value={daRisponderePRG} subtitle="Nessun PRG o ultimo ISP" onClick={() => setSelection({ type: "kpi", value: "da-rispondere-prg", label: "KPI", valueLabel: "Da rispondere PRG" })} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16 }}>
        <BarList title="Rilievi per disciplina" data={disciplineData} activeKey={selection?.type === "disciplina" ? selection.value : ""} onClick={(value: string) => setSelection({ type: "disciplina", value, label: "Disciplina" })} onExport={() => exportExcel("Rilievi_per_disciplina", toChartExportRows(disciplineData))} />
        <BarList title="Rilievi" data={esitiData} activeKey={selection?.type === "tipo" ? selection.value : ""} onClick={(value: string) => setSelection({ type: "tipo", value, label: "Rilievi" })} onExport={() => exportExcel("Rilievi", toChartExportRows(esitiData))} />
        <BarList title="NC / OSS per elaborato" data={rilieviPerElaboratoData} activeKey={selection?.type === "elaborato" ? selection.value : ""} onClick={(value: string, valueLabel: string) => setSelection({ type: "elaborato", value, label: "Elaborato", valueLabel })} onExport={() => exportExcel("NC_OSS_per_elaborato", toChartExportRows(rilieviPerElaboratoData))} />
      </div>

      <div style={{ marginTop: 24 }}>
        <DetailPanel title={selectionTitle} rows={filteredRows} onReset={() => setSelection(null)} />
      </div>

      <div
        style={{
          marginTop: 40,
          padding: 20,
          background: "#0f172a",
          color: "white",
          borderRadius: 12,
          textAlign: "center",
          fontSize: 12,
        }}
      >
        <div style={{ fontWeight: 600 }}>Nexcommon S.r.l.</div>
        <div style={{ opacity: 0.8 }}>
          © {new Date().getFullYear()} – Tutti i diritti riservati
        </div>
        <div style={{ marginTop: 6, opacity: 0.7 }}>
          Piattaforma Quality Control per ITS Controlli Tecnici S.p.A.
        </div>
      </div>
    </main>
  );
}
