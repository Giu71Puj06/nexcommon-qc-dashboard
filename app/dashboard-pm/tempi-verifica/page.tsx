"use client";

import React, { useMemo, useState } from "react";
import Link from "next/link";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import PizZip from "pizzip";

type CommessaRow = {
  commessa: string;
  file: string;
  fonte: string;
  primoEvento: string;
  ultimoEvento: string;
  durataGiorni: number;
  nc: number;
  oss: number;
  bcfIssue: number;
  totaleRilievi: number;
  giorniPerRilievo: number;
};

type ParsedEvent = {
  commessa: string;
  file: string;
  fonte: string;
  date: string;
  nc: number;
  oss: number;
  bcfIssue: number;
};

export default function TempiVerificaPage() {
  const [rows, setRows] = useState<CommessaRow[]>([]);
  const [loading, setLoading] = useState(false);

  const summary = useMemo(() => {
    const commesse = rows.length;
    const durataMedia = commesse > 0 ? Math.round(rows.reduce((sum, r) => sum + r.durataGiorni, 0) / commesse) : 0;
    const rilievi = rows.reduce((sum, r) => sum + r.totaleRilievi, 0);
    const giorniPerRilievo = rilievi > 0 ? Number((rows.reduce((sum, r) => sum + r.durataGiorni, 0) / rilievi).toFixed(2)) : 0;
    return { commesse, durataMedia, rilievi, giorniPerRilievo };
  }, [rows]);

  async function handleFiles(event: React.ChangeEvent<HTMLInputElement>) {
    const files = Array.from(event.target.files || []);
    if (files.length === 0) return;
    setLoading(true);

    try {
      const events: ParsedEvent[] = [];

      for (const file of files) {
        const name = file.name.toLowerCase();
        if (name.endsWith(".bcfzip") || name.endsWith(".bcf") || name.endsWith(".zip")) {
          events.push(...(await parseBcfFile(file)));
        } else if (name.endsWith(".docx")) {
          events.push(await parseInspectionDocx(file));
        }
      }

      setRows(buildRows(events));
    } catch (err) {
      console.error(err);
      alert("Errore durante l'elaborazione dei file.");
    } finally {
      setLoading(false);
      event.target.value = "";
    }
  }

  function exportXlsx() {
    if (rows.length === 0) {
      alert("Nessun dato da esportare.");
      return;
    }

    const data = rows.map((row) => ({
      Commessa: row.commessa,
      File: row.file,
      Fonte: row.fonte,
      "Primo evento": formatDate(row.primoEvento),
      "Ultimo evento": formatDate(row.ultimoEvento),
      "Durata verifica giorni": row.durataGiorni,
      NC: row.nc,
      OSS: row.oss,
      "Issue BCF": row.bcfIssue,
      "Totale rilievi": row.totaleRilievi,
      "Giorni per rilievo": row.giorniPerRilievo,
    }));

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [
      { wch: 40 }, { wch: 50 }, { wch: 18 }, { wch: 16 }, { wch: 16 },
      { wch: 22 }, { wch: 8 }, { wch: 8 }, { wch: 10 }, { wch: 14 }, { wch: 18 },
    ];
    XLSX.utils.book_append_sheet(wb, ws, "Tempi verifica");
    XLSX.writeFile(wb, "tempi_medi_verifica_commesse.xlsx");
  }

  return (
    <main style={pageStyle}>
      <div style={contentStyle}>
        <Link href="/dashboard-pm" style={backLinkStyle}>← Torna alla Dashboard PM</Link>

        <h1 style={titleStyle}>Stima tempi medi di verifica commesse</h1>

        <p style={leadStyle}>
          Carica file <b>BCF/BCFZIP</b> e schede ispettive <b>Word DOCX</b>. Il modulo stima la durata media
          delle verifiche, il numero di rilievi e i giorni medi per rilievo/commessa.
        </p>

        <section style={cardStyle}>
          <h2 style={sectionTitleStyle}>1. Carica file</h2>
          <input type="file" multiple accept=".bcf,.bcfzip,.zip,.docx" onChange={handleFiles} style={inputStyle} />
          <div style={helpStyle}>Puoi caricare insieme file BCF/BCFZIP e schede ispettive Word. I dati vengono aggregati per nome commessa.</div>
          <div style={actionsStyle}>
            <button type="button" onClick={exportXlsx} disabled={rows.length === 0} style={buttonStyle}>Esporta XLSX</button>
            <button type="button" onClick={() => setRows([])} disabled={rows.length === 0} style={secondaryButtonStyle}>Svuota dati</button>
          </div>
        </section>

        {loading && <p>Elaborazione in corso...</p>}

        {rows.length > 0 && (
          <>
            <section style={cardStyle}>
              <h2 style={sectionTitleStyle}>KPI generali</h2>
              <div style={statsGridStyle}>
                <Stat label="Commesse analizzate" value={summary.commesse} />
                <Stat label="Durata media verifica" value={`${summary.durataMedia} gg`} />
                <Stat label="Rilievi totali" value={summary.rilievi} />
                <Stat label="Giorni medi per rilievo" value={summary.giorniPerRilievo} />
              </div>
            </section>

            <section style={cardStyle}>
              <h2 style={sectionTitleStyle}>Report commesse</h2>
              <div style={tableWrapStyle}>
                <table style={tableStyle}>
                  <thead>
                    <tr>
                      <th style={thStyle}>Commessa</th>
                      <th style={thStyle}>Fonte</th>
                      <th style={thStyle}>Primo evento</th>
                      <th style={thStyle}>Ultimo evento</th>
                      <th style={thStyle}>Durata</th>
                      <th style={thStyle}>NC</th>
                      <th style={thStyle}>OSS</th>
                      <th style={thStyle}>Issue BCF</th>
                      <th style={thStyle}>Totale rilievi</th>
                      <th style={thStyle}>Giorni/rilievo</th>
                      <th style={thStyle}>File</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((row, index) => (
                      <tr key={`${row.commessa}-${index}`}>
                        <td style={tdStyle}><b>{row.commessa}</b></td>
                        <td style={tdStyle}>{row.fonte}</td>
                        <td style={tdStyle}>{formatDate(row.primoEvento)}</td>
                        <td style={tdStyle}>{formatDate(row.ultimoEvento)}</td>
                        <td style={tdStyle}>{row.durataGiorni} gg</td>
                        <td style={tdStyle}>{row.nc}</td>
                        <td style={tdStyle}>{row.oss}</td>
                        <td style={tdStyle}>{row.bcfIssue}</td>
                        <td style={tdStyle}>{row.totaleRilievi}</td>
                        <td style={tdStyle}>{row.giorniPerRilievo}</td>
                        <td style={tdStyle}>{row.file}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          </>
        )}
      </div>
    </main>
  );
}

async function parseBcfFile(file: File): Promise<ParsedEvent[]> {
  const events: ParsedEvent[] = [];
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const commessa = cleanProjectName(file.name);

  for (const entry of Object.values(zip.files)) {
    if (entry.dir) continue;
    if (!entry.name.toLowerCase().endsWith(".bcf")) continue;
    const xml = await entry.async("string");
    const dates = extractDates(xml);
    if (dates.length === 0) continue;

    events.push({ commessa, file: file.name, fonte: "BCF", date: dates[0], nc: 0, oss: 0, bcfIssue: 1 });

    if (dates.length > 1) {
      events.push({ commessa, file: file.name, fonte: "BCF", date: dates[dates.length - 1], nc: 0, oss: 0, bcfIssue: 0 });
    }
  }

  return events;
}

async function parseInspectionDocx(file: File): Promise<ParsedEvent> {
  const buffer = await file.arrayBuffer();
  const zip = new PizZip(buffer);
  const xml = zip.file("word/document.xml")?.asText() || "";

  const text = xml.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
  const nc = uniqueMatches(text, /\bNC\s*\d+/gi).length;
  const oss = uniqueMatches(text, /\bOSS\s*\d+/gi).length;
  const dates = extractItalianDates(text);

  return { commessa: cleanProjectName(file.name), file: file.name, fonte: "WORD", date: dates[0] || "", nc, oss, bcfIssue: 0 };
}

function buildRows(events: ParsedEvent[]): CommessaRow[] {
  const grouped = new Map<string, ParsedEvent[]>();
  for (const event of events) {
    if (!event.commessa) continue;
    grouped.set(event.commessa, [...(grouped.get(event.commessa) || []), event]);
  }

  return Array.from(grouped.entries()).map(([commessa, items]) => {
    const allDates = items.map((item) => item.date).filter(Boolean).sort((a, b) => new Date(a).getTime() - new Date(b).getTime());
    const primoEvento = allDates[0] || "";
    const ultimoEvento = allDates[allDates.length - 1] || "";
    const nc = items.reduce((sum, item) => sum + item.nc, 0);
    const oss = items.reduce((sum, item) => sum + item.oss, 0);
    const bcfIssue = items.reduce((sum, item) => sum + item.bcfIssue, 0);
    const totaleRilievi = nc + oss + bcfIssue;
    const durataGiorni = daysBetween(primoEvento, ultimoEvento);
    const giorniPerRilievo = totaleRilievi > 0 ? Number((durataGiorni / totaleRilievi).toFixed(2)) : 0;

    return {
      commessa,
      file: Array.from(new Set(items.map((item) => item.file))).join(", "),
      fonte: Array.from(new Set(items.map((item) => item.fonte))).join(" + "),
      primoEvento,
      ultimoEvento,
      durataGiorni,
      nc,
      oss,
      bcfIssue,
      totaleRilievi,
      giorniPerRilievo,
    };
  });
}

function extractDates(xml: string): string[] {
  const matches = xml.match(/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z?/g) || [];
  return matches.map((value) => value.slice(0, 10));
}

function extractItalianDates(text: string): string[] {
  const matches = text.match(/\b\d{2}\/\d{2}\/\d{4}\b/g) || [];
  return matches.map((date) => {
    const [day, month, year] = date.split("/");
    return `${year}-${month}-${day}`;
  }).sort((a, b) => new Date(a).getTime() - new Date(b).getTime());
}

function uniqueMatches(text: string, regex: RegExp) {
  return Array.from(new Set((text.match(regex) || []).map((x) => x.replace(/\s+/g, "").toUpperCase())));
}

function cleanProjectName(fileName: string): string {
  return fileName
    .replace(/\.bcfzip$/i, "")
    .replace(/\.bcf$/i, "")
    .replace(/\.zip$/i, "")
    .replace(/\.docx$/i, "")
    .replace(/\d{12,}/g, "")
    .replace(/[_-]+$/g, "")
    .trim();
}

function daysBetween(start: string, end: string): number {
  if (!start || !end) return 0;
  const startDate = new Date(start);
  const endDate = new Date(end);
  if (Number.isNaN(startDate.getTime()) || Number.isNaN(endDate.getTime())) return 0;
  return Math.max(0, Math.ceil((endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24)));
}

function formatDate(value: string): string {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "-";
  return date.toLocaleDateString("it-IT");
}

function Stat({ label, value }: { label: string; value: string | number }) {
  return (
    <div style={statStyle}>
      <div style={{ fontSize: 13, color: "#64748b", marginBottom: 6 }}>{label}</div>
      <div style={{ fontSize: 28, fontWeight: 800 }}>{value}</div>
    </div>
  );
}

const pageStyle: React.CSSProperties = { minHeight: "100vh", padding: 30, fontFamily: "Arial, sans-serif", background: "#f1f5f9", color: "#0f172a" };
const contentStyle: React.CSSProperties = { width: "100%", maxWidth: 1500, margin: "0 auto" };
const backLinkStyle: React.CSSProperties = { display: "inline-block", marginBottom: 24, color: "#0284c7", fontWeight: 700, textDecoration: "none" };
const titleStyle: React.CSSProperties = { fontSize: 42, marginBottom: 16 };
const leadStyle: React.CSSProperties = { fontSize: 19, lineHeight: 1.55, marginBottom: 24, maxWidth: 980 };
const cardStyle: React.CSSProperties = { background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: 16, padding: 22, marginBottom: 20, boxShadow: "0 8px 24px rgba(15, 23, 42, 0.06)" };
const sectionTitleStyle: React.CSSProperties = { fontSize: 22, marginTop: 0, marginBottom: 16 };
const inputStyle: React.CSSProperties = { display: "block", width: "100%", marginTop: 8, padding: 12, border: "1px solid #cbd5e1", borderRadius: 10, background: "#fff" };
const helpStyle: React.CSSProperties = { marginTop: 8, fontSize: 13, color: "#64748b" };
const actionsStyle: React.CSSProperties = { display: "flex", gap: 12, marginTop: 18 };
const buttonStyle: React.CSSProperties = { padding: "12px 18px", border: 0, borderRadius: 10, background: "#0284c7", color: "white", fontWeight: 800, cursor: "pointer" };
const secondaryButtonStyle: React.CSSProperties = { padding: "12px 18px", border: 0, borderRadius: 10, background: "#64748b", color: "white", fontWeight: 800, cursor: "pointer" };
const statsGridStyle: React.CSSProperties = { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))", gap: 14 };
const statStyle: React.CSSProperties = { border: "1px solid #e2e8f0", borderRadius: 12, padding: 16, background: "#f8fafc" };
const tableWrapStyle: React.CSSProperties = { overflowX: "auto", border: "1px solid #e2e8f0", borderRadius: 12 };
const tableStyle: React.CSSProperties = { width: "100%", minWidth: 1200, borderCollapse: "collapse", fontSize: 13 };
const thStyle: React.CSSProperties = { textAlign: "left", padding: 10, borderBottom: "1px solid #e2e8f0", background: "#f8fafc", whiteSpace: "nowrap" };
const tdStyle: React.CSSProperties = { padding: 10, borderBottom: "1px solid #e2e8f0", verticalAlign: "top" };
