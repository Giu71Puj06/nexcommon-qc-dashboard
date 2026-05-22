"use client";

import Link from "next/link";
import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

type AiReceptionResult = {
  commessa?: string;

  codiceElenco?: string;
  codiceDocumentoFile?: string;
  codiceDocumentoCartiglio?: string;

  revisioneElenco?: string;
  revisioneFile?: string;
  revisioneCartiglio?: string;

  formatoElenco?: string;
  formatoDocumento?: string;
  formatoCartiglio?: string;

  titoloElaborato?: string;
  titoloElenco?: string;

  dataRevisione?: string;
  disciplina?: string;
  faseProgettuale?: string;

  presenteInElenco?: boolean;
  presenteTraFile?: boolean;

  coerenze?: {
    presenteInElenco?: boolean;
    nomeFileCoerenteConElenco?: boolean;
    codiceFileCoerenteConCartiglio?: boolean;
    codiceElencoCoerenteConCartiglio?: boolean;
    revisioneCoerente?: boolean;
    formatoCoerente?: boolean;
    titoloPresente?: boolean;
    cartiglioLeggibile?: boolean;
    codiceCoerente?: boolean;
  };

  incoerenze?: string[];
  azioniConsigliate?: string[];
  confidenza?: number;
};

type Row = {
  fileName: string;
  result: AiReceptionResult;
};

function normalizeCode(value: string) {
  return String(value || "")
    .toUpperCase()
    .replace(/\.PDF$/i, "")
    .replace(/\s+/g, "")
    .replace(/[_\-.]/g, "")
    .trim();
}

function extractFileBaseName(fileName: string) {
  return String(fileName || "").replace(/\.pdf$/i, "").trim();
}

export default function NotaRicezioneAIPage() {
  const [elencoFile, setElencoFile] = useState<File | null>(null);
  const [files, setFiles] = useState<File[]>([]);
  const [rows, setRows] = useState<Row[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  async function analizzaPdf() {
    if (!elencoFile) {
      alert("Carica l'elenco elaborati PDF prodotto dai progettisti.");
      return;
    }

    if (!files.length) {
      alert("Carica gli elaborati PDF da verificare.");
      return;
    }

    setLoading(true);
    setError("");

    try {
      const newRows: Row[] = [];

      for (const file of files) {
        const fd = new FormData();

        fd.append("file", file);

        // Non inviamo tutto l'elenco elaborati PDF a ogni chiamata AI:
        // evitamo il superamento del context window del modello.
        fd.append(
          "elencoInfo",
          JSON.stringify({
            elencoFileName: elencoFile.name,
            fileDaVerificare: file.name,
            percorsoRelativo:
              (file as File & { webkitRelativePath?: string })
                .webkitRelativePath || file.name,
          })
        );

        fd.append("mode", "document-reception-check");

        const res = await fetch("/api/ai-document-reader", {
          method: "POST",
          body: fd,
        });

        const data = await res.json();

        if (!res.ok || !data.success) {
          throw new Error(data.error || "Errore durante l'analisi AI.");
        }

        let parsed: AiReceptionResult;

        try {
          parsed =
            typeof data.result === "string"
              ? JSON.parse(data.result)
              : data.result;
        } catch {
          parsed = {
            incoerenze: ["Risposta AI non leggibile come JSON."],
            azioniConsigliate: ["Verificare manualmente il documento."],
          };
        }

        const fileBaseName = extractFileBaseName(file.name);
        const codiceFile =
          parsed.codiceDocumentoFile || parsed.codiceElenco || fileBaseName;

        const codiceCoerente =
          parsed.coerenze?.codiceCoerente ??
          parsed.coerenze?.codiceFileCoerenteConCartiglio ??
          Boolean(
            normalizeCode(codiceFile) &&
              normalizeCode(parsed.codiceDocumentoCartiglio || "") &&
              normalizeCode(codiceFile) ===
                normalizeCode(parsed.codiceDocumentoCartiglio || "")
          );

        parsed = {
          ...parsed,
          codiceDocumentoFile: codiceFile,
          coerenze: {
            ...(parsed.coerenze || {}),
            codiceCoerente,
          },
        };

        newRows.push({
          fileName: file.name,
          result: parsed,
        });
      }

      setRows((prev) => [...prev, ...newRows]);
      setFiles([]);
    } catch (err: any) {
      setError(err?.message || "Errore imprevisto.");
    } finally {
      setLoading(false);
    }
  }

  function svuota() {
    setElencoFile(null);
    setFiles([]);
    setRows([]);
    setError("");
  }

  function esportaExcel() {
    const data = rows.map((row) => ({
      File: row.fileName,
      "Elenco elaborati PDF": elencoFile?.name || "",
      Commessa: row.result.commessa || "",
      "Codice da elenco": row.result.codiceElenco || "",
      "Codice da nome file": row.result.codiceDocumentoFile || "",
      "Codice da cartiglio": row.result.codiceDocumentoCartiglio || "",
      "Revisione elenco": row.result.revisioneElenco || "",
      "Revisione file": row.result.revisioneFile || "",
      "Revisione cartiglio": row.result.revisioneCartiglio || "",
      "Formato elenco": row.result.formatoElenco || "",
      "Formato documento": row.result.formatoDocumento || "",
      "Formato cartiglio": row.result.formatoCartiglio || "",
      "Titolo elenco": row.result.titoloElenco || "",
      "Titolo cartiglio": row.result.titoloElaborato || "",
      Disciplina: row.result.disciplina || "",
      "Fase progettuale": row.result.faseProgettuale || "",
      "Presente in elenco": row.result.coerenze?.presenteInElenco ? "OK" : "KO",
      "Nome file coerente con elenco": row.result.coerenze
        ?.nomeFileCoerenteConElenco
        ? "OK"
        : "KO",
      "Codice file/cartiglio coerente": row.result.coerenze
        ?.codiceFileCoerenteConCartiglio
        ? "OK"
        : "KO",
      "Codice elenco/cartiglio coerente": row.result.coerenze
        ?.codiceElencoCoerenteConCartiglio
        ? "OK"
        : "KO",
      "Revisione coerente": row.result.coerenze?.revisioneCoerente
        ? "OK"
        : "KO",
      "Formato coerente": row.result.coerenze?.formatoCoerente ? "OK" : "KO",
      "Cartiglio leggibile": row.result.coerenze?.cartiglioLeggibile
        ? "OK"
        : "KO",
      Incoerenze: (row.result.incoerenze || []).join(" | "),
      "Azioni consigliate": (row.result.azioniConsigliate || []).join(" | "),
      Confidenza: row.result.confidenza ?? "",
    }));

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(data);

    XLSX.utils.book_append_sheet(workbook, worksheet, "Nota Ricezione AI");
    XLSX.writeFile(workbook, "Nota_Ricezione_Elaborati_AI.xlsx");
  }

  const kpi = useMemo(() => {
    const totale = rows.length;
    const presentiInElenco = rows.filter(
      (r) => r.result.coerenze?.presenteInElenco
    ).length;
    const nomiFileCoerenti = rows.filter(
      (r) => r.result.coerenze?.nomeFileCoerenteConElenco
    ).length;
    const codiciCartiglioCoerenti = rows.filter(
      (r) =>
        r.result.coerenze?.codiceFileCoerenteConCartiglio ||
        r.result.coerenze?.codiceCoerente
    ).length;
    const revisioniCoerenti = rows.filter(
      (r) => r.result.coerenze?.revisioneCoerente
    ).length;
    const formatiCoerenti = rows.filter(
      (r) => r.result.coerenze?.formatoCoerente
    ).length;
    const cartigliLeggibili = rows.filter(
      (r) => r.result.coerenze?.cartiglioLeggibile
    ).length;
    const conIncoerenze = rows.filter(
      (r) => (r.result.incoerenze || []).length > 0
    ).length;

    return {
      totale,
      presentiInElenco,
      nomiFileCoerenti,
      codiciCartiglioCoerenti,
      revisioniCoerenti,
      formatiCoerenti,
      cartigliLeggibili,
      conIncoerenze,
    };
  }, [rows]);

  return (
    <main style={pageStyle}>
      <Link href="/dashboard-ai" style={backLinkStyle}>
        ← Torna alla Dashboard AI
      </Link>

      <h1 style={titleStyle}>Nota Ricezione Elaborati AI</h1>

      <p style={leadStyle}>
        Carica il PDF dell'elenco elaborati prodotto dai progettisti e i PDF
        degli elaborati ricevuti. L'AI confronta codice in elenco, nome file,
        codice cartiglio, revisione e formato del documento, generando un report
        di coerenze e anomalie.
      </p>

      <section style={cardStyle}>
        <h2>Verifica ricezione elaborati AI</h2>

        <div style={uploadGridStyle}>
          <div>
            <label style={labelStyle}>Elenco Elaborati PDF</label>
            <input
              type="file"
              accept=".pdf"
              onChange={(e) => setElencoFile(e.target.files?.[0] || null)}
              style={inputStyle}
            />

            {elencoFile && (
              <div style={infoStyle}>Elenco caricato: {elencoFile.name}</div>
            )}
          </div>

          <div>
            <label style={labelStyle}>Elaborati PDF da verificare</label>
            <input
  type="file"
  accept=".pdf"
  multiple
  {...({
    webkitdirectory: "",
    directory: "",
  } as any)}
  onChange={(e) => {
    const selectedFiles = Array.from(e.target.files || []).filter(
      (file) => file.name.toLowerCase().endsWith(".pdf")
    );

    setFiles(selectedFiles);
  }}
  style={inputStyle}
/>
            {files.length > 0 && (
              <div style={infoStyle}>PDF selezionati: {files.length}</div>
            )}
          </div>
        </div>

        <div style={noteStyle}>
          Il confronto atteso è: codice elenco elaborati = nome file PDF =
          codice cartiglio. Inoltre vengono verificati revisione, formato
          documento, disciplina, titolo e leggibilità del cartiglio.
        </div>

        <div style={actionsStyle}>
          <button onClick={analizzaPdf} disabled={loading} style={primaryButton}>
            {loading ? "Analisi AI in corso..." : "Analizza ricezione con AI"}
          </button>

          <button onClick={svuota} disabled={loading} style={secondaryButton}>
            Svuota dati
          </button>

          <button
            onClick={esportaExcel}
            disabled={!rows.length || loading}
            style={{
              ...exportButton,
              opacity: !rows.length || loading ? 0.5 : 1,
            }}
          >
            Esporta Excel
          </button>
        </div>

        {error && <div style={errorStyle}>{error}</div>}
      </section>

      <section style={gridStyle}>
        <Kpi title="Elaborati analizzati" value={String(kpi.totale)} />
        <Kpi title="Presenti in elenco" value={String(kpi.presentiInElenco)} />
        <Kpi title="Nomi file coerenti" value={String(kpi.nomiFileCoerenti)} />
        <Kpi
          title="Codici cartiglio coerenti"
          value={String(kpi.codiciCartiglioCoerenti)}
        />
        <Kpi title="Revisioni coerenti" value={String(kpi.revisioniCoerenti)} />
        <Kpi title="Formati coerenti" value={String(kpi.formatiCoerenti)} />
        <Kpi title="Cartigli leggibili" value={String(kpi.cartigliLeggibili)} />
        <Kpi title="Con incoerenze" value={String(kpi.conIncoerenze)} />
      </section>

      {rows.length > 0 && (
        <section style={cardStyle}>
          <h2>Report coerenza ricezione elaborati</h2>

          <div style={{ overflowX: "auto" }}>
            <table style={tableStyle}>
              <thead>
                <tr>
                  <th style={th}>File</th>
                  <th style={th}>Commessa</th>
                  <th style={th}>Codice elenco</th>
                  <th style={th}>Codice file</th>
                  <th style={th}>Codice cartiglio</th>
                  <th style={th}>Rev. elenco</th>
                  <th style={th}>Rev. file</th>
                  <th style={th}>Rev. cartiglio</th>
                  <th style={th}>Formato elenco</th>
                  <th style={th}>Formato doc.</th>
                  <th style={th}>Formato cartiglio</th>
                  <th style={th}>Titolo elenco</th>
                  <th style={th}>Titolo cartiglio</th>
                  <th style={th}>Disciplina</th>
                  <th style={th}>Presente elenco</th>
                  <th style={th}>Nome file</th>
                  <th style={th}>Codice cartiglio</th>
                  <th style={th}>Revisione</th>
                  <th style={th}>Formato</th>
                  <th style={th}>Cartiglio</th>
                  <th style={th}>Incoerenze</th>
                  <th style={th}>Azioni consigliate</th>
                  <th style={th}>Confidenza</th>
                </tr>
              </thead>

              <tbody>
                {rows.map((row, index) => (
                  <tr key={`${row.fileName}-${index}`}>
                    <td style={td}>{row.fileName}</td>
                    <td style={td}>{row.result.commessa || "-"}</td>
                    <td style={td}>{row.result.codiceElenco || "-"}</td>
                    <td style={td}>{row.result.codiceDocumentoFile || "-"}</td>
                    <td style={td}>
                      {row.result.codiceDocumentoCartiglio || "-"}
                    </td>
                    <td style={td}>{row.result.revisioneElenco || "-"}</td>
                    <td style={td}>{row.result.revisioneFile || "-"}</td>
                    <td style={td}>{row.result.revisioneCartiglio || "-"}</td>
                    <td style={td}>{row.result.formatoElenco || "-"}</td>
                    <td style={td}>{row.result.formatoDocumento || "-"}</td>
                    <td style={td}>{row.result.formatoCartiglio || "-"}</td>
                    <td style={td}>{row.result.titoloElenco || "-"}</td>
                    <td style={td}>{row.result.titoloElaborato || "-"}</td>
                    <td style={td}>{row.result.disciplina || "-"}</td>
                    <td style={td}>
                      {row.result.coerenze?.presenteInElenco ? "✅" : "❌"}
                    </td>
                    <td style={td}>
                      {row.result.coerenze?.nomeFileCoerenteConElenco
                        ? "✅"
                        : "❌"}
                    </td>
                    <td style={td}>
                      {row.result.coerenze?.codiceFileCoerenteConCartiglio ||
                      row.result.coerenze?.codiceCoerente
                        ? "✅"
                        : "❌"}
                    </td>
                    <td style={td}>
                      {row.result.coerenze?.revisioneCoerente ? "✅" : "❌"}
                    </td>
                    <td style={td}>
                      {row.result.coerenze?.formatoCoerente ? "✅" : "❌"}
                    </td>
                    <td style={td}>
                      {row.result.coerenze?.cartiglioLeggibile ? "✅" : "❌"}
                    </td>
                    <td style={td}>
                      {(row.result.incoerenze || []).join(" | ") || "-"}
                    </td>
                    <td style={td}>
                      {(row.result.azioniConsigliate || []).join(" | ") || "-"}
                    </td>
                    <td style={td}>{row.result.confidenza ?? "-"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}
    </main>
  );
}

function Kpi({ title, value }: { title: string; value: string }) {
  return (
    <div style={cardStyle}>
      <div style={{ color: "#64748b", fontSize: 14 }}>{title}</div>
      <div style={{ fontSize: 30, fontWeight: 800, marginTop: 8 }}>
        {value}
      </div>
    </div>
  );
}

const pageStyle: React.CSSProperties = {
  minHeight: "100vh",
  padding: 30,
  background: "#f1f5f9",
  fontFamily: "Arial, sans-serif",
  color: "#0f172a",
};

const backLinkStyle: React.CSSProperties = {
  display: "inline-block",
  marginBottom: 24,
  color: "#0284c7",
  fontWeight: 700,
  textDecoration: "none",
};

const titleStyle: React.CSSProperties = {
  fontSize: 42,
  marginBottom: 10,
};

const leadStyle: React.CSSProperties = {
  fontSize: 18,
  color: "#475569",
  maxWidth: 1200,
};

const cardStyle: React.CSSProperties = {
  background: "white",
  border: "1px solid #cbd5e1",
  borderRadius: 16,
  padding: 20,
  marginTop: 24,
  boxShadow: "0 8px 20px rgba(15,23,42,0.06)",
};

const uploadGridStyle: React.CSSProperties = {
  display: "grid",
  gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))",
  gap: 18,
};

const labelStyle: React.CSSProperties = {
  display: "block",
  fontWeight: 800,
  marginBottom: 8,
};

const inputStyle: React.CSSProperties = {
  width: "100%",
  padding: 12,
  border: "1px solid #cbd5e1",
  borderRadius: 10,
  marginBottom: 14,
};

const actionsStyle: React.CSSProperties = {
  display: "flex",
  gap: 12,
  marginTop: 18,
  flexWrap: "wrap",
};

const primaryButton: React.CSSProperties = {
  padding: "12px 18px",
  background: "#0f172a",
  color: "white",
  border: "none",
  borderRadius: 10,
  fontWeight: 700,
  cursor: "pointer",
};

const secondaryButton: React.CSSProperties = {
  padding: "12px 18px",
  background: "#94a3b8",
  color: "white",
  border: "none",
  borderRadius: 10,
  fontWeight: 700,
  cursor: "pointer",
};

const exportButton: React.CSSProperties = {
  padding: "12px 18px",
  background: "#16a34a",
  color: "white",
  border: "none",
  borderRadius: 10,
  fontWeight: 700,
  cursor: "pointer",
};

const infoStyle: React.CSSProperties = {
  marginTop: 6,
  padding: 12,
  borderRadius: 10,
  background: "#eff6ff",
  color: "#1d4ed8",
  fontWeight: 700,
};

const noteStyle: React.CSSProperties = {
  marginTop: 10,
  padding: 12,
  borderRadius: 10,
  background: "#f8fafc",
  color: "#475569",
  lineHeight: 1.5,
};

const errorStyle: React.CSSProperties = {
  marginTop: 16,
  padding: 14,
  borderRadius: 10,
  background: "#fee2e2",
  color: "#991b1b",
};

const gridStyle: React.CSSProperties = {
  display: "grid",
  gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
  gap: 16,
};

const tableStyle: React.CSSProperties = {
  width: "100%",
  borderCollapse: "collapse",
  fontSize: 14,
};

const th: React.CSSProperties = {
  textAlign: "left",
  borderBottom: "1px solid #cbd5e1",
  padding: 10,
  background: "#f8fafc",
  whiteSpace: "nowrap",
};

const td: React.CSSProperties = {
  borderBottom: "1px solid #e2e8f0",
  padding: 10,
  verticalAlign: "top",
};
