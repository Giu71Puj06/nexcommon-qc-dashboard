"use client";

import Link from "next/link";
import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

type AiReceptionResult = {
  commessa?: string;
  codiceDocumentoFile?: string;
  codiceDocumentoCartiglio?: string;
  titoloElaborato?: string;
  revisioneFile?: string;
  revisioneCartiglio?: string;
  dataRevisione?: string;
  disciplina?: string;
  faseProgettuale?: string;
  coerenze?: {
    codiceCoerente?: boolean;
    revisioneCoerente?: boolean;
    titoloPresente?: boolean;
    cartiglioLeggibile?: boolean;
  };
  incoerenze?: string[];
  azioniConsigliate?: string[];
  confidenza?: number;
};

type Row = {
  fileName: string;
  result: AiReceptionResult;
};

export default function NotaRicezioneAIPage() {
  const [files, setFiles] = useState<File[]>([]);
  const [rows, setRows] = useState<Row[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  async function analizzaPdf() {
    if (!files.length) {
      alert("Carica almeno un PDF.");
      return;
    }

    setLoading(true);
    setError("");

    try {
      const newRows: Row[] = [];

      for (const file of files) {
        const fd = new FormData();
        fd.append("file", file);
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
    setFiles([]);
    setRows([]);
    setError("");
  }

  function esportaExcel() {
    const data = rows.map((row) => ({
      File: row.fileName,
      Commessa: row.result.commessa || "",
      "Codice documento file": row.result.codiceDocumentoFile || "",
      "Codice documento cartiglio": row.result.codiceDocumentoCartiglio || "",
      "Titolo elaborato": row.result.titoloElaborato || "",
      "Revisione file": row.result.revisioneFile || "",
      "Revisione cartiglio": row.result.revisioneCartiglio || "",
      "Data revisione": row.result.dataRevisione || "",
      Disciplina: row.result.disciplina || "",
      "Fase progettuale": row.result.faseProgettuale || "",
      "Codice coerente": row.result.coerenze?.codiceCoerente ? "OK" : "KO",
      "Revisione coerente": row.result.coerenze?.revisioneCoerente
        ? "OK"
        : "KO",
      "Titolo presente": row.result.coerenze?.titoloPresente ? "OK" : "KO",
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

    const okCodice = rows.filter(
      (r) => r.result.coerenze?.codiceCoerente
    ).length;

    const okRevisione = rows.filter(
      (r) => r.result.coerenze?.revisioneCoerente
    ).length;

    const cartigliLeggibili = rows.filter(
      (r) => r.result.coerenze?.cartiglioLeggibile
    ).length;

    const conIncoerenze = rows.filter(
      (r) => (r.result.incoerenze || []).length > 0
    ).length;

    return {
      totale,
      okCodice,
      okRevisione,
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
        Carica gli elaborati PDF ricevuti. L'AI legge cartigli, codici,
        revisioni, titoli, discipline e controlla la coerenza tra nome file e
        contenuto del documento.
      </p>

      <section style={cardStyle}>
        <h2>Caricamento elaborati PDF</h2>

        <input
          type="file"
          accept=".pdf"
          multiple
          onChange={(e) => setFiles(Array.from(e.target.files || []))}
          style={inputStyle}
        />

        <div style={actionsStyle}>
          <button onClick={analizzaPdf} disabled={loading} style={primaryButton}>
            {loading ? "Analisi AI in corso..." : "Analizza elaborati con AI"}
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

        {files.length > 0 && (
          <div style={infoStyle}>PDF selezionati: {files.length}</div>
        )}

        {error && <div style={errorStyle}>{error}</div>}
      </section>

      <section style={gridStyle}>
        <Kpi title="Elaborati analizzati" value={String(kpi.totale)} />
        <Kpi title="Codici coerenti" value={String(kpi.okCodice)} />
        <Kpi title="Revisioni coerenti" value={String(kpi.okRevisione)} />
        <Kpi title="Cartigli leggibili" value={String(kpi.cartigliLeggibili)} />
        <Kpi title="Con incoerenze" value={String(kpi.conIncoerenze)} />
      </section>

      {rows.length > 0 && (
        <section style={cardStyle}>
          <h2>Report coerenza elaborati</h2>

          <div style={{ overflowX: "auto" }}>
            <table style={tableStyle}>
              <thead>
                <tr>
                  <th style={th}>File</th>
                  <th style={th}>Commessa</th>
                  <th style={th}>Codice file</th>
                  <th style={th}>Codice cartiglio</th>
                  <th style={th}>Titolo elaborato</th>
                  <th style={th}>Rev. file</th>
                  <th style={th}>Rev. cartiglio</th>
                  <th style={th}>Disciplina</th>
                  <th style={th}>Fase</th>
                  <th style={th}>Codice</th>
                  <th style={th}>Revisione</th>
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
                    <td style={td}>
                      {row.result.codiceDocumentoFile || "-"}
                    </td>
                    <td style={td}>
                      {row.result.codiceDocumentoCartiglio || "-"}
                    </td>
                    <td style={td}>{row.result.titoloElaborato || "-"}</td>
                    <td style={td}>{row.result.revisioneFile || "-"}</td>
                    <td style={td}>{row.result.revisioneCartiglio || "-"}</td>
                    <td style={td}>{row.result.disciplina || "-"}</td>
                    <td style={td}>{row.result.faseProgettuale || "-"}</td>
                    <td style={td}>
                      {row.result.coerenze?.codiceCoerente ? "✅" : "❌"}
                    </td>
                    <td style={td}>
                      {row.result.coerenze?.revisioneCoerente ? "✅" : "❌"}
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
  marginTop: 10,
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
  marginTop: 14,
  padding: 12,
  borderRadius: 10,
  background: "#eff6ff",
  color: "#1d4ed8",
  fontWeight: 700,
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
};

const td: React.CSSProperties = {
  borderBottom: "1px solid #e2e8f0",
  padding: 10,
  verticalAlign: "top",
};
