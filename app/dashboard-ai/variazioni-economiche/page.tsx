"use client";

import Link from "next/link";
import { useMemo, useState } from "react";

type AiEconomicResult = {
  commessa?: string;
  codiceElaborato?: string;
  titoloElaborato?: string;
  revisione?: string;
  tipoDocumento?: string;
  importoTotale?: number;
  computoEstimativo?: boolean;
  computoSenzaPrezzi?: boolean;
  coerenzaNomeFileCartiglio?: boolean;
  warning?: string;
  confidenza?: number;
};

type Row = {
  fileName: string;
  fase: "iniziale" | "finale";
  result: AiEconomicResult;
};

export default function VariazioniEconomicheAIPage() {
  const [fase, setFase] = useState<"iniziale" | "finale">("iniziale");
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
        fd.append("mode", "economic-analysis");

        const res = await fetch("/api/ai-document-reader", {
          method: "POST",
          body: fd,
        });

        const data = await res.json();

        if (!res.ok || !data.success) {
          throw new Error(data.error || "Errore durante l'analisi AI.");
        }

        let parsed: AiEconomicResult;

        try {
          parsed =
            typeof data.result === "string"
              ? JSON.parse(data.result)
              : data.result;
        } catch {
          parsed = {
            warning: "Risposta AI non leggibile come JSON.",
          };
        }

        newRows.push({
          fileName: file.name,
          fase,
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

  const kpi = useMemo(() => {
    const iniziali = rows.filter((r) => r.fase === "iniziale");
    const finali = rows.filter((r) => r.fase === "finale");

    const importoIniziale = iniziali.reduce(
      (sum, r) => sum + Number(r.result.importoTotale || 0),
      0
    );

    const importoFinale = finali.reduce(
      (sum, r) => sum + Number(r.result.importoTotale || 0),
      0
    );

    const delta = importoFinale - importoIniziale;
    const deltaPercent =
      importoIniziale > 0 ? (delta / importoIniziale) * 100 : 0;

    return {
      documenti: rows.length,
      importoIniziale,
      importoFinale,
      delta,
      deltaPercent,
    };
  }, [rows]);

  const euro = (value: number) =>
    new Intl.NumberFormat("it-IT", {
      style: "currency",
      currency: "EUR",
    }).format(value || 0);

  return (
    <main style={pageStyle}>
      <Link href="/dashboard-ai" style={backLinkStyle}>
        ← Torna alla Dashboard AI
      </Link>

      <h1 style={titleStyle}>Variazioni Economiche AI</h1>

      <p style={leadStyle}>
        Carica computi metrici, quadri economici o documenti economici. L'AI
        legge cartiglio, titolo elaborato, tipo documento e importo totale.
      </p>

      <section style={cardStyle}>
        <h2>Caricamento PDF economici</h2>

        <label style={labelStyle}>Fase documento</label>
        <select
          value={fase}
          onChange={(e) => setFase(e.target.value as "iniziale" | "finale")}
          style={inputStyle}
        >
          <option value="iniziale">Fase iniziale / arrivo progetto in ITS</option>
          <option value="finale">Fase finale / conclusione verifica</option>
        </select>

        <input
          type="file"
          accept=".pdf"
          multiple
          onChange={(e) => setFiles(Array.from(e.target.files || []))}
          style={inputStyle}
        />

        <div style={actionsStyle}>
          <button onClick={analizzaPdf} disabled={loading} style={primaryButton}>
            {loading ? "Analisi AI in corso..." : "Analizza con AI"}
          </button>

          <button onClick={svuota} disabled={loading} style={secondaryButton}>
            Svuota dati
          </button>
        </div>

        {error && <div style={errorStyle}>{error}</div>}
      </section>

      <section style={gridStyle}>
        <Kpi title="Documenti analizzati" value={String(kpi.documenti)} />
        <Kpi title="Importo iniziale" value={euro(kpi.importoIniziale)} />
        <Kpi title="Importo finale" value={euro(kpi.importoFinale)} />
        <Kpi title="Delta economico" value={euro(kpi.delta)} />
        <Kpi title="Delta %" value={`${kpi.deltaPercent.toFixed(1)}%`} />
      </section>

      {rows.length > 0 && (
        <section style={cardStyle}>
          <h2>Dati estratti dall'AI</h2>

          <div style={{ overflowX: "auto" }}>
            <table style={tableStyle}>
              <thead>
                <tr>
                  <th style={th}>File</th>
                  <th style={th}>Fase</th>
                  <th style={th}>Commessa</th>
                  <th style={th}>Codice elaborato</th>
                  <th style={th}>Titolo</th>
                  <th style={th}>Tipo documento</th>
                  <th style={th}>Importo totale</th>
                  <th style={th}>Confidenza</th>
                  <th style={th}>Warning</th>
                </tr>
              </thead>

              <tbody>
                {rows.map((row, index) => (
                  <tr key={`${row.fileName}-${index}`}>
                    <td style={td}>{row.fileName}</td>
                    <td style={td}>{row.fase}</td>
                    <td style={td}>{row.result.commessa || "-"}</td>
                    <td style={td}>{row.result.codiceElaborato || "-"}</td>
                    <td style={td}>{row.result.titoloElaborato || "-"}</td>
                    <td style={td}>{row.result.tipoDocumento || "-"}</td>
                    <td style={td}>
                      <b>{euro(Number(row.result.importoTotale || 0))}</b>
                    </td>
                    <td style={td}>{row.result.confidenza ?? "-"}</td>
                    <td style={td}>{row.result.warning || "-"}</td>
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
  maxWidth: 1100,
};

const cardStyle: React.CSSProperties = {
  background: "white",
  border: "1px solid #cbd5e1",
  borderRadius: 16,
  padding: 20,
  marginTop: 24,
  boxShadow: "0 8px 20px rgba(15,23,42,0.06)",
};

const labelStyle: React.CSSProperties = {
  display: "block",
  fontWeight: 700,
  marginTop: 16,
  marginBottom: 6,
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
