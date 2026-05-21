"use client";

import React, { useMemo, useState } from "react";
import Link from "next/link";

type EconomicFile = {
  id: string;
  commessa: string;
  fase: "iniziale" | "finale";
  fileName: string;
  importo: number;
  warning?: string;
};

type EconomicKpi = {
  commessa: string;
  importoIniziale: number;
  importoFinale: number;
  delta: number;
  deltaPercent: number;
  stato: "In aumento" | "In diminuzione" | "Stabile";
  fileIniziali: string[];
  fileFinali: string[];
};

type ApiExtractedFile = {
  fileName: string;
  commessa: string;
  importo: number;
  fase: "iniziale" | "finale";
  warning?: string;
};

export default function VariazioniEconomichePage() {
  const [fase, setFase] = useState<"iniziale" | "finale">("iniziale");
  const [files, setFiles] = useState<EconomicFile[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [lastMessage, setLastMessage] = useState("");

  const kpis = useMemo(() => buildKpis(files), [files]);

  const totals = useMemo(() => {
    const iniziale = kpis.reduce((sum, kpi) => sum + kpi.importoIniziale, 0);
    const finale = kpis.reduce((sum, kpi) => sum + kpi.importoFinale, 0);
    const delta = finale - iniziale;
    const deltaPercent = iniziale > 0 ? (delta / iniziale) * 100 : 0;

    return {
      iniziale,
      finale,
      delta,
      deltaPercent,
      aumento: kpis.filter((k) => k.stato === "In aumento").length,
      diminuzione: kpis.filter((k) => k.stato === "In diminuzione").length,
      stabile: kpis.filter((k) => k.stato === "Stabile").length,
    };
  }, [kpis]);

  async function handlePdfUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const selected = Array.from(event.target.files || []);
    if (!selected.length) return;

    const pdfFiles = selected.filter((file) => file.name.toLowerCase().endsWith(".pdf"));

    if (!pdfFiles.length) {
      alert("Carica almeno un file PDF.");
      event.target.value = "";
      return;
    }

    setIsLoading(true);
    setLastMessage("Analisi PDF in corso...");

    try {
      const formData = new FormData();
      formData.append("fase", fase);

      pdfFiles.forEach((file) => {
        formData.append("files", file);
      });

      const response = await fetch("/api/variazioni-economiche", {
        method: "POST",
        body: formData,
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data?.error || "Errore durante l'analisi dei PDF.");
      }

      const extractedFiles: ApiExtractedFile[] = data.files || [];

      const newRows: EconomicFile[] = extractedFiles.map((file) => ({
        id: `${file.fileName}-${Date.now()}-${Math.random()}`,
        commessa: file.commessa || cleanProjectName(file.fileName),
        fase: file.fase || fase,
        fileName: file.fileName,
        importo: Number(file.importo || 0),
        warning: file.warning,
      }));

      setFiles((prev) => [...prev, ...newRows]);

      const warnings = newRows.filter((row) => row.warning || row.importo <= 0);

      if (warnings.length > 0) {
        setLastMessage(
          `${newRows.length} PDF analizzati. ${warnings.length} documento/i richiedono verifica perché l'importo non è stato rilevato.`
        );
      } else {
        setLastMessage(`${newRows.length} PDF analizzati correttamente.`);
      }
    } catch (error) {
      alert(
        error instanceof Error
          ? error.message
          : "Errore durante l'analisi dei PDF economici."
      );
      setLastMessage("");
    } finally {
      setIsLoading(false);
      event.target.value = "";
    }
  }

  function clearData() {
    setFiles([]);
    setLastMessage("");
  }

  function updateFileField(
    id: string,
    field: "commessa" | "importo" | "fase",
    value: string
  ) {
    setFiles((prev) =>
      prev.map((file) => {
        if (file.id !== id) return file;

        if (field === "importo") {
          return { ...file, importo: parseImporto(value), warning: undefined };
        }

        if (field === "fase") {
          return { ...file, fase: value as "iniziale" | "finale" };
        }

        return { ...file, commessa: value };
      })
    );
  }

  function exportReport() {
    if (kpis.length === 0) {
      alert("Nessun dato da esportare.");
      return;
    }

    const rows = kpis.map((kpi) => ({
      Commessa: kpi.commessa,
      "Importo iniziale": kpi.importoIniziale,
      "Importo finale": kpi.importoFinale,
      "Delta economico": kpi.delta,
      "Delta %": Number(kpi.deltaPercent.toFixed(2)),
      Stato: kpi.stato,
      "File fase iniziale": kpi.fileIniziali.join(", "),
      "File fase finale": kpi.fileFinali.join(", "),
    }));

    const summaryRows = [
      {
        Commessa: "TOTALE COMMESSE",
        "Importo iniziale": totals.iniziale,
        "Importo finale": totals.finale,
        "Delta economico": totals.delta,
        "Delta %": Number(totals.deltaPercent.toFixed(2)),
        Stato: totals.delta > 0 ? "In aumento" : totals.delta < 0 ? "In diminuzione" : "Stabile",
        "File fase iniziale": "",
        "File fase finale": "",
      },
    ];

    const allRows = [...summaryRows, ...rows];
    const csvHeader = Object.keys(allRows[0]).join(";");
    const csvRows = allRows.map((row) =>
      Object.values(row)
        .map((value) => `"${String(value).replace(/"/g, '""')}"`)
        .join(";")
    );

    const csv = "\ufeff" + [csvHeader, ...csvRows].join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");

    a.href = url;
    a.download = "analisi_variazioni_economiche.csv";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  const maxDelta = Math.max(...kpis.map((k) => Math.abs(k.delta)), 1);

  return (
    <main style={pageStyle}>
      <div style={contentStyle}>
        <Link href="/dashboard-pm" style={backLinkStyle}>
          ← Torna alla Dashboard PM
        </Link>

        <h1 style={titleStyle}>Analisi variazioni economiche</h1>

        <p style={leadStyle}>
          Carica i PDF economici della fase iniziale e finale della verifica. Il sistema prova a leggere
          automaticamente nome commessa e importo dal documento; se un dato non viene rilevato, puoi correggerlo
          nella tabella di verifica.
        </p>

        <section style={sectionStyle}>
          <h2 style={h2Style}>Caricamento elaborati economici PDF</h2>

          <div style={formGridStyle}>
            <label>
              <b>Fase documento</b>
              <select
                value={fase}
                onChange={(e) => setFase(e.target.value as "iniziale" | "finale")}
                style={inputStyle}
              >
                <option value="iniziale">Fase iniziale / arrivo progetto in ITS</option>
                <option value="finale">Fase finale / conclusione verifica</option>
              </select>
            </label>
          </div>

          <div style={{ marginTop: 18 }}>
            <input type="file" multiple accept=".pdf" onChange={handlePdfUpload} disabled={isLoading} />
          </div>

          <p style={helpStyle}>
            Non devi inserire manualmente commessa e importo prima del caricamento. Il sistema li estrae dal PDF e
            mostra sotto i dati rilevati.
          </p>

          {lastMessage && (
            <div
              style={{
                marginTop: 12,
                padding: 12,
                borderRadius: 10,
                background: "#eff6ff",
                border: "1px solid #bfdbfe",
                color: "#1e40af",
                fontWeight: 700,
              }}
            >
              {lastMessage}
            </div>
          )}

          <div style={actionsStyle}>
            <button
              type="button"
              onClick={exportReport}
              disabled={kpis.length === 0}
              style={{
                ...buttonStyle,
                background: kpis.length === 0 ? "#cbd5e1" : "#0f172a",
                cursor: kpis.length === 0 ? "not-allowed" : "pointer",
              }}
            >
              Esporta report
            </button>

            <button
              type="button"
              onClick={clearData}
              disabled={files.length === 0}
              style={{
                ...buttonStyle,
                background: files.length === 0 ? "#cbd5e1" : "#94a3b8",
                cursor: files.length === 0 ? "not-allowed" : "pointer",
              }}
            >
              Svuota dati
            </button>
          </div>
        </section>

        <section style={sectionStyle}>
          <h2 style={h2Style}>Dati rilevati dai PDF</h2>

          <div style={{ overflowX: "auto" }}>
            <table style={tableStyle}>
              <thead>
                <tr style={{ background: "#f8fafc" }}>
                  <th style={thStyle}>File</th>
                  <th style={thStyle}>Commessa rilevata</th>
                  <th style={thStyle}>Fase</th>
                  <th style={thStyle}>Importo rilevato</th>
                  <th style={thStyle}>Esito</th>
                </tr>
              </thead>

              <tbody>
                {files.length === 0 ? (
                  <tr>
                    <td style={tdStyle} colSpan={5}>
                      Nessun PDF caricato.
                    </td>
                  </tr>
                ) : (
                  files.map((file) => (
                    <tr key={file.id}>
                      <td style={tdStyle}>{file.fileName}</td>

                      <td style={tdStyle}>
                        <input
                          value={file.commessa}
                          onChange={(e) => updateFileField(file.id, "commessa", e.target.value)}
                          style={smallInputStyle}
                        />
                      </td>

                      <td style={tdStyle}>
                        <select
                          value={file.fase}
                          onChange={(e) => updateFileField(file.id, "fase", e.target.value)}
                          style={smallInputStyle}
                        >
                          <option value="iniziale">Iniziale</option>
                          <option value="finale">Finale</option>
                        </select>
                      </td>

                      <td style={tdStyle}>
                        <input
                          value={file.importo > 0 ? String(file.importo) : ""}
                          onChange={(e) => updateFileField(file.id, "importo", e.target.value)}
                          placeholder="Da verificare"
                          style={smallInputStyle}
                        />
                      </td>

                      <td style={tdStyle}>
                        {file.importo > 0 ? "Letto automaticamente" : "Da verificare"}
                        {file.warning ? ` - ${file.warning}` : ""}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </section>

        <section style={sectionStyle}>
          <h2 style={h2Style}>KPI globali</h2>

          <div style={kpiGridStyle}>
            <KpiCard title="Commesse analizzate" value={String(kpis.length)} />
            <KpiCard title="Importo iniziale totale" value={formatCurrency(totals.iniziale)} />
            <KpiCard title="Importo finale totale" value={formatCurrency(totals.finale)} />
            <KpiCard title="Delta economico totale" value={formatCurrency(totals.delta)} />
            <KpiCard title="Delta % totale" value={formatPercent(totals.deltaPercent)} />
            <KpiCard title="Commesse in aumento" value={String(totals.aumento)} />
            <KpiCard title="Commesse in diminuzione" value={String(totals.diminuzione)} />
            <KpiCard title="Commesse stabili" value={String(totals.stabile)} />
          </div>
        </section>

        <section style={sectionStyle}>
          <h2 style={h2Style}>Grafico delta economico per commessa</h2>

          {kpis.length === 0 ? (
            <p style={{ color: "#64748b", margin: 0 }}>
              Carica almeno una fase iniziale e una fase finale per visualizzare il grafico.
            </p>
          ) : (
            <div style={{ display: "grid", gap: 16 }}>
              {kpis.map((kpi) => {
                const width = Math.max(4, (Math.abs(kpi.delta) / maxDelta) * 100);

                return (
                  <div key={kpi.commessa}>
                    <div style={barHeaderStyle}>
                      <span>{kpi.commessa}</span>
                      <span>
                        {formatCurrency(kpi.delta)} · {formatPercent(kpi.deltaPercent)}
                      </span>
                    </div>

                    <div style={barTrackStyle}>
                      <div
                        style={{
                          width: `${width}%`,
                          height: "100%",
                          background: kpi.delta > 0 ? "#dc2626" : kpi.delta < 0 ? "#16a34a" : "#64748b",
                        }}
                      />
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </section>

        <section style={sectionStyle}>
          <h2 style={h2Style}>Report commesse</h2>

          <div style={{ overflowX: "auto" }}>
            <table style={tableStyle}>
              <thead>
                <tr style={{ background: "#f8fafc" }}>
                  <th style={thStyle}>Commessa</th>
                  <th style={thStyle}>Importo iniziale</th>
                  <th style={thStyle}>Importo finale</th>
                  <th style={thStyle}>Delta €</th>
                  <th style={thStyle}>Delta %</th>
                  <th style={thStyle}>Stato</th>
                  <th style={thStyle}>PDF iniziali</th>
                  <th style={thStyle}>PDF finali</th>
                </tr>
              </thead>

              <tbody>
                {kpis.length === 0 ? (
                  <tr>
                    <td style={tdStyle} colSpan={8}>
                      Nessuna commessa economica analizzata.
                    </td>
                  </tr>
                ) : (
                  kpis.map((kpi) => (
                    <tr key={kpi.commessa}>
                      <td style={tdStyle}>{kpi.commessa}</td>
                      <td style={tdStyle}>{formatCurrency(kpi.importoIniziale)}</td>
                      <td style={tdStyle}>{formatCurrency(kpi.importoFinale)}</td>
                      <td style={tdStyle}>{formatCurrency(kpi.delta)}</td>
                      <td style={tdStyle}>{formatPercent(kpi.deltaPercent)}</td>
                      <td style={tdStyle}>{kpi.stato}</td>
                      <td style={tdStyle}>{kpi.fileIniziali.join(", ") || "-"}</td>
                      <td style={tdStyle}>{kpi.fileFinali.join(", ") || "-"}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    </main>
  );
}

function buildKpis(files: EconomicFile[]): EconomicKpi[] {
  const grouped = new Map<string, EconomicFile[]>();

  files
    .filter((file) => file.importo > 0 && file.commessa.trim())
    .forEach((file) => {
      grouped.set(file.commessa, [...(grouped.get(file.commessa) || []), file]);
    });

  return (Array.from(grouped.entries()) as [string, EconomicFile[]][])
    .map(([commessa, rows]) => {
      const iniziali = rows.filter((row) => row.fase === "iniziale");
      const finali = rows.filter((row) => row.fase === "finale");

      const importoIniziale = iniziali.reduce((sum, row) => sum + row.importo, 0);
      const importoFinale = finali.reduce((sum, row) => sum + row.importo, 0);
      const delta = importoFinale - importoIniziale;
      const deltaPercent = importoIniziale > 0 ? (delta / importoIniziale) * 100 : 0;

      const stato: EconomicKpi["stato"] =
        delta > 0 ? "In aumento" : delta < 0 ? "In diminuzione" : "Stabile";

      return {
        commessa,
        importoIniziale,
        importoFinale,
        delta,
        deltaPercent,
        stato,
        fileIniziali: iniziali.map((row) => row.fileName),
        fileFinali: finali.map((row) => row.fileName),
      };
    })
    .filter((kpi) => kpi.importoIniziale > 0 || kpi.importoFinale > 0);
}

function cleanProjectName(fileName: string): string {
  return fileName.replace(/\.pdf$/i, "").replace(/[_-]+/g, " ").trim();
}

function parseImporto(value: string): number {
  const normalized = value.replace(/\s/g, "").replace(/\./g, "").replace(",", ".");
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatCurrency(value: number): string {
  return new Intl.NumberFormat("it-IT", {
    style: "currency",
    currency: "EUR",
    maximumFractionDigits: 0,
  }).format(value || 0);
}

function formatPercent(value: number): string {
  if (!Number.isFinite(value)) return "0%";
  return `${value >= 0 ? "+" : ""}${value.toFixed(1)}%`;
}

function KpiCard({ title, value }: { title: string; value: string }) {
  return (
    <div style={kpiCardStyle}>
      <div style={{ fontSize: 14, color: "#64748b", marginBottom: 10 }}>{title}</div>
      <div style={{ fontSize: 28, fontWeight: 800 }}>{value}</div>
    </div>
  );
}

const pageStyle: React.CSSProperties = {
  minHeight: "100vh",
  background: "#f1f5f9",
  padding: 24,
  fontFamily: "Arial, sans-serif",
  color: "#0f172a",
};

const contentStyle: React.CSSProperties = {
  width: "100%",
  maxWidth: 1800,
  margin: "0 auto",
};

const backLinkStyle: React.CSSProperties = {
  display: "inline-block",
  marginBottom: 24,
  padding: "10px 18px",
  borderRadius: 10,
  background: "#0f172a",
  color: "#fff",
  textDecoration: "none",
  fontWeight: 700,
};

const titleStyle: React.CSSProperties = {
  fontSize: 34,
  fontWeight: 800,
  marginBottom: 10,
};

const leadStyle: React.CSSProperties = {
  fontSize: 16,
  color: "#475569",
  marginBottom: 28,
  lineHeight: 1.6,
};

const sectionStyle: React.CSSProperties = {
  background: "#ffffff",
  border: "1px solid #cbd5e1",
  borderRadius: 16,
  padding: 20,
  marginBottom: 24,
};

const h2Style: React.CSSProperties = {
  marginTop: 0,
  marginBottom: 14,
  fontSize: 22,
  fontWeight: 800,
};

const helpStyle: React.CSSProperties = {
  marginTop: 10,
  color: "#64748b",
  fontSize: 13,
};

const formGridStyle: React.CSSProperties = {
  display: "grid",
  gridTemplateColumns: "repeat(auto-fit, minmax(260px, 1fr))",
  gap: 16,
};

const inputStyle: React.CSSProperties = {
  display: "block",
  width: "100%",
  marginTop: 8,
  padding: 12,
  border: "1px solid #cbd5e1",
  borderRadius: 10,
  background: "#fff",
};

const smallInputStyle: React.CSSProperties = {
  width: "100%",
  minWidth: 160,
  padding: 8,
  border: "1px solid #cbd5e1",
  borderRadius: 8,
  background: "#fff",
};

const actionsStyle: React.CSSProperties = {
  display: "flex",
  gap: 12,
  marginTop: 18,
};

const buttonStyle: React.CSSProperties = {
  padding: "12px 18px",
  border: 0,
  borderRadius: 10,
  color: "#fff",
  fontWeight: 700,
};

const kpiGridStyle: React.CSSProperties = {
  display: "grid",
  gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
  gap: 16,
};

const kpiCardStyle: React.CSSProperties = {
  background: "#f8fafc",
  border: "1px solid #cbd5e1",
  borderRadius: 14,
  padding: 20,
};

const barHeaderStyle: React.CSSProperties = {
  display: "flex",
  justifyContent: "space-between",
  gap: 12,
  marginBottom: 6,
  fontSize: 14,
  fontWeight: 700,
};

const barTrackStyle: React.CSSProperties = {
  width: "100%",
  background: "#e2e8f0",
  borderRadius: 999,
  overflow: "hidden",
  height: 24,
};

const tableStyle: React.CSSProperties = {
  width: "100%",
  borderCollapse: "collapse",
  minWidth: 1100,
};

const thStyle: React.CSSProperties = {
  textAlign: "left",
  padding: 12,
  borderBottom: "1px solid #cbd5e1",
  fontSize: 14,
  whiteSpace: "nowrap",
};

const tdStyle: React.CSSProperties = {
  padding: 12,
  borderBottom: "1px solid #e2e8f0",
  fontSize: 14,
  verticalAlign: "top",
};
