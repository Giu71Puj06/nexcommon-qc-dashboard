"use client";

import React, { useMemo, useState } from "react";
import Link from "next/link";
import * as XLSX from "xlsx";

type CheckRow = {
  rowNumber: number;
  progressivo: number;
  label: string;
  title: string;
  codiceTitleTrimble: string;
  codiceReport: string;
  titleOk: boolean;
  tags: string;
  tagsOk: boolean;
  disciplina: string;
  disciplinaOk: boolean;
  status: string;
  statusOk: boolean;
  description: string;
  tr: string;
  bcfTitle: string;
  bcfDescription: string;
  rispostaProgettista: string;
  riscontroIspettore: string;
  storiaOk: boolean;
  esitoStoria:
    | "COMPLETA"
    | "Manca commento del progettista"
    | "Manca il riscontro dell'ispettore"
    | "NON APPLICABILE";
  esito: "OK" | "ERRORE";
  livello?: "OK" | "WARNING" | "ERRORE";
  warning?: string[];
  anomalie: string[];
};

type Filters = {
  n: string;
  codiceReport: string;
  codiceTitleTrimble: string;
  esitoCodice: string;
  tags: string;
  disciplina: string;
  status: string;
  tr: string;
  esitoStoria: string;
  esito: string;
  anomalie: string;
};

type ApiSummary = {
  totale: number;
  ok: number;
  errori: number;
  warning: number;
  storieComplete: number;
  bcfWarning: number;
  completezza: number;
};


function esitoIcon(ok: boolean) {
  return ok ? "✅" : "❌";
}


function isErroreBloccante(row: CheckRow) {
  const anomalie = row.anomalie || [];

  return anomalie.some((a) => {
    const t = String(a || "").toLowerCase();

    return (
      t.includes("title mancante") ||
      t.includes("title contiene .pdf") ||
      t.includes("disciplina mancante")
    );
  });
}

function hasDocumentiGenerali(value: string) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .includes("documenti generali");
}

function disciplinaOkReale(row: CheckRow) {
  if (!row.disciplina) return false;
  if (hasDocumentiGenerali(row.disciplina)) return true;

  // Una disciplina valorizzata è valida ai fini del report:
  // l'eventuale mancato allineamento con ELENCO_ELABORATI resta warning.
  return true;
}

function titleOkReale(row: CheckRow) {
  return !isErroreBloccante(row) || row.titleOk;
}

function tagsOkReale(row: CheckRow) {
  // Tags mancanti / non perfettamente riconosciuti sono warning.
  return true;
}

function statusOkReale(row: CheckRow) {
  return !String(row.status || "").trim() ? true : row.statusOk;
}

function livelloReale(row: CheckRow): "OK" | "WARNING" | "ERRORE" {
  if (isErroreBloccante(row)) return "ERRORE";

  const anomalie = row.anomalie || [];
  const warning = row.warning || [];
  const hasWarning =
    warning.length > 0 ||
    anomalie.some((a) => {
      const t = String(a || "").toLowerCase();

      return (
        t.includes("codice elaborato non presente") ||
        t.includes("disciplina non presente") ||
        t.includes("bcf non trovato") ||
        t.includes("manca risposta") ||
        t.includes("manca riscontro") ||
        t.includes("chiuso senza riscontro") ||
        t.includes("tags mancanti")
      );
    });

  return hasWarning ? "WARNING" : "OK";
}

function esitoReale(row: CheckRow): "OK" | "ERRORE" {
  return isErroreBloccante(row) ? "ERRORE" : "OK";
}

export default function ControlloTodoIspettoriPage() {
  const [todoFile, setTodoFile] = useState<File | null>(null);
  const [reportFile, setReportFile] = useState<File | null>(null);
  const [elencoFile, setElencoFile] = useState<File | null>(null);
  const [bcfFiles, setBcfFiles] = useState<File[]>([]);

  const [checks, setChecks] = useState<CheckRow[]>([]);
  const [summary, setSummary] = useState<ApiSummary | null>(null);
  const [bcfTopicsCount, setBcfTopicsCount] = useState(0);
  const [error, setError] = useState("");

  const [loading, setLoading] = useState(false);

  const [filters, setFilters] = useState<Filters>({
    n: "",
    codiceReport: "",
    codiceTitleTrimble: "",
    esitoCodice: "",
    tags: "",
    disciplina: "",
    status: "",
    tr: "",
    esitoStoria: "",
    esito: "",
    anomalie: "",
  });

  function updateFilter(key: keyof Filters, value: string) {
    setFilters((prev) => ({
      ...prev,
      [key]: value,
    }));
  }

  async function eseguiControllo() {
    if (!todoFile || !reportFile || !elencoFile) {
      alert("Carica ToDo XLSX, Report_Completo.xlsx ed ELENCO_ELABORATI.xlsx.");
      return;
    }

    setLoading(true);
    setError("");

    try {
      const fd = new FormData();
      fd.append("todo", todoFile);
      fd.append("report", reportFile);
      fd.append("elenco", elencoFile);
      bcfFiles.forEach((file) => fd.append("bcf", file));

      const res = await fetch("/api/controllo-todo-ispettori", {
        method: "POST",
        body: fd,
      });

      const data = await res.json();

      if (!res.ok || data.ok === false) {
        throw new Error(data.error || "Errore durante il controllo ToDo");
      }

      setChecks(data.checks || []);
      setSummary(data.summary || null);
      setBcfTopicsCount(data.bcfTopicsCount || 0);
    } catch (err: any) {
      setError(err.message || "Errore imprevisto");
      setChecks([]);
      setSummary(null);
      setBcfTopicsCount(0);
    } finally {
      setLoading(false);
    }
  }

  const filteredChecks = useMemo(() => {
    return checks.filter((row) => {
      const n = `${row.progressivo}${row.label ? ` (${row.label})` : ""}`;
      const esitoCodice = titleOkReale(row) ? "OK" : "ERRORE";
      const anomalie = [...(row.anomalie || []), ...(row.warning || [])].join(" | ");
      const esitoEffettivo = livelloReale(row);

      return (
        n.toLowerCase().includes(filters.n.toLowerCase()) &&
        row.codiceReport
          .toLowerCase()
          .includes(filters.codiceReport.toLowerCase()) &&
        row.codiceTitleTrimble
          .toLowerCase()
          .includes(filters.codiceTitleTrimble.toLowerCase()) &&
        (filters.esitoCodice === "" ||
          esitoCodice === filters.esitoCodice) &&
        row.tags.toLowerCase().includes(filters.tags.toLowerCase()) &&
        row.disciplina
          .toLowerCase()
          .includes(filters.disciplina.toLowerCase()) &&
        row.status.toLowerCase().includes(filters.status.toLowerCase()) &&
        row.tr.toLowerCase().includes(filters.tr.toLowerCase()) &&
        (filters.esitoStoria === "" || row.esitoStoria === filters.esitoStoria) &&
        (filters.esito === "" || esitoEffettivo === filters.esito) &&
        anomalie.toLowerCase().includes(filters.anomalie.toLowerCase())
      );
    });
  }, [checks, filters]);

  const totale = checks.length;
  const errori = checks.filter((r) => livelloReale(r) === "ERRORE").length;
  const warning = checks.filter((r) => livelloReale(r) === "WARNING").length;
  const ok = checks.filter((r) => livelloReale(r) === "OK").length;
  const completezza = totale > 0 ? Math.round((ok / totale) * 100) : 0;
  const storieComplete = summary?.storieComplete || checks.filter((r) => r.esitoStoria === "COMPLETA").length;
  const bcfWarning = summary?.bcfWarning || checks.filter(
    (r) =>
      r.esitoStoria === "Manca commento del progettista" ||
      r.esitoStoria === "Manca il riscontro dell'ispettore"
  ).length;

  const completezzaColor =
    completezza === 100 ? "#16a34a" : completezza >= 80 ? "#f59e0b" : "#dc2626";

  function esportaExcel() {
    const rows = filteredChecks.map((row) => ({
      "N.": `${row.progressivo}${row.label ? ` (${row.label})` : ""}`,
      TR: row.tr || "",
      "Codice elaborato Report": row.codiceReport || "",
      "Codice elaborato nel Title Trimble": row.codiceTitleTrimble || "",
      "Esito codice": titleOkReale(row) ? "OK" : "ERRORE",
      Tags: row.tags || "",
      "Esito Tags": tagsOkReale(row) ? "OK" : "WARNING",
      Disciplina: row.disciplina || "",
      "Esito Disciplina": disciplinaOkReale(row) ? "OK" : "WARNING",
      Status: row.status || "",
      "Esito Status": statusOkReale(row) ? "OK" : "WARNING",
      "Description ToDo": row.description || "",
      "Titolo BCF": row.bcfTitle || "",
      "Descrizione BCF": row.bcfDescription || "",
      "Risposta progettista": row.rispostaProgettista || "",
      "Riscontro ispettore ITS": row.riscontroIspettore || "",
      "Esito storia rilievo": row.esitoStoria,
      Livello: livelloReale(row),
      Esito: esitoReale(row),
      Warning: (row.warning || []).join(" | "),
      Anomalie: (row.anomalie || []).join(" | "),
    }));

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(rows),
      "Controllo ToDo"
    );

    XLSX.writeFile(workbook, "Report_Controllo_ToDo_BCF.xlsx");
  }

  return (
    <main
      style={{
        minHeight: "100vh",
        padding: 30,
        background: "#f1f5f9",
        fontFamily: "Arial, sans-serif",
      }}
    >
      <Link
        href="/"
        style={{
          color: "#0284c7",
          fontWeight: 700,
          textDecoration: "none",
        }}
      >
        ← Torna alla dashboard
      </Link>

      <h1 style={{ fontSize: 36, marginTop: 24 }}>
        Controllo ToDo Ispettori
      </h1>

      <p style={{ color: "#475569", fontSize: 18, maxWidth: 1150 }}>
        Verifica automatica di Title, Tags, Disciplina e Status. La storia BCF è mostrata come warning:
        non abbassa la completezza e non genera falsi errori se le schede Word sono già corrette.
      </p>

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(4, 1fr)",
          gap: 16,
          marginTop: 24,
          maxWidth: 1500,
        }}
      >
        <div style={cardStyle}>
          <b>ToDo XLSX</b>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => setTodoFile(e.target.files?.[0] || null)}
            style={inputStyle}
          />
        </div>

        <div style={cardStyle}>
          <b>Report_Completo.xlsx</b>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => setReportFile(e.target.files?.[0] || null)}
            style={inputStyle}
          />
        </div>

        <div style={cardStyle}>
          <b>ELENCO_ELABORATI.xlsx</b>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => setElencoFile(e.target.files?.[0] || null)}
            style={inputStyle}
          />
        </div>

        <div style={cardStyle}>
          <b>BCFZIP / ZIP</b>
          <input
            type="file"
            accept=".bcfzip,.zip,.bcf,.xml"
            multiple
            onChange={(e) => setBcfFiles(Array.from(e.target.files || []))}
            style={inputStyle}
          />
          <div style={{ marginTop: 8, color: "#64748b", fontSize: 12 }}>
            Opzionale, ma consigliato per verificare la storia dei rilievi.
          </div>
          {bcfFiles.length > 0 && (
            <div style={{ marginTop: 8, color: "#64748b", fontSize: 12 }}>
              File BCF caricati: {bcfFiles.length}
            </div>
          )}
        </div>
      </div>

      <button
        onClick={eseguiControllo}
        disabled={loading}
        style={{
          marginTop: 20,
          padding: "14px 22px",
          background: loading ? "#334155" : "#0f172a",
          color: "white",
          border: "none",
          borderRadius: 12,
          fontWeight: 700,
          cursor: loading ? "not-allowed" : "pointer",
        }}
      >
        {loading ? "Controllo in corso..." : "Esegui controllo ToDo"}
      </button>

      {error && (
        <div style={{ marginTop: 18, padding: 14, borderRadius: 12, background: "#fee2e2", color: "#991b1b" }}>
          {error}
        </div>
      )}

      {totale > 0 && (
        <>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(7, 1fr)",
              gap: 12,
              marginTop: 28,
            }}
          >
            <div style={cardStyle}>
              <div style={kpiLabel}>Completezza senza warning</div>
              <div
                style={{
                  fontSize: 36,
                  fontWeight: 800,
                  color: completezzaColor,
                }}
              >
                {completezza}%
              </div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>Righe ToDo controllate</div>
              <div style={kpiValue}>{totale}</div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>Righe OK</div>
              <div style={{ ...kpiValue, color: "#16a34a" }}>{ok}</div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>Errori veri</div>
              <div style={{ ...kpiValue, color: "#dc2626" }}>{errori}</div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>Warning BCF</div>
              <div style={{ ...kpiValue, color: "#f59e0b" }}>{warning}</div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>Storie complete</div>
              <div style={{ ...kpiValue, color: "#16a34a" }}>{storieComplete}</div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>Topic BCF letti</div>
              <div style={kpiValue}>{bcfTopicsCount}</div>
            </div>
          </div>

          <div style={{ marginTop: 28, ...cardStyle }}>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                gap: 16,
                marginBottom: 16,
              }}
            >
              <div>
                <h2 style={{ margin: 0 }}>Report controllo ToDo</h2>
                <div style={{ marginTop: 6, color: "#64748b", fontSize: 13 }}>
                  Righe visualizzate: {filteredChecks.length} su {checks.length}
                </div>
              </div>

              <button
                onClick={esportaExcel}
                style={{
                  padding: "10px 16px",
                  background: "#16a34a",
                  color: "white",
                  border: "none",
                  borderRadius: 10,
                  fontWeight: 700,
                  cursor: "pointer",
                }}
              >
                Esporta Excel
              </button>
            </div>

            <div style={{ overflowX: "auto" }}>
              <table
                style={{
                  width: "100%",
                  borderCollapse: "collapse",
                  fontSize: 13,
                }}
              >
                <thead>
                  <tr style={{ background: "#f8fafc" }}>
                    <th style={th}>N.</th>
                    <th style={th}>TR</th>
                    <th style={th}>Codice elaborato Report</th>
                    <th style={th}>Codice elaborato nel Title Trimble</th>
                    <th style={th}>Esito codice</th>
                    <th style={th}>Tags</th>
                    <th style={th}>Disciplina</th>
                    <th style={th}>Status</th>
                    <th style={th}>Description ToDo</th>
                    <th style={th}>Risposta progettista</th>
                    <th style={th}>Riscontro ITS</th>
                    <th style={th}>Storia rilievo</th>
                    <th style={th}>Livello</th>
                    <th style={th}>Anomalie / Warning</th>
                  </tr>

                  <tr style={{ background: "white" }}>
                    <th style={th}>
                      <input
                        value={filters.n}
                        onChange={(e) => updateFilter("n", e.target.value)}
                        placeholder="Filtra N."
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.tr}
                        onChange={(e) => updateFilter("tr", e.target.value)}
                        placeholder="Filtra TR"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.codiceReport}
                        onChange={(e) =>
                          updateFilter("codiceReport", e.target.value)
                        }
                        placeholder="Filtra codice"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.codiceTitleTrimble}
                        onChange={(e) =>
                          updateFilter("codiceTitleTrimble", e.target.value)
                        }
                        placeholder="Filtra title"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <select
                        value={filters.esitoCodice}
                        onChange={(e) =>
                          updateFilter("esitoCodice", e.target.value)
                        }
                        style={filterInput}
                      >
                        <option value="">Tutti</option>
                        <option value="OK">OK</option>
                        <option value="WARNING">WARNING</option>
                        <option value="ERRORE">ERRORE</option>
                      </select>
                    </th>

                    <th style={th}>
                      <input
                        value={filters.tags}
                        onChange={(e) => updateFilter("tags", e.target.value)}
                        placeholder="Filtra tags"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.disciplina}
                        onChange={(e) =>
                          updateFilter("disciplina", e.target.value)
                        }
                        placeholder="Filtra disciplina"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.status}
                        onChange={(e) => updateFilter("status", e.target.value)}
                        placeholder="Filtra status"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}></th>
                    <th style={th}></th>
                    <th style={th}></th>
                    <th style={th}>
                      <select
                        value={filters.esitoStoria}
                        onChange={(e) =>
                          updateFilter("esitoStoria", e.target.value)
                        }
                        style={filterInput}
                      >
                        <option value="">Tutte</option>
                        <option value="COMPLETA">COMPLETA</option>
                        <option value="Manca commento del progettista">
                          Manca commento del progettista
                        </option>
                        <option value="Manca il riscontro dell'ispettore">
                          Manca il riscontro dell'ispettore
                        </option>
                        <option value="NON APPLICABILE">NON APPLICABILE</option>
                      </select>
                    </th>

                    <th style={th}>
                      <select
                        value={filters.esito}
                        onChange={(e) => updateFilter("esito", e.target.value)}
                        style={filterInput}
                      >
                        <option value="">Tutti</option>
                        <option value="OK">OK</option>
                        <option value="ERRORE">ERRORE</option>
                      </select>
                    </th>

                    <th style={th}>
                      <input
                        value={filters.anomalie}
                        onChange={(e) =>
                          updateFilter("anomalie", e.target.value)
                        }
                        placeholder="Filtra anomalie"
                        style={filterInput}
                      />
                    </th>
                  </tr>
                </thead>

                <tbody>
                  {filteredChecks.map((row) => (
                    <tr key={row.rowNumber}>
                      <td style={td}>
                        {row.progressivo}
                        {row.label ? ` (${row.label})` : ""}
                      </td>

                      <td style={td}>{row.tr || "-"}</td>

                      <td style={td}>{row.codiceReport || ""}</td>

                      <td style={td}>{row.codiceTitleTrimble || "-"}</td>

                      <td style={td}>
                        <b
                          style={{
                            color: titleOkReale(row) ? "#16a34a" : "#dc2626",
                          }}
                        >
                          {esitoIcon(titleOkReale(row))}
                        </b>
                      </td>

                      <td style={td}>
                        {row.tags || "-"}{" "}
                        <b
                          style={{
                            color: tagsOkReale(row) ? "#16a34a" : "#f59e0b",
                          }}
                        >
                          {esitoIcon(tagsOkReale(row))}
                        </b>
                      </td>

                      <td style={td}>
                        {row.disciplina || "-"}{" "}
                        <b
                          style={{
                            color: disciplinaOkReale(row) ? "#16a34a" : "#f59e0b",
                          }}
                        >
                          {esitoIcon(disciplinaOkReale(row))}
                        </b>
                      </td>

                      <td style={td}>
                        {row.status || "-"}{" "}
                        <b
                          style={{
                            color: statusOkReale(row) ? "#16a34a" : "#f59e0b",
                          }}
                        >
                          {esitoIcon(statusOkReale(row))}
                        </b>
                      </td>

                      <td style={{ ...td, minWidth: 260, whiteSpace: "pre-wrap" }}>
                        {row.description || "-"}
                      </td>

                      <td style={{ ...td, minWidth: 260, whiteSpace: "pre-wrap" }}>
                        {row.rispostaProgettista || "-"}
                      </td>

                      <td style={{ ...td, minWidth: 260, whiteSpace: "pre-wrap" }}>
                        {row.riscontroIspettore || "-"}
                      </td>

                      <td
                        style={{
                          ...td,
                          fontWeight: 700,
                          color:
                            row.esitoStoria === "COMPLETA"
                              ? "#16a34a"
                              : row.esitoStoria === "NON APPLICABILE"
                              ? "#64748b"
                              : "#f59e0b",
                        }}
                      >
                        {row.esitoStoria}
                      </td>

                      <td
                        style={{
                          ...td,
                          fontWeight: 700,
                          color:
                            livelloReale(row) === "OK"
                              ? "#16a34a"
                              : livelloReale(row) === "WARNING"
                              ? "#f59e0b"
                              : "#dc2626",
                        }}
                      >
                        {livelloReale(row)}
                      </td>

                      <td style={td}>{[...(row.anomalie || []), ...(row.warning || [])].join(" | ")}</td>
                    </tr>
                  ))}

                  {filteredChecks.length === 0 && (
                    <tr>
                      <td style={td} colSpan={14}>
                        Nessuna riga trovata con i filtri impostati.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
    </main>
  );
}

const cardStyle: React.CSSProperties = {
  background: "white",
  border: "1px solid #e2e8f0",
  borderRadius: 16,
  padding: 16,
  boxShadow: "0 6px 18px rgba(15,23,42,.06)",
};

const inputStyle: React.CSSProperties = {
  display: "block",
  width: "100%",
  marginTop: 10,
};

const filterInput: React.CSSProperties = {
  width: "100%",
  padding: "6px 8px",
  border: "1px solid #cbd5e1",
  borderRadius: 6,
  fontSize: 12,
  boxSizing: "border-box",
};

const kpiLabel: React.CSSProperties = {
  fontSize: 13,
  color: "#64748b",
};

const kpiValue: React.CSSProperties = {
  fontSize: 36,
  fontWeight: 800,
  color: "#0f172a",
};

const th: React.CSSProperties = {
  border: "1px solid #e2e8f0",
  padding: 8,
  textAlign: "left",
};

const td: React.CSSProperties = {
  border: "1px solid #e2e8f0",
  padding: 8,
  verticalAlign: "top",
};
