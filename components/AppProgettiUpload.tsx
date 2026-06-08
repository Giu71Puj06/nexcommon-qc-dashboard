"use client";

import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
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

function toDashboardExportRows(rows: any[]) {
  return rows.map((r: any) => ({
    ID: r.id || "",
    Rilievi: r.tipo || "",
    Disciplina: r.disciplina || "",
    Elaborato: getElaboratoKey(r),
    Descrizione: r.descrizione || "",
    Stato: translateStatus(r.stato),
    "Storico commenti": commentsToText(Array.isArray(r.comments) ? r.comments : []),
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

      {data.map((d: any) => (
        <div
          key={d.label}
          onClick={() => onClick(d.label)}
          style={{
            marginBottom: 12,
            cursor: "pointer",
            background: activeKey === d.label ? "#e0f2fe" : "transparent",
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
      ))}
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
              {f.type === "bcfzip" &&
                `- BCF letto: ${f.markupCount || 0} topic, ${f.comments || 0} commenti`}
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
  if (!title) return null;

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

      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#f1f5f9" }}>
              <th style={th}>ID</th>
              <th style={th}>Rilievi</th>
              <th style={th}>Disciplina</th>
              <th style={th}>Elaborato</th>
              <th style={th}>Descrizione</th>
              <th style={th}>Stato</th>
              <th style={th}>Storico commenti</th>
            </tr>
          </thead>

          <tbody>
            {rows.map((r: any, i: number) => {
              const allComments = Array.isArray(r.comments) ? r.comments : [];

              return (
                <tr key={`${r.id}-${i}`}>
                  <td style={td}>{r.id}</td>
                  <td style={td}>{r.tipo}</td>
                  <td style={td}>{r.disciplina}</td>
                  <td style={td}>{getElaboratoKey(r)}</td>
                  <td style={td}>{r.descrizione}</td>
                  <td style={td}>{translateStatus(r.stato)}</td>
                  <td style={td}>
                    <CommentList comments={allComments} emptyText="Nessun commento" />
                  </td>
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
      elaboratoKey: getElaboratoKey(r),
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

    const elaborato = r.elaboratoKey || "Elaborato non identificato";

    if (!rilieviPerElaborato[elaborato]) {
      rilieviPerElaborato[elaborato] = {
        label: elaborato,
        value: 0,
        nc: 0,
        oss: 0,
      };
    }

    if (r.tipo === "NC") {
      rilieviPerElaborato[elaborato].value += 1;
      rilieviPerElaborato[elaborato].nc += 1;
    }

    if (r.tipo === "OSS") {
      rilieviPerElaborato[elaborato].value += 1;
      rilieviPerElaborato[elaborato].oss += 1;
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
              accept=".xlsx,.xls,.bcfzip,.zip,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
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
              Analizza ToDo / BCF / BCFZIP
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
            Carica uno o più file ToDo .xlsx, .bcf o .bcfzip esportati da Trimble Connect
            per generare la sintesi NC/OSS.
          </div>
        </Card>
      )}

      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          gap: 10,
          marginTop: 24,
          marginBottom: 12,
          flexWrap: "wrap",
        }}
      >
        <ExportButton
          onClick={() =>
            exportExcel(
              "Dashboard_NC_OSS_completa",
              toDashboardExportRows(enrichedRows)
            )
          }
        >
          Export Dashboard Excel
        </ExportButton>

        <ExportButton
          onClick={() =>
            exportExcel(
              "Rilievi_per_disciplina",
              toChartExportRows(disciplineData)
            )
          }
        >
          Export Rilievi per disciplina
        </ExportButton>

        <ExportButton
          onClick={() =>
            exportExcel(
              "Rilievi",
              toChartExportRows(esitiData)
            )
          }
        >
          Export Rilievi
        </ExportButton>

        <ExportButton
          onClick={() =>
            exportExcel(
              "NC_OSS_per_elaborato",
              toChartExportRows(rilieviPerElaboratoData)
            )
          }
        >
          Export NC/OSS per elaborato
        </ExportButton>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(6, 1fr)", gap: 12, marginTop: 24, marginBottom: 12 }}>
        <KPI title="Elaborati totali" value={elaboratiTot} onClick={() => setSelection({ type: "kpi", value: "totali", label: "KPI", valueLabel: "Elaborati totali" })} />
        <KPI title="Elaborati con NC" value={elaboratiNC} onClick={() => setSelection({ type: "kpi", value: "nc", label: "KPI", valueLabel: "Elaborati con NC" })} />
        <KPI title="Elaborati con OSS" value={elaboratiOSS} onClick={() => setSelection({ type: "kpi", value: "oss", label: "KPI", valueLabel: "Elaborati con OSS" })} />
        <KPI title="Elaborati senza rilievi" value={elaboratiOK} onClick={() => setSelection({ type: "kpi", value: "nessun", label: "KPI", valueLabel: "Elaborati senza rilievi" })} />
        <KPI title="Totale NC" value={totaleNC} onClick={() => setSelection({ type: "tipo", value: "NC", label: "Rilievi" })} />
        <KPI title="Totale OSS" value={totaleOSS} onClick={() => setSelection({ type: "tipo", value: "OSS", label: "Rilievi" })} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 24 }}>
        <KPI title="Totale NC + OSS" value={totaleNCOSS} onClick={() => setSelection({ type: "kpi", value: "totali", label: "KPI", valueLabel: "Tutti i rilievi" })} />
        <KPI title="In attesa di riscontro dell'ispettore" value={daVerificareISP} subtitle="Ultimo commento PRG" onClick={() => setSelection({ type: "kpi", value: "da-verificare-isp", label: "KPI", valueLabel: "Da verificare ISP" })} />
        <KPI title="In attesa di risposta del progettista" value={daRisponderePRG} subtitle="Nessun PRG o ultimo ISP" onClick={() => setSelection({ type: "kpi", value: "da-rispondere-prg", label: "KPI", valueLabel: "Da rispondere PRG" })} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16 }}>
        <BarList title="Rilievi per disciplina" data={disciplineData} activeKey={selection?.type === "disciplina" ? selection.value : ""} onClick={(value: string) => setSelection({ type: "disciplina", value, label: "Disciplina" })} onExport={() => exportExcel("Rilievi_per_disciplina", toChartExportRows(disciplineData))} />
        <BarList title="Rilievi" data={esitiData} activeKey={selection?.type === "tipo" ? selection.value : ""} onClick={(value: string) => setSelection({ type: "tipo", value, label: "Rilievi" })} onExport={() => exportExcel("Rilievi", toChartExportRows(esitiData))} />
        <BarList title="NC / OSS per elaborato" data={rilieviPerElaboratoData} activeKey={selection?.type === "elaborato" ? selection.value : ""} onClick={(value: string) => setSelection({ type: "elaborato", value, label: "Elaborato" })} onExport={() => exportExcel("NC_OSS_per_elaborato", toChartExportRows(rilieviPerElaboratoData))} />
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
