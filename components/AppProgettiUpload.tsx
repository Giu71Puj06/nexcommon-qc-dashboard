"use client";

import React, { useMemo, useState } from "react";

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

function BarList({ title, data, onClick, activeKey }: any) {
  const max = Math.max(...data.map((d: any) => d.value), 1);

  return (
    <Card>
      <h3 style={{ marginTop: 0 }}>{title}</h3>

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

function DetailPanel({ rows, title, onReset }: any) {
  if (!title) return null;

  return (
    <Card>
      <div style={{ display: "flex", justifyContent: "space-between", gap: 12 }}>
        <h2 style={{ marginTop: 0 }}>Dettaglio selezione: {title}</h2>
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

      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#f1f5f9" }}>
              <th style={th}>ID</th>
              <th style={th}>Esito</th>
              <th style={th}>Tipologia NC/OSS</th>
              <th style={th}>Disciplina</th>
              <th style={th}>Elaborato</th>
              <th style={th}>Descrizione NC/OSS</th>
              <th style={th}>Stato</th>
              <th style={th}>Ispettore</th>
              <th style={th}>PRG</th>
              <th style={th}>ISP</th>
              <th style={th}>Azione</th>
              <th style={th}>Stato risoluzione</th>
              <th style={th}>Ultimo commento</th>
            </tr>
          </thead>

          <tbody>
            {rows.map((r: any, i: number) => (
              <tr key={`${r.id}-${i}`}>
                <td style={td}>{r.id}</td>
                <td style={td}>{r.tipo}</td>
                <td style={td}>{r.tipologiaNcOss || r.tipologiaDocumento || ""}</td>
                <td style={td}>{r.disciplina}</td>
                <td style={td}>{getElaboratoKey(r)}</td>
                <td style={td}>{r.descrizione}</td>
                <td style={td}>{r.stato}</td>
                <td style={td}>{r.ispettore || r.creatoDa || r.assegnatari}</td>
                <td style={td}>{r.numeroCommentiPrg || 0}</td>
                <td style={td}>{r.numeroCommentiIsp || 0}</td>
                <td style={td}>{r.chiDeveAgire}</td>
                <td style={td}>{r.statoRisoluzione}</td>
                <td style={td}>{r.ultimoCommento}</td>
              </tr>
            ))}
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

const inputStyle = {
  width: "100%",
  border: "1px solid #cbd5e1",
  borderRadius: 10,
  padding: 10,
  background: "white",
  boxSizing: "border-box" as const,
};

const labelStyle = {
  display: "block",
  fontWeight: 800,
  fontSize: 15,
  marginBottom: 8,
  color: "#020617",
};

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
  const [error, setError] = useState("");

  const [todoFile, setTodoFile] = useState<File | null>(null);
  const [bcfFiles, setBcfFiles] = useState<File[]>([]);
  const [elencoFile, setElencoFile] = useState<File | null>(null);
  const [reportFile, setReportFile] = useState<File | null>(null);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [progettisti, setProgettisti] = useState("");
  const [ispettori, setIspettori] = useState("");
  const [schedeError, setSchedeError] = useState("");
  const [schedeLoading, setSchedeLoading] = useState(false);

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

  async function generaSchedeIspettive() {
    setSchedeError("");

    if (!todoFile || bcfFiles.length === 0 || !elencoFile || !templateFile) {
      setSchedeError(
        "Carica ToDo XLSX, almeno un BCFZIP/ZIP, Elenco Elaborati XLSX e Template DOCX."
      );
      return;
    }

    setSchedeLoading(true);

    try {
      const fd = new FormData();
      fd.append("todo", todoFile);
      bcfFiles.forEach((f) => fd.append("bcf", f));
      fd.append("elenco", elencoFile);
      fd.append("template", templateFile);

      if (reportFile) {
        fd.append("files", reportFile);
      }

      fd.append("progettisti", progettisti);
      fd.append("ispettori", ispettori);

      const res = await fetch("/api/schede-ispettive", {
        method: "POST",
        body: fd,
      });

      if (!res.ok) {
        let message = "Errore durante la generazione delle schede ispettive.";

        try {
          const data = await res.json();
          message = data.error || message;
        } catch {
          // La risposta potrebbe non essere JSON.
        }

        setSchedeError(message);
        return;
      }

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "SCHEDE_ISPETTIVE_OUTPUT.zip";
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (err: any) {
      setSchedeError(err?.message || "Errore imprevisto durante la generazione.");
    } finally {
      setSchedeLoading(false);
    }
  }

  const enrichedRows = useMemo(() => {
    return rows.map((r) => ({
      ...r,
      tipologiaNcOss: r.tipologiaNcOss || r.tipologiaDocumento || "",
      tipologia: r.tipologiaNcOss || r.tipologiaDocumento || "",
      elaboratoKey: getElaboratoKey(r),
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

    if (selection.type === "tipologia") {
      return enrichedRows.filter(
        (r) =>
          (r.tipo === "NC" || r.tipo === "OSS" || r.tipo === "Da NC a OSS") &&
          (r.tipologiaNcOss || r.tipologiaDocumento) === selection.value
      );
    }

    return [];
  }, [enrichedRows, selection, rilieviNCOSS]);

  const discipline: any = {};
  const tipologie: any = {};
  const esiti: any = {};
  const rilieviPerElaborato: any = {};

  enrichedRows.forEach((r) => {
    const d = r.disciplina || "Non assegnata";
    discipline[d] = (discipline[d] || 0) + 1;

    if (r.tipo === "NC" || r.tipo === "OSS" || r.tipo === "Da NC a OSS") {
      const t = r.tipologiaNcOss || r.tipologiaDocumento || "Altre";
      tipologie[t] = (tipologie[t] || 0) + 1;
    }

    const e = r.tipo || "Esito mancante";
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
    .map(([label, value]) => ({ label, value }))
    .sort((a: any, b: any) => b.value - a.value);

  const tipologieData = Object.entries(tipologie)
    .map(([label, value]) => ({ label, value }))
    .sort((a: any, b: any) => {
      const getOrder = (label: string) => {
        const match = String(label).match(/^(\d+)\./);
        if (match) return Number(match[1]);
        if (label === "Altre") return 999;
        return 998;
      };

      return getOrder(a.label) - getOrder(b.label);
    });

  const esitiData = Object.entries(esiti)
    .map(([label, value]) => ({ label, value }))
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
              Leggi XLSX + BCFZIP
            </button>

            {files.length > 0 && (
              <div style={{ marginTop: 10, color: "#64748b", fontSize: 12 }}>
                File selezionati: {files.map((f) => f.name).join(", ")}
              </div>
            )}
          </Card>
        </div>
      </div>

      <section style={{ marginTop: 24 }}>
        <Card>
          <h1 style={{ marginTop: 0, marginBottom: 10, fontSize: 34 }}>
            Generatore Schede Ispettive
          </h1>

          <p style={{ marginTop: 0, color: "#0f172a", fontSize: 17 }}>
            Modulo per generare schede ispettive da ToDo Trimble, uno o più BCFZIP,
            elenco elaborati, file elaborati progettisti e template Word.
          </p>

          <div style={{ display: "grid", gap: 16, maxWidth: 860 }}>
            <div>
              <label style={labelStyle}>ToDo Trimble XLSX</label>
              <input
                type="file"
                accept=".xlsx,.xls"
                style={inputStyle}
                onChange={(e) => setTodoFile(e.target.files?.[0] || null)}
              />
            </div>

            <div>
              <label style={labelStyle}>File BCFZIP / ZIP - anche multipli</label>
              <input
                type="file"
                multiple
                accept=".bcfzip,.zip"
                style={inputStyle}
                onChange={(e) => setBcfFiles(Array.from(e.target.files || []))}
              />
            </div>

            <div>
              <label style={labelStyle}>Elenco Elaborati XLSX</label>
              <input
                type="file"
                accept=".xlsx,.xls"
                style={inputStyle}
                onChange={(e) => setElencoFile(e.target.files?.[0] || null)}
              />
            </div>

            <div>
              <label style={labelStyle}>Report_Completo XLSX</label>
              <input
                type="file"
                accept=".xlsx,.xls"
                style={inputStyle}
                onChange={(e) => setReportFile(e.target.files?.[0] || null)}
              />
            </div>

            <div>
              <label style={labelStyle}>Template Scheda Ispettiva DOCX</label>
              <input
                type="file"
                accept=".docx"
                style={inputStyle}
                onChange={(e) => setTemplateFile(e.target.files?.[0] || null)}
              />
            </div>

            <div>
              <label style={labelStyle}>Progettisti</label>
              <textarea
                name="progettisti"
                rows={4}
                value={progettisti}
                onChange={(e) => setProgettisti(e.target.value)}
                placeholder={"Giuseppe Pizzi\nMario Rossi"}
                style={{
                  ...inputStyle,
                  resize: "vertical",
                  fontFamily: "Arial, sans-serif",
                }}
              />
              <div style={{ marginTop: 6, color: "#64748b", fontSize: 12 }}>
                Inserire i nomi o account esattamente come compaiono nei commenti ToDo Trimble.
                Uno per riga oppure separati da virgola.
              </div>
            </div>

            <div>
              <label style={labelStyle}>Ispettori</label>
              <textarea
                name="ispettori"
                rows={4}
                value={ispettori}
                onChange={(e) => setIspettori(e.target.value)}
                placeholder={"Luca Bianchi\nAnna Verdi"}
                style={{
                  ...inputStyle,
                  resize: "vertical",
                  fontFamily: "Arial, sans-serif",
                }}
              />
              <div style={{ marginTop: 6, color: "#64748b", fontSize: 12 }}>
                I commenti degli autori presenti in questa lista verranno stampati come riscontro ispettore.
              </div>
            </div>

            {schedeError && (
              <div
                style={{
                  background: "#fee2e2",
                  color: "#991b1b",
                  border: "1px solid #fecaca",
                  borderRadius: 12,
                  padding: 12,
                }}
              >
                {schedeError}
              </div>
            )}

            <button
              onClick={generaSchedeIspettive}
              disabled={schedeLoading}
              style={{
                width: "100%",
                padding: 18,
                background: "#0f172a",
                color: "white",
                borderRadius: 12,
                border: "none",
                cursor: schedeLoading ? "not-allowed" : "pointer",
                fontWeight: 800,
                fontSize: 17,
              }}
            >
              {schedeLoading ? "Generazione in corso..." : "Genera Schede Ispettive"}
            </button>
          </div>
        </Card>
      </section>

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
          gridTemplateColumns: "repeat(7, 1fr)",
          gap: 12,
          marginTop: 24,
          marginBottom: 24,
        }}
      >
        {[
          {
            title: "Nota di Ricezione Elaborati",
            subtitle: "Modulo operativo attivo",
            url: "https://verifica-elaborati-production.up.railway.app",
            active: true,
            external: true,
          },
          { title: "Verifiche preliminari", subtitle: "Coming soon", url: "", active: false, external: false },
          {
            title: "Schede ispettive",
            subtitle: "Modulo operativo attivo",
            url: "/schede-ispettive",
            active: true,
            external: false,
          },
          {
            title: "Controllo ToDo ispettori",
            subtitle: "Verifica Title, Tags, disciplina e codici",
            url: "/controllo-todo-ispettori",
            active: true,
            external: false,
          },
          { title: "Rapporto intermedio", subtitle: "Coming soon", url: "", active: false, external: false },
          { title: "Rapporto conclusivo", subtitle: "Coming soon", url: "", active: false, external: false },
          { title: "Dashboard PM", subtitle: "Coming soon", url: "", active: false, external: false },
        ].map((m) => (
          <Card
            key={m.title}
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

            {m.active && (
              <div style={{ marginTop: 12, fontSize: 12, fontWeight: 700, color: "#0284c7" }}>
                Apri modulo →
              </div>
            )}
          </Card>
        ))}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(6, 1fr)", gap: 12, marginBottom: 12 }}>
        <KPI title="Elaborati totali" value={elaboratiTot} onClick={() => setSelection({ type: "kpi", value: "totali", label: "KPI", valueLabel: "Elaborati totali" })} />
        <KPI title="Elaborati con NC" value={elaboratiNC} onClick={() => setSelection({ type: "kpi", value: "nc", label: "KPI", valueLabel: "Elaborati con NC" })} />
        <KPI title="Elaborati con OSS" value={elaboratiOSS} onClick={() => setSelection({ type: "kpi", value: "oss", label: "KPI", valueLabel: "Elaborati con OSS" })} />
        <KPI title="Elaborati senza rilievi" value={elaboratiOK} onClick={() => setSelection({ type: "kpi", value: "nessun", label: "KPI", valueLabel: "Elaborati senza rilievi" })} />
        <KPI title="Totale NC" value={totaleNC} onClick={() => setSelection({ type: "tipo", value: "NC", label: "Esito" })} />
        <KPI title="Totale OSS" value={totaleOSS} onClick={() => setSelection({ type: "tipo", value: "OSS", label: "Esito" })} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 24 }}>
        <KPI title="Totale NC + OSS" value={totaleNCOSS} onClick={() => setSelection({ type: "kpi", value: "totali", label: "KPI", valueLabel: "Tutti i rilievi" })} />
        <KPI title="In attesa di riscontro dell'ispettore" value={daVerificareISP} subtitle="Ultimo commento PRG" onClick={() => setSelection({ type: "kpi", value: "da-verificare-isp", label: "KPI", valueLabel: "Da verificare ISP" })} />
        <KPI title="In attesa di risposta del progettista" value={daRisponderePRG} subtitle="Nessun PRG o ultimo ISP" onClick={() => setSelection({ type: "kpi", value: "da-rispondere-prg", label: "KPI", valueLabel: "Da rispondere PRG" })} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 16 }}>
        <BarList title="Elaborati per disciplina" data={disciplineData} activeKey={selection?.type === "disciplina" ? selection.value : ""} onClick={(value: string) => setSelection({ type: "disciplina", value, label: "Disciplina" })} />
        <BarList title="NC / OSS / Nessun rilievo" data={esitiData} activeKey={selection?.type === "tipo" ? selection.value : ""} onClick={(value: string) => setSelection({ type: "tipo", value, label: "Esito" })} />
        <BarList title="Tipologie NC / OSS" data={tipologieData} activeKey={selection?.type === "tipologia" ? selection.value : ""} onClick={(value: string) => setSelection({ type: "tipologia", value, label: "Tipologia" })} />
        <BarList title="NC / OSS per elaborato" data={rilieviPerElaboratoData} activeKey={selection?.type === "elaborato" ? selection.value : ""} onClick={(value: string) => setSelection({ type: "elaborato", value, label: "Elaborato" })} />
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
