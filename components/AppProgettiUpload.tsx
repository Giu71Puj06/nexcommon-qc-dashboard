"use client";

import React, { useMemo, useState } from "react";

function getKpiColor(value: number) {
  if (value >= 80) return "#16a34a";
  if (value >= 50) return "#f59e0b";
  return "#dc2626";
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
      <div
        style={{
          fontSize: 34,
          fontWeight: 800,
          color: colorValue || "#0f172a",
        }}
      >
        {value}
      </div>
      {subtitle && (
        <div style={{ fontSize: 12, color: "#64748b" }}>{subtitle}</div>
      )}
    </Card>
  );
}

function BarList({ title, data, onClick, activeKey }: any) {
  const max = Math.max(...data.map((d: any) => d.value), 1);

  return (
    <Card>
      <h3 style={{ marginTop: 0 }}>{title}</h3>

      {data.length === 0 && (
        <div style={{ color: "#64748b", fontSize: 13 }}>
          Nessun dato disponibile
        </div>
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
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              fontSize: 13,
            }}
          >
            <span>{d.label}</span>
            <b>{d.value}</b>
          </div>

          <div
            style={{
              height: 12,
              background: "#e2e8f0",
              borderRadius: 99,
              overflow: "hidden",
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
              {f.type === "xlsx" &&
                `- Excel letto: ${f.rows || 0} righe`}
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
                <td style={td}>{r.elaborato || r.titolo}</td>
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

export default function AppProgettiUpload() {
  const [files, setFiles] = useState<File[]>([]);
  const [rows, setRows] = useState<any[]>([]);
  const [selection, setSelection] = useState<any>(null);
  const [importedFiles, setImportedFiles] = useState<any[]>([]);
  const [error, setError] = useState("");

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
    }));
  }, [rows]);

  const elaboratiTot = new Set(enrichedRows.map((r) => r.id).filter(Boolean)).size;
  const elaboratiNC = new Set(enrichedRows.filter((r) => r.tipo === "NC").map((r) => r.id)).size;
  const elaboratiOSS = new Set(enrichedRows.filter((r) => r.tipo === "OSS").map((r) => r.id)).size;
  const elaboratiOK = new Set(enrichedRows.filter((r) => r.tipo === "Nessun rilievo").map((r) => r.id)).size;

  const controlliCompleti = enrichedRows.filter((r) => r.controlloIspettoreCompleto).length;
  const controlliIncompleti = enrichedRows.filter((r) => !r.controlloIspettoreCompleto);

  const completezzaControlloIspettori =
    enrichedRows.length > 0
      ? Math.round((controlliCompleti / enrichedRows.length) * 100)
      : 0;

  const rilieviNCOSS = enrichedRows.filter(
    (r) => r.tipo === "NC" || r.tipo === "OSS" || r.tipo === "Da NC a OSS"
  );

  const rilieviConRiscontroPRG = rilieviNCOSS.filter((r) => r.hasPrgComment).length;

  const completezzaRisoluzioneProgettisti =
    rilieviNCOSS.length > 0
      ? Math.round((rilieviConRiscontroPRG / rilieviNCOSS.length) * 100)
      : 0;

  const daVerificareISP = enrichedRows.filter((r) => r.chiDeveAgire === "ISP").length;
  const daRisponderePRG = enrichedRows.filter((r) => r.chiDeveAgire === "PRG").length;

  const filteredRows = useMemo(() => {
    if (!selection) return [];

    if (selection.type === "kpi") {
      if (selection.value === "totali") return enrichedRows;
      if (selection.value === "nc") return enrichedRows.filter((r) => r.tipo === "NC");
      if (selection.value === "oss") return enrichedRows.filter((r) => r.tipo === "OSS");
      if (selection.value === "nessun") return enrichedRows.filter((r) => r.tipo === "Nessun rilievo");
      if (selection.value === "controllo-incompleto") return controlliIncompleti;
      if (selection.value === "risoluzione-prg") return rilieviNCOSS.filter((r) => r.hasPrgComment);
      if (selection.value === "da-verificare-isp") return enrichedRows.filter((r) => r.chiDeveAgire === "ISP");
      if (selection.value === "da-rispondere-prg") return enrichedRows.filter((r) => r.chiDeveAgire === "PRG");
    }

    if (selection.type === "tipo") return enrichedRows.filter((r) => r.tipo === selection.value);
    if (selection.type === "disciplina") return enrichedRows.filter((r) => r.disciplina === selection.value);
    if (selection.type === "tipologia") {
      return enrichedRows.filter(
        (r) =>
          (r.tipo === "NC" || r.tipo === "OSS" || r.tipo === "Da NC a OSS") &&
          (r.tipologiaNcOss || r.tipologiaDocumento) === selection.value
      );
    }

    return [];
  }, [enrichedRows, selection, controlliIncompleti, rilieviNCOSS]);

  const discipline: any = {};
  const tipologie: any = {};
  const esiti: any = {};

  enrichedRows.forEach((r) => {
    const d = r.disciplina || "Non assegnata";
    discipline[d] = (discipline[d] || 0) + 1;

    if (r.tipo === "NC" || r.tipo === "OSS" || r.tipo === "Da NC a OSS") {
      const t = r.tipologiaNcOss || r.tipologiaDocumento || "Altre";
      tipologie[t] = (tipologie[t] || 0) + 1;
    }

    const e = r.tipo || "Esito mancante";
    esiti[e] = (esiti[e] || 0) + 1;
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
      <div style={{ display: "flex", justifyContent: "space-between", gap: 20 }}>
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <img
              src="/logo_nexcommon.png"
              alt="Nexcommon"
              style={{ height: 34, objectFit: "contain" }}
            />
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
          gridTemplateColumns: "repeat(6, 1fr)",
          gap: 12,
          marginTop: 24,
          marginBottom: 24,
        }}
      >
        {[
          "Nota di trascrizione",
          "Verifiche preliminari",
          "Schede ispettive",
          "Rapporto intermedio",
          "Rapporto conclusivo",
          "Dashboard PM",
        ].map((m) => (
          <Card key={m}>
            <b>{m}</b>
            <div style={{ marginTop: 6, color: "#64748b", fontSize: 13 }}>
              Modulo operativo
            </div>
          </Card>
        ))}
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12, marginBottom: 12 }}>
        <KPI title="Elaborati totali" value={elaboratiTot} onClick={() => setSelection({ type: "kpi", value: "totali", label: "KPI", valueLabel: "Elaborati totali" })} />
        <KPI title="Elaborati con NC" value={elaboratiNC} onClick={() => setSelection({ type: "kpi", value: "nc", label: "KPI", valueLabel: "Elaborati con NC" })} />
        <KPI title="Elaborati con OSS" value={elaboratiOSS} onClick={() => setSelection({ type: "kpi", value: "oss", label: "KPI", valueLabel: "Elaborati con OSS" })} />
        <KPI title="Elaborati senza rilievi" value={elaboratiOK} onClick={() => setSelection({ type: "kpi", value: "nessun", label: "KPI", valueLabel: "Elaborati senza rilievi" })} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12, marginBottom: 24 }}>
        <KPI
          title="Completezza controllo ispettori"
          value={`${completezzaControlloIspettori}%`}
          colorValue={getKpiColor(completezzaControlloIspettori)}
          subtitle={`${controlliCompleti}/${enrichedRows.length} controlli completi`}
          active={selection?.value === "controllo-incompleto"}
          onClick={() =>
            setSelection({
              type: "kpi",
              value: "controllo-incompleto",
              label: "KPI",
              valueLabel: "Controlli ispettori incompleti",
            })
          }
        />

        <KPI
          title="Completezza risoluzione progettisti"
          value={`${completezzaRisoluzioneProgettisti}%`}
          colorValue={getKpiColor(completezzaRisoluzioneProgettisti)}
          subtitle={`${rilieviConRiscontroPRG}/${rilieviNCOSS.length} NC/OSS con riscontro PRG`}
          onClick={() =>
            setSelection({
              type: "kpi",
              value: "risoluzione-prg",
              label: "KPI",
              valueLabel: "Rilievi con riscontro PRG",
            })
          }
        />

       <KPI
  title="In attesa di riscontro dell'ispettore"
  value={daVerificareISP}
  subtitle="Ultimo commento PRG"
          onClick={() =>
            setSelection({
              type: "kpi",
              value: "da-verificare-isp",
              label: "KPI",
              valueLabel: "Da verificare ISP",
            })
          }
        />

        <KPI
  title="In attesa di risposta del progettista"
  value={daRisponderePRG}
  subtitle="Nessun PRG o ultimo ISP"
          onClick={() =>
            setSelection({
              type: "kpi",
              value: "da-rispondere-prg",
              label: "KPI",
              valueLabel: "Da rispondere PRG",
            })
          }
        />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16 }}>
        <BarList
          title="Elaborati per disciplina"
          data={disciplineData}
          activeKey={selection?.type === "disciplina" ? selection.value : ""}
          onClick={(value: string) =>
            setSelection({ type: "disciplina", value, label: "Disciplina" })
          }
        />

        <BarList
          title="NC / OSS / Nessun rilievo"
          data={esitiData}
          activeKey={selection?.type === "tipo" ? selection.value : ""}
          onClick={(value: string) =>
            setSelection({ type: "tipo", value, label: "Esito" })
          }
        />

        <BarList
          title="Tipologie NC / OSS"
          data={tipologieData}
          activeKey={selection?.type === "tipologia" ? selection.value : ""}
          onClick={(value: string) =>
            setSelection({ type: "tipologia", value, label: "Tipologia" })
          }
        />
      </div>

            <div style={{ marginTop: 24 }}>
        <DetailPanel
          title={selectionTitle}
          rows={filteredRows}
          onReset={() => setSelection(null)}
        />
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
