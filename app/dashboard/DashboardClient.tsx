"use client";

import React, { useMemo, useState } from "react";

/* ---------- UI ---------- */

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
      }}
    >
      {children}
    </div>
  );
}

function KPI({ title, value, onClick, active }: any) {
  return (
    <Card onClick={onClick} active={active}>
      <div style={{ fontSize: 13, color: "#64748b" }}>{title}</div>
      <div style={{ fontSize: 34, fontWeight: 800 }}>{value}</div>
    </Card>
  );
}

/* ---------- CLASSIFICAZIONE ---------- */

function classify(desc = "", title = "") {
  const d = `${desc} ${title}`.toLowerCase();

  if (d.includes("interferenz")) return "Interferenze";
  if (d.includes("mancant")) return "Informazioni mancanti";
  if (d.includes("quota")) return "Quote / dimensioni";
  if (d.includes("armatur")) return "Armature";
  if (d.includes("sicurezza")) return "Sicurezza";
  if (d.includes("impiant")) return "Impianti";

  return "Altro";
}

/* ---------- COMPONENT ---------- */

export default function DashboardClient() {
  const [rows, setRows] = useState<any[]>([]);
  const [files, setFiles] = useState<File[]>([]);
  const [selection, setSelection] = useState<any>(null);

  async function generaDashboard() {
    const fd = new FormData();
    files.forEach((f) => fd.append("files", f));

    const res = await fetch("/api/parse", {
      method: "POST",
      body: fd,
    });

    const data = await res.json();
    setRows(data.rows || []);
    setSelection(null);
  }

  const enriched = useMemo(
    () =>
      rows.map((r) => ({
        ...r,
        tipologia: classify(r.descrizione, r.titolo),
      })),
    [rows]
  );

  /* ---------- KPI ---------- */

  const tot = new Set(enriched.map((r) => r.id)).size;
  const nc = new Set(enriched.filter((r) => r.tipo === "NC").map((r) => r.id)).size;
  const oss = new Set(enriched.filter((r) => r.tipo === "OSS").map((r) => r.id)).size;
  const ok = new Set(
    enriched.filter((r) => r.tipo === "Nessun rilievo").map((r) => r.id)
  ).size;

  /* ---------- AGGREGA ---------- */

  const discipline: any = {};
  const tipologie: any = {};
  const esiti: any = {};

  enriched.forEach((r) => {
    discipline[r.disciplina || "Non assegnata"] =
      (discipline[r.disciplina] || 0) + 1;

    esiti[r.tipo] = (esiti[r.tipo] || 0) + 1;

    if (r.tipo === "NC" || r.tipo === "OSS") {
      tipologie[r.tipologia] = (tipologie[r.tipologia] || 0) + 1;
    }
  });

  /* ---------- FILTER ---------- */

  const filtered = useMemo(() => {
    if (!selection) return enriched;

    if (selection.type === "kpi") {
      if (selection.value === "nc") return enriched.filter((r) => r.tipo === "NC");
      if (selection.value === "oss") return enriched.filter((r) => r.tipo === "OSS");
      if (selection.value === "ok")
        return enriched.filter((r) => r.tipo === "Nessun rilievo");
    }

    if (selection.type === "disciplina")
      return enriched.filter((r) => r.disciplina === selection.value);

    if (selection.type === "tipologia")
      return enriched.filter((r) => r.tipologia === selection.value);

    return enriched;
  }, [selection, enriched]);

  /* ---------- UI ---------- */

  return (
    <main style={{ padding: 30, background: "#f1f5f9", minHeight: "100vh" }}>
      {/* HEADER */}
      <div style={{ display: "flex", justifyContent: "space-between" }}>
        <div>
          {/* Nexcommon */}
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <img src="/logo_nexcommon.png" style={{ height: 32 }} />
            <div style={{ fontSize: 13, color: "#64748b" }}>
              Piattaforma creata da Nexcommon S.r.l.
            </div>
          </div>

          {/* CLIENTE */}
          <div style={{ marginTop: 20, display: "flex", alignItems: "center", gap: 20 }}>
            <img
              src="/logo_its.png"
              style={{ height: 60, background: "#0f172a", padding: 8, borderRadius: 8 }}
            />
            <div>
              <h1 style={{ margin: 0, fontSize: 30 }}>
                ITS Controlli Tecnici S.p.A.
              </h1>
              <div style={{ color: "#64748b" }}>
                Dashboard verifiche elaborati
              </div>
            </div>
          </div>
        </div>

        {/* UPLOAD */}
        <Card>
          <input type="file" multiple onChange={(e) => setFiles(Array.from(e.target.files || []))} />
          <button
            onClick={generaDashboard}
            style={{
              marginTop: 10,
              width: "100%",
              padding: 10,
              background: "#0f172a",
              color: "white",
              borderRadius: 10,
            }}
          >
            Aggiorna dashboard
          </button>
        </Card>
      </div>

      {/* KPI */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 12, marginTop: 30 }}>
        <KPI title="Elaborati totali" value={tot} />
        <KPI title="Elaborati con NC" value={nc} onClick={() => setSelection({ type: "kpi", value: "nc" })} />
        <KPI title="Elaborati con OSS" value={oss} onClick={() => setSelection({ type: "kpi", value: "oss" })} />
        <KPI title="Senza rilievi" value={ok} onClick={() => setSelection({ type: "kpi", value: "ok" })} />
      </div>

      {/* LISTE */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginTop: 30 }}>
        <Card>
          <h3>Elaborati per disciplina</h3>
          {Object.entries(discipline).map(([k, v]: any) => (
            <div key={k} onClick={() => setSelection({ type: "disciplina", value: k })} style={{ cursor: "pointer" }}>
              {k} — <b>{v}</b>
            </div>
          ))}
        </Card>

        <Card>
          <h3>Tipologia NC / OSS</h3>
          {Object.entries(tipologie).map(([k, v]: any) => (
            <div key={k} onClick={() => setSelection({ type: "tipologia", value: k })} style={{ cursor: "pointer" }}>
              {k} — <b>{v}</b>
            </div>
          ))}
        </Card>
      </div>

      {/* DETTAGLIO */}
      {selection && (
        <Card>
          <h2>Dettaglio</h2>
          {filtered.map((r, i) => (
            <div key={i} style={{ borderBottom: "1px solid #eee", padding: 6 }}>
              {r.id} — {r.titolo} — {r.tipo}
            </div>
          ))}
        </Card>
      )}
    </main>
  );
}