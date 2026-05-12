"use client";

import React, { useState } from "react";

export default function Page() {
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<any>(null);
  const [error, setError] = useState("");

  async function analizzaElaborati() {
    if (!file) return;

    setLoading(true);
    setError("");
    setResult(null);

    try {
      const fd = new FormData();
      fd.append("file", file);

      const res = await fetch("/api/verifiche-preliminari", {
        method: "POST",
        body: fd,
      });

      const data = await res.json();

      if (!res.ok || data.success === false) {
        throw new Error(data.error || "Errore analisi elaborati");
      }

      setResult(data);
    } catch (err: any) {
      setError(err.message || "Errore imprevisto");
    } finally {
      setLoading(false);
    }
  }

  return (
    <main style={{ padding: 30, background: "#f1f5f9", minHeight: "100vh", fontFamily: "Arial, sans-serif" }}>
      <button
        onClick={() => (window.location.href = "/")}
        style={{ background: "#0f172a", color: "white", border: "none", borderRadius: 10, padding: "10px 16px", cursor: "pointer", fontWeight: 700 }}
      >
        Torna dashboard
      </button>

      <h1>Verifiche preliminari</h1>
      <p style={{ color: "#64748b" }}>Analisi automatica dell’elenco elaborati per la scheda verifiche preliminari.</p>

      <div style={{ background: "white", borderRadius: 16, padding: 24, marginTop: 20 }}>
        <h2>Upload elenco elaborati</h2>

        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={(e) => setFile(e.target.files?.[0] || null)}
        />

        <br />

        <button
          onClick={analizzaElaborati}
          disabled={!file || loading}
          style={{ marginTop: 16, background: "#0284c7", color: "white", border: "none", borderRadius: 10, padding: "12px 18px", cursor: !file || loading ? "not-allowed" : "pointer", fontWeight: 700 }}
        >
          {loading ? "Analisi in corso..." : "Analizza elaborati"}
        </button>

        {error && <div style={{ marginTop: 16, color: "#991b1b" }}>{error}</div>}
      </div>

      {result && (
        <div style={{ background: "white", borderRadius: 16, padding: 24, marginTop: 24 }}>
          <h2>Risultato analisi</h2>
          <p>Elaborati letti: <b>{result.count}</b></p>

          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr>
                <th style={th}>Codice</th>
                <th style={th}>Titolo</th>
                <th style={th}>Disciplina</th>
              </tr>
            </thead>
            <tbody>
              {(result.elaborati || []).map((e: any, i: number) => (
                <tr key={i}>
                  <td style={td}>{e.codice}</td>
                  <td style={td}>{e.titolo}</td>
                  <td style={td}>{e.disciplina}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </main>
  );
}

const th = {
  border: "1px solid #e2e8f0",
  padding: 8,
  textAlign: "left" as const,
  background: "#f8fafc",
};

const td = {
  border: "1px solid #e2e8f0",
  padding: 8,
};
