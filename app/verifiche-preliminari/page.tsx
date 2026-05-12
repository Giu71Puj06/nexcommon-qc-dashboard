"use client";

import React, { useState } from "react";

export default function Page() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [elaboratiFiles, setElaboratiFiles] = useState<File[]>([]);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<any>(null);
  const [error, setError] = useState("");

  async function analizzaVerifichePreliminari() {
    if (!excelFile || !templateFile || elaboratiFiles.length === 0) {
      setError("Carica Excel, template Word e cartella elaborati PDF.");
      return;
    }

    setLoading(true);
    setError("");
    setResult(null);

    try {
      const fd = new FormData();

      fd.append("excel", excelFile);
      fd.append("template", templateFile);

      elaboratiFiles.forEach((file) => {
        fd.append("elaborati", file);
        fd.append("paths", (file as any).webkitRelativePath || file.name);
      });

      const res = await fetch("/api/verifiche-preliminari", {
        method: "POST",
        body: fd,
      });

      const data = await res.json();

      if (!res.ok || data.success === false) {
        throw new Error(data.error || "Errore analisi verifiche preliminari");
      }

      setResult(data);
    } catch (err: any) {
      setError(err.message || "Errore imprevisto");
    } finally {
      setLoading(false);
    }
  }

  return (
    <main
      style={{
        padding: 30,
        background: "#f1f5f9",
        minHeight: "100vh",
        fontFamily: "Arial, sans-serif",
      }}
    >
      <button
        onClick={() => (window.location.href = "/")}
        style={{
          background: "#0f172a",
          color: "white",
          border: "none",
          borderRadius: 10,
          padding: "10px 16px",
          cursor: "pointer",
          fontWeight: 700,
        }}
      >
        Torna dashboard
      </button>

      <h1>Verifiche preliminari</h1>

      <p style={{ color: "#64748b" }}>
        Carica elenco elaborati, template Word e cartella PDF completa con eventuali sottocartelle.
      </p>

      <div style={card}>
        <h2>1. Elenco elaborati Excel</h2>
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={(e) => setExcelFile(e.target.files?.[0] || null)}
        />
        {excelFile && <p>File selezionato: <b>{excelFile.name}</b></p>}
      </div>

      <div style={card}>
        <h2>2. Template Word verifiche preliminari</h2>
        <input
          type="file"
          accept=".docx"
          onChange={(e) => setTemplateFile(e.target.files?.[0] || null)}
        />
        {templateFile && <p>File selezionato: <b>{templateFile.name}</b></p>}
      </div>

      <div style={card}>
        <h2>3. Cartella elaborati PDF</h2>
        <input
          type="file"
          multiple
          // @ts-ignore
          webkitdirectory="true"
          onChange={(e) => setElaboratiFiles(Array.from(e.target.files || []))}
        />

        {elaboratiFiles.length > 0 && (
          <div style={{ marginTop: 12 }}>
            <b>File caricati: {elaboratiFiles.length}</b>
            <div style={{ marginTop: 8, maxHeight: 160, overflowY: "auto", fontSize: 12, color: "#64748b" }}>
              {elaboratiFiles.slice(0, 30).map((file: any, i) => (
                <div key={i}>{file.webkitRelativePath || file.name}</div>
              ))}
              {elaboratiFiles.length > 30 && <div>... altri file non mostrati</div>}
            </div>
          </div>
        )}
      </div>

      <button
        onClick={analizzaVerifichePreliminari}
        disabled={loading}
        style={{
          marginTop: 20,
          background: "#0284c7",
          color: "white",
          border: "none",
          borderRadius: 10,
          padding: "14px 22px",
          cursor: loading ? "not-allowed" : "pointer",
          fontWeight: 700,
        }}
      >
        {loading ? "Analisi in corso..." : "Analizza verifiche preliminari"}
      </button>

      {error && (
        <div style={{ marginTop: 20, color: "#991b1b", background: "#fee2e2", padding: 12, borderRadius: 10 }}>
          {error}
        </div>
      )}

      {result && (
        <div style={card}>
          <h2>Risultato analisi</h2>
          <p>Elaborati da Excel: <b>{result.count || 0}</b></p>
          <p>File PDF/cartella ricevuti: <b>{result.pdfCount || elaboratiFiles.length}</b></p>

          {result.elaborati && (
            <table style={{ width: "100%", borderCollapse: "collapse", marginTop: 16 }}>
              <thead>
                <tr>
                  <th style={th}>Codice</th>
                  <th style={th}>Titolo</th>
                  <th style={th}>Disciplina</th>
                </tr>
              </thead>
              <tbody>
                {result.elaborati.map((e: any, i: number) => (
                  <tr key={i}>
                    <td style={td}>{e.codice}</td>
                    <td style={td}>{e.titolo}</td>
                    <td style={td}>{e.disciplina}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      )}
    </main>
  );
}

const card = {
  background: "white",
  borderRadius: 16,
  padding: 24,
  marginTop: 20,
  boxShadow: "0 6px 18px rgba(15,23,42,.06)",
};

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
