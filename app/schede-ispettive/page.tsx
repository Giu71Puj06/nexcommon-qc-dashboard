"use client";

import React, { useState } from "react";

export default function SchedeIspettivePage() {
  const [todo, setTodo] = useState<File | null>(null);
  const [bcf, setBcf] = useState<File | null>(null);
  const [elenco, setElenco] = useState<File | null>(null);
  const [filesXlsx, setFilesXlsx] = useState<File | null>(null);
  const [template, setTemplate] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);

  async function generaSchede() {
    if (!todo || !bcf || !elenco || !template) {
      alert(
        "Carica ToDo XLSX, BCFZIP, Elenco Elaborati XLSX e Template DOCX."
      );
      return;
    }

    setLoading(true);

    const fd = new FormData();
    fd.append("todo", todo);
    fd.append("bcf", bcf);
    fd.append("elenco", elenco);
    fd.append("template", template);

    if (filesXlsx) {
      fd.append("files", filesXlsx);
    }

    const res = await fetch("/api/genera-schede", {
      method: "POST",
      body: fd,
    });

    if (!res.ok) {
      alert("Errore durante la generazione delle schede.");
      setLoading(false);
      return;
    }

    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "SCHEDE_ISPETTIVE_OUTPUT.zip";
    a.click();

    window.URL.revokeObjectURL(url);
    setLoading(false);
  }

  return (
    <main style={{ padding: 30, fontFamily: "Arial, sans-serif" }}>
      <h1>Generatore Schede Ispettive</h1>

      <p>
        Modulo per generare schede ispettive da ToDo Trimble, BCFZIP,
        elenco elaborati, file elaborati progettisti e template Word.
      </p>

      <div style={{ display: "grid", gap: 16, maxWidth: 700 }}>
        <label>
          <b>ToDo Trimble XLSX</b>
          <input
            type="file"
            accept=".xlsx"
            onChange={(e) => setTodo(e.target.files?.[0] || null)}
          />
        </label>

        <label>
          <b>File BCFZIP</b>
          <input
            type="file"
            accept=".bcfzip,.zip"
            onChange={(e) => setBcf(e.target.files?.[0] || null)}
          />
        </label>

        <label>
          <b>Elenco Elaborati XLSX</b>
          <input
            type="file"
            accept=".xlsx"
            onChange={(e) => setElenco(e.target.files?.[0] || null)}
          />
        </label>

        <label>
          <b>Report_Completo XLSX</b>
          <input
            type="file"
            accept=".xlsx"
            onChange={(e) => setFilesXlsx(e.target.files?.[0] || null)}
          />
        </label>

        <label>
          <b>Template Scheda Ispettiva DOCX</b>
          <input
            type="file"
            accept=".docx"
            onChange={(e) => setTemplate(e.target.files?.[0] || null)}
          />
        </label>

        <button
          onClick={generaSchede}
          disabled={loading}
          style={{
            padding: 12,
            background: "#0f172a",
            color: "white",
            border: "none",
            borderRadius: 10,
            cursor: loading ? "not-allowed" : "pointer",
          }}
        >
          {loading ? "Generazione in corso..." : "Genera Schede Ispettive"}
        </button>
      </div>
    </main>
  );
}