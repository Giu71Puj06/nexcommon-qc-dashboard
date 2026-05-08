"use client";

import React, { useState } from "react";
import Link from "next/link";

export default function SchedeIspettivePage() {
  const [todo, setTodo] = useState<File | null>(null);
  const [bcfFiles, setBcfFiles] = useState<File[]>([]);
  const [elenco, setElenco] = useState<File | null>(null);
  const [filesXlsx, setFilesXlsx] = useState<File | null>(null);
  const [template, setTemplate] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);

  async function generaSchede() {
    if (!todo || bcfFiles.length === 0 || !elenco || !filesXlsx || !template) {
      alert(
        "Carica ToDo XLSX, almeno un BCFZIP, Elenco Elaborati XLSX, Report_Completo XLSX e Template DOCX."
      );
      return;
    }

    setLoading(true);

    try {
      const fd = new FormData();

      fd.append("todo", todo);
      fd.append("elenco", elenco);
      fd.append("files", filesXlsx);
      fd.append("report", filesXlsx);
      fd.append("template", template);

      bcfFiles.forEach((file) => {
        fd.append("bcf", file);
        fd.append("bcfzip", file);
      });

      const res = await fetch("/api/genera-schede", {
        method: "POST",
        body: fd,
      });

      if (!res.ok) {
        const errorText = await res.text().catch(() => "");
        console.error("Errore API genera-schede:", errorText);
        alert("Errore durante la generazione delle schede.");
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
    } catch (err) {
      console.error(err);
      alert("Errore durante la generazione delle schede.");
    } finally {
      setLoading(false);
    }
  }

  return (
    <main
      style={{
        minHeight: "100vh",
        padding: 30,
        fontFamily: "Arial, sans-serif",
        background: "#f1f5f9",
        color: "#0f172a",
      }}
    >
      <div style={{ maxWidth: 980 }}>
        <Link
          href="/"
          style={{
            display: "inline-block",
            marginBottom: 24,
            color: "#0284c7",
            fontWeight: 700,
            textDecoration: "none",
          }}
        >
          ← Torna alla dashboard
        </Link>

        <h1 style={{ fontSize: 42, marginBottom: 16 }}>
          Generatore Schede Ispettive
        </h1>

        <p style={{ fontSize: 20, marginBottom: 24 }}>
          Modulo per generare schede ispettive da ToDo Trimble, uno o più
          BCFZIP, elenco elaborati, file elaborati progettisti e template Word.
        </p>

        <div style={{ display: "grid", gap: 18, maxWidth: 820 }}>
          <label>
            <b>ToDo Trimble XLSX</b>
            <input
              type="file"
              accept=".xlsx"
              onChange={(e) => setTodo(e.target.files?.[0] || null)}
              style={inputStyle}
            />
          </label>

          <label>
            <b>File BCFZIP / ZIP - anche multipli</b>
            <input
              type="file"
              accept=".bcfzip,.zip"
              multiple
              onChange={(e) => {
                if (!e.target.files) {
                  setBcfFiles([]);
                  return;
                }

                setBcfFiles(Array.from(e.target.files));
              }}
              style={inputStyle}
            />
          </label>

          {bcfFiles.length > 0 && (
            <div style={{ fontSize: 14, color: "#475569" }}>
              File BCF caricati: {bcfFiles.length}
              <ul style={{ marginTop: 6 }}>
                {bcfFiles.map((file, index) => (
                  <li key={`${file.name}-${index}`}>{file.name}</li>
                ))}
              </ul>
            </div>
          )}

          <label>
            <b>Elenco Elaborati XLSX</b>
            <input
              type="file"
              accept=".xlsx"
              onChange={(e) => setElenco(e.target.files?.[0] || null)}
              style={inputStyle}
            />
          </label>

          <label>
            <b>Report_Completo XLSX</b>
            <input
              type="file"
              accept=".xlsx"
              onChange={(e) => setFilesXlsx(e.target.files?.[0] || null)}
              style={inputStyle}
            />
          </label>

          <label>
            <b>Template Scheda Ispettiva DOCX</b>
            <input
              type="file"
              accept=".docx"
              onChange={(e) => setTemplate(e.target.files?.[0] || null)}
              style={inputStyle}
            />
          </label>

          <button
            onClick={generaSchede}
            disabled={loading}
            style={{
              padding: 16,
              background: loading ? "#334155" : "#0f172a",
              color: "white",
              border: "none",
              borderRadius: 12,
              cursor: loading ? "not-allowed" : "pointer",
              fontSize: 18,
              fontWeight: 700,
            }}
          >
            {loading ? "Generazione in corso..." : "Genera Schede Ispettive"}
          </button>
        </div>
      </div>
    </main>
  );
}

const inputStyle: React.CSSProperties = {
  display: "block",
  width: "100%",
  marginTop: 8,
  padding: 12,
  border: "1px solid #cbd5e1",
  borderRadius: 10,
  background: "white",
};