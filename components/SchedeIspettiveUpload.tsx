"use client";

import React, { useState } from "react";

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

export default function SchedeIspettiveUpload() {
  React.useEffect(() => {
    const isAuthenticated = localStorage.getItem("nexcommon_verify_auth");

    if (isAuthenticated !== "true") {
      window.location.href = "/login";
    }
  }, []);

  const [todoFile, setTodoFile] = useState<File | null>(null);
  const [bcfFiles, setBcfFiles] = useState<File[]>([]);
  const [elencoFile, setElencoFile] = useState<File | null>(null);
  const [reportFile, setReportFile] = useState<File | null>(null);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [progettisti, setProgettisti] = useState("");
  const [ispettori, setIspettori] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  async function generaSchedeIspettive() {
    setError("");

    if (!todoFile || bcfFiles.length === 0 || !elencoFile || !templateFile) {
      setError(
        "Carica ToDo XLSX, almeno un BCFZIP/ZIP, Elenco Elaborati XLSX e Template DOCX."
      );
      return;
    }

    setLoading(true);

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

      const res = await fetch("/api/genera-schede", {
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

        setError(message);
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
      setError(err?.message || "Errore imprevisto durante la generazione.");
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
      <a
        href="/"
        style={{
          display: "inline-block",
          marginBottom: 28,
          color: "#0284c7",
          fontWeight: 700,
          textDecoration: "none",
        }}
      >
        ← Torna alla dashboard
      </a>

      <section
        style={{
          background: "white",
          borderRadius: 16,
          padding: 24,
          border: "1px solid #e2e8f0",
          boxShadow: "0 6px 18px rgba(15,23,42,.06)",
          maxWidth: 980,
        }}
      >
        <h1 style={{ marginTop: 0, marginBottom: 10, fontSize: 34 }}>
          Generatore Schede Ispettive
        </h1>

        <p style={{ marginTop: 0, color: "#0f172a", fontSize: 17 }}>
          Modulo per generare schede ispettive da ToDo Trimble, uno o più BCFZIP,
          elenco elaborati, file elaborati progettisti e template Word.
        </p>

        <div style={{ display: "grid", gap: 16 }}>
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

          {error && (
            <div
              style={{
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

          <button
            onClick={generaSchedeIspettive}
            disabled={loading}
            style={{
              width: "100%",
              padding: 18,
              background: "#0f172a",
              color: "white",
              borderRadius: 12,
              border: "none",
              cursor: loading ? "not-allowed" : "pointer",
              fontWeight: 800,
              fontSize: 17,
            }}
          >
            {loading ? "Generazione in corso..." : "Genera Schede Ispettive"}
          </button>
        </div>
      </section>
    </main>
  );
}
