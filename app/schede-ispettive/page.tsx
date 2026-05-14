"use client";

import React, { useState } from "react";
import Link from "next/link";

export default function SchedeIspettivePage() {
  const [todo, setTodo] = useState<File | null>(null);
  const [bcfFiles, setBcfFiles] = useState<File[]>([]);
  const [fotoFiles, setFotoFiles] = useState<File[]>([]);
  const [fotoZip, setFotoZip] = useState<File | null>(null);
  const [elenco, setElenco] = useState<File | null>(null);
  const [filesXlsx, setFilesXlsx] = useState<File | null>(null);
  const [template, setTemplate] = useState<File | null>(null);
  const [progettisti, setProgettisti] = useState("");
  const [ispettori, setIspettori] = useState("");
  const [revisioneScheda, setRevisioneScheda] = useState("0");
  const [dataRevisioneScheda, setDataRevisioneScheda] = useState("");
  const [responsabileIts, setResponsabileIts] = useState("");
  const [responsabilePcq, setResponsabilePcq] = useState("");
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
      fd.append("progettisti", progettisti);
      fd.append("ispettori", ispettori);

      fd.append("revisione_scheda", revisioneScheda);
      fd.append("data_revisione_scheda", dataRevisioneScheda);
      fd.append("responsabile_its", responsabileIts);
      fd.append("responsabile_pcq", responsabilePcq);

      bcfFiles.forEach((file) => {
        fd.append("bcf", file);
        fd.append("bcfzip", file);
      });

      fotoFiles.forEach((file) => {
        fd.append("foto", file);
        fd.append("immagini", file);
      });

      if (fotoZip) {
        fd.append("fotoZip", fotoZip);
        fd.append("immaginiZip", fotoZip);
      }

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
          BCFZIP, elenco elaborati, file elaborati progettisti, immagini
          allegate e template Word.
        </p>

        <div style={{ display: "grid", gap: 18, maxWidth: 820 }}>
          <label>
            <b>ToDo Trimble XLSX</b>
            <input
              type="file"
              accept=".xlsx,.xls"
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
            <b>Cartella immagini allegate</b>
            <input
              type="file"
              accept=".png,.jpg,.jpeg,.webp"
              multiple
              // @ts-expect-error webkitdirectory è supportato dai browser Chromium
              webkitdirectory="true"
              onChange={(e) => {
                if (!e.target.files) {
                  setFotoFiles([]);
                  return;
                }

                setFotoFiles(Array.from(e.target.files));
              }}
              style={inputStyle}
            />
            <div style={helpStyle}>
              Carica la cartella sicurezza con immagini tipo
              IT22-026_foto.png. Il sistema le collegherà a TR-26,
              ignorando gli zeri iniziali.
            </div>
          </label>

          {fotoFiles.length > 0 && (
            <div style={{ fontSize: 14, color: "#475569" }}>
              Immagini caricate da cartella: {fotoFiles.length}
              <ul style={{ marginTop: 6 }}>
                {fotoFiles.slice(0, 10).map((file, index) => (
                  <li key={`${file.name}-${index}`}>{file.name}</li>
                ))}
              </ul>
              {fotoFiles.length > 10 && (
                <div>...altre {fotoFiles.length - 10} immagini</div>
              )}
            </div>
          )}

          <label>
            <b>ZIP immagini allegate</b>
            <input
              type="file"
              accept=".zip"
              onChange={(e) => setFotoZip(e.target.files?.[0] || null)}
              style={inputStyle}
            />
            <div style={helpStyle}>
              In alternativa alla cartella, puoi caricare uno ZIP con le
              immagini allegate.
            </div>
          </label>

          <label>
            <b>Elenco Elaborati XLSX</b>
            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={(e) => setElenco(e.target.files?.[0] || null)}
              style={inputStyle}
            />
          </label>

          <label>
            <b>Report_Completo XLSX</b>
            <input
              type="file"
              accept=".xlsx,.xls"
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

          <label>
            <b>Progettisti</b>
            <textarea
              value={progettisti}
              onChange={(e) => setProgettisti(e.target.value)}
              placeholder={"Giuseppe Pizzi\nMario Rossi"}
              rows={4}
              style={textareaStyle}
            />
            <div style={helpStyle}>
              Inserire i nomi/account esattamente come compaiono nei commenti
              ToDo Trimble. Uno per riga oppure separati da virgola.
            </div>
          </label>

          <label>
            <b>Ispettori</b>
            <textarea
              value={ispettori}
              onChange={(e) => setIspettori(e.target.value)}
              placeholder={"Luca Bianchi\nAnna Verdi"}
              rows={4}
              style={textareaStyle}
            />
            <div style={helpStyle}>
              I commenti degli autori presenti in questa lista verranno stampati
              come riscontro ispettore.
            </div>
          </label>

          <div style={sectionStyle}>
            <h2 style={{ margin: "0 0 12px 0", fontSize: 22 }}>
              Dati revisione scheda
            </h2>

            <label>
              <b>Rev. scheda ispettiva</b>
              <input
                type="text"
                value={revisioneScheda}
                onChange={(e) => setRevisioneScheda(e.target.value)}
                placeholder="0"
                style={inputStyle}
              />
              <div style={helpStyle}>
                0 = Prima Emissione - Rilievi; 1 = Seconda emissione -
                Riscontri; 2 = Terza emissione - Riscontri.
              </div>
            </label>

            <label>
              <b>Data revisione scheda</b>
              <input
                type="text"
                value={dataRevisioneScheda}
                onChange={(e) => setDataRevisioneScheda(e.target.value)}
                placeholder="gg/mm/aaaa"
                style={inputStyle}
              />
            </label>

            <label>
              <b>Responsabile tecnico ITS</b>
              <input
                type="text"
                value={responsabileIts}
                onChange={(e) => setResponsabileIts(e.target.value)}
                placeholder="Ing. Nome Cognome"
                style={inputStyle}
              />
            </label>

            <label>
              <b>Responsabile tecnico PCQ</b>
              <input
                type="text"
                value={responsabilePcq}
                onChange={(e) => setResponsabilePcq(e.target.value)}
                placeholder="Compilare solo se previsto"
                style={inputStyle}
              />
              <div style={helpStyle}>
                Campo opzionale: da usare solo nei casi in cui è previsto un
                responsabile tecnico PCQ distinto.
              </div>
            </label>
          </div>

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

const textareaStyle: React.CSSProperties = {
  display: "block",
  width: "100%",
  marginTop: 8,
  padding: 12,
  border: "1px solid #cbd5e1",
  borderRadius: 10,
  background: "white",
  resize: "vertical",
  fontFamily: "Arial, sans-serif",
};

const sectionStyle: React.CSSProperties = {
  display: "grid",
  gap: 18,
  padding: 18,
  border: "1px solid #cbd5e1",
  borderRadius: 12,
  background: "#e2e8f0",
};

const helpStyle: React.CSSProperties = {
  marginTop: 6,
  color: "#64748b",
  fontSize: 12,
};
