"use client";

import React, { useState } from "react";
import Link from "next/link";

export default function SchedeIspettivePage() {
  const [todo, setTodo] = useState<File | null>(null);
  const [bcfFiles, setBcfFiles] = useState<File[]>([]);
  const [elenco, setElenco] = useState<File | null>(null);
  const [filesXlsx, setFilesXlsx] = useState<File | null>(null);
  const [template, setTemplate] = useState<File | null>(null);
  const [schedeEmesseZip, setSchedeEmesseZip] = useState<File | null>(null);
  const [progettisti, setProgettisti] = useState("");
  const [ispettori, setIspettori] = useState("");
  const [revisioneScheda, setRevisioneScheda] = useState("0");
  const [dataRevisioneScheda, setDataRevisioneScheda] = useState("");
  const [dataRev0, setDataRev0] = useState("");
  const [dataRev1, setDataRev1] = useState("");
  const [dataRev2, setDataRev2] = useState("");
  const [dataRev3, setDataRev3] = useState("");
  const [dataRev4, setDataRev4] = useState("");
  const [dateRispostaProgettista, setDateRispostaProgettista] = useState<string[]>(["", "", "", "", ""]);
  const [dateRiscontroIspettore, setDateRiscontroIspettore] = useState<string[]>(["", "", "", "", ""]);
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

    const revNumber = Number(String(revisioneScheda || "0").trim());
    const isEmissioneSuccessiva = Number.isFinite(revNumber) && revNumber > 0;

    if (isEmissioneSuccessiva && !schedeEmesseZip) {
      alert(
        "Per le emissioni successive alla prima devi caricare lo ZIP delle schede ispettive gia emesse nell emissione precedente."
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

      if (schedeEmesseZip) {
        fd.append("schede_emesse_zip", schedeEmesseZip);
      }
      fd.append("progettisti", progettisti);
      fd.append("ispettori", ispettori);
      fd.append("revisione_scheda", revisioneScheda);
      fd.append("data_revisione_scheda", dataRevisioneScheda);
      fd.append("data_rev_0", dataRev0);
      fd.append("data_rev_1", dataRev1);
      fd.append("data_rev_2", dataRev2);
      fd.append("data_rev_3", dataRev3);
      fd.append("data_rev_4", dataRev4);
      fd.append("data_risposta_progettista", dateRispostaProgettista[0] || "");
      fd.append("data_riscontro_ispettore", dateRiscontroIspettore[0] || "");
      dateRispostaProgettista.forEach((data, index) => {
        fd.append(`data_risposta_progettista_${index}`, data);
      });
      dateRiscontroIspettore.forEach((data, index) => {
        fd.append(`data_riscontro_ispettore_${index}`, data);
      });
      fd.append("responsabile_its", responsabileIts);
      fd.append("responsabile_pcq", responsabilePcq);

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
            <b>ZIP schede emissione precedente</b>
            <input
              type="file"
              accept=".zip"
              onChange={(e) => setSchedeEmesseZip(e.target.files?.[0] || null)}
              style={inputStyle}
            />
            <div style={helpStyle}>
              Obbligatorio dalla seconda emissione in poi. Caricare lo ZIP delle schede gia emesse nell emissione precedente; non richiesto per la Rev. 0.
            </div>
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
                Riscontri; 2 = Terza emissione - Riscontri; 3 = Quarta
                emissione - Riscontri; 4 = Quinta emissione - Riscontri.
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
              <div style={helpStyle}>
                Campo mantenuto per compatibilità: se la data della Rev. corrente
                non viene compilata nello storico emissioni, verrà usata questa.
              </div>
            </label>

            <div
              style={{
                display: "grid",
                gap: 12,
                padding: 14,
                border: "1px solid #cbd5e1",
                borderRadius: 10,
                background: "#f8fafc",
              }}
            >
              <h3 style={{ margin: 0, fontSize: 18 }}>
                Storico emissioni da riportare nella prima pagina
              </h3>

              <label>
                <b>Data Rev. 0 - Prima Emissione - Rilievi</b>
                <input type="text" value={dataRev0} onChange={(e) => setDataRev0(e.target.value)} placeholder="gg/mm/aaaa" style={inputStyle} />
              </label>

              <label>
                <b>Data Rev. 1 - Seconda Emissione - Riscontri</b>
                <input type="text" value={dataRev1} onChange={(e) => setDataRev1(e.target.value)} placeholder="gg/mm/aaaa" style={inputStyle} />
              </label>

              <label>
                <b>Data Rev. 2 - Terza Emissione - Riscontri</b>
                <input type="text" value={dataRev2} onChange={(e) => setDataRev2(e.target.value)} placeholder="gg/mm/aaaa" style={inputStyle} />
              </label>

              <label>
                <b>Data Rev. 3 - Quarta Emissione - Riscontri</b>
                <input type="text" value={dataRev3} onChange={(e) => setDataRev3(e.target.value)} placeholder="gg/mm/aaaa" style={inputStyle} />
              </label>

              <label>
                <b>Data Rev. 4 - Quinta Emissione - Riscontri</b>
                <input type="text" value={dataRev4} onChange={(e) => setDataRev4(e.target.value)} placeholder="gg/mm/aaaa" style={inputStyle} />
              </label>
            </div>

            <div
              style={{
                display: "grid",
                gap: 12,
                padding: 14,
                border: "1px solid #cbd5e1",
                borderRadius: 10,
                background: "#f8fafc",
              }}
            >
              <h3 style={{ margin: 0, fontSize: 18 }}>
                Date commenti progettista da inserire nelle schede
              </h3>
              <div style={helpStyle}>
                Compilare una data per ogni commento distinto. Se piu commenti dello stesso autore hanno la stessa data, la data verra riportata una sola volta.
              </div>

              {dateRispostaProgettista.map((data, index) => (
                <label key={`data-risposta-progettista-${index}`}>
                  <b>Data risposta progettista {index + 1}</b>
                  <input
                    type="text"
                    value={data}
                    onChange={(e) => {
                      const next = [...dateRispostaProgettista];
                      next[index] = e.target.value;
                      setDateRispostaProgettista(next);
                    }}
                    placeholder="gg/mm/aaaa"
                    style={inputStyle}
                  />
                </label>
              ))}
            </div>

            <div
              style={{
                display: "grid",
                gap: 12,
                padding: 14,
                border: "1px solid #cbd5e1",
                borderRadius: 10,
                background: "#f8fafc",
              }}
            >
              <h3 style={{ margin: 0, fontSize: 18 }}>
                Date commenti ispettore ITS da inserire nelle schede
              </h3>
              <div style={helpStyle}>
                Compilare una data per ogni riscontro distinto. Le date sono usate solo quando Trimble/BCF non fornisce gia la data del commento.
              </div>

              {dateRiscontroIspettore.map((data, index) => (
                <label key={`data-riscontro-ispettore-${index}`}>
                  <b>Data riscontro ispettore ITS {index + 1}</b>
                  <input
                    type="text"
                    value={data}
                    onChange={(e) => {
                      const next = [...dateRiscontroIspettore];
                      next[index] = e.target.value;
                      setDateRiscontroIspettore(next);
                    }}
                    placeholder="gg/mm/aaaa"
                    style={inputStyle}
                  />
                </label>
              ))}
            </div>

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
