"use client";

import Link from "next/link";
import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

type ElencoElaborato = {
  codice?: string;
  titolo?: string;
  revisione?: string;
  disciplina?: string;
  formato?: string;
};

type AiReceptionResult = {
  commessa?: string;
  codiceDocumentoFile?: string;
  codiceDocumentoCartiglio?: string;
  titoloElaborato?: string;
  revisioneFile?: string;
  revisioneCartiglio?: string;
  dataRevisione?: string;
  disciplina?: string;
  faseProgettuale?: string;
  formatoDocumento?: string;
  formatoCartiglio?: string;
  coerenze?: {
    cartiglioLeggibile?: boolean;
    titoloPresente?: boolean;
  };
  incoerenze?: string[];
  azioniConsigliate?: string[];
  confidenza?: number;
};

type Row = {
  cartella: string;
  percorso: string;
  fileName: string;

  codiceElenco: string;
  codiceFile: string;
  codiceCartiglio: string;

  titoloElenco: string;
  titoloCartiglio: string;

  revisioneElenco: string;
  revisioneFile: string;
  revisioneCartiglio: string;

  disciplinaElenco: string;
  disciplinaCartiglio: string;

  formatoElenco: string;
  formatoDocumento: string;

  statoControllo: "PRESENTE" | "INCOERENTE";
  anomalia: string;
  azioneRichiesta: string;
  confidenza: number | string;
};

function normalizeCode(value: string) {
  return String(value || "")
    .toUpperCase()
    .replace(/\.PDF$/i, "")
    .replace(/[-_\s.]/g, "")
    .trim();
}

function normalizeCodeWithoutRevision(value: string) {
  return normalizeCode(value).replace(/\d{2}$/, "");
}

function normalizeText(value: string) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeRevision(value: string) {
  const text = String(value || "").trim().toUpperCase();
  const revMatch = text.match(/(?:REV(?:ISIONE)?\.?\s*)?([0-9]{1,2}|[A-Z])$/i);

  if (!revMatch) return text;

  const rev = revMatch[1];

  if (/^\d+$/.test(rev)) {
    return rev.padStart(2, "0");
  }

  return rev;
}

function extractFileBaseName(fileName: string) {
  return String(fileName || "").replace(/\.pdf$/i, "").trim();
}

function extractRevisionFromFileName(fileName: string) {
  const base = extractFileBaseName(fileName);
  const match = base.match(/[-_ ]([0-9]{1,2}|[A-Z])$/i);

  if (!match) return "";

  return normalizeRevision(match[1]);
}

function getRelativePath(file: File) {
  return (
    (file as File & { webkitRelativePath?: string }).webkitRelativePath ||
    file.name
  );
}

function getFolderPath(file: File) {
  const path = getRelativePath(file);
  const parts = path.split("/");

  if (parts.length <= 1) return "-";

  return parts.slice(0, -1).join("/");
}

function getFolderName(file: File) {
  const folderPath = getFolderPath(file);

  if (folderPath === "-") return "-";

  const parts = folderPath.split("/").filter(Boolean);

  return parts[parts.length - 1] || "-";
}

function findMatchingElaborato(
  fileName: string,
  cartiglioCode: string,
  elenco: ElencoElaborato[]
) {
  const fileBase = extractFileBaseName(fileName);

  const fileNorm = normalizeCode(fileBase);
  const fileNormNoRev = normalizeCodeWithoutRevision(fileBase);
  const cartiglioNorm = normalizeCode(cartiglioCode);
  const cartiglioNormNoRev = normalizeCodeWithoutRevision(cartiglioCode);

  return (
    elenco.find((item) => normalizeCode(item.codice || "") === fileNorm) ||
    elenco.find(
      (item) => normalizeCodeWithoutRevision(item.codice || "") === fileNormNoRev
    ) ||
    elenco.find((item) => normalizeCode(item.codice || "") === cartiglioNorm) ||
    elenco.find(
      (item) =>
        normalizeCodeWithoutRevision(item.codice || "") === cartiglioNormNoRev
    ) ||
    null
  );
}

function buildAnomalie(params: {
  codiceOk: boolean;
  titoloOk: boolean;
  revisioneOk: boolean;
  formatoOk: boolean;
  cartiglioLeggibile: boolean;
  presenteInElenco: boolean;
}) {
  const anomalie: string[] = [];

  if (!params.presenteInElenco) {
    anomalie.push("Elaborato non presente nell'elenco elaborati");
  }

  if (!params.cartiglioLeggibile) {
    anomalie.push("Cartiglio non leggibile o dati cartiglio non estratti");
  }

  if (!params.codiceOk) {
    anomalie.push("Codice elaborato non coerente tra elenco/file/cartiglio");
  }

  if (!params.titoloOk) {
    anomalie.push("Titolo elaborato differente tra elenco e cartiglio");
  }

  if (!params.revisioneOk) {
    anomalie.push("Revisione differente tra elenco/file/cartiglio");
  }

  if (!params.formatoOk) {
    anomalie.push("Formato documento differente tra elenco e cartiglio/documento");
  }

  return anomalie;
}

function buildAzioneRichiesta(anomalie: string[]) {
  if (!anomalie.length) return "Nessuna azione richiesta";

  if (anomalie.some((a) => a.includes("non presente"))) {
    return "Verificare trasmissione/elenco elaborati e integrare il documento mancante o aggiornare l'elenco";
  }

  if (anomalie.some((a) => a.includes("Cartiglio"))) {
    return "Ritrasmettere il PDF con cartiglio leggibile e correttamente compilato";
  }

  if (anomalie.some((a) => a.includes("Codice"))) {
    return "Correggere codice elaborato nel nome file, nell'elenco o nel cartiglio";
  }

  if (anomalie.some((a) => a.includes("Revisione"))) {
    return "Allineare revisione tra elenco elaborati, nome file e cartiglio";
  }

  if (anomalie.some((a) => a.includes("Titolo"))) {
    return "Allineare titolo elaborato tra elenco elaborati e cartiglio";
  }

  if (anomalie.some((a) => a.includes("Formato"))) {
    return "Verificare e correggere il formato documento indicato";
  }

  return "Verificare e ritrasmettere gli elaborati corretti";
}

function buildEmailText(rows: Row[]) {
  const incoerenti = rows.filter((row) => row.statoControllo === "INCOERENTE");

  if (!incoerenti.length) {
    return `Buongiorno,

a seguito della verifica documentale effettuata sugli elaborati ricevuti, non risultano anomalie tra elenco elaborati, file PDF e cartigli.

Cordiali saluti`;
  }

  const elenco = incoerenti
    .map(
      (row, index) =>
        `${index + 1}. ${row.fileName} - ${row.anomalia} - ${row.azioneRichiesta}`
    )
    .join("\n");

  return `Buongiorno,

a seguito della verifica documentale effettuata sugli elaborati ricevuti, sono state rilevate le seguenti incoerenze tra elenco elaborati, file PDF e cartigli:

${elenco}

Si richiede verifica e ritrasmissione degli elaborati corretti o aggiornamento dell'elenco elaborati.

In allegato si trasmette il report di controllo.

Cordiali saluti`;
}

export default function NotaRicezioneAIPage() {
  const [elencoFile, setElencoFile] = useState<File | null>(null);
  const [elenco, setElenco] = useState<ElencoElaborato[]>([]);
  const [files, setFiles] = useState<File[]>([]);
  const [rows, setRows] = useState<Row[]>([]);
  const [emailText, setEmailText] = useState("");
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState("");
  const [error, setError] = useState("");

  async function estraiElenco(file: File) {
    const fd = new FormData();
    fd.append("file", file);

    const res = await fetch("/api/ai-extract-elenco", {
      method: "POST",
      body: fd,
    });

    const data = await res.json();

    if (!res.ok || !data.success) {
      throw new Error(data.error || "Errore durante l'estrazione dell'elenco elaborati.");
    }

    return (data.elaborati || []) as ElencoElaborato[];
  }

  async function leggiCartiglio(file: File) {
    const fd = new FormData();

    fd.append("file", file);
    fd.append("mode", "document-reception-check");

    const res = await fetch("/api/ai-document-reader", {
      method: "POST",
      body: fd,
    });

    const data = await res.json();

    if (!res.ok || !data.success) {
      throw new Error(data.error || "Errore durante la lettura del cartiglio.");
    }

    try {
      return typeof data.result === "string"
        ? (JSON.parse(data.result) as AiReceptionResult)
        : (data.result as AiReceptionResult);
    } catch {
      return {
        incoerenze: ["Risposta AI non leggibile come JSON."],
        azioniConsigliate: ["Verificare manualmente il documento."],
      } as AiReceptionResult;
    }
  }

  async function analizzaPdf() {
    if (!elencoFile) {
      alert("Carica l'elenco elaborati PDF prodotto dai progettisti.");
      return;
    }

    if (!files.length) {
      alert("Carica la cartella principale contenente gli elaborati PDF da verificare.");
      return;
    }

    setLoading(true);
    setError("");
    setProgress("");
    setRows([]);
    setEmailText("");

    try {
      setProgress("Estrazione elenco elaborati in corso...");
      const elencoEstratto = await estraiElenco(elencoFile);

      setElenco(elencoEstratto);

      const newRows: Row[] = [];

      for (let i = 0; i < files.length; i += 1) {
        const file = files[i];

        setProgress(`Lettura cartiglio ${i + 1} di ${files.length}: ${file.name}`);

        const ai = await leggiCartiglio(file);

        const codiceFile = extractFileBaseName(file.name);
        const revisioneFile =
          ai.revisioneFile || extractRevisionFromFileName(file.name);

        const codiceCartiglio =
          ai.codiceDocumentoCartiglio || ai.codiceDocumentoFile || "";
        const titoloCartiglio = ai.titoloElaborato || "";
        const revisioneCartiglio = ai.revisioneCartiglio || "";

        const match = findMatchingElaborato(file.name, codiceCartiglio, elencoEstratto);

        const codiceElenco = match?.codice || "";
        const titoloElenco = match?.titolo || "";
        const revisioneElenco = match?.revisione || "";
        const formatoElenco = match?.formato || "";
        const disciplinaElenco = match?.disciplina || "";

        const codiceOk =
          !!match &&
          (
            normalizeCodeWithoutRevision(codiceElenco) ===
              normalizeCodeWithoutRevision(codiceFile) ||
            normalizeCodeWithoutRevision(codiceElenco) ===
              normalizeCodeWithoutRevision(codiceCartiglio) ||
            normalizeCodeWithoutRevision(codiceFile) ===
              normalizeCodeWithoutRevision(codiceCartiglio)
          );

        const titoloOk =
          !titoloElenco ||
          !titoloCartiglio ||
          normalizeText(titoloElenco) === normalizeText(titoloCartiglio) ||
          normalizeText(titoloCartiglio).includes(normalizeText(titoloElenco)) ||
          normalizeText(titoloElenco).includes(normalizeText(titoloCartiglio));

        const revElenco = normalizeRevision(revisioneElenco);
        const revFile = normalizeRevision(revisioneFile);
        const revCartiglio = normalizeRevision(revisioneCartiglio);

        const revisioneOk =
          !revElenco ||
          revElenco === revFile ||
          revElenco === revCartiglio ||
          (!!revFile && !!revCartiglio && revFile === revCartiglio);

        const formatoDocumento = ai.formatoDocumento || ai.formatoCartiglio || "";

        const formatoOk =
          !formatoElenco ||
          !formatoDocumento ||
          normalizeText(formatoElenco) === normalizeText(formatoDocumento);

        const cartiglioLeggibile =
          ai.coerenze?.cartiglioLeggibile !== false &&
          (!!codiceCartiglio || !!titoloCartiglio || !!revisioneCartiglio);

        const anomalie = buildAnomalie({
          codiceOk,
          titoloOk,
          revisioneOk,
          formatoOk,
          cartiglioLeggibile,
          presenteInElenco: !!match,
        });

        const statoControllo: "PRESENTE" | "INCOERENTE" =
          anomalie.length === 0 ? "PRESENTE" : "INCOERENTE";

        newRows.push({
          cartella: getFolderName(file),
          percorso: getFolderPath(file),
          fileName: file.name,

          codiceElenco,
          codiceFile,
          codiceCartiglio,

          titoloElenco,
          titoloCartiglio,

          revisioneElenco,
          revisioneFile,
          revisioneCartiglio,

          disciplinaElenco,
          disciplinaCartiglio: ai.disciplina || "",

          formatoElenco,
          formatoDocumento,

          statoControllo,
          anomalia: anomalie.join(" | ") || "-",
          azioneRichiesta: buildAzioneRichiesta(anomalie),
          confidenza: ai.confidenza ?? "",
        });
      }

      setRows(newRows);
      setEmailText(buildEmailText(newRows));
      setProgress("Analisi completata.");
    } catch (err: any) {
      setError(err?.message || "Errore imprevisto.");
      setProgress("");
    } finally {
      setLoading(false);
    }
  }

  function svuota() {
    setElencoFile(null);
    setElenco([]);
    setFiles([]);
    setRows([]);
    setEmailText("");
    setProgress("");
    setError("");
  }

  async function esportaExcel() {
    try {
      const templateResponse = await fetch("/Nota_Ricezione_Template.xlsx");

      if (!templateResponse.ok) {
        throw new Error("Template Nota_Ricezione_Template.xlsx non trovato in /public.");
      }

      const templateArrayBuffer = await templateResponse.arrayBuffer();

      const workbook = XLSX.read(templateArrayBuffer, {
        type: "array",
      });

      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      const reportRows = rows.map((row) => ({
        CARTELLA: row.cartella || "-",
        "FILE PDF": row.fileName || "-",
        "CODICE ELENCO": row.codiceElenco || "-",
        "CODICE CARTIGLIO": row.codiceCartiglio || "-",
        "TITOLO ELENCO": row.titoloElenco || "-",
        "TITOLO CARTIGLIO": row.titoloCartiglio || "-",
        "REV ELENCO": row.revisioneElenco || "-",
        "REV CARTIGLIO": row.revisioneCartiglio || "-",
        STATO: row.statoControllo || "-",
        ANOMALIA: row.anomalia || "-",
        "AZIONE RICHIESTA": row.azioneRichiesta || "-",
      }));

      XLSX.utils.sheet_add_json(worksheet, reportRows, {
        origin: "A10",
        skipHeader: false,
      });

      const riepilogoSheet = XLSX.utils.aoa_to_sheet([
        ["RIEPILOGO CONTROLLO DOCUMENTALE"],
        [""],
        ["Elaborati analizzati", rows.length],
        ["PRESENTE", rows.filter((row) => row.statoControllo === "PRESENTE").length],
        ["INCOERENTE", rows.filter((row) => row.statoControllo === "INCOERENTE").length],
        ["Elaborati in elenco", elenco.length],
        [""],
        ["NOTE"],
        [
          'Il carattere "-" e il carattere "_" nel codice elaborato sono considerati equivalenti.',
        ],
      ]);

      XLSX.utils.book_append_sheet(workbook, riepilogoSheet, "Riepilogo");

      const emailSheet = XLSX.utils.aoa_to_sheet([
        ["TESTO EMAIL AI PROGETTISTI"],
        [""],
        [emailText || ""],
      ]);

      XLSX.utils.book_append_sheet(workbook, emailSheet, "Email_Progettisti");

      XLSX.writeFile(workbook, "Report_Completo.xlsx");
    } catch (error) {
      console.error(error);

      alert(
        error instanceof Error
          ? error.message
          : "Errore durante la generazione del report Excel completo."
      );
    }
  }

  const kpi = useMemo(() => {
    const totale = rows.length;
    const presenti = rows.filter((row) => row.statoControllo === "PRESENTE").length;
    const incoerenti = rows.filter(
      (row) => row.statoControllo === "INCOERENTE"
    ).length;
    const cartigliLeggibili = rows.filter(
      (row) => row.codiceCartiglio || row.titoloCartiglio || row.revisioneCartiglio
    ).length;

    return {
      totale,
      presenti,
      incoerenti,
      cartigliLeggibili,
    };
  }, [rows]);

  return (
    <main style={pageStyle}>
      <Link href="/dashboard-ai" style={backLinkStyle}>
        ← Torna alla Dashboard AI
      </Link>

      <h1 style={titleStyle}>Nota Ricezione Elaborati AI</h1>

      <p style={leadStyle}>
        Controllo documentale tra elenco elaborati, PDF ricevuti e cartigli.
        Il report indica se il documento è PRESENTE oppure INCOERENTE, con
        anomalia rilevata e azione richiesta per l'inoltro ai progettisti.
      </p>

      <section style={cardStyle}>
        <h2>Verifica ricezione elaborati</h2>

        <div style={uploadGridStyle}>
          <div>
            <label style={labelStyle}>Elenco Elaborati PDF</label>
            <input
              type="file"
              accept=".pdf"
              onChange={(e) => setElencoFile(e.target.files?.[0] || null)}
              style={inputStyle}
            />

            {elencoFile && (
              <div style={infoStyle}>Elenco caricato: {elencoFile.name}</div>
            )}
          </div>

          <div>
            <label style={labelStyle}>
              Cartella principale con elaborati PDF
            </label>

            <input
              type="file"
              accept=".pdf"
              multiple
              {...({
                webkitdirectory: "",
                directory: "",
              } as any)}
              onChange={(e) => {
                const selectedFiles = Array.from(e.target.files || []).filter(
                  (file) => file.name.toLowerCase().endsWith(".pdf")
                );

                setFiles(selectedFiles);
              }}
              style={inputStyle}
            />

            {files.length > 0 && (
              <div style={infoStyle}>PDF selezionati: {files.length}</div>
            )}
          </div>
        </div>

        <div style={noteStyle}>
          Il carattere "-" e il carattere "_" nel codice elaborato sono
          considerati equivalenti. Il controllo confronta codice, titolo e
          revisione tra elenco, file e cartiglio.
        </div>

        <div style={actionsStyle}>
          <button onClick={analizzaPdf} disabled={loading} style={primaryButton}>
            {loading ? "Analisi in corso..." : "Esegui controllo documentale"}
          </button>

          <button onClick={svuota} disabled={loading} style={secondaryButton}>
            Svuota dati
          </button>

          <button
            onClick={esportaExcel}
            disabled={!rows.length || loading}
            style={{
              ...exportButton,
              opacity: !rows.length || loading ? 0.5 : 1,
            }}
          >
            Esporta Excel
          </button>
        </div>

        {progress && <div style={progressStyle}>{progress}</div>}

        {error && <div style={errorStyle}>{error}</div>}
      </section>

      <section style={gridStyle}>
        <Kpi title="Elaborati analizzati" value={String(kpi.totale)} />
        <Kpi title="PRESENTE" value={String(kpi.presenti)} />
        <Kpi title="INCOERENTE" value={String(kpi.incoerenti)} />
        <Kpi title="Cartigli leggibili" value={String(kpi.cartigliLeggibili)} />
        <Kpi title="Elaborati in elenco" value={String(elenco.length)} />
      </section>

      {rows.length > 0 && (
        <section style={cardStyle}>
          <h2>Report sintetico controllo documentale</h2>

          <div style={{ overflowX: "auto" }}>
            <table style={tableStyle}>
              <thead>
                <tr>
                  <th style={th}>Cartella</th>
                  <th style={th}>File PDF</th>
                  <th style={th}>Codice elenco</th>
                  <th style={th}>Codice cartiglio</th>
                  <th style={th}>Titolo elenco</th>
                  <th style={th}>Titolo cartiglio</th>
                  <th style={th}>Rev elenco</th>
                  <th style={th}>Rev cartiglio</th>
                  <th style={th}>Stato controllo</th>
                  <th style={th}>Anomalia rilevata</th>
                  <th style={th}>Azione richiesta</th>
                </tr>
              </thead>

              <tbody>
                {rows.map((row, index) => (
                  <tr key={`${row.fileName}-${index}`}>
                    <td style={td}>{row.cartella}</td>
                    <td style={td}>{row.fileName}</td>
                    <td style={td}>{row.codiceElenco || "-"}</td>
                    <td style={td}>{row.codiceCartiglio || "-"}</td>
                    <td style={td}>{row.titoloElenco || "-"}</td>
                    <td style={td}>{row.titoloCartiglio || "-"}</td>
                    <td style={td}>{row.revisioneElenco || "-"}</td>
                    <td style={td}>{row.revisioneCartiglio || "-"}</td>
                    <td
                      style={{
                        ...td,
                        fontWeight: 800,
                        color:
                          row.statoControllo === "PRESENTE"
                            ? "#16a34a"
                            : "#dc2626",
                      }}
                    >
                      {row.statoControllo}
                    </td>
                    <td style={td}>{row.anomalia}</td>
                    <td style={td}>{row.azioneRichiesta}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      {emailText && (
        <section style={cardStyle}>
          <h2>Testo email ai progettisti</h2>

          <textarea
            value={emailText}
            readOnly
            style={emailTextAreaStyle}
          />
        </section>
      )}
    </main>
  );
}

function Kpi({ title, value }: { title: string; value: string }) {
  return (
    <div style={cardStyle}>
      <div style={{ color: "#64748b", fontSize: 14 }}>{title}</div>
      <div style={{ fontSize: 30, fontWeight: 800, marginTop: 8 }}>
        {value}
      </div>
    </div>
  );
}

const pageStyle: React.CSSProperties = {
  minHeight: "100vh",
  padding: 30,
  background: "#f1f5f9",
  fontFamily: "Arial, sans-serif",
  color: "#0f172a",
};

const backLinkStyle: React.CSSProperties = {
  display: "inline-block",
  marginBottom: 24,
  color: "#0284c7",
  fontWeight: 700,
  textDecoration: "none",
};

const titleStyle: React.CSSProperties = {
  fontSize: 42,
  marginBottom: 10,
};

const leadStyle: React.CSSProperties = {
  fontSize: 18,
  color: "#475569",
  maxWidth: 1200,
};

const cardStyle: React.CSSProperties = {
  background: "white",
  border: "1px solid #cbd5e1",
  borderRadius: 16,
  padding: 20,
  marginTop: 24,
  boxShadow: "0 8px 20px rgba(15,23,42,0.06)",
};

const uploadGridStyle: React.CSSProperties = {
  display: "grid",
  gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))",
  gap: 18,
};

const labelStyle: React.CSSProperties = {
  display: "block",
  fontWeight: 800,
  marginBottom: 8,
};

const inputStyle: React.CSSProperties = {
  width: "100%",
  padding: 12,
  border: "1px solid #cbd5e1",
  borderRadius: 10,
  marginBottom: 14,
};

const actionsStyle: React.CSSProperties = {
  display: "flex",
  gap: 12,
  marginTop: 18,
  flexWrap: "wrap",
};

const primaryButton: React.CSSProperties = {
  padding: "12px 18px",
  background: "#0f172a",
  color: "white",
  border: "none",
  borderRadius: 10,
  fontWeight: 700,
  cursor: "pointer",
};

const secondaryButton: React.CSSProperties = {
  padding: "12px 18px",
  background: "#94a3b8",
  color: "white",
  border: "none",
  borderRadius: 10,
  fontWeight: 700,
  cursor: "pointer",
};

const exportButton: React.CSSProperties = {
  padding: "12px 18px",
  background: "#16a34a",
  color: "white",
  border: "none",
  borderRadius: 10,
  fontWeight: 700,
  cursor: "pointer",
};

const infoStyle: React.CSSProperties = {
  marginTop: 6,
  padding: 12,
  borderRadius: 10,
  background: "#eff6ff",
  color: "#1d4ed8",
  fontWeight: 700,
};

const noteStyle: React.CSSProperties = {
  marginTop: 10,
  padding: 12,
  borderRadius: 10,
  background: "#f8fafc",
  color: "#475569",
  lineHeight: 1.5,
};

const progressStyle: React.CSSProperties = {
  marginTop: 16,
  padding: 14,
  borderRadius: 10,
  background: "#ecfeff",
  color: "#155e75",
  fontWeight: 700,
};

const errorStyle: React.CSSProperties = {
  marginTop: 16,
  padding: 14,
  borderRadius: 10,
  background: "#fee2e2",
  color: "#991b1b",
};

const gridStyle: React.CSSProperties = {
  display: "grid",
  gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
  gap: 16,
};

const tableStyle: React.CSSProperties = {
  width: "100%",
  borderCollapse: "collapse",
  fontSize: 14,
};

const th: React.CSSProperties = {
  textAlign: "left",
  borderBottom: "1px solid #cbd5e1",
  padding: 10,
  background: "#f8fafc",
  whiteSpace: "nowrap",
};

const td: React.CSSProperties = {
  borderBottom: "1px solid #e2e8f0",
  padding: 10,
  verticalAlign: "top",
};

const emailTextAreaStyle: React.CSSProperties = {
  width: "100%",
  minHeight: 260,
  padding: 14,
  border: "1px solid #cbd5e1",
  borderRadius: 10,
  fontFamily: "Arial, sans-serif",
  fontSize: 14,
  lineHeight: 1.5,
};
