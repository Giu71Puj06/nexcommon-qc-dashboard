"use client";

import React, { useMemo, useState } from "react";
import Link from "next/link";
import * as XLSX from "xlsx";
import JSZip from "jszip";

type CheckRow = {
  rowNumber: number;
  progressivo: number;
  label: string;
  title: string;
  codiceTitleTrimble: string;
  codiceReport: string;
  titleOk: boolean;
  tags: string;
  tagsOk: boolean;
  disciplina: string;
  disciplinaOk: boolean;
  status: string;
  statusOk: boolean;
  description: string;
  tr: string;
  bcfTitle: string;
  bcfDescription: string;
  rispostaProgettista: string;
  riscontroIspettore: string;
  storiaOk: boolean;
  esitoStoria:
    | "COMPLETA"
    | "BCF NON TROVATO"
    | "MANCA RISPOSTA PROGETTISTA"
    | "MANCA RISCONTRO ITS"
    | "CHIUSO SENZA RISCONTRO"
    | "NON APPLICABILE";
  esito: "OK" | "ERRORE";
  anomalie: string[];
};

type BcfTopic = {
  tr: string;
  title: string;
  description: string;
  commentsPrg: string;
  commentsIsp: string;
  allComments: string;
};

type Filters = {
  n: string;
  codiceReport: string;
  codiceTitleTrimble: string;
  esitoCodice: string;
  tags: string;
  disciplina: string;
  status: string;
  tr: string;
  esitoStoria: string;
  esito: string;
  anomalie: string;
};

const TAGS_AMMESSI = ["NC", "OSS", "Nessun rilievo", "Da NC a OSS"];
const STATUS_AMMESSI = ["New", "Closed"];

function normalizeCode(value: string) {
  return String(value || "")
    .toUpperCase()
    .replace(/\.PDF$/i, "")
    .replace(/\s+/g, "")
    .replace(/[_\-.]/g, "")
    .trim();
}

function normalizeText(value: string) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function cleanPdf(value: string) {
  return String(value || "").replace(/\.pdf$/i, "").trim();
}

function getXmlText(parent: Element | Document, tagName: string) {
  const node = parent.getElementsByTagName(tagName)[0];
  return node?.textContent?.trim() || "";
}

function getXmlTexts(parent: Element | Document, tagName: string) {
  return Array.from(parent.getElementsByTagName(tagName))
    .map((node) => node.textContent?.trim() || "")
    .filter(Boolean);
}

function getByLocalName(parent: Element | Document, localName: string) {
  return Array.from(parent.getElementsByTagName("*")).filter(
    (node) => node.localName.toLowerCase() === localName.toLowerCase()
  );
}

function getLocalText(parent: Element | Document, localName: string) {
  const node = getByLocalName(parent, localName)[0];
  return node?.textContent?.trim() || "";
}

function getLocalTexts(parent: Element | Document, localName: string) {
  return getByLocalName(parent, localName)
    .map((node) => node.textContent?.trim() || "")
    .filter(Boolean);
}


function extractTR(value: string) {
  const text = String(value || "").toUpperCase();

  const tr = text.match(/\bTR[-_\s]*0*(\d+[A-Z]?)\b/i);
  if (tr) return `TR-${tr[1]}`;

  const it = text.match(/\bIT(?:\d+)?[-_\s]*0*(\d+[A-Z]?)\b/i);
  if (it) return `TR-${it[1]}`;

  const bracket = text.match(/\((?:TR|IT)[-_\s]*0*(\d+[A-Z]?)\)/i);
  if (bracket) return `TR-${bracket[1]}`;

  return "";
}

function extractTRFromTodo(label: string, title: string, description: string) {
  return extractTR(label) || extractTR(title) || extractTR(description);
}

function extractComparableCode(value: string) {
  const text = String(value || "");
  const match = text.match(/PV\d{3}-[A-Z0-9]+-[A-Z0-9]+-[A-Z]{3}-\d{5}-[A-Z]{3}-\d{6}(?:[-_ ]?\d+)?/i);
  return match ? normalizeCode(match[0]) : "";
}

function isKnownInspectorAuthor(author: string) {
  const a = normalizeText(author);

  return (
    a.includes("massimo tamberi") ||
    a.includes("edoardo oddo") ||
    a.includes("guido bonin") ||
    a.includes("ilaria martarelli") ||
    a.includes("michea sciorra") ||
    a.includes("stefano arcangeli") ||
    a.includes("massimo") ||
    a.includes("edoardo") ||
    a.includes("guido") ||
    a.includes("ilaria") ||
    a.includes("michea") ||
    a.includes("stefano")
  );
}


function isPrgAuthor(author: string) {
  const a = normalizeText(author);

  return (
    a.includes("prg") ||
    a.includes("progett") ||
    a.includes("pcq") ||
    a.includes("rtp") ||
    a.includes("mandante")
  );
}

function isIspAuthor(author: string) {
  const a = normalizeText(author);

  return (
    a.includes("its") ||
    a.includes("isp") ||
    a.includes("controlli tecnici") ||
    a.includes("odi") ||
    a.includes("ispett") ||
    isKnownInspectorAuthor(author)
  );
}

function formatComment(date: string, author: string, text: string) {
  const cleanDate = String(date || "").slice(0, 10);
  const parts = [cleanDate, author].filter(Boolean).join(" - ");
  return parts ? `${parts}: ${text}` : text;
}

async function parseBcfFiles(files: File[]) {
  const topics: BcfTopic[] = [];

  async function parseMarkupXml(xml: string, sourceName: string) {
    const doc = new DOMParser().parseFromString(xml, "application/xml");

    const topicNode =
      getByLocalName(doc, "Topic")[0] ||
      doc.getElementsByTagName("Topic")[0];

    const title = topicNode
      ? getLocalText(topicNode, "Title") || getXmlText(topicNode, "Title")
      : "";
    const description = topicNode
      ? getLocalText(topicNode, "Description") || getXmlText(topicNode, "Description")
      : "";
    const topicLabels = topicNode
      ? getLocalTexts(topicNode, "Label").join(" ") || getXmlTexts(topicNode, "Label").join(" ")
      : "";

    const comments = getByLocalName(doc, "Comment").filter((node) =>
      Array.from(node.children || []).some((child) =>
        ["Date", "Author", "Comment", "ModifiedAuthor", "ModifiedDate"].includes(child.localName)
      )
    );

    const prgComments: string[] = [];
    const ispComments: string[] = [];
    const allComments: string[] = [];

    comments.forEach((comment) => {
      const author =
        getLocalText(comment, "Author") ||
        getLocalText(comment, "ModifiedAuthor") ||
        "";
      const date =
        getLocalText(comment, "Date") ||
        getLocalText(comment, "ModifiedDate") ||
        "";

      const commentTexts = getByLocalName(comment, "Comment")
        .filter((node) => node !== comment)
        .map((node) => node.textContent?.trim() || "")
        .filter(Boolean);

      const text = commentTexts[0] || "";
      if (!text) return;

      const formatted = formatComment(date, author, text);
      allComments.push(formatted);

      if (isIspAuthor(author)) {
        ispComments.push(formatted);
      } else {
        prgComments.push(formatted);
      }
    });

    const tr =
      extractTR(sourceName) ||
      extractTR(topicLabels) ||
      extractTR(title) ||
      extractTR(description) ||
      extractTR(allComments.join(" "));

    const comparableCode =
      extractComparableCode(title) ||
      extractComparableCode(description) ||
      extractComparableCode(sourceName) ||
      extractComparableCode(allComments.join(" "));

    if (!tr && !title && !description && allComments.length === 0 && !comparableCode) {
      return;
    }

    topics.push({
      tr,
      title,
      description: comparableCode ? `${description}\n${comparableCode}`.trim() : description,
      commentsPrg: prgComments.join("\n\n"),
      commentsIsp: ispComments.join("\n\n"),
      allComments: allComments.join("\n\n"),
    });
  }

  for (const file of files) {
    const lowerName = file.name.toLowerCase();

    if (lowerName.endsWith(".bcf") || lowerName.endsWith(".xml")) {
      await parseMarkupXml(await file.text(), file.name);
      continue;
    }

    const zip = await JSZip.loadAsync(await file.arrayBuffer());

    const markupFiles = Object.values(zip.files).filter((entry) => {
      const name = entry.name.toLowerCase();

      return (
        !entry.dir &&
        (name.endsWith("markup.bcf") ||
          name.endsWith(".bcf") ||
          name.endsWith(".xml"))
      );
    });

    for (const entry of markupFiles) {
      try {
        const xml = await entry.async("text");
        await parseMarkupXml(xml, entry.name);
      } catch {
        // ignora file non XML dentro lo ZIP
      }
    }
  }

  return topics;
}

function findBcfByTR(
  tr: string,
  topics: BcfTopic[],
  title = "",
  description = "",
  codiceReport = ""
) {
  const normalizedTr = normalizeCode(tr);
  const comparableCode =
    extractComparableCode(title) ||
    extractComparableCode(codiceReport) ||
    extractComparableCode(description);

  if (normalizedTr) {
    const exact = topics.find((topic) => normalizeCode(topic.tr) === normalizedTr);
    if (exact) return exact;

    const byText = topics.find((topic) =>
      normalizeCode(`${topic.tr} ${topic.title} ${topic.description} ${topic.allComments}`).includes(normalizedTr)
    );
    if (byText) return byText;
  }

  if (comparableCode) {
    const byCode = topics.find((topic) =>
      normalizeCode(`${topic.title} ${topic.description} ${topic.allComments}`).includes(comparableCode)
    );
    if (byCode) return byCode;
  }

  return null;
}

function isDescrizioneGenerale(value: string) {
  const text = String(value || "").trim();
  const normalized = normalizeCode(text);

  if (!text) return false;
  if (/\.pdf/i.test(text)) return false;

  const hasNumber = /\d/.test(text);
  const hasCodeSeparators = /[-_]/.test(text);
  const looksLikeCode = hasNumber && (hasCodeSeparators || normalized.length > 8);

  return !looksLikeCode;
}

function findBestReportCodeFromTitle(
  title: string,
  reportCodes: Map<string, string>
) {
  const normalizedTitle = normalizeCode(title);

  let bestNormalized = "";
  let bestOriginal = "";

  for (const [normalizedReportCode, originalReportCode] of reportCodes.entries()) {
    if (
      normalizedTitle === normalizedReportCode ||
      normalizedTitle.includes(normalizedReportCode)
    ) {
      if (normalizedReportCode.length > bestNormalized.length) {
        bestNormalized = normalizedReportCode;
        bestOriginal = originalReportCode;
      }
    }
  }

  return bestOriginal;
}

async function readXlsxRows(file: File, preferredSheetName?: string) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });

  const sheetName =
    preferredSheetName && workbook.SheetNames.includes(preferredSheetName)
      ? preferredSheetName
      : workbook.SheetNames[0];

  const sheet = workbook.Sheets[sheetName];

  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
  }) as any[][];
}

function esitoIcon(ok: boolean) {
  return ok ? "✅" : "❌";
}

export default function ControlloTodoIspettoriPage() {
  const [todoFile, setTodoFile] = useState<File | null>(null);
  const [reportFile, setReportFile] = useState<File | null>(null);
  const [elencoFile, setElencoFile] = useState<File | null>(null);
  const [bcfFiles, setBcfFiles] = useState<File[]>([]);

  const [todoRows, setTodoRows] = useState<any[][]>([]);
  const [reportRows, setReportRows] = useState<any[][]>([]);
  const [elencoRows, setElencoRows] = useState<any[][]>([]);
  const [bcfTopics, setBcfTopics] = useState<BcfTopic[]>([]);

  const [loading, setLoading] = useState(false);

  const [filters, setFilters] = useState<Filters>({
    n: "",
    codiceReport: "",
    codiceTitleTrimble: "",
    esitoCodice: "",
    tags: "",
    disciplina: "",
    status: "",
    tr: "",
    esitoStoria: "",
    esito: "",
    anomalie: "",
  });

  function updateFilter(key: keyof Filters, value: string) {
    setFilters((prev) => ({
      ...prev,
      [key]: value,
    }));
  }

  async function eseguiControllo() {
    if (!todoFile || !reportFile || !elencoFile) {
      alert("Carica ToDo XLSX, Report_Completo.xlsx ed ELENCO_ELABORATI.xlsx.");
      return;
    }

    setLoading(true);

    try {
      const todo = await readXlsxRows(todoFile);
      const report = await readXlsxRows(reportFile, "Verifica Elaborati");
      const elenco = await readXlsxRows(elencoFile);
      const bcf = bcfFiles.length > 0 ? await parseBcfFiles(bcfFiles) : [];

      setTodoRows(todo);
      setReportRows(report);
      setElencoRows(elenco);
      setBcfTopics(bcf);
    } catch (error) {
      console.error(error);
      alert("Errore durante la lettura dei file XLSX/BCFZIP.");
    } finally {
      setLoading(false);
    }
  }

  const reportCodes = useMemo(() => {
    const map = new Map<string, string>();

    reportRows.slice(1).forEach((row) => {
      const codiceReport = String(row[2] || "").trim();
      const normalized = normalizeCode(codiceReport);

      if (normalized && normalized !== "NAN") {
        map.set(normalized, codiceReport);
      }
    });

    return map;
  }, [reportRows]);

  const disciplineAmmesse = useMemo(() => {
    const set = new Set<string>();

    elencoRows.slice(1).forEach((row) => {
      const disciplina = String(row[5] || "").trim();

      if (disciplina) {
        set.add(disciplina.toLowerCase());
      }
    });

    return set;
  }, [elencoRows]);

  const checks: CheckRow[] = useMemo(() => {
    return todoRows.slice(1).map((row, index) => {
      const label = String(row[0] || "").trim();
      const title = String(row[1] || "").trim();
      const description = String(row[2] || "").trim();
      const status = String(row[5] || "").trim();
      const disciplina = String(row[8] || "").trim();
      const tags = String(row[9] || "").trim();

      const anomalie: string[] = [];

      const codiceTitleTrimble = cleanPdf(title);
      const titleContienePdf = /\.pdf/i.test(title);
      const titleDescrittivo = isDescrizioneGenerale(title);
      const codiceReport = titleDescrittivo
        ? ""
        : findBestReportCodeFromTitle(title, reportCodes);

      let titleOk = false;

      if (!title) {
        anomalie.push("Title mancante");
      } else if (titleContienePdf) {
        anomalie.push("Il Title contiene .pdf");
      } else if (titleDescrittivo) {
        titleOk = true;
      } else if (!codiceReport) {
        anomalie.push("Codice elaborato non presente nel Report_Completo.xlsx");
      } else {
        titleOk = true;
      }

      const tagsOk = TAGS_AMMESSI.some(
        (tag) => tag.toLowerCase() === tags.toLowerCase()
      );

      if (!tags) anomalie.push("Tags mancanti");
      else if (!tagsOk) anomalie.push("Tags non valido");

      const disciplinaOk =
        !!disciplina && disciplineAmmesse.has(disciplina.toLowerCase());

      if (!disciplina) anomalie.push("Disciplina mancante");
      else if (!disciplinaOk)
        anomalie.push("Disciplina non presente in ELENCO_ELABORATI.xlsx");

      const statusOk = STATUS_AMMESSI.some(
        (s) => s.toLowerCase() === status.toLowerCase()
      );

      if (!status) anomalie.push("Status mancante");
      else if (!statusOk) anomalie.push("Status non valido");

      const tr = extractTRFromTodo(label, title, description);
      const bcf = findBcfByTR(tr, bcfTopics, title, description, codiceReport);
      const isRilievo = ["NC", "OSS", "Da NC a OSS"].some(
        (tag) => tag.toLowerCase() === tags.toLowerCase()
      );
      const isClosed = status.toLowerCase() === "closed";

      let esitoStoria: CheckRow["esitoStoria"] = "NON APPLICABILE";
      let storiaOk = true;

      if (isRilievo) {
        if (!tr) {
          esitoStoria = "BCF NON TROVATO";
          storiaOk = false;
          anomalie.push("Codice TR non trovato nel ToDo");
        } else if (!bcf) {
          esitoStoria = "BCF NON TROVATO";
          storiaOk = false;
          anomalie.push(`BCF non trovato per ${tr}`);
        } else if (!bcf.commentsPrg && !bcf.allComments) {
          esitoStoria = "MANCA RISPOSTA PROGETTISTA";
          storiaOk = false;
          anomalie.push("Manca risposta progettista nei commenti BCF");
        } else if (!bcf.commentsIsp && isClosed) {
          esitoStoria = "CHIUSO SENZA RISCONTRO";
          storiaOk = false;
          anomalie.push("ToDo chiuso senza riscontro ispettore nei BCF");
        } else if (!bcf.commentsIsp) {
          esitoStoria = "MANCA RISCONTRO ITS";
          storiaOk = false;
          anomalie.push("Manca riscontro ispettore nei commenti BCF");
        } else {
          esitoStoria = "COMPLETA";
        }
      }

      const esito =
        titleOk && tagsOk && disciplinaOk && statusOk && storiaOk
          ? "OK"
          : "ERRORE";

      return {
        rowNumber: index + 2,
        progressivo: index + 1,
        label,
        title,
        description,
        codiceTitleTrimble,
        codiceReport,
        titleOk,
        tags,
        tagsOk,
        disciplina,
        disciplinaOk,
        status,
        statusOk,
        tr,
        bcfTitle: bcf?.title || "",
        bcfDescription: bcf?.description || "",
        rispostaProgettista: bcf?.commentsPrg || (bcf?.commentsIsp ? "" : bcf?.allComments || ""),
        riscontroIspettore: bcf?.commentsIsp || "",
        storiaOk,
        esitoStoria,
        esito,
        anomalie,
      };
    });
  }, [todoRows, reportCodes, disciplineAmmesse, bcfTopics]);

  const filteredChecks = useMemo(() => {
    return checks.filter((row) => {
      const n = `${row.progressivo}${row.label ? ` (${row.label})` : ""}`;
      const esitoCodice = row.titleOk ? "OK" : "ERRORE";
      const anomalie = row.anomalie.join(" | ");

      return (
        n.toLowerCase().includes(filters.n.toLowerCase()) &&
        row.codiceReport
          .toLowerCase()
          .includes(filters.codiceReport.toLowerCase()) &&
        row.codiceTitleTrimble
          .toLowerCase()
          .includes(filters.codiceTitleTrimble.toLowerCase()) &&
        (filters.esitoCodice === "" ||
          esitoCodice === filters.esitoCodice) &&
        row.tags.toLowerCase().includes(filters.tags.toLowerCase()) &&
        row.disciplina
          .toLowerCase()
          .includes(filters.disciplina.toLowerCase()) &&
        row.status.toLowerCase().includes(filters.status.toLowerCase()) &&
        row.tr.toLowerCase().includes(filters.tr.toLowerCase()) &&
        (filters.esitoStoria === "" || row.esitoStoria === filters.esitoStoria) &&
        (filters.esito === "" || row.esito === filters.esito) &&
        anomalie.toLowerCase().includes(filters.anomalie.toLowerCase())
      );
    });
  }, [checks, filters]);

  const totale = checks.length;
  const ok = checks.filter((r) => r.esito === "OK").length;
  const errori = checks.filter((r) => r.esito === "ERRORE").length;
  const completezza = totale > 0 ? Math.round((ok / totale) * 100) : 0;
  const storieComplete = checks.filter((r) => r.esitoStoria === "COMPLETA").length;
  const bcfNonTrovati = checks.filter((r) => r.esitoStoria === "BCF NON TROVATO").length;
  const mancanoRisposte = checks.filter(
    (r) =>
      r.esitoStoria === "MANCA RISPOSTA PROGETTISTA" ||
      r.esitoStoria === "MANCA RISCONTRO ITS" ||
      r.esitoStoria === "CHIUSO SENZA RISCONTRO"
  ).length;

  const completezzaColor =
    completezza === 100 ? "#16a34a" : completezza >= 51 ? "#f59e0b" : "#dc2626";

  function esportaExcel() {
    const rows = filteredChecks.map((row) => ({
      "N.": `${row.progressivo}${row.label ? ` (${row.label})` : ""}`,
      TR: row.tr || "",
      "Codice elaborato Report": row.codiceReport || "",
      "Codice elaborato nel Title Trimble": row.codiceTitleTrimble || "",
      "Esito codice": row.titleOk ? "OK" : "ERRORE",
      Tags: row.tags || "",
      "Esito Tags": row.tagsOk ? "OK" : "ERRORE",
      Disciplina: row.disciplina || "",
      "Esito Disciplina": row.disciplinaOk ? "OK" : "ERRORE",
      Status: row.status || "",
      "Esito Status": row.statusOk ? "OK" : "ERRORE",
      "Description ToDo": row.description || "",
      "Titolo BCF": row.bcfTitle || "",
      "Descrizione BCF": row.bcfDescription || "",
      "Risposta progettista": row.rispostaProgettista || "",
      "Riscontro ispettore ITS": row.riscontroIspettore || "",
      "Esito storia rilievo": row.esitoStoria,
      Esito: row.esito,
      Anomalie: row.anomalie.join(" | "),
    }));

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(rows),
      "Controllo ToDo"
    );

    XLSX.writeFile(workbook, "Report_Controllo_ToDo_BCF.xlsx");
  }

  return (
    <main
      style={{
        minHeight: "100vh",
        padding: 30,
        background: "#f1f5f9",
        fontFamily: "Arial, sans-serif",
      }}
    >
      <Link
        href="/"
        style={{
          color: "#0284c7",
          fontWeight: 700,
          textDecoration: "none",
        }}
      >
        ← Torna alla dashboard
      </Link>

      <h1 style={{ fontSize: 36, marginTop: 24 }}>
        Controllo ToDo Ispettori
      </h1>

      <p style={{ color: "#475569", fontSize: 18, maxWidth: 1150 }}>
        Verifica automatica di Title, Tags, Disciplina, Status e storia del rilievo.
        I file BCFZIP permettono di confrontare la Description iniziale del ToDo
        con i commenti del progettista e il riscontro finale dell'ispettore ITS.
      </p>

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(4, 1fr)",
          gap: 16,
          marginTop: 24,
          maxWidth: 1500,
        }}
      >
        <div style={cardStyle}>
          <b>ToDo XLSX</b>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => setTodoFile(e.target.files?.[0] || null)}
            style={inputStyle}
          />
        </div>

        <div style={cardStyle}>
          <b>Report_Completo.xlsx</b>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => setReportFile(e.target.files?.[0] || null)}
            style={inputStyle}
          />
        </div>

        <div style={cardStyle}>
          <b>ELENCO_ELABORATI.xlsx</b>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => setElencoFile(e.target.files?.[0] || null)}
            style={inputStyle}
          />
        </div>

        <div style={cardStyle}>
          <b>BCFZIP / ZIP</b>
          <input
            type="file"
            accept=".bcfzip,.zip"
            multiple
            onChange={(e) => setBcfFiles(Array.from(e.target.files || []))}
            style={inputStyle}
          />
          <div style={{ marginTop: 8, color: "#64748b", fontSize: 12 }}>
            Opzionale, ma consigliato per verificare la storia dei rilievi.
          </div>
          {bcfFiles.length > 0 && (
            <div style={{ marginTop: 8, color: "#64748b", fontSize: 12 }}>
              File BCF caricati: {bcfFiles.length}
            </div>
          )}
        </div>
      </div>

      <button
        onClick={eseguiControllo}
        disabled={loading}
        style={{
          marginTop: 20,
          padding: "14px 22px",
          background: loading ? "#334155" : "#0f172a",
          color: "white",
          border: "none",
          borderRadius: 12,
          fontWeight: 700,
          cursor: loading ? "not-allowed" : "pointer",
        }}
      >
        {loading ? "Controllo in corso..." : "Esegui controllo ToDo"}
      </button>

      {totale > 0 && (
        <>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(6, 1fr)",
              gap: 12,
              marginTop: 28,
            }}
          >
            <div style={cardStyle}>
              <div style={kpiLabel}>Completezza rilievi ispettori</div>
              <div
                style={{
                  fontSize: 36,
                  fontWeight: 800,
                  color: completezzaColor,
                }}
              >
                {completezza}%
              </div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>Righe ToDo controllate</div>
              <div style={kpiValue}>{totale}</div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>Righe corrette</div>
              <div style={{ ...kpiValue, color: "#16a34a" }}>{ok}</div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>Righe con errori</div>
              <div style={{ ...kpiValue, color: "#dc2626" }}>{errori}</div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>Storie complete</div>
              <div style={{ ...kpiValue, color: "#16a34a" }}>{storieComplete}</div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>BCF/riscontri mancanti</div>
              <div style={{ ...kpiValue, color: "#dc2626" }}>
                {bcfNonTrovati + mancanoRisposte}
              </div>
            </div>
          </div>

          <div style={{ marginTop: 28, ...cardStyle }}>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                gap: 16,
                marginBottom: 16,
              }}
            >
              <div>
                <h2 style={{ margin: 0 }}>Report controllo ToDo</h2>
                <div style={{ marginTop: 6, color: "#64748b", fontSize: 13 }}>
                  Righe visualizzate: {filteredChecks.length} su {checks.length}
                </div>
              </div>

              <button
                onClick={esportaExcel}
                style={{
                  padding: "10px 16px",
                  background: "#16a34a",
                  color: "white",
                  border: "none",
                  borderRadius: 10,
                  fontWeight: 700,
                  cursor: "pointer",
                }}
              >
                Esporta Excel
              </button>
            </div>

            <div style={{ overflowX: "auto" }}>
              <table
                style={{
                  width: "100%",
                  borderCollapse: "collapse",
                  fontSize: 13,
                }}
              >
                <thead>
                  <tr style={{ background: "#f8fafc" }}>
                    <th style={th}>N.</th>
                    <th style={th}>TR</th>
                    <th style={th}>Codice elaborato Report</th>
                    <th style={th}>Codice elaborato nel Title Trimble</th>
                    <th style={th}>Esito codice</th>
                    <th style={th}>Tags</th>
                    <th style={th}>Disciplina</th>
                    <th style={th}>Status</th>
                    <th style={th}>Description ToDo</th>
                    <th style={th}>Risposta progettista</th>
                    <th style={th}>Riscontro ITS</th>
                    <th style={th}>Storia rilievo</th>
                    <th style={th}>Esito</th>
                    <th style={th}>Anomalie</th>
                  </tr>

                  <tr style={{ background: "white" }}>
                    <th style={th}>
                      <input
                        value={filters.n}
                        onChange={(e) => updateFilter("n", e.target.value)}
                        placeholder="Filtra N."
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.tr}
                        onChange={(e) => updateFilter("tr", e.target.value)}
                        placeholder="Filtra TR"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.codiceReport}
                        onChange={(e) =>
                          updateFilter("codiceReport", e.target.value)
                        }
                        placeholder="Filtra codice"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.codiceTitleTrimble}
                        onChange={(e) =>
                          updateFilter("codiceTitleTrimble", e.target.value)
                        }
                        placeholder="Filtra title"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <select
                        value={filters.esitoCodice}
                        onChange={(e) =>
                          updateFilter("esitoCodice", e.target.value)
                        }
                        style={filterInput}
                      >
                        <option value="">Tutti</option>
                        <option value="OK">OK</option>
                        <option value="ERRORE">ERRORE</option>
                      </select>
                    </th>

                    <th style={th}>
                      <input
                        value={filters.tags}
                        onChange={(e) => updateFilter("tags", e.target.value)}
                        placeholder="Filtra tags"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.disciplina}
                        onChange={(e) =>
                          updateFilter("disciplina", e.target.value)
                        }
                        placeholder="Filtra disciplina"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.status}
                        onChange={(e) => updateFilter("status", e.target.value)}
                        placeholder="Filtra status"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}></th>
                    <th style={th}></th>
                    <th style={th}></th>
                    <th style={th}>
                      <select
                        value={filters.esitoStoria}
                        onChange={(e) =>
                          updateFilter("esitoStoria", e.target.value)
                        }
                        style={filterInput}
                      >
                        <option value="">Tutte</option>
                        <option value="COMPLETA">COMPLETA</option>
                        <option value="BCF NON TROVATO">BCF NON TROVATO</option>
                        <option value="MANCA RISPOSTA PROGETTISTA">
                          MANCA RISPOSTA PROGETTISTA
                        </option>
                        <option value="MANCA RISCONTRO ITS">
                          MANCA RISCONTRO ITS
                        </option>
                        <option value="CHIUSO SENZA RISCONTRO">
                          CHIUSO SENZA RISCONTRO
                        </option>
                        <option value="NON APPLICABILE">NON APPLICABILE</option>
                      </select>
                    </th>

                    <th style={th}>
                      <select
                        value={filters.esito}
                        onChange={(e) => updateFilter("esito", e.target.value)}
                        style={filterInput}
                      >
                        <option value="">Tutti</option>
                        <option value="OK">OK</option>
                        <option value="ERRORE">ERRORE</option>
                      </select>
                    </th>

                    <th style={th}>
                      <input
                        value={filters.anomalie}
                        onChange={(e) =>
                          updateFilter("anomalie", e.target.value)
                        }
                        placeholder="Filtra anomalie"
                        style={filterInput}
                      />
                    </th>
                  </tr>
                </thead>

                <tbody>
                  {filteredChecks.map((row) => (
                    <tr key={row.rowNumber}>
                      <td style={td}>
                        {row.progressivo}
                        {row.label ? ` (${row.label})` : ""}
                      </td>

                      <td style={td}>{row.tr || "-"}</td>

                      <td style={td}>{row.codiceReport || ""}</td>

                      <td style={td}>{row.codiceTitleTrimble || "-"}</td>

                      <td style={td}>
                        <b
                          style={{
                            color: row.titleOk ? "#16a34a" : "#dc2626",
                          }}
                        >
                          {esitoIcon(row.titleOk)}
                        </b>
                      </td>

                      <td style={td}>
                        {row.tags || "-"}{" "}
                        <b
                          style={{
                            color: row.tagsOk ? "#16a34a" : "#dc2626",
                          }}
                        >
                          {esitoIcon(row.tagsOk)}
                        </b>
                      </td>

                      <td style={td}>
                        {row.disciplina || "-"}{" "}
                        <b
                          style={{
                            color: row.disciplinaOk ? "#16a34a" : "#dc2626",
                          }}
                        >
                          {esitoIcon(row.disciplinaOk)}
                        </b>
                      </td>

                      <td style={td}>
                        {row.status || "-"}{" "}
                        <b
                          style={{
                            color: row.statusOk ? "#16a34a" : "#dc2626",
                          }}
                        >
                          {esitoIcon(row.statusOk)}
                        </b>
                      </td>

                      <td style={{ ...td, minWidth: 260, whiteSpace: "pre-wrap" }}>
                        {row.description || "-"}
                      </td>

                      <td style={{ ...td, minWidth: 260, whiteSpace: "pre-wrap" }}>
                        {row.rispostaProgettista || "-"}
                      </td>

                      <td style={{ ...td, minWidth: 260, whiteSpace: "pre-wrap" }}>
                        {row.riscontroIspettore || "-"}
                      </td>

                      <td
                        style={{
                          ...td,
                          fontWeight: 700,
                          color: row.storiaOk ? "#16a34a" : "#dc2626",
                        }}
                      >
                        {row.esitoStoria}
                      </td>

                      <td
                        style={{
                          ...td,
                          fontWeight: 700,
                          color: row.esito === "OK" ? "#16a34a" : "#dc2626",
                        }}
                      >
                        {row.esito}
                      </td>

                      <td style={td}>{row.anomalie.join(" | ")}</td>
                    </tr>
                  ))}

                  {filteredChecks.length === 0 && (
                    <tr>
                      <td style={td} colSpan={14}>
                        Nessuna riga trovata con i filtri impostati.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}
    </main>
  );
}

const cardStyle: React.CSSProperties = {
  background: "white",
  border: "1px solid #e2e8f0",
  borderRadius: 16,
  padding: 16,
  boxShadow: "0 6px 18px rgba(15,23,42,.06)",
};

const inputStyle: React.CSSProperties = {
  display: "block",
  width: "100%",
  marginTop: 10,
};

const filterInput: React.CSSProperties = {
  width: "100%",
  padding: "6px 8px",
  border: "1px solid #cbd5e1",
  borderRadius: 6,
  fontSize: 12,
  boxSizing: "border-box",
};

const kpiLabel: React.CSSProperties = {
  fontSize: 13,
  color: "#64748b",
};

const kpiValue: React.CSSProperties = {
  fontSize: 36,
  fontWeight: 800,
  color: "#0f172a",
};

const th: React.CSSProperties = {
  border: "1px solid #e2e8f0",
  padding: 8,
  textAlign: "left",
};

const td: React.CSSProperties = {
  border: "1px solid #e2e8f0",
  padding: 8,
  verticalAlign: "top",
};
