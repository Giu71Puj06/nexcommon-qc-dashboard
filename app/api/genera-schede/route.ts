import { NextResponse } from "next/server";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

const FASE_PROGETTO = "Progetto Esecutivo";
const REVISIONE_SCHEDA = "0";

const ISPETTORE_ELENCO_COLUMNS = [
  "Ispettore",
  "ISPETTORE",
  "Nome ispettore",
  "Nome Ispettore",
  "Nome_Ispettore",
  "Verificatore",
  "VERIFICATORE",
  "Nome verificatore",
  "Nome Verificatore",
  "Nome_verificatore",
];

const ISPETTORE_TODO_COLUMNS = [
  "Created by",
  "CREATED BY",
  "Creato da",
  "Autore",
  "Author",
  "Ispettore",
  "ISPETTORE",
  "Nome ispettore",
  "Nome Ispettore",
  "Nome_Ispettore",
  "Verificatore",
  "Nome verificatore",
];

const TITOLO_ELABORATO_COLUMNS = [
  "Titolo elenco",
  "Titolo Elenco",
  "TITOLO ELENCO",
  "Titolo elaborato",
  "Titolo Elaborato",
  "TITOLO ELABORATO",
  "Nome elaborato",
  "Nome Elaborato",
];

const REPORT_CODICE_COLUMNS = [
  "Codice elenco",
  "Codice Elenco",
  "CODICE ELENCO",
  "Codice cartiglio",
  "Codice Cartiglio",
  "CODICE CARTIGLIO",
  "Nome file PDF",
  "Nome File PDF",
  "NOME FILE PDF",
  "Codice elaborato",
  "Codice Elaborato",
  "CODICE ELABORATO",
  "Codice",
  "CODICE",
];

type BcfTopicData = {
  topicGuid: string;
  titolo: string;
  descrizione: string;
  descrizioneConData: string;
  ispettore: string;
  ispettoreNomeBcf: string;
  labels: string;
  stato: string;
  commentiPRG: string;
  commentiISP: string;
  ultimoCommento: string;
};

type CommentEntry = {
  author: string;
  date: string;
  text: string;
  order: number;
};

type ElaboratoVerificatoRow = {
  codice_elaborato: string;
  codice_file: string;
  revisione: string;
  titolo_elaborato: string;
  disciplina: string;
  presenza_nc: string;
  presenza_oss: string;
  assenza_nc_oss: string;
};

type SchedaIspettivaSintesi = {
  totaleElaboratiAnalizzati: number;
  totaleNC: number;
  totaleOSS: number;
  totaleChiuse: number;
};

type RevisioneSchedaRow = {
  rev: string;
  data: string;
  descrizione: string;
  responsabile_pcq: string;
  responsabile_its: string;
};

type StoricoRilievoRow = {
  tr: string;
  tipoBase: string;
  descrizioneRilievo: string;
};


function descrizioneRevisioneScheda(rev: string) {
  const n = Number(String(rev || "0").trim());

  if (!Number.isFinite(n) || n <= 0) return "Prima Emissione - Rilievi";
  if (n === 1) return "Seconda Emissione - Riscontri";
  if (n === 2) return "Terza Emissione - Riscontri";
  if (n === 3) return "Quarta Emissione - Riscontri";
  if (n === 4) return "Quinta Emissione - Riscontri";

  return `${n + 1}ª Emissione - Riscontri`;
}

function clampRevisioneScheda(value: string) {
  const n = Number(String(value || "0").trim());
  if (!Number.isFinite(n) || n < 0) return 0;
  return Math.min(Math.floor(n), 4);
}

function readFormString(formData: FormData, name: string) {
  return String(formData.get(name) || "").trim();
}

function buildRevisioniSchedaRows(
  formData: FormData,
  revisioneScheda: string,
  dataRevisioneScheda: string,
  dataRiscontroIspettore: string,
  responsabilePcq: string,
  responsabileIts: string
): RevisioneSchedaRow[] {
  const currentRev = clampRevisioneScheda(revisioneScheda);
  const rows: RevisioneSchedaRow[] = [];

  for (let rev = 0; rev <= currentRev; rev += 1) {
    let data =
      readFormString(formData, `data_rev_${rev}`) ||
      (rev === currentRev ? dataRevisioneScheda : "");

    // Per le revisioni successive alla prima, se non viene compilata la data specifica,
    // usa la data del riscontro ispettore inserita dal PM.
    if (!data && rev > 0) data = dataRiscontroIspettore;

    rows.push({
      rev: String(rev),
      data,
      descrizione: descrizioneRevisioneScheda(String(rev)),
      responsabile_pcq: responsabilePcq,
      responsabile_its: responsabileIts,
    });
  }

  return rows;
}


function safeName(value: string) {
  return String(value || "SENZA_DISCIPLINA")
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, "_");
}

function normalizeKey(value: string) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "")
    .replace(/[^A-Z0-9]/g, "");
}

function normalizeText(value: string) {
  return String(value || "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeAccount(value: string) {
  return String(value || "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9@._ -]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function parsePeopleList(value: FormDataEntryValue | null) {
  return String(value || "")
    .split(/[\n,;]+/)
    .map(normalizeAccount)
    .filter(Boolean);
}

function getCommentAuthor(c: any) {
  return String(
    c?.Author ||
      c?.CreationAuthor ||
      c?.ModifiedAuthor ||
      c?.["@_Author"] ||
      c?.["@_CreationAuthor"] ||
      c?.["@_ModifiedAuthor"] ||
      ""
  ).trim();
}

function getCommentDateValue(c: any) {
  return String(
    c?.Date ||
      c?.CreationDate ||
      c?.ModifiedDate ||
      c?.["@_Date"] ||
      c?.["@_CreationDate"] ||
      c?.["@_ModifiedDate"] ||
      ""
  ).trim();
}

function getTopicDateValue(topic: any) {
  return String(
    topic?.CreationDate ||
      topic?.Date ||
      topic?.ModifiedDate ||
      topic?.["@_CreationDate"] ||
      topic?.["@_Date"] ||
      topic?.["@_ModifiedDate"] ||
      ""
  ).trim();
}

function formatBcfCommentDate(value: string) {
  const raw = String(value || "").trim();
  if (!raw) return "";

  const datePart = raw.split("T")[0].split(" ")[0];

  const isoMatch = datePart.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) return `${isoMatch[3]}/${isoMatch[2]}/${isoMatch[1]}`;

  const italianMatch = datePart.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
  if (italianMatch) {
    const day = italianMatch[1].padStart(2, "0");
    const month = italianMatch[2].padStart(2, "0");
    const year = italianMatch[3].length === 2 ? `20${italianMatch[3]}` : italianMatch[3];
    return `${day}/${month}/${year}`;
  }

  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) {
    const day = String(parsed.getDate()).padStart(2, "0");
    const month = String(parsed.getMonth() + 1).padStart(2, "0");
    const year = parsed.getFullYear();
    return `${day}/${month}/${year}`;
  }

  return raw;
}

function prefixCommentWithDate(text: string, dateValue: string) {
  const cleanText = String(text || "").trim();
  if (!cleanText) return "";

  const formattedDate = formatBcfCommentDate(dateValue);
  if (!formattedDate) return cleanText;

  if (cleanText.startsWith(`${formattedDate} - `)) return cleanText;
  if (cleanText.startsWith(`[${formattedDate}]`)) return cleanText;
  if (cleanText.startsWith(`${formattedDate}\n`)) return cleanText;

  return `${formattedDate}\n${cleanText}`;
}

function isStandaloneItalianDate(value: string) {
  return /^\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}$/.test(String(value || "").trim());
}

function coalesceCommentBlocksByDate(text: string) {
  const cleanText = String(text || "").trim();
  if (!cleanText) return "";

  const groups: Record<string, string[]> = {};
  const orderedKeys: string[] = [];
  let undatedCounter = 0;

  cleanText
    .split(/\n{2,}/g)
    .map((block) => block.trim())
    .filter(Boolean)
    .forEach((block) => {
      const lines = block.split(/\n/g);
      const firstLine = lines[0]?.trim() || "";
      const hasDate = isStandaloneItalianDate(firstLine);
      const key = hasDate ? `DATE__${firstLine}` : `NO_DATE__${undatedCounter++}`;
      const body = hasDate ? lines.slice(1).join("\n").trim() : block;

      if (!body) return;

      if (!groups[key]) {
        groups[key] = [];
        orderedKeys.push(key);
      }

      if (!groups[key].includes(body)) groups[key].push(body);
    });

  return orderedKeys
    .map((key) => {
      const body = groups[key].join("\n\n").trim();
      if (!body) return "";

      if (key.startsWith("DATE__")) {
        const date = key.replace(/^DATE__/, "");
        return `${date}\n${body}`;
      }

      return body;
    })
    .filter(Boolean)
    .join("\n\n");
}

function buildCommentBlocks(entries: CommentEntry[]) {
  const groups: Record<string, CommentEntry[]> = {};
  const orderedKeys: string[] = [];

  entries
    .filter((entry) => String(entry.text || "").trim())
    .sort((a, b) => a.order - b.order)
    .forEach((entry) => {
      // REGOLA CORRETTA:
      // nella stessa NC/OSS la data deve comparire una sola volta.
      // Se piu' commenti hanno la stessa data, vengono accorpati sotto quella data.
      const dateKey = entry.date || `SENZA_DATA_${entry.order}`;
      const key = `DATE__${dateKey}`;

      if (!groups[key]) {
        groups[key] = [];
        orderedKeys.push(key);
      }

      groups[key].push(entry);
    });

  const text = orderedKeys
    .map((key) => {
      const blockEntries = groups[key];
      const date = blockEntries[0]?.date || "";
      const body = Array.from(
        new Set(blockEntries.map((entry) => String(entry.text || "").trim()).filter(Boolean))
      ).join("\n\n");

      if (!body) return "";
      return date ? `${date}\n${body}` : body;
    })
    .filter(Boolean)
    .join("\n\n");

  return coalesceCommentBlocksByDate(text);
}

function mergeCommentBlocks(existing: string, next: string) {
  const values = [existing || "", next || ""]
    .map((value) => String(value || "").trim())
    .filter(Boolean);

  return coalesceCommentBlocksByDate(Array.from(new Set(values)).join("\n\n"));
}

function readIndexedFormDates(formData: FormData, prefix: string, fallback: string) {
  const values: string[] = [];

  for (let i = 0; i < 10; i += 1) {
    const value = readFormString(formData, `${prefix}_${i}`);
    if (value) values.push(value);
  }

  if (fallback) values.push(fallback);

  return Array.from(new Set(values));
}

function applyFallbackDatesToUndatedBlocks(text: string, dates: string[]) {
  const cleanText = String(text || "").trim();
  if (!cleanText) return "";

  const availableDates = dates.map((d) => String(d || "").trim()).filter(Boolean);
  if (availableDates.length === 0) return coalesceCommentBlocksByDate(cleanText);

  const blocks = cleanText.split(/\n{2,}/g).map((block) => block.trim()).filter(Boolean);
  let dateIndex = 0;

  const withDates = blocks
    .map((block) => {
      const firstLine = block.split(/\n/g)[0]?.trim() || "";
      const alreadyDated = isStandaloneItalianDate(firstLine);
      if (alreadyDated) return block;

      const date = availableDates[Math.min(dateIndex, availableDates.length - 1)];
      dateIndex += 1;
      return `${date}\n${block}`;
    })
    .join("\n\n");

  return coalesceCommentBlocksByDate(withDates);
}

function stripLeadingCommentDate(text: string) {
  return String(text || "")
    .replace(/^\s*\d{4}-\d{2}-\d{2}\s*[-–]\s*/g, "")
    .replace(/^\s*\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}\s*[-–]\s*/g, "")
    .replace(/^\s*\[\s*\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}\s*\]\s*/g, "")
    .trim();
}

function prefixCommentWithPmDate(text: string, pmDate: string) {
  const cleanText = stripLeadingCommentDate(text);
  const cleanDate = String(pmDate || "").trim();

  if (!cleanText) return "";
  if (!cleanDate) return cleanText;

  return `${cleanDate} - ${cleanText}`;
}

function applyPmDateToMultilineComments(text: string, pmDate: string) {
  const cleanText = String(text || "").trim();
  if (!cleanText) return "";

  return cleanText
    .split(/\n+/g)
    .map((line) => prefixCommentWithPmDate(line, pmDate))
    .filter(Boolean)
    .join("\n");
}

function isInPeopleList(author: string, people: string[]) {
  const a = normalizeAccount(author);
  if (!a) return false;
  return people.some((p) => a === p || a.includes(p) || p.includes(a));
}

function findValue(row: any, names: string[]) {
  for (const name of names) {
    if (
      row?.[name] !== undefined &&
      row?.[name] !== null &&
      String(row[name]).trim() !== ""
    ) {
      return String(row[name]).trim();
    }
  }
  return "";
}

function getTitoloProgetto(row: any) {
  return findValue(row, [
    "Titolo_progetto",
    "Titiolo_progetto",
    "Titolo progetto",
    "Titolo Progetto",
    "TITOLO PROGETTO",
  ]);
}

function getFaseProgetto(row: any) {
  return (
    findValue(row, [
      "Fase_di_progetto",
      "Fase_di_ progetto",
      "Fase di progetto",
      "Fase Progetto",
      "FASE DI PROGETTO",
    ]) || FASE_PROGETTO
  );
}

function getIspettoreFromTodo(row: any) {
  return findValue(row, ISPETTORE_TODO_COLUMNS);
}

function getTitoloElaboratoFromTodo(row: any) {
  return findValue(row, TITOLO_ELABORATO_COLUMNS);
}

function cleanCodiceElaborato(value: string) {
  return String(value || "")
    .replace(/\.pdf$/i, "")
    .replace(/[.,;:]+$/g, "")
    .trim();
}

function extractCodiceElaborato(value: string) {
  const text = cleanCodiceElaborato(value);
  const match = text.match(/[A-Z0-9]{2,}(?:[-_][A-Z0-9]+){5,}(?:[-_]\d{1,2})?/i);
  return match ? cleanCodiceElaborato(match[0]) : text;
}

function getElaboratoBase(value: string) {
  return cleanCodiceElaborato(value)
    .replace(/\.[A-Za-z0-9]+$/g, "")
    .replace(/[_-]\d{1,2}$/g, "")
    .trim();
}

function sameElaboratoCode(a: string, b: string) {
  const aa = normalizeKey(getElaboratoBase(extractCodiceElaborato(a)));
  const bb = normalizeKey(getElaboratoBase(extractCodiceElaborato(b)));
  return aa && bb && aa === bb;
}

function getRevisioneDaCodice(value: string) {
  const clean = cleanCodiceElaborato(value);
  const match = clean.match(/[_-](\d{1,2})$/);
  return match ? match[1] : "";
}

function labelToTR(label: string) {
  const value = String(label || "").trim();
  const match = value.match(/-(\d+[A-Z]?)$/i);
  if (match) return `TR-${match[1]}`;
  return value.replace(/^.{5}/, "TR-");
}

function extractTR(value: string) {
  const text = String(value || "").toUpperCase();
  const match = text.match(/TR[-_\s]?(\d+[A-Z]?)/i);
  if (!match) return "";
  return `TR-${match[1]}`;
}

function normalizeTR(value: string) {
  const text = String(value || "").toUpperCase();

  const trMatch = text.match(/TR[-_\s]?0*(\d+[A-Z]?)/i);
  if (trMatch) return `TR-${trMatch[1]}`;

  const itMatch = text.match(/IT\d+[-_\s]?0*(\d+[A-Z]?)/i);
  if (itMatch) return `TR-${itMatch[1]}`;

  return "";
}


function cleanRolePrefix(text: string) {
  return String(text || "")
    .replace(/\(\s*ISP\s*\)/gi, "")
    .replace(/\(\s*PRG\s*\)/gi, "")
    .trim();
}

function mergeText(existing: string, additions: string[]) {
  const values = [existing || "", ...additions]
    .map((v) => String(v || "").trim())
    .filter(Boolean);

  return Array.from(new Set(values)).join("\n");
}

function extractAcronimiIspettoriUfficiali(nomeRedattore: string) {
  return Array.from(
    new Set(
      String(nomeRedattore || "")
        .match(/\([A-Z]{2,4}\)/g)
        ?.map((value) => value.trim().toUpperCase()) || []
    )
  );
}

function firstAllowedAcronimo(allowed: string[], fallback = "") {
  return allowed.length > 0 ? allowed[0] : fallback;
}

function remapIspettoreByDisciplina(sigla: string, disciplina: string) {
  const s = String(sigla || "").trim().toUpperCase();
  const d = normalizeKey(disciplina);

  if (d === "SICUREZZACANTIERE") {
    if (s === "(CR)" || s === "(CA)" || s === "(GB)") return "(FM)";
  }

  if (d === "DOCUMENTAZIONEECONOMICA") {
    if (s === "(CS)" || s === "(CR)" || s === "(CA)" || s === "(GB)") return "(MG)";
  }

  return sigla;
}

function remapIspettoreFinale(sigla: string, disciplina: string) {
  const s = String(sigla || "").trim().toUpperCase();
  const d = normalizeKey(disciplina);

  if (d === "DOCUMENTAZIONEECONOMICA") {
    if (s === "(CS)" || s === "(CR)" || s === "(CA)" || s === "(GB)") return "(MG)";
  }

  if (d === "IMPIANTI") {
    if (s === "(CR)" || s === "(CA)" || s === "(GB)") return "(OB)";
  }

  if (d === "DOCUMENTAZIONEGENERALE" || d === "DOCUMENTIGENERALI") {
    if (s === "(CS)") return "(MG)";
    if (s === "(CR)" || s === "(CA)" || s === "(GB)") return "(OB)";
  }

  if (s === "(CR)" || s === "(CA)") return "(OB)";
  if (d === "SICUREZZACANTIERE" && s === "(GB)") return "(FM)";

  return remapIspettoreByDisciplina(sigla, disciplina);
}

function remapIspettoreUfficiale(
  sigla: string,
  disciplina: string,
  nomeRedattore: string
) {
  const allowed = extractAcronimiIspettoriUfficiali(nomeRedattore);
  const mapped = remapIspettoreFinale(sigla, disciplina);
  const normalized = String(mapped || "").trim().toUpperCase();

  if (!normalized) return firstAllowedAcronimo(allowed, "");
  if (allowed.length === 0) return normalized;
  if (allowed.includes(normalized)) return normalized;

  const d = normalizeKey(disciplina);

  if (d === "DOCUMENTAZIONEECONOMICA" && allowed.includes("(MG)")) return "(MG)";
  if (d === "IMPIANTI" && allowed.includes("(OB)")) return "(OB)";

  if (d === "DOCUMENTAZIONEGENERALE" || d === "DOCUMENTIGENERALI") {
    if ((normalized === "(CS)" || normalized === "(MG)") && allowed.includes("(MG)")) return "(MG)";
    if ((normalized === "(CR)" || normalized === "(CA)" || normalized === "(GB)" || normalized === "(OB)") && allowed.includes("(OB)")) return "(OB)";
  }

  return firstAllowedAcronimo(allowed, normalized);
}

function siglaDaNome(nome: string) {
  const value = String(nome || "")
    .replace(/\bArch\.?\b/gi, "")
    .replace(/\bIng\.?\b/gi, "")
    .replace(/\bGeom\.?\b/gi, "")
    .replace(/\bDott\.?\b/gi, "")
    .replace(/\bP\.?\s*I\.?\b/gi, "")
    .replace(/\./g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const parts = value.split(" ").filter(Boolean);
  if (parts.length === 0) return "";

  const sigla =
    parts.length === 1
      ? parts[0].slice(0, 2).toUpperCase()
      : `${parts[0][0] || ""}${parts[parts.length - 1][0] || ""}`.toUpperCase();

  return sigla ? `(${sigla})` : "";
}

function getCommentText(c: any) {
  const value = c?.Comment || c?.CommentText || c?.Text || "";
  return typeof value === "string" ? value : String(value || "");
}

function isGiuseppePizzi(value: string) {
  const normalized = normalizeText(value);
  return (
    normalized === "GIUSEPPE PIZZI" ||
    normalized.includes("GIUSEPPE PIZZI") ||
    normalized === "PIZZI GIUSEPPE"
  );
}

function isIspettoreGiuseppePizzi(sigla: string, nomeBcf: string) {
  if (isGiuseppePizzi(nomeBcf)) return true;
  return String(sigla || "").trim().toUpperCase() === "(GP)";
}

function resolveIspettoreFinale(
  ispettoreBcf: string,
  nomeBcf: string,
  disciplina: string,
  ispettoreSostitutivo: string
) {
  const ispettoreRemappato = remapIspettoreFinale(ispettoreBcf || "", disciplina);

  if (isIspettoreGiuseppePizzi(ispettoreBcf, nomeBcf) && ispettoreSostitutivo) {
    return remapIspettoreFinale(siglaDaNome(ispettoreSostitutivo), disciplina);
  }

  return ispettoreRemappato;
}

function isDaNcAOss(tags: string, status?: string) {
  const text = normalizeText(`${tags || ""} ${status || ""}`);

  // Riconosce solo indicazioni esplicite di declassamento.
  // Evita falsi positivi generici su stringhe che contengono solo NC o OSS.
  return (
    text.includes("DA NC A OSS") ||
    text.includes("DA NC AD OSS") ||
    text.includes("NC A OSS") ||
    text.includes("NC AD OSS") ||
    text.includes("NC DECLASSATA A OSS") ||
    text.includes("NC DECLASSATO A OSS") ||
    text.includes("DECLASSATA AD OSSERVAZIONE") ||
    text.includes("DECLASSATO AD OSSERVAZIONE") ||
    text.includes("DECLASSATA A OSSERVAZIONE") ||
    text.includes("DECLASSATO A OSSERVAZIONE")
  );
}

function determineRilievoStatus(
  status: string,
  tags: string,
  riscontroIspettore: string,
  rispostaProgettista: string = ""
) {
  const statusText = normalizeText(status);
  const tagsText = normalizeText(tags);
  const riscontroText = normalizeText(riscontroIspettore);
  const rispostaText = normalizeText(rispostaProgettista);
  const allText = normalizeText(`${status || ""} ${tags || ""} ${riscontroIspettore || ""} ${rispostaProgettista || ""}`);

  // Caso specifico: una NC declassata ad OSS resta evidenziata come aperta/nera,
  // ma deve essere conteggiata e marcata come OSS nella tabella finale.
  if (isDaNcAOss(tags, status) || isDaNcAOss(riscontroIspettore, rispostaProgettista)) {
    return "Da NC a OSS";
  }

  // Prima intercettiamo le formule negative, perché contengono spesso la parola "risolta".
  // Esempi: "Non conformità non risolta", "NC permane", "non si condivide".
  if (
    allText.includes("NON CONFORMITA NON RISOLTA") ||
    allText.includes("NC NON RISOLTA") ||
    allText.includes("NON RISOLTA") ||
    allText.includes("NON RISOLTO") ||
    allText.includes("NON SUPERATA") ||
    allText.includes("NC PERMANE") ||
    allText.includes("PERMANE LA NON CONFORMITA") ||
    allText.includes("PERMANE IL RILIEVO") ||
    allText.includes("PERMANE") ||
    allText.includes("NON SI CONDIVIDE") ||
    allText.includes("NON CONDIVIDE") ||
    allText.includes("CHIARIMENTO NON RICEVUTO") ||
    allText.includes("SI CHIEDE DI FORNIRE RISPOSTA") ||
    allText.includes("SI RIBADISCE") ||
    allText.includes("SI RIMANDA ALLA VALUTAZIONE") ||
    allText.includes("DA DISCUTERE IN SEDE DI CONTRADDITTORIO")
  ) {
    return "Aperta";
  }

  if (
    statusText.includes("CLOSED") ||
    statusText.includes("CLOSE") ||
    statusText.includes("CHIUSA") ||
    statusText.includes("CHIUSO") ||
    statusText.includes("CHIUSE") ||
    statusText.includes("CHIUSI") ||
    riscontroText.includes("NON CONFORMITA RISOLTA") ||
    riscontroText.includes("NC RISOLTA") ||
    riscontroText.includes("NC RISOLTO") ||
    riscontroText.includes("NC SUPERATA") ||
    riscontroText.includes("NON CONFORMITA SUPERATA") ||
    riscontroText.includes("OSSERVAZIONE RISOLTA") ||
    riscontroText.includes("OSS SUPERATA") ||
    riscontroText.includes("OSSERVAZIONE SUPERATA") ||
    riscontroText.includes("SI CONDIVIDE IL CHIARIMENTO") ||
    riscontroText.includes("SI PRENDE ATTO DEL CHIARIMENTO") ||
    riscontroText.includes("SI CHIUDE LA NC") ||
    riscontroText.includes("RILIEVO CHIUSO") ||
    riscontroText.includes("RILIEVO SUPERATO") ||
    riscontroText.includes("RISOLTA") ||
    riscontroText.includes("RISOLTO") ||
    riscontroText.includes("SUPERATA") ||
    riscontroText.includes("SUPERATO")
  ) {
    return "Chiusa";
  }

  if (
    statusText.includes("OPEN") ||
    statusText.includes("APERTA") ||
    statusText.includes("APERTO") ||
    statusText.includes("NEW") ||
    statusText.includes("DA VERIFICARE") ||
    statusText.includes("IN ATTESA") ||
    statusText.includes("WAITING")
  ) {
    return "Aperta";
  }

  return status || "Aperta";
}

function normalizeStatus(status: string, tags: string) {
  return determineRilievoStatus(status, tags, "");
}

function isDaNcAOssStatus(status: string) {
  const s = normalizeText(status);
  return s.includes("DA NC A OSS") || s.includes("NC DECLASSATA A OSS") || s.includes("NC DECLASSATO A OSS");
}

function isClosedStatus(status: string) {
  const s = normalizeText(status);
  return s === "CHIUSA" || s.includes("CLOSED") || s.includes("CHIUS");
}

function isOpenStatus(status: string) {
  const s = normalizeText(status);

  if (!s) return false;
  if (isClosedStatus(s)) return false;

  // Le NC declassate ad OSS non vanno in grigio e devono restare considerate aperte
  // ai fini della X nella colonna OSS.
  if (isDaNcAOssStatus(s)) return true;

  return (
    s.includes("OPEN") ||
    s.includes("APERTA") ||
    s.includes("APERTO") ||
    s.includes("APERTE") ||
    s.includes("APERTI") ||
    s.includes("NEW") ||
    s.includes("DA VERIFICARE") ||
    s.includes("IN ATTESA") ||
    s.includes("WAITING")
  );
}

function isOpenRilievoRow(row: any) {
  return isOpenStatus(row?.Stato || row?.Status || row?.stato || "");
}

function isNessunRilievo(tags: string, descrizione: string) {
  return `${tags || ""} ${descrizione || ""}`
    .toUpperCase()
    .includes("NESSUN RILIEVO");
}


function disciplinaFromCodice(codice: string) {
  const c = String(codice || "").toUpperCase();

  if (c.includes("_ARC_") || c.includes("-ARC-") || c.includes("ARC")) return "Architettonico";
  if (c.includes("_STR_") || c.includes("-STR-") || c.includes("STR")) return "Strutturale";
  if (c.includes("_MEP_") || c.includes("-MEP-") || c.includes("MEP")) return "Impianti";
  if (c.includes("_IMP_") || c.includes("-IMP-") || c.includes("IMP")) return "Impianti";
  if (c.includes("_ECO_") || c.includes("-ECO-") || c.includes("ECO")) return "Documentazione economica";
  if (c.includes("_SIC_") || c.includes("-SIC-") || c.includes("SIC")) return "Sicurezza Cantiere";
  if (c.includes("_GEN_") || c.includes("-GEN-") || c.includes("GEN")) return "Documentazione generale";
  if (c.includes("_AMB_") || c.includes("-AMB-") || c.includes("AMB")) return "Documentazione generale";

  return "";
}

function disciplinaFromReportCartella(value: string) {
  const v = normalizeText(value);

  if (v.includes("GENERALE")) return "Documenti generali";
  if (v.includes("STRUTTURE")) return "Strutture";
  if (v.includes("AMBIENTE")) return "Ambiente e vincoli";
  if (v.includes("SICUREZZA")) return "Sicurezza cantierizzazione e BOB";
  if (v.includes("INTERFERENZE") || v.includes("ESPROPRI")) return "Interferenze e espropri";
  if (v.includes("ECONOMICO")) return "Economico";

  return "";
}

function sameDisciplina(a: string, b: string) {
  const aa = normalizeText(a);
  const bb = normalizeText(b);

  if (!aa || !bb) return false;

  if (aa === bb) return true;
  if (aa.includes(bb)) return true;
  if (bb.includes(aa)) return true;

  const ka = normalizeKey(a);
  const kb = normalizeKey(b);

  const aliases: Record<string, string[]> = {
    DOCUMENTAZIONEECONOMICA: [
      "ECONOMICO",
      "ECONOMICA",
      "DOCUMENTAZIONE ECONOMICA",
    ],
    ECONOMICO: ["DOCUMENTAZIONE ECONOMICA", "ECONOMICA", "ECONOMICO"],
    DOCUMENTAZIONEGENERALE: [
      "GENERALE",
      "DOCUMENTI GENERALI",
      "DOCUMENTAZIONE GENERALE",
      "DOCUMENTI GENERALI AMBIENTE E VINCOLI",
      "DOCUMENTI GENERALI STRUTTURE",
    ],
    DOCUMENTIGENERALI: [
      "GENERALE",
      "DOCUMENTAZIONE GENERALE",
      "DOCUMENTI GENERALI",
      "DOCUMENTI GENERALI AMBIENTE E VINCOLI",
      "DOCUMENTI GENERALI STRUTTURE",
    ],
    SICUREZZACANTIERE: [
      "SICUREZZA",
      "SICUREZZA CANTIERE",
      "SICUREZZA CANTIERIZZAZIONE E BOB",
      "SICUREZZACANTIERIZZAZIONEEBOB",
    ],
    SICUREZZACANTIERIZZAZIONEEBOB: [
      "SICUREZZA",
      "SICUREZZA CANTIERE",
      "SICUREZZA CANTIERIZZAZIONE E BOB",
    ],
    IMPIANTI: ["MEP", "IMPIANTISTICO", "IMPIANTI"],
    ARCHITETTONICO: ["ARCHITETTURA", "ARCHITETTONICO"],
    STRUTTURALE: ["STRUTTURE", "STRUTTURALE"],
    STRUTTURE: ["STRUTTURALE", "STRUTTURE", "DOCUMENTI GENERALI STRUTTURE"],
    AMBIENTEEVINCOLI: [
      "AMBIENTE",
      "VINCOLI",
      "AMBIENTE E VINCOLI",
      "DOCUMENTI GENERALI AMBIENTE E VINCOLI",
    ],
    INTERFERENZEEESPROPRI: [
      "INTERFERENZE",
      "ESPROPRI",
      "INTERFERENZE E ESPROPRI",
      "INTERFERENZE ED ESPROPRI",
    ],
  };

  const aliasesA = aliases[ka] || [];
  const aliasesB = aliases[kb] || [];

  return (
    aliasesA.some((x) => bb.includes(normalizeText(x))) ||
    aliasesB.some((x) => aa.includes(normalizeText(x)))
  );
}

function topicKey(title: string, description: string) {
  return `${normalizeKey(title)}__${normalizeKey(description).slice(0, 120)}`;
}

function descriptionScore(a: string, b: string) {
  const aa = normalizeText(a);
  const bb = normalizeText(b);
  if (!aa || !bb) return 0;
  if (aa === bb) return 1000;
  if (aa.includes(bb) || bb.includes(aa)) return 900;

  const aWords = new Set(aa.split(" ").filter((w) => w.length > 3));
  const bWords = new Set(bb.split(" ").filter((w) => w.length > 3));
  if (aWords.size === 0 || bWords.size === 0) return 0;

  let common = 0;
  aWords.forEach((w) => {
    if (bWords.has(w)) common += 1;
  });

  return common / Math.max(aWords.size, bWords.size);
}

function getElencoInfoByCode(elencoInfoMap: Record<string, any>, codice: string) {
  return (
    elencoInfoMap[normalizeKey(cleanCodiceElaborato(codice))] ||
    elencoInfoMap[normalizeKey(extractCodiceElaborato(codice))] ||
    elencoInfoMap[normalizeKey(getElaboratoBase(codice))] ||
    {}
  );
}

function findReportInfo(reportInfoMap: Record<string, any>, codice: string) {
  const estratto = extractCodiceElaborato(codice);
  return (
    reportInfoMap[normalizeKey(cleanCodiceElaborato(codice))] ||
    reportInfoMap[normalizeKey(estratto)] ||
    reportInfoMap[normalizeKey(getElaboratoBase(estratto))] ||
    {}
  );
}

function elaboratoAggregationKey(value: string) {
  return normalizeKey(getElaboratoBase(extractCodiceElaborato(value)));
}

function isRilievoOSS(row: any) {
  if (isDaNcAOss(row?.Tipo || row?.["Codice Rilievo"] || "", row?.Stato || "")) {
    return true;
  }

  const tipo = normalizeKey(row?.TipoBase || row?.Tipo || row?.["Codice Rilievo"] || "");
  return tipo.includes("OSS");
}

function buildRilieviFlagsByElaborato(rows: any[]) {
  const map: Record<string, { hasNC: boolean; hasOSS: boolean; sourceRow?: any }> = {};

  for (const row of rows) {
    // REGOLA CORRETTA:
    // nella tabella finale NC/OSS devono essere marcate solo se il rilievo e' aperto.
    // Se tutti i rilievi dell'elaborato sono chiusi, andra' marcata solo ASSENZA NC/OSS.
    if (!isOpenRilievoRow(row)) continue;

    const key = elaboratoAggregationKeyFromValues(
      row?.["Codice Elaborato"] || "",
      row?.["Titolo Elaborato"] || ""
    );
    if (!key) continue;

    if (!map[key]) map[key] = { hasNC: false, hasOSS: false, sourceRow: row };

    if (isRilievoOSS(row)) {
      map[key].hasOSS = true;
    } else {
      map[key].hasNC = true;
    }

    if (!map[key].sourceRow) map[key].sourceRow = row;
  }

  return map;
}

function applyRilieviFlagsToElaborati(
  elaborati: ElaboratoVerificatoRow[],
  rows: any[]
) {
  const flagsByElaborato = buildRilieviFlagsByElaborato(rows);
  const existingKeys = new Set<string>();

  const withFlags = elaborati.map((e) => {
    const key = elaboratoAggregationKeyFromValues(
      e.codice_file || e.codice_elaborato,
      e.titolo_elaborato || ""
    );
    if (key) existingKeys.add(key);

    const flags = key ? flagsByElaborato[key] : null;
    const hasNC = !!flags?.hasNC;
    const hasOSS = !!flags?.hasOSS;

    return {
      ...e,
      presenza_nc: hasNC ? "X" : "",
      presenza_oss: hasOSS ? "X" : "",
      assenza_nc_oss: !hasNC && !hasOSS ? "X" : "",
    };
  });

  for (const [key, flags] of Object.entries(flagsByElaborato)) {
    if (!key || existingKeys.has(key)) continue;

    const row = flags.sourceRow || {};
    const codiceFile = extractCodiceElaborato(row?.["Codice Elaborato"] || "");
    if (!codiceFile) continue;

    withFlags.push({
      codice_elaborato: getElaboratoBase(codiceFile),
      codice_file: codiceFile,
      revisione: row?.Revisione || getRevisioneDaCodice(codiceFile),
      titolo_elaborato: row?.["Titolo Elaborato"] || codiceFile,
      disciplina: row?.Disciplina || disciplinaFromCodice(codiceFile),
      presenza_nc: flags.hasNC ? "X" : "",
      presenza_oss: flags.hasOSS ? "X" : "",
      assenza_nc_oss: !flags.hasNC && !flags.hasOSS ? "X" : "",
    });

    existingKeys.add(key);
  }

  return withFlags;
}

function getTodoRilievoText(todo: any) {
  // REGOLA CORRETTA:
  // nella colonna "RILIEVI ITS CONTROLLI TECNICI" deve essere inserito SOLO
  // il contenuto della colonna C "Description" del file Excel ToDo.
  return findValue(todo, [
    "Description",
    "Descrizione",
    "DESCRIZIONE",
  ]);
}

function getRilievoItsText(todo: any, _bcf?: BcfTopicData | null) {
  return getTodoRilievoText(todo);
}

function getRispostaProgettistaText(
  bcf: BcfTopicData | null | undefined,
  dateRispostaProgettista: string[]
) {
  return applyFallbackDatesToUndatedBlocks(bcf?.commentiPRG || "", dateRispostaProgettista);
}

function getRiscontroIspettoreText(
  bcf: BcfTopicData | null | undefined,
  dateRiscontroIspettore: string[]
) {
  return applyFallbackDatesToUndatedBlocks(bcf?.commentiISP || "", dateRiscontroIspettore);
}

function buildElaboratiFromRowsDisciplina(rows: any[]) {
  const map: Record<string, ElaboratoVerificatoRow> = {};

  rows.forEach((row: any) => {
    const codiceFile = cleanCodiceElaborato(row?.["Codice Elaborato"] || "");
    if (!codiceFile) return;

    const titolo = row?.["Titolo Elaborato"] || codiceFile;
    const key = elaboratoAggregationKeyFromValues(codiceFile, titolo);
    if (!key || map[key]) return;

    map[key] = {
      codice_elaborato: getElaboratoBase(codiceFile),
      codice_file: codiceFile,
      revisione: row?.Revisione || getRevisioneDaCodice(codiceFile),
      titolo_elaborato: titolo,
      disciplina: row?.Disciplina || disciplinaFromCodice(codiceFile),
      presenza_nc: "",
      presenza_oss: "",
      assenza_nc_oss: "X",
    };
  });

  return Object.values(map);
}

function extractAnyTR(value: string) {
  const text = String(value || "").toUpperCase();

  const trMatch = text.match(/TR[-_\s]*0*(\d+[A-Z]?)/i);
  if (trMatch) return `TR-${trMatch[1]}`;

  const itMatch = text.match(/IT\d+[-_\s]*0*(\d+[A-Z]?)/i);
  if (itMatch) return `TR-${itMatch[1]}`;

  return "";
}

function getTodoTR(todo: any) {
  const label = findValue(todo, ["Label", "Etichetta", "Labels"]);
  const title = findValue(todo, ["Title", "Titolo", "TITLE", "Topic", "Nome"]);
  const description = findValue(todo, ["Description", "Descrizione", "DESCRIZIONE"]);

  return (
    extractAnyTR(label) ||
    normalizeTR(label) ||
    labelToTR(label) ||
    extractAnyTR(title) ||
    normalizeTR(title) ||
    extractAnyTR(description) ||
    normalizeTR(description)
  );
}

function getBcfTR(bcf: BcfTopicData) {
  return (
    extractAnyTR(bcf.titolo) ||
    normalizeTR(bcf.titolo) ||
    extractAnyTR(bcf.labels) ||
    normalizeTR(bcf.labels) ||
    extractAnyTR(bcf.descrizione) ||
    normalizeTR(bcf.descrizione)
  );
}


function getRilievoTipoBase(row: any) {
  const tipo = normalizeKey(row?.TipoBase || row?.Tipo || row?.["Codice Rilievo"] || "");
  return tipo.includes("OSS") ? "OSS" : "NC";
}

function getTrNumericValue(value: string) {
  const tr = normalizeTR(value) || extractAnyTR(value);
  const match = tr.match(/TR-(\d+)/i);
  if (!match) return Number.MAX_SAFE_INTEGER;
  return Number(match[1]);
}

function getTrAlphaSuffix(value: string) {
  const tr = normalizeTR(value) || extractAnyTR(value);
  const match = tr.match(/TR-\d+([A-Z]+)/i);
  return match ? match[1].toUpperCase() : "";
}

function getRilievoStableIdentity(row: any, index: number) {
  // REGOLA CORRETTA:
  // il TR e' l'identificativo principale del rilievo.
  // La descrizione NON deve essere usata come chiave perche' puo' essere copiata
  // o uguale tra rilievi diversi.
  const tr = normalizeTR(row?.CodiceTR || row?.["Codice Rilievo"] || row?.Label || "");
  if (tr) return tr;

  // Fallback solo per casi anomali senza TR: evita di fondere descrizioni uguali.
  return `TR-ND-${index + 1}`;
}

function buildStableProgressiviByTR(rows: any[]) {
  const identitiesByType: Record<string, string[]> = { NC: [], OSS: [] };
  const firstIndexByIdentity: Record<string, number> = {};
  const progressivoByIdentity: Record<string, string> = {};

  rows.forEach((row, index) => {
    const tipo = getRilievoTipoBase(row);
    const identity = getRilievoStableIdentity(row, index);

    if (firstIndexByIdentity[identity] === undefined) {
      firstIndexByIdentity[identity] = index;
    }

    if (!identitiesByType[tipo].includes(identity)) {
      identitiesByType[tipo].push(identity);
    }
  });

  (["NC", "OSS"] as const).forEach((tipo) => {
    identitiesByType[tipo]
      .sort((a, b) => {
        const numA = getTrNumericValue(a);
        const numB = getTrNumericValue(b);
        if (numA !== numB) return numA - numB;

        const suffixA = getTrAlphaSuffix(a);
        const suffixB = getTrAlphaSuffix(b);
        if (suffixA !== suffixB) return suffixA.localeCompare(suffixB);

        return (firstIndexByIdentity[a] ?? 0) - (firstIndexByIdentity[b] ?? 0);
      })
      .forEach((identity, index) => {
        progressivoByIdentity[identity] = `${tipo}${index + 1}`;
      });
  });

  return progressivoByIdentity;
}


function readStoricoRilieviRows(workbook: XLSX.WorkBook) {
  const sheetName =
    workbook.SheetNames.find((n) => normalizeKey(n) === "SCHEDEISPETTIVE") ||
    workbook.SheetNames[0];

  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]) as any[];
}

function buildStoricoRilieviMap(rows: any[]) {
  const map: Record<string, StoricoRilievoRow> = {};

  rows.forEach((row: any) => {
    const tr = normalizeTR(
      findValue(row, ["CodiceTR", "Codice TR", "TR", "Codice Rilievo", "Label", "Etichetta"])
    ) || extractAnyTR(
      findValue(row, ["CodiceTR", "Codice TR", "TR", "Codice Rilievo", "Label", "Etichetta"])
    );

    if (!tr || map[tr]) return;

    const tipoRaw = findValue(row, ["TipoBase", "Tipo Base", "Tipo", "Tags", "Tag", "Esito"]);
    const tipoBase = normalizeKey(tipoRaw).includes("OSS") ? "OSS" : "NC";

    const descrizioneRilievo = findValue(row, [
      "Descrizione Rilievo",
      "RILIEVI ITS CONTROLLI TECNICI",
      "Rilievi ITS Controlli Tecnici",
      "Description",
      "Descrizione",
      "DESCRIZIONE",
    ]);

    map[tr] = {
      tr,
      tipoBase,
      descrizioneRilievo,
    };
  });

  return map;
}


function extractAllegatoNumber(value: string) {
  const text = String(value || "");
  const match = text.match(/allegato[\s_-]*(\d+)/i);
  return match ? match[1] : "";
}

function getTodoAllegatoNumber(todo: any, bcf?: BcfTopicData | null) {
  const values = [
    findValue(todo, ["Title", "Titolo", "TITLE", "Topic", "Nome"]),
    findValue(todo, ["Description", "Descrizione", "DESCRIZIONE"]),
    bcf?.titolo || "",
    bcf?.descrizione || "",
    bcf?.descrizioneConData || "",
  ];

  for (const value of values) {
    const allegato = extractAllegatoNumber(value);
    if (allegato) return allegato;
  }

  return "";
}

function appendAllegatoToCodice(codice: string, allegato: string) {
  const clean = cleanCodiceElaborato(codice);
  if (!clean || !allegato) return clean;
  if (extractAllegatoNumber(clean)) return clean;
  return `${clean} - Allegato_${allegato}`;
}

function buildTitoloAllegato(titoloBase: string, allegato: string) {
  const base = String(titoloBase || "Relazione di calcolo impalcato").trim();

  if (!allegato) return base;
  if (extractAllegatoNumber(base)) return base;

  return `${base}- Allegato_${allegato}.pdf`;
}

function elaboratoAggregationKeyFromValues(codice: string, titolo: string) {
  const baseKey = elaboratoAggregationKey(codice);
  const allegato = extractAllegatoNumber(`${codice} ${titolo}`);

  if (baseKey && allegato) return `${baseKey}__ALLEGATO_${allegato}`;
  return baseKey;
}


function mergeElaboratiFromRows(
  elaborati: ElaboratoVerificatoRow[],
  rows: any[]
) {
  const map: Record<string, ElaboratoVerificatoRow> = {};

  function addRow(e: ElaboratoVerificatoRow) {
    const key = elaboratoAggregationKeyFromValues(
      e.codice_file || e.codice_elaborato,
      e.titolo_elaborato || ""
    );

    if (!key) return;

    const existing = map[key];

    if (!existing) {
      map[key] = e;
      return;
    }

    map[key] = {
      ...existing,
      codice_elaborato: existing.codice_elaborato || e.codice_elaborato,
      codice_file: existing.codice_file || e.codice_file,
      revisione: existing.revisione || e.revisione,
      titolo_elaborato: existing.titolo_elaborato || e.titolo_elaborato,
      disciplina: existing.disciplina || e.disciplina,
      presenza_nc: existing.presenza_nc || e.presenza_nc,
      presenza_oss: existing.presenza_oss || e.presenza_oss,
      assenza_nc_oss:
        existing.presenza_nc || existing.presenza_oss || e.presenza_nc || e.presenza_oss
          ? ""
          : existing.assenza_nc_oss || e.assenza_nc_oss || "X",
    };
  }

  elaborati.forEach(addRow);

  buildElaboratiFromRowsDisciplina(rows).forEach(addRow);

  return Object.values(map);
}


function bestBcfMatchForTodo(bcfTopics: BcfTopicData[], todo: any) {
  const titoloTodo = findValue(todo, ["Title", "Titolo", "TITLE", "Topic", "Nome"]);
  const descrizioneTodo = findValue(todo, ["Description", "Descrizione", "DESCRIZIONE"]);
  const labelTodo = findValue(todo, ["Label", "Etichetta", "Labels"]);

  const codiceTodo = extractCodiceElaborato(titoloTodo || descrizioneTodo);
  const trTodo = getTodoTR(todo);

  const ranked = bcfTopics
    .map((bcf) => {
      const codiceBcf = extractCodiceElaborato(bcf.titolo || bcf.descrizione);
      const trBcf = getBcfTR(bcf);

      let score = 0;

      // Match principale: TR-xxx.
      // Serve per recuperare risposta progettista e riscontro ispettore
      // anche quando titolo/descrizione ToDo e BCF non coincidono perfettamente.
      if (trTodo && trBcf && trTodo === trBcf) score += 100;

      if (codiceTodo && sameElaboratoCode(codiceBcf, codiceTodo)) score += 10;
      if (normalizeKey(bcf.titolo) === normalizeKey(titoloTodo)) score += 5;

      if (labelTodo && normalizeKey(bcf.labels).includes(normalizeKey(labelTodo))) {
        score += 5;
      }

      score += descriptionScore(bcf.descrizione, descrizioneTodo);

      return { row: bcf, score };
    })
    .sort((a, b) => b.score - a.score);

  return ranked[0] && ranked[0].score > 0.05 ? ranked[0].row : null;
}

function readFirstSheet(workbook: XLSX.WorkBook) {
  return XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
}

function readReportRows(workbook: XLSX.WorkBook) {
  const sheetName =
    workbook.SheetNames.find((n) => normalizeKey(n) === "VERIFICAELABORATI") ||
    workbook.SheetNames[0];

  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]) as any[];
}

function escapeXml(value: string | number) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function getPlainTextFromWordXml(xml: string) {
  return String(xml || "")
    .replace(/<w:tab\/>/g, "\t")
    .replace(/<w:br\/>/g, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function rowContainsStatoChiusa(rowXml: string) {
  const plain = getPlainTextFromWordXml(rowXml);
  return (
    /(^|\s)Chiusa(\s|$)/i.test(plain) ||
    /(^|\s)Chiuso(\s|$)/i.test(plain) ||
    /(^|\s)Closed(\s|$)/i.test(plain)
  );
}

function applyGreyTextToWordRow(rowXml: string) {
  return rowXml.replace(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g, (runXml) => {
    const runWithoutColor = runXml.replace(
      /<w:color\b[^>]*(?:\/>|>[\s\S]*?<\/w:color>)/g,
      ""
    );

    if (/<w:rPr\b[^>]*>/.test(runWithoutColor)) {
      return runWithoutColor.replace(
        /<w:rPr\b([^>]*)>/,
        '<w:rPr$1><w:color w:val="808080"/>'
      );
    }

    return runWithoutColor.replace(
      /<w:r\b([^>]*)>/,
      '<w:r$1><w:rPr><w:color w:val="808080"/></w:rPr>'
    );
  });
}

function applyClosedRowsGreyText(documentXml: string) {
  return documentXml.replace(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g, (rowXml) => {
    if (!rowContainsStatoChiusa(rowXml)) return rowXml;
    return applyGreyTextToWordRow(rowXml);
  });
}

function buildSintesiFinaleDocxXml(sintesi: SchedaIspettivaSintesi) {
  const rows: Array<[string, number]> = [
    ["Totale elaborati verificati", sintesi.totaleElaboratiAnalizzati],
    ["Totale NC rilevate", sintesi.totaleNC],
    ["Totale OSS rilevate", sintesi.totaleOSS],
    ["Totale rilievi chiusi", sintesi.totaleChiuse],
  ];

  const tableRows = rows
    .map(
      ([label, value]) => `
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="7000" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t>${escapeXml(label)}</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>
          <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>${escapeXml(value)}</w:t></w:r></w:p>
        </w:tc>
      </w:tr>`
    )
    .join("");

  return `
  <w:p>
    <w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>
    <w:r><w:rPr><w:b/></w:rPr><w:t>SINTESI FINALE</w:t></w:r>
  </w:p>
  <w:tbl>
    <w:tblPr>
      <w:tblW w:w="9000" w:type="dxa"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      </w:tblBorders>
    </w:tblPr>
    ${tableRows}
  </w:tbl>
  <w:p/>`;
}

function appendSintesiAfterLastTable(documentXml: string, sintesi: SchedaIspettivaSintesi) {
  const marker = "</w:tbl>";
  const lastTableIndex = documentXml.lastIndexOf(marker);
  const sintesiXml = buildSintesiFinaleDocxXml(sintesi);

  if (lastTableIndex < 0) {
    const bodyCloseIndex = documentXml.lastIndexOf("</w:body>");
    if (bodyCloseIndex < 0) return documentXml + sintesiXml;
    return documentXml.slice(0, bodyCloseIndex) + sintesiXml + documentXml.slice(bodyCloseIndex);
  }

  const insertIndex = lastTableIndex + marker.length;
  return documentXml.slice(0, insertIndex) + sintesiXml + documentXml.slice(insertIndex);
}


function ensureDocumentNamespaces(documentXml: string) {
  let xml = documentXml;

  const namespaces: Array<[string, string]> = [
    ["xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
    ["xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"],
    ["xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main"],
    ["xmlns:pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"],
  ];

  namespaces.forEach(([name, value]) => {
    if (xml.includes(`${name}=`)) return;
    xml = xml.replace("<w:document", `<w:document ${name}="${value}"`);
  });

  return xml;
}



function buildWordTextRuns(value: string) {
  const lines = String(value || "").split(/\n/g);

  return lines
    .map((line, index) => {
      const br = index === 0 ? "" : "<w:br/>";
      return `${br}<w:t xml:space="preserve">${escapeXml(line)}</w:t>`;
    })
    .join("");
}

function buildCellWithText(cellXml: string, value: string) {
  const tcPr = cellXml.match(/<w:tcPr\b[^>]*>[\s\S]*?<\/w:tcPr>/)?.[0] || "";
  const pPr = cellXml.match(/<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/)?.[0] || "";
  const rPr = cellXml.match(/<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/)?.[0] || "";

  return `<w:tc>${tcPr}<w:p>${pPr}<w:r>${rPr}${buildWordTextRuns(value)}</w:r></w:p></w:tc>`;
}

function buildRevisionRowFromTemplate(rowXml: string, revisione: RevisioneSchedaRow) {
  const cells = rowXml.match(/<w:tc\b[^>]*>[\s\S]*?<\/w:tc>/g);
  if (!cells || cells.length < 4) return rowXml;

  const open = rowXml.match(/^<w:tr\b[^>]*>/)?.[0] || "<w:tr>";
  const trPr = rowXml.match(/<w:trPr\b[^>]*>[\s\S]*?<\/w:trPr>/)?.[0] || "";
  const close = "</w:tr>";

  // Supporta sia template con 5 colonne (ASPI: PCQ + ITS)
  // sia template con 4 colonne (ORERO: solo ITS).
  const values =
    cells.length >= 5
      ? [
          revisione.rev,
          revisione.data,
          revisione.descrizione,
          revisione.responsabile_pcq,
          revisione.responsabile_its,
        ]
      : [
          revisione.rev,
          revisione.data,
          revisione.descrizione,
          revisione.responsabile_its,
        ];

  const newCells = cells.map((cell, index) => {
    if (index >= values.length) return cell;
    return buildCellWithText(cell, values[index]);
  });

  return `${open}${trPr}${newCells.join("")}${close}`;
}

function patchRevisioniSchedaTableRows(
  documentXml: string,
  revisioniScheda: RevisioneSchedaRow[]
) {
  if (!revisioniScheda.length) return documentXml;

  return documentXml.replace(/<w:tbl\b[^>]*>[\s\S]*?<\/w:tbl>/g, (tableXml) => {
    const rows = tableXml.match(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g);
    if (!rows || rows.length < 2) return tableXml;

    const headerIndex = rows.findIndex((row) => {
      const plain = normalizeText(getPlainTextFromWordXml(row));
      return (
        plain.includes("REV") &&
        plain.includes("DATA") &&
        plain.includes("DESCRIZIONE") &&
        plain.includes("APPROVATO")
      );
    });

    if (headerIndex <= 0) return tableXml;

    const dataRowIndex = headerIndex - 1;
    const templateRow = rows[dataRowIndex];
    const revisionRowsXml = revisioniScheda
      .map((rev) => buildRevisionRowFromTemplate(templateRow, rev))
      .join("");

    const newRows = [
      ...rows.slice(0, dataRowIndex),
      revisionRowsXml,
      ...rows.slice(headerIndex),
    ];

    let rowCounter = 0;
    return tableXml.replace(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/g, () => {
      const replacement = newRows[rowCounter] || "";
      rowCounter += 1;
      return replacement;
    });
  });
}

async function postProcessSchedaIspettivaDocx(
  buffer: Buffer,
  sintesi: SchedaIspettivaSintesi,
  revisioniScheda: RevisioneSchedaRow[]
) {
  const zip = await JSZip.loadAsync(buffer);
  const documentFile = zip.file("word/document.xml");
  if (!documentFile) return buffer;

  let documentXml = await documentFile.async("string");

  documentXml = applyClosedRowsGreyText(documentXml);
  documentXml = patchRevisioniSchedaTableRows(documentXml, revisioniScheda);
  documentXml = appendSintesiAfterLastTable(documentXml, sintesi);

  zip.file("word/document.xml", documentXml);

  return Buffer.from(
    await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" })
  );
}

export async function POST(req: Request) {
  try {
    const formData = await req.formData();

    const todoFile = formData.get("todo") as File;
    const bcfFiles = formData.getAll("bcf") as File[];
    const elencoFile = formData.get("elenco") as File;
    const templateFile = formData.get("template") as File;
    const reportFile = (formData.get("report") || formData.get("files")) as File | null;
    const storicoFile = (
      formData.get("storico") ||
      formData.get("storico_rilievi") ||
      formData.get("schede_precedenti")
    ) as File | null;

    const progettisti = parsePeopleList(formData.get("progettisti"));
    const ispettori = parsePeopleList(formData.get("ispettori"));

    const revisioneScheda = String(
      formData.get("revisione_scheda") || REVISIONE_SCHEDA
    ).trim();
    const dataRevisioneScheda = String(
      formData.get("data_revisione_scheda") || "xx/xx/xxxx"
    ).trim();
    const dataRispostaProgettista = String(
      formData.get("data_risposta_progettista") || ""
    ).trim();
    const dataRiscontroIspettore = String(
      formData.get("data_riscontro_ispettore") || ""
    ).trim();
    const dateRispostaProgettista = readIndexedFormDates(
      formData,
      "data_risposta_progettista",
      dataRispostaProgettista
    );
    const dateRiscontroIspettore = readIndexedFormDates(
      formData,
      "data_riscontro_ispettore",
      dataRiscontroIspettore
    );
    const responsabilePcq = String(formData.get("responsabile_pcq") || "").trim();
    const responsabileIts = String(formData.get("responsabile_its") || "").trim();

    const revisioniScheda = buildRevisioniSchedaRows(
      formData,
      revisioneScheda,
      dataRevisioneScheda,
      dataRiscontroIspettore,
      responsabilePcq,
      responsabileIts
    );

    if (!todoFile || bcfFiles.length === 0 || !elencoFile || !templateFile) {
      return NextResponse.json({
        ok: false,
        error:
          "Carica ToDo XLSX, almeno un BCFZIP, Elenco Elaborati XLSX e Template DOCX.",
      });
    }

    const todoWorkbook = XLSX.read(Buffer.from(await todoFile.arrayBuffer()), {
      type: "buffer",
    });
    const todoRowsRaw: any[] = readFirstSheet(todoWorkbook) as any[];

    const elencoWorkbook = XLSX.read(Buffer.from(await elencoFile.arrayBuffer()), {
      type: "buffer",
    });
    const elencoRows: any[] = readFirstSheet(elencoWorkbook) as any[];

    let reportRows: any[] = [];
    if (reportFile) {
      const reportWorkbook = XLSX.read(Buffer.from(await reportFile.arrayBuffer()), {
        type: "buffer",
      });
      reportRows = readReportRows(reportWorkbook) as any[];
    }

    let storicoRows: any[] = [];
    if (storicoFile) {
      const storicoWorkbook = XLSX.read(Buffer.from(await storicoFile.arrayBuffer()), {
        type: "buffer",
      });
      storicoRows = readStoricoRilieviRows(storicoWorkbook) as any[];
    }

    const storicoRilieviMap = buildStoricoRilieviMap(storicoRows);

    const elencoInfoMap: Record<string, any> = {};
    const disciplinaInfoMap: Record<string, any> = {};
    const reportInfoMap: Record<string, any> = {};

    elencoRows.forEach((r: any) => {
      const codice = findValue(r, [
        "Codice_SP",
        "Codice SP",
        "Codice Elaborato",
        "CODICE ELABORATO",
        "Codice elaborato",
      ]);

      const titolo = getTitoloProgetto(r);
      const revisione = findValue(r, ["REV.", "REV", "Rev.", "Revisione"]);

      const nomeRedattore = findValue(r, [
        "Nome_redattore",
        "Nome redattore",
        "NOME REDATTORE",
        "Nome Redattore",
      ]);

      const notaRicezione = findValue(r, [
        "Nota_ricezione_elaborati",
        "Nota ricezione elaborati",
        "Nota Ricezione Elaborati",
      ]);

      const dataRicezione = findValue(r, [
        "Data_ricezione",
        "Data ricezione",
        "Data Ricezione",
      ]);

      const faseProgetto = getFaseProgetto(r);

      const disciplinaElenco =
        findValue(r, ["DISCIPLINA", "Disciplina", "Oggetto", "OGGETTO"]) ||
        disciplinaFromCodice(codice);

      const ispettoreElenco = findValue(r, ISPETTORE_ELENCO_COLUMNS);

      if (disciplinaElenco) {
        disciplinaInfoMap[normalizeKey(disciplinaElenco)] = {
          codiceScheda: codice,
          titoloProgetto: titolo,
          faseProgetto,
          notaRicezione,
          dataRicezione,
          nomeRedattore,
          ispettoreElenco,
        };
      }

      const keys = [
        normalizeKey(cleanCodiceElaborato(codice)),
        normalizeKey(getElaboratoBase(codice)),
      ].filter(Boolean);

      keys.forEach((key) => {
        elencoInfoMap[key] = {
          codice,
          titolo,
          revisione,
          nomeRedattore,
          notaRicezione,
          dataRicezione,
          disciplina: disciplinaElenco,
          ispettoreElenco,
        };
      });
    });

    reportRows.forEach((r: any) => {
      const codice =
        findValue(r, ["Codice elenco", "Codice Elenco", "CODICE ELENCO"]) ||
        findValue(r, ["Codice cartiglio", "Codice Cartiglio", "CODICE CARTIGLIO"]) ||
        findValue(r, ["Nome file PDF", "Nome File PDF", "NOME FILE PDF"]) ||
        findValue(r, REPORT_CODICE_COLUMNS);

      if (!codice) return;

      const nomeFilePdf = findValue(r, ["Nome file PDF", "Nome File PDF", "NOME FILE PDF"]);
      const codicePulito = extractCodiceElaborato(codice);
      const base = getElaboratoBase(codicePulito);

      const info = {
        codice: codicePulito,
        codiceBase: base,
        revisione:
          findValue(r, ["REV", "REV.", "Rev.", "Revisione", "REVISIONE"]) ||
          getRevisioneDaCodice(nomeFilePdf || codicePulito),
        titolo: findValue(r, TITOLO_ELABORATO_COLUMNS),
        disciplina:
          disciplinaFromReportCartella(findValue(r, ["Cartella", "CARTELLA"])) ||
          findValue(r, ["DISCIPLINA", "Disciplina", "Oggetto", "OGGETTO"]) ||
          disciplinaFromCodice(codicePulito),
      };

      [
        codicePulito,
        base,
        findValue(r, ["Codice cartiglio", "Codice Cartiglio", "CODICE CARTIGLIO"]),
        findValue(r, ["Nome file PDF", "Nome File PDF", "NOME FILE PDF"]),
      ]
        .map((v) => normalizeKey(getElaboratoBase(extractCodiceElaborato(v))))
        .filter(Boolean)
        .forEach((key) => {
          reportInfoMap[key] = info;
        });
    });

    const parser = new XMLParser({ ignoreAttributes: false });
    const commentiMap: Record<string, BcfTopicData> = {};
    const bcfTopics: BcfTopicData[] = [];

    for (const bcfFile of bcfFiles) {
      const bcfZip = await JSZip.loadAsync(
        Buffer.from(await bcfFile.arrayBuffer())
      );

      for (const fileName of Object.keys(bcfZip.files)) {
        if (!fileName.endsWith("markup.bcf")) continue;

        const xml = await bcfZip.files[fileName].async("text");
        const parsed: any = parser.parse(xml);

        const topic = parsed?.Markup?.Topic || {};
        const commentsRaw = parsed?.Markup?.Comment;
        const comments = commentsRaw
          ? Array.isArray(commentsRaw)
            ? commentsRaw
            : [commentsRaw]
          : [];

        const topicTitle = topic?.Title || "";
        const topicDescription = topic?.Description || "";
        const topicGuid = topic?.["@_Guid"] || topic?.Guid || "";
        const topicLabels = String(topic?.Labels || "");
        const topicStatus = String(
          topic?.["@_TopicStatus"] ||
            topic?.TopicStatus ||
            topic?.Status ||
            ""
        );

        const commentiPRGEntries: CommentEntry[] = [];
        const commentiISPEntries: CommentEntry[] = [];
        let lastIspAuthor = "";

        comments.forEach((c: any, index: number) => {
          const testo = getCommentText(c);
          const cleanText = cleanRolePrefix(testo);
          const author = getCommentAuthor(c);
          const date = formatBcfCommentDate(getCommentDateValue(c));

          if (!cleanText) return;

          const entry: CommentEntry = {
            author,
            date,
            text: cleanText,
            order: index,
          };

          const isPRGByAccount = isInPeopleList(author, progettisti);
          const isISPByAccount = isInPeopleList(author, ispettori);

          if (isPRGByAccount) {
            commentiPRGEntries.push(entry);
            return;
          }

          if (isISPByAccount) {
            lastIspAuthor = author || lastIspAuthor;
            commentiISPEntries.push(entry);
            return;
          }

          if (/\(\s*PRG\s*\)/i.test(testo)) {
            commentiPRGEntries.push(entry);
            return;
          }

          if (/\(\s*ISP\s*\)/i.test(testo)) {
            lastIspAuthor = author || lastIspAuthor;
            commentiISPEntries.push(entry);
            return;
          }
        });

        const commentiPRGText = buildCommentBlocks(commentiPRGEntries);
        const commentiISPText = buildCommentBlocks(commentiISPEntries);

        const topicAuthor =
          topic?.CreationAuthor ||
          topic?.Author ||
          topic?.["@_CreationAuthor"] ||
          topic?.["@_Author"] ||
          "";

        const ispettoreNomeBcf = String(lastIspAuthor || topicAuthor || "").trim();

        const existing =
          (topicGuid && commentiMap[normalizeKey(topicGuid)]) ||
          commentiMap[topicKey(topicTitle, topicDescription)];

        const lastComment = comments.length > 0 ? comments[comments.length - 1] : null;

        const dataTopic: BcfTopicData = {
          topicGuid: topicGuid || existing?.topicGuid || "",
          titolo: topicTitle || existing?.titolo || "",
          descrizione: topicDescription || existing?.descrizione || "",
          descrizioneConData:
            topicDescription
              ? prefixCommentWithDate(topicDescription, getTopicDateValue(topic))
              : existing?.descrizioneConData || existing?.descrizione || "",
          ispettore:
            existing?.ispettore ||
            siglaDaNome(ispettoreNomeBcf || topicAuthor),
          ispettoreNomeBcf:
            ispettoreNomeBcf || existing?.ispettoreNomeBcf || "",
          labels: topicLabels || existing?.labels || "",
          stato: topicStatus || existing?.stato || "",
          commentiPRG: mergeCommentBlocks(existing?.commentiPRG || "", commentiPRGText),
          commentiISP: mergeCommentBlocks(existing?.commentiISP || "", commentiISPText),
          ultimoCommento: lastComment
            ? prefixCommentWithDate(
                cleanRolePrefix(getCommentText(lastComment)),
                getCommentDateValue(lastComment)
              )
            : existing?.ultimoCommento || "",
        };

        const previousIndex = bcfTopics.findIndex(
          (t) =>
            (dataTopic.topicGuid &&
              normalizeKey(t.topicGuid) === normalizeKey(dataTopic.topicGuid)) ||
            topicKey(t.titolo, t.descrizione) ===
              topicKey(dataTopic.titolo, dataTopic.descrizione)
        );

        if (previousIndex >= 0) {
          bcfTopics[previousIndex] = dataTopic;
        } else {
          bcfTopics.push(dataTopic);
        }

        const key = topicKey(dataTopic.titolo, dataTopic.descrizione);
        commentiMap[key] = dataTopic;

        if (dataTopic.topicGuid) {
          commentiMap[normalizeKey(dataTopic.topicGuid)] = dataTopic;
        }
      }
    }

    const todoRows = todoRowsRaw.filter((r: any) => {
      const tags = findValue(r, ["Tags", "Tag", "Tipo", "Esito"]);
      const descrizione = findValue(r, ["Description", "Descrizione"]);
      const status = findValue(r, ["Status", "Stato"]);
      return !(isNessunRilievo(tags, descrizione) && isClosedStatus(status));
    });

    const finalRows = todoRows
      .filter((todo) => {
        const tags = findValue(todo, ["Tags", "Tag", "Tipo", "Esito"]);
        const descrizione = findValue(todo, ["Description", "Descrizione"]);
        return !isNessunRilievo(tags, descrizione);
      })
      .map((todo) => {
        const bcf = bestBcfMatchForTodo(bcfTopics, todo);

        const label = findValue(todo, ["Label", "Etichetta"]);
        const tags = findValue(todo, ["Tags", "Tag", "Tipo", "Esito"]);
        const titoloTodo = findValue(todo, ["Title", "Titolo", "TITLE", "Topic", "Nome"]);
        const descrizioneTodo = findValue(todo, ["Description", "Descrizione"]);

        const disciplina =
          findValue(todo, ["Assignee(s) ", "Assignee(s)", "Disciplina"]) ||
          disciplinaFromCodice(titoloTodo || descrizioneTodo);

        const tipoBase = String(tags || "").toUpperCase().includes("OSS")
          ? "OSS"
          : "NC";

        const codiceTR =
          getTodoTR(todo) ||
          (label ? labelToTR(label) : extractTR(titoloTodo) || extractTR(descrizioneTodo) || "");

        const storicoRilievo = storicoRilieviMap[normalizeTR(codiceTR)] || null;

        const codiceElaboratoBase = extractCodiceElaborato(titoloTodo || descrizioneTodo);
        const allegatoNumero = getTodoAllegatoNumber(todo, bcf);
        const codiceElaborato = appendAllegatoToCodice(codiceElaboratoBase, allegatoNumero);
        const reportInfo = allegatoNumero ? {} : findReportInfo(reportInfoMap, codiceElaboratoBase);
        const disciplinaInfo = disciplinaInfoMap[normalizeKey(disciplina)] || {};
        const infoElenco = getElencoInfoByCode(elencoInfoMap, codiceElaborato);

        const ispettoreTodo = getIspettoreFromTodo(todo);
        const ispettoreElenco =
          infoElenco.ispettoreElenco || disciplinaInfo.ispettoreElenco || "";

        const ispettoreGrezzo = ispettoreTodo
          ? remapIspettoreFinale(siglaDaNome(ispettoreTodo), disciplina)
          : resolveIspettoreFinale(
              bcf?.ispettore || "",
              bcf?.ispettoreNomeBcf || "",
              disciplina,
              ispettoreElenco
            );

        const ispettoreFinale = remapIspettoreUfficiale(
          ispettoreGrezzo,
          disciplina,
          disciplinaInfo.nomeRedattore || infoElenco.nomeRedattore || ispettoreElenco
        );

        const titoloElaboratoBase =
          getTitoloElaboratoFromTodo(todo) ||
          infoElenco.titolo ||
          reportInfo.titolo ||
          "";

        const titoloElaborato = allegatoNumero
          ? buildTitoloAllegato(titoloElaboratoBase || reportInfo.titolo, allegatoNumero)
          : reportInfo.titolo || titoloElaboratoBase || "";

        const rispostaProgettista = getRispostaProgettistaText(
          bcf,
          dateRispostaProgettista
        );
        const riscontroIspettore = getRiscontroIspettoreText(
          bcf,
          dateRiscontroIspettore
        );
        const statoFinale = determineRilievoStatus(
          findValue(todo, ["Status", "Stato"]) || bcf?.stato || "",
          tags,
          riscontroIspettore,
          rispostaProgettista
        );

        return {
          Disciplina: disciplina,
          Label: label,
          TipoBase: storicoRilievo?.tipoBase || tipoBase,
          CodiceTR: normalizeTR(codiceTR) || codiceTR,
          "Codice Rilievo": label || codiceTR,
          "Codice Elaborato": codiceElaborato || titoloTodo || "",
          "Titolo Elaborato": titoloElaborato || codiceElaborato || titoloTodo || "",
          Revisione:
            reportInfo.revisione ||
            getRevisioneDaCodice(codiceElaboratoBase || titoloTodo),
          Tipo: tags || storicoRilievo?.tipoBase || tipoBase,
          "Descrizione Rilievo":
            storicoRilievo?.descrizioneRilievo || getRilievoItsText(todo, bcf),
          Ispettore: ispettoreFinale,
          "Risposta Progettista PRG": rispostaProgettista,
          "Riscontro Ispettore ISP": riscontroIspettore,
          "Ultimo Commento": bcf?.ultimoCommento || "",
          "Azione Richiesta": "",
          Stato: statoFinale,
          "Nota Ricezione Elaborati": disciplinaInfo.notaRicezione || "",
          "Data Ricezione": disciplinaInfo.dataRicezione || "",
          "Nome Redattore": disciplinaInfo.nomeRedattore || "",
        };
      });

    const outputZip = new JSZip();

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(finalRows);
    XLSX.utils.book_append_sheet(wb, ws, "Schede_Ispettive");
    outputZip.file(
      "SCHEDE_ISPETTIVE_CONSOLIDATE.xlsx",
      XLSX.write(wb, { type: "buffer", bookType: "xlsx" })
    );

    const templateBuffer = Buffer.from(await templateFile.arrayBuffer());
    const elaboratiSource = reportRows.length ? reportRows : elencoRows;

    const elaboratiVerificatiAll = elaboratiSource
      .map((e: any) => {
        const codiceCompleto =
          findValue(e, ["Codice elenco", "Codice Elenco", "CODICE ELENCO"]) ||
          findValue(e, ["Codice cartiglio", "Codice Cartiglio", "CODICE CARTIGLIO"]) ||
          findValue(e, ["Nome file PDF", "Nome File PDF", "NOME FILE PDF"]) ||
          findValue(e, REPORT_CODICE_COLUMNS);

        const codicePulito = extractCodiceElaborato(codiceCompleto);
        if (!codicePulito) return null;

        const infoElenco = getElencoInfoByCode(elencoInfoMap, codicePulito);
        const reportInfo = findReportInfo(reportInfoMap, codicePulito);

        const titoloRaw =
          findValue(e, TITOLO_ELABORATO_COLUMNS) ||
          reportInfo.titolo ||
          infoElenco.titolo ||
          "";

        const allegatoNumero = extractAllegatoNumber(`${codiceCompleto} ${titoloRaw}`);
        const codiceConAllegato = appendAllegatoToCodice(codicePulito, allegatoNumero);

        const titoloElenco = allegatoNumero
          ? buildTitoloAllegato(titoloRaw || reportInfo.titolo, allegatoNumero)
          : reportInfo.titolo || titoloRaw || "";

        const revisione =
          reportInfo.revisione ||
          findValue(e, ["REV", "REV.", "Rev.", "Revisione", "REVISIONE"]) ||
          infoElenco.revisione ||
          getRevisioneDaCodice(codicePulito);

        const codiceSenzaRev = getElaboratoBase(codiceConAllegato);

        const disciplinaElaborato =
          reportInfo.disciplina ||
          disciplinaFromReportCartella(findValue(e, ["Cartella", "CARTELLA"])) ||
          infoElenco.disciplina ||
          findValue(e, ["Disciplina", "DISCIPLINA", "Oggetto", "OGGETTO"]) ||
          disciplinaFromCodice(codicePulito);

        return {
          codice_elaborato: codiceSenzaRev,
          codice_file: codiceConAllegato,
          revisione,
          titolo_elaborato: titoloElenco,
          disciplina: disciplinaElaborato,
          presenza_nc: "",
          presenza_oss: "",
          assenza_nc_oss: "X",
        };
      })
      .filter(Boolean) as ElaboratoVerificatoRow[];

    const discipline = Array.from(
      new Set(
        elencoRows
          .map((r: any) => findValue(r, ["DISCIPLINA", "Disciplina"]))
          .filter(Boolean)
      )
    );

    for (const disciplina of discipline) {
      if (!disciplina || normalizeKey(disciplina) === "SENZADISCIPLINA") {
        continue;
      }

      const rowsDisciplina = finalRows.filter((r) =>
        sameDisciplina(r.Disciplina || "", disciplina)
      );

      let elaboratiVerificati = elaboratiVerificatiAll.filter((e) =>
        sameDisciplina(e.disciplina || "", disciplina)
      );

      elaboratiVerificati = applyRilieviFlagsToElaborati(
        mergeElaboratiFromRows(elaboratiVerificati, rowsDisciplina),
        rowsDisciplina
      );

      const elaboratiDedup: Record<string, ElaboratoVerificatoRow> = {};
      elaboratiVerificati.forEach((e) => {
        const key = elaboratoAggregationKeyFromValues(e.codice_file || e.codice_elaborato, e.titolo_elaborato || "");
        if (!key) return;

        const existing = elaboratiDedup[key];
        if (!existing) {
          elaboratiDedup[key] = e;
          return;
        }

        elaboratiDedup[key] = {
          ...existing,
          presenza_nc: existing.presenza_nc || e.presenza_nc,
          presenza_oss: existing.presenza_oss || e.presenza_oss,
          assenza_nc_oss:
            existing.presenza_nc || existing.presenza_oss || e.presenza_nc || e.presenza_oss
              ? ""
              : existing.assenza_nc_oss || e.assenza_nc_oss,
        };
      });
      elaboratiVerificati = Object.values(elaboratiDedup);

      const elencoDisciplina = elencoRows.filter((e: any) =>
        sameDisciplina(
          findValue(e, ["DISCIPLINA", "Disciplina"]) ||
            disciplinaFromCodice(findValue(e, ["Codice_SP"])),
          disciplina
        )
      );

      if (rowsDisciplina.length === 0 && elaboratiVerificati.length === 0) {
        continue;
      }

      const primaRigaElenco = elencoDisciplina[0] || elencoRows[0] || {};
      const disciplinaInfo = disciplinaInfoMap[normalizeKey(disciplina)] || {};

      const codiceScheda =
        disciplinaInfo.codiceScheda ||
        findValue(primaRigaElenco, ["Codice_SP"]) ||
        findValue(primaRigaElenco, ["Codice SP"]) ||
        "SCHEDA_ISPETTIVA";

      const titoloProgetto =
        disciplinaInfo.titoloProgetto || getTitoloProgetto(primaRigaElenco) || "";

      const faseProgetto =
        disciplinaInfo.faseProgetto || getFaseProgetto(primaRigaElenco);

      const notaRicezione =
        disciplinaInfo.notaRicezione ||
        findValue(primaRigaElenco, ["Nota_ricezione_elaborati"]) ||
        "";

      const dataRicezione =
        disciplinaInfo.dataRicezione ||
        findValue(primaRigaElenco, ["Data_ricezione"]) ||
        "";

      const nomeRedattore =
        disciplinaInfo.nomeRedattore ||
        findValue(primaRigaElenco, ["Nome_redattore"]) ||
        "";

      if (elaboratiVerificati.length === 0) {
        // Primo fallback: righe dell'elenco elaborati della disciplina.
        elaboratiVerificati = elencoDisciplina
          .map((e: any) => {
            const codiceCompleto =
              findValue(e, ["Codice_SP", "Codice SP", "Codice Elaborato", "CODICE ELABORATO"]) ||
              findValue(e, REPORT_CODICE_COLUMNS);

            const codicePulito = extractCodiceElaborato(codiceCompleto);
            if (!codicePulito) return null;

            const titoloRaw =
              findValue(e, TITOLO_ELABORATO_COLUMNS) ||
              getTitoloProgetto(e) ||
              codicePulito;
            const allegatoNumero = extractAllegatoNumber(`${codiceCompleto} ${titoloRaw}`);
            const codiceConAllegato = appendAllegatoToCodice(codicePulito, allegatoNumero);
            const codiceSenzaRev = getElaboratoBase(codiceConAllegato);

            return {
              codice_elaborato: codiceSenzaRev,
              codice_file: codiceConAllegato,
              revisione: findValue(e, ["REV.", "REV", "Rev.", "Revisione"]),
              titolo_elaborato: allegatoNumero
                ? buildTitoloAllegato(titoloRaw, allegatoNumero)
                : titoloRaw,
              disciplina: findValue(e, ["DISCIPLINA", "Disciplina"]) || disciplina,
              presenza_nc: "",
              presenza_oss: "",
              assenza_nc_oss: "X",
            };
          })
          .filter(Boolean) as ElaboratoVerificatoRow[];

        // Secondo fallback: se l'elenco non contiene gli elaborati della disciplina,
        // usa almeno tutti gli elaborati citati nei rilievi della disciplina.
        if (elaboratiVerificati.length === 0) {
          elaboratiVerificati = buildElaboratiFromRowsDisciplina(rowsDisciplina);
        }

        elaboratiVerificati = applyRilieviFlagsToElaborati(
          mergeElaboratiFromRows(elaboratiVerificati, rowsDisciplina),
          rowsDisciplina
        );
      }

      const progressivoByTR = buildStableProgressiviByTR(rowsDisciplina);

      const rilievi = rowsDisciplina.map((r, index) => {
        const identity = getRilievoStableIdentity(r, index);
        const tr = normalizeTR(r.CodiceTR || r["Codice Rilievo"] || r.Label || "") || r.CodiceTR || "TR-ND";
        const progressivo = progressivoByTR[identity] || `${getRilievoTipoBase(r)}${index + 1}`;

        return {
          tipo_progressivo: `${progressivo}\n(${tr})`,
          codice_elaborato: r["Codice Elaborato"] || "",
          titolo_elaborato: r["Titolo Elaborato"] || "",
          rilievo_its: r["Descrizione Rilievo"] || "",
          ispettore: r.Ispettore || "",
          risposta_prg: r["Risposta Progettista PRG"] || "",
          riscontro_isp: r["Riscontro Ispettore ISP"] || "",
          stato: r.Stato || "",
        };
      });

      const numeroNC = rilievi.filter(
        (r) => !isRilievoOSS({ ...r, Stato: r.stato, TipoBase: String(r.tipo_progressivo).startsWith("OSS") ? "OSS" : "NC" }) && isOpenStatus(r.stato)
      ).length;

      const numeroOSS = rilievi.filter(
        (r) => isRilievoOSS({ ...r, Stato: r.stato, TipoBase: String(r.tipo_progressivo).startsWith("OSS") ? "OSS" : "NC" }) && isOpenStatus(r.stato)
      ).length;

      const numeroChiuse = rilievi.filter((r) => isClosedStatus(r.stato)).length;

      const totaleDocumenti = elaboratiVerificati.length;

      let buffer: Buffer;

      try {
        const zipDocx = new PizZip(templateBuffer);

        const doc = new Docxtemplater(zipDocx, {
          paragraphLoop: true,
          linebreaks: true,
        });

        doc.render({
          rev_scheda: revisioneScheda,
          data_rev_scheda: dataRevisioneScheda,
          descrizione_rev_scheda: descrizioneRevisioneScheda(revisioneScheda),
          responsabile_pcq: responsabilePcq,
          responsabile_its: responsabileIts,
          data_risposta_progettista: dataRispostaProgettista,
          data_riscontro_ispettore: dataRiscontroIspettore,
          revisioni_scheda: revisioniScheda,
          Codice_SP: codiceScheda,
          Titolo_progetto: titoloProgetto,
          Fase_di_progetto: faseProgetto,
          "Fase_ di_ progetto": faseProgetto,
          DISCIPLINA: disciplina,
          Nota_ricezione_elaborati: notaRicezione,
          Data_ricezione: dataRicezione,
          Nome_redattore: nomeRedattore,
          rilievi,
          elaborati: elencoRows,
          elaborati_verificati: elaboratiVerificati,
          numero_nc: numeroNC,
          numero_oss: numeroOSS,
          numero_chiuse: numeroChiuse,
          totale_documenti: totaleDocumenti,
          riepilogo_finale: `NC=${numeroNC}
OSS=${numeroOSS}
Chiuse=${numeroChiuse}
Totale documenti=${totaleDocumenti}`,
        });

        buffer = doc.getZip().generate({
          type: "nodebuffer",
          compression: "DEFLATE",
        });
        buffer = await postProcessSchedaIspettivaDocx(
          buffer,
          {
            totaleElaboratiAnalizzati: totaleDocumenti,
            totaleNC: numeroNC,
            totaleOSS: numeroOSS,
            totaleChiuse: numeroChiuse,
          },
          revisioniScheda
        );
      } catch (e: any) {
        outputZip.file(
          `ERRORE_TEMPLATE_${safeName(disciplina)}.txt`,
          JSON.stringify(e, null, 2)
        );
        continue;
      }

      outputZip.file(
        `${codiceScheda}_${safeName(disciplina)}_${revisioneScheda}.docx`,
        buffer
      );
    }

    const zipBuffer = await outputZip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
    });

    return new NextResponse(zipBuffer as any, {
      status: 200,
      headers: {
        "Content-Type": "application/zip",
        "Content-Disposition": 'attachment; filename="SCHEDE_ISPETTIVE_OUTPUT_V5.zip"',
        "Content-Length": zipBuffer.length.toString(),
      },
    });
  } catch (err: any) {
    console.error(err);

    return NextResponse.json({
      ok: false,
      error: err.message,
    });
  }
}
