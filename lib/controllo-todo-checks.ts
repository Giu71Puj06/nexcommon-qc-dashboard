// lib/controllo-todo-checks.ts
//
// Costruisce le righe di controllo del "Controllo ToDo ispettori" a partire dai
// ToDo importati direttamente da Trimble (endpoint backend /tc/todos con
// with_comments=true e with_snapshots=true).
//
// A differenza del flusso a file (todo.xlsx + .bcf), qui ogni ToDo porta con sé
// i propri commenti e le immagini dei markup 2D, quindi NON serve il matching
// fuzzy: ogni ToDo genera una riga con i suoi commenti e le sue miniature.
//
// Le immagini sono restituite come data URL pronte per <img src>.

export type TrimbleComment = {
  author?: string;
  date?: string;
  text?: string;
};

export type TrimbleSnapshot = {
  attachment_id?: string;
  name?: string;
  type?: string;
  mime?: string;
  image_base64?: string;
};

export type TrimbleTodo = {
  id?: string;
  label?: string;
  title?: string;
  description?: string;
  status?: string;
  assignees?: string;
  tags?: string[] | string;
  comments?: TrimbleComment[];
  snapshots?: TrimbleSnapshot[];
};

export type CheckRow = {
  rowNumber: number;
  progressivo: number;
  label: string;
  title: string;
  description: string;
  codiceTitleTrimble: string;
  codiceReport: string;
  titleOk: boolean;
  tags: string;
  tagsOk: boolean;
  disciplina: string;
  disciplinaOk: boolean;
  status: string;
  statusOk: boolean;
  tr: string;
  bcfTitle: string;
  bcfDescription: string;
  rispostaProgettista: string;
  riscontroIspettore: string;
  storiaOk: boolean;
  esitoStoria:
    | "COMPLETA"
    | "Manca commento del progettista"
    | "Manca il riscontro dell'ispettore"
    | "NON APPLICABILE";
  esito: "OK" | "ERRORE";
  livello: "OK" | "WARNING" | "ERRORE";
  warning: string[];
  anomalie: string[];
  images: string[];
};

const TAGS_AMMESSI = ["NC", "OSS", "Nessun rilievo", "Da NC a OSS", "Nessun Rilievo"];
// Il backend normalizza lo stato Trimble in italiano: Aperta / In attesa / Fatto / Chiusa.
// Manteniamo anche gli stati inglesi per compatibilità con eventuali export.
const STATUS_AMMESSI = [
  "New", "Closed", "Waiting", "Open",
  "Aperta", "Chiusa", "In attesa", "Fatto",
];

function clean(v: any): string {
  return String(v ?? "").trim();
}

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

function splitDiscipline(value: string) {
  return String(value || "")
    .split(/[,;\/|]+|\s+-\s+|\n+/g)
    .map((p) => p.trim())
    .filter(Boolean);
}

function hasDocumentiGenerali(value: string) {
  return splitDiscipline(value).some((p) => normalizeText(p) === "documenti generali");
}

function hasAnyValidDiscipline(value: string, disciplineAmmesse: Set<string>) {
  const parts = splitDiscipline(value);
  if (parts.length === 0) return false;
  if (hasDocumentiGenerali(value)) return true;
  return parts.some((p) => disciplineAmmesse.has(p.toLowerCase()));
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

function extractComparableCode(value: string) {
  const match = String(value || "").match(
    /PV\d{3}-[A-Z0-9]+-[A-Z0-9]+-[A-Z]{3}-\d{5}-[A-Z]{3}-\d{6}(?:[-_ ]?\d+)?/i
  );
  return match ? normalizeCode(match[0]) : "";
}

function isDescrizioneGenerale(value: string) {
  const text = String(value || "").trim();
  if (!text) return false;
  if (/\.pdf/i.test(text)) return false;
  return !extractComparableCode(text);
}

function isKnownInspectorAuthor(author: string) {
  const a = normalizeText(author);
  return (
    a.includes("massimo tamberi") || a.includes("edoardo oddo") ||
    a.includes("guido bonin") || a.includes("ilaria martarelli") ||
    a.includes("michea sciorra") || a.includes("stefano arcangeli") ||
    a.includes("clara soliman") || a.includes("clara") ||
    a.includes("(isp)") || a.includes(" isp ") ||
    a.includes("non conformita superata") || a.includes("non conformità superata") ||
    a.includes("osservazione superata") || a.includes("si prende atto") ||
    a.includes("massimo") || a.includes("edoardo") || a.includes("guido") ||
    a.includes("ilaria") || a.includes("michea") || a.includes("stefano")
  );
}

function isIspAuthor(author: string) {
  const a = normalizeText(author);
  return (
    a.includes("its") || a.includes("isp") || a.includes("controlli tecnici") ||
    a.includes("odi") || a.includes("ispett") || isKnownInspectorAuthor(author)
  );
}

function isPrgAuthor(author: string) {
  const a = normalizeText(author);
  return (
    a.includes("prg") || a.includes("progett") || a.includes("pcq") ||
    a.includes("rtp") || a.includes("mandante") || a.includes("consorzio") ||
    a.includes("committente") || a.includes("(prg)") || a.includes(" prg ") ||
    a.includes("giuseppe pizzi") || a.includes("recepito") ||
    a.includes("aggiornato") || a.includes("eliminato")
  );
}

function formatComment(date: string, author: string, text: string) {
  const cleanDate = String(date || "").slice(0, 10);
  const parts = [cleanDate, author].filter(Boolean).join(" - ");
  return parts ? `${parts}: ${text}` : text;
}

function findBestReportCodeFromTitle(title: string, reportCodes: Map<string, string>) {
  const normalizedTitle = normalizeCode(title);
  let bestNormalized = "";
  let bestOriginal = "";
  for (const [normalizedReportCode, originalReportCode] of reportCodes.entries()) {
    if (normalizedTitle === normalizedReportCode || normalizedTitle.includes(normalizedReportCode)) {
      if (normalizedReportCode.length > bestNormalized.length) {
        bestNormalized = normalizedReportCode;
        bestOriginal = originalReportCode;
      }
    }
  }
  return bestOriginal;
}

export function buildReportCodes(reportRows: any[][]): Map<string, string> {
  const map = new Map<string, string>();
  (reportRows || []).slice(1).forEach((row) => {
    const codiceReport = clean(row?.[2]);
    const normalized = normalizeCode(codiceReport);
    if (normalized && normalized !== "NAN") map.set(normalized, codiceReport);
  });
  return map;
}

export function buildDiscipline(elencoRows: any[][]): Set<string> {
  const set = new Set<string>();
  (elencoRows || []).slice(1).forEach((row) => {
    [0, 1, 5].forEach((idx) => {
      const value = clean(row?.[idx]);
      if (value) set.add(value.toLowerCase());
    });
  });
  [
    "Documenti generali", "Generale", "Ambiente e vincoli", "Strutture",
    "Sicurezza cantierizzazione e BOB", "Sicurezza", "Interferenze e espropri",
    "Interferenze", "Economico", "Computi", "Bonifica bellica",
  ].forEach((d) => set.add(d.toLowerCase()));
  return set;
}

function todoTagsText(todo: TrimbleTodo): string {
  const t = todo.tags;
  if (Array.isArray(t)) return t.map((x) => clean(x)).filter(Boolean).join(", ");
  return clean(t);
}

function todoImages(todo: TrimbleTodo): string[] {
  const snaps = Array.isArray(todo.snapshots) ? todo.snapshots : [];
  return snaps
    .map((s) => {
      const b64 = clean(s?.image_base64);
      if (!b64) return "";
      const mime = clean(s?.mime) || "image/png";
      return `data:${mime};base64,${b64}`;
    })
    .filter(Boolean);
}

function buildRow(
  todo: TrimbleTodo,
  index: number,
  reportCodes: Map<string, string>,
  disciplineAmmesse: Set<string>
): CheckRow {
  const label = clean(todo.label);
  const title = clean(todo.title);
  const description = clean(todo.description);
  const status = clean(todo.status);
  const disciplina = clean(todo.assignees);
  const tags = todoTagsText(todo);

  const anomalie: string[] = [];
  const warning: string[] = [];

  const codiceTitleTrimble = title.replace(/\.pdf$/i, "").trim();
  const titleContienePdf = /\.pdf/i.test(title);
  const titleDescrittivo = isDescrizioneGenerale(title);
  const codiceReport = titleDescrittivo ? "" : findBestReportCodeFromTitle(title, reportCodes);

  let titleOk = false;
  if (!title) {
    anomalie.push("Title mancante");
  } else if (titleContienePdf) {
    anomalie.push("Il Title contiene .pdf");
  } else if (titleDescrittivo) {
    titleOk = true;
  } else if (!codiceReport) {
    titleOk = true;
    if (reportCodes.size > 0) warning.push("Codice elaborato non presente nel Report_Completo.xlsx");
  } else {
    titleOk = true;
  }

  const tagsOk = !tags || TAGS_AMMESSI.some((t) => t.toLowerCase() === tags.toLowerCase());
  if (!tags) warning.push("Tags mancanti");
  else if (!tagsOk) warning.push("Tags non riconosciuto");

  const disciplinaPresente = !!disciplina;
  const disciplinaOk =
    disciplinaPresente &&
    (hasDocumentiGenerali(disciplina) || hasAnyValidDiscipline(disciplina, disciplineAmmesse));
  if (!disciplinaPresente) anomalie.push("Disciplina mancante");
  else if (!disciplinaOk && disciplineAmmesse.size > 0)
    warning.push("Disciplina non presente in ELENCO_ELABORATI.xlsx");

  const statusOk = !status || STATUS_AMMESSI.some((s) => s.toLowerCase() === status.toLowerCase());
  if (!status) warning.push("Status mancante");
  else if (!statusOk) warning.push("Status non riconosciuto");

  const tr = extractTR(label) || extractTR(title) || extractTR(description);

  // Commenti dello stesso ToDo: classificazione progettista / ispettore.
  const prgComments: string[] = [];
  const ispComments: string[] = [];
  const allComments: string[] = [];
  (Array.isArray(todo.comments) ? todo.comments : []).forEach((c) => {
    const author = clean(c?.author);
    const date = clean(c?.date);
    const text = clean(c?.text);
    if (!text) return;
    const formatted = formatComment(date, author, text);
    allComments.push(formatted);
    const classifier = `${author} ${text}`;
    if (isIspAuthor(classifier)) ispComments.push(formatted);
    else prgComments.push(formatted);
  });

  const isRilievo = ["NC", "OSS", "Da NC a OSS"].some((t) => t.toLowerCase() === tags.toLowerCase());
  const s = status.toLowerCase();
  const isClosed = s === "closed" || s === "chiusa" || s === "fatto";

  let esitoStoria: CheckRow["esitoStoria"] = "NON APPLICABILE";
  if (isRilievo) {
    if (isClosed) {
      esitoStoria = "COMPLETA";
    } else if (!prgComments.length && !allComments.length) {
      esitoStoria = "Manca commento del progettista";
      warning.push("Risposta progettista mancante nei commenti");
    } else if (!prgComments.length) {
      esitoStoria = "Manca commento del progettista";
      warning.push("Risposta progettista non individuata nei commenti");
    } else if (!ispComments.length) {
      esitoStoria = "Manca il riscontro dell'ispettore";
      warning.push("Riscontro ispettore mancante nei commenti");
    } else {
      esitoStoria = "COMPLETA";
    }
  }

  const esito: "OK" | "ERRORE" = titleOk && disciplinaPresente && statusOk ? "OK" : "ERRORE";
  const livello: "OK" | "WARNING" | "ERRORE" =
    esito === "ERRORE" ? "ERRORE" : warning.length ? "WARNING" : "OK";

  return {
    rowNumber: index + 2,
    progressivo: index + 1,
    label,
    title,
    description: codiceReport ? description : description,
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
    bcfTitle: title,
    bcfDescription: description,
    rispostaProgettista: prgComments.join("\n\n") || (ispComments.length ? "" : allComments.join("\n\n")),
    riscontroIspettore: ispComments.join("\n\n"),
    storiaOk: esitoStoria === "COMPLETA" || esitoStoria === "NON APPLICABILE",
    esitoStoria,
    esito,
    livello,
    warning,
    anomalie,
    images: todoImages(todo),
  };
}

export function buildChecksFromTrimbleTodos(
  todos: TrimbleTodo[],
  reportRows: any[][] = [],
  elencoRows: any[][] = []
): { checks: CheckRow[]; summary: any } {
  const reportCodes = buildReportCodes(reportRows);
  const disciplineAmmesse = buildDiscipline(elencoRows);

  const checks = (Array.isArray(todos) ? todos : []).map((t, i) =>
    buildRow(t || {}, i, reportCodes, disciplineAmmesse)
  );

  const totale = checks.length;
  const ok = checks.filter((r) => r.esito === "OK").length;
  const errori = checks.filter((r) => r.esito === "ERRORE").length;
  const warning = checks.filter((r) => r.livello === "WARNING").length;
  const storieComplete = checks.filter((r) => r.esitoStoria === "COMPLETA").length;
  const bcfWarning = checks.filter(
    (r) =>
      r.esitoStoria === "Manca commento del progettista" ||
      r.esitoStoria === "Manca il riscontro dell'ispettore"
  ).length;
  const snapshotsTotal = checks.reduce((n, r) => n + r.images.length, 0);
  const completezza = totale > 0 ? Math.round((ok / totale) * 100) : 0;

  return {
    checks,
    summary: { totale, ok, errori, warning, storieComplete, bcfWarning, snapshotsTotal, completezza },
  };
}
