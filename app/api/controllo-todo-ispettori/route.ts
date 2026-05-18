import { NextRequest, NextResponse } from "next/server";
import * as XLSX from "xlsx";
import JSZip from "jszip";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const TAGS_AMMESSI = ["NC", "OSS", "Nessun rilievo", "Da NC a OSS", "Nessun Rilievo"];
const STATUS_AMMESSI = ["New", "Closed", "Waiting", "Open", "Aperta", "Chiusa"];

type BcfTopic = {
  sourceName: string;
  tr: string;
  title: string;
  description: string;
  commentsPrg: string;
  commentsIsp: string;
  allComments: string;
};

function xmlDecode(value: string) {
  return String(value || "")
    .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, "$1")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .trim();
}

function stripTags(value: string) {
  return xmlDecode(String(value || "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " "));
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
    .map((part) => part.trim())
    .filter(Boolean);
}

function hasDocumentiGenerali(value: string) {
  return splitDiscipline(value).some((part) => normalizeText(part) === "documenti generali");
}

function hasAnyValidDiscipline(value: string, disciplineAmmesse: Set<string>) {
  const parts = splitDiscipline(value);
  if (parts.length === 0) return false;
  if (hasDocumentiGenerali(value)) return true;
  return parts.some((part) => disciplineAmmesse.has(part.toLowerCase()));
}

function getTagText(xml: string, tagName: string) {
  const re = new RegExp(`<(?:[A-Za-z0-9_]+:)?${tagName}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/(?:[A-Za-z0-9_]+:)?${tagName}>`, "i");
  const match = String(xml || "").match(re);
  return match ? stripTags(match[1]) : "";
}

function getTagTexts(xml: string, tagName: string) {
  const re = new RegExp(`<(?:[A-Za-z0-9_]+:)?${tagName}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/(?:[A-Za-z0-9_]+:)?${tagName}>`, "gi");
  const values: string[] = [];
  let match: RegExpExecArray | null;

  while ((match = re.exec(String(xml || ""))) !== null) {
    const value = stripTags(match[1]);
    if (value) values.push(value);
  }

  return values;
}

function getTagBlocks(xml: string, tagName: string) {
  const re = new RegExp(`<(?:[A-Za-z0-9_]+:)?${tagName}(?:\\s[^>]*)?>[\\s\\S]*?<\\/(?:[A-Za-z0-9_]+:)?${tagName}>`, "gi");
  return String(xml || "").match(re) || [];
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

function extractITLabel(value: string) {
  const text = String(value || "").toUpperCase();
  const it = text.match(/\bIT(?:\d+)?[-_\s]*0*(\d+[A-Z]?)\b/i);
  return it ? `IT-${it[1]}` : "";
}

function extractComparableCode(value: string) {
  const match = String(value || "").match(/PV\d{3}-[A-Z0-9]+-[A-Z0-9]+-[A-Z]{3}-\d{5}-[A-Z]{3}-\d{6}(?:[-_ ]?\d+)?/i);
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

function isPrgAuthor(author: string) {
  const a = normalizeText(author);
  return (
    a.includes("prg") ||
    a.includes("progett") ||
    a.includes("pcq") ||
    a.includes("rtp") ||
    a.includes("mandante") ||
    a.includes("consorzio") ||
    a.includes("committente")
  );
}

function formatComment(date: string, author: string, text: string) {
  const cleanDate = String(date || "").slice(0, 10);
  const parts = [cleanDate, author].filter(Boolean).join(" - ");
  return parts ? `${parts}: ${text}` : text;
}

function getCommentBlocks(xml: string) {
  return getTagBlocks(xml, "Comment").filter((block) =>
    /<(?:[A-Za-z0-9_]+:)?(Date|Author|ModifiedAuthor)[\s>]/i.test(block)
  );
}

function extractCommentText(block: string) {
  const withoutOuterStart = block.replace(/^<[^>]*Comment[^>]*>/i, "");
  const inner = withoutOuterStart.match(/<(?:[A-Za-z0-9_]+:)?Comment(?:\s[^>]*)?>([\s\S]*?)<\/(?:[A-Za-z0-9_]+:)?Comment>/i);
  if (inner) return stripTags(inner[1]);

  return (
    getTagText(block, "Text") ||
    getTagText(block, "Description") ||
    getTagText(block, "Body") ||
    getTagText(block, "Content")
  );
}

async function readXlsxRows(file: File, preferredSheetName?: string) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetName =
    preferredSheetName && workbook.SheetNames.includes(preferredSheetName)
      ? preferredSheetName
      : workbook.SheetNames[0];

  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    defval: "",
  }) as any[][];
}

async function parseBcfFiles(files: File[]) {
  const topics: BcfTopic[] = [];

  async function parseMarkupXml(xml: string, sourceName: string) {
    const topicBlock = getTagBlocks(xml, "Topic")[0] || "";
    const title = getTagText(topicBlock, "Title") || getTagText(xml, "Title");
    const description = getTagText(topicBlock, "Description") || getTagText(xml, "Description");
    const topicLabels = getTagTexts(topicBlock, "Label").join(" ");

    const prgComments: string[] = [];
    const ispComments: string[] = [];
    const allComments: string[] = [];

    for (const block of getCommentBlocks(xml)) {
      const author = getTagText(block, "Author") || getTagText(block, "ModifiedAuthor") || "";
      const date = getTagText(block, "Date") || getTagText(block, "ModifiedDate") || "";
      const text = extractCommentText(block);
      if (!text) continue;

      const formatted = formatComment(date, author, text);
      allComments.push(formatted);

      if (isIspAuthor(author)) ispComments.push(formatted);
      else if (isPrgAuthor(author)) prgComments.push(formatted);
      else prgComments.push(formatted);
    }

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

    if (!tr && !title && !description && allComments.length === 0 && !comparableCode) return;

    topics.push({
      sourceName,
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
    const entries = Object.values(zip.files).filter((entry) => {
      const name = entry.name.toLowerCase();
      return !entry.dir && (name.endsWith("markup.bcf") || name.endsWith(".bcf") || name.endsWith(".xml"));
    });

    for (const entry of entries) {
      try {
        await parseMarkupXml(await entry.async("text"), entry.name);
      } catch {
        // ignora file non XML dentro BCFZIP/ZIP
      }
    }
  }

  return topics;
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

function textScore(source: string, needle: string, points: number) {
  if (!needle) return 0;
  return normalizeCode(source).includes(normalizeCode(needle)) ? points : 0;
}

function wordOverlapScore(source: string, value: string, pointsPerWord: number, maxPoints: number) {
  const sourceText = normalizeText(source);
  const words = normalizeText(value)
    .split(" ")
    .filter((word) => word.length >= 6)
    .slice(0, 20);

  let score = 0;
  words.forEach((word) => {
    if (sourceText.includes(word)) score += pointsPerWord;
  });

  return Math.min(score, maxPoints);
}

function findBcfTopic(
  tr: string,
  topics: BcfTopic[],
  title = "",
  description = "",
  codiceReport = "",
  label = "",
  disciplina = ""
): BcfTopic | null {
  const normalizedTr = normalizeCode(tr);
  const labelIT = extractITLabel(label);
  const comparableCode =
    extractComparableCode(title) ||
    extractComparableCode(codiceReport) ||
    extractComparableCode(description);

  let bestTopic: BcfTopic | null = null;
  let bestScore = 0;

  topics.forEach((topic) => {
    const haystack = `${topic.sourceName} ${topic.tr} ${topic.title} ${topic.description} ${topic.allComments}`;

    const labelMatch = !!labelIT && normalizeCode(haystack).includes(normalizeCode(labelIT));
    const codeMatch = !!comparableCode && normalizeCode(haystack).includes(normalizeCode(comparableCode));
    const trExact = !!normalizedTr && normalizeCode(topic.tr) === normalizedTr;
    const trText = !!normalizedTr && normalizeCode(haystack).includes(normalizedTr);

    const disciplinaScore = wordOverlapScore(haystack, disciplina, 25, 120);
    const titleScore = wordOverlapScore(haystack, title, 12, 150);
    const descriptionScore = wordOverlapScore(haystack, description, 8, 220);
    const semanticScore = disciplinaScore + titleScore + descriptionScore;

    let score = 0;

    if (labelMatch) score += 1200;

    if (codeMatch && (labelMatch || semanticScore >= 80)) score += 500;

    if (trExact) score += 180;
    else if (trText) score += 100;

    score += semanticScore;

    if (labelIT && !labelMatch && score < 350) score = 0;

    if (score > bestScore) {
      bestScore = score;
      bestTopic = topic;
    }
  });

  return bestScore >= 300 ? bestTopic : null;
}

function buildReportCodes(reportRows: any[][]) {
  const map = new Map<string, string>();
  reportRows.slice(1).forEach((row) => {
    const codiceReport = String(row[2] || "").trim();
    const normalized = normalizeCode(codiceReport);
    if (normalized && normalized !== "NAN") map.set(normalized, codiceReport);
  });
  return map;
}

function buildDiscipline(elencoRows: any[][]) {
  const set = new Set<string>();

  elencoRows.slice(1).forEach((row) => {
    [0, 1, 5].forEach((idx) => {
      const value = String(row[idx] || "").trim();
      if (value) set.add(value.toLowerCase());
    });
  });

  [
    "Documenti generali",
    "Generale",
    "Ambiente e vincoli",
    "Strutture",
    "Sicurezza cantierizzazione e BOB",
    "Sicurezza",
    "Interferenze e espropri",
    "Interferenze",
    "Economico",
    "Computi",
    "Bonifica bellica",
  ].forEach((d) => set.add(d.toLowerCase()));

  return set;
}

function buildChecks(todoRows: any[][], reportCodes: Map<string, string>, disciplineAmmesse: Set<string>, bcfTopics: BcfTopic[]) {
  return todoRows.slice(1).map((row, index) => {
    const label = String(row[0] || "").trim();
    const title = String(row[1] || "").trim();
    const description = String(row[2] || "").trim();
    const status = String(row[5] || "").trim();
    const disciplina = String(row[8] || "").trim();
    const tags = String(row[9] || "").trim();

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
      warning.push("Codice elaborato non presente nel Report_Completo.xlsx");
    } else {
      titleOk = true;
    }

    const tagsOk =
      !tags ||
      TAGS_AMMESSI.some((tag) => tag.toLowerCase() === tags.toLowerCase());

    if (!tags) warning.push("Tags mancanti");
    else if (!tagsOk) warning.push("Tags non riconosciuto");

    const disciplinaPresente = !!disciplina;
    const disciplinaOk =
      disciplinaPresente &&
      (hasDocumentiGenerali(disciplina) ||
        hasAnyValidDiscipline(disciplina, disciplineAmmesse));

    if (!disciplinaPresente) {
      anomalie.push("Disciplina mancante");
    } else if (!disciplinaOk) {
      warning.push("Disciplina non presente in ELENCO_ELABORATI.xlsx");
    }

    const statusOk =
      !status ||
      STATUS_AMMESSI.some((s) => s.toLowerCase() === status.toLowerCase());

    if (!status) warning.push("Status mancante");
    else if (!statusOk) warning.push("Status non riconosciuto");

    const tr = extractTR(label) || extractTR(title) || extractTR(description);
    const bcf: BcfTopic | null = findBcfTopic(
      tr,
      bcfTopics,
      title,
      description,
      codiceReport,
      label,
      disciplina
    );

    const isRilievo = ["NC", "OSS", "Da NC a OSS"].some((tag) => tag.toLowerCase() === tags.toLowerCase());
    const isClosed = status.toLowerCase() === "closed" || status.toLowerCase() === "chiusa";

    let esitoStoria = "NON APPLICABILE";

    if (isRilievo) {
      if (!tr && !extractComparableCode(title) && !extractComparableCode(codiceReport)) {
        esitoStoria = "BCF NON TROVATO";
        warning.push("Codice TR o codice elaborato non trovato nel ToDo");
      } else if (!bcf) {
        esitoStoria = "BCF NON TROVATO";
        warning.push(`BCF non trovato per ${tr || codiceReport || title}`);
      } else if (!bcf.commentsPrg && !bcf.allComments) {
        esitoStoria = "MANCA RISPOSTA PROGETTISTA";
        warning.push("Manca risposta progettista nei commenti BCF");
      } else if (!bcf.commentsIsp && isClosed) {
        esitoStoria = "CHIUSO SENZA RISCONTRO";
        warning.push("ToDo chiuso senza riscontro ispettore nei BCF");
      } else if (!bcf.commentsIsp) {
        esitoStoria = "MANCA RISCONTRO ITS";
        warning.push("Manca riscontro ispettore nei commenti BCF");
      } else {
        esitoStoria = "COMPLETA";
      }
    }

    const esito = titleOk && disciplinaPresente && statusOk ? "OK" : "ERRORE";
    const livello = esito === "ERRORE" ? "ERRORE" : warning.length ? "WARNING" : "OK";

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
      storiaOk: esitoStoria === "COMPLETA" || esitoStoria === "NON APPLICABILE",
      esitoStoria,
      esito,
      livello,
      warning,
      anomalie,
    };
  });
}

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();

    const todoFile = formData.get("todo");
    const reportFile = formData.get("report");
    const elencoFile = formData.get("elenco");
    const bcfFiles = formData
      .getAll("bcf")
      .filter((file): file is File => file instanceof File && file.size > 0);

    if (!(todoFile instanceof File)) {
      return NextResponse.json({ ok: false, error: "ToDo XLSX mancante" }, { status: 400 });
    }

    if (!(reportFile instanceof File)) {
      return NextResponse.json({ ok: false, error: "Report_Completo.xlsx mancante" }, { status: 400 });
    }

    if (!(elencoFile instanceof File)) {
      return NextResponse.json({ ok: false, error: "ELENCO_ELABORATI.xlsx mancante" }, { status: 400 });
    }

    const [todoRows, reportRows, elencoRows, bcfTopics] = await Promise.all([
      readXlsxRows(todoFile),
      readXlsxRows(reportFile, "Verifica Elaborati"),
      readXlsxRows(elencoFile),
      parseBcfFiles(bcfFiles),
    ]);

    const reportCodes = buildReportCodes(reportRows);
    const disciplineAmmesse = buildDiscipline(elencoRows);
    const checks = buildChecks(todoRows, reportCodes, disciplineAmmesse, bcfTopics);

    const totale = checks.length;
    const ok = checks.filter((row) => row.esito === "OK").length;
    const errori = checks.filter((row) => row.esito === "ERRORE").length;
    const warning = checks.filter((row) => row.livello === "WARNING").length;
    const storieComplete = checks.filter((row) => row.esitoStoria === "COMPLETA").length;
    const bcfWarning = checks.filter((row) =>
      row.esitoStoria === "BCF NON TROVATO" ||
      row.esitoStoria === "MANCA RISPOSTA PROGETTISTA" ||
      row.esitoStoria === "MANCA RISCONTRO ITS" ||
      row.esitoStoria === "CHIUSO SENZA RISCONTRO"
    ).length;
    const completezza = totale > 0 ? Math.round((ok / totale) * 100) : 0;

    return NextResponse.json({
      ok: true,
      checks,
      bcfTopicsCount: bcfTopics.length,
      summary: {
        totale,
        ok,
        errori,
        warning,
        storieComplete,
        bcfWarning,
        completezza,
      },
    });
  } catch (error: any) {
    console.error("Errore controllo ToDo ispettori:", error);

    return NextResponse.json(
      {
        ok: false,
        error: error?.message || "Errore durante il controllo ToDo ispettori",
      },
      { status: 500 }
    );
  }
}
