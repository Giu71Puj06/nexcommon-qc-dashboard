import { NextResponse } from "next/server";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

function arr<T>(v: T | T[] | undefined): T[] {
  if (!v) return [];
  return Array.isArray(v) ? v : [v];
}

function normalize(v = "") {
  return String(v)
    .toLowerCase()
    .replace(/\.pdf/gi, "")
    .replace(/[^a-z0-9àèéìòù]+/gi, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeElaboratoCode(v = "") {
  return String(v || "")
    .toUpperCase()
    .replace(/\.PDF$/i, "")
    .replace(/[^A-Z0-9]/g, "");
}

function cleanText(v: any) {
  return String(v ?? "")
    .replace(/\r/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/[ \t]{2,}/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function xmlDecode(value = "") {
  return String(value || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function getAny(obj: any, keys: string[]) {
  if (!obj || typeof obj !== "object") return "";
  for (const k of keys) {
    if (obj[k] !== undefined && obj[k] !== null) return obj[k];
  }
  return "";
}

function getXmlText(value: any) {
  if (!value) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number") return String(value);
  if (typeof value === "object") {
    return (
      value["#text"] ||
      value.text ||
      value.Text ||
      value.Value ||
      value.value ||
      ""
    );
  }
  return String(value);
}

function extractBcfLabels(value: any): string {
  if (!value) return "";

  if (typeof value === "string" || typeof value === "number") {
    return getXmlText(value);
  }

  if (Array.isArray(value)) {
    return value
      .map((item) => extractBcfLabels(item))
      .filter(Boolean)
      .join(" ");
  }

  if (typeof value === "object") {
    const directText = getXmlText(value);
    const nestedLabels = [
      value.Label,
      value.label,
      value.Labels,
      value.labels,
      value.TopicLabel,
      value.topicLabel,
      value.TopicLabels,
      value.topicLabels,
    ]
      .map((item) => extractBcfLabels(item))
      .filter(Boolean)
      .join(" ");

    return [directText, nestedLabels].filter(Boolean).join(" ").trim();
  }

  return "";
}

function roleFromText(text = "") {
  const m = String(text).match(/\((ISP|PRG)\)/i);
  return m ? m[1].toUpperCase() : "";
}

function detectTipo(tags = "", description = "") {
  const t = String(tags).toUpperCase();
  const d = String(description).toUpperCase();
  const text = `${t} ${d}`;

  if (text.includes("DA NC A OSS") || text.includes("DA NC A OS")) return "Da NC a OSS";

  if (text.includes("NESSUN RILIEVO")) {
    return "Nessun rilievo";
  }

  if (/(^|[^A-Z0-9])OSS([^A-Z0-9]|$)/.test(text)) return "OSS";
  if (/(^|[^A-Z0-9])NC([^A-Z0-9]|$)/.test(text)) return "NC";

  return "Esito mancante";
}

function includesAny(text = "", words: string[]) {
  const n = normalize(text);
  return words.some((w) => n.includes(normalize(w)));
}

function detectTipologiaNcOss(tags = "", description = "", title = "", tipo = "") {
  if (tipo === "Nessun rilievo" || tipo === "Esito mancante") return "";

  const text = `${tags} ${description} ${title}`;

  if (includesAny(text, ["normativa", "normative", "normativa vigente", "conformità", "conforme", "non conforme", "ntc", "eurocodice", "eurocodici", "codice appalti", "verifica normativa", "verifiche obbligatorie", "verifica obbligatoria", "prescrizione normativa", "classificazione opere", "autorizzazione", "autorizzativo", "vincolo", "vincoli", "prescrizioni", "prescrizione"])) return "1. Normative";
  if (includesAny(text, ["incoerenza", "incoerenze", "discordanza", "discordanze", "non coerente", "non coerenti", "differenza tra", "difformità tra", "non allineato", "non allineati", "disallineamento", "contraddizione", "contraddizioni", "relazione e tavola", "relazione tavola", "elaborati non coerenti", "tavole non coerenti"])) return "2. Incoerenze tra elaborati";
  if (includesAny(text, ["mancante", "mancanti", "manca", "non presente", "non presenti", "non indicato", "non indicati", "non riportato", "non riportati", "assente", "assenti", "incompleto", "incompleta", "incompleti", "incomplete", "omesso", "omessa", "non risulta", "necessario integrare", "integrare"])) return "3. Informazioni mancanti / incomplete";
  if (includesAny(text, ["chiarire", "si chiede", "si richiede", "richiesta di chiarimento", "chiarimento", "chiarimenti", "specificare", "precisare", "verificare", "si invita", "dettagliare", "approfondire", "motivare", "esplicitare"])) return "4. Richieste di chiarimento";
  if (includesAny(text, ["relazione", "relazioni", "documento", "documentazione", "elaborato", "elaborati", "relazione tecnica", "relazione specialistica", "rapporto", "allegato", "capitolo", "paragrafo", "pag", "pagina", "indice", "testo"])) return "5. Elaborati e relazioni";
  if (includesAny(text, ["quota", "quote", "dimensione", "dimensioni", "dimensionale", "dimensionali", "misura", "misure", "altezza", "larghezza", "spessore", "diametro", "sezione", "scala", "geometria", "geometrico", "geometrici"])) return "6. Errori dimensionali / quote";
  if (includesAny(text, ["dettaglio", "dettagli", "particolare", "particolari", "nodo", "nodi", "sezione costruttiva", "dettaglio costruttivo", "particolare costruttivo", "schema costruttivo", "dettagli esecutivi", "esecutivo", "esecutivi"])) return "7. Dettagli costruttivi insufficienti";
  if (includesAny(text, ["computo", "computi", "quantità", "quantita", "voce", "voci", "elenco prezzi", "prezzario", "prezzari", "prezzo", "prezzi", "stima", "stime", "stima economica", "quadro economico", "importo", "contabilità", "contabilita", "misurazione"])) return "8. Computi e quantità";
  if (includesAny(text, ["cantiere", "realizzazione", "realizzabile", "non realizzabile", "esecuzione", "esecutabilità", "esecutabilita", "costruttibilità", "costruttibilita", "fattibilità", "fattibilita", "posa", "montaggio", "lavorazione", "manutenzione", "manutenibilità", "manutenibilita", "accessibilità manutentiva", "accessibilita manutentiva"])) return "9. Costruttibilità / fattibilità";
  if (includesAny(text, ["interferenza", "interferenze", "clash", "sovrapposizione", "sovrapposizioni", "conflitto", "conflitti", "interferisce", "interferiscono", "collisione", "collisioni", "coordinamento interdisciplinare", "coordinamento bim", "bim"])) return "10. Interferenze / clash";

  return "Altre";
}

function similarity(a = "", b = "") {
  const aa = normalize(a);
  const bb = normalize(b);

  if (!aa || !bb) return 0;
  if (aa === bb) return 1;
  if (aa.includes(bb) || bb.includes(aa)) return 0.9;

  const aWords = new Set(aa.split(" "));
  const bWords = new Set(bb.split(" "));
  const intersection = [...aWords].filter((w) => bWords.has(w)).length;
  const union = new Set([...aWords, ...bWords]).size;

  return union ? intersection / union : 0;
}

function normalizeCommentKey(comment: any) {
  return [normalize(comment?.role || ""), normalize(comment?.author || ""), normalize(comment?.date || ""), normalize(comment?.comment || "")].join("|");
}

function uniqueComments(comments: any[] = []) {
  const seen = new Set<string>();
  const result: any[] = [];

  for (const c of comments || []) {
    const key = normalizeCommentKey(c);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    result.push(c);
  }

  return result;
}

function findBestComments(todo: any, commentsByTopic: Map<string, any[]>, matchedBcfTopic?: any) {
  const matchedGuid = normalize(matchedBcfTopic?.Guid || matchedBcfTopic?.GUID || matchedBcfTopic?.ID || matchedBcfTopic?.Label || "");

  if (matchedGuid && commentsByTopic.has(matchedGuid)) {
    return uniqueComments(commentsByTopic.get(matchedGuid) || []);
  }

  const candidates = [todo.Label, todo.ID, todo.Guid, todo.GUID].filter(Boolean);
  const collected: any[] = [];
  const usedKeys = new Set<string>();

  for (const c of candidates) {
    const key = normalize(c);
    if (!key || usedKeys.has(key)) continue;
    usedKeys.add(key);

    if (commentsByTopic.has(key)) {
      collected.push(...(commentsByTopic.get(key) || []));
    }
  }

  return uniqueComments(collected);
}

function translateStatus(status = "") {
  const s = String(status || "").trim();

  if (/^closed$/i.test(s)) return "Chiusa";
  if (/^new$/i.test(s)) return "Aperta";
  if (/^waiting$/i.test(s)) return "In attesa";
  if (/^unknown$/i.test(s)) return "Non definito";
  if (/^chiuso$/i.test(s)) return "Chiusa";
  if (/^chiusa$/i.test(s)) return "Chiusa";
  if (/^aperto$/i.test(s)) return "Aperta";
  if (/^aperta$/i.test(s)) return "Aperta";

  return s;
}

function getFirstColumnValue(row: any) {
  if (!row || typeof row !== "object") return "";
  const firstKey = Object.keys(row)[0];
  return firstKey ? row[firstKey] : "";
}

function getTodoLabel(todo: any) {
  return todo.Label || getFirstColumnValue(todo) || todo.ID || todo.Guid || todo.GUID || "";
}

function getTodoAssignees(todo: any) {
  return todo["Assignee(s)"] || todo["Assignee(s) "] || todo.Assignees || todo.Assignee || todo.AssignedTo || todo["Assigned to"] || "";
}

function findBestBcfTopic(todo: any, topicsByKey: Map<string, any>, allTopics: any[]) {
  const labelKey = normalize(getTodoLabel(todo));
  const idKey = normalize(todo.ID);
  const guidKey = normalize(todo.Guid || todo.GUID);
  const titleKey = normalize(todo.Title);
  const descriptionKey = normalize(todo.Description);

  for (const key of [labelKey, idKey, guidKey].filter(Boolean)) {
    if (topicsByKey.has(key)) return topicsByKey.get(key);
  }

  const sameTitleTopics = allTopics.filter((topic) => normalize(topic.Title) === titleKey);

  if (sameTitleTopics.length === 1) return sameTitleTopics[0];

  if (sameTitleTopics.length > 1 && descriptionKey) {
    let bestScore = 0;
    let bestTopic: any = null;

    for (const topic of sameTitleTopics) {
      const score = similarity(todo.Description, topic.Description);
      if (score > bestScore) {
        bestScore = score;
        bestTopic = topic;
      }
    }

    if (bestScore >= 0.75) return bestTopic;

    const todoInt = String(todo.Description || "").match(/\bINT[_\s-]*\d+/i)?.[0];
    if (todoInt) {
      const normalizedTodoInt = normalize(todoInt);
      const byInt = sameTitleTopics.find((topic) => normalize(topic.Description).includes(normalizedTodoInt));
      if (byInt) return byInt;
    }

    return null;
  }

  if (descriptionKey) {
    let bestScore = 0;
    let bestTopic: any = null;

    for (const topic of allTopics) {
      const score = similarity(todo.Description, topic.Description);
      if (score > bestScore) {
        bestScore = score;
        bestTopic = topic;
      }
    }

    if (bestScore >= 0.9) return bestTopic;
  }

  return null;
}

function buildBcfTopicUniqueKey(topic: any) {
  return [
    normalize(topic?.sourceFile || ""),
    normalize(topic?.Guid || topic?.GUID || topic?.ID || topic?.Label || ""),
    normalize(topic?.Title || ""),
    normalize(topic?.Description || ""),
  ].join("|");
}

function buildBcfTopicsByKey(topics: any[]) {
  const map = new Map<string, any>();

  for (const topic of topics) {
    const keys = Array.from(new Set([
      normalize(topic.Label),
      normalize(topic.ID),
      normalize(topic.Guid),
      normalize(topic.Title),
      normalize(String(topic.Title || "").replace(/\.pdf/gi, "")),
    ].filter(Boolean)));

    for (const key of keys) {
      if (!map.has(key)) map.set(key, topic);
    }
  }

  return map;
}

function buildTodoRowsWithStandaloneBcf(excelRows: any[], bcfTopicRows: any[]) {
  if (!excelRows.length) return bcfTopicRows;

  const topicsByKeyForMatch = buildBcfTopicsByKey(bcfTopicRows);
  const matchedTopicKeys = new Set<string>();

  for (const todo of excelRows) {
    const matched = findBestBcfTopic(todo, topicsByKeyForMatch, bcfTopicRows);
    if (matched) matchedTopicKeys.add(buildBcfTopicUniqueKey(matched));
  }

  const standaloneBcfTopics = bcfTopicRows.filter((topic) => {
    const key = buildBcfTopicUniqueKey(topic);
    return key && !matchedTopicKeys.has(key);
  });

  return [...excelRows, ...standaloneBcfTopics];
}

function extractMarkup(parsed: any) {
  return parsed?.Markup || parsed?.markup || parsed?.bcf?.Markup || parsed?.Bcf?.Markup || parsed;
}

function extractTopic(markup: any, fallbackGuid: string) {
  const topic = getAny(markup, ["Topic", "topic"]) || {};

  const title = getXmlText(getAny(topic, ["Title", "title", "TopicTitle", "Name"]));
  const description = getXmlText(getAny(topic, ["Description", "description"]));
  const labelsRaw = getAny(topic, ["Labels", "labels", "Label", "label", "TopicLabels", "topicLabels", "TopicLabel", "topicLabel"]);
  const labels = extractBcfLabels(labelsRaw);

  const guid = getXmlText(getAny(topic, ["Guid", "guid", "GUID"])) || topic.Guid || topic.guid || fallbackGuid;

  return {
    topic,
    topicTitle: title,
    topicDescription: description,
    topicLabels: labels,
    topicGuid: String(guid || fallbackGuid),
    topicStatus: getXmlText(getAny(topic, ["TopicStatus", "Status", "status"])) || topic.TopicStatus || "",
    topicPriority: getXmlText(getAny(topic, ["Priority", "priority"])),
    topicCreationDate: getXmlText(getAny(topic, ["CreationDate", "creationDate"])),
    topicCreationAuthor: getXmlText(getAny(topic, ["CreationAuthor", "creationAuthor"])),
    topicModifiedDate: getXmlText(getAny(topic, ["ModifiedDate", "modifiedDate"])),
    topicModifiedAuthor: getXmlText(getAny(topic, ["ModifiedAuthor", "modifiedAuthor"])),
    topicAssignedTo: getXmlText(getAny(topic, ["AssignedTo", "assignedTo"])),
  };
}

function extractComments(markup: any) {
  const raw = getAny(markup, ["Comment", "comment", "Comments", "comments"]) || getAny(markup?.Comments, ["Comment", "comment"]) || getAny(markup?.comments, ["Comment", "comment"]);

  if (Array.isArray(raw)) return raw;
  if (raw?.Comment) return arr(raw.Comment);
  if (raw?.comment) return arr(raw.comment);
  return arr(raw);
}

function getMimeTypeFromPath(path = "") {
  const p = String(path || "").toLowerCase();

  if (p.endsWith(".jpg") || p.endsWith(".jpeg")) return "image/jpeg";
  if (p.endsWith(".webp")) return "image/webp";
  return "image/png";
}

function getFolderFromPath(path = "") {
  const parts = String(path || "").split("/");
  parts.pop();
  return parts.join("/");
}

function isSolibriCheckingFile(fileName = "") {
  const baseName = String(fileName || "").split("/").pop() || "";
  return /^solibri[_\s-]/i.test(baseName) || /solibri/i.test(baseName);
}

function extractSolibriCheckingRevision(fileName = "") {
  const baseName = String(fileName || "").split("/").pop() || "";
  const match =
    baseName.match(/^solibri_checking_(.+?)\.(bcf|bcfzip|zip)$/i) ||
    baseName.match(/^solibri[_\s-](.+?)\.(bcf|bcfzip|zip)$/i);

  return match ? match[1] : "";
}

async function extractSnapshotDataUrl(zip: JSZip, markupPath: string) {
  const folder = getFolderFromPath(markupPath);
  const files = Object.keys(zip.files);

  const imagePath = files.find((path) => {
    const lower = path.toLowerCase();
    if (zip.files[path].dir) return false;
    if (!lower.endsWith(".png") && !lower.endsWith(".jpg") && !lower.endsWith(".jpeg") && !lower.endsWith(".webp")) return false;
    if (folder && !path.startsWith(`${folder}/`)) return false;
    return lower.includes("snapshot") || lower.endsWith(".png") || lower.endsWith(".jpg") || lower.endsWith(".jpeg") || lower.endsWith(".webp");
  });

  if (!imagePath) {
    return {
      snapshotPath: "",
      snapshotDataUrl: "",
    };
  }

  const base64 = await zip.files[imagePath].async("base64");

  return {
    snapshotPath: imagePath,
    snapshotDataUrl: `data:${getMimeTypeFromPath(imagePath)};base64,${base64}`,
  };
}

function normalizeBcfTopicStatus(status = "") {
  const s = String(status || "").trim();

  if (/^open$/i.test(s)) return "New";
  if (/^closed$/i.test(s)) return "Closed";
  if (/^active$/i.test(s)) return "New";
  if (/^resolved$/i.test(s)) return "Closed";

  return s;
}

async function readBcfZip(fileName: string, buffer: Buffer) {
  const bcfComments: any[] = [];
  const bcfTopics: any[] = [];

  const zip = await JSZip.loadAsync(buffer);
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "", textNodeName: "#text", trimValues: true });
  const isSolibriChecking = isSolibriCheckingFile(fileName);
  const solibriCheckingRevision = extractSolibriCheckingRevision(fileName);
  const origineFile = isSolibriChecking ? "Solibri" : "BCF";
  const tipoVerificaFile = isSolibriChecking ? "Checking Modelli" : "Verifica Documentale";

  const markupPaths = Object.keys(zip.files).filter((path) => path.toLowerCase().endsWith("markup.bcf"));

  for (const path of markupPaths) {
    const xml = await zip.files[path].async("text");
    const parsed = parser.parse(xml);
    const markup = extractMarkup(parsed);

    const folderGuid = path.split("/")[0] || "";
    const topicData = extractTopic(markup, folderGuid);
    const snapshot = await extractSnapshotDataUrl(zip, path);

    bcfTopics.push({
      sourceFile: fileName,
      origine: origineFile,
      tipoVerifica: tipoVerificaFile,
      revisioneChecking: solibriCheckingRevision,
      isSolibriChecking,
      markupPath: path,
      snapshotPath: snapshot.snapshotPath,
      snapshotDataUrl: snapshot.snapshotDataUrl,
      Label: topicData.topicGuid,
      ID: topicData.topicGuid,
      Guid: topicData.topicGuid,
      Title: topicData.topicTitle,
      Description: topicData.topicDescription,
      Tags: topicData.topicLabels,
      Status: normalizeBcfTopicStatus(topicData.topicStatus),
      Priority: topicData.topicPriority,
      "Created by": topicData.topicCreationAuthor,
      "Created on": topicData.topicCreationDate,
      "Last modified by": topicData.topicModifiedAuthor,
      "Last modified on": topicData.topicModifiedDate,
      "Assignee(s)": isSolibriChecking ? "BIM" : topicData.topicAssignedTo,
      Groups: isSolibriChecking ? "BIM" : "",
      disciplina: isSolibriChecking ? "BIM" : "",
      Type: "BCF Topic",
      __source: fileName.toLowerCase().endsWith(".bcf") ? "bcf" : "bcfzip",
    });

    const comments = extractComments(markup);

    for (const c of comments) {
      const text = getXmlText(getAny(c, ["Comment", "comment", "Text", "text"]));
      const role = roleFromText(text);

      bcfComments.push({
        sourceFile: fileName,
        markupPath: path,
        topicGuid: topicData.topicGuid,
        topicTitle: topicData.topicTitle,
        topicKey: normalize(topicData.topicTitle),
        author: getXmlText(getAny(c, ["Author", "author", "ModifiedAuthor"])) || "",
        date: getXmlText(getAny(c, ["Date", "date", "ModifiedDate"])) || "",
        role,
        comment: text,
      });
    }
  }

  return { bcfTopics, bcfComments, markupCount: markupPaths.length };
}

function extractDocxTables(documentXml: string) {
  const tables: string[][][] = [];
  const tableRegex = /<w:tbl[\s\S]*?<\/w:tbl>/g;
  const rowRegex = /<w:tr[\s\S]*?<\/w:tr>/g;
  const cellRegex = /<w:tc[\s\S]*?<\/w:tc>/g;

  const tableBlocks = documentXml.match(tableRegex) || [];

  for (const tableBlock of tableBlocks) {
    const rows: string[][] = [];
    const rowBlocks = tableBlock.match(rowRegex) || [];

    for (const rowBlock of rowBlocks) {
      const cells: string[] = [];
      const cellBlocks = rowBlock.match(cellRegex) || [];

      for (const cellBlock of cellBlocks) {
        const withBreaks = cellBlock
          .replace(/<w:br\s*\/>/g, "\n")
          .replace(/<w:tab\s*\/>/g, " ")
          .replace(/<\/w:p>/g, "\n");

        const texts = [...withBreaks.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => xmlDecode(m[1] || ""));
        cells.push(cleanText(texts.join("")));
      }

      if (cells.some((c) => cleanText(c))) rows.push(cells);
    }

    if (rows.length > 0) tables.push(rows);
  }

  return tables;
}

function extractTextFromDocxTables(tables: string[][][]) {
  return tables.flat(2).map((x) => cleanText(x)).filter(Boolean).join("\n");
}

function detectDisciplineFromDocx(fileName: string, tables: string[][][]) {
  for (const table of tables) {
    for (let i = 0; i < table.length; i++) {
      const row = table[i];
      const idx = row.findIndex((cell) => normalize(cell) === "oggetto");
      if (idx >= 0) {
        const sameRowValue = cleanText(row[idx + 1] || "");
        if (sameRowValue && !/nota ricezione/i.test(sameRowValue)) return sameRowValue;
        const nextRowValue = cleanText(table[i + 1]?.[idx] || table[i + 1]?.[idx + 1] || "");
        if (nextRowValue) return nextRowValue;
      }
    }
  }

  const base = fileName.replace(/\.docx$/i, "").split("_").pop() || "";
  return cleanText(base.replace(/[-_]+/g, " ")) || "Non assegnata";
}

function detectWordIssueTipo(idCell: string, riscontro = "") {
  const id = normalize(idCell).toUpperCase();
  const r = normalize(riscontro);

  if (id.includes("NESSUN RILIEVO")) return "Nessun rilievo";
  if (r.includes("declassa ad osservazione") || r.includes("declassata ad osservazione") || r.includes("declassato ad osservazione")) return "Da NC a OSS";
  if (/\bOSS\s*\d+/i.test(idCell)) return "OSS";
  if (/\bNC\s*\d+/i.test(idCell)) return "NC";

  return "Esito mancante";
}

function extractTrFromId(idCell: string) {
  const match = String(idCell || "").match(/\(\s*TR\s*[-_]?\s*0*(\d+[A-Z]?)\s*\)/i);
  return match ? `TR-${match[1]}` : "";
}

function normalizeIssueId(idCell: string) {
  const first = cleanText(idCell).split("\n").map((x) => x.trim()).filter(Boolean)[0] || "";
  const tr = extractTrFromId(idCell);
  return [first, tr ? `(${tr})` : ""].filter(Boolean).join(" ");
}

function rowLooksLikeIssue(row: string[]) {
  const first = cleanText(row[0] || "");
  return /\b(NC|OSS)\s*\d+/i.test(first) || /\(\s*TR\s*[-_]?\s*\d+/i.test(first);
}

function findRilieviTable(tables: string[][][]) {
  return tables.find((table) => {
    const text = normalize(table.slice(0, 6).flat().join(" "));
    return text.includes("codice elaborato") && text.includes("rilievi odi") && text.includes("risposta del progettista") && text.includes("stato");
  });
}

function findRiepilogoTable(tables: string[][][]) {
  return tables.find((table) => {
    const text = normalize(table.slice(0, 6).flat().join(" "));
    return text.includes("codice elaborato") && text.includes("revisione") && text.includes("presenza di nc") && text.includes("assenza nc oss");
  });
}


function getElaboratoUnivocoKey(row: any) {
  const value =
    row?.codiceElaborato ||
    row?.elaborato ||
    row?.titolo ||
    row?.Title ||
    "";

  const normalizedText = normalize(value);
  if (!normalizedText) return "";
  if (normalizedText.includes("codice elaborato")) return "";

  return normalizeElaboratoCode(value);
}

async function readDocxInspection(fileName: string, buffer: Buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.files["word/document.xml"]?.async("text");

  if (!documentXml) {
    return { rows: [], importedFile: { fileName, type: "docx", rows: 0, error: "word/document.xml non trovato" } };
  }

  const tables = extractDocxTables(documentXml);
  const allText = extractTextFromDocxTables(tables);
  const disciplina = detectDisciplineFromDocx(fileName, tables);
  const rilieviTable = findRilieviTable(tables);
  const riepilogoTable = findRiepilogoTable(tables);
  const rows: any[] = [];
  const issueCodes = new Set<string>();

  if (rilieviTable) {
    for (const row of rilieviTable) {
      if (!rowLooksLikeIssue(row)) continue;

      const padded = [...row];
      while (padded.length < 8) padded.push("");

      const idCell = cleanText(padded[0]);
      const codiceElaborato = cleanText(padded[1]);
      const titoloElaborato = cleanText(padded[2]);
      const rilievoOdi = cleanText(padded[3]);
      const ispettore = cleanText(padded[4]);
      const rispostaProgettista = cleanText(padded[5]);
      const riscontroIts = cleanText(padded[6]);
      const stato = translateStatus(cleanText(padded[7]));
      const tipo = detectWordIssueTipo(idCell, riscontroIts);
      const id = normalizeIssueId(idCell);
      const tr = extractTrFromId(idCell);
      const comments = uniqueComments([
        rispostaProgettista
          ? { sourceFile: fileName, role: "PRG", author: "Progettista", date: "", comment: rispostaProgettista }
          : null,
        riscontroIts
          ? { sourceFile: fileName, role: "ISP", author: ispettore || "ITS", date: "", comment: riscontroIts }
          : null,
      ].filter(Boolean) as any[]);

      issueCodes.add(normalize(codiceElaborato));

      const tipologiaNcOss = detectTipologiaNcOss(tipo, rilievoOdi, titoloElaborato, tipo);

      const hasPrgComment = comments.some((c) => c.role === "PRG");
      const hasIspComment = comments.some((c) => c.role === "ISP");
      const last = comments[comments.length - 1];
      const ultimoRuolo = last?.role || "";
      const isRilievo = tipo === "NC" || tipo === "OSS" || tipo === "Da NC a OSS";

      let chiDeveAgire = "";
      let statoRisoluzione = "Non applicabile";

      if (isRilievo) {
        if (stato === "Chiusa") {
          chiDeveAgire = "";
          statoRisoluzione = "Chiusa";
        } else if (!hasPrgComment) {
          chiDeveAgire = "PRG";
          statoRisoluzione = "In attesa riscontro progettista";
        } else if (ultimoRuolo === "PRG") {
          chiDeveAgire = "ISP";
          statoRisoluzione = "Risposto da progettista - da verificare ISP";
        } else if (ultimoRuolo === "ISP") {
          chiDeveAgire = "PRG";
          statoRisoluzione = "Riscontrato da ispettore - eventuale azione PRG";
        } else {
          chiDeveAgire = "PRG";
          statoRisoluzione = "Da riscontrare";
        }
      }

      rows.push({
        idRecord: `DOCX-${fileName}-${rows.length + 1}`,
        id,
        idTodo: id,
        tr,
        elaborato: codiceElaborato,
        codiceElaborato,
        titolo: codiceElaborato,
        titoloElaborato,
        descrizione: rilievoOdi,
        tipo,
        tipoOriginale: "Scheda Word ITS",
        tags: tipo,
        tipologiaNcOss,
        tipologiaDocumento: tipologiaNcOss,
        tipologia: tipologiaNcOss,
        disciplina: disciplina || "Non assegnata",
        gruppoTrimble: disciplina || "",
        assegnatari: disciplina || "",
        stato,
        statoOriginale: stato,
        priorita: "",
        completamento: "",
        scadenza: "",
        ispettore,
        creatoDa: ispettore,
        creatoIl: "",
        modificatoDa: "",
        modificatoIl: "",
        controlloIspettoreCompleto: Boolean(id && codiceElaborato && tipo && disciplina),
        campiMancantiControllo: [],
        statoCompilazioneIspettore: "Completo",
        hasPrgComment,
        hasIspComment,
        numeroCommentiPrg: comments.filter((c) => c.role === "PRG").length,
        numeroCommentiIsp: comments.filter((c) => c.role === "ISP").length,
        ultimoRuolo,
        ultimoCommento: last?.comment || "",
        ultimoAutore: last?.author || "",
        ultimaDataCommento: last?.date || "",
        chiDeveAgire,
        statoRisoluzione,
        nCommenti: comments.length,
        comments,
        sourceFile: fileName,
        sourceType: "docx",
      });
    }
  }

  if (riepilogoTable) {
    for (const row of riepilogoTable) {
      const codiceElaborato = cleanText(row[0] || "");
      const revisione = cleanText(row[1] || "");
      const titoloElaborato = cleanText(row[2] || "");
      const presenzaNc = cleanText(row[3] || "");
      const presenzaOss = cleanText(row[4] || "");
      const assenza = cleanText(row[5] || "");

      if (!codiceElaborato || normalize(codiceElaborato).includes("codice elaborato")) continue;
      if (issueCodes.has(normalize(codiceElaborato))) continue;
      if (!assenza && (presenzaNc || presenzaOss)) continue;

      rows.push({
        idRecord: `DOCX-${fileName}-ELAB-${rows.length + 1}`,
        id: `NESSUN RILIEVO - ${codiceElaborato}`,
        idTodo: `NESSUN RILIEVO - ${codiceElaborato}`,
        elaborato: codiceElaborato,
        codiceElaborato,
        titolo: codiceElaborato,
        titoloElaborato,
        descrizione: titoloElaborato || "Nessun rilievo",
        tipo: "Nessun rilievo",
        tipoOriginale: "Scheda Word ITS - riepilogo elaborati",
        tags: "Nessun rilievo",
        tipologiaNcOss: "",
        tipologiaDocumento: "",
        tipologia: "",
        disciplina: disciplina || "Non assegnata",
        gruppoTrimble: disciplina || "",
        assegnatari: disciplina || "",
        stato: "Chiusa",
        statoOriginale: "Chiusa",
        revisione,
        priorita: "",
        completamento: "",
        scadenza: "",
        ispettore: "",
        creatoDa: "",
        creatoIl: "",
        modificatoDa: "",
        modificatoIl: "",
        controlloIspettoreCompleto: true,
        campiMancantiControllo: [],
        statoCompilazioneIspettore: "Completo",
        hasPrgComment: false,
        hasIspComment: false,
        numeroCommentiPrg: 0,
        numeroCommentiIsp: 0,
        ultimoRuolo: "",
        ultimoCommento: "",
        ultimoAutore: "",
        ultimaDataCommento: "",
        chiDeveAgire: "",
        statoRisoluzione: "Non applicabile",
        nCommenti: 0,
        comments: [],
        sourceFile: fileName,
        sourceType: "docx",
      });
    }
  }

  return {
    rows,
    importedFile: {
      fileName,
      type: "docx",
      rows: rows.length,
      rilievi: rows.filter((r) => r.tipo === "NC" || r.tipo === "OSS" || r.tipo === "Da NC a OSS").length,
      elaboratiSenzaRilievi: rows.filter((r) => r.tipo === "Nessun rilievo").length,
      disciplina,
      tables: tables.length,
      textLength: allText.length,
    },
  };
}

export async function POST(req: Request) {
  try {
    const formData = await req.formData();
    const files = formData.getAll("files") as File[];

    let excelRows: any[] = [];
    const bcfTopicRows: any[] = [];
    const bcfComments: any[] = [];
    const docxRows: any[] = [];
    const importedFiles: any[] = [];

    for (const file of files) {
      const buffer = Buffer.from(await file.arrayBuffer());
      const name = file.name.toLowerCase();

      if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
        const workbook = XLSX.read(buffer, { type: "buffer" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        excelRows = [...excelRows, ...rows];

        importedFiles.push({ fileName: file.name, type: "xlsx", rows: rows.length });
        continue;
      }

      if (name.endsWith(".docx")) {
        const result = await readDocxInspection(file.name, buffer);
        docxRows.push(...result.rows);
        importedFiles.push(result.importedFile);
        continue;
      }

      if (name.endsWith(".bcf") || name.endsWith(".bcfzip") || name.endsWith(".zip")) {
        const zip = await JSZip.loadAsync(buffer);
        const docxEntries = Object.values(zip.files).filter((entry) => !entry.dir && entry.name.toLowerCase().endsWith(".docx"));

        for (const entry of docxEntries) {
          const docxBuffer = await entry.async("nodebuffer");
          const result = await readDocxInspection(entry.name.split("/").pop() || entry.name, docxBuffer);
          docxRows.push(...result.rows);
          importedFiles.push({ ...result.importedFile, fileName: `${file.name}/${entry.name}` });
        }

        const result = await readBcfZip(file.name, buffer);
        bcfTopicRows.push(...result.bcfTopics);
        bcfComments.push(...result.bcfComments);

        importedFiles.push({
          fileName: file.name,
          origine: isSolibriCheckingFile(file.name) ? "Solibri" : "BCF",
          tipoVerifica: isSolibriCheckingFile(file.name) ? "Checking Modelli" : "Verifica Documentale",
          revisioneChecking: isSolibriCheckingFile(file.name) ? extractSolibriCheckingRevision(file.name) : "",
          type: name.endsWith(".bcf") ? "bcf" : name.endsWith(".bcfzip") ? "bcfzip" : "zip",
          markupCount: result.markupCount,
          topics: result.bcfTopics.length,
          comments: result.bcfComments.length,
          snapshots: result.bcfTopics.filter((topic) => topic.snapshotDataUrl).length,
          docx: docxEntries.length,
        });
      }
    }

    const uniqueBcfComments = uniqueComments(bcfComments);
    const commentsByTopic = new Map<string, any[]>();

    for (const c of uniqueBcfComments) {
      const keys = Array.from(new Set([normalize(c.topicTitle), normalize(c.topicGuid), normalize(String(c.topicTitle || "").replace(/\.pdf/gi, ""))].filter(Boolean)));

      for (const key of keys) {
        if (!commentsByTopic.has(key)) commentsByTopic.set(key, []);
        commentsByTopic.get(key)!.push(c);
      }
    }

    const topicsByKey = new Map<string, any>();

    for (const topic of bcfTopicRows) {
      const keys = Array.from(new Set([normalize(topic.Label), normalize(topic.ID), normalize(topic.Guid), normalize(topic.Title), normalize(String(topic.Title || "").replace(/\.pdf/gi, ""))].filter(Boolean)));

      for (const key of keys) {
        if (!topicsByKey.has(key)) topicsByKey.set(key, topic);
      }
    }

    const todoRows = buildTodoRowsWithStandaloneBcf(excelRows, bcfTopicRows);

    const rowsFromTodoOrBcf = todoRows.map((todo: any, index: number) => {
      const matchedBcfTopic = todo.__source
        ? todo
        : findBestBcfTopic(todo, topicsByKey, bcfTopicRows);
      const label = getTodoLabel(todo);
      const title = todo.Title || matchedBcfTopic?.Title || "";
      const todoDescription = todo.Description || "";
      const bcfDescription = matchedBcfTopic?.Description || "";
      const description = String(bcfDescription || "").trim() ? bcfDescription : todoDescription;
      const tags = todo.Tags || todo.Labels || matchedBcfTopic?.Tags || "";
      const tipo = detectTipo(tags, description);
      const tipologiaNcOss = detectTipologiaNcOss(tags, description, title, tipo);
      const isSolibriCheckingRow = Boolean(
        todo.isSolibriChecking ||
        matchedBcfTopic?.isSolibriChecking ||
        isSolibriCheckingFile(todo.sourceFile || matchedBcfTopic?.sourceFile || "")
      );

      const disciplina = isSolibriCheckingRow
        ? "BIM"
        : getTodoAssignees(todo) || matchedBcfTopic?.["Assignee(s)"] || "";
      const statoOriginale = matchedBcfTopic?.Status || todo.Status || "";
      const statoTradotto = translateStatus(statoOriginale);
      const ispettore = todo["Created by"] || todo["Last modified by"] || todo.Owner || "";
      const comments = uniqueComments(findBestComments(todo, commentsByTopic, matchedBcfTopic)).sort((a, b) => String(a.date).localeCompare(String(b.date)));
      const prgComments = comments.filter((c) => c.role === "PRG");
      const ispComments = comments.filter((c) => c.role === "ISP");
      const last = comments[comments.length - 1];
      const hasPrgComment = prgComments.length > 0;
      const hasIspComment = ispComments.length > 0;
      const ultimoRuolo = last?.role || "";

      let chiDeveAgire = "";
      let statoRisoluzione = "Non applicabile";

      if (tipo === "NC" || tipo === "OSS" || tipo === "Da NC a OSS") {
        if (!hasPrgComment) {
          chiDeveAgire = "PRG";
          statoRisoluzione = "In attesa riscontro progettista";
        } else if (ultimoRuolo === "PRG") {
          chiDeveAgire = "ISP";
          statoRisoluzione = "Risposto da progettista - da verificare ISP";
        } else if (ultimoRuolo === "ISP") {
          chiDeveAgire = "PRG";
          statoRisoluzione = "Riscontrato da ispettore - eventuale azione PRG";
        } else {
          chiDeveAgire = "PRG";
          statoRisoluzione = "Da riscontrare";
        }
      }

      const esitoCompilato = tipo === "NC" || tipo === "OSS" || tipo === "Nessun rilievo" || tipo === "Da NC a OSS";
      const titleCompilato = Boolean(String(title).trim());
      const disciplinaCompilata = Boolean(String(disciplina).trim());
      const campiMancantiControllo: string[] = [];

      if (!esitoCompilato) campiMancantiControllo.push("Tags / Esito verifica mancante o non coerente");
      if (!titleCompilato) campiMancantiControllo.push("Title / Codice elaborato mancante");
      if (!disciplinaCompilata) campiMancantiControllo.push("Gruppo / Disciplina mancante");

      const controlloIspettoreCompleto = esitoCompilato && titleCompilato && disciplinaCompilata;

      return {
        idRecord: `IMPORT-${index + 1}`,
        id: label,
        idTodo: label,
        elaborato: title,
        titolo: title,
        descrizione: description,
        tipo,
        tipoOriginale: todo.Type || "",
        origine: todo.origine || matchedBcfTopic?.origine || (todo.__source ? "BCF" : "Trimble"),
        tipoVerifica: todo.tipoVerifica || matchedBcfTopic?.tipoVerifica || "Verifica Documentale",
        revisioneChecking: todo.revisioneChecking || matchedBcfTopic?.revisioneChecking || "",
        isSolibriChecking: Boolean(todo.isSolibriChecking || matchedBcfTopic?.isSolibriChecking),
        snapshotPath: todo.snapshotPath || matchedBcfTopic?.snapshotPath || "",
        snapshotDataUrl: todo.snapshotDataUrl || matchedBcfTopic?.snapshotDataUrl || "",
        tags,
        tipologiaNcOss,
        tipologiaDocumento: tipologiaNcOss,
        tipologia: tipologiaNcOss,
        disciplina: disciplina || "Non assegnata",
        gruppoTrimble: disciplina || "",
        assegnatari: disciplina || "",
        stato: statoTradotto,
        statoOriginale,
        priorita: todo.Priority || "",
        completamento: todo.Completion || "",
        scadenza: todo["Due date"] || "",
        ispettore,
        creatoDa: todo["Created by"] || "",
        creatoIl: todo["Created on"] || "",
        modificatoDa: todo["Last modified by"] || "",
        modificatoIl: todo["Last modified on"] || "",
        controlloIspettoreCompleto,
        campiMancantiControllo,
        statoCompilazioneIspettore: controlloIspettoreCompleto ? "Completo" : "Incompleto",
        hasPrgComment,
        hasIspComment,
        numeroCommentiPrg: prgComments.length,
        numeroCommentiIsp: ispComments.length,
        ultimoRuolo,
        ultimoCommento: last?.comment || "",
        ultimoAutore: last?.author || "",
        ultimaDataCommento: last?.date || "",
        chiDeveAgire,
        statoRisoluzione,
        nCommenti: comments.length,
        comments,
        sourceFile: todo.sourceFile || matchedBcfTopic?.sourceFile || "",
        sourceType: todo.__source || matchedBcfTopic?.__source || (excelRows.length > 0 ? "xlsx" : "bcf"),
      };
    });

    const rows = [...rowsFromTodoOrBcf, ...docxRows];

    const elaboratiEsaminatiKeys = new Set<string>();

    for (const row of rows) {
      const key = getElaboratoUnivocoKey(row);
      if (key) elaboratiEsaminatiKeys.add(key);
    }

    const elaboratiEsaminatiCount = elaboratiEsaminatiKeys.size;

    return NextResponse.json({
      ok: true,
      importedFiles,
      excelRowCount: excelRows.length,
      bcfTopicCount: bcfTopicRows.length,
      docxRowCount: docxRows.length,
      todoCount: elaboratiEsaminatiCount,
      todoRawCount: todoRows.length,
      elaboratiEsaminatiCount,
      elaboratiEsaminatiUnivoci: elaboratiEsaminatiCount,
      bcfCommentCount: uniqueBcfComments.length,
      rows,
      comments: uniqueBcfComments,
    });
  } catch (error: any) {
    return NextResponse.json({ ok: false, error: error?.message || "Errore durante la lettura dei file" }, { status: 500 });
  }
}
