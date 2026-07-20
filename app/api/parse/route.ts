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

function buildTopicMatchKey(title: any, description: any) {
  const titleKey = normalize(title);
  const descriptionKey = normalize(cleanText(description)).slice(0, 300);

  return [titleKey, descriptionKey].filter(Boolean).join("__");
}

function getCreatedBy(row: any) {
  return cleanText(getAny(row, [
    "Created by",
    "Created By",
    "CreationAuthor",
    "Creation Author",
    "Creation author",
    "Creator",
    "Owner",
  ]));
}

function getCreatedOn(row: any) {
  return cleanText(getAny(row, [
    "Created on",
    "Created On",
    "CreationDate",
    "Creation Date",
    "Creation date",
  ]));
}

function getModifiedBy(row: any) {
  return cleanText(getAny(row, [
    "Last modified by",
    "Last Modified By",
    "ModifiedAuthor",
    "Modified Author",
    "ModificationAuthor",
    "Modification Author",
  ]));
}

function getModifiedOn(row: any) {
  return cleanText(getAny(row, [
    "Last modified on",
    "Last Modified On",
    "ModifiedDate",
    "Modified Date",
    "ModificationDate",
    "Modification Date",
  ]));
}

function normalizeElaboratoCode(v = "") {
  return String(v || "")
    .toUpperCase()
    .replace(/\.PDF$/i, "")
    .replace(/[^A-Z0-9]/g, "");
}


function normalizeElaboratoForTrimble(value: any) {
  const raw = cleanText(value);
  const normalized = normalize(raw);

  const isRilievoGenerale = normalized === "rilievo generale";

  const parts = raw
    .split(";")
    .map((part) => part.trim())
    .filter(Boolean);

  const isValidCodeOrCodes =
    parts.length > 0 &&
    parts.every((part) => /^[A-Z0-9]+(?:_[A-Z0-9]+){2,}/i.test(part));

  if (isRilievoGenerale || isValidCodeOrCodes) {
    return {
      elaborato: raw,
      anomaliaElaborato: "",
    };
  }

  return {
    elaborato: "Rilievo Generale",
    anomaliaElaborato: raw
      ? `Elaborato non ammesso corretto in Rilievo Generale: ${raw}`
      : "Elaborato mancante corretto in Rilievo Generale",
  };
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

const BCF_PARSER_VERSION = "2026-07-20_v23_its_domain_roles";


const ISPETTORI_DISCIPLINE_ITS: Record<string, string> = {
  "Arch. Veronica Laino": "Progetto Architettonico",
  "Arch. Arianna Brunetti": "Progetto Architettonico",
  "Ing. Salvatore Grimaldi": "Progetto Strutturale",
  "Ing. Bruno Gabrielli": "Geotecnica",
  "Ing. Carlo Renda": "Impianti meccanici e relativa documentazione economica",
  "Ing. Gianluca Biaggioli": "Impianti elettrici e relativa documentazione economica",
  "Ing. Marta Dominijanni": "BIM",
  "P.I. Mauro Garofalo": "Documentazione economica opere civili",
  "Arch. Riccardo Hoops": "CAM e DNSH",
  "Ing. Marcello Caccialupi": "Acustica",
  "Geom. Massimo Tamberi": "Sicurezza e Cantierizzazione",
  "Arch. Stefano Arcangellelli": "Progetto Architettonico",
  "Ing. Edoardo Oddo Casano": "Progetto Strutturale",
};

const INSPECTOR_DOMAINS = [
  "itscontrollitecnici.it",
];

function hasInspectorDomain(value = "") {
  const raw = cleanText(value).toLowerCase();
  if (!raw) return false;

  const emailMatches = raw.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/gi) || [];

  return emailMatches.some((email) => {
    const domain = email.split("@")[1] || "";
    return INSPECTOR_DOMAINS.some(
      (allowedDomain) => domain === allowedDomain || domain.endsWith(`.${allowedDomain}`)
    );
  });
}

const ISPETTORI_ITS = [
  "Arch. Veronica Laino",
  "Arch. Arianna Brunetti",
  "Ing. Salvatore Grimaldi",
  "Ing. Bruno Gabrielli",
  "Ing. Carlo Renda",
  "Ing. Gianluca Biaggioli",
  "Ing. Marta Dominijanni",
  "P.I. Mauro Garofalo",
  "Arch. Riccardo Hoops",
  "Ing. Marcello Caccialupi",
  "Geom. Massimo Tamberi",
  "Arch. Stefano Arcangellelli",
  "Ing. Edoardo Oddo Casano",
];

const AUTHOR_NAME_MAP: Record<string, string> = {
  "m.dominijanni@itscontrollitecnici.it": "Ing. Marta Dominijanni",
  "marta.dominijanni@itscontrollitecnici.it": "Ing. Marta Dominijanni",
  "v.laino@itscontrollitecnici.it": "Arch. Veronica Laino",
  "veronica.laino@itscontrollitecnici.it": "Arch. Veronica Laino",
  "a.brunetti@itscontrollitecnici.it": "Arch. Arianna Brunetti",
  "s.grimaldi@itscontrollitecnici.it": "Ing. Salvatore Grimaldi",
  "b.gabrielli@itscontrollitecnici.it": "Ing. Bruno Gabrielli",
  "c.renda@itscontrollitecnici.it": "Ing. Carlo Renda",
  "g.biaggioli@itscontrollitecnici.it": "Ing. Gianluca Biaggioli",
  "m.garofalo@itscontrollitecnici.it": "P.I. Mauro Garofalo",
  "r.hoops@itscontrollitecnici.it": "Arch. Riccardo Hoops",
  "m.caccialupi@itscontrollitecnici.it": "Ing. Marcello Caccialupi",
  "m.tamberi@itscontrollitecnici.it": "Geom. Massimo Tamberi",
  "s.arcangellelli@itscontrollitecnici.it": "Arch. Stefano Arcangellelli",
  "e.cassano@itscontrollitecnici.it": "Ing. Edoardo Oddo Casano",
};

function normalizeAuthorName(value = "") {
  const raw = cleanText(value);
  if (!raw) return "";

  const key = raw.toLowerCase().trim();
  if (AUTHOR_NAME_MAP[key]) return AUTHOR_NAME_MAP[key];

  const emailMatch = key.match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i);
  if (emailMatch && AUTHOR_NAME_MAP[emailMatch[0]]) {
    return AUTHOR_NAME_MAP[emailMatch[0]];
  }

  return raw;
}

function normalizePersonName(value = "") {
  return normalize(value)
    .replace(/\barch\b/g, "")
    .replace(/\bing\b/g, "")
    .replace(/\bgeom\b/g, "")
    .replace(/\bp\s*i\b/g, "")
    .replace(/\bdott\b/g, "")
    .replace(/\bdr\b/g, "")
    .replace(/\s+/g, " ")
    .trim();
}


function getIspettoreDisciplineFromCreatedBy(author = "") {
  const authorKey = normalizePersonName(author);
  if (!authorKey) return "";

  for (const [name, discipline] of Object.entries(ISPETTORI_DISCIPLINE_ITS)) {
    const key = normalizePersonName(name);
    if (key && (authorKey === key || authorKey.includes(key) || key.includes(authorKey))) {
      return discipline;
    }
  }

  return "";
}

function roleFromAuthor(author = "", _text = "") {
  // Regola ITS:
  // 1. qualunque account del dominio ITS e sempre classificato come ISP;
  // 2. in assenza dell'e-mail, restano validi i nominativi dell'anagrafica;
  // 3. tutti gli altri autori sono PRG.
  // Le sigle ISP/PRG eventualmente presenti nel testo del commento
  // non modificano la classificazione basata sull'autore.
  if (hasInspectorDomain(author)) return "ISP";

  const normalizedAuthor = normalizeAuthorName(author);
  const authorKey = normalizePersonName(normalizedAuthor);

  if (!authorKey || authorKey === "autore non indicato") return "PRG";

  const isIspettore = ISPETTORI_ITS.some((name) => {
    const key = normalizePersonName(name);
    return Boolean(
      key &&
      (authorKey === key || authorKey.includes(key) || key.includes(authorKey))
    );
  });

  return isIspettore ? "ISP" : "PRG";
}

function getCommentTextFromBcfComment(c: any) {
  if (typeof c === "string" || typeof c === "number") return cleanText(c);

  return cleanText(
    getXmlText(getAny(c, ["Comment", "comment", "CommentText", "commentText", "Text", "text", "Description", "description"])) ||
      getXmlText(c?.Comment) ||
      getXmlText(c?.comment) ||
      getXmlText(c?.CommentText) ||
      getXmlText(c?.commentText) ||
      getXmlText(c?.Text) ||
      getXmlText(c?.text) ||
      ""
  );
}

function getCommentAuthorFromBcfComment(c: any) {
  if (!c || typeof c !== "object") return "Autore non indicato";

  return cleanText(
    getXmlText(getAny(c, ["Author", "author", "ModifiedAuthor", "modifiedAuthor", "CreationAuthor", "creationAuthor"])) ||
      getXmlText(c?.Author) ||
      getXmlText(c?.author) ||
      "Autore non indicato"
  );
}

function getCommentDateFromBcfComment(c: any) {
  if (!c || typeof c !== "object") return "";

  return cleanText(
    getXmlText(getAny(c, ["Date", "date", "ModifiedDate", "modifiedDate", "CreationDate", "creationDate"])) ||
      getXmlText(c?.Date) ||
      getXmlText(c?.date) ||
      ""
  );
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

function normalizeIssueCommentKey(comment: any) {
  // Deduplica i commenti nella singola issue senza usare la data.
  // Serve a evitare duplicazioni quando lo stesso commento BCF viene agganciato sia dal topic sia dai fallback di match.
  return [
    normalize(comment?.role || ""),
    normalize(comment?.author || ""),
    normalize(comment?.comment || ""),
  ].join("|");
}

function uniqueIssueComments(comments: any[] = []) {
  const seen = new Set<string>();
  const result: any[] = [];

  for (const c of comments || []) {
    const key = normalizeIssueCommentKey(c);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    result.push(c);
  }

  return result;
}

function findBestComments(todo: any, commentsByTopic: Map<string, any[]>, matchedBcfTopic?: any) {
  const collected: any[] = [];
  const usedKeys = new Set<string>();

  function addFromKey(value: any, isAlreadyNormalized = false) {
    const key = isAlreadyNormalized ? String(value || "") : normalize(value || "");
    if (!key || usedKeys.has(key)) return;
    usedKeys.add(key);

    if (commentsByTopic.has(key)) {
      collected.push(...(commentsByTopic.get(key) || []));
    }
  }

  const todoTitle = todo?.Title || todo?.title || todo?.titolo || todo?.elaborato || "";
  const todoDescription = todo?.Description || todo?.description || todo?.descrizione || "";
  const bcfTitle = matchedBcfTopic?.Title || matchedBcfTopic?.title || "";
  const bcfDescription = matchedBcfTopic?.Description || matchedBcfTopic?.description || "";

  // Regola rigida Trimble:
  // i commenti vengono associati solo al topic BCF effettivamente individuato
  // tramite Title + Description oppure tramite GUID/ID/Label.
  // Non si usano più fallback per solo Title o similarità, perché più rilievi
  // possono riferirsi allo stesso elaborato ma avere descrizioni diverse.
  [
    matchedBcfTopic?.Guid,
    matchedBcfTopic?.GUID,
    matchedBcfTopic?.ID,
    matchedBcfTopic?.Label,
    todo?.Guid,
    todo?.GUID,
  ].forEach((value) => addFromKey(value));

  [
    buildTopicMatchKey(todoTitle, todoDescription),
    buildTopicMatchKey(bcfTitle, bcfDescription),
  ]
    .filter(Boolean)
    .forEach((value) => addFromKey(value, true));

  if (Array.isArray(matchedBcfTopic?.comments)) {
    collected.push(...matchedBcfTopic.comments);
  }

  return uniqueComments(collected);
}

function translateStatus(status = "") {
  const s = String(status || "").trim().toLowerCase();

  if (s === "new") return "Aperta";
  if (s === "waiting") return "In attesa";
  if (s === "done") return "Fatto";
  if (s === "closed") return "Chiusa";
  if (s === "unknown") return "Non definito";

  if (s === "aperto" || s === "aperta") return "Aperta";
  if (s === "in attesa") return "In attesa";
  if (s === "fatto" || s === "fatta") return "Fatto";
  if (s === "chiuso" || s === "chiusa") return "Chiusa";

  return String(status || "").trim();
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
  const titleDescriptionKey = buildTopicMatchKey(todo.Title, todo.Description);

  // 1. Match principale: stesso elaborato (Title) + stessa descrizione del rilievo.
  if (titleDescriptionKey && topicsByKey.has(titleDescriptionKey)) {
    return topicsByKey.get(titleDescriptionKey);
  }

  // 2. Match tramite GUID/ID/Label se disponibile.
  for (const key of [labelKey, idKey, guidKey].filter(Boolean)) {
    if (topicsByKey.has(key)) return topicsByKey.get(key);
  }

  // 3. Match per solo Title consentito solo se quel Title è univoco nel BCF.
  // Se lo stesso elaborato ha più rilievi, il solo Title non è affidabile.
  if (titleKey && topicsByKey.has(titleKey)) {
    const candidate = topicsByKey.get(titleKey);
    const sameTitleCount = allTopics.filter((topic) => normalize(topic.Title) === titleKey).length;
    if (sameTitleCount === 1) return candidate;
  }

  // 4. Ultima sicurezza: se il title è uguale e la descrizione è molto simile.
  // Non trasferisce commenti da altri rilievi con descrizione diversa.
  if (titleKey && descriptionKey) {
    const sameTitleTopics = allTopics.filter((topic) => normalize(topic.Title) === titleKey);
    let bestScore = 0;
    let bestTopic: any = null;

    for (const topic of sameTitleTopics) {
      const score = similarity(todo.Description, topic.Description);
      if (score > bestScore) {
        bestScore = score;
        bestTopic = topic;
      }
    }

    if (bestScore >= 0.92) return bestTopic;
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
  const titleCounts = new Map<string, number>();

  for (const topic of topics) {
    const titleKey = normalize(topic.Title);
    if (!titleKey) continue;
    titleCounts.set(titleKey, (titleCounts.get(titleKey) || 0) + 1);
  }

  for (const topic of topics) {
    const titleKey = normalize(topic.Title);
    const titleDescriptionKey = buildTopicMatchKey(topic.Title, topic.Description);

    const keys = Array.from(new Set([
      normalize(topic.Label),
      normalize(topic.ID),
      normalize(topic.Guid),
      titleDescriptionKey,
      titleKey && titleCounts.get(titleKey) === 1 ? titleKey : "",
    ].filter(Boolean)));

    for (const key of keys) {
      if (!map.has(key)) map.set(key, topic);
    }
  }

  return map;
}

function isKnownIspettore(author = "") {
  const authorKey = normalizePersonName(normalizeAuthorName(author));
  if (!authorKey) return false;

  return ISPETTORI_ITS.some((name) => {
    const key = normalizePersonName(name);
    return Boolean(key && (authorKey === key || authorKey.includes(key) || key.includes(authorKey)));
  });
}

function getStandaloneBcfInspectorAuthor(topic: any) {
  // I Topic provenienti da Solibri appartengono alla verifica BIM ITS
  // e il redattore responsabile è Marta Dominijanni.
  if (topic?.isSolibriChecking || isSolibriCheckingFile(topic?.sourceFile || "")) {
    return "Ing. Marta Dominijanni";
  }

  const topicAuthor = normalizeAuthorName(
    getCreatedBy(topic) ||
    cleanText(topic?.topicCreationAuthor || "") ||
    cleanText(topic?.CreationAuthor || "") ||
    cleanText(topic?.creationAuthor || "")
  );

  if (isKnownIspettore(topicAuthor)) return topicAuthor;

  const inspectorCommentAuthors = (Array.isArray(topic?.comments) ? topic.comments : [])
    .filter((comment: any) => comment?.role === "ISP" || isKnownIspettore(comment?.author || ""))
    .map((comment: any) => normalizeAuthorName(comment?.author || ""))
    .filter((author: string) => isKnownIspettore(author));

  if (inspectorCommentAuthors.length) {
    return inspectorCommentAuthors[inspectorCommentAuthors.length - 1];
  }

  return topicAuthor;
}

function normalizeStandaloneBcfTopic(topic: any) {
  const isSolibriTopic = Boolean(
    topic?.isSolibriChecking || isSolibriCheckingFile(topic?.sourceFile || "")
  );
  const inspectorAuthor = isSolibriTopic
    ? "Ing. Marta Dominijanni"
    : getStandaloneBcfInspectorAuthor(topic);
  const disciplinaDaIspettore = getIspettoreDisciplineFromCreatedBy(inspectorAuthor);
  const disciplinaStandalone = isSolibriTopic
    ? "BIM"
    : disciplinaDaIspettore || cleanText(topic?.disciplina || topic?.["Assignee(s)"] || "");

  return {
    ...topic,
    __standaloneBcf: true,
    disciplina: disciplinaStandalone,
    gruppoTrimble: disciplinaStandalone,
    assegnatari: disciplinaStandalone,
    "Assignee(s)": disciplinaStandalone,
    "Created by": inspectorAuthor,
  };
}

function buildTodoRowsWithStandaloneBcf(excelRows: any[], bcfTopicRows: any[]) {
  if (!excelRows.length) return bcfTopicRows.map((topic) => normalizeStandaloneBcfTopic(topic));

  const topicsByKeyForMatch = buildBcfTopicsByKey(bcfTopicRows);
  const matchedTopicKeys = new Set<string>();

  for (const todo of excelRows) {
    const matched = findBestBcfTopic(todo, topicsByKeyForMatch, bcfTopicRows);
    if (matched) matchedTopicKeys.add(buildBcfTopicUniqueKey(matched));
  }

  const standaloneBcfTopics = bcfTopicRows
    .filter((topic) => {
      const key = buildBcfTopicUniqueKey(topic);
      return key && !matchedTopicKeys.has(key);
    })
    .map((topic) => normalizeStandaloneBcfTopic(topic));

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
  const commentsContainer = getAny(markup, ["Comments", "comments"]);

  if (commentsContainer) {
    if (Array.isArray(commentsContainer)) return commentsContainer;

    const nested = getAny(commentsContainer, ["Comment", "comment"]);
    if (nested) return arr(nested);

    // Alcuni export salvano direttamente l'oggetto commento dentro Comments.
    if (typeof commentsContainer === "object") return [commentsContainer];
  }

  const direct = getAny(markup, ["Comment", "comment"]);
  if (!direct) return [];

  // Caso Trimble reale:
  // <Comment Guid="..."><Date>...</Date><Author>...</Author><Comment>testo</Comment></Comment>
  // Il nodo esterno è già l'oggetto completo del commento. Non bisogna restituire direct.Comment,
  // altrimenti si perde Author e Date e resta solo una stringa.
  if (Array.isArray(direct)) return direct;
  if (typeof direct === "object") return [direct];

  return arr(direct);
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

  // Nel workflow ITS i file con estensione .bcf sono esportazioni Solibri.
  // I pacchetti Trimble vengono invece normalmente forniti come .bcfzip.
  return (
    /\.bcf$/i.test(baseName) ||
    /^solibri[_\s-]/i.test(baseName) ||
    /solibri/i.test(baseName)
  );
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

    const comments = extractComments(markup);
    const topicComments = comments
      .map((c: any) => {
        const comment = getCommentTextFromBcfComment(c);
        const author = getCommentAuthorFromBcfComment(c);
        const date = getCommentDateFromBcfComment(c);
        const role = roleFromAuthor(author, comment);

        return {
          sourceFile: fileName,
          markupPath: path,
          topicGuid: topicData.topicGuid,
          topicTitle: topicData.topicTitle,
          topicDescription: topicData.topicDescription,
          topicKey: normalize(topicData.topicTitle),
          topicMatchKey: buildTopicMatchKey(topicData.topicTitle, topicData.topicDescription),
          author,
          date,
          role,
          comment,
        };
      })
      .filter((c: any) => c.comment);

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
      topicCreationAuthor: normalizeAuthorName(topicData.topicCreationAuthor),
      topicCreationDate: topicData.topicCreationDate,
      "Created by": normalizeAuthorName(topicData.topicCreationAuthor),
      "Created on": topicData.topicCreationDate,
      "Last modified by": topicData.topicModifiedAuthor,
      "Last modified on": topicData.topicModifiedDate,
      "Assignee(s)": isSolibriChecking ? "BIM" : topicData.topicAssignedTo,
      Groups: isSolibriChecking ? "BIM" : "",
      disciplina: isSolibriChecking ? "BIM" : "",
      Type: "BCF Topic",
      __source: fileName.toLowerCase().endsWith(".bcf") ? "bcf" : "bcfzip",
      comments: topicComments,
    });

    bcfComments.push(...topicComments);
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
      const elaboratoNormalizzato = normalizeElaboratoForTrimble(codiceElaborato);

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
        elaborato: elaboratoNormalizzato.elaborato,
        codiceElaborato: elaboratoNormalizzato.elaborato,
        titolo: elaboratoNormalizzato.elaborato,
        elaboratoOriginale: codiceElaborato,
        anomaliaElaborato: elaboratoNormalizzato.anomaliaElaborato,
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

      const elaboratoNormalizzato = normalizeElaboratoForTrimble(codiceElaborato);

      rows.push({
        idRecord: `DOCX-${fileName}-ELAB-${rows.length + 1}`,
        id: `NESSUN RILIEVO - ${codiceElaborato}`,
        idTodo: `NESSUN RILIEVO - ${codiceElaborato}`,
        elaborato: elaboratoNormalizzato.elaborato,
        codiceElaborato: elaboratoNormalizzato.elaborato,
        titolo: elaboratoNormalizzato.elaborato,
        elaboratoOriginale: codiceElaborato,
        anomaliaElaborato: elaboratoNormalizzato.anomaliaElaborato,
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


function isTrimbleChronologyId(value = "") {
  return /^IT\d{2}-\d+$/i.test(cleanText(value));
}

function hasMeaningfulDiscipline(value = "") {
  const normalized = normalize(value);
  return Boolean(
    normalized &&
    normalized !== "null" &&
    normalized !== "null null" &&
    normalized !== "non assegnata"
  );
}

function buildFinalIssueMergeKey(row: any) {
  const elaborato = normalize(
    row?.elaboratoOriginale ||
    row?.elaborato ||
    row?.titolo ||
    ""
  );
  const descrizione = normalize(row?.descrizione || "");
  const tipo = normalize(row?.tipo || "");

  if (!elaborato || !descrizione) return "";
  return [elaborato, descrizione, tipo].join("__");
}

function mergeRowsPreservingTrimbleChronology(rows: any[]) {
  const grouped = new Map<string, any[]>();
  const ungrouped: any[] = [];

  for (const row of rows) {
    const key = buildFinalIssueMergeKey(row);

    if (!key) {
      ungrouped.push(row);
      continue;
    }

    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key)!.push(row);
  }

  function mergeCommentsIntoRow(primary: any, additionalRows: any[]) {
    const mergedComments = uniqueIssueComments([
      ...(Array.isArray(primary?.comments) ? primary.comments : []),
      ...additionalRows.flatMap((row) =>
        Array.isArray(row?.comments) ? row.comments : []
      ),
    ]).sort((a, b) =>
      String(a?.date || "").localeCompare(String(b?.date || ""))
    );

    const prgComments = mergedComments.filter(
      (comment) => comment?.role === "PRG"
    );
    const ispComments = mergedComments.filter(
      (comment) => comment?.role === "ISP"
    );
    const last = mergedComments[mergedComments.length - 1];

    return {
      ...primary,
      snapshotPath:
        primary?.snapshotPath ||
        additionalRows.find((row) => row?.snapshotPath)?.snapshotPath ||
        "",
      snapshotDataUrl:
        primary?.snapshotDataUrl ||
        additionalRows.find((row) => row?.snapshotDataUrl)?.snapshotDataUrl ||
        "",
      comments: mergedComments,
      nCommenti: mergedComments.length,
      hasPrgComment: prgComments.length > 0,
      hasIspComment: ispComments.length > 0,
      numeroCommentiPrg: prgComments.length,
      numeroCommentiIsp: ispComments.length,
      ultimoRuolo: last?.role || "",
      ultimoCommento: last?.comment || "",
      ultimoAutore: last?.author || "",
      ultimaDataCommento: last?.date || "",
    };
  }

  function commentFingerprints(row: any) {
    return new Set(
      (Array.isArray(row?.comments) ? row.comments : [])
        .map((comment: any) => normalizeIssueCommentKey(comment))
        .filter(Boolean)
    );
  }

  const mergedRows: any[] = [];

  for (const group of grouped.values()) {
    const trimbleRows = group.filter((row) =>
      isTrimbleChronologyId(row?.idTodo || row?.id || "")
    );
    const topicRows = group.filter(
      (row) => !isTrimbleChronologyId(row?.idTodo || row?.id || "")
    );

    // Nessun cronologico Trimble: conserva i Topic autonomi.
    if (!trimbleRows.length) {
      mergedRows.push(...topicRows);
      continue;
    }

    // Un solo ToDo Trimble: integra tutti i Topic BCF corrispondenti.
    if (trimbleRows.length === 1) {
      mergedRows.push(mergeCommentsIntoRow(trimbleRows[0], topicRows));
      continue;
    }

    // Più ToDo Trimble possono avere lo stesso elaborato e la stessa descrizione.
    // Non devono mai essere accorpati tra loro, perché possono avere status diversi
    // (NEW, WAITING, DONE, CLOSED) e cronologici IT25 differenti.
    const topicRowsByTrimbleIndex = new Map<number, any[]>();

    for (const topicRow of topicRows) {
      const topicKeys = commentFingerprints(topicRow);
      let bestIndex = -1;
      let bestOverlap = 0;

      trimbleRows.forEach((trimbleRow, index) => {
        const trimbleKeys = commentFingerprints(trimbleRow);
        const overlap = [...topicKeys].filter((key) =>
          trimbleKeys.has(key)
        ).length;

        if (overlap > bestOverlap) {
          bestOverlap = overlap;
          bestIndex = index;
        }
      });

      // Il Topic viene integrato soltanto quando i commenti permettono
      // di individuare senza ambiguità il ToDo corrispondente.
      // In caso contrario non genera una riga GUID duplicata.
      if (bestIndex >= 0 && bestOverlap > 0) {
        if (!topicRowsByTrimbleIndex.has(bestIndex)) {
          topicRowsByTrimbleIndex.set(bestIndex, []);
        }
        topicRowsByTrimbleIndex.get(bestIndex)!.push(topicRow);
      }
    }

    trimbleRows.forEach((trimbleRow, index) => {
      mergedRows.push(
        mergeCommentsIntoRow(
          trimbleRow,
          topicRowsByTrimbleIndex.get(index) || []
        )
      );
    });
  }

  return [...mergedRows, ...ungrouped].filter((row) => {
    const id = cleanText(row?.idTodo || row?.id || "");
    const isGuidOnly = !isTrimbleChronologyId(id);
    const isNoIssue = normalize(row?.tipo || "") === "nessun rilievo";
    const noDiscipline = !hasMeaningfulDiscipline(row?.disciplina || "");
    const author = normalizePersonName(
      row?.ispettore ||
      row?.redattore ||
      row?.creatoDa ||
      ""
    );
    const isGiuseppePizzi =
      author === "giuseppe pizzi" ||
      author === "gp";

    return !(isGuidOnly && isNoIssue && noDiscipline && isGiuseppePizzi);
  });
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
      const keys = Array.from(new Set([
        normalize(c.topicGuid),
        c.topicMatchKey || buildTopicMatchKey(c.topicTitle, c.topicDescription),
      ].filter(Boolean)));

      for (const key of keys) {
        if (!commentsByTopic.has(key)) commentsByTopic.set(key, []);
        commentsByTopic.get(key)!.push(c);
      }
    }

    const topicsByKey = buildBcfTopicsByKey(bcfTopicRows);

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

      const isStandaloneBcfRow = Boolean(todo.__standaloneBcf);

      const createdBy = normalizeAuthorName(
        isStandaloneBcfRow
          ? (
              getCreatedBy(todo) ||
              cleanText(todo?.topicCreationAuthor || "") ||
              cleanText(todo?.CreationAuthor || "") ||
              cleanText(todo?.creationAuthor || "") ||
              cleanText(todo?.["Created by"] || "") ||
              cleanText(matchedBcfTopic?.topicCreationAuthor || "") ||
              cleanText(matchedBcfTopic?.CreationAuthor || "") ||
              cleanText(matchedBcfTopic?.creationAuthor || "") ||
              cleanText(matchedBcfTopic?.["Created by"] || "")
            )
          : getCreatedBy(todo)
      );

      const modifiedBy = isStandaloneBcfRow ? "" : getModifiedBy(todo);
      const createdOn = getCreatedOn(todo);
      const modifiedOn = isStandaloneBcfRow ? "" : getModifiedOn(todo);

      // Per i Topic BCF standalone il redattore/ispettore coincide
      // esclusivamente con l'autore del Topic BCF.
      const ispettore = isStandaloneBcfRow && isSolibriCheckingRow
        ? "Ing. Marta Dominijanni"
        : createdBy;

      const disciplinaDaAssignee = cleanText(getTodoAssignees(todo));
      const disciplinaDaCreatore = getIspettoreDisciplineFromCreatedBy(createdBy);

      const disciplina = isStandaloneBcfRow
        ? isSolibriCheckingRow
          ? "BIM"
          : cleanText(todo?.disciplina || todo?.["Assignee(s)"] || "") || disciplinaDaCreatore || ""
        : isSolibriCheckingRow
          ? "BIM"
          : disciplinaDaAssignee || disciplinaDaCreatore || "";
      const statoOriginale = todo.Status || matchedBcfTopic?.Status || "";
      const statoTradotto = translateStatus(statoOriginale);
      const topicDirectComments = Array.isArray(matchedBcfTopic?.comments) ? matchedBcfTopic.comments : [];
      const commentsSource = topicDirectComments.length
        ? topicDirectComments
        : findBestComments(todo, commentsByTopic, matchedBcfTopic);
      const comments = uniqueIssueComments(commentsSource).sort((a, b) => String(a.date).localeCompare(String(b.date)));
      const prgComments = comments.filter((c) => c.role === "PRG");
      const ispComments = comments.filter((c) => c.role === "ISP");
      const last = comments[comments.length - 1];
      const hasPrgComment = prgComments.length > 0;
      const hasIspComment = ispComments.length > 0;
      const ultimoRuolo = last?.role || "";
      const isRilievoApertoSenzaRiscontro =
        (tipo === "NC" || tipo === "OSS" || tipo === "Da NC a OSS") &&
        statoTradotto === "Aperta" &&
        prgComments.length === 0;

      let chiDeveAgire = "";
      let statoRisoluzione = "Non applicabile";

      if (tipo === "NC" || tipo === "OSS" || tipo === "Da NC a OSS") {
        if (statoTradotto === "Aperta") {
          chiDeveAgire = "PRG";
          statoRisoluzione = "Aperta - in attesa del progettista";
        } else if (statoTradotto === "In attesa") {
          chiDeveAgire = "ISP";
          statoRisoluzione = "Risposta del progettista - verifica ispettore";
        } else if (statoTradotto === "Chiusa") {
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

      const elaboratoNormalizzato = normalizeElaboratoForTrimble(title);
      const esitoCompilato = tipo === "NC" || tipo === "OSS" || tipo === "Nessun rilievo" || tipo === "Da NC a OSS";
      const titleCompilato = Boolean(String(title).trim());
      const disciplinaCompilata = Boolean(String(disciplina).trim());
      const campiMancantiControllo: string[] = [];

      if (!esitoCompilato) campiMancantiControllo.push("Tags / Esito verifica mancante o non coerente");
      if (!titleCompilato) campiMancantiControllo.push("Title / Codice elaborato mancante");
      if (elaboratoNormalizzato.anomaliaElaborato) campiMancantiControllo.push(elaboratoNormalizzato.anomaliaElaborato);
      if (!disciplinaCompilata) campiMancantiControllo.push("Gruppo / Disciplina mancante");

      const controlloIspettoreCompleto = esitoCompilato && titleCompilato && disciplinaCompilata;

      return {
        idRecord: `IMPORT-${index + 1}`,
        id: label,
        idTodo: label,
        elaborato: elaboratoNormalizzato.elaborato,
        titolo: elaboratoNormalizzato.elaborato,
        elaboratoOriginale: title,
        anomaliaElaborato: elaboratoNormalizzato.anomaliaElaborato,
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
        status: statoTradotto,
        Status: statoTradotto,
        statoOriginale: statoTradotto,
        statusOriginale: statoOriginale,
        priorita: todo.Priority || "",
        completamento: todo.Completion || "",
        scadenza: todo["Due date"] || "",
        ispettore,
        redattore: ispettore,
        creatoDa: createdBy,
        creatoIl: createdOn,
        modificatoDa: modifiedBy,
        modificatoIl: modifiedOn,
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
        isRilievoApertoSenzaRiscontro,
        gestioneRilievoAlert: isRilievoApertoSenzaRiscontro ? "Manca riscontro del progettista" : "",
        nCommenti: comments.length,
        comments,
        sourceFile: todo.sourceFile || matchedBcfTopic?.sourceFile || "",
        sourceType: todo.__source || matchedBcfTopic?.__source || (excelRows.length > 0 ? "xlsx" : "bcf"),
      };
    });

    const rows = [
      ...mergeRowsPreservingTrimbleChronology(rowsFromTodoOrBcf),
      ...docxRows,
    ];

    const elaboratiEsaminatiKeys = new Set<string>();

    for (const row of rows) {
      const key = getElaboratoUnivocoKey(row);
      if (key) elaboratiEsaminatiKeys.add(key);
    }

    const elaboratiEsaminatiCount = elaboratiEsaminatiKeys.size;

    return NextResponse.json({
      ok: true,
      bcfParserVersion: BCF_PARSER_VERSION,
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
