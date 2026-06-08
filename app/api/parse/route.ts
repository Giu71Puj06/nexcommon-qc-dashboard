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

function roleFromText(text = "") {
  const m = String(text).match(/\((ISP|PRG)\)/i);
  return m ? m[1].toUpperCase() : "";
}

function detectTipo(tags = "", description = "") {
  const t = String(tags).toUpperCase();
  const d = String(description).toUpperCase();

  if (t.includes("DA NC A OSS") || t.includes("DA NC A OS")) return "Da NC a OSS";

  if (t.includes("NESSUN RILIEVO") || d.includes("NESSUN RILIEVO")) {
    return "Nessun rilievo";
  }

  if (t.includes("OSS")) return "OSS";
  if (t.includes("NC")) return "NC";

  return "Esito mancante";
}

function includesAny(text = "", words: string[]) {
  const n = normalize(text);
  return words.some((w) => n.includes(normalize(w)));
}

function detectTipologiaNcOss(
  tags = "",
  description = "",
  title = "",
  tipo = ""
) {
  // "Nessun rilievo" ed esiti mancanti NON devono comparire nelle Tipologie NC/OSS.
  if (tipo === "Nessun rilievo" || tipo === "Esito mancante") return "";

  const text = `${tags} ${description} ${title}`;

  // 1. Normative
  // Prioritaria perché spesso genera NC.
  if (
    includesAny(text, [
      "normativa",
      "normative",
      "normativa vigente",
      "conformità",
      "conforme",
      "non conforme",
      "ntc",
      "eurocodice",
      "eurocodici",
      "codice appalti",
      "verifica normativa",
      "verifiche obbligatorie",
      "verifica obbligatoria",
      "prescrizione normativa",
      "classificazione opere",
      "autorizzazione",
      "autorizzativo",
      "vincolo",
      "vincoli",
      "prescrizioni",
      "prescrizione",
    ])
  ) {
    return "1. Normative";
  }

  // 2. Incoerenze tra elaborati
  if (
    includesAny(text, [
      "incoerenza",
      "incoerenze",
      "discordanza",
      "discordanze",
      "non coerente",
      "non coerenti",
      "differenza tra",
      "difformità tra",
      "non allineato",
      "non allineati",
      "disallineamento",
      "contraddizione",
      "contraddizioni",
      "relazione e tavola",
      "relazione tavola",
      "elaborati non coerenti",
      "tavole non coerenti",
    ])
  ) {
    return "2. Incoerenze tra elaborati";
  }

  // 3. Informazioni mancanti / incomplete
  if (
    includesAny(text, [
      "mancante",
      "mancanti",
      "manca",
      "non presente",
      "non presenti",
      "non indicato",
      "non indicati",
      "non riportato",
      "non riportati",
      "assente",
      "assenti",
      "incompleto",
      "incompleta",
      "incompleti",
      "incomplete",
      "omesso",
      "omessa",
      "non risulta",
      "necessario integrare",
      "integrare",
    ])
  ) {
    return "3. Informazioni mancanti / incomplete";
  }

  // 4. Richieste di chiarimento
  if (
    includesAny(text, [
      "chiarire",
      "si chiede",
      "si richiede",
      "richiesta di chiarimento",
      "chiarimento",
      "chiarimenti",
      "specificare",
      "precisare",
      "verificare",
      "si invita",
      "dettagliare",
      "approfondire",
      "motivare",
      "esplicitare",
    ])
  ) {
    return "4. Richieste di chiarimento";
  }

  // 5. Elaborati e relazioni
  if (
    includesAny(text, [
      "relazione",
      "relazioni",
      "documento",
      "documentazione",
      "elaborato",
      "elaborati",
      "relazione tecnica",
      "relazione specialistica",
      "rapporto",
      "allegato",
      "capitolo",
      "paragrafo",
      "pag",
      "pagina",
      "indice",
      "testo",
    ])
  ) {
    return "5. Elaborati e relazioni";
  }

  // 6. Errori dimensionali / quote
  if (
    includesAny(text, [
      "quota",
      "quote",
      "dimensione",
      "dimensioni",
      "dimensionale",
      "dimensionali",
      "misura",
      "misure",
      "altezza",
      "larghezza",
      "spessore",
      "diametro",
      "sezione",
      "scala",
      "geometria",
      "geometrico",
      "geometrici",
    ])
  ) {
    return "6. Errori dimensionali / quote";
  }

  // 7. Dettagli costruttivi insufficienti
  if (
    includesAny(text, [
      "dettaglio",
      "dettagli",
      "particolare",
      "particolari",
      "nodo",
      "nodi",
      "sezione costruttiva",
      "dettaglio costruttivo",
      "particolare costruttivo",
      "schema costruttivo",
      "dettagli esecutivi",
      "esecutivo",
      "esecutivi",
    ])
  ) {
    return "7. Dettagli costruttivi insufficienti";
  }

  // 8. Computi e quantità
  if (
    includesAny(text, [
      "computo",
      "computi",
      "quantità",
      "quantita",
      "voce",
      "voci",
      "elenco prezzi",
      "prezzario",
      "prezzari",
      "prezzo",
      "prezzi",
      "stima",
      "stime",
      "stima economica",
      "quadro economico",
      "importo",
      "contabilità",
      "contabilita",
      "misurazione",
    ])
  ) {
    return "8. Computi e quantità";
  }

  // 9. Costruttibilità / fattibilità
  if (
    includesAny(text, [
      "cantiere",
      "realizzazione",
      "realizzabile",
      "non realizzabile",
      "esecuzione",
      "esecutabilità",
      "esecutabilita",
      "costruttibilità",
      "costruttibilita",
      "fattibilità",
      "fattibilita",
      "posa",
      "montaggio",
      "lavorazione",
      "manutenzione",
      "manutenibilità",
      "manutenibilita",
      "accessibilità manutentiva",
      "accessibilita manutentiva",
    ])
  ) {
    return "9. Costruttibilità / fattibilità";
  }

  // 10. Interferenze / clash
  if (
    includesAny(text, [
      "interferenza",
      "interferenze",
      "clash",
      "sovrapposizione",
      "sovrapposizioni",
      "conflitto",
      "conflitti",
      "interferisce",
      "interferiscono",
      "collisione",
      "collisioni",
      "coordinamento interdisciplinare",
      "coordinamento bim",
      "bim",
    ])
  ) {
    return "10. Interferenze / clash";
  }

  // Solo vere NC/OSS che non rientrano in nessuna delle tipologie/categorie sopra.
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

function findBestComments(todo: any, commentsByTopic: Map<string, any[]>) {
  const candidates = [
    todo.Label,
    todo.ID,
    todo.Guid,
    todo.GUID,
    todo.Title,
    String(todo.Title || "").replace(/\.pdf/gi, ""),
  ].filter(Boolean);

  for (const c of candidates) {
    const key = normalize(c);
    if (commentsByTopic.has(key)) return commentsByTopic.get(key) || [];
  }

  let bestScore = 0;
  let bestComments: any[] = [];

  for (const [topic, comments] of commentsByTopic.entries()) {
    const score = Math.max(...candidates.map((c) => similarity(c, topic)));

    if (score > bestScore) {
      bestScore = score;
      bestComments = comments;
    }
  }

  return bestScore >= 0.35 ? bestComments : [];
}

function translateStatus(status = "") {
  const s = String(status || "").trim();

  if (/^closed$/i.test(s)) return "Chiusa";
  if (/^new$/i.test(s)) return "Aperta";
  if (/^waiting$/i.test(s)) return "In attesa";
  if (/^unknown$/i.test(s)) return "Non definito";

  return s;
}

function getFirstColumnValue(row: any) {
  if (!row || typeof row !== "object") return "";
  const firstKey = Object.keys(row)[0];
  return firstKey ? row[firstKey] : "";
}

function getTodoLabel(todo: any) {
  return (
    todo.Label ||
    getFirstColumnValue(todo) ||
    todo.ID ||
    todo.Guid ||
    todo.GUID ||
    ""
  );
}

function getTodoAssignees(todo: any) {
  return (
    todo["Assignee(s)"] ||
    todo["Assignee(s) "] ||
    todo.Assignees ||
    todo.Assignee ||
    todo.AssignedTo ||
    todo["Assigned to"] ||
    ""
  );
}

function findBestBcfTopic(todo: any, topicsByKey: Map<string, any>) {
  const candidates = [
    getTodoLabel(todo),
    todo.ID,
    todo.Guid,
    todo.GUID,
    todo.Title,
    String(todo.Title || "").replace(/\.pdf/gi, ""),
  ].filter(Boolean);

  for (const c of candidates) {
    const key = normalize(c);
    if (topicsByKey.has(key)) return topicsByKey.get(key);
  }

  let bestScore = 0;
  let bestTopic: any = null;

  for (const [topicKey, topic] of topicsByKey.entries()) {
    const score = Math.max(...candidates.map((c) => similarity(c, topicKey)));

    if (score > bestScore) {
      bestScore = score;
      bestTopic = topic;
    }
  }

  return bestScore >= 0.35 ? bestTopic : null;
}

function extractMarkup(parsed: any) {
  return (
    parsed?.Markup ||
    parsed?.markup ||
    parsed?.bcf?.Markup ||
    parsed?.Bcf?.Markup ||
    parsed
  );
}

function extractTopic(markup: any, fallbackGuid: string) {
  const topic = getAny(markup, ["Topic", "topic"]) || {};

  const title = getXmlText(getAny(topic, ["Title", "title", "TopicTitle", "Name"]));
  const description = getXmlText(getAny(topic, ["Description", "description"]));
  const labelsRaw = getAny(topic, ["Labels", "labels", "Label", "label"]);
  const labels = Array.isArray(labelsRaw)
    ? labelsRaw.map((x) => getXmlText(x)).join(" ")
    : getXmlText(labelsRaw);

  const guid =
    getXmlText(getAny(topic, ["Guid", "guid", "GUID"])) ||
    topic.Guid ||
    topic.guid ||
    fallbackGuid;

  return {
    topic,
    topicTitle: title,
    topicDescription: description,
    topicLabels: labels,
    topicGuid: String(guid || fallbackGuid),
    topicStatus:
      getXmlText(getAny(topic, ["TopicStatus", "Status", "status"])) ||
      topic.TopicStatus ||
      "",
    topicPriority: getXmlText(getAny(topic, ["Priority", "priority"])),
    topicCreationDate: getXmlText(getAny(topic, ["CreationDate", "creationDate"])),
    topicCreationAuthor: getXmlText(getAny(topic, ["CreationAuthor", "creationAuthor"])),
    topicModifiedDate: getXmlText(getAny(topic, ["ModifiedDate", "modifiedDate"])),
    topicModifiedAuthor: getXmlText(getAny(topic, ["ModifiedAuthor", "modifiedAuthor"])),
    topicAssignedTo: getXmlText(getAny(topic, ["AssignedTo", "assignedTo"])),
  };
}

function extractComments(markup: any) {
  const raw =
    getAny(markup, ["Comment", "comment", "Comments", "comments"]) ||
    getAny(markup?.Comments, ["Comment", "comment"]) ||
    getAny(markup?.comments, ["Comment", "comment"]);

  if (Array.isArray(raw)) return raw;

  if (raw?.Comment) return arr(raw.Comment);
  if (raw?.comment) return arr(raw.comment);

  return arr(raw);
}

async function readBcfZip(file: File, buffer: Buffer) {
  const bcfComments: any[] = [];
  const bcfTopics: any[] = [];

  const zip = await JSZip.loadAsync(buffer);
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "",
    textNodeName: "#text",
    trimValues: true,
  });

  const markupPaths = Object.keys(zip.files).filter((path) =>
    path.toLowerCase().endsWith("markup.bcf")
  );

  for (const path of markupPaths) {
    const xml = await zip.files[path].async("text");
    const parsed = parser.parse(xml);
    const markup = extractMarkup(parsed);

    const folderGuid = path.split("/")[0] || "";
    const topicData = extractTopic(markup, folderGuid);

    bcfTopics.push({
      sourceFile: file.name,
      markupPath: path,
      Label: topicData.topicGuid,
      ID: topicData.topicGuid,
      Guid: topicData.topicGuid,
      Title: topicData.topicTitle,
      Description: topicData.topicDescription,
      Tags: topicData.topicLabels,
      Status: topicData.topicStatus,
      Priority: topicData.topicPriority,
      "Created by": topicData.topicCreationAuthor,
      "Created on": topicData.topicCreationDate,
      "Last modified by": topicData.topicModifiedAuthor,
      "Last modified on": topicData.topicModifiedDate,
      "Assignee(s)": topicData.topicAssignedTo,
      Groups: "",
      Type: "BCF Topic",
      __source: "bcfzip",
    });

    const comments = extractComments(markup);

    for (const c of comments) {
      const text = getXmlText(getAny(c, ["Comment", "comment", "Text", "text"]));
      const role = roleFromText(text);

      bcfComments.push({
        sourceFile: file.name,
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

  return {
    bcfTopics,
    bcfComments,
    markupCount: markupPaths.length,
  };
}

export async function POST(req: Request) {
  try {
    const formData = await req.formData();
    const files = formData.getAll("files") as File[];

    let excelRows: any[] = [];
    const bcfTopicRows: any[] = [];
    const bcfComments: any[] = [];
    const importedFiles: any[] = [];

    for (const file of files) {
      const buffer = Buffer.from(await file.arrayBuffer());
      const name = file.name.toLowerCase();

      if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
        const workbook = XLSX.read(buffer, { type: "buffer" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        excelRows = [...excelRows, ...rows];

        importedFiles.push({
          fileName: file.name,
          type: "xlsx",
          rows: rows.length,
        });
      }

      if (name.endsWith(".bcfzip") || name.endsWith(".zip")) {
        const result = await readBcfZip(file, buffer);

        bcfTopicRows.push(...result.bcfTopics);
        bcfComments.push(...result.bcfComments);

        importedFiles.push({
          fileName: file.name,
          type: "bcfzip",
          markupCount: result.markupCount,
          topics: result.bcfTopics.length,
          comments: result.bcfComments.length,
        });
      }
    }

    const commentsByTopic = new Map<string, any[]>();

    for (const c of bcfComments) {
      const keys = [
        normalize(c.topicTitle),
        normalize(c.topicGuid),
        normalize(String(c.topicTitle || "").replace(/\.pdf/gi, "")),
      ].filter(Boolean);

      for (const key of keys) {
        if (!commentsByTopic.has(key)) commentsByTopic.set(key, []);
        commentsByTopic.get(key)!.push(c);
      }
    }

    const topicsByKey = new Map<string, any>();

    for (const topic of bcfTopicRows) {
      const keys = [
        normalize(topic.Label),
        normalize(topic.ID),
        normalize(topic.Guid),
        normalize(topic.Title),
        normalize(String(topic.Title || "").replace(/\.pdf/gi, "")),
      ].filter(Boolean);

      for (const key of keys) {
        if (!topicsByKey.has(key)) topicsByKey.set(key, topic);
      }
    }

    // Se è presente l'Excel, l'Excel resta la base principale.
    // Se non è presente l'Excel, i Topic BCF diventano direttamente le righe dashboard.
    const todoRows = excelRows.length > 0 ? excelRows : bcfTopicRows;

    const rows = todoRows.map((todo: any, index: number) => {
      const matchedBcfTopic = findBestBcfTopic(todo, topicsByKey);

      // ID dashboard: prioritariamente il Label Trimble, cioè la prima colonna del ToDo.
      const label = getTodoLabel(todo);

      const title =
        todo.Title ||
        matchedBcfTopic?.Title ||
        "";

      const todoDescription = todo.Description || "";
      const bcfDescription = matchedBcfTopic?.Description || "";

      // Descrizione: prima il BCF, se presente; altrimenti la colonna Description del ToDo.
      const description = String(bcfDescription || "").trim()
        ? bcfDescription
        : todoDescription;

      const tags =
        todo.Tags ||
        todo.Labels ||
        matchedBcfTopic?.Tags ||
        "";

      const tipo = detectTipo(tags, description);

      const tipologiaNcOss = detectTipologiaNcOss(
        tags,
        description,
        title,
        tipo
      );

      // Disciplina: deriva dalla colonna Assignee(s) del ToDo.
      // Se manca l'Excel e la riga deriva dal BCF, viene usato l'AssignedTo del BCF come fallback.
      const disciplina =
        getTodoAssignees(todo) ||
        matchedBcfTopic?.["Assignee(s)"] ||
        "";

      const statoOriginale =
        matchedBcfTopic?.Status ||
        todo.Status ||
        "";

      const statoTradotto = translateStatus(statoOriginale);

      const ispettore =
        todo["Created by"] ||
        todo["Last modified by"] ||
        todo.Owner ||
        "";

      const comments = findBestComments(todo, commentsByTopic).sort((a, b) =>
        String(a.date).localeCompare(String(b.date))
      );

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

      const esitoCompilato =
        tipo === "NC" ||
        tipo === "OSS" ||
        tipo === "Nessun rilievo" ||
        tipo === "Da NC a OSS";

      const titleCompilato = Boolean(String(title).trim());
      const disciplinaCompilata = Boolean(String(disciplina).trim());

      const campiMancantiControllo: string[] = [];

      if (!esitoCompilato) {
        campiMancantiControllo.push("Tags / Esito verifica mancante o non coerente");
      }

      if (!titleCompilato) {
        campiMancantiControllo.push("Title / Codice elaborato mancante");
      }

      if (!disciplinaCompilata) {
        campiMancantiControllo.push("Gruppo / Disciplina mancante");
      }

      const controlloIspettoreCompleto =
        esitoCompilato &&
        titleCompilato &&
        disciplinaCompilata;

      return {
        idRecord: `IMPORT-${index + 1}`,
        id: label,
        idTodo: label,

        elaborato: title,
        titolo: title,
        descrizione: description,

        tipo,
        tipoOriginale: todo.Type || "",
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
        statoCompilazioneIspettore: controlloIspettoreCompleto
          ? "Completo"
          : "Incompleto",

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
      };
    });

    return NextResponse.json({
      ok: true,
      importedFiles,
      excelRowCount: excelRows.length,
      bcfTopicCount: bcfTopicRows.length,
      todoCount: todoRows.length,
      bcfCommentCount: bcfComments.length,
      rows,
      comments: bcfComments,
    });
  } catch (error: any) {
    return NextResponse.json(
      {
        ok: false,
        error: error?.message || "Errore durante la lettura dei file",
      },
      { status: 500 }
    );
  }
}
