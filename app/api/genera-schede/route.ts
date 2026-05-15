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

type ImageAttachment = {
  tr: string;
  fileName: string;
  contentType: string;
  ext: string;
  buffer: Buffer;
  bookmarkName: string;
};

function descrizioneRevisioneScheda(rev: string) {
  const n = Number(String(rev || "0").trim());

  if (!Number.isFinite(n) || n <= 0) return "Prima Emissione - Rilievi";
  if (n === 1) return "Seconda emissione - Riscontri";
  if (n === 2) return "Terza emissione - Riscontri";
  if (n === 3) return "Quarta emissione - Riscontri";
  if (n === 4) return "Quinta emissione - Riscontri";

  return `${n + 1}ª emissione - Riscontri`;
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

  return `${formattedDate} - ${cleanText}`;
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

function safeBookmarkName(value: string) {
  return `IMG_${normalizeTR(value).replace(/[^A-Z0-9]/g, "_")}`;
}

function getImageContentType(fileName: string) {
  const lower = String(fileName || "").toLowerCase();

  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  if (lower.endsWith(".webp")) return "image/webp";

  return "";
}

function getImageExt(fileName: string) {
  const lower = String(fileName || "").toLowerCase();

  if (lower.endsWith(".png")) return "png";
  if (lower.endsWith(".jpg")) return "jpg";
  if (lower.endsWith(".jpeg")) return "jpeg";
  if (lower.endsWith(".webp")) return "webp";

  return "";
}

async function collectImageAttachments(formData: FormData) {
  const images: ImageAttachment[] = [];
  const seen = new Set<string>();

  const hasImmagini = String(formData.get("has_immagini") || "false") === "true";

  if (!hasImmagini) {
    return images;
  }

  async function addImage(fileName: string, buffer: Buffer) {
    const tr = normalizeTR(fileName);
    const contentType = getImageContentType(fileName);
    const ext = getImageExt(fileName);

    if (!tr || !contentType || !ext) return;

    const key = `${tr}_${fileName}`;
    if (seen.has(key)) return;
    seen.add(key);

    images.push({
      tr,
      fileName,
      contentType,
      ext,
      buffer,
      bookmarkName: safeBookmarkName(tr),
    });
  }

  try {
    const directFiles = [
      ...formData.getAll("foto"),
      ...formData.getAll("immagini"),
    ];

    for (const entry of directFiles) {
      const file = entry as File;
      if (!file || typeof file.arrayBuffer !== "function") continue;

      await addImage(file.name, Buffer.from(await file.arrayBuffer()));
    }
  } catch (err) {
    console.warn("Errore lettura immagini singole. Continuo senza immagini.", err);
  }

  try {
    const zipFiles = [
      ...formData.getAll("fotoZip"),
      ...formData.getAll("immaginiZip"),
    ];

    for (const entry of zipFiles) {
      const zipFile = entry as File;
      if (!zipFile || typeof zipFile.arrayBuffer !== "function") continue;

      try {
        const zip = await JSZip.loadAsync(
          Buffer.from(await zipFile.arrayBuffer())
        );

        for (const fileName of Object.keys(zip.files)) {
          const zipEntry = zip.files[fileName];
          if (zipEntry.dir) continue;

          const contentType = getImageContentType(fileName);
          if (!contentType) continue;

          await addImage(
            fileName.split("/").pop() || fileName,
            await zipEntry.async("nodebuffer")
          );
        }
      } catch (zipErr) {
        console.warn(
          `ZIP immagini non valido o non leggibile: ${zipFile.name}. Continuo senza immagini.`,
          zipErr
        );
      }
    }
  } catch (err) {
    console.warn("Errore lettura ZIP immagini. Continuo senza immagini.", err);
  }

  return images;
}

function buildImagesByTR(images: ImageAttachment[]) {
  const map: Record<string, ImageAttachment[]> = {};

  images.forEach((img) => {
    if (!map[img.tr]) map[img.tr] = [];
    map[img.tr].push(img);
  });

  return map;
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

function remapIspettoreByDisciplina(sigla: string, disciplina: string) {
  const s = String(sigla || "").trim().toUpperCase();
  const d = normalizeKey(disciplina);

  if (d === "SICUREZZACANTIERE") {
    if (s === "(CR)" || s === "(CA)" || s === "(GB)") return "(FM)";
  }

  if (d === "DOCUMENTAZIONEECONOMICA") {
    if (s === "(CS)") return "(MG)";
    if (s === "(CR)" || s === "(CA)" || s === "(GB)") return "(OB)";
  }

  return sigla;
}

function remapIspettoreFinale(sigla: string, disciplina: string) {
  const s = String(sigla || "").trim().toUpperCase();
  const d = normalizeKey(disciplina);

  if (s === "(CR)" || s === "(CA)") return "(OB)";
  if (d === "SICUREZZACANTIERE" && s === "(GB)") return "(FM)";
  if (d === "DOCUMENTAZIONEECONOMICA" && s === "(GB)") return "(OB)";
  if (d === "DOCUMENTAZIONEECONOMICA" && s === "(CS)") return "(MG)";

  return remapIspettoreByDisciplina(sigla, disciplina);
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

function normalizeStatus(status: string, tags: string) {
  const s = String(status || "").toUpperCase();
  const t = String(tags || "").toUpperCase();

  if (
    t.includes("DA NC A OSS") ||
    t.includes("NC A OSS") ||
    t.includes("DA NC AD OSS")
  ) {
    return "NC declassata a OSS";
  }

  if (s.includes("CLOSED") || s.includes("CHIUS")) return "Chiusa";
  if (s.includes("NEW") || s.includes("OPEN") || s.includes("APERT")) return "Aperta";

  return status || "";
}

function isNessunRilievo(tags: string, descrizione: string) {
  return `${tags || ""} ${descrizione || ""}`
    .toUpperCase()
    .includes("NESSUN RILIEVO");
}

function isClosedStatus(status: string) {
  const s = String(status || "").toUpperCase();
  return s.includes("CLOSED") || s.includes("CHIUS");
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
  const tipo = normalizeKey(row?.TipoBase || row?.Tipo || row?.["Codice Rilievo"] || "");
  return tipo.includes("OSS");
}

function buildRilieviFlagsByElaborato(rows: any[]) {
  const map: Record<string, { hasNC: boolean; hasOSS: boolean; sourceRow?: any }> = {};

  for (const row of rows) {
    const key = elaboratoAggregationKey(row?.["Codice Elaborato"] || "");
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
    const key = elaboratoAggregationKey(e.codice_file || e.codice_elaborato);
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

function bestBcfMatchForTodo(bcfTopics: BcfTopicData[], todo: any) {
  const titoloTodo = findValue(todo, ["Title", "Titolo", "TITLE", "Topic", "Nome"]);
  const descrizioneTodo = findValue(todo, ["Description", "Descrizione"]);
  const labelTodo = findValue(todo, ["Label", "Etichetta"]);
  const codiceTodo = extractCodiceElaborato(titoloTodo || descrizioneTodo);

  const ranked = bcfTopics
    .map((bcf) => {
      const codiceBcf = extractCodiceElaborato(bcf.titolo || bcf.descrizione);
      let score = 0;

      if (codiceTodo && sameElaboratoCode(codiceBcf, codiceTodo)) score += 10;
      if (normalizeKey(bcf.titolo) === normalizeKey(titoloTodo)) score += 5;
      if (labelTodo && extractTR(labelTodo) && extractTR(labelTodo) === extractTR(bcf.titolo)) {
        score += 3;
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

    return tableXml.replace(/<w:tr\b[^>]*>[\s\S]*?<\/w:tr>/, (headerRowXml) => {
      let rowXml = headerRowXml;

      if (rowXml.includes("<w:tblHeader")) {
        if (!rowXml.includes("<w:cantSplit")) {
          rowXml = rowXml.replace(/<w:trPr\b([^>]*)>/, '<w:trPr$1><w:cantSplit/>');
        }
        return rowXml;
      }

      if (/<w:trPr\b[^>]*>/.test(rowXml)) {
        rowXml = rowXml.replace(
          /<w:trPr\b([^>]*)>/,
          '<w:trPr$1><w:tblHeader w:val="true"/><w:cantSplit/>'
        );
      } else {
        rowXml = rowXml.replace(
          /<w:tr\b([^>]*)>/,
          '<w:tr$1><w:trPr><w:tblHeader w:val="true"/><w:cantSplit/></w:trPr>'
        );
      }

      return rowXml;
    });
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

function getNextRelationshipId(relsXml: string) {
  const ids = Array.from(relsXml.matchAll(/Id="rId(\d+)"/g))
    .map((m) => Number(m[1]))
    .filter((n) => Number.isFinite(n));

  return ids.length ? Math.max(...ids) + 1 : 1;
}

function ensureImageContentTypes(contentTypesXml: string, images: ImageAttachment[]) {
  let xml = contentTypesXml;
  const required: Record<string, string> = {};

  images.forEach((img) => {
    required[img.ext] = img.contentType;
  });

  Object.entries(required).forEach(([ext, contentType]) => {
    if (xml.includes(`Extension="${ext}"`)) return;

    xml = xml.replace(
      "</Types>",
      `<Default Extension="${escapeXml(ext)}" ContentType="${escapeXml(contentType)}"/></Types>`
    );
  });

  return xml;
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

function buildImageDrawingXml(rId: string, name: string) {
  return `
  <w:p>
    <w:pPr><w:spacing w:before="120" w:after="120"/></w:pPr>
    <w:r>
      <w:drawing>
        <wp:inline distT="0" distB="0" distL="0" distR="0">
          <wp:extent cx="5200000" cy="3600000"/>
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
          <wp:docPr id="${Math.floor(Math.random() * 100000)}" name="${escapeXml(name)}"/>
          <wp:cNvGraphicFramePr>
            <a:graphicFrameLocks noChangeAspect="1"/>
          </wp:cNvGraphicFramePr>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic>
                <pic:nvPicPr>
                  <pic:cNvPr id="0" name="${escapeXml(name)}"/>
                  <pic:cNvPicPr/>
                </pic:nvPicPr>
                <pic:blipFill>
                  <a:blip r:embed="${escapeXml(rId)}"/>
                  <a:stretch><a:fillRect/></a:stretch>
                </pic:blipFill>
                <pic:spPr>
                  <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="5200000" cy="3600000"/>
                  </a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                </pic:spPr>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>
  </w:p>`;
}

function buildImmaginiAllegateXml(
  images: ImageAttachment[],
  rIdByImageKey: Record<string, string>
) {
  if (images.length === 0) return "";

  const blocks = images
    .map((img, index) => {
      const bookmarkId = 9000 + index;
      const title = `${img.tr} - ${img.fileName}`;
      const rId = rIdByImageKey[`${img.tr}_${img.fileName}`];

      return `
      <w:p>
        <w:pPr><w:spacing w:before="240" w:after="80"/></w:pPr>
        <w:bookmarkStart w:id="${bookmarkId}" w:name="${escapeXml(img.bookmarkName)}"/>
        <w:r><w:rPr><w:b/></w:rPr><w:t>${escapeXml(title)}</w:t></w:r>
        <w:bookmarkEnd w:id="${bookmarkId}"/>
      </w:p>
      ${buildImageDrawingXml(rId, img.fileName)}
      `;
    })
    .join("");

  return `
  <w:p>
    <w:pPr><w:spacing w:before="360" w:after="160"/></w:pPr>
    <w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>IMMAGINI ALLEGATE</w:t></w:r>
  </w:p>
  ${blocks}
  <w:p/>`;
}

function applyImageLinksToRilieviCells(documentXml: string, images: ImageAttachment[]) {
  if (images.length === 0) return documentXml;

  return documentXml.replace(/<w:tc\b[^>]*>[\s\S]*?<\/w:tc>/g, (cellXml) => {
    const plain = getPlainTextFromWordXml(cellXml);
    const plainTr = normalizeTR(plain);

    if (!plainTr || !/(NC|OSS)\s*\d+/i.test(plain)) return cellXml;

    const matchedImage = images.find((img) => normalizeTR(img.tr) === plainTr);
    if (!matchedImage) return cellXml;

    const display =
      plain.match(/(?:NC|OSS)\s*\d+\s*\(?\s*TR[-_\s]?\d+[A-Z]?\s*\)?/i)?.[0] ||
      plain;

    const tcPr = cellXml.match(/<w:tcPr\b[^>]*>[\s\S]*?<\/w:tcPr>/)?.[0] || "";

    return `
    <w:tc>
      ${tcPr}
      <w:p>
        <w:hyperlink w:anchor="${escapeXml(matchedImage.bookmarkName)}">
          <w:r>
            <w:rPr>
              <w:color w:val="0563C1"/>
              <w:u w:val="single"/>
            </w:rPr>
            <w:t>${escapeXml(display)}</w:t>
          </w:r>
        </w:hyperlink>
      </w:p>
    </w:tc>`;
  });
}

async function postProcessSchedaIspettivaDocx(
  buffer: Buffer,
  sintesi: SchedaIspettivaSintesi,
  immaginiAllegate: ImageAttachment[] = []
) {
  const zip = await JSZip.loadAsync(buffer);
  const documentFile = zip.file("word/document.xml");
  if (!documentFile) return buffer;

  const rIdByImageKey: Record<string, string> = {};

  let relsXml =
    (await zip.file("word/_rels/document.xml.rels")?.async("string")) ||
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

  let nextRid = getNextRelationshipId(relsXml);

  immaginiAllegate.forEach((img, index) => {
    const mediaName = `image_nc_oss_${index + 1}_${safeName(img.tr)}.${img.ext}`;
    const rId = `rId${nextRid++}`;

    zip.file(`word/media/${mediaName}`, img.buffer);

    relsXml = relsXml.replace(
      "</Relationships>",
      `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${escapeXml(mediaName)}"/></Relationships>`
    );

    rIdByImageKey[`${img.tr}_${img.fileName}`] = rId;
  });

  zip.file("word/_rels/document.xml.rels", relsXml);

  const contentTypesFile = zip.file("[Content_Types].xml");
  if (contentTypesFile) {
    const contentTypesXml = await contentTypesFile.async("string");
    zip.file("[Content_Types].xml", ensureImageContentTypes(contentTypesXml, immaginiAllegate));
  }

  let documentXml = await documentFile.async("string");

  documentXml = ensureDocumentNamespaces(documentXml);
  documentXml = applyClosedRowsGreyText(documentXml);
  documentXml = applyImageLinksToRilieviCells(documentXml, immaginiAllegate);
  documentXml = appendSintesiAfterLastTable(documentXml, sintesi);

  const immaginiXml = buildImmaginiAllegateXml(immaginiAllegate, rIdByImageKey);
  if (immaginiXml) {
    const bodyCloseIndex = documentXml.lastIndexOf("</w:body>");
    if (bodyCloseIndex >= 0) {
      documentXml =
        documentXml.slice(0, bodyCloseIndex) +
        immaginiXml +
        documentXml.slice(bodyCloseIndex);
    }
  }

  zip.file("word/document.xml", documentXml);

  return Buffer.from(
    await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" })
  );
}

export async function POST(req: Request) {
  try {
    const formData = await req.formData();

    let immaginiAllegate: ImageAttachment[] = [];

    try {
      immaginiAllegate = await collectImageAttachments(formData);
    } catch (err) {
      console.warn("Immagini non disponibili. Continuo senza immagini.", err);
      immaginiAllegate = [];
    }

    const immaginiByTR = buildImagesByTR(immaginiAllegate);

    const todoFile = formData.get("todo") as File;
    const bcfFiles = formData.getAll("bcf") as File[];
    const elencoFile = formData.get("elenco") as File;
    const templateFile = formData.get("template") as File;
    const reportFile = formData.get("files") as File | null;

    const progettisti = parsePeopleList(formData.get("progettisti"));
    const ispettori = parsePeopleList(formData.get("ispettori"));

    const revisioneScheda = String(
      formData.get("revisione_scheda") || REVISIONE_SCHEDA
    ).trim();
    const dataRevisioneScheda = String(
      formData.get("data_revisione_scheda") || "xx/xx/xxxx"
    ).trim();
    const responsabilePcq = String(formData.get("responsabile_pcq") || "").trim();
    const responsabileIts = String(formData.get("responsabile_its") || "").trim();

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

        const commentiPRGList: string[] = [];
        const commentiISPList: string[] = [];
        let lastIspAuthor = "";

        comments.forEach((c: any) => {
          const testo = getCommentText(c);
          const cleanText = cleanRolePrefix(testo);
          const datedCleanText = prefixCommentWithDate(cleanText, getCommentDateValue(c));
          const author = getCommentAuthor(c);

          if (!cleanText) return;

          const isPRGByAccount = isInPeopleList(author, progettisti);
          const isISPByAccount = isInPeopleList(author, ispettori);

          if (isPRGByAccount) {
            commentiPRGList.push(datedCleanText);
            return;
          }

          if (isISPByAccount) {
            lastIspAuthor = author || lastIspAuthor;
            commentiISPList.push(datedCleanText);
            return;
          }

          if (/\(\s*PRG\s*\)/i.test(testo)) {
            commentiPRGList.push(datedCleanText);
            return;
          }

          if (/\(\s*ISP\s*\)/i.test(testo)) {
            lastIspAuthor = author || lastIspAuthor;
            commentiISPList.push(datedCleanText);
            return;
          }
        });

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
          commentiPRG: mergeText(existing?.commentiPRG || "", commentiPRGList),
          commentiISP: mergeText(existing?.commentiISP || "", commentiISPList),
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
          label ? labelToTR(label) : extractTR(titoloTodo) || extractTR(descrizioneTodo) || "";

        const codiceElaborato = extractCodiceElaborato(titoloTodo || descrizioneTodo);
        const reportInfo = findReportInfo(reportInfoMap, codiceElaborato);
        const disciplinaInfo = disciplinaInfoMap[normalizeKey(disciplina)] || {};
        const infoElenco = getElencoInfoByCode(elencoInfoMap, codiceElaborato);

        const ispettoreTodo = getIspettoreFromTodo(todo);
        const ispettoreElenco =
          infoElenco.ispettoreElenco || disciplinaInfo.ispettoreElenco || "";

        const ispettoreFinale = ispettoreTodo
          ? remapIspettoreFinale(siglaDaNome(ispettoreTodo), disciplina)
          : resolveIspettoreFinale(
              bcf?.ispettore || "",
              bcf?.ispettoreNomeBcf || "",
              disciplina,
              ispettoreElenco
            );

        const titoloElaborato =
          reportInfo.titolo ||
          getTitoloElaboratoFromTodo(todo) ||
          infoElenco.titolo ||
          "";

        return {
          Disciplina: disciplina,
          Label: label,
          TipoBase: tipoBase,
          CodiceTR: codiceTR,
          "Codice Rilievo": label || codiceTR,
          "Codice Elaborato": codiceElaborato || titoloTodo || "",
          "Titolo Elaborato": titoloElaborato || codiceElaborato || titoloTodo || "",
          Revisione:
            reportInfo.revisione ||
            getRevisioneDaCodice(codiceElaborato || titoloTodo),
          Tipo: tags || tipoBase,
          "Descrizione Rilievo":
            bcf?.descrizioneConData ||
            bcf?.descrizione ||
            descrizioneTodo ||
            "",
          Ispettore: ispettoreFinale,
          "Risposta Progettista PRG": bcf?.commentiPRG || "",
          "Riscontro Ispettore ISP": bcf?.commentiISP || "",
          "Ultimo Commento": bcf?.ultimoCommento || "",
          "Azione Richiesta": "",
          Stato: normalizeStatus(findValue(todo, ["Status", "Stato"]) || bcf?.stato || "", tags),
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

        const titoloElenco =
          reportInfo.titolo ||
          findValue(e, TITOLO_ELABORATO_COLUMNS) ||
          infoElenco.titolo ||
          "";

        const revisione =
          reportInfo.revisione ||
          findValue(e, ["REV", "REV.", "Rev.", "Revisione", "REVISIONE"]) ||
          infoElenco.revisione ||
          getRevisioneDaCodice(codicePulito);

        const codiceSenzaRev = getElaboratoBase(codicePulito);

        const disciplinaElaborato =
          reportInfo.disciplina ||
          disciplinaFromReportCartella(findValue(e, ["Cartella", "CARTELLA"])) ||
          infoElenco.disciplina ||
          findValue(e, ["Disciplina", "DISCIPLINA", "Oggetto", "OGGETTO"]) ||
          disciplinaFromCodice(codicePulito);

        return {
          codice_elaborato: codiceSenzaRev,
          codice_file: codicePulito,
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
        elaboratiVerificati,
        rowsDisciplina
      );

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
        elaboratiVerificati = elencoDisciplina.map((e: any) => {
          const codiceCompleto = findValue(e, ["Codice_SP", "Codice SP"]);
          const codicePulito = extractCodiceElaborato(codiceCompleto);
          const codiceSenzaRev = getElaboratoBase(codicePulito);

          return {
            codice_elaborato: codiceSenzaRev,
            codice_file: codicePulito,
            revisione: findValue(e, ["REV.", "REV", "Rev.", "Revisione"]),
            titolo_elaborato: getTitoloProgetto(e),
            disciplina: findValue(e, ["DISCIPLINA", "Disciplina"]),
            presenza_nc: "",
            presenza_oss: "",
            assenza_nc_oss: "X",
          };
        });

        elaboratiVerificati = applyRilieviFlagsToElaborati(
          elaboratiVerificati,
          rowsDisciplina
        );
      }

      const progressivi: Record<string, number> = { NC: 0, OSS: 0 };

      const rilievi = rowsDisciplina.map((r) => {
        const tipo = String(r.TipoBase || "").includes("OSS") ? "OSS" : "NC";
        progressivi[tipo] += 1;

        return {
          tipo_progressivo: `${tipo}${progressivi[tipo]}\n(${r.CodiceTR || "TR-ND"})`,
          codice_elaborato: r["Codice Elaborato"] || "",
          titolo_elaborato: r["Titolo Elaborato"] || "",
          rilievo_its: r["Descrizione Rilievo"] || "",
          ispettore: r.Ispettore || "",
          risposta_prg: r["Risposta Progettista PRG"] || "",
          riscontro_isp: r["Riscontro Ispettore ISP"] || "",
          stato: r.Stato || "",
        };
      });

      const numeroNC = rilievi.filter((r) =>
        String(r.tipo_progressivo).startsWith("NC")
      ).length;

      const numeroOSS = rilievi.filter((r) =>
        String(r.tipo_progressivo).startsWith("OSS")
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

        const immaginiDisciplina = rowsDisciplina.flatMap((r) => {
          const tr = normalizeTR(r.CodiceTR || r.Label || r["Codice Rilievo"] || "");
          return tr ? immaginiByTR[tr] || [] : [];
        });

        buffer = await postProcessSchedaIspettivaDocx(
          buffer,
          {
            totaleElaboratiAnalizzati: totaleDocumenti,
            totaleNC: numeroNC,
            totaleOSS: numeroOSS,
            totaleChiuse: numeroChiuse,
          },
          immaginiDisciplina
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
        "Content-Disposition": 'attachment; filename="SCHEDE_ISPETTIVE_OUTPUT.zip"',
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
