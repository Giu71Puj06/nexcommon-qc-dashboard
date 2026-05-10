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
  "Titolo elaborato",
  "Titolo Elaborato",
  "TITOLO ELABORATO",
  "Titolo elenco",
  "Titolo Elenco",
  "TITOLO ELENCO",
  "Nome elaborato",
  "Nome Elaborato",
];

type BcfTopicData = {
  topicGuid: string;
  titolo: string;
  descrizione: string;
  ispettore: string;
  ispettoreNomeBcf: string;
  labels: string;
  stato: string;
  commentiPRG: string;
  commentiISP: string;
  ultimoCommento: string;
};

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

function getElaboratoBase(value: string) {
  return String(value || "")
    .replace(/\.pdf$/i, "")
    .replace(/\.[A-Za-z0-9]+$/g, "")
    .replace(/[_-]\d{1,2}$/g, "")
    .trim();
}

function getRevisioneDaCodice(value: string) {
  const clean = String(value || "").replace(/\.pdf$/i, "").trim();
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

function sameDisciplina(a: string, b: string) {
  const aa = normalizeKey(a);
  const bb = normalizeKey(b);
  if (aa === bb) return true;

  const aliases: Record<string, string[]> = {
    DOCUMENTAZIONEECONOMICA: ["ECONOMICO", "ECONOMICA"],
    DOCUMENTAZIONEGENERALE: ["GENERALE"],
    SICUREZZACANTIERE: ["SICUREZZA", "SICUREZZACANTIERIZZAZIONEEBOB"],
    IMPIANTI: ["MEP", "IMPIANTISTICO"],
    ARCHITETTONICO: ["ARCHITETTURA"],
    STRUTTURALE: ["STRUTTURE"],
    AMBIENTEEVINCOLI: ["AMBIENTE", "VINCOLI"],
    INTERFERENZEEESPROPRI: ["INTERFERENZE", "ESPROPRI"],
  };

  return (aliases[bb] || []).includes(aa) || (aliases[aa] || []).includes(bb);
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

function bestTodoMatch(todoRows: any[], bcf: BcfTopicData) {
  const titleKey = normalizeKey(bcf.titolo);

  const sameTitle = todoRows.filter((r: any) => {
    const titoloTodo = findValue(r, ["Title", "Titolo", "TITLE", "Topic", "Nome"]);
    return normalizeKey(titoloTodo) === titleKey;
  });

  if (sameTitle.length === 1) return sameTitle[0];

  if (sameTitle.length > 1) {
    const ranked = sameTitle
      .map((r: any) => ({
        row: r,
        score: descriptionScore(
          findValue(r, ["Description", "Descrizione"]),
          bcf.descrizione
        ),
      }))
      .sort((a, b) => b.score - a.score);

    if (ranked[0] && ranked[0].score > 0.03) return ranked[0].row;
  }

  const rankedAll = todoRows
    .map((r: any) => ({
      row: r,
      score:
        descriptionScore(
          findValue(r, ["Description", "Descrizione"]),
          bcf.descrizione
        ) +
        (normalizeKey(findValue(r, ["Title", "Titolo", "TITLE", "Topic", "Nome"])) === titleKey
          ? 1
          : 0),
    }))
    .sort((a, b) => b.score - a.score);

  if (rankedAll[0] && rankedAll[0].score > 0.08) return rankedAll[0].row;

  return {};
}

function getElencoInfoByCode(elencoInfoMap: Record<string, any>, codice: string) {
  return (
    elencoInfoMap[normalizeKey(codice)] ||
    elencoInfoMap[normalizeKey(getElaboratoBase(codice))] ||
    {}
  );
}

function readFirstSheet(workbook: XLSX.WorkBook) {
  return XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
}

function readReportRows(workbook: XLSX.WorkBook) {
  const sheetName =
    workbook.SheetNames.find((n) => normalizeKey(n) === "VERIFICAELABORATI") ||
    workbook.SheetNames[0];

  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
}

export async function POST(req: Request) {
  try {
    const formData = await req.formData();

    const todoFile = formData.get("todo") as File;
    const bcfFiles = formData.getAll("bcf") as File[];
    const elencoFile = formData.get("elenco") as File;
    const templateFile = formData.get("template") as File;
    const reportFile = formData.get("files") as File | null;

    const progettisti = parsePeopleList(formData.get("progettisti"));
    const ispettori = parsePeopleList(formData.get("ispettori"));

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
        normalizeKey(codice),
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
      const codice = findValue(r, [
        "Codice elaborato",
        "Codice Elaborato",
        "CODICE ELABORATO",
        "Codice_SP",
        "Codice SP",
      ]);

      if (!codice) return;

      const codicePulito = String(codice).replace(/\.pdf$/i, "").trim();
      const base = getElaboratoBase(codicePulito);

      const info = {
        codice: codicePulito,
        codiceBase: base,
        revisione:
          findValue(r, ["REV", "REV.", "Rev.", "Revisione", "REVISIONE"]) ||
          getRevisioneDaCodice(codicePulito),
        titolo: findValue(r, TITOLO_ELABORATO_COLUMNS),
        disciplina:
          findValue(r, ["DISCIPLINA", "Disciplina", "Oggetto", "OGGETTO"]) ||
          disciplinaFromCodice(codicePulito),
      };

      reportInfoMap[normalizeKey(codicePulito)] = info;
      reportInfoMap[normalizeKey(base)] = info;
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
          const author = getCommentAuthor(c);

          if (!cleanText) return;

          const isPRGByAccount = isInPeopleList(author, progettisti);
          const isISPByAccount = isInPeopleList(author, ispettori);

          if (isPRGByAccount) {
            commentiPRGList.push(cleanText);
            return;
          }

          if (isISPByAccount) {
            lastIspAuthor = author || lastIspAuthor;
            commentiISPList.push(cleanText);
            return;
          }

          if (/\(\s*PRG\s*\)/i.test(testo)) {
            commentiPRGList.push(cleanText);
            return;
          }

          if (/\(\s*ISP\s*\)/i.test(testo)) {
            lastIspAuthor = author || lastIspAuthor;
            commentiISPList.push(cleanText);
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

        const dataTopic: BcfTopicData = {
          topicGuid: topicGuid || existing?.topicGuid || "",
          titolo: topicTitle || existing?.titolo || "",
          descrizione: topicDescription || existing?.descrizione || "",
          ispettore:
            existing?.ispettore ||
            siglaDaNome(ispettoreNomeBcf || topicAuthor),
          ispettoreNomeBcf:
            ispettoreNomeBcf || existing?.ispettoreNomeBcf || "",
          labels: topicLabels || existing?.labels || "",
          stato: topicStatus || existing?.stato || "",
          commentiPRG: mergeText(existing?.commentiPRG || "", commentiPRGList),
          commentiISP: mergeText(existing?.commentiISP || "", commentiISPList),
          ultimoCommento:
            comments.length > 0
              ? getCommentText(comments[comments.length - 1])
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

    const finalRows = bcfTopics
      .filter((bcf) => !isNessunRilievo(bcf.labels, bcf.descrizione))
      .map((bcf) => {
        const todoMatch = bestTodoMatch(todoRows, bcf);

        const label = findValue(todoMatch, ["Label", "Etichetta"]);
        const tags =
          findValue(todoMatch, ["Tags", "Tag", "Tipo", "Esito"]) ||
          bcf.labels ||
          "";

        const disciplina =
          findValue(todoMatch, ["Assignee(s) ", "Assignee(s)", "Disciplina"]) ||
          disciplinaFromCodice(bcf.titolo);

        const tipoBase = String(tags || bcf.labels || "")
          .toUpperCase()
          .includes("OSS")
          ? "OSS"
          : "NC";

        const codiceTR =
          label ? labelToTR(label) : extractTR(bcf.titolo) || extractTR(bcf.descrizione) || "";

        const reportInfo =
          reportInfoMap[normalizeKey(bcf.titolo)] ||
          reportInfoMap[normalizeKey(getElaboratoBase(bcf.titolo))] ||
          {};

        const disciplinaInfo = disciplinaInfoMap[normalizeKey(disciplina)] || {};

        const infoElenco =
          getElencoInfoByCode(elencoInfoMap, bcf.titolo) ||
          getElencoInfoByCode(elencoInfoMap, getElaboratoBase(bcf.titolo)) ||
          {};

        const ispettoreTodo = getIspettoreFromTodo(todoMatch);
        const ispettoreElenco =
          infoElenco.ispettoreElenco || disciplinaInfo.ispettoreElenco || "";

        const ispettoreFinale = resolveIspettoreFinale(
          bcf.ispettore || "",
          bcf.ispettoreNomeBcf || "",
          disciplina,
          ispettoreTodo || ispettoreElenco
        );

        const titoloElaborato =
          getTitoloElaboratoFromTodo(todoMatch) ||
          reportInfo.titolo ||
          infoElenco.titolo ||
          "";

        return {
          Disciplina: disciplina,
          Label: label,
          TipoBase: tipoBase,
          CodiceTR: codiceTR,
          "Codice Rilievo": label || codiceTR,
          "Codice Elaborato": bcf.titolo || "",
          "Titolo Elaborato": titoloElaborato || bcf.titolo || "",
          Revisione: reportInfo.revisione || getRevisioneDaCodice(bcf.titolo),
          Tipo: tags || tipoBase,
          "Descrizione Rilievo": bcf.descrizione || "",
          Ispettore: ispettoreFinale,
          "Risposta Progettista PRG": bcf.commentiPRG || "",
          "Riscontro Ispettore ISP": bcf.commentiISP || "",
          "Ultimo Commento": cleanRolePrefix(bcf.ultimoCommento || ""),
          "Azione Richiesta": "",
          Stato: normalizeStatus(bcf.stato || "", tags),
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
        const codiceCompleto = findValue(e, [
          "Codice elaborato",
          "Codice Elaborato",
          "CODICE ELABORATO",
          "Codice_SP",
          "Codice SP",
        ]);

        const codicePulito = String(codiceCompleto || "")
          .replace(/\.pdf$/i, "")
          .trim();

        if (!codicePulito) return null;

        const infoElenco = getElencoInfoByCode(elencoInfoMap, codicePulito);
        const reportInfo =
          reportInfoMap[normalizeKey(codicePulito)] ||
          reportInfoMap[normalizeKey(getElaboratoBase(codicePulito))] ||
          {};

        const titoloElenco = reportInfo.titolo || infoElenco.titolo || "";
        const revisione =
          reportInfo.revisione ||
          infoElenco.revisione ||
          getRevisioneDaCodice(codicePulito);

        const codiceSenzaRev = getElaboratoBase(codicePulito);

        const disciplinaElaborato =
          infoElenco.disciplina ||
          reportInfo.disciplina ||
          findValue(e, ["Disciplina", "DISCIPLINA", "Oggetto", "OGGETTO"]) ||
          disciplinaFromCodice(codicePulito);

        const key = normalizeKey(codicePulito);
        const baseKey = normalizeKey(codiceSenzaRev);

        const hasNC = finalRows.some(
          (r) =>
            (normalizeKey(r["Codice Elaborato"]) === key ||
              normalizeKey(getElaboratoBase(r["Codice Elaborato"])) === baseKey) &&
            r.TipoBase === "NC"
        );

        const hasOSS = finalRows.some(
          (r) =>
            (normalizeKey(r["Codice Elaborato"]) === key ||
              normalizeKey(getElaboratoBase(r["Codice Elaborato"])) === baseKey) &&
            r.TipoBase === "OSS"
        );

        return {
          codice_elaborato: codiceSenzaRev,
          codice_file: codicePulito,
          revisione,
          titolo_elaborato: titoloElenco,
          disciplina: disciplinaElaborato,
          presenza_nc: hasNC ? "X" : "",
          presenza_oss: hasOSS ? "X" : "",
          assenza_nc_oss: !hasNC && !hasOSS ? "X" : "",
        };
      })
      .filter(Boolean) as any[];

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
          const codiceSenzaRev = getElaboratoBase(codiceCompleto);

          return {
            codice_elaborato: codiceSenzaRev,
            codice_file: codiceCompleto,
            revisione: findValue(e, ["REV.", "REV", "Rev.", "Revisione"]),
            titolo_elaborato: getTitoloProgetto(e),
            disciplina: findValue(e, ["DISCIPLINA", "Disciplina"]),
            presenza_nc: "",
            presenza_oss: "",
            assenza_nc_oss: "X",
          };
        });
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

      const numeroChiuse = elaboratiVerificati.filter(
        (e) => e.assenza_nc_oss === "X"
      ).length;

      const totaleDocumenti = elaboratiVerificati.length;

      let buffer: Buffer;

      try {
        const zipDocx = new PizZip(templateBuffer);

        const doc = new Docxtemplater(zipDocx, {
          paragraphLoop: true,
          linebreaks: true,
        });

        doc.render({
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
      } catch (e: any) {
        outputZip.file(
          `ERRORE_TEMPLATE_${safeName(disciplina)}.txt`,
          JSON.stringify(e, null, 2)
        );
        continue;
      }

      outputZip.file(
        `${codiceScheda}_${safeName(disciplina)}_${REVISIONE_SCHEDA}.docx`,
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
