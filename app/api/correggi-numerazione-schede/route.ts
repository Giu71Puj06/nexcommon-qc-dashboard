import { NextResponse } from "next/server";
import JSZip from "jszip";

type RilievoRecord = {
  file: string;
  rowIndex: number;
  codice: string;
  rilievo: string;
  key: string;
};

type Riferimenti = {
  byKey: Map<string, RilievoRecord>;
  byProgressivo: Map<string, RilievoRecord>;
  duplicateKeys: Set<string>;
  records: RilievoRecord[];
};

type ReportRow = {
  file: string;
  riga: number;
  codice_precedente: string;
  codice_originale: string;
  codice_finale: string;
  stato: "OK" | "CORRETTO" | "NON_TROVATO" | "SENZA_CODICE" | "DUPLICATO_RIFERIMENTO";
  rilievo_odi: string;
};

const WORD_DOCUMENT_XML = "word/document.xml";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(req: Request) {
  try {
    const formData = await req.formData();
    const precedente = formData.get("emissione_precedente");
    const daCorreggere = formData.get("emissione_da_correggere");

    if (
      !precedente ||
      !daCorreggere ||
      typeof (precedente as Blob).arrayBuffer !== "function" ||
      typeof (daCorreggere as Blob).arrayBuffer !== "function"
    ) {
      return new NextResponse(
        "Caricare entrambi gli ZIP: emissione precedente ed emissione da correggere.",
        { status: 400 }
      );
    }

    const zipPrecedente = await JSZip.loadAsync(await (precedente as Blob).arrayBuffer());
    const zipDaCorreggere = await JSZip.loadAsync(await (daCorreggere as Blob).arrayBuffer());

    const riferimenti = await estraiRiferimenti(zipPrecedente);

    if (riferimenti.records.length === 0) {
      return new NextResponse(
        "Nessun rilievo NC/OSS trovato nello ZIP dell'emissione precedente.",
        { status: 400 }
      );
    }

    const outputZip = new JSZip();
    const report: ReportRow[] = [];

    let totaleRighe = 0;
    let corretti = 0;
    let giaAllineati = 0;
    let mancanti = 0;
    let duplicati = 0;
    let riordinate = 0;
    let rinumerate = 0;

    for (const entry of Object.values(zipDaCorreggere.files)) {
      if (entry.dir) continue;

      const data = await entry.async("uint8array");
      const isDocx =
        entry.name.toLowerCase().endsWith(".docx") &&
        !entry.name.includes("__MACOSX/");

      if (!isDocx) {
        outputZip.file(entry.name, data);
        continue;
      }

      const docx = await JSZip.loadAsync(data);
      const documentFile = docx.file(WORD_DOCUMENT_XML);

      if (!documentFile) {
        outputZip.file(entry.name, data);
        continue;
      }

      const xml = await documentFile.async("string");

      const parsed = correggiDocumentXml(xml, entry.name, riferimenti, report);

      docx.file(WORD_DOCUMENT_XML, parsed.xml);

      totaleRighe += parsed.stats.totaleRighe;
      corretti += parsed.stats.corretti;
      giaAllineati += parsed.stats.giaAllineati;
      mancanti += parsed.stats.mancanti;
      duplicati += parsed.stats.duplicati;
      riordinate += parsed.stats.riordinate;
      rinumerate += parsed.stats.rinumerate;

      const correctedBuffer = await docx.generateAsync({ type: "uint8array" });
      outputZip.file(entry.name, correctedBuffer);
    }

    outputZip.file("report_correzione_numerazione.csv", buildCsv(report));
    outputZip.file(
      "log_correzione.txt",
      buildLog({ totaleRighe, corretti, giaAllineati, mancanti, duplicati, riordinate, rinumerate })
    );

    const out = await outputZip.generateAsync({ type: "uint8array" });

    return new NextResponse(Buffer.from(out), {
      status: 200,
      headers: {
        "Content-Type": "application/zip",
        "Content-Disposition": 'attachment; filename="SCHEDE_CORRETTE_NUMERAZIONE.zip"',
        "X-Correzione-Stats": JSON.stringify({
          totaleRighe,
          corretti,
          giaAllineati,
          mancanti,
          nonTrovati: mancanti,
          duplicati,
          riordinate,
          rinumerate,
        }),
      },
    });
  } catch (err) {
    console.error("Errore correzione numerazione schede:", err);

    return new NextResponse(
      err instanceof Error
        ? err.message
        : "Errore durante la correzione della numerazione NC/OSS.",
      { status: 500 }
    );
  }
}

async function estraiRiferimenti(zip: JSZip): Promise<Riferimenti> {
  const byKey = new Map<string, RilievoRecord>();
  const byProgressivo = new Map<string, RilievoRecord>();
  const duplicateKeys = new Set<string>();
  const records: RilievoRecord[] = [];

  for (const entry of Object.values(zip.files)) {
    if (entry.dir) continue;
    if (!entry.name.toLowerCase().endsWith(".docx")) continue;
    if (entry.name.includes("__MACOSX/")) continue;

    const docx = await JSZip.loadAsync(await entry.async("uint8array"));
    const documentFile = docx.file(WORD_DOCUMENT_XML);
    if (!documentFile) continue;

    const xml = await documentFile.async("string");
    const estratti = estraiRecordsDaDocumentXml(xml, entry.name);

    for (const record of estratti) {
      records.push(record);

      if (byKey.has(record.key)) {
        duplicateKeys.add(record.key);
      } else {
        byKey.set(record.key, record);
      }

      const progressivo = estraiProgressivo(record.codice);
      if (progressivo && !byProgressivo.has(progressivo)) {
        byProgressivo.set(progressivo, record);
      }
    }
  }

  return { byKey, byProgressivo, duplicateKeys, records };
}

function correggiDocumentXml(
  xml: string,
  fileName: string,
  riferimenti: Riferimenti,
  report: ReportRow[]
) {
  let output = xml;
  const replacements: Array<{ from: string; to: string }> = [];

  const stats = {
    totaleRighe: 0,
    corretti: 0,
    giaAllineati: 0,
    mancanti: 0,
    duplicati: 0,
    riordinate: 0,
    rinumerate: 0,
  };

  const tables = matchAll(xml, /<w:tbl[\s\S]*?<\/w:tbl>/g);

  for (const table of tables) {
    const rows = matchAll(table, /<w:tr[\s\S]*?<\/w:tr>/g);
    if (rows.length < 2) continue;

    const headerInfo = trovaColonneTabella(rows);
    if (!headerInfo) continue;

    for (let i = headerInfo.headerRowIndex + 1; i < rows.length; i += 1) {
      const row = rows[i];
      const cells = estraiCelle(row);

      if (cells.length <= Math.max(headerInfo.codiceCol, headerInfo.rilievoCol)) {
        continue;
      }

      const codiceOriginale = normalizzaCodice(estraiTesto(cells[headerInfo.codiceCol]));
      const rilievo = pulisciTestoRilievo(estraiTesto(cells[headerInfo.rilievoCol]));
      const key = normalizeRilievoKey(rilievo);

      if (!key || !rilievo || !sembraCodiceNcOss(codiceOriginale)) {
        continue;
      }

      stats.totaleRighe += 1;

      let ref = riferimenti.byKey.get(key);

      if (!ref) {
        const progressivo = estraiProgressivo(codiceOriginale);
        if (progressivo) {
          ref = riferimenti.byProgressivo.get(progressivo);
        }
      }

      if (!ref) {
        stats.mancanti += 1;
        report.push({
          file: fileName,
          riga: i + 1,
          codice_precedente: "",
          codice_originale: codiceOriginale,
          codice_finale: codiceOriginale,
          stato: "NON_TROVATO",
          rilievo_odi: rilievo,
        });
        continue;
      }

      if (riferimenti.duplicateKeys.has(key) && riferimenti.byKey.has(key)) {
        stats.duplicati += 1;
        report.push({
          file: fileName,
          riga: i + 1,
          codice_precedente: ref.codice,
          codice_originale: codiceOriginale,
          codice_finale: codiceOriginale,
          stato: "DUPLICATO_RIFERIMENTO",
          rilievo_odi: rilievo,
        });
        continue;
      }

      const codicePrecedente = normalizzaCodice(ref.codice);

      if (!codicePrecedente) {
        stats.mancanti += 1;
        report.push({
          file: fileName,
          riga: i + 1,
          codice_precedente: "",
          codice_originale: codiceOriginale,
          codice_finale: codiceOriginale,
          stato: "SENZA_CODICE",
          rilievo_odi: rilievo,
        });
        continue;
      }

      if (codicePrecedente === codiceOriginale) {
        stats.giaAllineati += 1;
        report.push({
          file: fileName,
          riga: i + 1,
          codice_precedente: codicePrecedente,
          codice_originale: codiceOriginale,
          codice_finale: codiceOriginale,
          stato: "OK",
          rilievo_odi: rilievo,
        });
        continue;
      }

      const newCell = sostituisciTestoCella(cells[headerInfo.codiceCol], codicePrecedente);
      const newCells = [...cells];
      newCells[headerInfo.codiceCol] = newCell;

      const newRow = ricostruisciRiga(row, newCells);
      replacements.push({ from: row, to: newRow });

      stats.corretti += 1;

      report.push({
        file: fileName,
        riga: i + 1,
        codice_precedente: codicePrecedente,
        codice_originale: codiceOriginale,
        codice_finale: codicePrecedente,
        stato: "CORRETTO",
        rilievo_odi: rilievo,
      });
    }
  }

  for (const replacement of replacements) {
    output = output.replace(replacement.from, replacement.to);
  }

  const riordino = riordinaTabelleCronologiche(output);
  output = riordino.xml;
  stats.riordinate += riordino.riordinate;

  const rinumerazione = rinumeraProgressiviFinali(output);
  output = rinumerazione.xml;
  stats.rinumerate += rinumerazione.rinumerate;

  return { xml: output, stats };
}

function riordinaTabelleCronologiche(xml: string) {
  let riordinate = 0;

  const newXml = xml.replace(/<w:tbl[\s\S]*?<\/w:tbl>/g, (tableXml) => {
    const rows = matchAll(tableXml, /<w:tr[\s\S]*?<\/w:tr>/g);
    if (rows.length < 3) return tableXml;

    const headerInfo = trovaColonneTabella(rows);
    if (!headerInfo) return tableXml;

    const headerRows = rows.slice(0, headerInfo.headerRowIndex + 1);
    const bodyRows = rows.slice(headerInfo.headerRowIndex + 1);

    const sortable = bodyRows.map((row, originalIndex) => {
      const cells = estraiCelle(row);
      const codice =
        cells.length > headerInfo.codiceCol
          ? normalizzaCodice(estraiTesto(cells[headerInfo.codiceCol]))
          : "";

      const parsed = parseCodiceOrdinamento(codice);

      return {
        row,
        originalIndex,
        tipo: parsed.tipo,
        numero: parsed.numero,
        ordinabile: parsed.ordinabile,
      };
    });

    const ordinabili = sortable.filter((x) => x.ordinabile);
    if (ordinabili.length < 2) return tableXml;

    const prima = ordinabili.map((x) => `${x.tipo}${x.numero}`).join("|");

    sortable.sort((a, b) => {
      if (!a.ordinabile && !b.ordinabile) return a.originalIndex - b.originalIndex;
      if (!a.ordinabile) return 1;
      if (!b.ordinabile) return -1;

      const tipoOrder = tipoSortOrder(a.tipo) - tipoSortOrder(b.tipo);
      if (tipoOrder !== 0) return tipoOrder;

      return a.numero - b.numero;
    });

    const dopo = sortable
      .filter((x) => x.ordinabile)
      .map((x) => `${x.tipo}${x.numero}`)
      .join("|");

    if (prima !== dopo) {
      riordinate += 1;
    }

    const orderedRows = [...headerRows, ...sortable.map((x) => x.row)];
    let idx = 0;

    return tableXml.replace(/<w:tr[\s\S]*?<\/w:tr>/g, () => {
      const row = orderedRows[idx];
      idx += 1;
      return row || "";
    });
  });

  return { xml: newXml, riordinate };
}

function rinumeraProgressiviFinali(xml: string) {
  let rinumerate = 0;

  const newXml = xml.replace(/<w:tbl[\s\S]*?<\/w:tbl>/g, (tableXml) => {
    const rows = matchAll(tableXml, /<w:tr[\s\S]*?<\/w:tr>/g);
    if (rows.length < 3) return tableXml;

    const headerInfo = trovaColonneTabella(rows);
    if (!headerInfo) return tableXml;

    let nextNC = 1;
    let nextOSS = 1;

    const newRows = rows.map((row, index) => {
      if (index <= headerInfo.headerRowIndex) return row;

      const cells = estraiCelle(row);
      if (cells.length <= headerInfo.codiceCol) return row;

      const codice = normalizzaCodice(estraiTesto(cells[headerInfo.codiceCol]));
      const match = codice.match(/^(NC|OSS)(\d+)(.*)$/i);

      if (!match) return row;

      const tipo = match[1].toUpperCase();
      const resto = match[3] || "";
      const nuovoNumero = tipo === "NC" ? nextNC++ : nextOSS++;
      const nuovoCodice = `${tipo}${nuovoNumero}${resto}`;

      if (nuovoCodice === codice) return row;

      const newCell = sostituisciTestoCella(cells[headerInfo.codiceCol], nuovoCodice);
      const newCells = [...cells];
      newCells[headerInfo.codiceCol] = newCell;

      rinumerate += 1;
      return ricostruisciRiga(row, newCells);
    });

    let idx = 0;
    return tableXml.replace(/<w:tr[\s\S]*?<\/w:tr>/g, () => newRows[idx++] || "");
  });

  return { xml: newXml, rinumerate };
}

function parseCodiceOrdinamento(codice: string) {
  const match = String(codice || "")
    .toUpperCase()
    .match(/^(NC|OSS)(\d+)/);

  if (!match) {
    return { tipo: "ZZZ", numero: 999999, ordinabile: false };
  }

  return {
    tipo: match[1],
    numero: Number(match[2]),
    ordinabile: true,
  };
}

function tipoSortOrder(tipo: string) {
  if (tipo === "NC") return 1;
  if (tipo === "OSS") return 2;
  return 99;
}

function estraiRecordsDaDocumentXml(
  xml: string,
  fileName: string
): RilievoRecord[] {
  const records: RilievoRecord[] = [];
  const tables = matchAll(xml, /<w:tbl[\s\S]*?<\/w:tbl>/g);

  for (const table of tables) {
    const rows = matchAll(table, /<w:tr[\s\S]*?<\/w:tr>/g);
    if (rows.length < 2) continue;

    const headerInfo = trovaColonneTabella(rows);
    if (!headerInfo) continue;

    for (let i = headerInfo.headerRowIndex + 1; i < rows.length; i += 1) {
      const cells = estraiCelle(rows[i]);

      if (cells.length <= Math.max(headerInfo.codiceCol, headerInfo.rilievoCol)) {
        continue;
      }

      const codice = normalizzaCodice(estraiTesto(cells[headerInfo.codiceCol]));
      const rilievo = pulisciTestoRilievo(estraiTesto(cells[headerInfo.rilievoCol]));
      const key = normalizeRilievoKey(rilievo);

      if (!key || !rilievo || !sembraCodiceNcOss(codice)) {
        continue;
      }

      records.push({
        file: fileName,
        rowIndex: i + 1,
        codice,
        rilievo,
        key,
      });
    }
  }

  return records;
}

function trovaColonneTabella(rows: string[]) {
  for (let rowIndex = 0; rowIndex < Math.min(rows.length, 8); rowIndex += 1) {
    const cells = estraiCelle(rows[rowIndex]);
    const texts = cells.map((cell) => normalizeHeader(estraiTesto(cell)));

    let codiceCol = texts.findIndex(
      (text) =>
        text.includes("CRONOLOGICO") ||
        text.includes("N CRONOLOGICO") ||
        text.includes("NUMERO CRONOLOGICO") ||
        text === "NCOSS" ||
        text.includes("NCOSS") ||
        text.includes("NC OSS") ||
        text.includes("CLASSIFICAZIONE") ||
        text.includes("TIPO RILIEVO") ||
        text === "CODICE"
    );

    let rilievoCol = texts.findIndex(
      (text) =>
        text.includes("RILIEVI ODI") ||
        text.includes("RILIEVO ODI") ||
        text.includes("RILIEVIODI") ||
        text.includes("RILIEVOODI") ||
        text.includes("RILIEVI ITS") ||
        text.includes("RILIEVO ITS") ||
        text.includes("ITS CONTROLLI TECNICI") ||
        text.includes("CONTROLLI TECNICI") ||
        text.includes("DESCRIZIONE")
    );

    if (codiceCol < 0 && cells.length > 0) {
      const righeDati = rows.slice(rowIndex + 1, rowIndex + 5);
      const primaColonnaContieneCodici = righeDati.some((r) => {
        const first = estraiCelle(r)[0];
        return first && sembraCodiceNcOss(estraiTesto(first));
      });

      if (primaColonnaContieneCodici) {
        codiceCol = 0;
      }
    }

    if (rilievoCol < 0 && cells.length >= 4) {
      rilievoCol = 3;
    }

    if (codiceCol >= 0 && rilievoCol >= 0 && codiceCol !== rilievoCol) {
      return { headerRowIndex: rowIndex, codiceCol, rilievoCol };
    }
  }

  return null;
}

function estraiCelle(rowXml: string) {
  return matchAll(rowXml, /<w:tc[\s\S]*?<\/w:tc>/g);
}

function ricostruisciRiga(rowXml: string, newCells: string[]) {
  let index = 0;
  return rowXml.replace(/<w:tc[\s\S]*?<\/w:tc>/g, () => {
    const cell = newCells[index];
    index += 1;
    return cell || "";
  });
}

function estraiTesto(xml: string) {
  return matchAll(xml, /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)
    .map((part) =>
      decodeXml(part.replace(/^<w:t(?:\s[^>]*)?>/, "").replace(/<\/w:t>$/, ""))
    )
    .join("")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function sostituisciTestoCella(cellXml: string, nuovoTesto: string) {
  let replaced = false;

  return cellXml.replace(/<w:t(\s[^>]*)?>([\s\S]*?)<\/w:t>/g, (match, attrs) => {
    if (!replaced) {
      replaced = true;
      return `<w:t${attrs || ""}>${escapeXml(nuovoTesto)}</w:t>`;
    }

    return match.replace(/>([\s\S]*?)<\/w:t>$/, `></w:t>`);
  });
}

function normalizzaCodice(value: string) {
  return String(value || "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, "")
    .replace(/^(NC|OSS)(\d+)/i, (_, tipo, numero) => {
      return `${String(tipo).toUpperCase()}${numero}`;
    })
    .trim();
}

function estraiProgressivo(codice: string) {
  const match = String(codice || "")
    .toUpperCase()
    .match(/^(NC|OSS)\s*(\d+)/);

  if (!match) return "";

  return `${match[1]}${match[2]}`;
}

function sembraCodiceNcOss(value: string) {
  return /^(NC|OSS)\s*\d+/i.test(String(value || "").trim());
}

function pulisciTestoRilievo(value: string) {
  return String(value || "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeRilievoKey(value: string) {
  return String(value || "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeHeader(value: string) {
  return normalizeRilievoKey(value);
}

function buildCsv(rows: ReportRow[]) {
  const header = [
    "file",
    "riga",
    "codice_precedente",
    "codice_originale",
    "codice_finale",
    "stato",
    "rilievo_odi",
  ];

  const lines = [header.join(";")];

  for (const row of rows) {
    lines.push(
      [
        row.file,
        String(row.riga),
        row.codice_precedente,
        row.codice_originale,
        row.codice_finale,
        row.stato,
        row.rilievo_odi,
      ]
        .map(csvCell)
        .join(";")
    );
  }

  return "\ufeff" + lines.join("\n");
}

function buildLog(stats: {
  totaleRighe: number;
  corretti: number;
  giaAllineati: number;
  mancanti: number;
  duplicati: number;
  riordinate: number;
  rinumerate: number;
}) {
  return [
    "Correzione numerazione NC/OSS tra emissioni",
    "Regola: l'emissione da correggere viene adeguata all'emissione precedente.",
    "Chiave di confronto: testo della colonna Rilievi ODI / Rilievi ITS Controlli Tecnici.",
    "Fallback: se il rilievo non coincide, usa il progressivo NC/OSS.",
    "Riordino: le righe Word vengono ordinate per progressivo NC e OSS.",
    "Rinumerazione finale: elimina duplicati e buchi nella progressione NC/OSS.",
    "",
    `Righe analizzate: ${stats.totaleRighe}`,
    `Numerazioni corrette da riferimento: ${stats.corretti}`,
    `Numerazioni gia allineate: ${stats.giaAllineati}`,
    `Rilievi non trovati o senza codice: ${stats.mancanti}`,
    `Rilievi duplicati nel riferimento: ${stats.duplicati}`,
    `Tabelle riordinate: ${stats.riordinate}`,
    `Progressivi rinumerati: ${stats.rinumerate}`,
    "",
  ].join("\n");
}

function csvCell(value: string) {
  const text = String(value || "");
  return `"${text.replace(/"/g, '""')}"`;
}

function matchAll(value: string, regex: RegExp) {
  return Array.from(value.matchAll(regex)).map((match) => match[0]);
}

function decodeXml(value: string) {
  return String(value || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function escapeXml(value: string) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
