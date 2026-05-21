import { NextRequest, NextResponse } from "next/server";

export const runtime = "nodejs";

type ExtractedEconomicFile = {
  fileName: string;
  commessa: string;
  importo: number;
  fase: "iniziale" | "finale";
  testoEstratto: string;
  warning?: string;
};

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();

    const files = formData.getAll("files") as File[];

    const fase = (formData.get("fase") ||
      "iniziale") as "iniziale" | "finale";

    if (!files.length) {
      return NextResponse.json(
        { error: "Nessun file PDF caricato." },
        { status: 400 }
      );
    }

    const results: ExtractedEconomicFile[] = [];

    for (const file of files) {
      const buffer = Buffer.from(await file.arrayBuffer());

      let text = "";

      try {
        // eslint-disable-next-line @typescript-eslint/no-var-requires
        const pdfParse = require("pdf-parse");

        const parsed = await pdfParse(buffer);

        text = parsed.text || "";
      } catch {
        text = "";
      }

      const commessa = extractCommessa(text, file.name);

      const importo = extractImporto(text);

      const isEstimativo = isComputoEstimativo(text);

      results.push({
        fileName: file.name,
        commessa,
        importo,
        fase,
        testoEstratto: text.slice(0, 5000),
        warning:
          !isEstimativo
            ? "Documento metrico senza prezzi: importo non calcolato."
            : importo <= 0
            ? "Importo totale non rilevato automaticamente. Verifica il PDF."
            : undefined,
      });
    }

    return NextResponse.json({ files: results });
  } catch (error) {
    return NextResponse.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "Errore durante l'analisi dei PDF economici.",
      },
      { status: 500 }
    );
  }
}

function extractCommessa(text: string, fileName: string): string {
  const cleanText = normalizeText(text);

  const patterns = [
    /SOTTOVIA[\s\S]{0,120}?A13/i,

    /INTERVENTO[\s\S]{0,150}?RISANAMENTO[\s\S]{0,120}?EVOLUTIVO/i,

    /PROGETTO[\s:]+([A-Z0-9À-ÿ\s._/-]{10,180})/i,

    /OGGETTO[\s:]+([A-Z0-9À-ÿ\s._/-]{10,180})/i,

    /OPERA[\s:]+([A-Z0-9À-ÿ\s._/-]{10,180})/i,
  ];

  for (const pattern of patterns) {
    const match = cleanText.match(pattern);

    if (match?.[0]) {
      return cleanupLabel(match[0]);
    }

    if (match?.[1]) {
      return cleanupLabel(match[1]);
    }
  }

  const codiceMatch = cleanText.match(
    /(PV[0-9]{3}-[A-Z]{2}-[A-Z]{4}-[A-Z]{3}-[0-9]{5}-[A-Z]{3}-[0-9]{6})/i
  );

  if (codiceMatch?.[1]) {
    return codiceMatch[1].trim();
  }

  return fileName
    .replace(/\.pdf$/i, "")
    .replace(/[_-]+/g, " ")
    .trim();
}

function extractImporto(text: string): number {
  const cleanText = normalizeText(text);

  if (!isComputoEstimativo(cleanText)) {
    return 0;
  }

  const foundAmounts: number[] = [];

  const patterns = [
    /T\s*O\s*T\s*A\s*L\s*E[\s\S]{0,200}?([0-9]{1,3}(?:[.'´\s][0-9]{3})*,[0-9]{2})/gi,

    /Totale\s+Generale[\s\S]{0,200}?([0-9]{1,3}(?:[.'´\s][0-9]{3})*,[0-9]{2})/gi,

    /Totale\s+Complessivo[\s\S]{0,200}?([0-9]{1,3}(?:[.'´\s][0-9]{3})*,[0-9]{2})/gi,

    /Importo\s+Totale[\s\S]{0,200}?([0-9]{1,3}(?:[.'´\s][0-9]{3})*,[0-9]{2})/gi,

    /Totale\s+Super\s+Categorie[\s\S]{0,200}?([0-9]{1,3}(?:[.'´\s][0-9]{3})*,[0-9]{2})/gi,

    /Totale\s+Categorie[\s\S]{0,200}?([0-9]{1,3}(?:[.'´\s][0-9]{3})*,[0-9]{2})/gi,

    /Totale\s+Sub\s+Categorie[\s\S]{0,200}?([0-9]{1,3}(?:[.'´\s][0-9]{3})*,[0-9]{2})/gi,

    /Parziale\s+LAVORI\s+A\s+MISURA[\s\S]{0,200}?([0-9]{1,3}(?:[.'´\s][0-9]{3})*,[0-9]{2})/gi,

    /Parziale\s+LAVORI\s+A\s+CORPO[\s\S]{0,200}?([0-9]{1,3}(?:[.'´\s][0-9]{3})*,[0-9]{2})/gi,

    /euro[\s:]{0,20}([0-9]{1,3}(?:[.'´\s][0-9]{3})*,[0-9]{2})/gi,
  ];

  for (const pattern of patterns) {
    const matches = Array.from(cleanText.matchAll(pattern));

    for (const match of matches) {
      if (!match?.[1]) continue;

      const value = parseEuro(match[1]);

      if (
        Number.isFinite(value) &&
        value > 10000 &&
        value < 1000000000
      ) {
        foundAmounts.push(value);
      }
    }
  }

  if (!foundAmounts.length) {
    return 0;
  }

  const sorted = foundAmounts.sort((a, b) => b - a);

  return sorted[0];
}

function isComputoEstimativo(text: string): boolean {
  const cleanText = normalizeText(text).toUpperCase();

  return (
    cleanText.includes("COMPUTO METRICO ESTIMATIVO") ||
    cleanText.includes("TOTALE EURO") ||
    cleanText.includes("T O T A L E") ||
    cleanText.includes("IMPORTO TOTALE") ||
    cleanText.includes("TOTALE GENERALE") ||
    cleanText.includes("TOTALE CATEGORIE") ||
    cleanText.includes("TOTALE SUB CATEGORIE") ||
    cleanText.includes("TOTALE SUPER CATEGORIE")
  );
}

function parseEuro(value: string): number {
  const normalized = value
    .replace(/\s/g, "")
    .replace(/'/g, "")
    .replace(/´/g, "")
    .replace(/\./g, "")
    .replace(",", ".");

  const parsed = Number(normalized);

  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeText(value: string): string {
  return value
    .replace(/\r/g, " ")
    .replace(/\n/g, " ")
    .replace(/\t/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function cleanupLabel(value: string): string {
  return value
    .replace(/\s{2,}/g, " ")
    .replace(/TOTALE.*$/i, "")
    .replace(/IMPORTO.*$/i, "")
    .replace(/CATEGORIE.*$/i, "")
    .trim()
    .slice(0, 180);
}
