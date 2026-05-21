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
    const fase = (formData.get("fase") || "iniziale") as "iniziale" | "finale";

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
        const pdfParse = (await import("pdf-parse")).default;
        const parsed = await pdfParse(buffer);
        text = parsed.text || "";
      } catch {
        text = "";
      }

      const commessa = extractCommessa(text, file.name);
      const importo = extractImporto(text);

      results.push({
        fileName: file.name,
        commessa,
        importo,
        fase,
        testoEstratto: text.slice(0, 3000),
        warning:
          importo <= 0
            ? "Importo non rilevato automaticamente. Verifica il PDF."
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
    /commessa[:\s]+([A-Z0-9À-ÿ\s._/-]{5,120})/i,
    /oggetto[:\s]+([A-Z0-9À-ÿ\s._/-]{5,160})/i,
    /intervento[:\s]+([A-Z0-9À-ÿ\s._/-]{5,160})/i,
    /progetto[:\s]+([A-Z0-9À-ÿ\s._/-]{5,160})/i,
    /cup[:\s]+([A-Z0-9]{10,20})/i,
  ];

  for (const pattern of patterns) {
    const match = cleanText.match(pattern);
    if (match?.[1]) {
      return cleanupLabel(match[1]);
    }
  }

  return fileName
    .replace(/\.pdf$/i, "")
    .replace(/[_-]+/g, " ")
    .trim();
}

function extractImporto(text: string): number {
  const cleanText = normalizeText(text);

  const priorityPatterns = [
    /totale\s+(?:complessivo|generale|quadro\s+economico|computo|intervento)[^\d€]{0,80}€?\s*([0-9.\s]+,[0-9]{2})/i,
    /importo\s+(?:complessivo|totale|lavori|progetto)[^\d€]{0,80}€?\s*([0-9.\s]+,[0-9]{2})/i,
    /€\s*([0-9.\s]+,[0-9]{2})/i,
  ];

  for (const pattern of priorityPatterns) {
    const match = cleanText.match(pattern);
    if (match?.[1]) {
      return parseEuro(match[1]);
    }
  }

  const allAmounts = Array.from(
    cleanText.matchAll(/([0-9]{1,3}(?:[.\s][0-9]{3})+,[0-9]{2})/g)
  )
    .map((match) => parseEuro(match[1]))
    .filter((value) => Number.isFinite(value) && value > 0);

  if (!allAmounts.length) return 0;

  return Math.max(...allAmounts);
}

function parseEuro(value: string): number {
  const normalized = value
    .replace(/\s/g, "")
    .replace(/\./g, "")
    .replace(",", ".");

  const parsed = Number(normalized);

  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeText(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function cleanupLabel(value: string): string {
  return value
    .replace(/\s{2,}/g, " ")
    .replace(/(?:fase|data|importo|totale|cup).*$/i, "")
    .trim()
    .slice(0, 140);
}
