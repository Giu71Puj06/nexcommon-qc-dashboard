import * as XLSX from "xlsx";

export type EconomicRevision = {
  projectName: string;
  revisionName: string;
  amount: number;
};

function normalize(value: unknown): number {
  if (typeof value === "number") return value;

  return Number(
    String(value || "0")
      .replace(/\./g, "")
      .replace(",", ".")
      .replace(/[^\d.-]/g, "")
  );
}

function cleanProjectName(fileName: string): string {
  return fileName
    .replace(/\.xlsx$/i, "")
    .replace(/\.pdf$/i, "")
    .replace(/\d{12}\+\d{4}$/i, "")
    .trim();
}

function extractAmountFromText(text: string): number {
  const normalizedText = text.replace(/\s+/g, " ");

  const patterns = [
    /(?:importo|totale|costo|quadro economico)[^\d€]{0,80}(?:€\s*)?([\d.]+,\d{2})/gi,
    /(?:importo|totale|costo|quadro economico)[^\d€]{0,80}(?:€\s*)?([\d.]+)/gi,
  ];

  let max = 0;

  for (const pattern of patterns) {
    let match;

    while ((match = pattern.exec(normalizedText)) !== null) {
      const value = normalize(match[1]);

      if (value > max) {
        max = value;
      }
    }
  }

  return max;
}

async function readPdfText(file: File): Promise<string> {
  const pdfjs = await import("pdfjs-dist");

  const buffer = await file.arrayBuffer();

  const loadingTask = pdfjs.getDocument({
    data: buffer,
    useWorkerFetch: false,
    disableFontFace: true,
  });

  const pdf = await loadingTask.promise;

  let text = "";

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();

    text +=
      content.items
        .map((item: any) => ("str" in item ? item.str : ""))
        .join(" ") + " ";
  }

  return text;
}

async function readXlsxAmount(file: File): Promise<number> {
  const buffer = await file.arrayBuffer();

  const workbook = XLSX.read(buffer, {
    type: "array",
  });

  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet);

  let amount = 0;

  for (const row of rows) {
    for (const key of Object.keys(row)) {
      const label = key.toLowerCase();

      if (
        label.includes("importo") ||
        label.includes("totale") ||
        label.includes("quadro economico") ||
        label.includes("costo")
      ) {
        const value = normalize(row[key]);

        if (value > amount) {
          amount = value;
        }
      }
    }
  }

  return amount;
}

export async function readEconomicFile(
  file: File
): Promise<EconomicRevision | null> {
  const lowerName = file.name.toLowerCase();

  let amount = 0;

  if (lowerName.endsWith(".xlsx")) {
    amount = await readXlsxAmount(file);
  }

  if (lowerName.endsWith(".pdf")) {
    const text = await readPdfText(file);
    amount = extractAmountFromText(text);
  }

  if (amount === 0) return null;

  return {
    projectName: cleanProjectName(file.name),
    revisionName: file.name,
    amount,
  };
}
