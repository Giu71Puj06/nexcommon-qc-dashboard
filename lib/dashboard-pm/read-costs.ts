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

export async function readEconomicFile(
  file: File
): Promise<EconomicRevision | null> {
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

  if (amount === 0) return null;

  return {
    projectName: file.name
      .replace(".xlsx", "")
      .replace(/\d{12}\+\d{4}$/i, "")
      .trim(),

    revisionName: file.name,
    amount,
  };
}
