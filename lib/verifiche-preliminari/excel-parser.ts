import * as XLSX from "xlsx";

export interface Elaborato {
  codice: string;
  titolo: string;
  disciplina?: string;
}

export async function parseElencoElaborati(
  file: File
): Promise<Elaborato[]> {

  const buffer = await file.arrayBuffer();

  const workbook = XLSX.read(buffer, {
    type: "array"
  });

  const sheetName = workbook.SheetNames[0];

  const sheet = workbook.Sheets[sheetName];

  const jsonData = XLSX.utils.sheet_to_json<any>(sheet);

  return jsonData.map((row: any) => ({

    codice:
      row["Codice"] ||
      row["CODICE"] ||
      row["codice"] ||
      "",

    titolo:
      row["Titolo"] ||
      row["TITOLO"] ||
      row["Descrizione"] ||
      row["descrizione"] ||
      "",

    disciplina:
      row["Disciplina"] ||
      row["DISCIPLINA"] ||
      ""

  }));

}