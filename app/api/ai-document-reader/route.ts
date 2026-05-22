import { NextRequest, NextResponse } from "next/server";
import OpenAI from "openai";

export const runtime = "nodejs";

type ReaderMode =
  | "economic-analysis"
  | "document-reception-check"
  | "generic-document-reading";

export async function POST(request: NextRequest) {
  try {
    if (!process.env.OPENAI_API_KEY) {
      return NextResponse.json(
        {
          error: "OPENAI_API_KEY non configurata",
        },
        { status: 500 }
      );
    }

    const client = new OpenAI({
      apiKey: process.env.OPENAI_API_KEY,
    });

    const formData = await request.formData();

    const file = formData.get("file") as File | null;

    const mode =
      (formData.get("mode") as ReaderMode) ||
      "generic-document-reading";

    if (!file) {
      return NextResponse.json(
        { error: "Nessun file caricato" },
        { status: 400 }
      );
    }

    const bytes = await file.arrayBuffer();

    const base64File = Buffer.from(bytes).toString("base64");

    let prompt = "";

    if (mode === "economic-analysis") {
      prompt = `
Analizza questo PDF tecnico/economico.

Estrai:
- nome commessa
- codice elaborato
- importo totale
- se è computo metrico estimativo
- se è computo senza prezzi
- eventuali incongruenze

Rispondi SOLO in JSON valido.

Formato:
{
  "commessa": "",
  "codiceElaborato": "",
  "importoTotale": 0,
  "tipoDocumento": "",
  "warning": ""
}
`;
    }

    if (mode === "document-reception-check") {
      prompt = `
Analizza il PDF tecnico.

Verifica:
- codice elaborato
- titolo elaborato
- coerenza cartiglio
- presenza revisione
- presenza data
- eventuali incongruenze

Rispondi SOLO in JSON valido.
`;
    }

    if (mode === "generic-document-reading") {
      prompt = `
Leggi il documento PDF e riassumi i contenuti principali.
`;
    }

    const response = await client.responses.create({
      model: "gpt-4.1-mini",
      input: [
        {
          role: "user",
          content: [
            {
              type: "input_text",
              text: prompt,
            },
            {
              type: "input_file",
              filename: file.name,
              file_data: `data:application/pdf;base64,${base64File}`,
            },
          ],
        },
      ],
    });

    return NextResponse.json({
      success: true,
      result: response.output_text,
    });
  } catch (error) {
    return NextResponse.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "Errore AI document reader",
      },
      { status: 500 }
    );
  }
}
