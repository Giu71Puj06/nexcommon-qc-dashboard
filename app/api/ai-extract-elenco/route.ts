import { NextRequest, NextResponse } from "next/server";
import OpenAI from "openai";

export const runtime = "nodejs";

const client = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

export async function POST(request: NextRequest) {
  try {
    if (!process.env.OPENAI_API_KEY) {
      return NextResponse.json(
        {
          success: false,
          error: "OPENAI_API_KEY non configurata",
        },
        { status: 500 }
      );
    }

    const formData = await request.formData();

    const file = formData.get("file") as File | null;

    if (!file) {
      return NextResponse.json(
        {
          success: false,
          error: "Nessun file PDF caricato",
        },
        { status: 400 }
      );
    }

    const bytes = await file.arrayBuffer();

    const base64File = Buffer.from(bytes).toString("base64");

    const prompt = `
Analizza questo PDF "Elenco Elaborati".

Estrai TUTTI gli elaborati presenti.

Per ogni elaborato restituisci:
- codice elaborato
- titolo elaborato
- revisione
- disciplina
- formato documento

Rispondi SOLO in JSON valido.

Formato richiesto:

{
  "elaborati": [
    {
      "codice": "",
      "titolo": "",
      "revisione": "",
      "disciplina": "",
      "formato": ""
    }
  ]
}
`;

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

    const raw = response.output_text;

    let parsed;

    try {
      parsed = JSON.parse(raw);
    } catch {
      return NextResponse.json(
        {
          success: false,
          error: "Risposta AI non valida",
          raw,
        },
        { status: 500 }
      );
    }

    return NextResponse.json({
      success: true,
      elaborati: parsed.elaborati || [],
    });
  } catch (error) {
    return NextResponse.json(
      {
        success: false,
        error:
          error instanceof Error
            ? error.message
            : "Errore AI extract elenco",
      },
      { status: 500 }
    );
  }
}
