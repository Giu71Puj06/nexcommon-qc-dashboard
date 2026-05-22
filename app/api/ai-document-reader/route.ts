import { NextRequest, NextResponse } from "next/server";
import OpenAI from "openai";

export const runtime = "nodejs";

const client = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

type ReaderMode =
  | "economic-analysis"
  | "document-reception-check"
  | "generic-document-reading";

export async function POST(request: NextRequest) {
  try {
    if (!process.env.OPENAI_API_KEY) {
      return NextResponse.json(
        { error: "OPENAI_API_KEY non configurata su Railway." },
        { status: 500 }
      );
    }

    const formData = await request.formData();
    const mode = String(
      formData.get("mode") || "generic-document-reading"
    ) as ReaderMode;

    const files = formData.getAll("files") as File[];

    if (!files.length) {
      return NextResponse.json(
        { error: "Nessun file caricato." },
        { status: 400 }
      );
    }

    const results = [];

    for (const file of files) {
      const arrayBuffer = await file.arrayBuffer();
      const base64 = Buffer.from(arrayBuffer).toString("base64");

      const response = await client.responses.create({
        model: "gpt-4.1-mini",
        input: [
          {
            role: "user",
            content: [
              {
                type: "input_text",
                text: buildPrompt(mode, file.name),
              },
              {
                type: "input_file",
                filename: file.name,
                file_data: `data:${file.type || "application/pdf"};base64,${base64}`,
              },
            ],
          },
        ],
        text: {
          format: {
            type: "json_object",
          },
        },
      });

      const outputText = response.output_text || "{}";

      let parsed;

      try {
        parsed = JSON.parse(outputText);
      } catch {
        parsed = {
          error: "Risposta AI non valida.",
          raw: outputText,
        };
      }

      results.push({
        fileName: file.name,
        mode,
        result: parsed,
      });
    }

    return NextResponse.json({ files: results });
  } catch (error) {
    return NextResponse.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "Errore durante la lettura AI del documento.",
      },
      { status: 500 }
    );
  }
}

function buildPrompt(mode: ReaderMode, fileName: string): string {
  const common = `
Sei un assistente tecnico esperto in documentazione di ingegneria, appalti,
computi metrici, cartigli, elaborati progettuali e verifiche QA/QC.

Analizza il PDF allegato.
Devi restituire SOLO JSON valido.
Non aggiungere markdown.
Non aggiungere testo fuori dal JSON.

Nome file:
${fileName}
`;

  if (mode === "economic-analysis") {
    return `
${common}

Obiettivo:
estrarre dati economici da computi metrici estimativi, quadri economici,
analisi prezzi, SAL, perizie o documenti economici.

Restituisci questo JSON:

{
  "tipo_documento": "",
  "commessa": "",
  "codice_documento": "",
  "titolo_elaborato": "",
  "revisione": "",
  "fase_progetto": "",
  "importo_totale": 0,
  "importo_lavori_a_misura": 0,
  "importo_lavori_a_corpo": 0,
  "oneri_sicurezza": 0,
  "categorie": [
    {
      "codice": "",
      "descrizione": "",
      "importo": 0
    }
  ],
  "valuta": "EUR",
  "confidenza": 0,
  "note": "",
  "coerenze": {
    "codice_file_coerente_con_cartiglio": true,
    "revisione_file_coerente_con_cartiglio": true,
    "titolo_rilevato": true,
    "importo_rilevato": true
  }
}

Regole:
- Se trovi "TOTALE euro", quello è l'importo totale principale.
- Nei computi metrici estimativi italiani il totale finale può trovarsi nelle ultime pagine.
- Non confondere quantità, riporti, prezzi unitari o subtotali con importo totale.
- Se il documento è un computo metrico senza prezzi, importo_totale deve essere 0.
- Usa numeri puri, senza simbolo euro e senza separatori migliaia.
`;
  }

  if (mode === "document-reception-check") {
    return `
${common}

Obiettivo:
controllare ricezione elaborati, cartigli, codici, revisioni e coerenza tra nome file e contenuto PDF.

Restituisci questo JSON:

{
  "tipo_documento": "",
  "commessa": "",
  "codice_documento_cartiglio": "",
  "codice_documento_file": "",
  "titolo_elaborato": "",
  "disciplina": "",
  "fase_progetto": "",
  "revisione_cartiglio": "",
  "revisione_file": "",
  "data_revisione": "",
  "stato": "",
  "coerenze": {
    "codice_file_coerente_con_cartiglio": true,
    "revisione_file_coerente_con_cartiglio": true,
    "titolo_presente": true,
    "cartiglio_leggibile": true,
    "disciplina_coerente": true
  },
  "incoerenze": [],
  "azioni_consigliate": [],
  "confidenza": 0
}

Regole:
- Confronta codice del nome file con codice nel cartiglio.
- Confronta revisione nel nome file con revisione nel cartiglio.
- Se ci sono discrepanze, inseriscile in "incoerenze".
- Non inventare dati non presenti.
`;
  }

  return `
${common}

Restituisci questo JSON:

{
  "tipo_documento": "",
  "commessa": "",
  "codice_documento": "",
  "titolo_elaborato": "",
  "revisione": "",
  "fase_progetto": "",
  "dati_estratti": {},
  "note": "",
  "confidenza": 0
}
`;
}
