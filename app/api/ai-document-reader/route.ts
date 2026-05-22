import { NextRequest, NextResponse } from "next/server";
import OpenAI from "openai";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type ReaderMode =
  | "economic-analysis"
  | "document-reception-check"
  | "generic-document-reading";

export async function GET() {
  return NextResponse.json({
    success: true,
    message: "AI Document Reader attivo. Usa POST con file e mode.",
  });
}

export async function POST(request: NextRequest) {
  try {
    const apiKey = process.env.OPENAI_API_KEY;

    if (!apiKey) {
      return NextResponse.json(
        { error: "OPENAI_API_KEY non configurata su Railway." },
        { status: 500 }
      );
    }

    const client = new OpenAI({
      apiKey,
    });

    const formData = await request.formData();
    const file = formData.get("file") as File | null;

    const mode =
      (formData.get("mode") as ReaderMode) || "generic-document-reading";

    if (!file) {
      return NextResponse.json(
        { error: "Nessun file caricato." },
        { status: 400 }
      );
    }

    const bytes = await file.arrayBuffer();
    const base64File = Buffer.from(bytes).toString("base64");

    const prompt = buildPrompt(mode, file.name);

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
              file_data: `data:${file.type || "application/pdf"};base64,${base64File}`,
            },
          ],
        },
      ],
    });

    return NextResponse.json({
      success: true,
      mode,
      fileName: file.name,
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

function buildPrompt(mode: ReaderMode, fileName: string): string {
  if (mode === "economic-analysis") {
    return `
Sei un assistente tecnico esperto in computi metrici italiani, quadri economici, cartigli e documentazione di appalto.

Analizza il PDF allegato.

Nome file:
${fileName}

Estrai:
- nome commessa
- codice elaborato
- titolo elaborato
- revisione
- tipo documento
- importo totale
- se è computo metrico estimativo
- se è computo metrico senza prezzi
- eventuali incongruenze tra nome file e cartiglio

Regole:
- Rispondi SOLO in JSON valido.
- Non usare markdown.
- Non aggiungere testo fuori dal JSON.
- Se trovi "TOTALE euro", quello è l'importo totale principale.
- Nei computi metrici estimativi italiani il totale finale può trovarsi nelle ultime pagine.
- Non confondere quantità, prezzi unitari, riporti o subtotali con importo totale.
- Se il documento è un computo metrico senza prezzi, importoTotale deve essere 0.

Formato JSON:
{
  "commessa": "",
  "codiceElaborato": "",
  "titoloElaborato": "",
  "revisione": "",
  "tipoDocumento": "",
  "importoTotale": 0,
  "computoEstimativo": false,
  "computoSenzaPrezzi": false,
  "coerenzaNomeFileCartiglio": true,
  "warning": "",
  "confidenza": 0
}
`;
  }

  if (mode === "document-reception-check") {
    return `
Sei un assistente tecnico esperto in controllo elaborati progettuali, cartigli, codici documento e revisioni.

Analizza il PDF allegato.

Nome file:
${fileName}

Verifica:
- codice elaborato nel nome file
- codice elaborato nel cartiglio
- titolo elaborato
- revisione
- data revisione
- disciplina
- fase progettuale
- coerenza tra nome file e cartiglio

Rispondi SOLO in JSON valido.

Formato JSON:
{
  "commessa": "",
  "codiceDocumentoFile": "",
  "codiceDocumentoCartiglio": "",
  "titoloElaborato": "",
  "revisioneFile": "",
  "revisioneCartiglio": "",
  "dataRevisione": "",
  "disciplina": "",
  "faseProgettuale": "",
  "coerenze": {
    "codiceCoerente": true,
    "revisioneCoerente": true,
    "titoloPresente": true,
    "cartiglioLeggibile": true
  },
  "incoerenze": [],
  "azioniConsigliate": [],
  "confidenza": 0
}
`;
  }

  return `
Leggi il PDF allegato e restituisci SOLO JSON valido.

Nome file:
${fileName}

Formato JSON:
{
  "tipoDocumento": "",
  "commessa": "",
  "codiceDocumento": "",
  "titoloElaborato": "",
  "revisione": "",
  "sintesi": "",
  "datiEstratti": {},
  "confidenza": 0
}
`;
}
