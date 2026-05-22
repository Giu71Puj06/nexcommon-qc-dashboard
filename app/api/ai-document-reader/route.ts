import { NextRequest, NextResponse } from "next/server";
import OpenAI from "openai";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type ReaderMode =
  | "economic-analysis"
  | "document-reception-check"
  | "document-list-extraction"
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

    const client = new OpenAI({ apiKey });

    const formData = await request.formData();
    const file = formData.get("file") as File | null;
    const elencoInfoRaw = formData.get("elencoInfo");

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

    const elencoInfo =
      typeof elencoInfoRaw === "string" && elencoInfoRaw
        ? elencoInfoRaw
        : "";

    const prompt = buildPrompt(mode, file.name, elencoInfo);

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

function buildPrompt(mode: ReaderMode, fileName: string, elencoInfo = ""): string {
  if (mode === "document-list-extraction") {
    return `
Sei un assistente tecnico esperto in elenchi elaborati progettuali italiani.

Analizza SOLO il PDF allegato, che è un elenco elaborati.

Nome file elenco:
${fileName}

Estrai tutte le righe dell'elenco elaborati.

Per ogni elaborato estrai:
- codice elaborato
- revisione
- titolo elaborato
- disciplina
- formato, se presente
- scala, se presente
- data, se presente

Regole:
- Rispondi SOLO in JSON valido.
- Non usare markdown.
- Non aggiungere testo fuori dal JSON.
- Se un dato non è presente, usa stringa vuota.
- Non inventare dati.
- Mantieni i codici esattamente come compaiono nel documento.

Formato JSON:
{
  "tipoDocumento": "elenco_elaborati",
  "fileName": "${fileName}",
  "commessa": "",
  "elaborati": [
    {
      "codice": "",
      "revisione": "",
      "titolo": "",
      "disciplina": "",
      "formato": "",
      "scala": "",
      "data": ""
    }
  ],
  "warning": "",
  "confidenza": 0
}
`;
  }

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
Sei un assistente tecnico esperto in controllo elaborati progettuali, cartigli, codici documento, revisioni e formati.

Analizza SOLO il PDF allegato.

Nome file PDF da verificare:
${fileName}

Informazioni leggere disponibili dal frontend:
${elencoInfo || "Nessuna informazione aggiuntiva."}

Estrai dal PDF:
- codice elaborato nel cartiglio
- titolo elaborato nel cartiglio
- revisione nel cartiglio
- data revisione
- disciplina
- fase progettuale
- formato documento o formato tavola, se presente
- commessa, se presente

Regole:
- Rispondi SOLO in JSON valido.
- Non usare markdown.
- Non aggiungere testo fuori dal JSON.
- Non leggere né confrontare altri PDF.
- Non inventare dati.
- Il confronto con l'elenco elaborati verrà fatto dal frontend in TypeScript.
- Qui devi solo leggere il cartiglio e i dati del singolo PDF.

Formato JSON:
{
  "commessa": "",
  "codiceDocumentoFile": "${fileName.replace(/\.pdf$/i, "")}",
  "codiceDocumentoCartiglio": "",
  "titoloElaborato": "",
  "revisioneFile": "",
  "revisioneCartiglio": "",
  "dataRevisione": "",
  "disciplina": "",
  "faseProgettuale": "",
  "formatoDocumento": "",
  "formatoCartiglio": "",
  "coerenze": {
    "cartiglioLeggibile": true,
    "titoloPresente": true
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
