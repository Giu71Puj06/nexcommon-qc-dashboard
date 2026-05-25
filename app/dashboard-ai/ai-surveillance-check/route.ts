import { NextRequest, NextResponse } from "next/server";
import OpenAI from "openai";

export const runtime = "nodejs";

type SurveillanceMode =
  | "architettonico"
  | "antincendio"
  | "demolizioni"
  | "art44bis"
  | "pfte-pnrr"
  | "generico";

function getSurveillancePrompt(mode: SurveillanceMode) {
  const common = `
Agisci come assistente tecnico per ispettori ITS.

Devi analizzare l'elaborato ricevuto rispetto al Piano di Sorveglianza selezionato.

NON inventare rilievi.
Se un'informazione non è leggibile o non verificabile, indicarlo chiaramente.
Classifica ogni rilievo come:
- OK = nessuna criticità evidente
- OSS = osservazione / approfondimento consigliato
- NC = non conformità evidente rispetto al punto di controllo
- NA = non applicabile

Rispondi SOLO in JSON valido.

Formato:
{
  "codiceElaborato": "",
  "titoloElaborato": "",
  "revisione": "",
  "disciplinaRilevata": "",
  "tipoElaborato": "",
  "esitoGenerale": "OK | OSS | NC | NA",
  "priorita": "BASSA | MEDIA | ALTA",
  "rilievi": [
    {
      "puntoPiano": "",
      "tipo": "OK | OSS | NC | NA",
      "descrizioneRilievo": "",
      "azioneRichiesta": "",
      "evidenzaNelDocumento": ""
    }
  ],
  "sintesiIspettore": ""
}
`;

  const plans: Record<SurveillanceMode, string> = {
    architettonico: `
Piano di Sorveglianza: PROGETTO ESECUTIVO EDIFICI - ARCHITETTURA.

Verifica in particolare:
- presenza e leggibilità di piante, sezioni, prospetti
- quote planimetriche e altimetriche
- indicazione nord geografico
- destinazioni d'uso degli ambienti
- altezze locali
- accessibilità e barriere architettoniche
- infissi, abachi, scale, parapetti, rampe
- dettagli costruttivi
- congruenza con elaborati impiantistici, strutturali e antincendio ove evidente
- presenza riferimenti a particolari costruttivi
- coerenza grafica generale
`,

    antincendio: `
Piano di Sorveglianza: CHECKLIST ANTINCENDIO.

Verifica in particolare:
- destinazione d'uso e attività soggetta
- compartimentazione
- vie di esodo
- uscite di emergenza
- resistenza/reazione al fuoco
- impianti antincendio
- rivelazione e allarme
- controllo fumi e calore
- accessibilità mezzi di soccorso
- segnaletica e illuminazione di emergenza
- presenza documentazione SCIA/CPI se richiamata
`,

    demolizioni: `
Piano di Sorveglianza: DEMOLIZIONI.

Verifica in particolare:
- analisi stato attuale e stato di progetto
- indagini strutturali preliminari
- modalità esecutive di demolizione
- gestione stabilità e puntellamenti
- divieto getto materiale dall'alto
- gestione polveri, rumore e vibrazioni
- materiali pericolosi e smaltimento
- layout area di demolizione
- fotografie e rilievi
- ponti di servizio e sicurezza lavoratori
`,

    art44bis: `
Piano di Sorveglianza: RELAZIONE A CORREDO ART. 44 BIS.

Verifica in particolare:
- completezza documentale
- protocollo di intesa
- nota di trasmissione
- relazione a corredo
- quadro conoscitivo
- coerenza con norme vigenti
- cantierizzazione
- manutenibilità opere
- indicazione autorizzazioni
- articolazione secondo le quattro tematiche richieste
`,

    "pfte-pnrr": `
Piano di Sorveglianza: PFTE PNRR.

Verifica in particolare:
- relazione generale
- relazione tecnica
- elaborati grafici
- quadro economico
- cronoprogramma
- sostenibilità opera
- DNSH
- studio impatto ambientale se pertinente
- piano preliminare manutenzione
- completezza rispetto alle linee guida PFTE PNRR
`,

    generico: `
Piano di Sorveglianza generico.

Verifica:
- completezza elaborato
- coerenza cartiglio
- leggibilità
- presenza contenuti tecnici minimi
- eventuali anomalie evidenti
`,
  };

  return `${common}\n\n${plans[mode] || plans.generico}`;
}

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

    const client = new OpenAI({
      apiKey: process.env.OPENAI_API_KEY,
    });

    const formData = await request.formData();

    const file = formData.get("file") as File | null;
    const mode =
      ((formData.get("mode") as SurveillanceMode) || "generico") as SurveillanceMode;

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

    const prompt = getSurveillancePrompt(mode);

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
          error: "Risposta AI non valida come JSON",
          raw,
        },
        { status: 500 }
      );
    }

    return NextResponse.json({
      success: true,
      result: parsed,
    });
  } catch (error) {
    return NextResponse.json(
      {
        success: false,
        error:
          error instanceof Error
            ? error.message
            : "Errore durante verifica elaborato AI",
      },
      { status: 500 }
    );
  }
}
