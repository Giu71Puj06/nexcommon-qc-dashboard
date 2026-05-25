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

FASE 1 - CLASSIFICAZIONE OBBLIGATORIA
Prima di applicare il Piano di Sorveglianza selezionato, devi classificare il documento ricevuto leggendo cartiglio, titolo, codice elaborato e contenuto.

Devi identificare:
- disciplina reale del documento
- tipo reale del documento
- pertinenza rispetto al Piano di Sorveglianza selezionato

Esempi:
- ESPROPRI / ELENCO DITTE non è pertinente al Piano Architettonico
- ELENCO ELABORATI non è pertinente a una verifica tecnica grafica
- COMPUTO METRICO non è pertinente a una verifica grafica architettonica
- TAVOLA ARCHITETTONICA è pertinente al Piano Architettonico
- RELAZIONE ANTINCENDIO è pertinente al Piano Antincendio
- PIANO DEMOLIZIONI è pertinente al Piano Demolizioni

REGOLA FONDAMENTALE
Se il documento NON è pertinente al Piano di Sorveglianza selezionato:
- NON generare rilievi tecnici inventati
- NON applicare checklist non coerenti
- restituisci esitoGenerale = "NA"
- inserisci un solo rilievo di tipo "NA"
- spiega che il documento non è pertinente al piano selezionato
- indica quale disciplina/tipo documento è stato rilevato

Devi analizzare l'elaborato ricevuto rispetto al Piano di Sorveglianza selezionato SOLO se il documento è pertinente.

NON inventare rilievi.
Se un'informazione non è leggibile o non verificabile, indicarlo chiaramente.

Classifica ogni rilievo come:
- OK = nessuna criticità evidente
- OSS = osservazione / approfondimento consigliato
- NC = non conformità evidente rispetto al punto di controllo
- NA = non applicabile o documento non pertinente

Rispondi SOLO in JSON valido.
Non usare markdown.
Non inserire testo prima o dopo il JSON.

Formato:
{
  "codiceElaborato": "",
  "titoloElaborato": "",
  "revisione": "",
  "disciplinaRilevata": "",
  "tipoElaborato": "",
  "pertinenteAlPiano": true,
  "esitoGenerale": "OK",
  "priorita": "BASSA",
  "rilievi": [
    {
      "puntoPiano": "",
      "tipo": "OK",
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

Verifica:
- piante, sezioni, prospetti
- quote planimetriche e altimetriche
- nord geografico
- destinazioni d'uso
- altezze locali
- accessibilità e barriere architettoniche
- infissi, abachi, scale, parapetti, rampe
- dettagli costruttivi
- coerenza grafica generale
`,

    antincendio: `
Piano di Sorveglianza: CHECKLIST ANTINCENDIO.

Verifica:
- destinazione d'uso
- attività soggetta
- compartimentazione
- vie di esodo
- uscite di emergenza
- resistenza/reazione al fuoco
- impianti antincendio
- rivelazione e allarme
- controllo fumi e calore
- accessibilità mezzi di soccorso
- segnaletica e illuminazione emergenza
`,

    demolizioni: `
Piano di Sorveglianza: DEMOLIZIONI.

Verifica:
- stato attuale e stato di progetto
- indagini strutturali preliminari
- modalità esecutive demolizione
- stabilità e puntellamenti
- gestione polveri, rumore, vibrazioni
- materiali pericolosi e smaltimento
- layout area demolizione
- fotografie e rilievi
- ponti di servizio e sicurezza lavoratori
`,

    art44bis: `
Piano di Sorveglianza: RELAZIONE A CORREDO ART. 44 BIS.

Verifica:
- completezza documentale
- protocollo di intesa
- nota di trasmissione
- relazione a corredo
- quadro conoscitivo
- coerenza con norme vigenti
- cantierizzazione
- manutenibilità opere
- autorizzazioni
- articolazione secondo le quattro tematiche richieste
`,

    "pfte-pnrr": `
Piano di Sorveglianza: PFTE PNRR.

Verifica:
- relazione generale
- relazione tecnica
- elaborati grafici
- quadro economico
- cronoprogramma
- sostenibilità opera
- DNSH
- studio impatto ambientale se pertinente
- piano preliminare manutenzione
- completezza rispetto linee guida PFTE PNRR
`,

    generico: `
Piano di Sorveglianza generico.

Verifica:
- completezza elaborato
- coerenza cartiglio
- leggibilità
- contenuti tecnici minimi
- anomalie evidenti
`,
  };

  return `${common}\n\n${plans[mode] || plans.generico}`;
}

function cleanJsonText(value: string) {
  return String(value || "")
    .trim()
    .replace(/^```json/i, "")
    .replace(/^```/i, "")
    .replace(/```$/i, "")
    .trim();
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
      ((formData.get("mode") as SurveillanceMode) ||
        "generico") as SurveillanceMode;

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
      model: "gpt-4o-mini",
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

    const raw = cleanJsonText(response.output_text || "");

    console.log("RAW AI RESPONSE:");
    console.log(raw);

    let parsed;

    try {
      parsed = JSON.parse(raw);
    } catch (err) {
      console.error("JSON PARSE ERROR:", err);

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
      result: parsed,
    });
  } catch (error) {
    console.error("AI SURVEILLANCE CHECK ERROR:", error);

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
