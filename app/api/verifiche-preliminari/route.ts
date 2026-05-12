import { NextRequest, NextResponse } from "next/server";
import { parseElencoElaborati } from "@/lib/verifiche-preliminari/excel-parser";

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();

    const file = formData.get("file");

    if (!file || !(file instanceof File)) {
      return NextResponse.json(
        {
          success: false,
          error: "Nessun file Excel ricevuto"
        },
        { status: 400 }
      );
    }

    const elaborati = await parseElencoElaborati(file);

    return NextResponse.json({
      success: true,
      count: elaborati.length,
      elaborati
    });

  } catch (error) {
    console.error("Errore verifiche preliminari:", error);

    return NextResponse.json(
      {
        success: false,
        error: "Errore durante la lettura dell'elenco elaborati"
      },
      { status: 500 }
    );
  }
}
