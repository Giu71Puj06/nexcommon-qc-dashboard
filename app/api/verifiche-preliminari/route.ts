import { NextRequest, NextResponse } from "next/server";
import { parseElencoElaborati } from "@/lib/verifiche-preliminari/excel-parser";

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();

    const excel = formData.get("excel");
    const template = formData.get("template");
    const elaborati = formData.getAll("elaborati");
    const paths = formData.getAll("paths").map(String);

    if (!excel || !(excel instanceof File)) {
      return NextResponse.json(
        { success: false, error: "Nessun file Excel ricevuto" },
        { status: 400 }
      );
    }

    if (!template || !(template instanceof File)) {
      return NextResponse.json(
        { success: false, error: "Nessun template Word ricevuto" },
        { status: 400 }
      );
    }

    const elaboratiFiles = elaborati.filter(
      (f): f is File => f instanceof File
    );

    if (elaboratiFiles.length === 0) {
      return NextResponse.json(
        { success: false, error: "Nessun elaborato PDF ricevuto" },
        { status: 400 }
      );
    }

    const elaboratiDaExcel = await parseElencoElaborati(excel);

    return NextResponse.json({
      success: true,
      count: elaboratiDaExcel.length,
      pdfCount: elaboratiFiles.length,
      templateName: template.name,
      paths,
      elaborati: elaboratiDaExcel,
    });
  } catch (error) {
    console.error("Errore verifiche preliminari:", error);

    return NextResponse.json(
      {
        success: false,
        error: "Errore durante l'analisi delle verifiche preliminari",
      },
      { status: 500 }
    );
  }
}
