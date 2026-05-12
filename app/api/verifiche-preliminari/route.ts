import { NextRequest, NextResponse } from "next/server";

export async function POST(req: NextRequest) {

  try {

    const body = await req.json();

    return NextResponse.json({
      success: true,
      message: "Modulo verifiche preliminari attivo",
      received: body
    });

  } catch (error) {

    return NextResponse.json(
      {
        success: false,
        error: "Errore API verifiche preliminari"
      },
      { status: 500 }
    );

  }

}
