import { NextResponse } from "next/server";

/**
 * Proxy server-side verso l'app Trimble (verifica-elaborati) per elencare i
 * progetti. Sta lato server per due motivi:
 *  - evita problemi di CORS (il browser non chiama direttamente l'altra app);
 *  - l'URL dell'app Trimble resta configurabile e non esposto al client.
 *
 * Configura su Railway (dashboard) la variabile:
 *   TRIMBLE_API_URL=https://verifica-elaborati-production.up.railway.app
 * Se non impostata usa il default qui sotto.
 */

const TRIMBLE_API_URL = (
  process.env.TRIMBLE_API_URL ||
  "https://verifica-elaborati-production.up.railway.app"
).replace(/\/+$/, "");

export const dynamic = "force-dynamic";

export async function GET() {
  try {
    const res = await fetch(`${TRIMBLE_API_URL}/tc/projects`, {
      headers: { Accept: "application/json" },
      cache: "no-store",
    });

    if (!res.ok) {
      const text = await res.text().catch(() => "");
      return NextResponse.json(
        { ok: false, error: `Trimble ${res.status}: ${text.slice(0, 300)}` },
        { status: 502 }
      );
    }

    const data = await res.json();
    const projectsRaw = Array.isArray(data?.projects) ? data.projects : [];

    // Normalizza a { id, name } indipendentemente dalle varianti di chiave.
    const projects = projectsRaw
      .map((p: any) => ({
        id: String(p?.id ?? p?.projectId ?? "").trim(),
        name: String(p?.name ?? p?.projectName ?? "").trim(),
      }))
      .filter((p: any) => p.id && p.name)
      .sort((a: any, b: any) => a.name.localeCompare(b.name));

    return NextResponse.json({ ok: true, count: projects.length, projects });
  } catch (e: any) {
    return NextResponse.json(
      { ok: false, error: e?.message || "Errore lettura progetti Trimble" },
      { status: 502 }
    );
  }
}
