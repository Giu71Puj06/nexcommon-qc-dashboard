"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import Link from "next/link";

type Rilievo = {
  puntoPiano: string;
  tipo: "OK" | "OSS" | "NC" | "NA";
  descrizioneRilievo: string;
  azioneRichiesta: string;
  evidenzaNelDocumento: string;
};

type RisultatoAI = {
  codiceElaborato: string;
  titoloElaborato: string;
  revisione: string;
  disciplinaRilevata: string;
  tipoElaborato: string;
  esitoGenerale: "OK" | "OSS" | "NC" | "NA";
  priorita: "BASSA" | "MEDIA" | "ALTA";
  rilievi: Rilievo[];
  sintesiIspettore: string;
};

const piani = [
  {
    value: "architettonico",
    label: "Architettonico",
  },
  {
    value: "antincendio",
    label: "Antincendio",
  },
  {
    value: "demolizioni",
    label: "Demolizioni",
  },
  {
    value: "art44bis",
    label: "Art. 44 Bis",
  },
  {
    value: "pfte-pnrr",
    label: "PFTE PNRR",
  },
  {
    value: "generico",
    label: "Generico",
  },
];

export default function VerificaElaboratiPage() {
  const [files, setFiles] = useState<File[]>([]);
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState("");
  const [errore, setErrore] = useState("");
  const [piano, setPiano] = useState("architettonico");
  const [risultati, setRisultati] = useState<RisultatoAI[]>([]);

  async function analizzaElaborati() {
    try {
      setErrore("");
      setLoading(true);
      setRisultati([]);

      const allResults: RisultatoAI[] = [];

      for (let i = 0; i < files.length; i++) {
        const file = files[i];

        setProgress(
          `Analisi elaborato ${i + 1} di ${files.length}: ${file.name}`
        );

        const formData = new FormData();
        formData.append("file", file);
        formData.append("mode", piano);

        const response = await fetch("/api/ai-surveillance-check", {
          method: "POST",
          body: formData,
        });

        const data = await response.json();

        if (!response.ok || !data.success) {
          console.error(data);
          continue;
        }

        allResults.push(data.result);
      }

      setRisultati(allResults);
      setProgress("Analisi completata.");
    } catch (error) {
      console.error(error);
      setErrore("Errore durante l'analisi AI.");
    } finally {
      setLoading(false);
    }
  }

  function svuota() {
    setFiles([]);
    setRisultati([]);
    setErrore("");
    setProgress("");
  }

  function esportaExcel() {
    const rows: any[] = [];

    risultati.forEach((r) => {
      if (!r.rilievi?.length) {
        rows.push({
          Elaborato: r.codiceElaborato,
          Titolo: r.titoloElaborato,
          Revisione: r.revisione,
          Disciplina: r.disciplinaRilevata,
          Esito: r.esitoGenerale,
          Priorita: r.priorita,
          Punto: "-",
          Tipo: "-",
          Rilievo: r.sintesiIspettore,
          Azione: "-",
        });
      } else {
        r.rilievi.forEach((rilievo) => {
          rows.push({
            Elaborato: r.codiceElaborato,
            Titolo: r.titoloElaborato,
            Revisione: r.revisione,
            Disciplina: r.disciplinaRilevata,
            Esito: r.esitoGenerale,
            Priorita: r.priorita,
            Punto: rilievo.puntoPiano,
            Tipo: rilievo.tipo,
            Rilievo: rilievo.descrizioneRilievo,
            Azione: rilievo.azioneRichiesta,
          });
        });
      }
    });

    const worksheet = XLSX.utils.json_to_sheet(rows);

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      "Verifica Elaborati AI"
    );

    XLSX.writeFile(workbook, "Report_Verifica_Elaborati_AI.xlsx");
  }

  const stats = useMemo(() => {
    let ok = 0;
    let oss = 0;
    let nc = 0;

    risultati.forEach((r) => {
      if (r.esitoGenerale === "OK") ok++;
      if (r.esitoGenerale === "OSS") oss++;
      if (r.esitoGenerale === "NC") nc++;
    });

    return {
      ok,
      oss,
      nc,
    };
  }, [risultati]);

  return (
    <div className="min-h-screen bg-slate-100 p-6">
      <Link
        href="/dashboard-ai"
        className="text-blue-600 font-semibold"
      >
        ← Torna alla Dashboard AI
      </Link>

      <div className="mt-6">
        <h1 className="text-5xl font-bold text-slate-900">
          Verifica Elaborati AI
        </h1>

        <p className="text-slate-600 mt-4 text-lg max-w-5xl">
          Analisi AI elaborati tecnici rispetto ai Piani di Sorveglianza.
          Il sistema genera rilievi NC / OSS / OK per supportare gli
          ispettori ITS.
        </p>
      </div>

      <div className="bg-white rounded-3xl shadow-sm border border-slate-200 p-6 mt-8">
        <h2 className="text-3xl font-bold text-slate-900 mb-8">
          Analisi elaborati
        </h2>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <div>
            <label className="font-semibold text-slate-900 block mb-3">
              Piano di Sorveglianza
            </label>

            <select
              value={piano}
              onChange={(e) => setPiano(e.target.value)}
              className="w-full border border-slate-300 rounded-xl px-4 py-3"
            >
              {piani.map((p) => (
                <option key={p.value} value={p.value}>
                  {p.label}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="font-semibold text-slate-900 block mb-3">
              Elaborati PDF
            </label>

            <input
              type="file"
              multiple
              accept=".pdf"
              onChange={(e) => {
                const selected = Array.from(e.target.files || []);
                setFiles(selected);
              }}
              className="w-full border border-slate-300 rounded-xl p-3"
            />
          </div>
        </div>

        <div className="bg-slate-50 border border-slate-200 rounded-2xl p-4 mt-6">
          <div className="font-semibold text-blue-700">
            PDF selezionati: {files.length}
          </div>
        </div>

        <div className="flex gap-4 mt-8 flex-wrap">
          <button
            onClick={analizzaElaborati}
            disabled={loading || files.length === 0}
            className="bg-slate-950 text-white px-6 py-3 rounded-2xl font-semibold disabled:opacity-50"
          >
            {loading
              ? "Analisi in corso..."
              : "Esegui verifica AI"}
          </button>

          <button
            onClick={svuota}
            className="bg-slate-400 text-white px-6 py-3 rounded-2xl font-semibold"
          >
            Svuota dati
          </button>

          <button
            onClick={esportaExcel}
            disabled={!risultati.length}
            className="bg-green-600 text-white px-6 py-3 rounded-2xl font-semibold disabled:opacity-50"
          >
            Esporta Excel
          </button>
        </div>

        {progress && (
          <div className="mt-6 bg-cyan-50 text-cyan-800 border border-cyan-200 rounded-2xl p-4">
            {progress}
          </div>
        )}

        {errore && (
          <div className="mt-6 bg-red-50 text-red-700 border border-red-200 rounded-2xl p-4">
            {errore}
          </div>
        )}
      </div>

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-6 mt-8">
        <div className="bg-white rounded-3xl border border-slate-200 p-6">
          <div className="text-slate-500">Elaborati analizzati</div>
          <div className="text-5xl font-bold mt-4">
            {risultati.length}
          </div>
        </div>

        <div className="bg-white rounded-3xl border border-slate-200 p-6">
          <div className="text-green-600">OK</div>
          <div className="text-5xl font-bold mt-4 text-green-600">
            {stats.ok}
          </div>
        </div>

        <div className="bg-white rounded-3xl border border-slate-200 p-6">
          <div className="text-orange-500">OSS</div>
          <div className="text-5xl font-bold mt-4 text-orange-500">
            {stats.oss}
          </div>
        </div>

        <div className="bg-white rounded-3xl border border-slate-200 p-6">
          <div className="text-red-600">NC</div>
          <div className="text-5xl font-bold mt-4 text-red-600">
            {stats.nc}
          </div>
        </div>
      </div>

      <div className="bg-white rounded-3xl shadow-sm border border-slate-200 p-6 mt-8 overflow-auto">
        <h2 className="text-3xl font-bold text-slate-900 mb-6">
          Report ispettivo AI
        </h2>

        <table className="min-w-full text-sm">
          <thead>
            <tr className="border-b border-slate-200 text-left">
              <th className="p-3">Elaborato</th>
              <th className="p-3">Disciplina</th>
              <th className="p-3">Esito</th>
              <th className="p-3">Priorità</th>
              <th className="p-3">Rilievo</th>
              <th className="p-3">Azione richiesta</th>
            </tr>
          </thead>

          <tbody>
            {risultati.map((r, index) => (
              <tr
                key={index}
                className="border-b border-slate-100 align-top"
              >
                <td className="p-3 font-semibold">
                  {r.codiceElaborato || "-"}
                  <div className="text-slate-500 font-normal mt-1">
                    Rev. {r.revisione || "-"}
                  </div>
                </td>

                <td className="p-3">
                  {r.disciplinaRilevata || "-"}
                </td>

                <td className="p-3">
                  <span
                    className={`font-bold ${
                      r.esitoGenerale === "OK"
                        ? "text-green-600"
                        : r.esitoGenerale === "OSS"
                        ? "text-orange-500"
                        : r.esitoGenerale === "NC"
                        ? "text-red-600"
                        : "text-slate-500"
                    }`}
                  >
                    {r.esitoGenerale}
                  </span>
                </td>

                <td className="p-3">
                  {r.priorita || "-"}
                </td>

                <td className="p-3">
                  <div className="space-y-3">
                    {r.rilievi?.length ? (
                      r.rilievi.map((rilievo, idx) => (
                        <div
                          key={idx}
                          className="border border-slate-200 rounded-xl p-3"
                        >
                          <div className="font-semibold">
                            [{rilievo.tipo}] {rilievo.puntoPiano}
                          </div>

                          <div className="mt-1 text-slate-700">
                            {rilievo.descrizioneRilievo}
                          </div>
                        </div>
                      ))
                    ) : (
                      <div>{r.sintesiIspettore}</div>
                    )}
                  </div>
                </td>

                <td className="p-3">
                  <div className="space-y-3">
                    {r.rilievi?.length ? (
                      r.rilievi.map((rilievo, idx) => (
                        <div key={idx}>
                          {rilievo.azioneRichiesta}
                        </div>
                      ))
                    ) : (
                      <div>-</div>
                    )}
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        {!risultati.length && (
          <div className="text-slate-500 mt-6">
            Nessun elaborato analizzato.
          </div>
        )}
      </div>
    </div>
  );
}
