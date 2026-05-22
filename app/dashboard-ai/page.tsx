"use client";

import Link from "next/link";

const modules = [
  {
    title: "Variazioni Economiche AI",
    description:
      "Analisi intelligente computi metrici, importi, cartigli e differenze economiche automatiche.",
    href: "/dashboard-ai/variazioni-economiche",
    color: "from-blue-600 to-cyan-500",
    icon: "💰",
  },
  {
    title: "Nota Ricezione Elaborati AI",
    description:
      "Verifica automatica cartigli, revisioni, naming elaborati e coerenza documentale.",
    href: "/dashboard-ai/nota-ricezione",
    color: "from-purple-600 to-fuchsia-500",
    icon: "📑",
  },
  {
    title: "Controllo Qualità AI",
    description:
      "Rilevamento anomalie, elaborati mancanti, discipline errate e revisioni incoerenti.",
    href: "/dashboard-ai/controllo-qualita",
    color: "from-emerald-600 to-green-500",
    icon: "✅",
  },
  {
    title: "Assistente Tecnico AI",
    description:
      "Assistente tecnico intelligente per analisi progetto, verifiche e report automatici.",
    href: "/dashboard-ai/assistente-tecnico",
    color: "from-orange-500 to-red-500",
    icon: "🤖",
  },
];

export default function DashboardAIPage() {
  return (
    <main className="min-h-screen bg-slate-100 p-8">
      <div className="max-w-7xl mx-auto">
        <div className="mb-10">
          <Link
            href="/dashboard-pm"
            className="inline-flex items-center text-blue-700 hover:text-blue-900 font-medium mb-6"
          >
            ← Torna alla Dashboard PM
          </Link>

          <h1 className="text-5xl font-bold text-slate-900 mb-4">
            Dashboard AI
          </h1>

          <p className="text-xl text-slate-600 max-w-4xl">
            Sistema AI avanzato per controllo qualità, verifiche documentali,
            analisi computi, coerenza elaborati e assistenza tecnica automatica.
          </p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-2 gap-8">
          {modules.map((module) => (
            <Link key={module.title} href={module.href}>
              <div className="bg-white rounded-3xl shadow-xl hover:shadow-2xl transition-all duration-300 overflow-hidden border border-slate-200 hover:scale-[1.02]">
                <div
                  className={`h-3 bg-gradient-to-r ${module.color}`}
                />

                <div className="p-8">
                  <div className="text-5xl mb-5">{module.icon}</div>

                  <h2 className="text-2xl font-bold text-slate-900 mb-4">
                    {module.title}
                  </h2>

                  <p className="text-slate-600 text-lg leading-relaxed mb-6">
                    {module.description}
                  </p>

                  <div className="inline-flex items-center font-semibold text-blue-700">
                    Apri modulo →
                  </div>
                </div>
              </div>
            </Link>
          ))}
        </div>

        <div className="mt-12 bg-white rounded-3xl shadow-xl border border-slate-200 p-8">
          <h3 className="text-2xl font-bold text-slate-900 mb-4">
            Stato AI Platform
          </h3>

          <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
            <div className="bg-slate-100 rounded-2xl p-6">
              <div className="text-sm text-slate-500 mb-2">
                AI Engine
              </div>
              <div className="text-2xl font-bold text-green-600">
                ONLINE
              </div>
            </div>

            <div className="bg-slate-100 rounded-2xl p-6">
              <div className="text-sm text-slate-500 mb-2">
                OpenAI API
              </div>
              <div className="text-2xl font-bold text-green-600">
                ATTIVA
              </div>
            </div>

            <div className="bg-slate-100 rounded-2xl p-6">
              <div className="text-sm text-slate-500 mb-2">
                AI Reader
              </div>
              <div className="text-2xl font-bold text-green-600">
                OPERATIVO
              </div>
            </div>

            <div className="bg-slate-100 rounded-2xl p-6">
              <div className="text-sm text-slate-500 mb-2">
                GPT Model
              </div>
              <div className="text-2xl font-bold text-blue-700">
                GPT-4.1-mini
              </div>
            </div>
          </div>
        </div>
      </div>
    </main>
  );
}
