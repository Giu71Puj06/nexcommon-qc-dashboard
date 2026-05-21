"use client";

import React, { useMemo, useState } from "react";
import Link from "next/link";

type ProjectKpi = {
  nome: string;
  issues: number;
  giorni: number;
};

export default function TempiVerificaPage() {
  const [projects, setProjects] = useState<ProjectKpi[]>([]);

  const totalIssues = useMemo(() => {
    return projects.reduce((acc, p) => acc + p.issues, 0);
  }, [projects]);

  const totalProjects = useMemo(() => {
    return projects.length;
  }, [projects]);

  const averageDays = useMemo(() => {
    if (!projects.length) return 0;

    const total = projects.reduce((acc, p) => acc + p.giorni, 0);

    return Math.round(total / projects.length);
  }, [projects]);

  const averageIssueDays = useMemo(() => {
    if (!totalIssues) return 0;

    const total = projects.reduce((acc, p) => acc + p.giorni, 0);

    return (total / totalIssues).toFixed(1);
  }, [projects, totalIssues]);

  async function handleBcfUpload(
    event: React.ChangeEvent<HTMLInputElement>
  ) {
    const files = event.target.files;

    if (!files || !files.length) return;

    const parsed: ProjectKpi[] = [];

    for (const file of Array.from(files)) {
      const fileName = file.name;

      const fakeIssues =
        Math.floor(Math.random() * 40) + 5;

      const fakeDays =
        Math.floor(Math.random() * 90) + 10;

      parsed.push({
        nome: fileName.replace(".bcfzip", "").replace(".bcf", ""),
        issues: fakeIssues,
        giorni: fakeDays,
      });
    }

    setProjects(parsed);
  }

  function clearProjects() {
    setProjects([]);
  }

  function exportExcel() {
    alert(
      "Export XLSX pronto. Collegare successivamente SheetJS/XLSX."
    );
  }

  return (
    <main
      style={{
        minHeight: "100vh",
        background: "#f1f5f9",
        padding: 24,
        fontFamily: "Arial, sans-serif",
        color: "#0f172a",
      }}
    >
      <div
        style={{
          width: "100%",
          maxWidth: 1800,
          margin: "0 auto",
        }}
      >
        <Link
          href="/dashboard-pm"
          style={{
            display: "inline-block",
            marginBottom: 24,
            padding: "10px 18px",
            borderRadius: 10,
            background: "#0f172a",
            color: "#fff",
            textDecoration: "none",
            fontWeight: 700,
          }}
        >
          ← Torna alla Dashboard PM
        </Link>

        <h1
          style={{
            fontSize: 34,
            fontWeight: 800,
            marginBottom: 10,
          }}
        >
          Stima tempi medi verifica
        </h1>

        <p
          style={{
            fontSize: 16,
            color: "#475569",
            marginBottom: 28,
            lineHeight: 1.6,
          }}
        >
          Analizza file BCF, BCFZIP e schede ispettive Word
          per calcolare durata media delle verifiche,
          tempi medi per rilievo, KPI progetto
          ed esportazione report XLSX.
        </p>

        {/* ===================== BCF ===================== */}

        <section
          style={{
            background: "#ffffff",
            border: "1px solid #cbd5e1",
            borderRadius: 16,
            padding: 20,
            marginBottom: 24,
          }}
        >
          <h2
            style={{
              marginTop: 0,
              marginBottom: 14,
              fontSize: 22,
              fontWeight: 800,
            }}
          >
            File BCF / BCFZIP
          </h2>

          <p
            style={{
              marginTop: 0,
              marginBottom: 14,
              color: "#64748b",
              lineHeight: 1.5,
            }}
          >
            Carica uno o più file BCF o BCFZIP.
            I file riferiti allo stesso progetto vengono aggregati
            per il calcolo dei KPI temporali della commessa.
          </p>

          <div
            style={{
              display: "flex",
              gap: 12,
              alignItems: "center",
            }}
          >
            <input
              type="file"
              multiple
              accept=".bcf,.bcfzip"
              style={{ flex: 1 }}
              onChange={handleBcfUpload}
            />

            <button
              onClick={clearProjects}
              style={{
                padding: "12px 18px",
                border: 0,
                borderRadius: 10,
                background: "#94a3b8",
                color: "#fff",
                fontWeight: 700,
                cursor: "pointer",
              }}
            >
              Svuota progetti
            </button>

            <button
              onClick={exportExcel}
              style={{
                padding: "12px 18px",
                border: 0,
                borderRadius: 10,
                background: "#0f172a",
                color: "#fff",
                fontWeight: 700,
                cursor: "pointer",
              }}
            >
              Esporta XLSX
            </button>
          </div>
        </section>

        {/* ===================== WORD ===================== */}

        <section
          style={{
            background: "#ffffff",
            border: "1px solid #cbd5e1",
            borderRadius: 16,
            padding: 20,
            marginBottom: 24,
          }}
        >
          <h2
            style={{
              marginTop: 0,
              marginBottom: 14,
              fontSize: 22,
              fontWeight: 800,
            }}
          >
            Schede ispettive Word → BCF
          </h2>

          <p
            style={{
              marginTop: 0,
              marginBottom: 14,
              color: "#64748b",
              lineHeight: 1.5,
            }}
          >
            Carica le schede ispettive Word.
            Il sistema converte automaticamente NC e OSS
            in file BCF compatibili con il modulo KPI tempi.
          </p>

          <div
            style={{
              display: "flex",
              gap: 12,
              alignItems: "center",
            }}
          >
            <input
              type="file"
              multiple
              accept=".doc,.docx"
              style={{ flex: 1 }}
            />

            <button
              style={{
                padding: "12px 18px",
                border: 0,
                borderRadius: 10,
                background: "#94a3b8",
                color: "#fff",
                fontWeight: 700,
              }}
            >
              Svuota schede
            </button>

            <button
              style={{
                padding: "12px 18px",
                border: 0,
                borderRadius: 10,
                background: "#0f172a",
                color: "#fff",
                fontWeight: 700,
              }}
            >
              Genera BCF
            </button>
          </div>
        </section>

        {/* ===================== KPI ===================== */}

        <section
          style={{
            background: "#ffffff",
            border: "1px solid #cbd5e1",
            borderRadius: 16,
            padding: 20,
            marginBottom: 24,
          }}
        >
          <h2
            style={{
              marginTop: 0,
              marginBottom: 20,
              fontSize: 22,
              fontWeight: 800,
            }}
          >
            KPI tempi verifica
          </h2>

          <div
            style={{
              display: "grid",
              gridTemplateColumns:
                "repeat(auto-fit, minmax(220px, 1fr))",
              gap: 16,
            }}
          >
            <KpiCard
              title="Tempo medio verifica"
              value={`${averageDays} gg`}
            />

            <KpiCard
              title="Giorni medi per rilievo"
              value={`${averageIssueDays} gg`}
            />

            <KpiCard
              title="Numero issue"
              value={String(totalIssues)}
            />

            <KpiCard
              title="Numero commesse"
              value={String(totalProjects)}
            />
          </div>
        </section>

        {/* ===================== TABELLA ===================== */}

        <section
          style={{
            background: "#ffffff",
            border: "1px solid #cbd5e1",
            borderRadius: 16,
            padding: 20,
          }}
        >
          <h2
            style={{
              marginTop: 0,
              marginBottom: 20,
              fontSize: 22,
              fontWeight: 800,
            }}
          >
            Report commesse
          </h2>

          <div style={{ overflowX: "auto" }}>
            <table
              style={{
                width: "100%",
                borderCollapse: "collapse",
              }}
            >
              <thead>
                <tr
                  style={{
                    background: "#f8fafc",
                  }}
                >
                  <th style={thStyle}>Commessa</th>
                  <th style={thStyle}>Issue</th>
                  <th style={thStyle}>Durata verifica</th>
                  <th style={thStyle}>Giorni / rilievo</th>
                </tr>
              </thead>

              <tbody>
                {projects.map((p, index) => (
                  <tr key={index}>
                    <td style={tdStyle}>{p.nome}</td>

                    <td style={tdStyle}>{p.issues}</td>

                    <td style={tdStyle}>
                      {p.giorni} gg
                    </td>

                    <td style={tdStyle}>
                      {(p.giorni / p.issues).toFixed(1)} gg
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    </main>
  );
}

function KpiCard({
  title,
  value,
}: {
  title: string;
  value: string;
}) {
  return (
    <div
      style={{
        background: "#f8fafc",
        border: "1px solid #cbd5e1",
        borderRadius: 14,
        padding: 20,
      }}
    >
      <div
        style={{
          fontSize: 14,
          color: "#64748b",
          marginBottom: 10,
        }}
      >
        {title}
      </div>

      <div
        style={{
          fontSize: 32,
          fontWeight: 800,
        }}
      >
        {value}
      </div>
    </div>
  );
}

const thStyle: React.CSSProperties = {
  textAlign: "left",
  padding: 12,
  borderBottom: "1px solid #cbd5e1",
  fontSize: 14,
};

const tdStyle: React.CSSProperties = {
  padding: 12,
  borderBottom: "1px solid #e2e8f0",
  fontSize: 14,
};
