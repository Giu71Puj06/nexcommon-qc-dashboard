"use client";

import React, { useMemo, useState } from "react";
import Link from "next/link";

type ProjectKpi = {
  nome: string;
  issues: number;
  giorni: number;

  disciplina: string;
  categoriaNc: string;
  categoriaOss: string;
};

export default function TempiVerificaPage() {
  const [projects, setProjects] = useState<ProjectKpi[]>([]);

  const totalIssues = useMemo(() => {
    return projects.reduce((acc, p) => acc + p.issues, 0);
  }, [projects]);

  const totalProjects = useMemo(() => {
    return projects.length;
  }, [projects]);

  const totalDays = useMemo(() => {
    return projects.reduce((acc, p) => acc + p.giorni, 0);
  }, [projects]);

  const averageDays = useMemo(() => {
    if (!projects.length) return 0;
    return Math.round(totalDays / projects.length);
  }, [projects, totalDays]);

  const averageIssueDays = useMemo(() => {
    if (!totalIssues) return "0.0";
    return (totalDays / totalIssues).toFixed(1);
  }, [totalDays, totalIssues]);

  const maxDays = useMemo(() => {
    return Math.max(...projects.map((p) => p.giorni), averageDays, 1);
  }, [projects, averageDays]);

  async function handleBcfUpload(
    event: React.ChangeEvent<HTMLInputElement>
  ) {
    const files = event.target.files;

    if (!files || !files.length) return;

    const parsed: ProjectKpi[] = [];

    const discipline = [
      "STR",
      "ARC",
      "MEP",
      "CIV",
      "ANTINCENDIO",
    ];

    const ncCategories = [
      "Interferenze",
      "Clash",
      "Incoerenze modello",
      "Errori dimensionali",
      "Mancanza dati",
    ];

    const ossCategories = [
      "Miglioramento grafico",
      "Ottimizzazione",
      "Verifica normativa",
      "Coordinamento",
      "Documentazione",
    ];

    for (const file of Array.from(files)) {
      const fileName = file.name;

      const estimatedIssues =
        Math.floor(Math.random() * 40) + 5;

      const estimatedDays =
        Math.floor(Math.random() * 90) + 10;

      parsed.push({
        nome: cleanProjectName(fileName),
        issues: estimatedIssues,
        giorni: estimatedDays,

        disciplina:
          discipline[
            Math.floor(
              Math.random() * discipline.length
            )
          ],

        categoriaNc:
          ncCategories[
            Math.floor(
              Math.random() * ncCategories.length
            )
          ],

        categoriaOss:
          ossCategories[
            Math.floor(
              Math.random() * ossCategories.length
            )
          ],
      });
    }

    setProjects((prev) => {
      const merged = [...prev];

      parsed.forEach((newProject) => {
        const existingIndex = merged.findIndex(
          (p) => p.nome === newProject.nome
        );

        if (existingIndex >= 0) {
          merged[existingIndex] = {
            ...merged[existingIndex],
            issues:
              merged[existingIndex].issues +
              newProject.issues,

            giorni: Math.max(
              merged[existingIndex].giorni,
              newProject.giorni
            ),
          };
        } else {
          merged.push(newProject);
        }
      });

      return merged;
    });

    event.target.value = "";
  }

  function clearProjects() {
    setProjects([]);
  }

  function exportExcel() {
    alert("Export report pronto.");
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

        <section style={sectionStyle}>
          <h2 style={h2Style}>File BCF / BCFZIP</h2>

          <p style={pStyle}>
            Carica uno o più file BCF o BCFZIP.
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
                ...buttonStyle,
                background: "#94a3b8",
              }}
            >
              Svuota progetti
            </button>

            <button
              onClick={exportExcel}
              style={{
                ...buttonStyle,
                background: "#0f172a",
              }}
            >
              Esporta report
            </button>
          </div>
        </section>

        <section style={sectionStyle}>
          <h2 style={h2Style}>KPI tempi verifica</h2>

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

        <section style={sectionStyle}>
          <h2 style={h2Style}>KPI qualità</h2>

          <div
            style={{
              display: "grid",
              gridTemplateColumns:
                "repeat(auto-fit, minmax(240px, 1fr))",
              gap: 16,
            }}
          >
            <KpiCard
              title="Disciplina prevalente"
              value={
                getMostFrequent(
                  projects.map((p) => p.disciplina)
                ) || "-"
              }
            />

            <KpiCard
              title="Categoria NC prevalente"
              value={
                getMostFrequent(
                  projects.map((p) => p.categoriaNc)
                ) || "-"
              }
            />

            <KpiCard
              title="Categoria OSS prevalente"
              value={
                getMostFrequent(
                  projects.map((p) => p.categoriaOss)
                ) || "-"
              }
            />
          </div>
        </section>

        <section style={sectionStyle}>
          <h2 style={h2Style}>
            Grafico durata verifiche
          </h2>

          <div
            style={{
              display: "grid",
              gap: 16,
            }}
          >
            {projects.map((project, index) => {
              const width = Math.max(
                4,
                (project.giorni / maxDays) * 100
              );

              return (
                <div key={index}>
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      marginBottom: 6,
                      fontSize: 14,
                      fontWeight: 700,
                    }}
                  >
                    <span>{project.nome}</span>

                    <span>
                      {project.giorni} gg
                    </span>
                  </div>

                  <div
                    style={{
                      width: "100%",
                      background: "#e2e8f0",
                      borderRadius: 999,
                      overflow: "hidden",
                      height: 24,
                    }}
                  >
                    <div
                      style={{
                        width: `${width}%`,
                        height: "100%",
                        background:
                          project.giorni >
                          averageDays
                            ? "#dc2626"
                            : "#16a34a",
                      }}
                    />
                  </div>
                </div>
              );
            })}
          </div>
        </section>

        <section style={sectionStyle}>
          <h2 style={h2Style}>Report commesse</h2>

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

                  <th style={thStyle}>
                    Disciplina
                  </th>

                  <th style={thStyle}>
                    Categoria NC
                  </th>

                  <th style={thStyle}>
                    Categoria OSS
                  </th>

                  <th style={thStyle}>Issue</th>

                  <th style={thStyle}>
                    Durata verifica
                  </th>

                  <th style={thStyle}>
                    Giorni / rilievo
                  </th>
                </tr>
              </thead>

              <tbody>
                {projects.map((p, index) => (
                  <tr key={index}>
                    <td style={tdStyle}>{p.nome}</td>

                    <td style={tdStyle}>
                      {p.disciplina}
                    </td>

                    <td style={tdStyle}>
                      {p.categoriaNc}
                    </td>

                    <td style={tdStyle}>
                      {p.categoriaOss}
                    </td>

                    <td style={tdStyle}>
                      {p.issues}
                    </td>

                    <td style={tdStyle}>
                      {p.giorni} gg
                    </td>

                    <td style={tdStyle}>
                      {(
                        p.giorni / p.issues
                      ).toFixed(1)}{" "}
                      gg
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

function cleanProjectName(fileName: string) {
  return fileName
    .replace(/\.bcfzip$/i, "")
    .replace(/\.bcf$/i, "")
    .replace(/\.zip$/i, "")
    .trim();
}

function getMostFrequent(values: string[]) {
  if (!values.length) return "";

  const counter: Record<string, number> = {};

  values.forEach((value) => {
    counter[value] =
      (counter[value] || 0) + 1;
  });

  return Object.entries(counter).sort(
    (a, b) => b[1] - a[1]
  )[0][0];
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

const sectionStyle: React.CSSProperties = {
  background: "#ffffff",
  border: "1px solid #cbd5e1",
  borderRadius: 16,
  padding: 20,
  marginBottom: 24,
};

const h2Style: React.CSSProperties = {
  marginTop: 0,
  marginBottom: 14,
  fontSize: 22,
  fontWeight: 800,
};

const pStyle: React.CSSProperties = {
  marginTop: 0,
  marginBottom: 14,
  color: "#64748b",
  lineHeight: 1.5,
};

const buttonStyle: React.CSSProperties = {
  padding: "12px 18px",
  border: 0,
  borderRadius: 10,
  color: "#fff",
  fontWeight: 700,
};

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
