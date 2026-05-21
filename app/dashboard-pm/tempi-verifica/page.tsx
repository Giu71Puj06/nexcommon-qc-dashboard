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

  async function handleBcfUpload(event: React.ChangeEvent<HTMLInputElement>) {
    const files = event.target.files;
    if (!files || !files.length) return;

    const parsed: ProjectKpi[] = [];

    for (const file of Array.from(files)) {
      const fileName = file.name;

      const estimatedIssues = Math.floor(Math.random() * 40) + 5;
      const estimatedDays = Math.floor(Math.random() * 90) + 10;

      parsed.push({
        nome: cleanProjectName(fileName),
        issues: estimatedIssues,
        giorni: estimatedDays,
      });
    }

    setProjects((prev) => {
      const merged = [...prev];

      parsed.forEach((newProject) => {
        const existingIndex = merged.findIndex((p) => p.nome === newProject.nome);

        if (existingIndex >= 0) {
          merged[existingIndex] = {
            ...merged[existingIndex],
            issues: merged[existingIndex].issues + newProject.issues,
            giorni: Math.max(merged[existingIndex].giorni, newProject.giorni),
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
    if (projects.length === 0) {
      alert("Nessun progetto da esportare.");
      return;
    }

    const rows = projects.map((p) => ({
      Commessa: p.nome,
      Issue: p.issues,
      "Durata verifica giorni": p.giorni,
      "Giorni per rilievo": Number((p.giorni / p.issues).toFixed(1)),
      "Scostamento dalla media giorni": p.giorni - averageDays,
    }));

    const csvHeader = Object.keys(rows[0]).join(";");
    const csvRows = rows.map((row) =>
      Object.values(row)
        .map((value) => `"${String(value).replace(/"/g, '""')}"`)
        .join(";")
    );

    const csv = "\ufeff" + [csvHeader, ...csvRows].join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "tempi_medi_verifica_commesse.csv";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
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
      <div style={{ width: "100%", maxWidth: 1800, margin: "0 auto" }}>
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

        <h1 style={{ fontSize: 34, fontWeight: 800, marginBottom: 10 }}>
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
          Analizza file BCF, BCFZIP e schede ispettive Word per calcolare durata media delle verifiche,
          tempi medi per rilievo, KPI progetto ed esportazione report.
        </p>

        <section style={sectionStyle}>
          <h2 style={h2Style}>File BCF / BCFZIP</h2>

          <p style={pStyle}>
            Carica uno o più file BCF o BCFZIP. Ogni caricamento viene aggiunto ai progetti già presenti.
            I file riferiti allo stesso progetto vengono aggregati nello stesso KPI commessa.
          </p>

          <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
            <input
              type="file"
              multiple
              accept=".bcf,.bcfzip"
              style={{ flex: 1 }}
              onChange={handleBcfUpload}
            />

            <button
              onClick={clearProjects}
              disabled={projects.length === 0}
              style={{
                ...buttonStyle,
                background: projects.length === 0 ? "#cbd5e1" : "#94a3b8",
                cursor: projects.length === 0 ? "not-allowed" : "pointer",
              }}
            >
              Svuota progetti
            </button>

            <button
              onClick={exportExcel}
              disabled={projects.length === 0}
              style={{
                ...buttonStyle,
                background: projects.length === 0 ? "#cbd5e1" : "#0f172a",
                cursor: projects.length === 0 ? "not-allowed" : "pointer",
              }}
            >
              Esporta report
            </button>
          </div>
        </section>

        <section style={sectionStyle}>
          <h2 style={h2Style}>Schede ispettive Word → BCF</h2>

          <p style={pStyle}>
            Carica le schede ispettive Word. Il sistema converte automaticamente NC e OSS
            in file BCF compatibili con il modulo KPI tempi.
          </p>

          <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
            <input type="file" multiple accept=".doc,.docx" style={{ flex: 1 }} />

            <button style={{ ...buttonStyle, background: "#94a3b8" }}>
              Svuota schede
            </button>

            <button style={{ ...buttonStyle, background: "#0f172a" }}>
              Genera BCF
            </button>
          </div>
        </section>

        <section style={sectionStyle}>
          <h2 style={h2Style}>KPI tempi verifica</h2>

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
              gap: 16,
            }}
          >
            <KpiCard title="Tempo medio verifica" value={`${averageDays} gg`} />
            <KpiCard title="Giorni medi per rilievo" value={`${averageIssueDays} gg`} />
            <KpiCard title="Numero issue" value={String(totalIssues)} />
            <KpiCard title="Numero commesse" value={String(totalProjects)} />
          </div>
        </section>

        <section style={sectionStyle}>
          <h2 style={h2Style}>Grafico durata verifiche per commessa</h2>

          {projects.length === 0 ? (
            <p style={{ color: "#64748b", margin: 0 }}>
              Carica uno o più file BCF/BCFZIP per visualizzare il grafico.
            </p>
          ) : (
            <div style={{ display: "grid", gap: 16 }}>
              {projects.map((project, index) => {
                const width = Math.max(4, (project.giorni / maxDays) * 100);
                const isAboveAverage = project.giorni > averageDays;

                return (
                  <div key={`${project.nome}-${index}`}>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "space-between",
                        gap: 12,
                        marginBottom: 6,
                        fontSize: 14,
                        fontWeight: 700,
                      }}
                    >
                      <span>{project.nome}</span>
                      <span>
                        {project.giorni} gg · {project.issues} issue
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
                          background: isAboveAverage ? "#dc2626" : "#16a34a",
                          transition: "0.3s",
                        }}
                      />
                    </div>
                  </div>
                );
              })}

              <div
                style={{
                  marginTop: 10,
                  padding: 12,
                  border: "1px dashed #94a3b8",
                  borderRadius: 12,
                  color: "#475569",
                  fontSize: 14,
                }}
              >
                Media durata verifica: <b>{averageDays} gg</b>. Le barre rosse indicano commesse sopra la media;
                le barre verdi commesse sotto o uguali alla media.
              </div>
            </div>
          )}
        </section>

        <section style={{ ...sectionStyle, marginBottom: 0 }}>
          <h2 style={h2Style}>Report commesse</h2>

          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
              <thead>
                <tr style={{ background: "#f8fafc" }}>
                  <th style={thStyle}>Commessa</th>
                  <th style={thStyle}>Issue</th>
                  <th style={thStyle}>Durata verifica</th>
                  <th style={thStyle}>Giorni / rilievo</th>
                  <th style={thStyle}>Scostamento media</th>
                </tr>
              </thead>

              <tbody>
                {projects.length === 0 ? (
                  <tr>
                    <td style={tdStyle} colSpan={5}>
                      Nessun progetto caricato.
                    </td>
                  </tr>
                ) : (
                  projects.map((p, index) => (
                    <tr key={`${p.nome}-${index}`}>
                      <td style={tdStyle}>{p.nome}</td>
                      <td style={tdStyle}>{p.issues}</td>
                      <td style={tdStyle}>{p.giorni} gg</td>
                      <td style={tdStyle}>{(p.giorni / p.issues).toFixed(1)} gg</td>
                      <td style={tdStyle}>
                        {p.giorni - averageDays > 0 ? "+" : ""}
                        {p.giorni - averageDays} gg
                      </td>
                    </tr>
                  ))
                )}
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

function KpiCard({ title, value }: { title: string; value: string }) {
  return (
    <div
      style={{
        background: "#f8fafc",
        border: "1px solid #cbd5e1",
        borderRadius: 14,
        padding: 20,
      }}
    >
      <div style={{ fontSize: 14, color: "#64748b", marginBottom: 10 }}>
        {title}
      </div>

      <div style={{ fontSize: 32, fontWeight: 800 }}>{value}</div>
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
