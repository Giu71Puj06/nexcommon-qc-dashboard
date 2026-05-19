"use client";

import { useState } from "react";
import * as XLSX from "xlsx";

import { readBCF } from "@/lib/dashboard-pm/read-bcf";
import {
  extractBCFHistory,
  BCFIssue,
} from "@/lib/dashboard-pm/extract-history";

type ProjectIssues = {
  projectName: string;
  fileName: string;
  issues: BCFIssue[];
};

type InspectionDoc = {
  id: string;
  projectName: string;
  fileName: string;
  url: string;
};

type Selection = {
  projectName: string;
  type: "all" | "open" | "closed";
  label: string;
} | null;

type ProjectKpi = {
  projectName: string;
  fileName: string;
  total: number;
  closed: number;
  open: number;
  closure: number;
  firstDate: string;
  lastDate: string;
  durationDays: number;
};

function daysBetween(start: string, end: string): number {
  const startDate = new Date(start);
  const endDate = new Date(end);

  if (
    Number.isNaN(startDate.getTime()) ||
    Number.isNaN(endDate.getTime())
  ) {
    return 0;
  }

  return Math.ceil(
    (endDate.getTime() - startDate.getTime()) /
      (1000 * 60 * 60 * 24)
  );
}

function formatDate(value: string): string {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "-";
  return date.toLocaleDateString("it-IT");
}

function cleanProjectName(fileName: string): string {
  return fileName
    .replace(/\.bcfzip$/i, "")
    .replace(/\.bcf$/i, "")
    .replace(/\.zip$/i, "")
    .replace(/\.docx$/i, "")
    .replace(/\d{12}\+\d{4}$/i, "")
    .trim();
}

function getProjectKpi(project: ProjectIssues): ProjectKpi {
  const issues = project.issues;

  const firstDate =
    [...issues].sort(
      (a, b) =>
        new Date(a.creationDate).getTime() -
        new Date(b.creationDate).getTime()
    )[0]?.creationDate || "";

  const lastDate =
    [...issues].sort(
      (a, b) =>
        new Date(b.lastActivityDate).getTime() -
        new Date(a.lastActivityDate).getTime()
    )[0]?.lastActivityDate || "";

  const closed = issues.filter((i) => i.isClosed).length;
  const open = issues.length - closed;

  const closure =
    issues.length > 0 ? Math.round((closed / issues.length) * 100) : 0;

  return {
    projectName: project.projectName,
    fileName: project.fileName,
    total: issues.length,
    closed,
    open,
    closure,
    firstDate,
    lastDate,
    durationDays: daysBetween(firstDate, lastDate),
  };
}

export default function DashboardPMPage() {
  const [projects, setProjects] = useState<ProjectIssues[]>([]);
  const [inspectionDocs, setInspectionDocs] = useState<InspectionDoc[]>([]);
  const [loading, setLoading] = useState(false);
  const [selection, setSelection] = useState<Selection>(null);

  async function handleFiles(event: React.ChangeEvent<HTMLInputElement>) {
    const files = Array.from(event.target.files || []);
    setLoading(true);

    const parsedProjects: ProjectIssues[] = [];

    for (const file of files) {
      const name = file.name.toLowerCase();

      if (
        !name.endsWith(".bcfzip") &&
        !name.endsWith(".bcf") &&
        !name.endsWith(".zip")
      ) {
        continue;
      }

      const markups = await readBCF(file);
      const issues: BCFIssue[] = [];

      for (const markup of markups) {
        const issue = extractBCFHistory(markup.rawXml);
        if (issue) issues.push(issue);
      }

      parsedProjects.push({
        projectName: cleanProjectName(file.name),
        fileName: file.name,
        issues,
      });
    }

    setProjects((prev) => {
      const mergedProjects = [...prev];

      for (const newProject of parsedProjects) {
        const existingProject = mergedProjects.find(
          (project) => project.projectName === newProject.projectName
        );

        if (existingProject) {
          existingProject.issues = [
            ...existingProject.issues,
            ...newProject.issues,
          ];
          existingProject.fileName = `${existingProject.fileName}, ${newProject.fileName}`;
        } else {
          mergedProjects.push(newProject);
        }
      }

      return mergedProjects;
    });

    setLoading(false);
    event.target.value = "";
  }

  function handleInspectionDocs(
    event: React.ChangeEvent<HTMLInputElement>
  ) {
    const files = Array.from(event.target.files || []);

    const docs = files
      .filter((file) => file.name.toLowerCase().endsWith(".docx"))
      .map((file) => ({
        id: `${file.name}-${Date.now()}-${Math.random()}`,
        projectName: cleanProjectName(file.name),
        fileName: file.name,
        url: URL.createObjectURL(file),
      }));

    setInspectionDocs((prev) => [...prev, ...docs]);
    event.target.value = "";
  }

  function exportExcel() {
    const projectKpis = projects.map(getProjectKpi);

    const summaryRows = projectKpis.map((p) => ({
      Progetto: p.projectName,
      File: p.fileName,
      "Issue totali": p.total,
      "Issue chiuse": p.closed,
      "Issue aperte": p.open,
      "% chiusura": p.closure,
      "Arrivo progetto": formatDate(p.firstDate),
      "Ultima attività": formatDate(p.lastDate),
      "Durata giorni": p.durationDays,
    }));

    const issueRows = projects.flatMap((project) =>
      project.issues.map((issue) => ({
        Progetto: project.projectName,
        File: project.fileName,
        Titolo: issue.title,
        Stato: issue.isClosed ? "Chiusa" : "Aperta",
        "Data creazione": formatDate(issue.creationDate),
        "Ultima attività": formatDate(issue.lastActivityDate),
        "Durata giorni": daysBetween(
          issue.creationDate,
          issue.lastActivityDate
        ),
        Commenti: issue.commentsCount,
        Autore: issue.author,
        GUID: issue.guid,
      }))
    );

    const inspectionRows = inspectionDocs.map((doc) => ({
      Progetto: doc.projectName,
      "Scheda ispettiva Word": doc.fileName,
    }));

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(summaryRows),
      "KPI Progetti"
    );

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(issueRows),
      "Dettaglio Issue"
    );

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(inspectionRows),
      "Schede Ispettive"
    );

    XLSX.writeFile(workbook, "dashboard-pm-report.xlsx");
  }

  const allIssues = projects.flatMap((project) => project.issues);
  const projectKpis = projects.map(getProjectKpi);

  const selectedProject = selection
    ? projects.find((p) => p.projectName === selection.projectName)
    : null;

  const selectedRows = selectedProject
    ? selectedProject.issues.filter((issue) => {
        if (selection?.type === "open") return !issue.isClosed;
        if (selection?.type === "closed") return issue.isClosed;
        return true;
      })
    : [];

  return (
    <main style={{ padding: 24, fontFamily: "Arial, sans-serif" }}>
      <button
        onClick={() => {
          window.location.href = "/";
        }}
        style={darkButton}
      >
        ← Torna alla dashboard
      </button>

      <h1>Dashboard PM</h1>

      <p>
        Carica uno o più file BCF / BCFZIP. I file riferiti allo stesso
        progetto vengono sommati in un unico KPI progetto.
      </p>

      <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
        <input
          type="file"
          multiple
          accept=".bcf,.bcfzip,.zip"
          onChange={handleFiles}
        />

        <button
          onClick={() => {
            setProjects([]);
            setSelection(null);
          }}
          disabled={projects.length === 0}
          style={{
            ...dangerButton,
            background: projects.length === 0 ? "#94a3b8" : "#dc2626",
            cursor: projects.length === 0 ? "not-allowed" : "pointer",
          }}
        >
          Svuota progetti
        </button>

        <button
          onClick={exportExcel}
          disabled={projects.length === 0 && inspectionDocs.length === 0}
          style={{
            ...successButton,
            background:
              projects.length === 0 && inspectionDocs.length === 0
                ? "#94a3b8"
                : "#16a34a",
            cursor:
              projects.length === 0 && inspectionDocs.length === 0
                ? "not-allowed"
                : "pointer",
          }}
        >
          Esporta Excel
        </button>
      </div>

      <section
        style={{
          marginTop: 18,
          border: "1px solid #e2e8f0",
          borderRadius: 14,
          padding: 16,
          background: "white",
        }}
      >
        <h2 style={{ marginTop: 0 }}>Schede ispettive Word</h2>

        <p style={{ color: "#64748b", fontSize: 14 }}>
          Carica le schede ispettive storiche in formato Word. Le schede
          vengono associate al progetto e rese disponibili come archivio
          documentale della commessa.
        </p>

        <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
          <input
            type="file"
            multiple
            accept=".docx"
            onChange={handleInspectionDocs}
          />

          <button
            onClick={() => setInspectionDocs([])}
            disabled={inspectionDocs.length === 0}
            style={{
              ...dangerButton,
              background:
                inspectionDocs.length === 0 ? "#94a3b8" : "#dc2626",
              cursor:
                inspectionDocs.length === 0 ? "not-allowed" : "pointer",
            }}
          >
            Svuota schede
          </button>
        </div>

        {inspectionDocs.length > 0 && (
          <div style={{ marginTop: 16, display: "grid", gap: 10 }}>
            {inspectionDocs.map((doc) => (
              <div
                key={doc.id}
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  gap: 12,
                  border: "1px solid #e2e8f0",
                  borderRadius: 10,
                  padding: 10,
                  background: "#f8fafc",
                }}
              >
                <div>
                  <b>{doc.fileName}</b>
                  <div style={{ fontSize: 12, color: "#64748b" }}>
                    Progetto associato: {doc.projectName}
                  </div>
                </div>

                <a
                  href={doc.url}
                  download={doc.fileName}
                  style={{
                    background: "#0f172a",
                    color: "white",
                    textDecoration: "none",
                    borderRadius: 10,
                    padding: "8px 12px",
                    fontWeight: 700,
                    fontSize: 13,
                  }}
                >
                  Scarica Word
                </a>
              </div>
            ))}
          </div>
        )}
      </section>

      {loading && <p>Caricamento file BCF...</p>}

      {!loading && projects.length > 0 && (
        <>
          <h2>Riepilogo generale</h2>

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(4, 1fr)",
              gap: 16,
              marginBottom: 28,
            }}
          >
            <KpiCard title="Progetti caricati" value={projects.length} />
            <KpiCard title="Issue totali" value={allIssues.length} />
            <KpiCard
              title="Issue chiuse"
              value={allIssues.filter((i) => i.isClosed).length}
            />
            <KpiCard
              title="% chiusura globale"
              value={
                allIssues.length
                  ? `${Math.round(
                      (allIssues.filter((i) => i.isClosed).length /
                        allIssues.length) *
                        100
                    )}%`
                  : "0%"
              }
            />
          </div>

          <h2>Grafico durata verifica</h2>

          <ProjectBarChart data={projectKpis} />

          <AverageDurationChart data={projectKpis} />

          <h2>KPI per progetto</h2>

          <div style={{ display: "grid", gap: 20 }}>
            {projectKpis.map((project) => (
              <section
                key={project.projectName}
                style={{
                  border: "1px solid #ddd",
                  borderRadius: 14,
                  padding: 18,
                  background: "white",
                }}
              >
                <h3 style={{ marginTop: 0 }}>{project.projectName}</h3>

                <div
                  style={{
                    display: "grid",
                    gridTemplateColumns: "repeat(7, 1fr)",
                    gap: 12,
                  }}
                >
                  <KpiCard
                    title="Issue"
                    value={project.total}
                    onClick={() =>
                      setSelection({
                        projectName: project.projectName,
                        type: "all",
                        label: "Tutte le issue",
                      })
                    }
                  />
                  <KpiCard
                    title="Chiuse"
                    value={project.closed}
                    onClick={() =>
                      setSelection({
                        projectName: project.projectName,
                        type: "closed",
                        label: "Issue chiuse",
                      })
                    }
                  />
                  <KpiCard
                    title="Aperte"
                    value={project.open}
                    onClick={() =>
                      setSelection({
                        projectName: project.projectName,
                        type: "open",
                        label: "Issue aperte",
                      })
                    }
                  />
                  <KpiCard title="% chiusura" value={`${project.closure}%`} />
                  <KpiCard
                    title="Arrivo progetto"
                    value={formatDate(project.firstDate)}
                  />
                  <KpiCard
                    title="Ultima attività"
                    value={formatDate(project.lastDate)}
                  />
                  <KpiCard
                    title="Durata verifica"
                    value={`${project.durationDays} gg`}
                  />
                </div>
              </section>
            ))}
          </div>

          {selection && selectedProject && (
            <section
              style={{
                marginTop: 28,
                border: "1px solid #ddd",
                borderRadius: 14,
                padding: 18,
                background: "white",
              }}
            >
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  gap: 12,
                }}
              >
                <h2 style={{ marginTop: 0 }}>
                  Report: {selection.label} - {selection.projectName}
                </h2>

                <button
                  onClick={() => setSelection(null)}
                  style={darkButton}
                >
                  Chiudi report
                </button>
              </div>

              <div style={{ overflowX: "auto" }}>
                <table
                  style={{
                    width: "100%",
                    borderCollapse: "collapse",
                    fontSize: 13,
                  }}
                >
                  <thead>
                    <tr style={{ background: "#f1f5f9" }}>
                      <th style={th}>Titolo</th>
                      <th style={th}>Stato</th>
                      <th style={th}>Data creazione</th>
                      <th style={th}>Ultima attività</th>
                      <th style={th}>Durata</th>
                      <th style={th}>Commenti</th>
                      <th style={th}>Autore</th>
                      <th style={th}>GUID</th>
                    </tr>
                  </thead>

                  <tbody>
                    {selectedRows.map((issue, index) => (
                      <tr key={`${issue.guid}-${index}`}>
                        <td style={td}>{issue.title}</td>
                        <td style={td}>
                          {issue.isClosed ? "Chiusa" : "Aperta"}
                        </td>
                        <td style={td}>{formatDate(issue.creationDate)}</td>
                        <td style={td}>
                          {formatDate(issue.lastActivityDate)}
                        </td>
                        <td style={td}>
                          {daysBetween(
                            issue.creationDate,
                            issue.lastActivityDate
                          )}{" "}
                          gg
                        </td>
                        <td style={td}>{issue.commentsCount}</td>
                        <td style={td}>{issue.author}</td>
                        <td style={td}>{issue.guid}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>
          )}
        </>
      )}
    </main>
  );
}

function ProjectBarChart({ data }: { data: ProjectKpi[] }) {
  const maxDuration = Math.max(...data.map((d) => d.durationDays), 1);

  return (
    <div
      style={{
        border: "1px solid #e2e8f0",
        borderRadius: 14,
        padding: 18,
        background: "white",
        marginBottom: 28,
      }}
    >
      <h3 style={{ marginTop: 0 }}>
        Durata temporale del progetto di verifica
      </h3>

      <div style={{ display: "grid", gap: 14 }}>
        {data.map((project) => (
          <div key={project.projectName}>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                fontSize: 13,
                marginBottom: 4,
              }}
            >
              <b>{project.projectName}</b>
              <span>{project.durationDays} gg</span>
            </div>

            <div
              style={{
                height: 18,
                background: "#e2e8f0",
                borderRadius: 999,
                overflow: "hidden",
              }}
            >
              <div
                style={{
                  width: `${Math.max(
                    4,
                    (project.durationDays / maxDuration) * 100
                  )}%`,
                  height: "100%",
                  background: "#0f172a",
                }}
              />
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function AverageDurationChart({ data }: { data: ProjectKpi[] }) {
  if (data.length === 0) return null;

  const average =
    data.reduce((sum, project) => sum + project.durationDays, 0) /
    data.length;

  const maxDuration = Math.max(
    ...data.map((project) => project.durationDays),
    average,
    1
  );

  return (
    <div
      style={{
        border: "1px solid #e2e8f0",
        borderRadius: 14,
        padding: 18,
        background: "white",
        marginBottom: 28,
      }}
    >
      <h3 style={{ marginTop: 0 }}>
        Durata progetto vs durata media
      </h3>

      <div
        style={{
          position: "relative",
          height: 320,
          borderLeft: "2px solid #cbd5e1",
          borderBottom: "2px solid #cbd5e1",
          marginTop: 40,
          marginLeft: 40,
          padding: "0 20px",
        }}
      >
        <div
          style={{
            position: "absolute",
            left: 0,
            right: 0,
            bottom: `${(average / maxDuration) * 100}%`,
            borderTop: "3px solid #ef4444",
            zIndex: 1,
          }}
        />

        <div
          style={{
            position: "absolute",
            right: 10,
            bottom: `calc(${(average / maxDuration) * 100}% + 8px)`,
            color: "#ef4444",
            fontSize: 12,
            fontWeight: 700,
          }}
        >
          Media: {Math.round(average)} gg
        </div>

        <div
          style={{
            position: "absolute",
            inset: 0,
            display: "flex",
            alignItems: "flex-end",
            justifyContent: "space-around",
            padding: "0 24px",
          }}
        >
          {data.map((project) => {
            const height = (project.durationDays / maxDuration) * 260;

            return (
              <div
                key={project.projectName}
                style={{
                  display: "flex",
                  flexDirection: "column",
                  alignItems: "center",
                  width: 100,
                }}
              >
                <div
                  style={{
                    width: 14,
                    height: 14,
                    borderRadius: "50%",
                    background: "#2563eb",
                    marginBottom: 6,
                    position: "relative",
                    bottom: `${height}px`,
                    zIndex: 2,
                  }}
                />

                <div
                  style={{
                    width: 3,
                    height: `${height}px`,
                    background: "#2563eb",
                    marginTop: `-${height}px`,
                  }}
                />

                <div
                  style={{
                    fontSize: 12,
                    fontWeight: 700,
                    marginTop: 6,
                  }}
                >
                  {project.durationDays} gg
                </div>

                <div
                  style={{
                    fontSize: 11,
                    textAlign: "center",
                    marginTop: 4,
                    wordBreak: "break-word",
                  }}
                >
                  {project.projectName}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      <div
        style={{
          display: "flex",
          gap: 18,
          marginTop: 18,
          fontSize: 13,
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <div
            style={{
              width: 18,
              height: 3,
              background: "#2563eb",
            }}
          />
          <span>Durata progetto</span>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <div
            style={{
              width: 18,
              height: 3,
              background: "#ef4444",
            }}
          />
          <span>Durata media</span>
        </div>
      </div>
    </div>
  );
}

function KpiCard({
  title,
  value,
  onClick,
}: {
  title: string;
  value: string | number;
  onClick?: () => void;
}) {
  return (
    <div
      onClick={onClick}
      style={{
        border: "1px solid #e2e8f0",
        borderRadius: 12,
        padding: 14,
        background: "#f8fafc",
        cursor: onClick ? "pointer" : "default",
      }}
    >
      <div style={{ fontSize: 12, color: "#64748b" }}>{title}</div>
      <div
        style={{
          fontSize: 24,
          fontWeight: 800,
          marginTop: 6,
        }}
      >
        {value}
      </div>
    </div>
  );
}

const th = {
  border: "1px solid #e2e8f0",
  padding: 8,
  textAlign: "left" as const,
};

const td = {
  border: "1px solid #e2e8f0",
  padding: 8,
  verticalAlign: "top" as const,
};

const darkButton = {
  marginBottom: 20,
  background: "#0f172a",
  color: "white",
  border: "none",
  borderRadius: 10,
  padding: "10px 14px",
  fontWeight: 700,
  cursor: "pointer",
};

const dangerButton = {
  color: "white",
  border: "none",
  borderRadius: 10,
  padding: "9px 12px",
  fontWeight: 700,
};

const successButton = {
  color: "white",
  border: "none",
  borderRadius: 10,
  padding: "9px 12px",
  fontWeight: 700,
};
