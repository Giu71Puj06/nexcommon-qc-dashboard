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
    .replace(".bcfzip", "")
    .replace(".bcf", "")
    .replace(".zip", "");
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
  const [loading, setLoading] = useState(false);

  async function handleFiles(
    event: React.ChangeEvent<HTMLInputElement>
  ) {
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

    setProjects((prev) => [...prev, ...parsedProjects]);
    setLoading(false);

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

    const workbook = XLSX.utils.book_new();

    const summarySheet = XLSX.utils.json_to_sheet(summaryRows);
    const issueSheet = XLSX.utils.json_to_sheet(issueRows);

    XLSX.utils.book_append_sheet(workbook, summarySheet, "KPI Progetti");
    XLSX.utils.book_append_sheet(workbook, issueSheet, "Dettaglio Issue");

    XLSX.writeFile(workbook, "dashboard-pm-report.xlsx");
  }

  const allIssues = projects.flatMap((project) => project.issues);
  const projectKpis = projects.map(getProjectKpi);

  return (
    <main style={{ padding: 24, fontFamily: "Arial, sans-serif" }}>
      <button
        onClick={() => {
          window.location.href = "/";
        }}
        style={{
          marginBottom: 20,
          background: "#0f172a",
          color: "white",
          border: "none",
          borderRadius: 10,
          padding: "10px 14px",
          fontWeight: 700,
          cursor: "pointer",
        }}
      >
        ← Torna alla dashboard
      </button>

      <h1>Dashboard PM</h1>

      <p>
        Carica uno o più file BCF / BCFZIP. Ogni nuovo caricamento si
        somma ai progetti già presenti.
      </p>

      <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
        <input
          type="file"
          multiple
          accept=".bcf,.bcfzip,.zip"
          onChange={handleFiles}
        />

        <button
          onClick={() => setProjects([])}
          disabled={projects.length === 0}
          style={{
            background: projects.length === 0 ? "#94a3b8" : "#dc2626",
            color: "white",
            border: "none",
            borderRadius: 10,
            padding: "9px 12px",
            fontWeight: 700,
            cursor: projects.length === 0 ? "not-allowed" : "pointer",
          }}
        >
          Svuota progetti
        </button>

        <button
          onClick={exportExcel}
          disabled={projects.length === 0}
          style={{
            background: projects.length === 0 ? "#94a3b8" : "#16a34a",
            color: "white",
            border: "none",
            borderRadius: 10,
            padding: "9px 12px",
            fontWeight: 700,
            cursor: projects.length === 0 ? "not-allowed" : "pointer",
          }}
        >
          Esporta Excel
        </button>
      </div>

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

          <h2>Grafico sintetico</h2>

          <ProjectBarChart data={projectKpis} />

          <h2>KPI per progetto</h2>

          <div style={{ display: "grid", gap: 20 }}>
            {projectKpis.map((project, projectIndex) => (
              <section
                key={`${project.fileName}-${projectIndex}`}
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
                  <KpiCard title="Issue" value={project.total} />
                  <KpiCard title="Chiuse" value={project.closed} />
                  <KpiCard title="Aperte" value={project.open} />
                  <KpiCard
                    title="% chiusura"
                    value={`${project.closure}%`}
                  />
                  <KpiCard
                    title="Arrivo progetto"
                    value={formatDate(project.firstDate)}
                  />
                  <KpiCard
                    title="Ultima attività"
                    value={formatDate(project.lastDate)}
                  />
                  <KpiCard
                    title="Durata"
                    value={`${project.durationDays} gg`}
                  />
                </div>
              </section>
            ))}
          </div>
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
        Durata tracciamento per progetto
      </h3>

      <div style={{ display: "grid", gap: 14 }}>
        {data.map((project) => (
          <div key={project.fileName}>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                fontSize: 13,
                marginBottom: 4,
              }}
            >
              <b>{project.projectName}</b>
              <span>
                {project.durationDays} gg · {project.closure}% chiuso
              </span>
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

function KpiCard({
  title,
  value,
}: {
  title: string;
  value: string | number;
}) {
  return (
    <div
      style={{
        border: "1px solid #e2e8f0",
        borderRadius: 12,
        padding: 14,
        background: "#f8fafc",
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
