"use client";

import { useState } from "react";

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

export default function DashboardPMPage() {
  const [projects, setProjects] = useState<ProjectIssues[]>([]);
  const [loading, setLoading] = useState(false);

  async function handleFiles(
    event: React.ChangeEvent<HTMLInputElement>
  ) {
    const files = Array.from(event.target.files || []);

    setLoading(true);
    setProjects([]);

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

    setProjects(parsedProjects);
    setLoading(false);
  }

  const allIssues = projects.flatMap((project) => project.issues);

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
        Carica più file BCF / BCFZIP, uno per ogni progetto o commessa.
      </p>

      <input
        type="file"
        multiple
        accept=".bcf,.bcfzip,.zip"
        onChange={handleFiles}
      />

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

          <h2>KPI per progetto</h2>

          <div style={{ display: "grid", gap: 20 }}>
            {projects.map((project) => {
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
                issues.length > 0
                  ? Math.round((closed / issues.length) * 100)
                  : 0;

              return (
                <section
                  key={project.fileName}
                  style={{
                    border: "1px solid #ddd",
                    borderRadius: 14,
                    padding: 18,
                    background: "white",
                  }}
                >
                  <h3 style={{ marginTop: 0 }}>
                    {project.projectName}
                  </h3>

                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns: "repeat(7, 1fr)",
                      gap: 12,
                    }}
                  >
                    <KpiCard title="Issue" value={issues.length} />
                    <KpiCard title="Chiuse" value={closed} />
                    <KpiCard title="Aperte" value={open} />
                    <KpiCard title="% chiusura" value={`${closure}%`} />
                    <KpiCard
                      title="Arrivo progetto"
                      value={formatDate(firstDate)}
                    />
                    <KpiCard
                      title="Ultima attività"
                      value={formatDate(lastDate)}
                    />
                    <KpiCard
                      title="Durata"
                      value={`${daysBetween(firstDate, lastDate)} gg`}
                    />
                  </div>
                </section>
              );
            })}
          </div>
        </>
      )}
    </main>
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
