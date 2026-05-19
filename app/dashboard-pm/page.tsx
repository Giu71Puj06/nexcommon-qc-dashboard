"use client";

import { useState } from "react";

import { readBCF } from "@/lib/dashboard-pm/read-bcf";
import {
  extractBCFHistory,
  BCFIssue,
} from "@/lib/dashboard-pm/extract-history";

function daysBetween(start: string, end: string): number {
  const startDate = new Date(start);
  const endDate = new Date(end);

  if (
    Number.isNaN(startDate.getTime()) ||
    Number.isNaN(endDate.getTime())
  ) {
    return 0;
  }

  const diff = endDate.getTime() - startDate.getTime();
  return Math.ceil(diff / (1000 * 60 * 60 * 24));
}

function formatDate(value: string): string {
  if (!value) return "-";

  const date = new Date(value);

  if (Number.isNaN(date.getTime())) return "-";

  return date.toLocaleDateString("it-IT");
}

export default function DashboardPMPage() {
  const [issues, setIssues] = useState<BCFIssue[]>([]);
  const [loading, setLoading] = useState(false);
  const [fileNames, setFileNames] = useState<string[]>([]);

  async function handleFiles(
    event: React.ChangeEvent<HTMLInputElement>
  ) {
    const files = Array.from(event.target.files || []);

    setLoading(true);
    setIssues([]);
    setFileNames(files.map((file) => file.name));

    const parsedIssues: BCFIssue[] = [];

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

      for (const markup of markups) {
        const issue = extractBCFHistory(markup.rawXml);

        if (issue) {
          parsedIssues.push(issue);
        }
      }
    }

    setIssues(parsedIssues);
    setLoading(false);
  }

  const sortedByStart = [...issues].sort(
    (a, b) =>
      new Date(a.creationDate).getTime() -
      new Date(b.creationDate).getTime()
  );

  const sortedByEnd = [...issues].sort(
    (a, b) =>
      new Date(b.lastActivityDate).getTime() -
      new Date(a.lastActivityDate).getTime()
  );

  const firstDate = sortedByStart[0]?.creationDate || "";
  const lastDate = sortedByEnd[0]?.lastActivityDate || "";

  const totalIssues = issues.length;
  const closedIssues = issues.filter((issue) => issue.isClosed).length;
  const openIssues = totalIssues - closedIssues;

  const closurePercentage =
    totalIssues > 0
      ? Math.round((closedIssues / totalIssues) * 100)
      : 0;

  const durationDays = daysBetween(firstDate, lastDate);

  return (
    <div style={{ padding: "24px", fontFamily: "Arial, sans-serif" }}>
      <h1>Dashboard PM</h1>

      <p>
        KPI temporali di commessa da file BCF / BCFZIP.
      </p>

      <input
        type="file"
        multiple
        accept=".bcf,.bcfzip,.zip"
        onChange={handleFiles}
      />

      {fileNames.length > 0 && (
        <p>
          File caricati: {fileNames.join(", ")}
        </p>
      )}

      {loading && <p>Caricamento BCF...</p>}

      {!loading && issues.length > 0 && (
        <>
          <h2>KPI commessa</h2>

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(4, 1fr)",
              gap: "16px",
              marginBottom: "24px",
            }}
          >
            <KpiCard title="Issue totali" value={totalIssues} />
            <KpiCard title="Issue chiuse" value={closedIssues} />
            <KpiCard title="Issue aperte" value={openIssues} />
            <KpiCard
              title="% chiusura"
              value={`${closurePercentage}%`}
            />
            <KpiCard
              title="Arrivo progetto"
              value={formatDate(firstDate)}
            />
            <KpiCard
              title="Ultima attività"
              value={formatDate(lastDate)}
            />
            <KpiCard
              title="Durata tracciamento"
              value={`${durationDays} giorni`}
            />
            <KpiCard
              title="Stato commessa"
              value={
                closurePercentage === 100
                  ? "Conclusa"
                  : "In corso"
              }
            />
          </div>

          <h2>Dettaglio issue</h2>

          <table
            style={{
              width: "100%",
              borderCollapse: "collapse",
              marginTop: "16px",
            }}
          >
            <thead>
              <tr>
                <th style={cellStyle}>Titolo</th>
                <th style={cellStyle}>Stato</th>
                <th style={cellStyle}>Arrivo</th>
                <th style={cellStyle}>Ultima attività</th>
                <th style={cellStyle}>Durata</th>
                <th style={cellStyle}>Commenti</th>
              </tr>
            </thead>

            <tbody>
              {issues.map((issue, index) => (
                <tr key={`${issue.guid}-${index}`}>
                  <td style={cellStyle}>{issue.title}</td>
                  <td style={cellStyle}>
                    {issue.isClosed ? "Chiusa" : "Aperta"}
                  </td>
                  <td style={cellStyle}>
                    {formatDate(issue.creationDate)}
                  </td>
                  <td style={cellStyle}>
                    {formatDate(issue.lastActivityDate)}
                  </td>
                  <td style={cellStyle}>
                    {daysBetween(
                      issue.creationDate,
                      issue.lastActivityDate
                    )}{" "}
                    giorni
                  </td>
                  <td style={cellStyle}>
                    {issue.commentsCount}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </>
      )}
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
        border: "1px solid #ddd",
        borderRadius: "12px",
        padding: "16px",
        background: "#fff",
      }}
    >
      <div style={{ fontSize: "13px", color: "#555" }}>
        {title}
      </div>
      <div
        style={{
          fontSize: "28px",
          fontWeight: "bold",
          marginTop: "8px",
        }}
      >
        {value}
      </div>
    </div>
  );
}

const cellStyle: React.CSSProperties = {
  border: "1px solid #ccc",
  padding: "8px",
};
