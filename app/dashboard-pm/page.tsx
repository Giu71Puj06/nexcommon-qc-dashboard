"use client";

import { useState } from "react";

import { readBCF } from "@/lib/dashboard-pm/read-bcf";
import {
  extractBCFHistory,
  BCFIssue,
} from "@/lib/dashboard-pm/extract-history";

export default function DashboardPMPage() {
  const [issues, setIssues] = useState<BCFIssue[]>([]);
  const [loading, setLoading] = useState(false);

  async function handleFiles(
    event: React.ChangeEvent<HTMLInputElement>
  ) {
    const files = event.target.files;

    if (!files) return;

    setLoading(true);

    const parsedIssues: BCFIssue[] = [];

    for (const file of Array.from(files)) {
      if (!file.name.toLowerCase().endsWith(".bcfzip")) {
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

  return (
    <div style={{ padding: "24px" }}>
      <h1>Dashboard PM</h1>

      <p>Analisi timeline issue BCF</p>

      <input
        type="file"
        multiple
        accept=".bcfzip"
        onChange={handleFiles}
      />

      {loading && <p>Caricamento BCF...</p>}

      {!loading && issues.length > 0 && (
        <div style={{ marginTop: "24px" }}>
          <h2>Issue trovate: {issues.length}</h2>

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
                <th style={cellStyle}>Status</th>
                <th style={cellStyle}>Creation</th>
                <th style={cellStyle}>Modified</th>
                <th style={cellStyle}>Commenti</th>
              </tr>
            </thead>

            <tbody>
              {issues.map((issue) => (
                <tr key={issue.guid}>
                  <td style={cellStyle}>{issue.title}</td>
                  <td style={cellStyle}>{issue.status}</td>
                  <td style={cellStyle}>{issue.creationDate}</td>
                  <td style={cellStyle}>{issue.modifiedDate}</td>
                  <td style={cellStyle}>
                    {issue.commentsCount}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

const cellStyle: React.CSSProperties = {
  border: "1px solid #ccc",
  padding: "8px",
};
