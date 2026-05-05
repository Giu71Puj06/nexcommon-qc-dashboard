"use client";

import Link from "next/link";
import { useState } from "react";
import * as XLSX from "xlsx";

export default function ProjectsPage() {
  const [message, setMessage] = useState("");
  const [rows, setRows] = useState<any[]>([]);

  async function handleExcelUpload(file: File | null) {
    if (!file) return;

    try {
      setMessage("Caricamento Excel in corso...");

      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json<any>(sheet);

      sessionStorage.setItem("trimbleTodos", JSON.stringify(data));

      setRows(data);
      setMessage(`Excel caricato correttamente: ${file.name} — ${data.length} righe lette`);
    } catch (error) {
      console.error(error);
      setMessage("Errore nel caricamento Excel");
    }
  }

  function goToDashboard() {
    if (rows.length === 0) {
      alert("Carica prima un file Excel");
      return;
    }

    window.location.href = "/dashboard";
  }

  const columns = rows.length > 0 ? Object.keys(rows[0]).slice(0, 8) : [];

  return (
    <main style={{ padding: 40, fontFamily: "Arial", background: "#f5f7fa", minHeight: "100vh" }}>
      <Link href="/" style={{ color: "#00796b", textDecoration: "none" }}>
        ← Torna alla home
      </Link>

      <h1>Seleziona progetto</h1>
      <p>Carica il file Excel ToDo esportato da Trimble per visualizzare i dati.</p>

      <section style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 20, marginTop: 30 }}>
        <div style={cardStyle}>
          <h3>ACDat</h3>
          <p>Collega archivio ufficiale della commessa.</p>
          <button style={disabledBtn} disabled>Prossimamente</button>
        </div>

        <div style={cardStyle}>
          <h3>Trimble Connect</h3>
          <p>Importa ToDo, BCF, gruppi, tag e discipline.</p>
          <button style={disabledBtn} disabled>Prossimamente</button>
        </div>

        <div style={cardStyle}>
          <h3>Import Excel ToDo</h3>
          <p>Carica il file Excel esportato da Trimble.</p>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => handleExcelUpload(e.target.files?.[0] ?? null)}
          />
        </div>

        <div style={cardStyle}>
          <h3>Import BCF immagini</h3>
          <p>Carica file .bcfzip per immagini e viewpoint.</p>
          <input type="file" accept=".bcfzip,.zip" />
        </div>
      </section>

      {message && (
        <p style={{ marginTop: 20, color: "#00796b", fontWeight: 700 }}>
          {message}
        </p>
      )}

      {rows.length > 0 && (
        <section style={{ marginTop: 30, background: "white", padding: 24, borderRadius: 14, border: "1px solid #e2e8f0" }}>
          <h2>Anteprima dati Excel</h2>
          <p>Mostro le prime 10 righe e le prime 8 colonne.</p>

          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", marginTop: 16 }}>
              <thead>
                <tr>
                  {columns.map((col) => (
                    <th key={col} style={thStyle}>{col}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.slice(0, 10).map((row, index) => (
                  <tr key={index}>
                    {columns.map((col) => (
                      <td key={col} style={tdStyle}>
                        {String(row[col] ?? "")}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      <button
        onClick={goToDashboard}
        style={{
          marginTop: 30,
          padding: "14px 22px",
          background: rows.length > 0 ? "#00796b" : "#9ca3af",
          color: "white",
          border: "none",
          borderRadius: 8,
          cursor: "pointer",
          fontSize: 16,
        }}
      >
        Apri dashboard progetto
      </button>
    </main>
  );
}

const cardStyle = {
  background: "white",
  padding: 24,
  borderRadius: 14,
  border: "1px solid #e2e8f0",
};

const disabledBtn = {
  padding: "10px 14px",
  background: "#9ca3af",
  color: "white",
  border: "none",
  borderRadius: 8,
};

const thStyle = {
  textAlign: "left" as const,
  padding: 10,
  borderBottom: "1px solid #e2e8f0",
  fontSize: 13,
};

const tdStyle = {
  padding: 10,
  borderBottom: "1px solid #e2e8f0",
  fontSize: 13,
  maxWidth: 240,
};