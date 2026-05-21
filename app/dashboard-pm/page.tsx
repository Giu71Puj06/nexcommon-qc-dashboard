"use client";

import Link from "next/link";

export default function DashboardPMPage() {
  return (
    <main style={pageStyle}>
      <div style={contentStyle}>
        <Link href="/" style={backLinkStyle}>
          ← Torna alla dashboard
        </Link>

        <h1 style={titleStyle}>Dashboard PM</h1>

        <p style={leadStyle}>
          Area strumenti per il Project Manager. Seleziona il modulo operativo da utilizzare.
        </p>

        <section style={cardStyle}>
          <h2 style={sectionTitleStyle}>Strumenti PM</h2>

          <div style={toolsGridStyle}>
            <a href="/dashboard-pm/correggi-numerazione" style={toolCardStyle}>
              <div style={{ ...badgeStyle, color: "#0284c7" }}>MODULO PM</div>
              <div style={toolTitleStyle}>Correzione numerazione schede</div>
              <div style={toolDescriptionStyle}>
                Allinea automaticamente NC e OSS tra due emissioni di schede ispettive Word,
                corregge i cronologici, elimina duplicati, riordina le tabelle ed esporta il report Excel.
              </div>
            </a>

            <a href="/dashboard-pm/tempi-verifica" style={toolCardStyle}>
              <div style={{ ...badgeStyle, color: "#16a34a" }}>MODULO KPI</div>
              <div style={toolTitleStyle}>Stima tempi medi verifica</div>
              <div style={toolDescriptionStyle}>
                Analizza file BCF e schede ispettive Word per stimare tempi medi di verifica,
                produttività ispettiva, discipline prevalenti e KPI QA/QC.
              </div>
            </a>

            <a href="/dashboard-pm/variazioni-economiche" style={toolCardStyle}>
              <div style={{ ...badgeStyle, color: "#dc2626" }}>MODULO ECONOMICO</div>
              <div style={toolTitleStyle}>Analisi variazioni economiche</div>
              <div style={toolDescriptionStyle}>
                Confronta i PDF economici della fase iniziale e finale della verifica,
                calcola incremento/decremento per commessa, delta percentuali,
                KPI economici e report esportabili.
              </div>
            </a>
          </div>
        </section>
      </div>
    </main>
  );
}

const pageStyle: React.CSSProperties = {
  minHeight: "100vh",
  padding: 30,
  fontFamily: "Arial, sans-serif",
  background: "#f1f5f9",
  color: "#0f172a",
};

const contentStyle: React.CSSProperties = {
  width: "100%",
  maxWidth: 1300,
  margin: "0 auto",
};

const backLinkStyle: React.CSSProperties = {
  display: "inline-block",
  marginBottom: 24,
  padding: "10px 14px",
  borderRadius: 10,
  background: "#0f172a",
  color: "white",
  fontWeight: 700,
  textDecoration: "none",
};

const titleStyle: React.CSSProperties = {
  fontSize: 42,
  marginBottom: 10,
};

const leadStyle: React.CSSProperties = {
  fontSize: 18,
  lineHeight: 1.5,
  marginBottom: 24,
  color: "#475569",
};

const cardStyle: React.CSSProperties = {
  background: "#ffffff",
  border: "1px solid #e2e8f0",
  borderRadius: 16,
  padding: 22,
  boxShadow: "0 8px 24px rgba(15, 23, 42, 0.06)",
};

const sectionTitleStyle: React.CSSProperties = {
  fontSize: 22,
  marginTop: 0,
  marginBottom: 18,
};

const toolsGridStyle: React.CSSProperties = {
  display: "grid",
  gridTemplateColumns: "repeat(auto-fit, minmax(360px, 1fr))",
  gap: 18,
};

const toolCardStyle: React.CSSProperties = {
  textDecoration: "none",
  color: "#0f172a",
  border: "1px solid #cbd5e1",
  borderRadius: 14,
  padding: 20,
  background: "#f8fafc",
  display: "block",
  boxShadow: "0 1px 3px rgba(0,0,0,0.05)",
};

const badgeStyle: React.CSSProperties = {
  fontSize: 13,
  fontWeight: 800,
  marginBottom: 10,
  letterSpacing: 0.4,
};

const toolTitleStyle: React.CSSProperties = {
  fontSize: 24,
  fontWeight: 800,
  marginBottom: 12,
};

const toolDescriptionStyle: React.CSSProperties = {
  fontSize: 14,
  lineHeight: 1.6,
  color: "#475569",
};
