"use client";

import React, { useState } from "react";
import Link from "next/link";

type StatoRisultato = {
  totaleRighe?: number;
  corretti?: number;
  giaAllineati?: number;
  mancanti?: number;
  nonTrovati?: number;
};

export default function CorreggiNumerazioneSchedePage() {
  const [zipPrecedente, setZipPrecedente] = useState<File | null>(null);
  const [zipDaCorreggere, setZipDaCorreggere] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [risultato, setRisultato] = useState<StatoRisultato | null>(null);

  async function correggiNumerazione() {
    if (!zipPrecedente || !zipDaCorreggere) {
      alert("Carica lo ZIP dell'emissione precedente e lo ZIP dell'emissione da correggere.");
      return;
    }

    setLoading(true);
    setRisultato(null);

    try {
      const fd = new FormData();
      fd.append("emissione_precedente", zipPrecedente);
      fd.append("emissione_da_correggere", zipDaCorreggere);

      const res = await fetch("/api/correggi-numerazione-schede", {
        method: "POST",
        body: fd,
      });

      if (!res.ok) {
        const errorText = await res.text().catch(() => "");
        console.error("Errore API correggi-numerazione-schede:", errorText);
        alert(errorText || "Errore durante la correzione della numerazione.");
        return;
      }

      const headerStats = res.headers.get("x-correzione-stats");
      if (headerStats) {
        try {
          setRisultato(JSON.parse(headerStats));
        } catch {
          setRisultato(null);
        }
      }

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);

      const a = document.createElement("a");
      a.href = url;
      a.download = "SCHEDE_CORRETTE_NUMERAZIONE.zip";
      document.body.appendChild(a);
      a.click();
      a.remove();

      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error(err);
      alert("Errore durante la correzione della numerazione.");
    } finally {
      setLoading(false);
    }
  }

  return (
    <main style={pageStyle}>
      <div style={{ maxWidth: 980 }}>
        <Link href="/" style={backLinkStyle}>← Torna alla dashboard</Link>

        <h1 style={titleStyle}>Correzione numerazione NC/OSS</h1>

        <p style={leadStyle}>
          Modulo separato per allineare la numerazione NC/OSS tra due emissioni di schede ispettive Word.
          L'emissione da correggere viene sempre adeguata all'emissione precedente, usando come chiave il testo
          della colonna <b>Rilievi ODI</b>.
        </p>

        <section style={cardStyle}>
          <h2 style={sectionTitleStyle}>1. Carica le due emissioni</h2>

          <div style={{ display: "grid", gap: 18 }}>
            <label>
              <b>ZIP emissione precedente / riferimento</b>
              <input
                type="file"
                accept=".zip"
                onChange={(e) => setZipPrecedente(e.target.files?.[0] || null)}
                style={inputStyle}
              />
              <div style={helpStyle}>Esempio: prima emissione, seconda emissione, revisione precedente.</div>
            </label>

            <label>
              <b>ZIP emissione da correggere</b>
              <input
                type="file"
                accept=".zip"
                onChange={(e) => setZipDaCorreggere(e.target.files?.[0] || null)}
                style={inputStyle}
              />
              <div style={helpStyle}>Esempio: seconda emissione, terza emissione, revisione successiva.</div>
            </label>
          </div>
        </section>

        <section style={cardStyle}>
          <h2 style={sectionTitleStyle}>2. Regola di correzione</h2>
          <div style={ruleBoxStyle}>
            Se il testo <b>Rilievi ODI</b> è uguale, il codice NC/OSS dell'emissione da correggere viene sostituito
            con il codice NC/OSS dell'emissione precedente.
          </div>
        </section>

        <button type="button" onClick={correggiNumerazione} disabled={loading} style={buttonStyle}>
          {loading ? "Correzione in corso..." : "Correggi numerazione e scarica ZIP"}
        </button>

        {risultato && (
          <section style={cardStyle}>
            <h2 style={sectionTitleStyle}>Esito ultima elaborazione</h2>
            <div style={statsGridStyle}>
              <Stat label="Righe analizzate" value={risultato.totaleRighe} />
              <Stat label="Corrette" value={risultato.corretti} />
              <Stat label="Già allineate" value={risultato.giaAllineati} />
              <Stat label="Senza riferimento" value={(risultato.mancanti || 0) + (risultato.nonTrovati || 0)} />
            </div>
          </section>
        )}
      </div>
    </main>
  );
}

function Stat({ label, value }: { label: string; value?: number }) {
  return (
    <div style={statStyle}>
      <div style={{ fontSize: 13, color: "#64748b", marginBottom: 6 }}>{label}</div>
      <div style={{ fontSize: 28, fontWeight: 800 }}>{value ?? 0}</div>
    </div>
  );
}

const pageStyle: React.CSSProperties = {
  minHeight: "100vh",
  padding: 30,
  fontFamily: "Arial, sans-serif",
  background: "#f1f5f9",
  color: "#0f172a",
};

const backLinkStyle: React.CSSProperties = {
  display: "inline-block",
  marginBottom: 24,
  color: "#0284c7",
  fontWeight: 700,
  textDecoration: "none",
};

const titleStyle: React.CSSProperties = {
  fontSize: 42,
  marginBottom: 16,
};

const leadStyle: React.CSSProperties = {
  fontSize: 19,
  lineHeight: 1.55,
  marginBottom: 24,
  maxWidth: 860,
};

const cardStyle: React.CSSProperties = {
  background: "#ffffff",
  border: "1px solid #e2e8f0",
  borderRadius: 16,
  padding: 22,
  marginBottom: 20,
  boxShadow: "0 8px 24px rgba(15, 23, 42, 0.06)",
};

const sectionTitleStyle: React.CSSProperties = {
  fontSize: 22,
  marginTop: 0,
  marginBottom: 16,
};

const inputStyle: React.CSSProperties = {
  display: "block",
  width: "100%",
  marginTop: 8,
  padding: 12,
  border: "1px solid #cbd5e1",
  borderRadius: 10,
  background: "#fff",
};

const helpStyle: React.CSSProperties = {
  marginTop: 6,
  fontSize: 13,
  color: "#64748b",
};

const ruleBoxStyle: React.CSSProperties = {
  padding: 16,
  borderRadius: 12,
  background: "#eff6ff",
  border: "1px solid #bfdbfe",
  color: "#1e3a8a",
  lineHeight: 1.5,
};

const buttonStyle: React.CSSProperties = {
  marginTop: 6,
  marginBottom: 24,
  padding: "14px 22px",
  border: 0,
  borderRadius: 12,
  background: "#0284c7",
  color: "white",
  fontSize: 17,
  fontWeight: 800,
  cursor: "pointer",
};

const statsGridStyle: React.CSSProperties = {
  display: "grid",
  gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))",
  gap: 14,
};

const statStyle: React.CSSProperties = {
  border: "1px solid #e2e8f0",
  borderRadius: 12,
  padding: 16,
  background: "#f8fafc",
};
