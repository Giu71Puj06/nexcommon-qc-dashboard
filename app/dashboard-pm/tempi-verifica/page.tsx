"use client";

import React from "react";
import Link from "next/link";

export default function DashboardPmPage() {
  return (
    <main
      style={{
        minHeight: "100vh",
        background: "#f1f5f9",
        padding: 16,
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
          href="/"
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
          ← Torna alla dashboard
        </Link>

        <h1
          style={{
            fontSize: 28,
            fontWeight: 800,
            marginBottom: 28,
          }}
        >
          Dashboard PM
        </h1>

        <section
          style={{
            background: "#ffffff",
            border: "1px solid #cbd5e1",
            borderRadius: 16,
            padding: 20,
            marginBottom: 28,
          }}
        >
          <h2
            style={{
              marginTop: 0,
              marginBottom: 18,
              fontSize: 20,
              fontWeight: 800,
            }}
          >
            Strumenti PM
          </h2>

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fit, minmax(420px, 1fr))",
              gap: 18,
            }}
          >
            <a
              href="/dashboard-pm/correggi-numerazione"
              style={{
                textDecoration: "none",
                color: "#0f172a",
                border: "1px solid #cbd5e1",
                borderRadius: 14,
                padding: 18,
                background: "#f8fafc",
              }}
            >
              <div
                style={{
                  fontSize: 14,
                  fontWeight: 700,
                  color: "#0284c7",
                  marginBottom: 10,
                }}
              >
                MODULO PM
              </div>

              <div
                style={{
                  fontSize: 22,
                  fontWeight: 800,
                  marginBottom: 10,
                }}
              >
                Correzione numerazione schede
              </div>

              <div
                style={{
                  fontSize: 14,
                  lineHeight: 1.5,
                  color: "#475569",
                }}
              >
                Allinea automaticamente NC e OSS tra due emissioni di schede ispettive Word,
                corregge i cronologici, elimina duplicati, riordina le tabelle ed esporta il report Excel.
              </div>
            </a>

            <a
              href="/dashboard-pm/tempi-verifica"
              style={{
                textDecoration: "none",
                color: "#0f172a",
                border: "1px solid #cbd5e1",
                borderRadius: 14,
                padding: 18,
                background: "#f8fafc",
              }}
            >
              <div
                style={{
                  fontSize: 14,
                  fontWeight: 700,
                  color: "#0284c7",
                  marginBottom: 10,
                }}
              >
                MODULO PM
              </div>

              <div
                style={{
                  fontSize: 22,
                  fontWeight: 800,
                  marginBottom: 10,
                }}
              >
                Stima tempi medi verifica
              </div>

              <div
                style={{
                  fontSize: 14,
                  lineHeight: 1.5,
                  color: "#475569",
                }}
              >
                Analizza file BCF, BCFZIP e schede Word per stimare durata media delle verifiche,
                giorni medi per rilievo, KPI commessa ed esportazione report XLSX.
              </div>
            </a>
          </div>
        </section>

        <section
          style={{
            background: "#ffffff",
            border: "1px solid #cbd5e1",
            borderRadius: 16,
            padding: 20,
            marginBottom: 20,
          }}
        >
          <p
            style={{
              marginTop: 0,
              marginBottom: 12,
              fontSize: 16,
              color: "#334155",
            }}
          >
            Carica uno o più file BCF / BCFZIP. I file riferiti allo stesso progetto
            vengono sommati in un unico KPI progetto.
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
              style={{
                flex: 1,
              }}
            />

            <button
              style={{
                padding: "12px 18px",
                border: 0,
                borderRadius: 10,
                background: "#94a3b8",
                color: "#fff",
                fontWeight: 700,
              }}
            >
              Svuota progetti
            </button>

            <button
              style={{
                padding: "12px 18px",
                border: 0,
                borderRadius: 10,
                background: "#94a3b8",
                color: "#fff",
                fontWeight: 700,
              }}
            >
              Esporta Excel
            </button>
          </div>
        </section>

        <section
          style={{
            background: "#ffffff",
            border: "1px solid #cbd5e1",
            borderRadius: 16,
            padding: 20,
            marginBottom: 20,
          }}
        >
          <h2
            style={{
              marginTop: 0,
              marginBottom: 14,
              fontSize: 20,
              fontWeight: 800,
            }}
          >
            Schede ispettive Word
          </h2>

          <p
            style={{
              marginTop: 0,
              marginBottom: 12,
              color: "#64748b",
            }}
          >
            Carica le schede ispettive storiche in formato Word.
            Le schede vengono associate al progetto e rese disponibili come archivio documentale della commessa.
          </p>

          <div
            style={{
              display: "flex",
              gap: 12,
              alignItems: "center",
            }}
          >
            <input type="file" multiple style={{ flex: 1 }} />

            <button
              style={{
                padding: "12px 18px",
                border: 0,
                borderRadius: 10,
                background: "#94a3b8",
                color: "#fff",
                fontWeight: 700,
              }}
            >
              Svuota schede
            </button>

            <button
              style={{
                padding: "12px 18px",
                border: 0,
                borderRadius: 10,
                background: "#94a3b8",
                color: "#fff",
                fontWeight: 700,
              }}
            >
              Genera BCF
            </button>
          </div>
        </section>

        <section
          style={{
            background: "#ffffff",
            border: "1px solid #cbd5e1",
            borderRadius: 16,
            padding: 20,
          }}
        >
          <h2
            style={{
              marginTop: 0,
              marginBottom: 14,
              fontSize: 20,
              fontWeight: 800,
            }}
          >
            Elaborati economici
          </h2>

          <p
            style={{
              marginTop: 0,
              marginBottom: 12,
              color: "#64748b",
            }}
          >
            Carica gli elaborati economici in formato Excel delle diverse consegne.
            Il sistema legge gli importi, costruisce i trend economici e confronta la variazione
            con le issue BCF e i rilievi ispettivi.
          </p>

          <div
            style={{
              display: "flex",
              gap: 12,
              alignItems: "center",
            }}
          >
            <input type="file" multiple style={{ flex: 1 }} />

            <button
              style={{
                padding: "12px 18px",
                border: 0,
                borderRadius: 10,
                background: "#94a3b8",
                color: "#fff",
                fontWeight: 700,
              }}
            >
              Svuota costi
            </button>
          </div>
        </section>
      </div>
    </main>
  );
}
