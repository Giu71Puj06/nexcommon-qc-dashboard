"use client";

import React from "react";
import Link from "next/link";

export default function TempiVerificaPage() {
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
      <div
        style={{
          width: "100%",
          maxWidth: 1800,
          margin: "0 auto",
        }}
      >
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

        <h1
          style={{
            fontSize: 34,
            fontWeight: 800,
            marginBottom: 10,
          }}
        >
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
          Analizza file BCF, BCFZIP e schede ispettive Word per calcolare
          durata media delle verifiche, tempi medi per rilievo,
          KPI progetto ed esportazione report XLSX.
        </p>

        {/* ===================== BCF ===================== */}

        <section
          style={{
            background: "#ffffff",
            border: "1px solid #cbd5e1",
            borderRadius: 16,
            padding: 20,
            marginBottom: 24,
          }}
        >
          <h2
            style={{
              marginTop: 0,
              marginBottom: 14,
              fontSize: 22,
              fontWeight: 800,
            }}
          >
            File BCF / BCFZIP
          </h2>

          <p
            style={{
              marginTop: 0,
              marginBottom: 14,
              color: "#64748b",
              lineHeight: 1.5,
            }}
          >
            Carica uno o più file BCF o BCFZIP.
            I file riferiti allo stesso progetto vengono aggregati
            per il calcolo dei KPI temporali della commessa.
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
              Svuota progetti
            </button>

            <button
              style={{
                padding: "12px 18px",
                border: 0,
                borderRadius: 10,
                background: "#0f172a",
                color: "#fff",
                fontWeight: 700,
              }}
            >
              Esporta XLSX
            </button>
          </div>
        </section>

        {/* ===================== WORD ===================== */}

        <section
          style={{
            background: "#ffffff",
            border: "1px solid #cbd5e1",
            borderRadius: 16,
            padding: 20,
            marginBottom: 24,
          }}
        >
          <h2
            style={{
              marginTop: 0,
              marginBottom: 14,
              fontSize: 22,
              fontWeight: 800,
            }}
          >
            Schede ispettive Word → BCF
          </h2>

          <p
            style={{
              marginTop: 0,
              marginBottom: 14,
              color: "#64748b",
              lineHeight: 1.5,
            }}
          >
            Carica le schede ispettive Word.
            Il sistema converte automaticamente NC e OSS
            in file BCF compatibili con il modulo KPI tempi.
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
                background: "#0f172a",
                color: "#fff",
                fontWeight: 700,
              }}
            >
              Genera BCF
            </button>
          </div>
        </section>

        {/* ===================== KPI ===================== */}

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
              marginBottom: 20,
              fontSize: 22,
              fontWeight: 800,
            }}
          >
            KPI tempi verifica
          </h2>

          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
              gap: 16,
            }}
          >
            {[
              "Tempo medio verifica",
              "Giorni medi per rilievo",
              "Numero issue",
              "Numero commesse",
            ].map((item) => (
              <div
                key={item}
                style={{
                  background: "#f8fafc",
                  border: "1px solid #cbd5e1",
                  borderRadius: 14,
                  padding: 20,
                }}
              >
                <div
                  style={{
                    fontSize: 14,
                    color: "#64748b",
                    marginBottom: 10,
                  }}
                >
                  {item}
                </div>

                <div
                  style={{
                    fontSize: 32,
                    fontWeight: 800,
                  }}
                >
                  -
                </div>
              </div>
            ))}
          </div>
        </section>
      </div>
    </main>
  );
}
