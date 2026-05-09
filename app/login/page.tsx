"use client";

import React, { useState } from "react";
import { supabase } from "@/lib/supabase";

const ALLOWED_DOMAIN = "@itscontrollitecnici.it";

export default function LoginPage() {
  const [mode, setMode] = useState<"login" | "register">("login");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  function isAllowedEmail(value: string) {
    return value.trim().toLowerCase().endsWith(ALLOWED_DOMAIN);
  }

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    setError("");
    setMessage("");

    const cleanEmail = email.trim().toLowerCase();
    const cleanPassword = password.trim();

    if (!isAllowedEmail(cleanEmail)) {
      setError("Puoi usare solo email @itscontrollitecnici.it");
      return;
    }

    if (!cleanPassword) {
      setError("Inserisci una password.");
      return;
    }

    setLoading(true);

    try {
      if (mode === "login") {
        const { error } = await supabase.auth.signInWithPassword({
          email: cleanEmail,
          password: cleanPassword,
        });

        if (error) {
          setError("Credenziali non valide o email non confermata.");
          return;
        }

        localStorage.setItem("nexcommon_verify_auth", "true");
        localStorage.setItem(
          "nexcommon_verify_user",
          JSON.stringify({ username: cleanEmail, role: "ITS" })
        );

        window.location.href = "/";
        return;
      }

      const { error } = await supabase.auth.signUp({
        email: cleanEmail,
        password: cleanPassword,
      });

      if (error) {
        setError(error.message || "Errore durante la registrazione.");
        return;
      }

      setMessage(
        "Registrazione inviata. Controlla la tua email aziendale per confermare l’account."
      );
      setMode("login");
      setPassword("");
    } finally {
      setLoading(false);
    }
  }

  return (
    <main
      style={{
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background: "#f1f5f9",
        fontFamily: "Arial, sans-serif",
      }}
    >
      <div
        style={{
          width: 440,
          background: "white",
          borderRadius: 20,
          padding: 36,
          boxShadow: "0 10px 30px rgba(15,23,42,.10)",
          border: "1px solid #e2e8f0",
        }}
      >
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <img
            src="/logo_nexcommon.png"
            alt="Nexcommon"
            style={{ height: 42, objectFit: "contain" }}
          />

          <h1 style={{ marginTop: 18, marginBottom: 4, fontSize: 32, color: "#0f172a" }}>
            Nexcommon Verify
          </h1>

          <div style={{ color: "#64748b", fontSize: 14 }}>
            Quality Control Platform
          </div>
        </div>

        <div
          style={{
            display: "grid",
            gridTemplateColumns: "1fr 1fr",
            gap: 8,
            marginBottom: 22,
          }}
        >
          <button
            type="button"
            onClick={() => {
              setMode("login");
              setError("");
              setMessage("");
            }}
            style={{
              padding: 10,
              borderRadius: 10,
              border: "1px solid #cbd5e1",
              background: mode === "login" ? "#0f172a" : "white",
              color: mode === "login" ? "white" : "#0f172a",
              fontWeight: 700,
              cursor: "pointer",
            }}
          >
            Accedi
          </button>

          <button
            type="button"
            onClick={() => {
              setMode("register");
              setError("");
              setMessage("");
            }}
            style={{
              padding: 10,
              borderRadius: 10,
              border: "1px solid #cbd5e1",
              background: mode === "register" ? "#0f172a" : "white",
              color: mode === "register" ? "white" : "#0f172a",
              fontWeight: 700,
              cursor: "pointer",
            }}
          >
            Registrati
          </button>
        </div>

        <form onSubmit={handleSubmit}>
          <div style={{ marginBottom: 18 }}>
            <div style={labelStyle}>Email aziendale</div>
            <input
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              placeholder="nome.cognome@itscontrollitecnici.it"
              style={inputStyle}
            />
          </div>

          <div style={{ marginBottom: 18 }}>
            <div style={labelStyle}>Password</div>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder={mode === "login" ? "Inserisci password" : "Crea una password"}
              style={inputStyle}
            />
          </div>

          {error && <div style={errorStyle}>{error}</div>}
          {message && <div style={successStyle}>{message}</div>}

          <button
            type="submit"
            disabled={loading}
            style={{
              width: "100%",
              padding: 14,
              borderRadius: 12,
              border: "none",
              background: loading ? "#334155" : "#0f172a",
              color: "white",
              fontWeight: 700,
              cursor: loading ? "not-allowed" : "pointer",
              fontSize: 15,
            }}
          >
            {loading
              ? "Attendere..."
              : mode === "login"
              ? "Accedi alla piattaforma"
              : "Crea account"}
          </button>
        </form>

        <div
          style={{
            marginTop: 24,
            paddingTop: 18,
            borderTop: "1px solid #e2e8f0",
            fontSize: 12,
            color: "#64748b",
            textAlign: "center",
          }}
        >
          Registrazione consentita solo con email @itscontrollitecnici.it
        </div>
      </div>
    </main>
  );
}

const labelStyle: React.CSSProperties = {
  marginBottom: 6,
  fontSize: 13,
  fontWeight: 700,
  color: "#334155",
};

const inputStyle: React.CSSProperties = {
  width: "100%",
  padding: 12,
  borderRadius: 12,
  border: "1px solid #cbd5e1",
  fontSize: 14,
};

const errorStyle: React.CSSProperties = {
  marginBottom: 18,
  background: "#fee2e2",
  border: "1px solid #fecaca",
  color: "#991b1b",
  padding: 10,
  borderRadius: 10,
  fontSize: 13,
};

const successStyle: React.CSSProperties = {
  marginBottom: 18,
  background: "#dcfce7",
  border: "1px solid #bbf7d0",
  color: "#166534",
  padding: 10,
  borderRadius: 10,
  fontSize: 13,
};