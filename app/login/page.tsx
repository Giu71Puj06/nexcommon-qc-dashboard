"use client";

import React, { useState } from "react";

const USERS = [
  {
    username: "admin@nexcommon.it",
    password: "admin123",
    role: "Admin",
  },
  {
    username: "a.albani@itscontrollitecnici.it",
    password: "its2026",
    role: "ITS Roma",
  },
  {
    username: "v.guccione@itscontrollitecnici.it",
    password: "its2026",
    role: "ITS Roma",
  },
];

export default function LoginPage() {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  function handleLogin(e: React.FormEvent) {
    e.preventDefault();

    const user = USERS.find(
      (u) =>
        u.username.toLowerCase() === username.toLowerCase() &&
        u.password === password
    );

    if (!user) {
      setError("Credenziali non valide");
      return;
    }

    localStorage.setItem("nexcommon_verify_auth", "true");
    localStorage.setItem(
      "nexcommon_verify_user",
      JSON.stringify({
        username: user.username,
        role: user.role,
      })
    );

    window.location.href = "/";
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
          width: 420,
          background: "white",
          borderRadius: 20,
          padding: 36,
          boxShadow: "0 10px 30px rgba(15,23,42,.10)",
          border: "1px solid #e2e8f0",
        }}
      >
        <div style={{ textAlign: "center", marginBottom: 30 }}>
          <img
            src="/logo_nexcommon.png"
            alt="Nexcommon"
            style={{ height: 42, objectFit: "contain" }}
          />

          <h1
            style={{
              marginTop: 18,
              marginBottom: 4,
              fontSize: 32,
              color: "#0f172a",
            }}
          >
            Nexcommon Verify
          </h1>

          <div style={{ color: "#64748b", fontSize: 14 }}>
            Quality Control Platform
          </div>
        </div>

        <form onSubmit={handleLogin}>
          <div style={{ marginBottom: 18 }}>
            <div style={labelStyle}>Username</div>
            <input
              type="text"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              placeholder="Inserisci username"
              style={inputStyle}
            />
          </div>

          <div style={{ marginBottom: 18 }}>
            <div style={labelStyle}>Password</div>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="Inserisci password"
              style={inputStyle}
            />
          </div>

          {error && <div style={errorStyle}>{error}</div>}

          <button
            type="submit"
            style={{
              width: "100%",
              padding: 14,
              borderRadius: 12,
              border: "none",
              background: "#0f172a",
              color: "white",
              fontWeight: 700,
              cursor: "pointer",
              fontSize: 15,
            }}
          >
            Accedi alla piattaforma
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
          Nexcommon S.r.l. · QA/QC Platform
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