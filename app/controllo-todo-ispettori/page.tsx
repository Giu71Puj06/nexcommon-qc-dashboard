"use client";

import React, { useMemo, useState } from "react";
import Link from "next/link";
import * as XLSX from "xlsx";

type CheckRow = {
  rowNumber: number;
  progressivo: number;
  label: string;
  title: string;
  codiceTitleTrimble: string;
  codiceReport: string;
  titleOk: boolean;
  tags: string;
  tagsOk: boolean;
  disciplina: string;
  disciplinaOk: boolean;
  status: string;
  statusOk: boolean;
  ispettoreTodo: string;
  ispettoriAssegnati: string;
  schedaAssegnata: string;
  assegnazioneOk: boolean;
  verificaAssegnazione: "OK" | "NON ASSEGNATO" | "NON VERIFICATO" | "ISPETTORE NON COERENTE" | "SCHEDA ERRATA";
  esito: "OK" | "ERRORE";
  anomalie: string[];
};

type AssignmentRow = {
  codice: string;
  codiceNorm: string;
  titolo: string;
  ispettori: string;
  scheda: string;
  disciplina: string;
  verificato: boolean;
  anomalie: string[];
};

type AssignmentCheckRow = AssignmentRow & {
  codiceReport: string;
  presenteInTodo: boolean;
  presenteInReport: boolean;
  esito: "OK" | "ERRORE";
};

type Filters = {
  n: string;
  codiceReport: string;
  codiceTitleTrimble: string;
  esitoCodice: string;
  tags: string;
  disciplina: string;
  status: string;
  ispettoriAssegnati: string;
  schedaAssegnata: string;
  verificaAssegnazione: string;
  esito: string;
  anomalie: string;
};

type AssignmentFilters = {
  codice: string;
  titolo: string;
  ispettori: string;
  scheda: string;
  disciplina: string;
  presenteInTodo: string;
  presenteInReport: string;
  esito: string;
  anomalie: string;
};

const TAGS_AMMESSI = ["NC", "OSS", "Nessun rilievo", "Da NC a OSS"];
const STATUS_AMMESSI = ["New", "Closed"];

function normalizeCode(value: string) {
  return String(value || "")
    .toUpperCase()
    .replace(/\.PDF$/i, "")
    .replace(/\s+/g, "")
    .replace(/[_\-.]/g, "")
    .trim();
}

function normalizeText(value: string) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function cleanPdf(value: string) {
  return String(value || "").replace(/\.pdf$/i, "").trim();
}

function looksLikeElaboratoCode(value: string) {
  const text = String(value || "").trim();

  return /^PV\d{3}-[A-Z0-9]+-[A-Z0-9]+-[A-Z]{3}-\d{5}-[A-Z]{3}-\d{6}/i.test(text);
}

function normalizeInspectorList(value: string) {
  return String(value || "")
    .replace(/\s+-\s+/g, ";")
    .replace(/\s*;\s*/g, ";")
    .replace(/\s*,\s*/g, ";")
    .replace(/\s*\/\s*/g, ";")
    .replace(/\s*\|\s*/g, ";")
    .trim();
}

function disciplinaCodeFromCodice(value: string) {
  const match = String(value || "").match(/^PV\d{3}-[A-Z0-9]+-[A-Z0-9]+-([A-Z]{3})-/i);

  return match ? match[1].toUpperCase() : "";
}

function getAssignmentSchedaDisciplina(value: string) {
  const text = String(value || "").trim();
  const parts = text.split(/\s+-\s+/);

  return parts.length > 1 ? parts.slice(1).join(" - ").trim() : text;
}

function getCell(row: any[], indexes: number[]) {
  for (const i of indexes) {
    const value = String(row[i] || "").trim();
    if (value) return value;
  }
  return "";
}

function findHeaderIndex(headers: any[], names: string[], fallback = -1) {
  const normalizedNames = names.map(normalizeText);

  const found = headers.findIndex((h) => {
    const header = normalizeText(String(h || ""));
    return normalizedNames.some((n) => header === n || header.includes(n));
  });

  return found >= 0 ? found : fallback;
}

function isDescrizioneGenerale(value: string) {
  const text = String(value || "").trim();
  const normalized = normalizeCode(text);

  if (!text) return false;
  if (/\.pdf/i.test(text)) return false;

  const hasNumber = /\d/.test(text);
  const hasCodeSeparators = /[-_]/.test(text);
  const looksLikeCode = hasNumber && (hasCodeSeparators || normalized.length > 8);

  return !looksLikeCode;
}

function findBestReportCodeFromTitle(
  title: string,
  reportCodes: Map<string, string>
) {
  const normalizedTitle = normalizeCode(title);

  let bestNormalized = "";
  let bestOriginal = "";

  for (const [normalizedReportCode, originalReportCode] of reportCodes.entries()) {
    if (
      normalizedTitle === normalizedReportCode ||
      normalizedTitle.includes(normalizedReportCode)
    ) {
      if (normalizedReportCode.length > bestNormalized.length) {
        bestNormalized = normalizedReportCode;
        bestOriginal = originalReportCode;
      }
    }
  }

  return bestOriginal;
}

function findAssignmentFromCode(
  code: string,
  assignments: Map<string, AssignmentRow>
) {
  const normalizedCode = normalizeCode(code);

  if (!normalizedCode) return null;

  if (assignments.has(normalizedCode)) {
    return assignments.get(normalizedCode) || null;
  }

  let bestKey = "";
  let bestAssignment: AssignmentRow | null = null;

  for (const [assignmentCode, assignment] of assignments.entries()) {
    if (
      normalizedCode === assignmentCode ||
      normalizedCode.includes(assignmentCode) ||
      assignmentCode.includes(normalizedCode)
    ) {
      if (assignmentCode.length > bestKey.length) {
        bestKey = assignmentCode;
        bestAssignment = assignment;
      }
    }
  }

  return bestAssignment;
}

function ispettoreCoerente(todoValue: string, assignedValue: string) {
  const todo = normalizeText(todoValue);
  const assigned = normalizeText(assignedValue);

  if (!assigned) return false;

  // Nel file PM la dicitura "Tutti" significa che qualunque ispettore del ToDo
  // e' coerente con l'assegnazione.
  if (assigned === "tutti" || assigned.includes("tutti")) return true;

  if (!todo) return true;

  const assignedParts = normalizeInspectorList(assignedValue)
    .split(";")
    .map((p) => normalizeText(p))
    .filter(Boolean);

  if (assignedParts.length === 0) return true;

  return assignedParts.some((p) => todo.includes(p) || p.includes(todo));
}

async function readXlsxRows(file: File, preferredSheetName?: string) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });

  const sheetName =
    preferredSheetName && workbook.SheetNames.includes(preferredSheetName)
      ? preferredSheetName
      : workbook.SheetNames[0];

  const sheet = workbook.Sheets[sheetName];

  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
  }) as any[][];
}

function esitoIcon(ok: boolean) {
  return ok ? "✅" : "❌";
}

export default function ControlloTodoIspettoriPage() {
  const [todoFile, setTodoFile] = useState<File | null>(null);
  const [reportFile, setReportFile] = useState<File | null>(null);
  const [elencoFile, setElencoFile] = useState<File | null>(null);
  const [ispettoriFile, setIspettoriFile] = useState<File | null>(null);

  const [todoRows, setTodoRows] = useState<any[][]>([]);
  const [reportRows, setReportRows] = useState<any[][]>([]);
  const [elencoRows, setElencoRows] = useState<any[][]>([]);
  const [ispettoriRows, setIspettoriRows] = useState<any[][]>([]);

  const [loading, setLoading] = useState(false);

  const [filters, setFilters] = useState<Filters>({
    n: "",
    codiceReport: "",
    codiceTitleTrimble: "",
    esitoCodice: "",
    tags: "",
    disciplina: "",
    status: "",
    ispettoriAssegnati: "",
    schedaAssegnata: "",
    verificaAssegnazione: "",
    esito: "",
    anomalie: "",
  });

  const [assignmentFilters, setAssignmentFilters] = useState<AssignmentFilters>({
    codice: "",
    titolo: "",
    ispettori: "",
    scheda: "",
    disciplina: "",
    presenteInTodo: "",
    presenteInReport: "",
    esito: "",
    anomalie: "",
  });

  function updateFilter(key: keyof Filters, value: string) {
    setFilters((prev) => ({
      ...prev,
      [key]: value,
    }));
  }

  function updateAssignmentFilter(key: keyof AssignmentFilters, value: string) {
    setAssignmentFilters((prev) => ({
      ...prev,
      [key]: value,
    }));
  }

  async function eseguiControllo() {
    if (!todoFile || !reportFile || !elencoFile) {
      alert("Carica ToDo XLSX, Report_Completo.xlsx ed ELENCO_ELABORATI.xlsx.");
      return;
    }

    setLoading(true);

    try {
      const todo = await readXlsxRows(todoFile);
      const report = await readXlsxRows(reportFile, "Verifica Elaborati");
      const elenco = await readXlsxRows(elencoFile);
      const ispettori = ispettoriFile ? await readXlsxRows(ispettoriFile) : [];

      setTodoRows(todo);
      setReportRows(report);
      setElencoRows(elenco);
      setIspettoriRows(ispettori);
    } catch (error) {
      console.error(error);
      alert("Errore durante la lettura dei file XLSX.");
    } finally {
      setLoading(false);
    }
  }

  const reportCodes = useMemo(() => {
    const map = new Map<string, string>();

    reportRows.slice(1).forEach((row) => {
      const codiceReport = String(row[2] || "").trim();
      const normalized = normalizeCode(codiceReport);

      if (normalized && normalized !== "NAN") {
        map.set(normalized, codiceReport);
      }
    });

    return map;
  }, [reportRows]);

  const disciplineAmmesse = useMemo(() => {
    const set = new Set<string>();

    elencoRows.slice(1).forEach((row) => {
      const disciplina = String(row[5] || "").trim();

      if (disciplina) {
        set.add(disciplina.toLowerCase());
      }
    });

    return set;
  }, [elencoRows]);

  const assignments = useMemo(() => {
    const map = new Map<string, AssignmentRow>();

    if (ispettoriRows.length === 0) return map;

    let currentDisciplinaSezione = "";

    // Struttura file PM/Ispettori.xlsx:
    // A = Codice Elaborato oppure titolo sezione/disciplina in azzurro
    // B = Titolo elaborato
    // C = Commessa
    // D = Fase
    // E = Origine
    // F = Disciplina sintetica, es. GEN, STR, SIC
    // G = WBS
    // H = Tipo
    // I = Progressivo
    // J = Revisione
    // K = Commenti PM
    // L = Ispettori assegnati
    // M = Scheda ispettiva
    ispettoriRows.forEach((row) => {
      const colA = String(row[0] || "").trim();
      const colB = String(row[1] || "").trim();

      if (!colA) return;

      const normalizedA = normalizeText(colA);
      if (
        normalizedA.includes("codice elaborato") ||
        normalizedA.includes("elenco elaborati") ||
        normalizedA.includes("intervento")
      ) {
        return;
      }

      // Le righe sezione hanno la disciplina in colonna A e titolo vuoto in colonna B.
      if (!looksLikeElaboratoCode(colA)) {
        if (!colB) currentDisciplinaSezione = colA;
        return;
      }

      const codice = colA;
      const codiceNorm = normalizeCode(codice);

      if (!codiceNorm) return;

      const disciplinaCodice = String(row[5] || "").trim();
      const ispettori = String(row[11] || "").trim();
      const scheda = String(row[12] || "").trim();

      map.set(codiceNorm, {
        codice,
        codiceNorm,
        titolo: colB,
        ispettori,
        scheda,
        disciplina:
          getAssignmentSchedaDisciplina(scheda) ||
          currentDisciplinaSezione ||
          disciplinaCodice ||
          disciplinaCodeFromCodice(codice),
        verificato: false,
        anomalie: [],
      });
    });

    return map;
  }, [ispettoriRows]);

  const checks: CheckRow[] = useMemo(() => {
    return todoRows.slice(1).map((row, index) => {
      const label = String(row[0] || "").trim();
      const title = String(row[1] || "").trim();
      const status = String(row[5] || "").trim();
      const ispettoreTodo = getCell(row, [6, 7, 10, 11]);
      const disciplina = String(row[8] || "").trim();
      const tags = String(row[9] || "").trim();

      const anomalie: string[] = [];

      const codiceTitleTrimble = cleanPdf(title);
      const titleContienePdf = /\.pdf/i.test(title);
      const titleDescrittivo = isDescrizioneGenerale(title);
      const codiceReport = titleDescrittivo
        ? ""
        : findBestReportCodeFromTitle(title, reportCodes);

      let titleOk = false;

      if (!title) {
        anomalie.push("Title mancante");
      } else if (titleContienePdf) {
        anomalie.push("Il Title contiene .pdf");
      } else if (titleDescrittivo) {
        titleOk = true;
      } else if (!codiceReport) {
        anomalie.push("Codice elaborato non presente nel Report_Completo.xlsx");
      } else {
        titleOk = true;
      }

      const tagsOk = TAGS_AMMESSI.some(
        (tag) => tag.toLowerCase() === tags.toLowerCase()
      );

      if (!tags) anomalie.push("Tags mancanti");
      else if (!tagsOk) anomalie.push("Tags non valido");

      const disciplinaOk =
        !!disciplina && disciplineAmmesse.has(disciplina.toLowerCase());

      if (!disciplina) anomalie.push("Disciplina mancante");
      else if (!disciplinaOk)
        anomalie.push("Disciplina non presente in ELENCO_ELABORATI.xlsx");

      const statusOk = STATUS_AMMESSI.some(
        (s) => s.toLowerCase() === status.toLowerCase()
      );

      if (!status) anomalie.push("Status mancante");
      else if (!statusOk) anomalie.push("Status non valido");

      const assignment =
        findAssignmentFromCode(codiceReport || codiceTitleTrimble, assignments) ||
        null;

      let verificaAssegnazione: CheckRow["verificaAssegnazione"] = "OK";
      let assegnazioneOk = true;

      if (assignments.size > 0) {
        if (!assignment) {
          verificaAssegnazione = "NON ASSEGNATO";
          assegnazioneOk = false;
          anomalie.push("Elaborato non presente nel file assegnazioni PM/Ispettori");
        } else if (!ispettoreCoerente(ispettoreTodo, assignment.ispettori)) {
          verificaAssegnazione = "ISPETTORE NON COERENTE";
          assegnazioneOk = false;
          anomalie.push(
            `Ispettore ToDo non coerente con assegnazione PM: ${assignment.ispettori}`
          );
        } else {
          // La coerenza della scheda PM viene derivata dal codice elaborato e
          // dall'assegnazione trovata. Non confrontiamo il testo esteso della
          // disciplina ToDo con il codice sintetico del file PM (GEN, STR, SIC...),
          // per evitare falsi errori.
        }
      }

      const esito =
        titleOk && tagsOk && disciplinaOk && statusOk && assegnazioneOk
          ? "OK"
          : "ERRORE";

      return {
        rowNumber: index + 2,
        progressivo: index + 1,
        label,
        title,
        codiceTitleTrimble,
        codiceReport,
        titleOk,
        tags,
        tagsOk,
        disciplina,
        disciplinaOk,
        status,
        statusOk,
        ispettoreTodo,
        ispettoriAssegnati: assignment?.ispettori || "",
        schedaAssegnata: assignment?.scheda || assignment?.disciplina || "",
        assegnazioneOk,
        verificaAssegnazione,
        esito,
        anomalie,
      };
    });
  }, [todoRows, reportCodes, disciplineAmmesse, assignments]);

  const assignmentChecks: AssignmentCheckRow[] = useMemo(() => {
    if (assignments.size === 0) return [];

    const todoCodes = new Set<string>();
    const reportCodeSet = new Set<string>();

    checks.forEach((row) => {
      const normalized = normalizeCode(row.codiceReport || row.codiceTitleTrimble);
      if (normalized) todoCodes.add(normalized);
    });

    reportCodes.forEach((_original, normalized) => {
      if (normalized) reportCodeSet.add(normalized);
    });

    return Array.from(assignments.values()).map((assignment) => {
      const presenteInTodo = Array.from(todoCodes).some(
        (code) =>
          code === assignment.codiceNorm ||
          code.includes(assignment.codiceNorm) ||
          assignment.codiceNorm.includes(code)
      );

      const presenteInReport = Array.from(reportCodeSet).some(
        (code) =>
          code === assignment.codiceNorm ||
          code.includes(assignment.codiceNorm) ||
          assignment.codiceNorm.includes(code)
      );

      const anomalie: string[] = [];

      if (!presenteInTodo) {
        anomalie.push("Elaborato assegnato dal PM ma non presente nel ToDo");
      }

      if (!presenteInReport) {
        anomalie.push("Elaborato assegnato dal PM ma non presente nel Report_Completo");
      }

      const esito = presenteInTodo && presenteInReport ? "OK" : "ERRORE";

      return {
        ...assignment,
        codiceReport: presenteInReport ? assignment.codice : "",
        presenteInTodo,
        presenteInReport,
        esito,
        anomalie,
      };
    });
  }, [assignments, checks, reportCodes]);

  const filteredChecks = useMemo(() => {
    return checks.filter((row) => {
      const n = `${row.progressivo}${row.label ? ` (${row.label})` : ""}`;
      const esitoCodice = row.titleOk ? "OK" : "ERRORE";
      const anomalie = row.anomalie.join(" | ");

      return (
        n.toLowerCase().includes(filters.n.toLowerCase()) &&
        row.codiceReport
          .toLowerCase()
          .includes(filters.codiceReport.toLowerCase()) &&
        row.codiceTitleTrimble
          .toLowerCase()
          .includes(filters.codiceTitleTrimble.toLowerCase()) &&
        (filters.esitoCodice === "" ||
          esitoCodice === filters.esitoCodice) &&
        row.tags.toLowerCase().includes(filters.tags.toLowerCase()) &&
        row.disciplina
          .toLowerCase()
          .includes(filters.disciplina.toLowerCase()) &&
        row.status.toLowerCase().includes(filters.status.toLowerCase()) &&
        row.ispettoriAssegnati
          .toLowerCase()
          .includes(filters.ispettoriAssegnati.toLowerCase()) &&
        row.schedaAssegnata
          .toLowerCase()
          .includes(filters.schedaAssegnata.toLowerCase()) &&
        (filters.verificaAssegnazione === "" ||
          row.verificaAssegnazione === filters.verificaAssegnazione) &&
        (filters.esito === "" || row.esito === filters.esito) &&
        anomalie.toLowerCase().includes(filters.anomalie.toLowerCase())
      );
    });
  }, [checks, filters]);

  const filteredAssignmentChecks = useMemo(() => {
    return assignmentChecks.filter((row) => {
      const anomalie = row.anomalie.join(" | ");

      return (
        row.codice.toLowerCase().includes(assignmentFilters.codice.toLowerCase()) &&
        row.titolo.toLowerCase().includes(assignmentFilters.titolo.toLowerCase()) &&
        row.ispettori
          .toLowerCase()
          .includes(assignmentFilters.ispettori.toLowerCase()) &&
        row.scheda.toLowerCase().includes(assignmentFilters.scheda.toLowerCase()) &&
        row.disciplina
          .toLowerCase()
          .includes(assignmentFilters.disciplina.toLowerCase()) &&
        (assignmentFilters.presenteInTodo === "" ||
          (row.presenteInTodo ? "SI" : "NO") === assignmentFilters.presenteInTodo) &&
        (assignmentFilters.presenteInReport === "" ||
          (row.presenteInReport ? "SI" : "NO") ===
            assignmentFilters.presenteInReport) &&
        (assignmentFilters.esito === "" || row.esito === assignmentFilters.esito) &&
        anomalie.toLowerCase().includes(assignmentFilters.anomalie.toLowerCase())
      );
    });
  }, [assignmentChecks, assignmentFilters]);

  const totale = checks.length;
  const ok = checks.filter((r) => r.esito === "OK").length;
  const errori = checks.filter((r) => r.esito === "ERRORE").length;
  const completezza = totale > 0 ? Math.round((ok / totale) * 100) : 0;

  const assegnatiTotale = assignmentChecks.length;
  const assegnatiOk = assignmentChecks.filter((r) => r.esito === "OK").length;
  const assegnatiErrore = assignmentChecks.filter((r) => r.esito === "ERRORE").length;
  const assegnatiNonVerificati = assignmentChecks.filter((r) => !r.presenteInTodo).length;

  const completezzaColor =
    completezza === 100 ? "#16a34a" : completezza >= 51 ? "#f59e0b" : "#dc2626";

  function esportaExcel() {
    const rows = filteredChecks.map((row) => ({
      "N.": `${row.progressivo}${row.label ? ` (${row.label})` : ""}`,
      "Codice elaborato Report": row.codiceReport || "",
      "Codice elaborato nel Title Trimble": row.codiceTitleTrimble || "",
      "Esito codice": row.titleOk ? "OK" : "ERRORE",
      Tags: row.tags || "",
      "Esito Tags": row.tagsOk ? "OK" : "ERRORE",
      Disciplina: row.disciplina || "",
      "Esito Disciplina": row.disciplinaOk ? "OK" : "ERRORE",
      Status: row.status || "",
      "Esito Status": row.statusOk ? "OK" : "ERRORE",
      "Ispettore ToDo": row.ispettoreTodo || "",
      "Ispettori assegnati PM": row.ispettoriAssegnati || "",
      "Scheda assegnata PM": row.schedaAssegnata || "",
      "Verifica assegnazione": row.verificaAssegnazione,
      Esito: row.esito,
      Anomalie: row.anomalie.join(" | "),
    }));

    const assignmentRows = filteredAssignmentChecks.map((row) => ({
      "Codice elaborato assegnato": row.codice,
      Titolo: row.titolo,
      Ispettori: row.ispettori,
      Scheda: row.scheda,
      Disciplina: row.disciplina,
      "Presente nel ToDo": row.presenteInTodo ? "SI" : "NO",
      "Presente nel Report_Completo": row.presenteInReport ? "SI" : "NO",
      Esito: row.esito,
      Anomalie: row.anomalie.join(" | "),
    }));

    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(rows),
      "Controllo ToDo"
    );

    if (assignmentRows.length > 0) {
      XLSX.utils.book_append_sheet(
        workbook,
        XLSX.utils.json_to_sheet(assignmentRows),
        "Controllo Assegnazioni PM"
      );
    }

    XLSX.writeFile(workbook, "Report_Controllo_ToDo_Ispettori.xlsx");
  }

  return (
    <main
      style={{
        minHeight: "100vh",
        padding: 30,
        background: "#f1f5f9",
        fontFamily: "Arial, sans-serif",
      }}
    >
      <Link
        href="/"
        style={{
          color: "#0284c7",
          fontWeight: 700,
          textDecoration: "none",
        }}
      >
        ← Torna alla dashboard
      </Link>

      <h1 style={{ fontSize: 36, marginTop: 24 }}>
        Controllo ToDo Ispettori
      </h1>

      <p style={{ color: "#475569", fontSize: 18, maxWidth: 1150 }}>
        Verifica automatica di Title, Tags, Disciplina, Status e assegnazioni PM.
        Il file Ispettori.xlsx permette di controllare se gli elaborati assegnati
        dal PM sono stati effettivamente verificati dagli ispettori e se risultano
        nella scheda corretta.
      </p>

      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(4, 1fr)",
          gap: 16,
          marginTop: 24,
          maxWidth: 1500,
        }}
      >
        <div style={cardStyle}>
          <b>ToDo XLSX</b>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => setTodoFile(e.target.files?.[0] || null)}
            style={inputStyle}
          />
        </div>

        <div style={cardStyle}>
          <b>Report_Completo.xlsx</b>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => setReportFile(e.target.files?.[0] || null)}
            style={inputStyle}
          />
        </div>

        <div style={cardStyle}>
          <b>ELENCO_ELABORATI.xlsx</b>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => setElencoFile(e.target.files?.[0] || null)}
            style={inputStyle}
          />
        </div>

        <div style={cardStyle}>
          <b>Assegnazioni PM / Ispettori.xlsx</b>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => setIspettoriFile(e.target.files?.[0] || null)}
            style={inputStyle}
          />
          <div style={{ marginTop: 8, color: "#64748b", fontSize: 12 }}>
            Opzionale, ma consigliato per verificare gli elaborati assegnati.
          </div>
        </div>
      </div>

      <button
        onClick={eseguiControllo}
        disabled={loading}
        style={{
          marginTop: 20,
          padding: "14px 22px",
          background: loading ? "#334155" : "#0f172a",
          color: "white",
          border: "none",
          borderRadius: 12,
          fontWeight: 700,
          cursor: loading ? "not-allowed" : "pointer",
        }}
      >
        {loading ? "Controllo in corso..." : "Esegui controllo ToDo"}
      </button>

      {totale > 0 && (
        <>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(7, 1fr)",
              gap: 12,
              marginTop: 28,
            }}
          >
            <div style={cardStyle}>
              <div style={kpiLabel}>Completezza rilievi ispettori</div>
              <div
                style={{
                  fontSize: 36,
                  fontWeight: 800,
                  color: completezzaColor,
                }}
              >
                {completezza}%
              </div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>Righe ToDo controllate</div>
              <div style={kpiValue}>{totale}</div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>Righe corrette</div>
              <div style={{ ...kpiValue, color: "#16a34a" }}>{ok}</div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>Righe con errori</div>
              <div style={{ ...kpiValue, color: "#dc2626" }}>{errori}</div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>Elaborati assegnati PM</div>
              <div style={kpiValue}>{assegnatiTotale}</div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>Assegnati verificati</div>
              <div style={{ ...kpiValue, color: "#16a34a" }}>{assegnatiOk}</div>
            </div>

            <div style={cardStyle}>
              <div style={kpiLabel}>Assegnati non verificati</div>
              <div style={{ ...kpiValue, color: "#dc2626" }}>{assegnatiNonVerificati}</div>
            </div>
          </div>

          <div style={{ marginTop: 28, ...cardStyle }}>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                gap: 16,
                marginBottom: 16,
              }}
            >
              <div>
                <h2 style={{ margin: 0 }}>Report controllo ToDo</h2>
                <div style={{ marginTop: 6, color: "#64748b", fontSize: 13 }}>
                  Righe visualizzate: {filteredChecks.length} su {checks.length}
                </div>
              </div>

              <button
                onClick={esportaExcel}
                style={{
                  padding: "10px 16px",
                  background: "#16a34a",
                  color: "white",
                  border: "none",
                  borderRadius: 10,
                  fontWeight: 700,
                  cursor: "pointer",
                }}
              >
                Esporta Excel
              </button>
            </div>

            <div style={{ overflowX: "auto" }}>
              <table
                style={{
                  width: "100%",
                  borderCollapse: "collapse",
                  fontSize: 13,
                }}
              >
                <thead>
                  <tr style={{ background: "#f8fafc" }}>
                    <th style={th}>N.</th>
                    <th style={th}>Codice elaborato Report</th>
                    <th style={th}>Codice elaborato nel Title Trimble</th>
                    <th style={th}>Esito codice</th>
                    <th style={th}>Tags</th>
                    <th style={th}>Disciplina</th>
                    <th style={th}>Status</th>
                    <th style={th}>Ispettore ToDo</th>
                    <th style={th}>Ispettori PM</th>
                    <th style={th}>Scheda PM</th>
                    <th style={th}>Assegnazione</th>
                    <th style={th}>Esito</th>
                    <th style={th}>Anomalie</th>
                  </tr>

                  <tr style={{ background: "white" }}>
                    <th style={th}>
                      <input
                        value={filters.n}
                        onChange={(e) => updateFilter("n", e.target.value)}
                        placeholder="Filtra N."
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.codiceReport}
                        onChange={(e) =>
                          updateFilter("codiceReport", e.target.value)
                        }
                        placeholder="Filtra codice"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.codiceTitleTrimble}
                        onChange={(e) =>
                          updateFilter("codiceTitleTrimble", e.target.value)
                        }
                        placeholder="Filtra title"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <select
                        value={filters.esitoCodice}
                        onChange={(e) =>
                          updateFilter("esitoCodice", e.target.value)
                        }
                        style={filterInput}
                      >
                        <option value="">Tutti</option>
                        <option value="OK">OK</option>
                        <option value="ERRORE">ERRORE</option>
                      </select>
                    </th>

                    <th style={th}>
                      <input
                        value={filters.tags}
                        onChange={(e) => updateFilter("tags", e.target.value)}
                        placeholder="Filtra tags"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.disciplina}
                        onChange={(e) =>
                          updateFilter("disciplina", e.target.value)
                        }
                        placeholder="Filtra disciplina"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.status}
                        onChange={(e) => updateFilter("status", e.target.value)}
                        placeholder="Filtra status"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.ispettoriAssegnati}
                        onChange={(e) =>
                          updateFilter("ispettoriAssegnati", e.target.value)
                        }
                        placeholder="Filtra ispettori"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.ispettoriAssegnati}
                        onChange={(e) =>
                          updateFilter("ispettoriAssegnati", e.target.value)
                        }
                        placeholder="Filtra PM"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <input
                        value={filters.schedaAssegnata}
                        onChange={(e) =>
                          updateFilter("schedaAssegnata", e.target.value)
                        }
                        placeholder="Filtra scheda"
                        style={filterInput}
                      />
                    </th>

                    <th style={th}>
                      <select
                        value={filters.verificaAssegnazione}
                        onChange={(e) =>
                          updateFilter("verificaAssegnazione", e.target.value)
                        }
                        style={filterInput}
                      >
                        <option value="">Tutte</option>
                        <option value="OK">OK</option>
                        <option value="NON ASSEGNATO">NON ASSEGNATO</option>
                        <option value="ISPETTORE NON COERENTE">
                          ISPETTORE NON COERENTE
                        </option>
                        <option value="SCHEDA ERRATA">SCHEDA ERRATA</option>
                      </select>
                    </th>

                    <th style={th}>
                      <select
                        value={filters.esito}
                        onChange={(e) => updateFilter("esito", e.target.value)}
                        style={filterInput}
                      >
                        <option value="">Tutti</option>
                        <option value="OK">OK</option>
                        <option value="ERRORE">ERRORE</option>
                      </select>
                    </th>

                    <th style={th}>
                      <input
                        value={filters.anomalie}
                        onChange={(e) =>
                          updateFilter("anomalie", e.target.value)
                        }
                        placeholder="Filtra anomalie"
                        style={filterInput}
                      />
                    </th>
                  </tr>
                </thead>

                <tbody>
                  {filteredChecks.map((row) => (
                    <tr key={row.rowNumber}>
                      <td style={td}>
                        {row.progressivo}
                        {row.label ? ` (${row.label})` : ""}
                      </td>

                      <td style={td}>{row.codiceReport || ""}</td>

                      <td style={td}>{row.codiceTitleTrimble || "-"}</td>

                      <td style={td}>
                        <b
                          style={{
                            color: row.titleOk ? "#16a34a" : "#dc2626",
                          }}
                        >
                          {esitoIcon(row.titleOk)}
                        </b>
                      </td>

                      <td style={td}>
                        {row.tags || "-"}{" "}
                        <b
                          style={{
                            color: row.tagsOk ? "#16a34a" : "#dc2626",
                          }}
                        >
                          {esitoIcon(row.tagsOk)}
                        </b>
                      </td>

                      <td style={td}>
                        {row.disciplina || "-"}{" "}
                        <b
                          style={{
                            color: row.disciplinaOk ? "#16a34a" : "#dc2626",
                          }}
                        >
                          {esitoIcon(row.disciplinaOk)}
                        </b>
                      </td>

                      <td style={td}>
                        {row.status || "-"}{" "}
                        <b
                          style={{
                            color: row.statusOk ? "#16a34a" : "#dc2626",
                          }}
                        >
                          {esitoIcon(row.statusOk)}
                        </b>
                      </td>

                      <td style={td}>{row.ispettoreTodo || "-"}</td>
                      <td style={td}>{row.ispettoriAssegnati || "-"}</td>
                      <td style={td}>{row.schedaAssegnata || "-"}</td>

                      <td
                        style={{
                          ...td,
                          fontWeight: 700,
                          color:
                            row.verificaAssegnazione === "OK"
                              ? "#16a34a"
                              : "#dc2626",
                        }}
                      >
                        {row.verificaAssegnazione}
                      </td>

                      <td
                        style={{
                          ...td,
                          fontWeight: 700,
                          color: row.esito === "OK" ? "#16a34a" : "#dc2626",
                        }}
                      >
                        {row.esito}
                      </td>

                      <td style={td}>{row.anomalie.join(" | ")}</td>
                    </tr>
                  ))}

                  {filteredChecks.length === 0 && (
                    <tr>
                      <td style={td} colSpan={13}>
                        Nessuna riga trovata con i filtri impostati.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {assignmentChecks.length > 0 && (
            <div style={{ marginTop: 28, ...cardStyle }}>
              <h2 style={{ marginTop: 0 }}>Controllo elaborati assegnati dal PM</h2>

              <div style={{ marginBottom: 14, color: "#64748b", fontSize: 13 }}>
                Verifica che ogni elaborato assegnato nel file Ispettori.xlsx sia
                presente nel ToDo e nel Report_Completo.
              </div>

              <div style={{ overflowX: "auto" }}>
                <table
                  style={{
                    width: "100%",
                    borderCollapse: "collapse",
                    fontSize: 13,
                  }}
                >
                  <thead>
                    <tr style={{ background: "#f8fafc" }}>
                      <th style={th}>Codice elaborato assegnato</th>
                      <th style={th}>Titolo</th>
                      <th style={th}>Ispettori</th>
                      <th style={th}>Scheda</th>
                      <th style={th}>Disciplina</th>
                      <th style={th}>Presente ToDo</th>
                      <th style={th}>Presente Report</th>
                      <th style={th}>Esito</th>
                      <th style={th}>Anomalie</th>
                    </tr>

                    <tr style={{ background: "white" }}>
                      <th style={th}>
                        <input
                          value={assignmentFilters.codice}
                          onChange={(e) =>
                            updateAssignmentFilter("codice", e.target.value)
                          }
                          placeholder="Filtra codice"
                          style={filterInput}
                        />
                      </th>
                      <th style={th}>
                        <input
                          value={assignmentFilters.titolo}
                          onChange={(e) =>
                            updateAssignmentFilter("titolo", e.target.value)
                          }
                          placeholder="Filtra titolo"
                          style={filterInput}
                        />
                      </th>
                      <th style={th}>
                        <input
                          value={assignmentFilters.ispettori}
                          onChange={(e) =>
                            updateAssignmentFilter("ispettori", e.target.value)
                          }
                          placeholder="Filtra ispettori"
                          style={filterInput}
                        />
                      </th>
                      <th style={th}>
                        <input
                          value={assignmentFilters.scheda}
                          onChange={(e) =>
                            updateAssignmentFilter("scheda", e.target.value)
                          }
                          placeholder="Filtra scheda"
                          style={filterInput}
                        />
                      </th>
                      <th style={th}>
                        <input
                          value={assignmentFilters.disciplina}
                          onChange={(e) =>
                            updateAssignmentFilter("disciplina", e.target.value)
                          }
                          placeholder="Filtra disciplina"
                          style={filterInput}
                        />
                      </th>
                      <th style={th}>
                        <select
                          value={assignmentFilters.presenteInTodo}
                          onChange={(e) =>
                            updateAssignmentFilter("presenteInTodo", e.target.value)
                          }
                          style={filterInput}
                        >
                          <option value="">Tutti</option>
                          <option value="SI">SI</option>
                          <option value="NO">NO</option>
                        </select>
                      </th>
                      <th style={th}>
                        <select
                          value={assignmentFilters.presenteInReport}
                          onChange={(e) =>
                            updateAssignmentFilter("presenteInReport", e.target.value)
                          }
                          style={filterInput}
                        >
                          <option value="">Tutti</option>
                          <option value="SI">SI</option>
                          <option value="NO">NO</option>
                        </select>
                      </th>
                      <th style={th}>
                        <select
                          value={assignmentFilters.esito}
                          onChange={(e) =>
                            updateAssignmentFilter("esito", e.target.value)
                          }
                          style={filterInput}
                        >
                          <option value="">Tutti</option>
                          <option value="OK">OK</option>
                          <option value="ERRORE">ERRORE</option>
                        </select>
                      </th>
                      <th style={th}>
                        <input
                          value={assignmentFilters.anomalie}
                          onChange={(e) =>
                            updateAssignmentFilter("anomalie", e.target.value)
                          }
                          placeholder="Filtra anomalie"
                          style={filterInput}
                        />
                      </th>
                    </tr>
                  </thead>

                  <tbody>
                    {filteredAssignmentChecks.map((row) => (
                      <tr key={row.codiceNorm}>
                        <td style={td}>{row.codice}</td>
                        <td style={td}>{row.titolo}</td>
                        <td style={td}>{row.ispettori || "-"}</td>
                        <td style={td}>{row.scheda || "-"}</td>
                        <td style={td}>{row.disciplina || "-"}</td>
                        <td style={td}>{row.presenteInTodo ? "✅ SI" : "❌ NO"}</td>
                        <td style={td}>
                          {row.presenteInReport ? "✅ SI" : "❌ NO"}
                        </td>
                        <td
                          style={{
                            ...td,
                            fontWeight: 700,
                            color: row.esito === "OK" ? "#16a34a" : "#dc2626",
                          }}
                        >
                          {row.esito}
                        </td>
                        <td style={td}>{row.anomalie.join(" | ")}</td>
                      </tr>
                    ))}

                    {filteredAssignmentChecks.length === 0 && (
                      <tr>
                        <td style={td} colSpan={9}>
                          Nessuna assegnazione trovata con i filtri impostati.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </>
      )}
    </main>
  );
}

const cardStyle: React.CSSProperties = {
  background: "white",
  border: "1px solid #e2e8f0",
  borderRadius: 16,
  padding: 16,
  boxShadow: "0 6px 18px rgba(15,23,42,.06)",
};

const inputStyle: React.CSSProperties = {
  display: "block",
  width: "100%",
  marginTop: 10,
};

const filterInput: React.CSSProperties = {
  width: "100%",
  padding: "6px 8px",
  border: "1px solid #cbd5e1",
  borderRadius: 6,
  fontSize: 12,
  boxSizing: "border-box",
};

const kpiLabel: React.CSSProperties = {
  fontSize: 13,
  color: "#64748b",
};

const kpiValue: React.CSSProperties = {
  fontSize: 36,
  fontWeight: 800,
  color: "#0f172a",
};

const th: React.CSSProperties = {
  border: "1px solid #e2e8f0",
  padding: 8,
  textAlign: "left",
};

const td: React.CSSProperties = {
  border: "1px solid #e2e8f0",
  padding: 8,
  verticalAlign: "top",
};
