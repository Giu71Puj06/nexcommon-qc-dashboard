// Regola CANONICA di classificazione del tipo rilievo (NC / OSS / Nessun rilievo).
// "Da NC a OSS" e' un rilievo declassato: conta come OSS (coerente con copertura, tabelle per-disciplina e PDF).
// Case-insensitive e tollerante a spazi (fix F3).

export function normTipo(tipo: unknown): string {
  return String(tipo ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

export function isNcTipo(tipo: unknown): boolean {
  return normTipo(tipo) === "NC";
}

export function isOssTipo(tipo: unknown): boolean {
  const t = normTipo(tipo);
  return t === "OSS" || t === "DA NC A OSS";
}

export function isNessunRilievoTipo(tipo: unknown): boolean {
  return normTipo(tipo) === "NESSUN RILIEVO";
}

export function isRilievo(tipo: unknown): boolean {
  return isNcTipo(tipo) || isOssTipo(tipo);
}
