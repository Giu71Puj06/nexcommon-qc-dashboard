// Assegnazione dei topic BCF standalone ai ToDo Trimble quando >=2 ToDo condividono
// la stessa chiave di merge. Un topic viene assegnato al ToDo con il maggior numero di
// commenti in comune; se non ha commenti in comune con NESSUNO, non viene scartato ma
// restituito in `unassigned`, cosi' il chiamante puo' emetterlo come riga autonoma
// (FIX M3/B1: preserva i commenti PRG/ISP, unica fonte dal BCF Trimble).
export function assignTopicsToTrimble<T>(
  trimbleRows: T[],
  topicRows: T[],
  fingerprints: (row: T) => Set<string>
): { byTrimbleIndex: Map<number, T[]>; unassigned: T[] } {
  const byTrimbleIndex = new Map<number, T[]>();
  const unassigned: T[] = [];
  const trimbleFps = trimbleRows.map(fingerprints);

  for (const topicRow of topicRows) {
    const topicKeys = fingerprints(topicRow);
    let bestIndex = -1;
    let bestOverlap = 0;

    trimbleFps.forEach((tk, index) => {
      let overlap = 0;
      topicKeys.forEach((key) => {
        if (tk.has(key)) overlap += 1;
      });
      if (overlap > bestOverlap) {
        bestOverlap = overlap;
        bestIndex = index;
      }
    });

    if (bestIndex >= 0 && bestOverlap > 0) {
      if (!byTrimbleIndex.has(bestIndex)) byTrimbleIndex.set(bestIndex, []);
      byTrimbleIndex.get(bestIndex)!.push(topicRow);
    } else {
      unassigned.push(topicRow);
    }
  }

  return { byTrimbleIndex, unassigned };
}
