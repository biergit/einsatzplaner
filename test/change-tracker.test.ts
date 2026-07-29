import { describe, it, expect } from 'vitest';

interface SheetRowSnapshot {
  row: number;
  key: string;
  display: string;
}

interface ChangeEntry {
  timestamp: number;
  sheetName: string;
  rangeA1: string;
  alterWert: string;
  neuerWert: string;
  bearbeiter: string;
}

function computeSheetDiff(
  snapshot: SheetRowSnapshot[],
  current: SheetRowSnapshot[],
  sheetName: string,
  bearbeiter: string,
  out: ChangeEntry[]
): void {
  const now = Date.now();
  const snapByKey = new Map<string, SheetRowSnapshot[]>();
  const snapByRow = new Map<number, SheetRowSnapshot>();
  for (const r of snapshot) {
    if (!snapByKey.has(r.key)) snapByKey.set(r.key, []);
    snapByKey.get(r.key)!.push(r);
    snapByRow.set(r.row, r);
  }

  const currByKey = new Map<string, SheetRowSnapshot[]>();
  const currByRow = new Map<number, SheetRowSnapshot>();
  for (const r of current) {
    if (!currByKey.has(r.key)) currByKey.set(r.key, []);
    currByKey.get(r.key)!.push(r);
    currByRow.set(r.row, r);
  }

  const allKeys = new Set([...snapByKey.keys(), ...currByKey.keys()]);

  const unmatchedSnap: SheetRowSnapshot[] = [];
  const unmatchedCurr: SheetRowSnapshot[] = [];

  for (const key of allKeys) {
    const snapRows = snapByKey.get(key) || [];
    const currRows = currByKey.get(key) || [];

    while (snapRows.length > 0 && currRows.length > 0) {
      snapRows.shift();
      currRows.shift();
    }

    for (const s of snapRows) unmatchedSnap.push(s);
    for (const c of currRows) unmatchedCurr.push(c);
  }

  for (const s of [...unmatchedSnap]) {
    const sc = currByRow.get(s.row);
    if (sc && unmatchedCurr.includes(sc)) {
      unmatchedSnap.splice(unmatchedSnap.indexOf(s), 1);
      unmatchedCurr.splice(unmatchedCurr.indexOf(sc), 1);
      out.push({
        timestamp: now,
        sheetName,
        rangeA1: `A${s.row}`,
        alterWert: s.display,
        neuerWert: sc.display,
        bearbeiter,
      });
    }
  }

  for (const s of unmatchedSnap) {
    out.push({
      timestamp: now,
      sheetName,
      rangeA1: `A${s.row}`,
      alterWert: s.display,
      neuerWert: '(gelöscht)',
      bearbeiter,
    });
  }

  for (const c of unmatchedCurr) {
    out.push({
      timestamp: now,
      sheetName,
      rangeA1: `A${c.row}`,
      alterWert: '(neu)',
      neuerWert: c.display,
      bearbeiter,
    });
  }
}

function sn(row: number, key: string, display?: string): SheetRowSnapshot {
  return { row, key, display: display ?? `Row ${row} — ${key}` };
}

function diff(snapshot: SheetRowSnapshot[], current: SheetRowSnapshot[]): ChangeEntry[] {
  const out: ChangeEntry[] = [];
  computeSheetDiff(snapshot, current, 'Test', 'bearbeiter@test', out);
  return out;
}

describe('computeSheetDiff', () => {
  it('empty snapshot and empty current → no entries', () => {
    expect(diff([], [])).toHaveLength(0);
  });

  it('no changes (snapshot == current) → no entries', () => {
    const rows = [sn(2, 'a'), sn(3, 'b')];
    expect(diff(rows, rows)).toHaveLength(0);
  });

  it('cut+paste (same key, different row) → no entries', () => {
    expect(diff(
      [sn(2, 'Max|Urlaub', 'Max | Urlaub')],
      [sn(5, 'Max|Urlaub', 'Max | Urlaub')],
    )).toHaveLength(0);
  });

  it('cut+paste with multiple rows → matched silently', () => {
    expect(diff(
      [sn(2, 'a'), sn(3, 'b'), sn(4, 'c')],
      [sn(2, 'a'), sn(4, 'b'), sn(5, 'c')],
    )).toHaveLength(0);
  });

  it('delete (in snapshot, not in current) → gelöscht', () => {
    const entries = diff([sn(2, 'a')], []);
    expect(entries).toHaveLength(1);
    expect(entries[0].neuerWert).toBe('(gelöscht)');
    expect(entries[0].alterWert).toContain('Row 2');
    expect(entries[0].sheetName).toBe('Test');
    expect(entries[0].rangeA1).toBe('A2');
  });

  it('new entry (in current, not in snapshot) → neu', () => {
    const entries = diff([], [sn(3, 'b')]);
    expect(entries).toHaveLength(1);
    expect(entries[0].alterWert).toBe('(neu)');
    expect(entries[0].neuerWert).toContain('Row 3');
    expect(entries[0].rangeA1).toBe('A3');
  });

  it('modified (same row, different key) → alt.display → neu.display', () => {
    const entries = diff(
      [sn(2, 'Max|Urlaub', 'Max | Urlaub')],
      [sn(2, 'Max|Krank', 'Max | Krank')],
    );
    expect(entries).toHaveLength(1);
    expect(entries[0].alterWert).toBe('Max | Urlaub');
    expect(entries[0].neuerWert).toBe('Max | Krank');
    expect(entries[0].rangeA1).toBe('A2');
  });

  it('mixed: delete row 2, add row 4, modify row 3', () => {
    const entries = diff(
      [sn(2, 'a'), sn(3, 'b')],
      [sn(3, 'b-mod', 'B modified'), sn(4, 'c')],
    );
    expect(entries).toHaveLength(3);

    const deleted = entries.find(e => e.neuerWert === '(gelöscht)');
    const added = entries.find(e => e.alterWert === '(neu)');
    const modified = entries.find(e => e.alterWert !== '(neu)' && e.neuerWert !== '(gelöscht)');

    expect(deleted).toBeDefined();
    expect(deleted!.rangeA1).toBe('A2');

    expect(added).toBeDefined();
    expect(added!.neuerWert).toContain('c');

    expect(modified).toBeDefined();
    expect(modified!.alterWert).toContain('b');
    expect(modified!.neuerWert).toContain('B modified');
    expect(modified!.rangeA1).toBe('A3');
  });

  it('two identical rows in snapshot + two in current → all matched', () => {
    expect(diff(
      [sn(2, 'dup'), sn(3, 'dup')],
      [sn(4, 'dup'), sn(5, 'dup')],
    )).toHaveLength(0);
  });

  it('three identical in snapshot, two in current → one deleted', () => {
    const entries = diff(
      [sn(2, 'dup'), sn(3, 'dup'), sn(4, 'dup')],
      [sn(5, 'dup'), sn(6, 'dup')],
    );
    expect(entries).toHaveLength(1);
    expect(entries[0].neuerWert).toBe('(gelöscht)');
  });

  it('one in snapshot, three in current → two new', () => {
    const entries = diff(
      [sn(2, 'dup')],
      [sn(3, 'dup'), sn(4, 'dup'), sn(5, 'dup')],
    );
    expect(entries).toHaveLength(2);
    expect(entries.every(e => e.alterWert === '(neu)')).toBe(true);
  });

  it('row-based fallback: different keys at same row → modified', () => {
    const entries = diff(
      [sn(3, 'k1', 'alt')],
      [sn(3, 'k2', 'neu')],
    );
    expect(entries).toHaveLength(1);
    expect(entries[0].alterWert).toBe('alt');
    expect(entries[0].neuerWert).toBe('neu');
    expect(entries[0].rangeA1).toBe('A3');
  });

  it('row-based fallback does not match across different rows', () => {
    const entries = diff(
      [sn(3, 'k1', 'row3')],
      [sn(5, 'k2', 'row5')],
    );
    expect(entries).toHaveLength(2);
    const deleted = entries.find(e => e.neuerWert === '(gelöscht)');
    const added = entries.find(e => e.alterWert === '(neu)');
    expect(deleted!.rangeA1).toBe('A3');
    expect(added!.rangeA1).toBe('A5');
  });

  it('move + modify: key moved to new row, old row gets new content → move ignored, modify logged', () => {
    const entries = diff(
      [sn(2, 'a', 'A'), sn(3, 'b', 'B')],
      [sn(2, 'c', 'C'), sn(4, 'b', 'B')],
    );
    const modified = entries.find(e => e.rangeA1 === 'A2');
    expect(modified).toBeDefined();
    expect(modified!.alterWert).toBe('A');
    expect(modified!.neuerWert).toBe('C');
    expect(entries).toHaveLength(1);
  });

  it('delete all (snapshot has entries, current empty) → all gelöscht', () => {
    const entries = diff(
      [sn(2, 'a', 'A'), sn(3, 'b', 'B'), sn(4, 'c', 'C')],
      [],
    );
    expect(entries).toHaveLength(3);
    expect(entries.every(e => e.neuerWert === '(gelöscht)')).toBe(true);
    expect(entries.every(e => e.sheetName === 'Test')).toBe(true);
  });

  it('add all (snapshot empty, current has entries) → all neu', () => {
    const entries = diff(
      [],
      [sn(2, 'a', 'A'), sn(3, 'b', 'B')],
    );
    expect(entries).toHaveLength(2);
    expect(entries.every(e => e.alterWert === '(neu)')).toBe(true);
  });
});

// Saison-Diff Tests (diffSaisonRow + computeSaisonDiff)
interface SaisonRowSnapshot {
  datum: string; gegner: string; startzeit: string; heimAuswaerts: string;
  playerAssignments: string[]; ersatz: string[]; status: string; kommentar: string;
}

interface ChangedCell { label: string; oldVal: string; newVal: string; }
interface SaisonModifiedRow { datum: string; gegner: string; oldRow: SaisonRowSnapshot; newRow: SaisonRowSnapshot; changedCells: ChangedCell[]; }
interface SaisonDiffResult { modified: SaisonModifiedRow[]; added: SaisonRowSnapshot[]; deleted: SaisonRowSnapshot[]; }

function computeSaisonDiff(snapshot: SaisonRowSnapshot[], current: SaisonRowSnapshot[]): SaisonDiffResult {
  function key(r: SaisonRowSnapshot) { return `${r.datum}|${r.gegner}`; }
  const oldByKey = new Map<string, SaisonRowSnapshot>();
  for (const r of snapshot) oldByKey.set(key(r), r);
  const newByKey = new Map<string, SaisonRowSnapshot>();
  for (const r of current) newByKey.set(key(r), r);
  const allKeys = new Set([...oldByKey.keys(), ...newByKey.keys()]);
  const modified: SaisonModifiedRow[] = [];
  const added: SaisonRowSnapshot[] = [];
  const deleted: SaisonRowSnapshot[] = [];
  for (const k of allKeys) {
    const oldR = oldByKey.get(k);
    const newR = newByKey.get(k);
    if (oldR && newR) {
      const changes = diffSaisonRow(oldR, newR);
      if (changes.length > 0) modified.push({ datum: oldR.datum, gegner: oldR.gegner, oldRow: oldR, newRow: newR, changedCells: changes });
    } else if (!oldR && newR) { added.push(newR); }
    else if (oldR && !newR) { deleted.push(oldR); }
  }
  return { modified, added, deleted };
}

function diffSaisonRow(oldR: SaisonRowSnapshot, newR: SaisonRowSnapshot): ChangedCell[] {
  const changes: ChangedCell[] = [];
  if (oldR.gegner !== newR.gegner) changes.push({ label: 'Gegner', oldVal: oldR.gegner, newVal: newR.gegner });
  if (oldR.startzeit !== newR.startzeit) changes.push({ label: 'Startzeit', oldVal: oldR.startzeit, newVal: newR.startzeit });
  if (oldR.heimAuswaerts !== newR.heimAuswaerts) changes.push({ label: 'Ort', oldVal: oldR.heimAuswaerts, newVal: newR.heimAuswaerts });
  for (let pi = 0; pi < oldR.playerAssignments.length; pi++) {
    const ov = oldR.playerAssignments[pi]; const nv = newR.playerAssignments[pi];
    if (ov !== nv) changes.push({ label: `Spieler${pi + 1}`, oldVal: ov, newVal: nv });
  }
  for (let ei = 0; ei < 3; ei++) {
    if (oldR.ersatz[ei] !== newR.ersatz[ei]) changes.push({ label: `Ersatz ${ei + 1}`, oldVal: oldR.ersatz[ei], newVal: newR.ersatz[ei] });
  }
  if (oldR.status !== newR.status) changes.push({ label: 'Status', oldVal: oldR.status, newVal: newR.status });
  if (oldR.kommentar !== newR.kommentar) changes.push({ label: 'Kommentar', oldVal: oldR.kommentar, newVal: newR.kommentar });
  return changes;
}

function sRow(datum: string, gegner: string, overrides: Partial<SaisonRowSnapshot> = {}): SaisonRowSnapshot {
  return {
    datum, gegner,
    startzeit: '', heimAuswaerts: '',
    playerAssignments: ['Einzel+Doppel', 'Einzel+Doppel', '', '', ''],
    ersatz: ['', '', ''],
    status: 'Final', kommentar: '',
    ...overrides,
  };
}

describe('computeSaisonDiff', () => {
  it('detects Kommentar change', () => {
    const oldSnap = [sRow('01.09.2026', 'ABC', { kommentar: '' })];
    const newSnap = [sRow('01.09.2026', 'ABC', { kommentar: 'Treffpunkt 18:30' })];
    const result = computeSaisonDiff(oldSnap, newSnap);
    expect(result.modified).toHaveLength(1);
    expect(result.modified[0].changedCells).toContainEqual({ label: 'Kommentar', oldVal: '', newVal: 'Treffpunkt 18:30' });
  });

  it('detects no change when Kommentar unchanged', () => {
    const row = sRow('01.09.2026', 'ABC', { kommentar: 'Treffpunkt' });
    const result = computeSaisonDiff([row], [row]);
    expect(result.modified).toHaveLength(0);
  });

  it('detects new Spieltag with Kommentar', () => {
    const result = computeSaisonDiff([], [sRow('01.09.2026', 'ABC', { kommentar: 'Info' })]);
    expect(result.added).toHaveLength(1);
    expect(result.added[0].kommentar).toBe('Info');
  });

  it('detects deleted Spieltag', () => {
    const result = computeSaisonDiff([sRow('01.09.2026', 'ABC')], []);
    expect(result.deleted).toHaveLength(1);
  });

  it('detects Status change', () => {
    const oldRow = sRow('01.09.2026', 'ABC', { status: 'Geplant' });
    const newRow = sRow('01.09.2026', 'ABC', { status: 'Final' });
    const result = computeSaisonDiff([oldRow], [newRow]);
    expect(result.modified[0].changedCells).toContainEqual({ label: 'Status', oldVal: 'Geplant', newVal: 'Final' });
  });

  it('ignores Geplant in added count (no-op, just ensure structure)', () => {
    const result = computeSaisonDiff([], [sRow('01.09.2026', 'ABC', { status: 'Geplant' })]);
    expect(result.added).toHaveLength(1);
  });
});

// ─── Benachrichtigungs-Logik (sendChangeNotification, flushPendingChanges) ──

/**
 * AbwAffectedEntry – wie in ChangeTracker.ts definiert.
 * Wird von rebuildAbwesenheitenInSaison() befüllt und beim Mail-Versand
 * ausgewertet.
 */
interface AbwAffectedEntry {
  datumStr: string;
  startzeit: string;
  heimAuswaerts: string;
  gegner: string;
  spielerName: string;
  warAufgestellt: boolean;
  warFinal: boolean;
}

function ae(overrides: Partial<AbwAffectedEntry> = {}): AbwAffectedEntry {
  return {
    datumStr: '01.09.2026',
    startzeit: '20:00',
    heimAuswaerts: 'Heim',
    gegner: 'Gegner ABC',
    spielerName: 'Max',
    warAufgestellt: false,
    warFinal: false,
    ...overrides,
  };
}

describe('hasAffectedFinals detection', () => {
  it('empty array → false', () => {
    const result = [].some(a => a.warFinal);
    expect(result).toBe(false);
  });

  it('single entry with warFinal=false → false', () => {
    const result = [ae({ warFinal: false })].some(a => a.warFinal);
    expect(result).toBe(false);
  });

  it('single entry with warFinal=true → true', () => {
    const result = [ae({ warFinal: true })].some(a => a.warFinal);
    expect(result).toBe(true);
  });

  it('multiple entries, none warFinal → false', () => {
    const result = [
      ae({ warFinal: false, spielerName: 'A' }),
      ae({ warFinal: false, spielerName: 'B' }),
      ae({ warFinal: false, spielerName: 'C' }),
    ].some(a => a.warFinal);
    expect(result).toBe(false);
  });

  it('multiple entries, one warFinal → true', () => {
    const result = [
      ae({ warFinal: false, spielerName: 'A' }),
      ae({ warFinal: true, spielerName: 'B' }),   // Final→Geplant
      ae({ warFinal: false, spielerName: 'C' }),
    ].some(a => a.warFinal);
    expect(result).toBe(true);
  });

  it('null input → false', () => {
    const abwAffected: AbwAffectedEntry[] | null = null;
    const result = abwAffected ? abwAffected.some(a => a.warFinal) : false;
    expect(result).toBe(false);
  });
});

describe('affectedNames from abwAffected', () => {
  it('player with warAufgestellt=true is added', () => {
    const affected: AbwAffectedEntry[] = [
      ae({ spielerName: 'Max', warAufgestellt: true, warFinal: true }),
    ];
    const names = new Set<string>();
    for (const a of affected) {
      if (a.warAufgestellt) {
        names.add(a.spielerName);
      }
    }
    expect(names.has('Max')).toBe(true);
  });

  it('player with warAufgestellt=false is NOT added', () => {
    const affected: AbwAffectedEntry[] = [
      ae({ spielerName: 'ErsatzSpieler', warAufgestellt: false }),
    ];
    const names = new Set<string>();
    for (const a of affected) {
      if (a.warAufgestellt) {
        names.add(a.spielerName);
      }
    }
    expect(names.size).toBe(0);
  });

  it('multiple players: only warAufgestellt are added', () => {
    const affected: AbwAffectedEntry[] = [
      ae({ spielerName: 'Max', warAufgestellt: true }),
      ae({ spielerName: 'Anna', warAufgestellt: false }),
      ae({ spielerName: 'Tom', warAufgestellt: true }),
    ];
    const names = new Set<string>();
    for (const a of affected) {
      if (a.warAufgestellt) {
        names.add(a.spielerName);
      }
    }
    expect(names.has('Max')).toBe(true);
    expect(names.has('Tom')).toBe(true);
    expect(names.has('Anna')).toBe(false);
    expect(names.size).toBe(2);
  });

  it('abwAffected is null → no names added', () => {
    const affected: AbwAffectedEntry[] | null = null;
    const names = new Set<string>();
    if (affected) {
      for (const a of affected) {
        if (a.warAufgestellt) {
          names.add(a.spielerName);
        }
      }
    }
    expect(names.size).toBe(0);
  });
});

describe('sendChangeNotification guard logic', () => {
  /**
   * Entscheidungs-Logik aus sendChangeNotification (vereinfacht, ohne GAS-Abhängigkeiten):
   * Die Mail wird unterdrückt, wenn ALLE drei Bedingungen zutreffen:
   *   1. visibleSaisonCount === 0  (keine sichtbaren Saison-Änderungen)
   *   2. abwEntries.length === 0   (keine Abwesenheits-Diff-Einträge)
   *   3. !hasAffectedFinals        (kein Final→Geplant-Revert)
   */
  function shouldSuppressEmail(
    visibleSaisonCount: number,
    abwEntriesCount: number,
    abwAffected: AbwAffectedEntry[] | null,
  ): boolean {
    const hasAffectedFinals = abwAffected ? abwAffected.some(a => a.warFinal) : false;
    return visibleSaisonCount === 0 && abwEntriesCount === 0 && !hasAffectedFinals;
  }

  it('no saison changes, no abw entries, no affected finals → suppressed', () => {
    expect(shouldSuppressEmail(0, 0, null)).toBe(true);
  });

  it('no saison changes, no abw entries, no final→Geplant → suppressed', () => {
    expect(shouldSuppressEmail(0, 0, [ae({ warFinal: false })])).toBe(true);
  });

  it('no saison changes, no abw entries, BUT final→Geplant revert exists → NOT suppressed', () => {
    expect(shouldSuppressEmail(0, 0, [ae({ warFinal: true })])).toBe(false);
  });

  it('no saison changes, BUT abw entries exist → NOT suppressed', () => {
    expect(shouldSuppressEmail(0, 1, null)).toBe(false);
  });

  it('saison changes exist, no abw entries → NOT suppressed', () => {
    expect(shouldSuppressEmail(1, 0, null)).toBe(false);
  });

  it('saison changes + abw entries + final→Geplant → NOT suppressed', () => {
    expect(shouldSuppressEmail(1, 2, [ae({ warFinal: true })])).toBe(false);
  });

  it('only abw entries → NOT suppressed', () => {
    expect(shouldSuppressEmail(0, 3, [ae({ warFinal: false })])).toBe(false);
  });

  it('multiple affected finals → NOT suppressed', () => {
    expect(shouldSuppressEmail(0, 0, [
      ae({ warFinal: true, spielerName: 'A' }),
      ae({ warFinal: true, spielerName: 'B' }),
    ])).toBe(false);
  });
});

describe('Abw section display logic', () => {
  /**
   * Die Abwesenheits-Sektion im HTML wird angezeigt wenn:
   *   showAbw && (abwEntriesCount > 0 || hasFinals)
   * (showAbw ist true wenn abwAffected nicht-null und nicht-leer ist)
   */
  function shouldShowAbwSection(
    abwAffected: AbwAffectedEntry[] | null,
    abwEntriesCount: number,
  ): boolean {
    const showAbw = abwAffected && abwAffected.length > 0;
    if (!showAbw) return false;
    const hasFinals = abwAffected!.some(a => a.warFinal);
    return abwEntriesCount > 0 || hasFinals;
  }

  it('abwAffected has entries with warFinal → shown even with 0 entries', () => {
    expect(shouldShowAbwSection([ae({ warFinal: true })], 0)).toBe(true);
  });

  it('abwAffected has entries without warFinal, 0 entries → NOT shown', () => {
    expect(shouldShowAbwSection([ae({ warFinal: false })], 0)).toBe(false);
  });

  it('abwAffected has entries without warFinal, entries=1 → shown', () => {
    expect(shouldShowAbwSection([ae({ warFinal: false })], 1)).toBe(true);
  });

  it('abwAffected null → NOT shown', () => {
    expect(shouldShowAbwSection(null, 5)).toBe(false);
  });

  it('abwAffected empty → NOT shown', () => {
    expect(shouldShowAbwSection([], 5)).toBe(false);
  });

  it('abwAffected with warFinal + entries → shown', () => {
    expect(shouldShowAbwSection([ae({ warFinal: true })], 3)).toBe(true);
  });
});

// ─── onEdit-Fallback-Logik (nur bei DEBOUNCE_FAILED) ───────────────────────

describe('onEdit fallback processing logic', () => {
  /**
   * Im onEdit-Handler wird NUR dann sofort verarbeitet, wenn
   * resetDebounceTimer keinen Trigger erstellen konnte (Quota, Berechtigungen).
   * Normale PENDING_EDITS während der laufenden Debounce-Periode sind
   * kein Grund für sofortige Verarbeitung — der Debounce existiert, damit
   * der Nutzer die Aufstellung vor der Mail korrigieren kann.
   */
  function shouldProcessImmediately(debounceFailed: boolean): boolean {
    return debounceFailed;
  }

  it('debounce OK → NO immediate processing (waits for timer)', () => {
    expect(shouldProcessImmediately(false)).toBe(false);
  });

  it('debounce FAILED → process immediately (no timer to wait for)', () => {
    expect(shouldProcessImmediately(true)).toBe(true);
  });
});

describe('rebuildAbwesenheitenInSaison: ABW_AFFECTED warFinal field', () => {
  /**
   * Das warFinal-Feld wird nur dann true, wenn BOTH:
   *   wasFinalBefore (Status vorher "Final") AND warAufgestellt (Spieler war zugewiesen)
   * wahr sind.
   */
  function computeWarFinal(wasFinalBefore: boolean, warAufgestellt: boolean): boolean {
    return wasFinalBefore && warAufgestellt;
  }

  it('Final + aufgestellt → warFinal ~ true', () => {
    expect(computeWarFinal(true, true)).toBe(true);
  });

  it('Geplant + aufgestellt → warFinal ~ false', () => {
    expect(computeWarFinal(false, true)).toBe(false);
  });

  it('Final + NICHT aufgestellt → warFinal ~ false', () => {
    expect(computeWarFinal(true, false)).toBe(false);
  });

  it('Geplant + NICHT aufgestellt → warFinal ~ false', () => {
    expect(computeWarFinal(false, false)).toBe(false);
  });
});

// ─── Empfänger-Inclusion-Logik (getAenderungenMeldenEmpfaenger) ─────────────

describe('Empfänger inclusion logic', () => {
  /**
   * Entscheidungs-Logik pro Spieler:
   *
   *   Editor (email === currentUser): IMMER excluded
   *   isKapitaen:     included (wenn nicht Editor)
   *   isAffected:     included (wenn nicht Editor)
   *   melden:         included (wenn nicht Editor)
   */
  function shouldInclude(
    isKapitaen: boolean,
    isAffected: boolean,
    melden: boolean,
    isEditor: boolean,
  ): boolean {
    if (isEditor) return false;
    if (isKapitaen) return true;
    if (isAffected) return true;
    if (melden) return true;
    return false;
  }

  it('Kapitän, nicht Editor → included', () => {
    expect(shouldInclude(true, false, false, false)).toBe(true);
  });

  it('Kapitän = Editor → excluded (Selbst-Notifikation)', () => {
    expect(shouldInclude(true, false, false, true)).toBe(false);
  });

  it('betroffener Spieler, nicht Editor → included', () => {
    expect(shouldInclude(false, true, false, false)).toBe(true);
  });

  it('betroffener Spieler = Editor → excluded', () => {
    expect(shouldInclude(false, true, false, true)).toBe(false);
  });

  it('melden-Checkbox, nicht Editor → included', () => {
    expect(shouldInclude(false, false, true, false)).toBe(true);
  });

  it('melden-Checkbox = Editor → excluded', () => {
    expect(shouldInclude(false, false, true, true)).toBe(false);
  });

  it('kein Grund, nicht Editor → excluded', () => {
    expect(shouldInclude(false, false, false, false)).toBe(false);
  });

  it('kein Grund, Editor → excluded', () => {
    expect(shouldInclude(false, false, false, true)).toBe(false);
  });
});

describe('melden checkbox recognition', () => {
  /**
   * Robuster Check: true (boolean), 'TRUE', 'true', 1
   */
  function isMelden(raw: unknown): boolean {
    return raw === true
      || String(raw).toUpperCase() === 'TRUE'
      || raw === 1;
  }

  it('boolean true → true', () => {
    expect(isMelden(true)).toBe(true);
  });

  it('boolean false → false', () => {
    expect(isMelden(false)).toBe(false);
  });

  it('string "TRUE" → true', () => {
    expect(isMelden('TRUE')).toBe(true);
  });

  it('string "true" → true', () => {
    expect(isMelden('true')).toBe(true);
  });

  it('string "false" → false', () => {
    expect(isMelden('false')).toBe(false);
  });

  it('number 1 → true', () => {
    expect(isMelden(1)).toBe(true);
  });

  it('number 0 → false', () => {
    expect(isMelden(0)).toBe(false);
  });

  it('undefined → false', () => {
    expect(isMelden(undefined)).toBe(false);
  });
});

// ─── Saison-Zeilenformat für Änderungslog ───────────────────────────────────

describe('formatSaisonRowForLog', () => {
  /** Nachbildung der formatSaisonRowForLog-Helperfunktion aus ChangeTracker.ts */
  function formatSaisonRowForLog(row: SaisonRowSnapshot, playerNames: string[]): string {
    const parts: string[] = [];
    if (row.gegner) parts.push(row.gegner);
    if (row.startzeit) parts.push(row.startzeit);
    if (row.heimAuswaerts) parts.push(row.heimAuswaerts);
    for (let pi = 0; pi < row.playerAssignments.length; pi++) {
      const v = row.playerAssignments[pi];
      if (v) parts.push(`${playerNames[pi]}: ${v}`);
    }
    for (let ei = 0; ei < row.ersatz.length; ei++) {
      if (row.ersatz[ei]) parts.push(`Ersatz ${ei + 1}: ${row.ersatz[ei]}`);
    }
    parts.push(`Status: ${row.status}`);
    if (row.kommentar) parts.push(`Kommentar: ${row.kommentar}`);
    return parts.join(' | ');
  }

  const names = ['Max', 'Anna', 'Tom'];

  it('full row with all fields', () => {
    const row: SaisonRowSnapshot = {
      datum: '01.09.2026', gegner: 'ABC', startzeit: '20:00', heimAuswaerts: 'Heim',
      playerAssignments: ['Einzel+Doppel', 'Einzel+Doppel', ''],
      ersatz: ['Gast1', '', ''],
      status: 'Final', kommentar: 'Treffpunkt 19:30',
    };
    const result = formatSaisonRowForLog(row, names);
    expect(result).toBe('ABC | 20:00 | Heim | Max: Einzel+Doppel | Anna: Einzel+Doppel | Ersatz 1: Gast1 | Status: Final | Kommentar: Treffpunkt 19:30');
  });

  it('minimal row (only gegner and status)', () => {
    const row: SaisonRowSnapshot = {
      datum: '', gegner: 'DEF', startzeit: '', heimAuswaerts: '',
      playerAssignments: ['', '', ''],
      ersatz: ['', '', ''],
      status: 'Geplant', kommentar: '',
    };
    expect(formatSaisonRowForLog(row, names)).toBe('DEF | Status: Geplant');
  });

  it('row with ✗ markers (absent players)', () => {
    const row: SaisonRowSnapshot = {
      datum: '', gegner: 'GHI', startzeit: '', heimAuswaerts: 'Auswärts',
      playerAssignments: ['Einzel+Doppel', '✗ Urlaub', ''],
      ersatz: ['', '', ''],
      status: 'Geplant', kommentar: '',
    };
    const result = formatSaisonRowForLog(row, names);
    expect(result).toContain('Max: Einzel+Doppel');
    expect(result).toContain('Anna: ✗ Urlaub');
  });
});

// ─── buildAbwesenheitenIndex: überlappende Abwesenheiten ────────────────────

describe('buildAbwesenheitenIndex merge logic', () => {
  /**
   * Merge-Logik: mehrere Abwesenheiten für denselben Spieler am selben Tag
   * werden zu einem String zusammengeführt: ✗ Komm1, Komm2
   */
  function mergeAbsences(abwesenheiten: { name: string; kommentar: string }[]): string {
    const labels = abwesenheiten.map(a => a.kommentar || 'abwesend');
    return `✗ ${labels.join(', ')}`;
  }

  it('single absence → ✗ Urlaub', () => {
    expect(mergeAbsences([{ name: 'Max', kommentar: 'Urlaub' }])).toBe('✗ Urlaub');
  });

  it('no comment → ✗ abwesend', () => {
    expect(mergeAbsences([{ name: 'Max', kommentar: '' }])).toBe('✗ abwesend');
  });

  it('two overlapping absences → ✗ Urlaub, Verletzung', () => {
    expect(mergeAbsences([
      { name: 'Max', kommentar: 'Urlaub' },
      { name: 'Max', kommentar: 'Verletzung' },
    ])).toBe('✗ Urlaub, Verletzung');
  });

  it('three overlapping → ✗ Urlaub, Verletzung, Krank', () => {
    expect(mergeAbsences([
      { name: 'Max', kommentar: 'Urlaub' },
      { name: 'Max', kommentar: 'Verletzung' },
      { name: 'Max', kommentar: 'Krank' },
    ])).toBe('✗ Urlaub, Verletzung, Krank');
  });
});

// ─── Status-Meldung in der Mail (AbwAffectedEntry) ──────────────────────────

describe('AbwAffectedEntry status message', () => {
  /**
   * Die Status-Meldung in der Mail hängt von warAufgestellt, warFinal
   * und der validierung ab.
   */
  function statusText(warAufgestellt: boolean, warFinal: boolean, validierung: string): string {
    if (warAufgestellt) {
      if (warFinal) {
        if (validierung) {
          return `War aufgestellt (Final → Geplant). Validierung: ${validierung}`;
        }
        return 'War aufgestellt (Final → Geplant) — Bitte Spieltag prüfen und Status auf Final setzen.';
      }
      return 'War aufgestellt.';
    }
    return 'Stand als Ersatz zur Verfügung';
  }

  it('warFinal=true + Validierung → zeigt Validierungsmeldungen', () => {
    expect(statusText(true, true, 'Nur 3/4 Einzel | Nur 2/4 Doppel'))
      .toBe('War aufgestellt (Final → Geplant). Validierung: Nur 3/4 Einzel | Nur 2/4 Doppel');
  });

  it('warFinal=true, keine Validierung → Aufforderung zum Prüfen', () => {
    expect(statusText(true, true, ''))
      .toBe('War aufgestellt (Final → Geplant) — Bitte Spieltag prüfen und Status auf Final setzen.');
  });

  it('warAufgestellt=true, warFinal=false (korrigiert) → kein Alarm', () => {
    expect(statusText(true, false, ''))
      .toBe('War aufgestellt.');
  });

  it('nicht aufgestellt → Ersatz-Hinweis', () => {
    expect(statusText(false, false, ''))
      .toBe('Stand als Ersatz zur Verfügung');
  });
});
