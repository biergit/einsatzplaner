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
