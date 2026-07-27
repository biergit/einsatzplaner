import { describe, it, expect } from 'vitest';

function validateAufstellungLogic(
  einsatzarten: string[],
  ersatzCount: number,
  einzelReq: number,
  doppelReq: number
): string[] {
  let einzel = 0;
  let doppel = 0;
  const warnungen: string[] = [];

  for (const val of einsatzarten) {
    if (!val || val.startsWith('✗')) continue;
    if (val === 'Einzel+Doppel') { einzel++; doppel++; }
    else if (val === 'Einzel') einzel++;
    else if (val === 'Doppel') doppel++;
  }

  einzel += ersatzCount;
  doppel += ersatzCount;

  const gesamt = einzel > doppel ? einzel : doppel;
  if (gesamt > 6) warnungen.push(`Mehr als 6 Spieler aufgestellt (${gesamt})`);
  if (einzel < einzelReq) warnungen.push(`Nur ${einzel}/${einzelReq} Einzel-Spieler`);
  if (doppel < doppelReq) warnungen.push(`Nur ${doppel}/${doppelReq} Doppel-Spieler`);

  return warnungen;
}

describe('Aufstellung validation logic', () => {
  const E = 4; // einzel requirement
  const D = 2; // doppel requirement

  it('4x Einzel+Doppel → keine Warnungen', () => {
    expect(validateAufstellungLogic(
      ['Einzel+Doppel', 'Einzel+Doppel', 'Einzel+Doppel', 'Einzel+Doppel'], 0, E, D
    )).toHaveLength(0);
  });

  it('warns when not enough Einzel', () => {
    const w = validateAufstellungLogic(
      ['Einzel', 'Einzel', 'Einzel', 'Doppel', 'Doppel'], 0, E, D
    );
    expect(w).toContain(`Nur 3/${E} Einzel-Spieler`);
  });

  it('warns when not enough Doppel', () => {
    const w = validateAufstellungLogic(
      ['Einzel', 'Einzel', 'Einzel', 'Einzel', 'Doppel'], 0, E, D
    );
    expect(w).toContain(`Nur 1/${D} Doppel-Spieler`);
  });

  it('1 Ersatzspieler zählt für beide', () => {
    const w = validateAufstellungLogic(
      ['Einzel', 'Einzel', 'Einzel', 'Doppel'], 1, E, D
    );
    expect(w).toHaveLength(0);
  });

  it('3 Ersatzspieler liefern 3 Einzel, brauchen noch 1 Kader-Einzel', () => {
    const w = validateAufstellungLogic([], 3, E, D);
    expect(w).toContain(`Nur 3/${E} Einzel-Spieler`);
    expect(w).not.toContain(`Nur 3/${D} Doppel-Spieler`); // Doppel: 3 ist ok
  });

  it('warns at 7 total players (6 Einzel+Doppel + 1 Ersatz)', () => {
    const w = validateAufstellungLogic(
      ['Einzel+Doppel', 'Einzel+Doppel', 'Einzel+Doppel', 'Einzel+Doppel', 'Einzel+Doppel', 'Einzel+Doppel'], 1, E, D
    );
    expect(w.some(x => x.includes('Mehr als 6'))).toBe(true);
  });

  it('skips ✗ entries (absent players)', () => {
    const w = validateAufstellungLogic(
      ['Einzel+Doppel', '✗ Urlaub', 'Einzel+Doppel', 'Einzel+Doppel', 'Einzel+Doppel'], 0, E, D
    );
    expect(w).toHaveLength(0);
  });

  it('empty array → warns for both', () => {
    const w = validateAufstellungLogic([], 0, E, D);
    expect(w).toContain(`Nur 0/${E} Einzel-Spieler`);
    expect(w).toContain(`Nur 0/${D} Doppel-Spieler`);
  });

  it('mixed types count correctly', () => {
    const w = validateAufstellungLogic(
      ['Einzel+Doppel', 'Einzel+Doppel', 'Einzel', 'Einzel', 'Doppel', 'Doppel'], 0, E, D
    );
    expect(w).toHaveLength(0);
  });
});
