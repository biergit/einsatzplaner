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

  // Ersatz füllt zuerst fehlende Plätze auf
  let missingEinzel = Math.max(0, einzelReq - einzel);
  let missingDoppel = Math.max(0, doppelReq * 2 - doppel);
  for (let i = 0; i < ersatzCount; i++) {
    if (missingEinzel > 0) { einzel++; missingEinzel--; }
    else if (missingDoppel > 0) { doppel++; missingDoppel--; }
    else { einzel++; doppel++; }
  }

  const gesamt = einzel > doppel ? einzel : doppel;
  if (gesamt > 6) warnungen.push(`Mehr als 6 Spieler aufgestellt (${gesamt})`);
  if (einzel < einzelReq) warnungen.push(`Nur ${einzel}/${einzelReq} Einzel-Spieler`);
  if (einzel > einzelReq) warnungen.push(`Mehr als ${einzelReq} Einzel-Spieler (${einzel})`);
  if (doppel < doppelReq * 2) warnungen.push(`Nur ${doppel}/${doppelReq * 2} Doppel-Spieler`);
  if (doppel > doppelReq * 2) warnungen.push(`Mehr als ${doppelReq * 2} Doppel-Spieler (${doppel})`);

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
    expect(w).toContain(`Nur 1/${D * 2} Doppel-Spieler`);
  });

  it('1 Ersatzspieler füllt zuerst Einzel-Lücke', () => {
    const w = validateAufstellungLogic(
      ['Einzel', 'Einzel', 'Einzel', 'Doppel', 'Doppel', 'Doppel'], 1, E, D
    );
    expect(w).toContain(`Nur 3/${D * 2} Doppel-Spieler`);
    expect(w).not.toContain(`Nur ${E}/${E} Einzel-Spieler`);
  });

  it('3 Ersatzspieler füllen alle die Einzel-Lücke, Doppel bleibt leer', () => {
    const w = validateAufstellungLogic([], 3, E, D);
    expect(w).toContain(`Nur 3/${E} Einzel-Spieler`);
    expect(w).toContain(`Nur 0/${D * 2} Doppel-Spieler`);
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
    expect(w).toContain(`Nur 0/${D * 2} Doppel-Spieler`);
  });

  it('mixed types count correctly', () => {
    const w = validateAufstellungLogic(
      ['Einzel+Doppel', 'Einzel+Doppel', 'Einzel', 'Einzel', 'Doppel', 'Doppel'], 0, E, D
    );
    expect(w).toHaveLength(0);
  });

  it('warns when too many Einzel', () => {
    const w = validateAufstellungLogic(
      ['Einzel+Doppel', 'Einzel+Doppel', 'Einzel+Doppel', 'Einzel+Doppel', 'Einzel'], 0, E, D
    );
    expect(w).toContain(`Mehr als ${E} Einzel-Spieler (5)`);
  });

  it('warns when too many Doppel', () => {
    const w = validateAufstellungLogic(
      ['Einzel+Doppel', 'Einzel+Doppel', 'Einzel+Doppel', 'Einzel+Doppel', 'Doppel'], 0, E, D
    );
    expect(w).toContain(`Mehr als ${D * 2} Doppel-Spieler (5)`);
  });

  it('1 Ersatz with 4 Einzel+Doppel → warns too many Einzel and Doppel', () => {
    const w = validateAufstellungLogic(
      ['Einzel+Doppel', 'Einzel+Doppel', 'Einzel+Doppel', 'Einzel+Doppel'], 1, E, D
    );
    expect(w).toContain(`Mehr als ${E} Einzel-Spieler (5)`);
    expect(w).toContain(`Mehr als ${D * 2} Doppel-Spieler (5)`);
  });
});
