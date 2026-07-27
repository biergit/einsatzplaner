/// <reference path="ConfigTypes.ts" />

function generateAufstellungen(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  const abwesenheitenSheet = ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN);
  const saisonSheet = ss.getSheetByName(SHEET_NAMES.SAISON);

  if (!spielerSheet || !abwesenheitenSheet || !saisonSheet) {
    SpreadsheetApp.getUi().alert('Fehler', 'Nicht alle Sheets vorhanden.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const aktiveSpieler = readAktiveSpieler(spielerSheet);
  const lastRow = saisonSheet.getLastRow();

  fillPlayerPresence(saisonSheet, abwesenheitenSheet, aktiveSpieler, lastRow);

  for (let row = 2; row <= lastRow; row++) {
    const gegner = String(saisonSheet.getRange(row, saisonGegnerCol()).getValue() || '').trim();
    if (!gegner) continue;

    const datum = saisonSheet.getRange(row, saisonDatumCol()).getValue();
    if (!(datum instanceof Date)) continue;

    const abwesende = readAbwesenheitenFuerDatum(abwesenheitenSheet, datum);
    const abwesendHard = new Set(
      abwesende.filter(a => !a.kommentar.toLowerCase().includes('muss')).map(a => a.spieler)
    );

    fillEmptyEinsatzarten(saisonSheet, row, aktiveSpieler, abwesendHard);
    validateRow(saisonSheet, row, aktiveSpieler, abwesende);
  }
}

function fillPlayerPresence(
  saisonSheet: GoogleAppsScript.Spreadsheet.Sheet,
  abwesenheitenSheet: GoogleAppsScript.Spreadsheet.Sheet,
  aktiveSpieler: Spieler[],
  lastRow: number
): void {
  for (let row = 2; row <= lastRow; row++) {
    const datum = saisonSheet.getRange(row, saisonDatumCol()).getValue();
    if (!(datum instanceof Date)) continue;

    const abwesende = readAbwesenheitenFuerDatum(abwesenheitenSheet, datum);
    const abwMap = new Map<string, Abwesenheit>();
    for (const a of abwesende) abwMap.set(a.spieler, a);

    for (let pi = 0; pi < aktiveSpieler.length; pi++) {
      const name = aktiveSpieler[pi].name;
      const abw = abwMap.get(name);
      const cell = saisonSheet.getRange(row, saisonSpielerCol(pi));
      const current = String(cell.getValue() || '').trim();

      if (abw) {
        const isHard = !abw.kommentar.toLowerCase().includes('muss');
        if (isHard) {
          cell.setValue(`✗ ${abw.kommentar || 'abwesend'}`);
        }
      } else if (!current || current.startsWith('✗')) {
        cell.setValue('');
      }
    }
  }
}

function fillEmptyEinsatzarten(
  saisonSheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  aktiveSpieler: Spieler[],
  abwesendHard: Set<string>
): void {
  const verfuegbare = aktiveSpieler
    .filter(s => !abwesendHard.has(s.name))
    .sort((a, b) => a.rang - b.rang);

  for (let pi = 0; pi < aktiveSpieler.length; pi++) {
    const cell = saisonSheet.getRange(row, saisonSpielerCol(pi));
    const current = String(cell.getValue() || '').trim();

    if (current && !current.startsWith('✗')) continue;

    const name = aktiveSpieler[pi].name;
    if (abwesendHard.has(name)) continue;

    const rank = verfuegbare.findIndex(s => s.name === name);
    if (rank >= 0 && rank < 4) {
      cell.setValue('Einzel+Doppel');
    }
  }
}

function validateRow(
  saisonSheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  aktiveSpieler: Spieler[],
  abwesende: Abwesenheit[]
): void {
  const v = validateAufstellung(saisonSheet, row, aktiveSpieler, abwesende);
  saisonSheet.getRange(row, saisonHinweisCol()).setValue(v.warnungen.join(' | '));
}

function validateAufstellung(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  aktiveSpieler: Spieler[],
  abwesende: Abwesenheit[]
): SaisonValidierung {
  let einzel = 0;
  let doppel = 0;
  let ersatz = 0;
  const warnungen: string[] = [];
  const abwMap = new Map<string, string>();
  for (const a of abwesende) abwMap.set(a.spieler, a.kommentar);

  for (let pi = 0; pi < aktiveSpieler.length; pi++) {
    const val = String(sheet.getRange(row, saisonSpielerCol(pi)).getValue() || '').trim();
    if (!val || val.startsWith('✗')) continue;

    const name = aktiveSpieler[pi].name;
    if (val === 'Einzel+Doppel') { einzel++; doppel++; }
    else if (val === 'Einzel') einzel++;
    else if (val === 'Doppel') doppel++;

    if (abwMap.has(name)) {
      warnungen.push(`${name}: abwesend (${abwMap.get(name)})`);
    }
  }

  for (let ei = 0; ei < 3; ei++) {
    const eName = String(sheet.getRange(row, saisonErsatzCol(ei)).getValue() || '').trim();
    if (eName) {
      ersatz++;
      einzel++;
      doppel++;
    }
  }

  const gesamt = einzel > doppel ? einzel : doppel;
  if (gesamt > 6) warnungen.push(`Mehr als 6 Spieler aufgestellt (${gesamt})`);

  const f = SHEET_CONFIG.einstellungen.spielformat;
  if (einzel < f.einzel) warnungen.push(`Nur ${einzel}/${f.einzel} Einzel-Spieler`);
  if (doppel < f.doppel) warnungen.push(`Nur ${doppel}/${f.doppel} Doppel-Spieler`);

  return { einzelCount: einzel, doppelCount: doppel, ersatzCount: ersatz, gesamtCount: gesamt, warnungen };
}

function readAktiveSpieler(sheet: GoogleAppsScript.Spreadsheet.Sheet): Spieler[] {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, COL_SPIELER.Rolle).getValues();
  const spieler: Spieler[] = [];

  for (const row of data) {
    if (row[COL_SPIELER.Aktiv - 1] === true || row[COL_SPIELER.Aktiv - 1] === 'TRUE') {
      const name = String(row[COL_SPIELER.Name - 1]).trim();
      if (!name) continue;
      let rang = Number(row[COL_SPIELER.Rang - 1]);
      if (isNaN(rang) || rang <= 0) rang = 99;
      spieler.push({
        name,
        email: String(row[COL_SPIELER.Email - 1]),
        rang,
        aenderungenMelden: row[COL_SPIELER.AenderungenMelden - 1] === true,
        rolle: '',
      });
    }
  }
  return spieler;
}

function readAbwesenheitenFuerDatum(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  datum: Date
): Abwesenheit[] {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, COL_ABWESENHEITEN.Kommentar).getValues();
  const result: Abwesenheit[] = [];
  const d = new Date(datum.getFullYear(), datum.getMonth(), datum.getDate());

  for (const row of data) {
    const von = row[COL_ABWESENHEITEN.Von - 1] instanceof Date
      ? row[COL_ABWESENHEITEN.Von - 1] : new Date(row[COL_ABWESENHEITEN.Von - 1]);
    const bis = row[COL_ABWESENHEITEN.Bis - 1] instanceof Date
      ? row[COL_ABWESENHEITEN.Bis - 1] : new Date(row[COL_ABWESENHEITEN.Bis - 1]);
    if (isNaN(von.getTime()) || isNaN(bis.getTime())) continue;
    if (d < von || d > bis) continue;

    result.push({
      spieler: String(row[COL_ABWESENHEITEN.Spieler - 1]),
      von,
      bis,
      kommentar: String(row[COL_ABWESENHEITEN.Kommentar - 1]),
    });
  }
  return result;
}
