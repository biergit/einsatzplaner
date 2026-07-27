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
  if (lastRow <= 1) return;

  const allAbw = buildAbwesenheitenIndex(abwesenheitenSheet);
  const dates = readSaisonDates(saisonSheet, lastRow);

  fillPlayerPresenceFast(saisonSheet, dates, aktiveSpieler, allAbw);
  fillEinsatzartenFast(saisonSheet, dates, aktiveSpieler, allAbw);
  validateAllRowsFast(saisonSheet, dates, aktiveSpieler, allAbw);
}

function buildAbwesenheitenIndex(
  sheet: GoogleAppsScript.Spreadsheet.Sheet
): Map<string, Map<string, string>> {
  const lastRow = sheet.getLastRow();
  const index = new Map<string, Map<string, string>>();

  if (lastRow <= 1) return index;

  const data = sheet.getRange(2, 1, lastRow - 1, COL_ABWESENHEITEN.Kommentar).getValues();

  for (const row of data) {
    const name = String(row[COL_ABWESENHEITEN.Spieler - 1]).trim();
    const von = toDate(row[COL_ABWESENHEITEN.Von - 1]);
    const bis = toDate(row[COL_ABWESENHEITEN.Bis - 1]);
    const kommentar = String(row[COL_ABWESENHEITEN.Kommentar - 1]).trim();
    if (!name || !von || !bis) continue;

    const cursor = new Date(von);
    cursor.setHours(0, 0, 0, 0);
    const end = new Date(bis);
    end.setHours(0, 0, 0, 0);

    const display = `✗ ${kommentar || 'abwesend'}`;

    while (cursor <= end) {
      const key = dateKey(cursor);
      if (!index.has(key)) index.set(key, new Map());
      index.get(key)!.set(name, display);
      cursor.setDate(cursor.getDate() + 1);
    }
  }

  return index;
}

function readSaisonDates(sheet: GoogleAppsScript.Spreadsheet.Sheet, lastRow: number): Date[] {
  const data = sheet.getRange(2, saisonDatumCol(), lastRow - 1, 1).getValues();
  return data.map((r: unknown[]) => toDate(r[0])).filter(Boolean) as Date[];
}

function fillPlayerPresenceFast(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  dates: Date[],
  aktiveSpieler: Spieler[],
  allAbw: Map<string, Map<string, string>>
): void {
  const numPlayers = aktiveSpieler.length;
  const numRows = dates.length;

  for (let pi = 0; pi < numPlayers; pi++) {
    const col = saisonSpielerCol(pi);
    const range = sheet.getRange(2, col, numRows, 1);
    const values = range.getValues() as string[][];
    const name = aktiveSpieler[pi].name;
    let changed = false;

    for (let r = 0; r < numRows; r++) {
      const key = dateKey(dates[r]);
      const dayMap = allAbw.get(key);
      const abwDisplay = dayMap?.get(name);
      const current = String(values[r][0] || '').trim();

      if (abwDisplay && abwDisplay.startsWith('✗')) {
        if (current !== abwDisplay) {
          values[r][0] = abwDisplay;
          changed = true;
        }
      } else if (current && current.startsWith('✗') && !abwDisplay) {
        values[r][0] = '';
        changed = true;
      }
    }

    if (changed) range.setValues(values);
  }
}

function fillEinsatzartenFast(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  dates: Date[],
  aktiveSpieler: Spieler[],
  allAbw: Map<string, Map<string, string>>
): void {
  const numPlayers = aktiveSpieler.length;
  const numRows = dates.length;

  for (let r = 0; r < numRows; r++) {
    const row = r + 2;
    const gegner = String(sheet.getRange(row, saisonGegnerCol()).getValue() || '').trim();
    if (!gegner) continue;

    const statusCell = sheet.getRange(row, saisonStatusCol());
    if (!String(statusCell.getValue() || '').trim()) {
      statusCell.setValue('Geplant');
    }

    const key = dateKey(dates[r]);
    const dayMap = allAbw.get(key);

    const available = aktiveSpieler
      .filter(s => {
        const abw = dayMap?.get(s.name);
        return !abw || !abw.startsWith('✗');
      })
      .sort((a, b) => a.rang - b.rang);

    const top4Names = new Set(available.slice(0, 4).map(s => s.name));

    for (let pi = 0; pi < numPlayers; pi++) {
      const col = saisonSpielerCol(pi);
      const cell = sheet.getRange(row, col);
      const current = String(cell.getValue() || '').trim();
      if (current && !current.startsWith('✗')) continue;

      const name = aktiveSpieler[pi].name;
      if (top4Names.has(name)) {
        cell.setValue('Einzel+Doppel');
      }
    }
  }
}

function validateAllRowsFast(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  dates: Date[],
  aktiveSpieler: Spieler[],
  allAbw: Map<string, Map<string, string>>
): void {
  const numPlayers = aktiveSpieler.length;
  const numRows = dates.length;

  const gegnerData = sheet.getRange(2, saisonGegnerCol(), numRows, 1).getValues() as string[][];
  const playerCols = sheet.getRange(2, saisonSpielerCol(0), numRows, numPlayers).getValues() as string[][];

  const hinweisValues: string[][] = [];
  const NON_TOP4_BG = '#FFF9E6';

  for (let r = 0; r < numRows; r++) {
    const gegner = String(gegnerData[r][0] || '').trim();
    if (!gegner) {
      hinweisValues.push(['']);
      continue;
    }

    const key = dateKey(dates[r]);
    const dayMap = allAbw.get(key);

    let einzel = 0;
    let doppel = 0;
    const warnungen: string[] = [];

    for (let pi = 0; pi < numPlayers; pi++) {
      const val = String(playerCols[r][pi] || '').trim();
      if (!val || val.startsWith('✗')) continue;

      if (val === 'Einzel+Doppel') { einzel++; doppel++; }
      else if (val === 'Einzel') einzel++;
      else if (val === 'Doppel') doppel++;

      const name = aktiveSpieler[pi].name;
      const abwDisplay = dayMap?.get(name);
      if (abwDisplay) warnungen.push(`${name}: ${abwDisplay}`);
    }

    let ersatz = 0;
    for (let ei = 0; ei < 3; ei++) {
      const eName = String(sheet.getRange(r + 2, saisonErsatzCol(ei)).getValue() || '').trim();
      if (eName) { ersatz++; einzel++; doppel++; }
    }

    const gesamt = einzel > doppel ? einzel : doppel;
    const f = SHEET_CONFIG.einstellungen.spielformat;
    if (gesamt > 6) warnungen.push(`Mehr als 6 Spieler aufgestellt (${gesamt})`);
    if (einzel < f.einzel) warnungen.push(`Nur ${einzel}/${f.einzel} Einzel-Spieler`);
    if (doppel < f.doppel * 2) warnungen.push(`Nur ${doppel}/${f.doppel * 2} Doppel-Spieler`);

    hinweisValues.push([warnungen.join(' | ')]);
  }

  sheet.getRange(2, saisonHinweisCol(), numRows, 1).setValues(hinweisValues);

  for (let r = 0; r < numRows; r++) {
    const gegner = String(gegnerData[r][0] || '').trim();
    if (!gegner) continue;

    const key = dateKey(dates[r]);
    const dayMap = allAbw.get(key);

    const available = aktiveSpieler
      .filter(s => {
        const abw = dayMap?.get(s.name);
        return !abw || !abw.startsWith('✗');
      })
      .sort((a, b) => a.rang - b.rang);

    const top4Names = new Set(available.slice(0, 4).map(s => s.name));

    const bgColors: (string | null)[] = [];
    for (let pi = 0; pi < numPlayers; pi++) {
      const val = String(playerCols[r][pi] || '').trim();
      const name = aktiveSpieler[pi].name;
      const notTop4 = val && !val.startsWith('✗') && !top4Names.has(name);
      bgColors.push(notTop4 ? NON_TOP4_BG : null);
    }

    sheet.getRange(r + 2, saisonSpielerCol(0), 1, numPlayers).setBackgrounds([bgColors]);
  }
}

function toDate(val: unknown): Date | null {
  if (val instanceof Date) return val;
  if (typeof val === 'string' && val) {
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

function dateKey(d: Date): string {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

function pad2(n: number): string {
  return n.toString().padStart(2, '0');
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
    const von = toDate(row[COL_ABWESENHEITEN.Von - 1]);
    const bis = toDate(row[COL_ABWESENHEITEN.Bis - 1]);
    if (!von || !bis) continue;
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
    if (eName) { ersatz++; einzel++; doppel++; }
  }

  const gesamt = einzel > doppel ? einzel : doppel;
  if (gesamt > 6) warnungen.push(`Mehr als 6 Spieler aufgestellt (${gesamt})`);

  const f = SHEET_CONFIG.einstellungen.spielformat;
  if (einzel < f.einzel) warnungen.push(`Nur ${einzel}/${f.einzel} Einzel-Spieler`);
  if (einzel > f.einzel) warnungen.push(`Mehr als ${f.einzel} Einzel-Spieler (${einzel})`);
  if (doppel < f.doppel * 2) warnungen.push(`Nur ${doppel}/${f.doppel * 2} Doppel-Spieler`);
  if (doppel > f.doppel * 2) warnungen.push(`Mehr als ${f.doppel * 2} Doppel-Spieler (${doppel})`);

  return { einzelCount: einzel, doppelCount: doppel, ersatzCount: ersatz, gesamtCount: gesamt, warnungen };
}
