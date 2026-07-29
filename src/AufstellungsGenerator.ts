/// <reference path="ConfigTypes.ts" />

function buildNameToColIndex(): Map<string, number> {
  const m = new Map<string, number>();
  SHEET_CONFIG.spieler.forEach((s, i) => m.set(s.name, i));
  return m;
}

function generateAufstellungen(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  const abwesenheitenSheet = ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN);
  const saisonSheet = ss.getSheetByName(SHEET_NAMES.SAISON);

  if (!spielerSheet || !abwesenheitenSheet || !saisonSheet) {
    SpreadsheetApp.getUi().alert('Fehler', 'Nicht alle Sheets vorhanden.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const allSpieler = readSpieler(spielerSheet);
  const lastRow = saisonSheet.getLastRow();
  if (lastRow <= 1) return;

  const allAbw = buildAbwesenheitenIndex(abwesenheitenSheet);
  const dates = readSaisonDates(saisonSheet, lastRow);

  fillPlayerPresenceFast(saisonSheet, dates, allSpieler.map(s => s.name), allAbw);
  fillEinsatzartenFast(saisonSheet, dates, allSpieler, allAbw);
  validateAllRowsFast(saisonSheet, dates, allSpieler, allAbw);
}

function buildAbwesenheitenIndex(
  sheet: GoogleAppsScript.Spreadsheet.Sheet
): Map<string, Map<string, string>> {
  const lastRow = sheet.getLastRow();
  const raw = new Map<string, Map<string, string[]>>();

  if (lastRow <= 1) return new Map();

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

    const label = kommentar || 'abwesend';

    while (cursor <= end) {
      const key = dateKey(cursor);
      if (!raw.has(key)) raw.set(key, new Map());
      const dayMap = raw.get(key)!;
      if (!dayMap.has(name)) dayMap.set(name, []);
      dayMap.get(name)!.push(label);
      cursor.setDate(cursor.getDate() + 1);
    }
  }

  const index = new Map<string, Map<string, string>>();
  for (const [key, dayMap] of raw) {
    const merged = new Map<string, string>();
    for (const [name, kommentare] of dayMap) {
      merged.set(name, `✗ ${kommentare.join(', ')}`);
    }
    index.set(key, merged);
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
  playerNames: string[],
  allAbw: Map<string, Map<string, string>>
): void {
  const numRows = dates.length;

  const nameToColIndex = buildNameToColIndex();

  for (const name of playerNames) {
    const colIndex = nameToColIndex.get(name);
    if (colIndex === undefined) continue;

    const col = saisonSpielerCol(colIndex);
    const range = sheet.getRange(2, col, numRows, 1);
    const values = range.getValues() as string[][];
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
  allSpieler: Spieler[],
  allAbw: Map<string, Map<string, string>>
): void {
  const numRows = dates.length;

  const nameToColIndex = buildNameToColIndex();

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

    const available = allSpieler
      .filter(s => {
        const abw = dayMap?.get(s.name);
        return !abw || !abw.startsWith('✗');
      })
      .sort((a, b) => a.rang - b.rang);

    const top4Names = new Set(available.slice(0, 4).map(s => s.name));

    for (const spieler of allSpieler) {
      const colIndex = nameToColIndex.get(spieler.name);
      if (colIndex === undefined) continue;

      const col = saisonSpielerCol(colIndex);
      const cell = sheet.getRange(row, col);
      const current = String(cell.getValue() || '').trim();
      if (current && !current.startsWith('✗')) continue;

      if (top4Names.has(spieler.name)) {
        cell.setValue('Einzel+Doppel');
      }
    }
  }
}

function validateAllRowsFast(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  dates: Date[],
  allSpieler: Spieler[],
  allAbw: Map<string, Map<string, string>>
): void {
  const configPlayerCount = SHEET_CONFIG.spieler.length;
  const numRows = dates.length;

  const gegnerData = sheet.getRange(2, saisonGegnerCol(), numRows, 1).getValues() as string[][];
  const playerCols = sheet.getRange(2, saisonSpielerCol(0), numRows, configPlayerCount).getValues() as string[][];
  const ersatzCols = sheet.getRange(2, saisonErsatzCol(0), numRows, 3).getValues();

  const validierungValues: string[][] = [];

  const spielerList: Spieler[] = SHEET_CONFIG.spieler.map(s => ({
    name: s.name, email: '', rang: s.rang, aenderungenMelden: false, rolle: ''
  }));

  for (let r = 0; r < numRows; r++) {
    const gegner = String(gegnerData[r][0] || '').trim();
    if (!gegner) {
      validierungValues.push(['']);
      continue;
    }

    const ersatzVals = [
      String(ersatzCols[r][0] || ''),
      String(ersatzCols[r][1] || ''),
      String(ersatzCols[r][2] || ''),
    ];
    validierungValues.push([computeValidierung(playerCols[r], dates[r], spielerList, allAbw, ersatzVals)]);
  }

  sheet.getRange(2, saisonValidierungCol(), numRows, 1).setValues(validierungValues);
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

function readSpieler(sheet: GoogleAppsScript.Spreadsheet.Sheet): Spieler[] {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, COL_SPIELER.Rolle).getValues();
  const spieler: Spieler[] = [];

  for (const row of data) {
    const name = String(row[COL_SPIELER.Name - 1]).trim();
    if (!name) continue;
    let rang = Number(row[COL_SPIELER.Rang - 1]);
    if (isNaN(rang) || rang <= 0) rang = 99;
    const meldenRaw = row[COL_SPIELER.AenderungenMelden - 1];
    spieler.push({
      name,
      email: String(row[COL_SPIELER.Email - 1]),
      rang,
      aenderungenMelden: meldenRaw === true || String(meldenRaw).toUpperCase() === 'TRUE',
      rolle: '',
    });
  }
  return spieler;
}
