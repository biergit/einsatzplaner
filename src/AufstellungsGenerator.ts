/// <reference path="ConfigTypes.ts" />

function generateAufstellungen(): string[] {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  const abwesenheitenSheet = ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN);
  const saisonSheet = ss.getSheetByName(SHEET_NAMES.SAISON);

  if (!spielerSheet || !abwesenheitenSheet || !saisonSheet) {
    SpreadsheetApp.getUi().alert('Fehler', 'Nicht alle Sheets vorhanden.', SpreadsheetApp.getUi().ButtonSet.OK);
    return [];
  }

  const allSpieler = readSpieler(spielerSheet);
  const lastRow = saisonSheet.getLastRow();
  if (lastRow <= 1) return [];

  const warnings: string[] = [];
  if (allSpieler.length > SHEET_CONFIG.spieler.length) {
    const extra = allSpieler.slice(SHEET_CONFIG.spieler.length).map(s => s.name).join(', ');
    warnings.push(`Keine Saison-Spalte für: ${extra}. Bitte die Konfiguration (data/spieler.tsv) erweitern und das Sheet neu aufbauen.`);
  }

  syncSaisonSheet(saisonSheet, abwesenheitenSheet, allSpieler);

  const allAbw = buildAbwesenheitenIndex(abwesenheitenSheet);
  const dates = readSaisonDates(saisonSheet, lastRow);

  fillPlayerPresenceFast(saisonSheet, dates, allSpieler.map(s => s.name), allAbw);
  fillEinsatzartenFast(saisonSheet, dates, allSpieler, allAbw);
  validateAllRowsFast(saisonSheet, dates, allSpieler, allAbw);
  refreshSaisonFormatting(saisonSheet, allSpieler);
  return warnings;
}

/**
 * Gleicht das Saison-Sheet mit den aktuellen Spielern des Spieler-Sheets ab.
 * Die Saison-Spalten folgen positionsbasiert der Reihenfolge des Spieler-Sheets:
 * - Spaltenüberschriften werden auf die aktuellen Namen gesetzt
 * - Abwesenheitszeilen mit alten (Config-)Namen werden auf die neuen Namen umgeschrieben
 */
function syncSaisonSheet(
  saisonSheet: GoogleAppsScript.Spreadsheet.Sheet,
  abwesenheitenSheet: GoogleAppsScript.Spreadsheet.Sheet,
  allSpieler: Spieler[]
): void {
  const playerCount = Math.min(allSpieler.length, SHEET_CONFIG.spieler.length);
  const currentNames = allSpieler.slice(0, playerCount).map(s => s.name);
  saisonSheet.getRange(1, saisonSpielerCol(0), 1, playerCount).setValues([currentNames]);
  renameAbwesenheitenSpieler(abwesenheitenSheet, currentNames);
}

/**
 * Schreibt Abwesenheitszeilen mit veralteten Spielernamen auf die aktuellen
 * Namen um. Zuordnung positional: Config-Position i → aktuelle Position i.
 * Einmaliger In-Memory-Durchlauf (keine Kaskaden-Umbenennungen).
 */
function renameAbwesenheitenSpieler(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  currentNames: string[]
): void {
  const oldToNew = new Map<string, string>();
  for (let i = 0; i < currentNames.length; i++) {
    const old = SHEET_CONFIG.spieler[i].name;
    const neu = currentNames[i];
    if (old !== neu) oldToNew.set(old, neu);
  }
  if (oldToNew.size === 0) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const range = sheet.getRange(2, COL_ABWESENHEITEN.Spieler, lastRow - 1, 1);
  const values = range.getValues();
  let changed = false;
  for (const row of values) {
    const neu = oldToNew.get(String(row[0] || '').trim());
    if (neu) {
      row[0] = neu;
      changed = true;
    }
  }
  if (changed) range.setValues(values);
}

/**
 * Erneuert die bedingten Formatierungen des Saison-Sheets aus den aktuellen
 * Rängen des Spieler-Sheets (nicht aus der Build-Config), damit die
 * Gelb-Markierung (Rang > 4) nach Rang-Änderungen korrekt bleibt.
 */
function refreshSaisonFormatting(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  allSpieler: Spieler[]
): void {
  const spielerMitRang = allSpieler.slice(0, SHEET_CONFIG.spieler.length).map(s => ({
    name: s.name,
    rang: s.rang,
  }));
  sheet.setConditionalFormatRules(buildSaisonConditionalFormats(sheet, sheet.getLastRow(), spielerMitRang));
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
  const playerCount = Math.min(playerNames.length, SHEET_CONFIG.spieler.length);

  for (let pi = 0; pi < playerCount; pi++) {
    const name = playerNames[pi];
    const range = sheet.getRange(2, saisonSpielerCol(pi), numRows, 1);
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
  const playerCount = Math.min(allSpieler.length, SHEET_CONFIG.spieler.length);

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

    for (let pi = 0; pi < playerCount; pi++) {
      const spieler = allSpieler[pi];
      const cell = sheet.getRange(row, saisonSpielerCol(pi));
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

  const spielerList: Spieler[] = [];
  for (let i = 0; i < configPlayerCount; i++) {
    const s = allSpieler[i];
    spielerList.push(s
      ? { name: s.name, email: '', rang: s.rang, aenderungenMelden: false, rolle: '' }
      : { name: SHEET_CONFIG.spieler[i].name, email: '', rang: SHEET_CONFIG.spieler[i].rang, aenderungenMelden: false, rolle: '' });
  }

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
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val === 'string' && val) {
    const trimmed = val.trim();
    const german = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(trimmed);
    if (german) {
      const d = new Date(Number(german[3]), Number(german[2]) - 1, Number(german[1]));
      return isNaN(d.getTime()) ? null : d;
    }
    const d = new Date(trimmed);
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
