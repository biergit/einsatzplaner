/// <reference path="ConfigTypes.ts" />

interface ChangeEntry {
  timestamp: Date;
  bereich: string;
  alterWert: string;
  neuerWert: string;
  bearbeiter: string;
}

let debounceTimer: GoogleAppsScript.Script.Trigger | null = null;

function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SHEET_BUILDER_RUNNING') === 'true') return;

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName === SHEET_NAMES.AENDERUNGSLOG || sheetName === SHEET_NAMES.DOKUMENTATION) return;

  const row = range.getRow();
  const col = range.getColumn();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (sheetName === SHEET_NAMES.SAISON && row >= 2) {
    if (col === saisonGegnerCol()) {
      const newValue = e.value !== undefined ? String(e.value) : '';
      const statusCell = sheet.getRange(row, saisonStatusCol());
      if (newValue.trim()) {
        if (!String(statusCell.getValue() || '').trim()) statusCell.setValue('Geplant');
      } else {
        statusCell.setValue('');
      }
    }
    validateAndUpdateSaisonRow(range);
  }

  if (sheetName === SHEET_NAMES.ABWESENHEITEN && row >= 2) {
    rebuildAbwesenheitenInSaison(ss);
  }

  const isSingleCell = range.getNumRows() === 1 && range.getNumColumns() === 1;
  const oldValue = e.oldValue !== undefined ? String(e.oldValue) : '';
  const newValue = e.value !== undefined ? String(e.value) : '';

  const userEmail = Session.getActiveUser().getEmail();
  const bearbeiter = userEmail && userEmail.includes('@') ? userEmail : 'Unbekannt';

  logChange({
    timestamp: new Date(),
    bereich: `${sheetName}!${range.getA1Notation()}`,
    alterWert: oldValue,
    neuerWert: isSingleCell ? newValue : formatMultiCellEdit(sheet, range),
    bearbeiter: bearbeiter,
  });

  resetDebounceTimer();
}

function rebuildAbwesenheitenInSaison(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): void {
  const saisonSheet = ss.getSheetByName(SHEET_NAMES.SAISON);
  if (!saisonSheet) return;
  const lastRow = saisonSheet.getLastRow();
  if (lastRow <= 1) return;
  const numRows = lastRow - 1;

  const spielerNames = readSpielerNames(ss);
  const numPlayers = spielerNames.length;
  if (numPlayers === 0) return;

  const aktiveSpieler = readAktiveSpieler(ss.getSheetByName(SHEET_NAMES.SPIELER)!);
  if (!aktiveSpieler || aktiveSpieler.length === 0) return;

  const allAbw = buildAbwesenheitenIndex(ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN)!);
  const dates = readSaisonDates(saisonSheet, lastRow);

  const allPlayerCols = saisonSheet.getRange(2, saisonSpielerCol(0), numRows, numPlayers).getValues() as string[][];
  const statusValues = saisonSheet.getRange(2, saisonStatusCol(), numRows, 1).getValues() as string[][];

  let playerChanged = false;
  let statusChanged = false;

  for (let r = 0; r < numRows; r++) {
    const d = dates[r];
    if (!d) continue;
    const key = dateKey(d);
    const dayMap = allAbw.get(key);

    for (let pi = 0; pi < numPlayers; pi++) {
      const current = String(allPlayerCols[r][pi] || '').trim();
      const abwDisplay = dayMap?.get(spielerNames[pi]);

      if (abwDisplay) {
        if (abwDisplay.startsWith('✗')) {
          if (!current || !current.startsWith('✗')) {
            if (current && !current.startsWith('✗') && String(statusValues[r][0] || '').trim() === 'Final') {
              statusValues[r][0] = 'Geplant';
              statusChanged = true;
            }
            allPlayerCols[r][pi] = abwDisplay;
            playerChanged = true;
          }
        }
      } else {
        if (current.startsWith('✗')) {
          allPlayerCols[r][pi] = '';
          playerChanged = true;
        }
      }
    }
  }

  if (playerChanged) {
    saisonSheet.getRange(2, saisonSpielerCol(0), numRows, numPlayers).setValues(allPlayerCols);
  }
  if (statusChanged) {
    saisonSheet.getRange(2, saisonStatusCol(), numRows, 1).setValues(statusValues);
  }

  if (playerChanged || statusChanged) {
    const refreshedAllAbw = buildAbwesenheitenIndex(ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN)!);
    const hinweisValues: string[][] = [];
    for (let r = 0; r < numRows; r++) {
      hinweisValues.push([computeHinweis(allPlayerCols[r], dates[r], aktiveSpieler, refreshedAllAbw, saisonSheet, r + 2)]);
    }
    saisonSheet.getRange(2, saisonHinweisCol(), numRows, 1).setValues(hinweisValues);
  }
}

function formatMultiCellEdit(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  range: GoogleAppsScript.Spreadsheet.Range
): string {
  if (sheet.getName() === SHEET_NAMES.ABWESENHEITEN) {
    const row = range.getRow();
    const data = sheet.getRange(row, 1, 1, COL_ABWESENHEITEN.Kommentar).getValues()[0];
    const spieler = String(data[COL_ABWESENHEITEN.Spieler - 1] || '').trim();
    const von = toDateSafe(data[COL_ABWESENHEITEN.Von - 1]);
    const bis = toDateSafe(data[COL_ABWESENHEITEN.Bis - 1]);
    const kommentar = String(data[COL_ABWESENHEITEN.Kommentar - 1] || '').trim();
    const vonStr = von ? Utilities.formatDate(von, 'Europe/Berlin', 'dd.MM.yyyy') : '-';
    const bisStr = bis ? Utilities.formatDate(bis, 'Europe/Berlin', 'dd.MM.yyyy') : '-';
    return `Spieler: ${spieler || '-'} | Von: ${vonStr} | Bis: ${bisStr} | Kommentar: ${kommentar || '-'}`;
  }
  return `(Mehrere Zellen, ${range.getNumRows()}×${range.getNumColumns()})`;
}

function computeHinweis(
  playerRow: string[],
  datum: Date,
  aktiveSpieler: Spieler[],
  allAbw: Map<string, Map<string, string>>,
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number
): string {
  const key = dateKey(datum);
  const dayMap = allAbw.get(key);
  let einzel = 0;
  let doppel = 0;
  const warnungen: string[] = [];

  for (let pi = 0; pi < playerRow.length; pi++) {
    const val = String(playerRow[pi] || '').trim();
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
    if (String(sheet.getRange(row, saisonErsatzCol(ei)).getValue() || '').trim()) {
      ersatz++; einzel++; doppel++;
    }
  }

  const gesamt = einzel > doppel ? einzel : doppel;
  const f = SHEET_CONFIG.einstellungen.spielformat;
  if (gesamt > 6) warnungen.push(`Mehr als 6 Spieler aufgestellt (${gesamt})`);
  if (einzel < f.einzel) warnungen.push(`Nur ${einzel}/${f.einzel} Einzel-Spieler`);
  if (doppel < f.doppel * 2) warnungen.push(`Nur ${doppel}/${f.doppel * 2} Doppel-Spieler`);

  return warnungen.join(' | ');
}

function toDateSafe(val: unknown): Date | null {
  if (val instanceof Date) return val;
  if (typeof val === 'string' && val) {
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

function validateAndUpdateSaisonRow(editedRange: GoogleAppsScript.Spreadsheet.Range): void {
  const row = editedRange.getRow();
  const sheet = editedRange.getSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const gegner = String(sheet.getRange(row, saisonGegnerCol()).getValue() || '').trim();
  if (!gegner) {
    sheet.getRange(row, saisonHinweisCol()).setValue('');
    return;
  }

  const datum = toDateSafe(sheet.getRange(row, saisonDatumCol()).getValue());
  if (!datum) return;

  const aktiveSpieler = readAktiveSpieler(ss.getSheetByName(SHEET_NAMES.SPIELER)!);
  if (!aktiveSpieler || aktiveSpieler.length === 0) return;

  const allAbw = buildAbwesenheitenIndex(ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN)!);
  const playerRow = sheet.getRange(row, saisonSpielerCol(0), 1, aktiveSpieler.length).getValues()[0] as string[];
  const hinweis = computeHinweis(playerRow, datum, aktiveSpieler, allAbw, sheet, row);
  sheet.getRange(row, saisonHinweisCol()).setValue(hinweis);
}

function logChange(entry: ChangeEntry): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAMES.AENDERUNGSLOG);
  if (!logSheet) return;

  logSheet.showSheet();
  const lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow + 1, COL_AENDERUNGSLOG.Zeitstempel, 1, 5).setValues([[
    entry.timestamp, entry.bereich, entry.alterWert, entry.neuerWert, entry.bearbeiter,
  ]]);
  logSheet.hideSheet();
}

function resetDebounceTimer(): void {
  if (debounceTimer) ScriptApp.deleteTrigger(debounceTimer);
  const minuten = SHEET_CONFIG.einstellungen.debounceMinuten;
  debounceTimer = ScriptApp.newTrigger('onDebounceTimer')
    .timeBased().after(minuten * 60 * 1000).create();
}

function onDebounceTimer(): void {
  if (debounceTimer) { ScriptApp.deleteTrigger(debounceTimer); debounceTimer = null; }
  sendChangeNotification();
}

function sendChangeNotification(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAMES.AENDERUNGSLOG);
  if (!logSheet) return;

  logSheet.showSheet();
  const lastRow = logSheet.getLastRow();
  const recentChanges = readRecentChanges(logSheet, lastRow);
  logSheet.hideSheet();

  if (recentChanges.length === 0) return;

  const empfaenger = getAenderungenMeldenEmpfaenger(ss);
  if (empfaenger.length === 0) return;

  const body = `Änderungen im Einsatzplaner:\n\n${recentChanges.join('\n')}`;
  const subject = `Einsatzplaner – Änderungen vom ${Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM.yyyy HH:mm')}`;

  for (const email of empfaenger) {
    MailApp.sendEmail({ to: email, subject, body });
  }
}

function readRecentChanges(logSheet: GoogleAppsScript.Spreadsheet.Sheet, lastRow: number): string[] {
  const changes: string[] = [];
  for (let row = lastRow; row > Math.max(1, lastRow - 20); row--) {
    const values = logSheet.getRange(row, 1, 1, 5).getValues()[0];
    if (values[COL_AENDERUNGSLOG.Zeitstempel - 1]) {
      changes.push(
        `${values[COL_AENDERUNGSLOG.Zeitstempel - 1]}: ${values[COL_AENDERUNGSLOG.Bereich - 1]} || ${values[COL_AENDERUNGSLOG.AlterWert - 1]} → ${values[COL_AENDERUNGSLOG.NeuerWert - 1]}`
      );
    }
  }
  return changes;
}

function getAenderungenMeldenEmpfaenger(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): string[] {
  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (!spielerSheet) return [];

  const lastRow = spielerSheet.getLastRow();
  if (lastRow <= 1) return [];

  const currentUser = Session.getActiveUser().getEmail();
  const data = spielerSheet.getRange(2, 1, lastRow - 1, COL_SPIELER.AenderungenMelden).getValues();
  const empfaenger: string[] = [];

  for (const row of data) {
    const email = String(row[COL_SPIELER.Email - 1]).trim();
    if (!email || !email.includes('@') || email === currentUser) continue;
    if (row[COL_SPIELER.AenderungenMelden - 1] === true || row[COL_SPIELER.AenderungenMelden - 1] === 'TRUE') {
      empfaenger.push(email);
    }
  }
  return empfaenger;
}
