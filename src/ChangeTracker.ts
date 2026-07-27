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

  const oldValue = e.oldValue !== undefined ? String(e.oldValue) : '';
  const newValue = e.value !== undefined ? String(e.value) : '';

  logChange({
    timestamp: new Date(),
    bereich: `${sheetName}!${range.getA1Notation()}`,
    alterWert: oldValue,
    neuerWert: newValue,
    bearbeiter: Session.getActiveUser().getEmail(),
  });

  resetDebounceTimer();

  if (sheetName === SHEET_NAMES.ABWESENHEITEN && range.getRow() >= 2) {
    updateSaisonForAbsenceChange(range);
  }

  if (sheetName === SHEET_NAMES.SAISON && range.getRow() >= 2) {
    validateAndUpdateSaisonRow(range);
  }
}

function updateSaisonForAbsenceChange(editedRange: GoogleAppsScript.Spreadsheet.Range): void {
  const row = editedRange.getRow();
  const sheet = editedRange.getSheet();
  const data = sheet.getRange(row, 1, 1, COL_ABWESENHEITEN.Kommentar).getValues()[0];

  const spielerName = String(data[COL_ABWESENHEITEN.Spieler - 1] || '').trim();
  const von = data[COL_ABWESENHEITEN.Von - 1] instanceof Date
    ? data[COL_ABWESENHEITEN.Von - 1] : new Date(data[COL_ABWESENHEITEN.Von - 1]);
  const bis = data[COL_ABWESENHEITEN.Bis - 1] instanceof Date
    ? data[COL_ABWESENHEITEN.Bis - 1] : new Date(data[COL_ABWESENHEITEN.Bis - 1]);

  if (!spielerName || isNaN(von.getTime()) || isNaN(bis.getTime())) return;

  const kommentar = String(data[COL_ABWESENHEITEN.Kommentar - 1]).trim();
  if (kommentar.toLowerCase().includes('muss')) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const saisonSheet = ss.getSheetByName(SHEET_NAMES.SAISON);
  if (!saisonSheet) return;

  const aktiveSpieler = readAktiveSpieler(ss.getSheetByName(SHEET_NAMES.SPIELER)!);
  if (!aktiveSpieler) return;

  const pi = aktiveSpieler.findIndex(s => s.name === spielerName);
  if (pi < 0) return;

  const saisonLastRow = saisonSheet.getLastRow();
  const col = saisonSpielerCol(pi);

  for (let r = 2; r <= saisonLastRow; r++) {
    const datum = saisonSheet.getRange(r, saisonDatumCol()).getValue();
    if (!(datum instanceof Date)) continue;

    const d = new Date(datum.getFullYear(), datum.getMonth(), datum.getDate());
    if (d < von || d > bis) continue;

    const current = String(saisonSheet.getRange(r, col).getValue() || '').trim();
    if (!current || current.startsWith('✗')) continue;

    saisonSheet.getRange(r, col).setValue(`✗ ${kommentar || 'abwesend'}`);

    const statusCell = saisonSheet.getRange(r, saisonStatusCol());
    if (String(statusCell.getValue()).trim() === 'Final') {
      statusCell.setValue('Geplant');
    }

    const abwesende = readAbwesenheitenFuerDatum(
      ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN)!, datum
    );
    const v = validateAufstellung(saisonSheet, r, aktiveSpieler, abwesende);
    saisonSheet.getRange(r, saisonHinweisCol()).setValue(v.warnungen.join(' | '));
  }
}

function validateAndUpdateSaisonRow(editedRange: GoogleAppsScript.Spreadsheet.Range): void {
  const row = editedRange.getRow();
  const sheet = editedRange.getSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const datum = sheet.getRange(row, saisonDatumCol()).getValue();
  if (!(datum instanceof Date)) return;

  const aktiveSpieler = readAktiveSpieler(ss.getSheetByName(SHEET_NAMES.SPIELER)!);
  if (!aktiveSpieler) return;

  const abwesende = readAbwesenheitenFuerDatum(
    ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN)!, datum
  );

  const v = validateAufstellung(sheet, row, aktiveSpieler, abwesende);
  sheet.getRange(row, saisonHinweisCol()).setValue(v.warnungen.join(' | '));
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
