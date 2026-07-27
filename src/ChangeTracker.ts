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
  if (props.getProperty('SHEET_BUILDER_RUNNING') === 'true') {
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  if (
    sheetName === SHEET_NAMES.AENDERUNGSLOG ||
    sheetName === SHEET_NAMES.DOKUMENTATION
  ) {
    return;
  }

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
}

function logChange(entry: ChangeEntry): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAMES.AENDERUNGSLOG);
  if (!logSheet) {
    return;
  }

  logSheet.showSheet();
  const lastRow = logSheet.getLastRow();
  const row = lastRow + 1;

  logSheet.getRange(row, COL_AENDERUNGSLOG.Zeitstempel, 1, 5).setValues([[
    entry.timestamp,
    entry.bereich,
    entry.alterWert,
    entry.neuerWert,
    entry.bearbeiter,
  ]]);

  logSheet.hideSheet();
}

function resetDebounceTimer(): void {
  if (debounceTimer) {
    ScriptApp.deleteTrigger(debounceTimer);
  }

  const minuten = SHEET_CONFIG.einstellungen.debounceMinuten;
  debounceTimer = ScriptApp.newTrigger('onDebounceTimer')
    .timeBased()
    .after(minuten * 60 * 1000)
    .create();
}

function onDebounceTimer(): void {
  if (debounceTimer) {
    ScriptApp.deleteTrigger(debounceTimer);
    debounceTimer = null;
  }

  sendChangeNotification();
}

function sendChangeNotification(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAMES.AENDERUNGSLOG);
  if (!logSheet) {
    return;
  }

  logSheet.showSheet();
  const lastRow = logSheet.getLastRow();

  const recentChanges = readRecentChanges(logSheet, lastRow);
  logSheet.hideSheet();

  if (recentChanges.length === 0) {
    return;
  }

  const empfaenger = getAenderungenMeldenEmpfaenger(ss);
  if (empfaenger.length === 0) {
    return;
  }

  const body = `Die folgenden Änderungen wurden im Einsatzplaner vorgenommen:\n\n${recentChanges.join('\n')}`;
  const subject = `Einsatzplaner – Änderungen vom ${Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM.yyyy HH:mm')}`;

  for (const email of empfaenger) {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
    });
  }
}

function readRecentChanges(
  logSheet: GoogleAppsScript.Spreadsheet.Sheet,
  lastRow: number
): string[] {
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
  if (!spielerSheet) {
    return [];
  }

  const lastRow = spielerSheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  const currentUser = Session.getActiveUser().getEmail();
  const data = spielerSheet.getRange(2, 1, lastRow - 1, COL_SPIELER.AenderungenMelden).getValues();
  const empfaenger: string[] = [];

  for (const row of data) {
    const email = String(row[COL_SPIELER.Email - 1]).trim();
    const melden = row[COL_SPIELER.AenderungenMelden - 1];

    if (!email || !email.includes('@')) {
      continue;
    }

    if (email === currentUser) {
      continue;
    }

    if (melden === true || melden === 'TRUE') {
      empfaenger.push(email);
    }
  }

  return empfaenger;
}
