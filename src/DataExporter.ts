/// <reference path="ConfigTypes.ts" />

interface ExportResult {
  fileId: string;
  fileName: string;
}

function exportAllData(): void {
  const userEmail = Session.getActiveUser().getEmail();
  if (!userEmail || !userEmail.includes('@')) {
    SpreadsheetApp.getUi().alert(
      'Fehler',
      'Export per E-Mail benötigt ein Google-Konto. Deine E-Mail-Adresse konnte nicht ermittelt werden.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const blobs: GoogleAppsScript.Base.BlobSource[] = [];
  const names: string[] = [];

  const spielerBlob = exportSpielerTSV(ss);
  if (spielerBlob) { blobs.push(spielerBlob); names.push('spieler.tsv'); }

  const abwBlob = exportAbwesenheitenTSV(ss);
  if (abwBlob) { blobs.push(abwBlob); names.push('abwesenheiten.tsv'); }

  const saisonBlob = exportSaisonTSV(ss);
  if (saisonBlob) { blobs.push(saisonBlob); names.push('saison.tsv'); }

  const einstellungenJson = JSON.stringify(SHEET_CONFIG.einstellungen, null, 2);
  blobs.push(Utilities.newBlob(einstellungenJson, 'application/json', 'einstellungen.json'));
  names.push('einstellungen.json');

  const timestamp = Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM.yyyy HH:mm');
  MailApp.sendEmail({
    to: userEmail,
    subject: `Einsatzplaner – Datenexport ${timestamp}`,
    body: `Hallo,\n\nanbei die exportierten Rohdaten des Einsatzplaners vom ${timestamp}.\n\n` +
      `Die Dateien können in das data/-Verzeichnis des Projekts kopiert werden,\n` +
      `um das Sheet mit 'Sheet neu aufbauen' wiederherzustellen.\n\n` +
      `Dateien:\n` +
      `${names.map(n => `  – ${n}`).join('\n')}\n\n` +
      `Dein Einsatzplaner-Team`,
    attachments: blobs,
  });
}

function exportSpielerTSV(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): GoogleAppsScript.Base.Blob | null {
  const sheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (!sheet) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, COL_SPIELER.Rolle).getValues();
  let tsv = 'Name\tEmail\tRang\tÄnderungen melden\tRolle\n';
  for (const row of data) {
    const name = String(row[COL_SPIELER.Name - 1]).trim();
    if (!name) continue;
    const email = String(row[COL_SPIELER.Email - 1]).trim();
    const rang = row[COL_SPIELER.Rang - 1];
    const melden = row[COL_SPIELER.AenderungenMelden - 1] === true ? 'ja' : '';
    const rolle = String(row[COL_SPIELER.Rolle - 1]).trim();
    tsv += `${name}\t${email}\t${rang}\t${melden}\t${rolle}\n`;
  }
  return Utilities.newBlob(tsv, 'text/tab-separated-values', 'spieler.tsv');
}

function exportAbwesenheitenTSV(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): GoogleAppsScript.Base.Blob | null {
  const sheet = ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN);
  if (!sheet) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, COL_ABWESENHEITEN.Kommentar).getValues();
  let tsv = 'Spieler/in\tAbwesenheit von\tAbwesenheit bis\tKommentar\n';
  for (const row of data) {
    const spieler = String(row[COL_ABWESENHEITEN.Spieler - 1]).trim();
    if (!spieler) continue;
    const von = formatDate(row[COL_ABWESENHEITEN.Von - 1]);
    const bis = formatDate(row[COL_ABWESENHEITEN.Bis - 1]);
    if (!von || !bis) continue;
    const kommentar = String(row[COL_ABWESENHEITEN.Kommentar - 1]).trim();
    tsv += `${spieler}\t${von}\t${bis}\t${kommentar}\n`;
  }
  return Utilities.newBlob(tsv, 'text/tab-separated-values', 'abwesenheiten.tsv');
}

function exportSaisonTSV(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): GoogleAppsScript.Base.Blob | null {
  const sheet = ss.getSheetByName(SHEET_NAMES.SAISON);
  if (!sheet) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;

  const numCols = saisonColCount();
  const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  let tsv = '';
  for (const row of data) {
    const rowArr: string[] = [];
    for (const cell of row) {
      if (cell instanceof Date) {
        rowArr.push(Utilities.formatDate(cell, 'Europe/Berlin', 'dd.MM.yyyy'));
      } else {
        rowArr.push(String(cell || '').trim());
      }
    }
    tsv += rowArr.join('\t') + '\n';
  }
  return Utilities.newBlob(tsv, 'text/tab-separated-values', 'saison.tsv');
}

function formatDate(val: unknown): string {
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Europe/Berlin', 'dd.MM.yyyy');
  }
  return '';
}
