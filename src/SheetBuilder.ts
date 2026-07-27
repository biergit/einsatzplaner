/// <reference path="ConfigTypes.ts" />

const SHEET_NAMES = {
  DOKUMENTATION: 'Dokumentation',
  SPIELER: 'Spieler',
  ABWESENHEITEN: 'Abwesenheiten',
  SAISON: 'Saison',
  AENDERUNGSLOG: 'Änderungslog',
} as const;

const HEADER_COLOR = '#4A90D9';
const HEADER_FONT_COLOR = '#FFFFFF';

function getOrCreateSheet(name: string): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName(name);
  if (existing) {
    ss.deleteSheet(existing);
  }
  return ss.insertSheet(name);
}

function formatHeader(sheet: GoogleAppsScript.Spreadsheet.Sheet, numCols: number): void {
  const range = sheet.getRange(1, 1, 1, numCols);
  range.setFontWeight('bold');
  range.setBackground(HEADER_COLOR);
  range.setFontColor(HEADER_FONT_COLOR);
  range.setHorizontalAlignment('center');
}

function buildDokumentationSheet(einstellungen: Einstellungen): GoogleAppsScript.Spreadsheet.Sheet {
  const sheet = getOrCreateSheet(SHEET_NAMES.DOKUMENTATION);
  sheet.setColumnWidth(1, 380);
  sheet.setColumnWidth(2, 480);
  sheet.setColumnWidth(3, 480);

  const format = einstellungen.spielformat;

  const rows: string[][] = [
    ['EINSATZPLANER – DOKUMENTATION', '', ''],
    [`${einstellungen.teamName} – Saison ${einstellungen.saison}`, '', ''],
    ['', '', ''],
    ['SHEETS IM ÜBERBLICK', '', ''],
    ['Spieler', 'Stammdaten aller Teammitglieder', 'Name, Email (optional), Rang, Aktiv, Änderungen melden, Rolle'],
    ['Abwesenheiten', 'Von den Spielern gepflegt', 'Spieler, Von, Bis, Kommentar'],
    ['Saison', 'Ein Tag pro Zeile. Spieltage werden durch Gegner-Eintrag markiert.', 'Datum–Heim/Auswärts, pro Spieler eine Spalte mit Einsatzart, 3 Ersatzspieler, Status, Hinweis'],
    ['Änderungslog', 'Automatisch (versteckt)', 'Protokoll aller Änderungen'],
    ['', '', ''],
    ['SPIELER-SPALTEN', '', ''],
    ['Name, Email, Rang, Aktiv, Änderungen melden, Rolle', '', 'siehe oben'],
    ['', '', ''],
    ['ABWESENHEITEN-SPALTEN', '', ''],
    ['Spieler, Von, Bis, Kommentar', '', ''],
    ['', '', ''],
    ['SAISON-SPALTEN', '', ''],
    ['Datum', 'Vorausgefüllt (saisonBeginn–saisonEnde)', ''],
    ['Gegner', 'Vom Kapitän gefüllt – macht den Tag zum Spieltag', ''],
    ['Startzeit', 'Optional', ''],
    ['Heim / Auswärts', 'Dropdown: Heim, Auswärts', ''],
    ['Pro Spieler (Name als Spaltenkopf)', `Dropdown: ${ALLE_AUFSTELLUNGS_TYPEN.join(', ')}. Bei Abwesenheit erscheint ✗ Kommentar.`, ''],
    ['Ersatzspieler 1–3', 'Freitext – Namen von Gastspielern', ''],
    ['Status', 'Geplant oder Final', 'Bei Abwesenheits-Änderung wird Final automatisch auf Geplant zurückgesetzt'],
    ['Hinweis', 'Automatisch', `Warnt wenn nicht genau ${format.einzel} Einzel + ${format.doppel} Doppel abgedeckt sind, oder abwesende Spieler eingesetzt sind`],
    ['', '', ''],
    ['VALIDIERUNG', '', ''],
    ['Einzel', `Mindestens ${format.einzel} Spieler mit Einzel oder Einzel+Doppel`, 'Ersatzspieler zählen als Einzel+Doppel'],
    ['Doppel', `Mindestens ${format.doppel} Spieler mit Doppel oder Einzel+Doppel`, 'Ersatzspieler zählen als Einzel+Doppel'],
    ['Maximum', '6 Spieler insgesamt (Kader + Ersatz)', ''],
    ['', '', ''],
    ['MENÜ', '', ''],
    ['Sheet neu aufbauen', 'Löscht alles und baut neu. Keine E-Mails.', ''],
    ['Daten exportieren', 'JSON nach Google Drive', ''],
    ['Aufstellungen generieren', 'Füllt leere Einsatzart-Zellen nach Rang. Setzt ✗ bei Abwesenden.', ''],
    ['Finalisieren + Emails senden', 'Geplant→Final, versendet E-Mails an eingesetzte Spieler', ''],
    ['', '', ''],
    ['BENACHRICHTIGUNGEN', '', ''],
    ['Änderungs-Mail', `Nach ${einstellungen.debounceMinuten} Min. Inaktivität an alle mit "Änderungen melden"`, 'Nicht an den Bearbeiter selbst'],
    ['Aufstellungs-Mail', 'Bei Finalisierung an eingesetzte Spieler', ''],
    ['Kapitän-Hinweis', 'Spieler ohne E-Mail werden dem Kapitän gemeldet', ''],
  ];

  const numCols = 3;
  sheet.getRange(1, 1, rows.length, numCols).setValues(rows);
  sheet.getRange(1, 1, 2, numCols).setFontWeight('bold').setFontSize(14).setBackground('#E8F0FE');

  const sectionRows = [4, 10, 14, 17, 22, 28, 33, 37, 41, 49];
  for (const r of sectionRows) {
    sheet.getRange(r, 1, 1, numCols).setFontWeight('bold').setFontSize(11)
      .setBackground(HEADER_COLOR).setFontColor(HEADER_FONT_COLOR);
  }

  sheet.setFrozenRows(0);
  return sheet;
}

function buildSpielerSheet(config: SheetConfig): GoogleAppsScript.Spreadsheet.Sheet {
  const sheet = getOrCreateSheet(SHEET_NAMES.SPIELER);
  const headers = ['Name', 'Email', 'Rang', 'Aktiv', 'Änderungen melden', 'Rolle'];
  const numCols = headers.length;

  sheet.getRange(1, 1, 1, numCols).setValues([headers]);
  formatHeader(sheet, numCols);

  if (config.spieler.length > 0) {
    const data = config.spieler.map(s => [
      s.name, s.email, s.rang, true, s.aenderungenMelden, s.rolle,
    ]);
    sheet.getRange(2, 1, data.length, numCols).setValues(data);
  }

  sheet.setColumnWidth(COL_SPIELER.Name, 200);
  sheet.setColumnWidth(COL_SPIELER.Email, 280);
  sheet.setColumnWidth(COL_SPIELER.Rang, 60);
  sheet.setColumnWidth(COL_SPIELER.Aktiv, 60);
  sheet.setColumnWidth(COL_SPIELER.AenderungenMelden, 140);
  sheet.setColumnWidth(COL_SPIELER.Rolle, 100);

  const lastRow = Math.max(sheet.getMaxRows() - 1, config.spieler.length + 50);

  sheet.getRange(2, COL_SPIELER.Rang, lastRow, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberGreaterThan(0).setAllowInvalid(true).build()
  );
  sheet.getRange(2, COL_SPIELER.Aktiv, lastRow, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().build()
  );
  sheet.getRange(2, COL_SPIELER.AenderungenMelden, lastRow, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().build()
  );
  sheet.getRange(2, COL_SPIELER.Rolle, lastRow, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Kapitän']).setAllowInvalid(true).build()
  );

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  return sheet;
}

function buildAbwesenheitenSheet(config: SheetConfig): GoogleAppsScript.Spreadsheet.Sheet {
  const sheet = getOrCreateSheet(SHEET_NAMES.ABWESENHEITEN);
  const headers = ['Spieler', 'Abwesenheit von', 'Abwesenheit bis', 'Kommentar'];
  const numCols = headers.length;

  sheet.getRange(1, 1, 1, numCols).setValues([headers]);
  formatHeader(sheet, numCols);

  if (config.abwesenheiten.length > 0) {
    const data = config.abwesenheiten.map(a => [
      a.spieler,
      `${pad(a.von.getDate())}.${pad(a.von.getMonth() + 1)}.${a.von.getFullYear()}`,
      `${pad(a.bis.getDate())}.${pad(a.bis.getMonth() + 1)}.${a.bis.getFullYear()}`,
      a.kommentar,
    ]);
    sheet.getRange(2, 1, data.length, numCols).setValues(data);
  }

  sheet.setColumnWidth(COL_ABWESENHEITEN.Spieler, 200);
  sheet.setColumnWidth(COL_ABWESENHEITEN.Von, 140);
  sheet.setColumnWidth(COL_ABWESENHEITEN.Bis, 140);
  sheet.setColumnWidth(COL_ABWESENHEITEN.Kommentar, 350);

  const lastRow = sheet.getMaxRows() - 1;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (spielerSheet) {
    const spielerLastRow = spielerSheet.getLastRow();
    if (spielerLastRow > 1) {
      sheet.getRange(2, COL_ABWESENHEITEN.Spieler, lastRow, 1).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInRange(spielerSheet.getRange(2, COL_SPIELER.Name, spielerLastRow - 1, 1))
          .setAllowInvalid(true).build()
      );
    }
  }

  sheet.getRange(2, COL_ABWESENHEITEN.Von, lastRow, 2).setNumberFormat('DD.MM.YYYY');
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('muss')
      .setBackground('#FFF2CC')
      .setRanges([sheet.getRange(1, COL_ABWESENHEITEN.Kommentar, sheet.getMaxRows(), 1)])
      .build(),
  ]);
  sheet.setFrozenRows(1);
  return sheet;
}

function buildSaisonSheet(config: SheetConfig): GoogleAppsScript.Spreadsheet.Sheet {
  const sheet = getOrCreateSheet(SHEET_NAMES.SAISON);
  const numCols = saisonColCount();

  const headers: string[] = ['Datum', 'Gegner', 'Startzeit', 'Heim / Auswärts'];
  for (const s of config.spieler) {
    headers.push(s.name);
  }
  headers.push('Ersatzspieler 1', 'Ersatzspieler 2', 'Ersatzspieler 3', 'Status', 'Hinweis');

  sheet.getRange(1, 1, 1, numCols).setValues([headers]);
  formatHeader(sheet, numCols);

  sheet.setColumnWidth(saisonDatumCol(), 100);
  sheet.setColumnWidth(saisonGegnerCol(), 200);
  sheet.setColumnWidth(saisonStartzeitCol(), 80);
  sheet.setColumnWidth(saisonHeimAuswaertsCol(), 100);
  for (let i = 0; i < config.spieler.length; i++) {
    sheet.setColumnWidth(saisonSpielerCol(i), 130);
  }
  sheet.setColumnWidth(saisonErsatzCol(0), 150);
  sheet.setColumnWidth(saisonErsatzCol(1), 150);
  sheet.setColumnWidth(saisonErsatzCol(2), 150);
  sheet.setColumnWidth(saisonStatusCol(), 100);
  sheet.setColumnWidth(saisonHinweisCol(), 350);

  const beginn = new Date(config.einstellungen.saisonBeginn);
  beginn.setHours(0, 0, 0, 0);
  const ende = new Date(config.einstellungen.saisonEnde);
  ende.setHours(0, 0, 0, 0);
  const dayCount = Math.floor((ende.getTime() - beginn.getTime()) / 86400000) + 1;

  for (let row = 0; row < dayCount; row++) {
    const d = new Date(beginn.getTime() + row * 86400000);
    sheet.getRange(row + 2, saisonDatumCol()).setValue(d).setNumberFormat('DD.MM.YYYY');
  }

  const lastRow = dayCount + 1;

  sheet.getRange(2, saisonDatumCol(), lastRow - 1, 1).setNumberFormat('DD.MM.YYYY');

  sheet.getRange(2, saisonHeimAuswaertsCol(), lastRow - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Heim', 'Auswärts']).setAllowInvalid(true).build()
  );

  for (let i = 0; i < config.spieler.length; i++) {
    sheet.getRange(2, saisonSpielerCol(i), lastRow - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(ALLE_AUFSTELLUNGS_TYPEN).setAllowInvalid(false).build()
    );
  }

  sheet.getRange(2, saisonStatusCol(), lastRow - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Geplant', 'Final']).setAllowInvalid(true).build()
  );

  const filterRange = sheet.getRange(1, 1, lastRow, numCols);
  if (filterRange.getFilter()) filterRange.getFilter()!.remove();

  const filter = filterRange.createFilter();
  filter.setColumnFilterCriteria(saisonGegnerCol(),
    SpreadsheetApp.newFilterCriteria().whenCellNotEmpty().build()
  );
  filter.remove();

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(saisonHeimAuswaertsCol());
  return sheet;
}

function buildAenderungslogSheet(): GoogleAppsScript.Spreadsheet.Sheet {
  const sheet = getOrCreateSheet(SHEET_NAMES.AENDERUNGSLOG);
  const headers = ['Zeitstempel', 'Bereich', 'Alter Wert', 'Neuer Wert', 'Bearbeiter'];
  const numCols = headers.length;

  sheet.getRange(1, 1, 1, numCols).setValues([headers]);
  formatHeader(sheet, numCols);

  sheet.setColumnWidth(COL_AENDERUNGSLOG.Zeitstempel, 160);
  sheet.setColumnWidth(COL_AENDERUNGSLOG.Bereich, 150);
  sheet.setColumnWidth(COL_AENDERUNGSLOG.AlterWert, 250);
  sheet.setColumnWidth(COL_AENDERUNGSLOG.NeuerWert, 250);
  sheet.setColumnWidth(COL_AENDERUNGSLOG.Bearbeiter, 200);
  sheet.getRange(1, COL_AENDERUNGSLOG.Zeitstempel, sheet.getMaxRows(), 1)
    .setNumberFormat('DD.MM.YYYY HH:MM:SS');

  sheet.setFrozenRows(1);
  sheet.hideSheet();
  return sheet;
}

function pad(n: number): string {
  return n.toString().padStart(2, '0');
}

function buildAllSheets(config: SheetConfig): void {
  buildDokumentationSheet(config.einstellungen);
  buildSpielerSheet(config);
  buildAbwesenheitenSheet(config);
  buildSaisonSheet(config);
  buildAenderungslogSheet();

  const knownNames: ReadonlySet<string> = new Set(Object.values(SHEET_NAMES));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  for (const s of allSheets) {
    if (!knownNames.has(s.getName())) {
      ss.deleteSheet(s);
    }
  }
  ss.setActiveSheet(ss.getSheetByName(SHEET_NAMES.SAISON)!);
}
