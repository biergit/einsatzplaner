/// <reference path="ConfigTypes.ts" />

const SHEET_NAMES = {
  DOKUMENTATION: 'Dokumentation',
  SPIELER: 'Spieler',
  ABWESENHEITEN: 'Abwesenheiten',
  SPIELTERMINE: 'Spieltermine',
  AUFSTELLUNGEN: 'Aufstellungen',
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

  const rows: string[][] = [
    ['EINSATZPLANER – DOKUMENTATION', '', ''],
    [`${einstellungen.teamName} – Saison ${einstellungen.saison}`, '', ''],
    ['', '', ''],
    ['SHEETS IM ÜBERBLICK', '', ''],
    ['Spieler', 'Stammdaten aller Teammitglieder', 'Name, Email (optional), Rang (1 = stärkster), Aktiv, Änderungen melden, Rolle'],
    ['Abwesenheiten', 'Von den Spielern selbst gepflegt', 'Spieler (Dropdown), Von, Bis, Kommentar'],
    ['Spieltermine', 'Vom Kapitän erfasst', 'Datum, Heim/Gast, Gegner, Ort (optional), Status'],
    ['Aufstellungen', 'Autom. generiert, manuell anpassbar', 'Termin, Typ (Einzel 1-4 / Doppel), Spieler'],
    ['Änderungslog', 'Automatisch, versteckt', 'Protokoll aller Änderungen durch Teammitglieder'],
    ['', '', ''],
    ['SPALTEN-ERKLÄRUNG: SPIELER', '', ''],
    ['Name', 'Vor- und Nachname', 'Wird in allen Dropdowns verwendet'],
    ['Email', 'Optional – Google-Mail-Adresse', 'Fehlt sie, wird der Kapitän nach dem E-Mail-Versand informiert'],
    ['Rang', '1 = stärkster Spieler, 2 = zweitstärkster usw.', 'Bestimmt die Position in der automatischen Aufstellung'],
    ['Aktiv', 'Checkbox', 'Nur aktive Spieler werden bei der Aufstellungs-Generierung berücksichtigt. Inaktive Spieler (kein Haken) tauchen in keiner Aufstellung auf.'],
    ['Änderungen melden', 'Checkbox', 'Wer hier einen Haken setzt, bekommt nach Sheet-Änderungen durch andere eine E-Mail mit den letzten Änderungen'],
    ['Rolle', 'Dropdown: "Kapitän" oder leer', 'Nur der Kapitän bekommt nach dem E-Mail-Versand einen Hinweis über eingesetzte Spieler ohne E-Mail-Adresse'],
    ['', '', ''],
    ['SPALTEN-ERKLÄRUNG: ABWESENHEITEN', '', ''],
    ['Spieler', 'Dropdown aus Spieler-Liste', 'Von wem die Abwesenheit gemeldet wird'],
    ['Abwesenheit von / bis', 'Datum im Format TT.MM.JJJJ', 'Erster und letzter Tag der Abwesenheit'],
    ['Kommentar', 'Freitext', 'Hinweise wie "nur wenn es sein muss" werden gelb markiert – Spieler wird dann trotzdem bei Aufstellung berücksichtigt'],
    ['', '', ''],
    ['SPALTEN-ERKLÄRUNG: SPIELTERMINE', '', ''],
    ['Datum', 'Datum im Format TT.MM.JJJJ', 'Tag des Spiels'],
    ['Heim / Gast', 'Dropdown: Heim oder Gast', ''],
    ['Gegner', 'Name des gegnerischen Teams', ''],
    ['Ort', 'Optional – Adresse der Spielstätte', ''],
    ['Status', 'Dropdown: Geplant / Finalisiert / Versendet', 'Grün = Geplant, Gelb = Finalisiert, Lila = Versendet'],
    ['', '', ''],
    ['SPALTEN-ERKLÄRUNG: AUFSTELLUNGEN', '', ''],
    ['Termin', 'Dropdown aus Spieltermine', 'Für welches Spiel die Aufstellung gilt'],
    ['Typ', 'Dropdown: Einzel 1-4 oder Doppel', '4 Einzel + 2 Doppel pro Termin. Doppel-Paare legt der Kapitän vor Ort fest.'],
    ['Spieler', 'Dropdown aus Spieler-Liste', 'Wer auf dieser Position spielt'],
    ['', '', ''],
    ['MENÜ "EINSATZPLANER"', '', ''],
    ['Sheet neu aufbauen', '', 'Löscht ALLE Daten und baut die Sheets aus der Konfiguration neu auf. Vorher Daten exportieren!'],
    ['Daten exportieren', '', 'Sichert alle Sheet-Daten als JSON-Datei auf Google Drive (Ordner "Einsatzplaner Exports")'],
    ['Aufstellungen generieren', '', 'Ermittelt verfügbare Spieler (aktiv + nicht abwesend) und weist sie nach Rang den Positionen zu'],
    ['Aufstellungen finalisieren', '', 'Setzt den Status aller geplanten Spieltermine auf "Finalisiert"'],
    ['Emails senden', '', 'Sendet jedem eingesetzten Spieler seinen Einsatzplan per E-Mail. Spieler ohne E-Mail werden dem Kapitän gemeldet.'],
    ['', '', ''],
    ['BENACHRICHTIGUNGEN (AUTOMATISCH)', '', ''],
    ['Änderungslog', '', 'Jede Bearbeitung an Spieler-, Abwesenheiten- oder Spieltermine-Sheet wird protokolliert'],
    ['Änderungs-Mail', '', 'Nach 5 Minuten ohne weitere Änderung erhalten alle mit Haken bei "Änderungen melden" eine E-Mail – außer dem Bearbeiter selbst'],
    ['Aufstellungs-Mail', '', 'Über das Menü manuell ausgelöst. Jeder eingesetzte Spieler mit E-Mail bekommt seinen Plan.'],
    ['Kapitän-Hinweis', '', 'Spieler ohne E-Mail werden dem Kapitän nach dem E-Mail-Versand gemeldet.'],
  ];

  const numCols = 3;
  sheet.getRange(1, 1, rows.length, numCols).setValues(rows);

  sheet.getRange(1, 1, 2, numCols).setFontWeight('bold').setFontSize(14).setBackground('#E8F0FE');
  sheet.getRange(4, 1, 1, numCols).setFontWeight('bold').setFontSize(12).setBackground(HEADER_COLOR).setFontColor(HEADER_FONT_COLOR);
  sheet.getRange(6, 1, 1, numCols).setFontWeight('bold').setBackground('#E8F0FE');
  sheet.getRange(13, 1, 1, numCols).setFontWeight('bold').setBackground(HEADER_COLOR).setFontColor(HEADER_FONT_COLOR);
  sheet.getRange(21, 1, 1, numCols).setFontWeight('bold').setBackground(HEADER_COLOR).setFontColor(HEADER_FONT_COLOR);
  sheet.getRange(30, 1, 1, numCols).setFontWeight('bold').setBackground(HEADER_COLOR).setFontColor(HEADER_FONT_COLOR);
  sheet.getRange(37, 1, 1, numCols).setFontWeight('bold').setBackground(HEADER_COLOR).setFontColor(HEADER_FONT_COLOR);

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
      s.name,
      s.email,
      s.rang,
      true,
      s.aenderungenMelden,
      s.rolle,
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
    SpreadsheetApp.newDataValidation()
      .requireNumberGreaterThan(0)
      .setAllowInvalid(true)
      .build()
  );

  sheet.getRange(2, COL_SPIELER.Aktiv, lastRow, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .build()
  );

  sheet.getRange(2, COL_SPIELER.AenderungenMelden, lastRow, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .build()
  );

  sheet.getRange(2, COL_SPIELER.Rolle, lastRow, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Kapitän'])
      .setAllowInvalid(true)
      .build()
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
    const data = config.abwesenheiten.map(a => {
      const vonStr = `${pad(a.von.getDate())}.${pad(a.von.getMonth() + 1)}.${a.von.getFullYear()}`;
      const bisStr = `${pad(a.bis.getDate())}.${pad(a.bis.getMonth() + 1)}.${a.bis.getFullYear()}`;
      return [a.spieler, vonStr, bisStr, a.kommentar];
    });
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
      const spielerRange = spielerSheet.getRange(2, COL_SPIELER.Name, spielerLastRow - 1, 1);
      sheet.getRange(2, COL_ABWESENHEITEN.Spieler, lastRow, 1).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInRange(spielerRange)
          .setAllowInvalid(true)
          .build()
      );
    }
  }

  sheet.getRange(2, COL_ABWESENHEITEN.Von, lastRow, 2).setNumberFormat('DD.MM.YYYY');

  const kommentarRange = sheet.getRange(1, COL_ABWESENHEITEN.Kommentar, sheet.getMaxRows(), 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('muss')
      .setBackground('#FFF2CC')
      .setRanges([kommentarRange])
      .build(),
  ]);

  sheet.setFrozenRows(1);
  return sheet;
}

function buildSpieltermineSheet(config: SheetConfig): GoogleAppsScript.Spreadsheet.Sheet {
  const sheet = getOrCreateSheet(SHEET_NAMES.SPIELTERMINE);
  const headers = ['Datum', 'Heim / Gast', 'Gegner', 'Ort', 'Status'];
  const numCols = headers.length;

  sheet.getRange(1, 1, 1, numCols).setValues([headers]);
  formatHeader(sheet, numCols);

  if (config.spieltermine.length > 0) {
    const data = config.spieltermine.map(t => {
      const datumStr = `${pad(t.datum.getDate())}.${pad(t.datum.getMonth() + 1)}.${t.datum.getFullYear()}`;
      return [datumStr, t.heimGast, t.gegner, t.ort, t.status];
    });
    sheet.getRange(2, 1, data.length, numCols).setValues(data);
  }

  sheet.setColumnWidth(COL_SPIELTERMINE.Datum, 120);
  sheet.setColumnWidth(COL_SPIELTERMINE.HeimGast, 100);
  sheet.setColumnWidth(COL_SPIELTERMINE.Gegner, 220);
  sheet.setColumnWidth(COL_SPIELTERMINE.Ort, 300);
  sheet.setColumnWidth(COL_SPIELTERMINE.Status, 120);

  const lastRow = sheet.getMaxRows() - 1;

  sheet.getRange(2, COL_SPIELTERMINE.Datum, lastRow, 1).setNumberFormat('DD.MM.YYYY');

  sheet.getRange(2, COL_SPIELTERMINE.HeimGast, lastRow, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Heim', 'Gast'])
      .setAllowInvalid(true)
      .build()
  );

  sheet.getRange(2, COL_SPIELTERMINE.Status, lastRow, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Geplant', 'Finalisiert', 'Versendet'])
      .setAllowInvalid(true)
      .build()
  );

  const statusRange = sheet.getRange(2, COL_SPIELTERMINE.Status, lastRow, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Geplant')
      .setBackground('#D9EAD3')
      .setRanges([statusRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Finalisiert')
      .setBackground('#FFF2CC')
      .setRanges([statusRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Versendet')
      .setBackground('#D9D2E9')
      .setRanges([statusRange])
      .build(),
  ]);

  sheet.setFrozenRows(1);
  return sheet;
}

function buildAufstellungenSheet(): GoogleAppsScript.Spreadsheet.Sheet {
  const sheet = getOrCreateSheet(SHEET_NAMES.AUFSTELLUNGEN);
  const headers = ['Termin', 'Typ', 'Spieler'];
  const numCols = headers.length;

  sheet.getRange(1, 1, 1, numCols).setValues([headers]);
  formatHeader(sheet, numCols);

  sheet.setColumnWidth(COL_AUFSTELLUNGEN.Termin, 120);
  sheet.setColumnWidth(COL_AUFSTELLUNGEN.Typ, 120);
  sheet.setColumnWidth(COL_AUFSTELLUNGEN.Spieler, 200);

  sheet.getRange(1, 1, sheet.getMaxRows(), 1).setNumberFormat('DD.MM.YYYY');

  const lastRow = sheet.getMaxRows() - 1;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (spielerSheet) {
    const spielerLastRow = spielerSheet.getLastRow();
    if (spielerLastRow > 1) {
      const spielerRange = spielerSheet.getRange(2, COL_SPIELER.Name, spielerLastRow - 1, 1);
      sheet.getRange(2, COL_AUFSTELLUNGEN.Spieler, lastRow, 1).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInRange(spielerRange)
          .setAllowInvalid(true)
          .build()
      );
    }
  }

  const termineSheet = ss.getSheetByName(SHEET_NAMES.SPIELTERMINE);
  if (termineSheet) {
    const termineLastRow = termineSheet.getLastRow();
    if (termineLastRow > 1) {
      const termineRange = termineSheet.getRange(2, COL_SPIELTERMINE.Datum, termineLastRow - 1, 1);
      sheet.getRange(2, COL_AUFSTELLUNGEN.Termin, lastRow, 1).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInRange(termineRange)
          .setAllowInvalid(true)
          .build()
      );
    }
  }

  sheet.getRange(2, COL_AUFSTELLUNGEN.Typ, lastRow, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList([
        'Einzel 1', 'Einzel 2', 'Einzel 3', 'Einzel 4', 'Doppel',
      ])
      .setAllowInvalid(true)
      .build()
  );

  sheet.setFrozenRows(1);
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

  sheet.getRange(1, COL_AENDERUNGSLOG.Zeitstempel, sheet.getMaxRows(), 1).setNumberFormat('DD.MM.YYYY HH:MM:SS');

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
  buildSpieltermineSheet(config);
  buildAbwesenheitenSheet(config);
  buildAufstellungenSheet();
  buildAenderungslogSheet();

  const knownNames: ReadonlySet<string> = new Set(Object.values(SHEET_NAMES));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  for (const s of allSheets) {
    if (!knownNames.has(s.getName())) {
      ss.deleteSheet(s);
    }
  }

  ss.setActiveSheet(ss.getSheetByName(SHEET_NAMES.DOKUMENTATION)!);
}
