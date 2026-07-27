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

  const rows: string[][] = [
    ['EINSATZPLANER – DOKUMENTATION', '', ''],
    [`${einstellungen.teamName} – Saison ${einstellungen.saison}`, '', ''],
    ['', '', ''],
    ['SHEETS IM ÜBERBLICK', '', ''],
    ['Spieler', 'Stammdaten aller Teammitglieder', 'Name, Email (optional), Rang (1 = stärkster), Aktiv, Änderungen melden, Rolle'],
    ['Abwesenheiten', 'Von den Spielern selbst gepflegt', 'Spieler (Dropdown), Von, Bis, Kommentar'],
    ['Saison', 'Zentrale Übersicht – ein Tab pro Kalendertag', 'Datum, Gegner, Startzeit, Heim/Auswärts, Aufstellung, Status, Spieler-Anwesenheit'],
    ['Änderungslog', 'Automatisch, versteckt', 'Protokoll aller Änderungen durch Teammitglieder'],
    ['', '', ''],
    ['SPALTEN-ERKLÄRUNG: SPIELER', '', ''],
    ['Name', 'Vor- und Nachname', 'Wird in allen Dropdowns verwendet'],
    ['Email', 'Optional – Google-Mail-Adresse', 'Fehlt sie, wird der Kapitän nach dem E-Mail-Versand informiert'],
    ['Rang', '1 = stärkster Spieler', 'Bestimmt die Reihenfolge in der automatischen Aufstellung'],
    ['Aktiv', 'Checkbox', 'Nur aktive Spieler werden bei der Aufstellungs-Generierung berücksichtigt'],
    ['Änderungen melden', 'Checkbox', 'Wer hier einen Haken setzt, bekommt nach Sheet-Änderungen durch andere eine E-Mail'],
    ['Rolle', 'Dropdown: "Kapitän" oder leer', 'Der Kapitän bekommt nach dem E-Mail-Versand einen Hinweis über eingesetzte Spieler ohne E-Mail-Adresse'],
    ['', '', ''],
    ['SPALTEN-ERKLÄRUNG: ABWESENHEITEN', '', ''],
    ['Spieler', 'Dropdown aus Spieler-Liste', ''],
    ['Abwesenheit von / bis', 'Datum TT.MM.JJJJ', 'Erster und letzter Tag'],
    ['Kommentar', 'Freitext', '"nur wenn es sein muss" wird gelb markiert – Spieler wird dennoch aufgestellt'],
    ['', '', ''],
    ['SPALTEN-ERKLÄRUNG: SAISON', '', ''],
    ['Datum', 'Automatisch vorausgefüllt', 'Alle Tage von saisonBeginn bis saisonEnde'],
    ['Gegner', 'Vom Kapitän ausgefüllt', 'Ein Eintrag hier macht den Tag zum Spieltag'],
    ['Startzeit', 'Optional', 'Uhrzeit des Spielbeginns'],
    ['Heim/Auswärts', 'Dropdown: Heim oder Auswärts', ''],
    ['Einsatzart (1-6)', 'Einzel+Doppel / Einzel / Doppel', 'Pro Spieler-Slot: welche Rolle der Spieler übernimmt'],
    ['Spieler (1-6)', 'Dropdown + Freitext', 'Wer auf dieser Position spielt (auch Gastspieler möglich)'],
    ['Status', 'Dropdown: Geplant / Final', 'Geplant = Entwurf, Final = bestätigt (löst E-Mails aus)'],
    ['Hinweis', 'Automatisch gefüllt', 'Warnt vor abwesenden/inaktiven Spielern in der Aufstellung'],
    ['Spieler-Spalten', 'Formel (automatisch)', 'Zeigt An- oder Abwesenheit jedes Spielers am jeweiligen Tag'],
    ['', '', ''],
    ['MENÜ "EINSATZPLANER"', '', ''],
    ['Sheet neu aufbauen', '', 'Löscht alle Daten und baut Sheets aus der Konfiguration neu. Keine E-Mails.'],
    ['Daten exportieren', '', 'Sichert alle Daten als JSON nach Google Drive'],
    ['Aufstellungen generieren', '', 'Füllt leere Aufstellungs-Zellen basierend auf Rang + Verfügbarkeit'],
    ['Aufstellung finalisieren + senden', '', 'Setzt Geplant→Final und versendet E-Mails an betroffene Spieler'],
    ['', '', ''],
    ['BENACHRICHTIGUNGEN', '', ''],
    ['Änderungs-Mail', '', 'Nach debounceMinuten Inaktivität erhalten alle mit "Änderungen melden"-Haken eine Mail (nicht der Bearbeiter selbst)'],
    ['Aufstellungs-Mail', '', 'Bei Finalisierung werden neu eingesetzte, entfernte und interessierte Spieler informiert'],
    ['Kapitän-Hinweis', '', 'Spieler ohne E-Mail werden dem Kapitän gemeldet'],
  ];

  const numCols = 3;
  sheet.getRange(1, 1, rows.length, numCols).setValues(rows);

  sheet.getRange(1, 1, 2, numCols).setFontWeight('bold').setFontSize(14).setBackground('#E8F0FE');
  sheet.getRange(4, 1, 1, numCols).setFontWeight('bold').setFontSize(12).setBackground(HEADER_COLOR).setFontColor(HEADER_FONT_COLOR);
  sheet.getRange(6, 1, 1, numCols).setFontWeight('bold').setBackground('#E8F0FE');
  sheet.getRange(14, 1, 1, numCols).setFontWeight('bold').setBackground(HEADER_COLOR).setFontColor(HEADER_FONT_COLOR);
  sheet.getRange(21, 1, 1, numCols).setFontWeight('bold').setBackground(HEADER_COLOR).setFontColor(HEADER_FONT_COLOR);
  sheet.getRange(29, 1, 1, numCols).setFontWeight('bold').setBackground(HEADER_COLOR).setFontColor(HEADER_FONT_COLOR);
  sheet.getRange(42, 1, 1, numCols).setFontWeight('bold').setBackground(HEADER_COLOR).setFontColor(HEADER_FONT_COLOR);
  sheet.getRange(49, 1, 1, numCols).setFontWeight('bold').setBackground(HEADER_COLOR).setFontColor(HEADER_FONT_COLOR);

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

function buildSaisonSheet(config: SheetConfig): GoogleAppsScript.Spreadsheet.Sheet {
  const sheet = getOrCreateSheet(SHEET_NAMES.SAISON);
  const maxSlots = saisonMaxSlots();
  const firstSpielerCol = saisonErsterSpielerCol();
  const lastCol = firstSpielerCol - 1 + config.spieler.length;

  const headers: string[] = ['Datum', 'Gegner', 'Startzeit', 'Heim / Auswärts'];
  for (let i = 0; i < maxSlots; i++) {
    headers.push(`Einsatzart ${i + 1}`);
    headers.push(`Spieler ${i + 1}`);
  }
  headers.push('Status');
  headers.push('Hinweis');
  for (const s of config.spieler) {
    headers.push(s.name);
  }

  const numCols = headers.length;
  sheet.getRange(1, 1, 1, numCols).setValues([headers]);
  formatHeader(sheet, numCols);

  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 100);
  for (let i = 0; i < maxSlots; i++) {
    sheet.setColumnWidth(saisonEinsatzartCol(i), 120);
    sheet.setColumnWidth(saisonSpielerCol(i), 180);
  }
  sheet.setColumnWidth(saisonStatusCol(), 100);
  sheet.setColumnWidth(saisonHinweisCol(), 300);
  for (let i = 0; i < config.spieler.length; i++) {
    sheet.setColumnWidth(firstSpielerCol + i, 130);
  }

  const beginn = new Date(config.einstellungen.saisonBeginn);
  const ende = new Date(config.einstellungen.saisonEnde);
  const dayCount = Math.floor((ende.getTime() - beginn.getTime()) / (1000 * 60 * 60 * 24)) + 1;

  for (let row = 0; row < dayCount; row++) {
    const d = new Date(beginn.getTime() + row * 86400000);
    sheet.getRange(row + 2, 1).setValue(d).setNumberFormat('DD.MM.YYYY');
  }

  const lastRow = dayCount + 1;

  sheet.getRange(2, 1, lastRow - 1, 1).setNumberFormat('DD.MM.YYYY');

  sheet.getRange(2, 4, lastRow - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Heim', 'Auswärts'])
      .setAllowInvalid(true)
      .build()
  );

  for (let i = 0; i < maxSlots; i++) {
    sheet.getRange(2, saisonEinsatzartCol(i), lastRow - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Einzel+Doppel', 'Einzel', 'Doppel'])
        .setAllowInvalid(true)
        .build()
    );
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (spielerSheet) {
    const spielerLastRow = spielerSheet.getLastRow();
    if (spielerLastRow > 1) {
      const spielerRange = spielerSheet.getRange(2, COL_SPIELER.Name, spielerLastRow - 1, 1);
      for (let i = 0; i < maxSlots; i++) {
        sheet.getRange(2, saisonSpielerCol(i), lastRow - 1, 1).setDataValidation(
          SpreadsheetApp.newDataValidation()
            .requireValueInRange(spielerRange)
            .setAllowInvalid(true)
            .build()
        );
      }
    }
  }

  sheet.getRange(2, saisonStatusCol(), lastRow - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Geplant', 'Final'])
      .setAllowInvalid(true)
      .build()
  );

  for (let pi = 0; pi < config.spieler.length; pi++) {
    const playerCol = firstSpielerCol + pi;
    const spielerName = config.spieler[pi].name;
    const colLetter = colToLetter(playerCol);

    for (let row = 2; row <= lastRow; row++) {
      const col2Letter = colToLetter(COL_ABWESENHEITEN.Spieler);
      const vonLetter = colToLetter(COL_ABWESENHEITEN.Von);
      const bisLetter = colToLetter(COL_ABWESENHEITEN.Bis);
      const komLetter = colToLetter(COL_ABWESENHEITEN.Kommentar);

      const formula = `=LET(abw; FILTER(${col2Letter}:${col2Letter}; ${col2Letter}:${col2Letter}="${spielerName}"; ${vonLetter}:${vonLetter}<=A${row}; ${bisLetter}:${bisLetter}>=A${row}); IF(COUNTA(abw)>0; "✗ " & INDEX(abw;1;1); "✓"))`;

      sheet.getRange(row, playerCol).setFormula(formula);
    }

    const playerRange = sheet.getRange(2, playerCol, lastRow - 1, 1);
    sheet.setConditionalFormatRules([
      ...sheet.getConditionalFormatRules(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextStartsWith('✓')
        .setBackground('#D9EAD3')
        .setRanges([playerRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextStartsWith('✗')
        .setBackground('#F4CCCC')
        .setRanges([playerRange])
        .build(),
    ]);
  }

  const filterRange = sheet.getRange(1, 1, lastRow, numCols);
  if (filterRange.getFilter()) {
    filterRange.getFilter()!.remove();
  }
  filterRange.createFilter();
  const filterCol2 = 2;
  const criteria = SpreadsheetApp.newFilterCriteria()
    .whenCellNotEmpty()
    .build();
  sheet.getFilter()!.setColumnFilterCriteria(filterCol2, criteria);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  return sheet;
}

function colToLetter(col: number): string {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
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
