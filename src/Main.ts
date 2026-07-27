/// <reference path="ConfigTypes.ts" />

function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Einsatzplaner')
    .addItem('Sheet neu aufbauen', 'menuSheetNeuAufbauen')
    .addSeparator()
    .addItem('Daten exportieren', 'menuDatenExportieren')
    .addItem('Aufstellungen generieren', 'menuAufstellungenGenerieren')
    .addSeparator()
    .addItem('Aufstellungen finalisieren', 'menuAufstellungenFinalisieren')
    .addItem('Emails an eingesetzte Spieler senden', 'menuEmailsSenden')
    .addToUi();
}

function menuSheetNeuAufbauen(): void {
  const ui = SpreadsheetApp.getUi();
  const antwort = ui.alert(
    'Sheet neu aufbauen',
    'Achtung: Alle bestehenden Daten in diesem Sheet werden gelöscht und durch die Struktur aus der Konfiguration ersetzt.\n\n' +
    'Bist du sicher?\n\n' +
    'Vorher solltest du einen Export der aktuellen Daten machen!',
    ui.ButtonSet.YES_NO
  );

  if (antwort !== ui.Button.YES) {
    return;
  }

  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SHEET_BUILDER_RUNNING', 'true');
    buildAllSheets(SHEET_CONFIG);
    props.setProperty('SHEET_BUILDER_RUNNING', 'false');
    ui.alert('Fertig', 'Das Sheet wurde erfolgreich neu aufgebaut.', ui.ButtonSet.OK);
  } catch (e) {
    PropertiesService.getScriptProperties().setProperty('SHEET_BUILDER_RUNNING', 'false');
    ui.alert('Fehler', `Beim Aufbau des Sheets ist ein Fehler aufgetreten:\n${e}`, ui.ButtonSet.OK);
  }
}

function menuDatenExportieren(): void {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = exportAllData();
    ui.alert(
      'Export erfolgreich',
      `Die Daten wurden exportiert.\n\nExport-ID: ${result.fileId}\n\nDatei: ${result.fileName}`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Fehler', `Beim Export ist ein Fehler aufgetreten:\n${e}`, ui.ButtonSet.OK);
  }
}

function menuAufstellungenGenerieren(): void {
  try {
    generateAufstellungen();
    SpreadsheetApp.getUi().alert('Fertig', 'Die Aufstellungen wurden generiert.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Fehler', `Beim Generieren ist ein Fehler aufgetreten:\n${e}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function menuAufstellungenFinalisieren(): void {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const termineSheet = ss.getSheetByName(SHEET_NAMES.SPIELTERMINE);

  if (!termineSheet) {
    ui.alert('Fehler', 'Sheet "Spieltermine" nicht gefunden.', ui.ButtonSet.OK);
    return;
  }

  const lastRow = termineSheet.getLastRow();
  if (lastRow <= 1) {
    ui.alert('Keine Termine', 'Es sind keine Spieltermine vorhanden.', ui.ButtonSet.OK);
    return;
  }

  for (let row = 2; row <= lastRow; row++) {
    const statusCell = termineSheet.getRange(row, COL_SPIELTERMINE.Status);
    const currentStatus = String(statusCell.getValue());

    if (currentStatus === 'Geplant') {
      statusCell.setValue('Finalisiert');
    }
  }

  ui.alert('Fertig', 'Alle geplanten Spieltermine wurden auf "Finalisiert" gesetzt.', ui.ButtonSet.OK);
}

function menuEmailsSenden(): void {
  try {
    sendEinsatzEmails();
    SpreadsheetApp.getUi().alert('Fertig', 'Die Einsatzpläne wurden per Email versendet.', SpreadsheetApp.getUi().ButtonSet.OK);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const termineSheet = ss.getSheetByName(SHEET_NAMES.SPIELTERMINE);
    if (termineSheet) {
      const lastRow = termineSheet.getLastRow();
      for (let row = 2; row <= lastRow; row++) {
        const statusCell = termineSheet.getRange(row, COL_SPIELTERMINE.Status);
        if (String(statusCell.getValue()) === 'Finalisiert') {
          statusCell.setValue('Versendet');
        }
      }
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('Fehler', `Beim Versenden ist ein Fehler aufgetreten:\n${e}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
