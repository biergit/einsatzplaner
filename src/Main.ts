/// <reference path="ConfigTypes.ts" />

function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Einsatzplaner')
    .addItem('Aufstellungen generieren', 'menuAufstellungenGenerieren')
    .addSeparator()
    .addItem('Finalisieren + Emails senden', 'menuFinalisierenUndSenden')
    .addSeparator()
    .addItem('Daten exportieren', 'menuDatenExportieren')
    .addSeparator()
    .addSubMenu(ui.createMenu('Danger Zone')
      .addItem('Sheet neu aufbauen', 'menuSheetNeuAufbauen'))
    .addToUi();
}

function menuSheetNeuAufbauen(): void {
  const ui = SpreadsheetApp.getUi();
  const antwort = ui.alert(
    'Sheet neu aufbauen',
    'Achtung: Alle bestehenden Daten werden gelöscht!\n\nBist du sicher?',
    ui.ButtonSet.YES_NO
  );

  if (antwort !== ui.Button.YES) return;

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
    exportAllData();
    ui.alert(
      'Export erfolgreich',
      'Die Rohdaten wurden als E-Mail-Anhang verschickt.',
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

function menuFinalisierenUndSenden(): void {
  const ui = SpreadsheetApp.getUi();
  const antwort = ui.alert(
    'Finalisieren + Emails senden',
    'Alle Aufstellungen mit Status "Geplant" werden auf "Final" gesetzt und Einsatz-Mails an die Spieler versendet.\n\nFortfahren?',
    ui.ButtonSet.YES_NO
  );
  if (antwort !== ui.Button.YES) return;

  PropertiesService.getScriptProperties().setProperty('SUPPRESS_NOTIFICATION', 'true');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const saisonSheet = ss.getSheetByName(SHEET_NAMES.SAISON);
  if (!saisonSheet) {
    ui.alert('Fehler', 'Saison-Sheet nicht gefunden.', ui.ButtonSet.OK);
    return;
  }

  const lastRow = saisonSheet.getLastRow();
  let count = 0;
  for (let row = 2; row <= lastRow; row++) {
    const statusCell = saisonSheet.getRange(row, saisonStatusCol());
    if (String(statusCell.getValue()).trim() === 'Geplant') {
      const gegner = String(saisonSheet.getRange(row, 2).getValue() || '').trim();
      if (gegner) {
        statusCell.setValue('Final');
        count++;
      }
    }
  }

  try {
    sendEinsatzEmails();
    ui.alert('Fertig', `${count} Spieltermine finalisiert. Einsatz-Mails wurden versendet.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Fehler', `Finalisierung ok (${count} Termine), aber E-Mail-Versand fehlgeschlagen:\n${e}`, ui.ButtonSet.OK);
  }
}

function autorisiere(): void {
  ScriptApp.getProjectTriggers();
}
