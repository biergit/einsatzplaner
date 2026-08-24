/// <reference path="ConfigTypes.ts" />

function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Einsatzplaner')
    .addItem('Aufstellungen generieren', 'menuAufstellungenGenerieren')
    .addSeparator()
    .addItem('Finalisieren + Emails senden', 'menuFinalisierenUndSenden')
    .addSeparator()
    .addItem('Spieltag-Filter setzen', 'menuSpieltagFilterSetzen')
    .addItem('Spieltag-Filter entfernen', 'menuSpieltagFilterEntfernen')
    .addSeparator()
    .addItem('Daten exportieren', 'menuDatenExportieren')
    .addSeparator()
    .addItem('Autorisierung prüfen', 'menuAutorisierungPruefen')
    .addSeparator()
    .addSubMenu(ui.createMenu('Danger Zone')
      .addItem('Sheet neu aufbauen', 'menuSheetNeuAufbauen'))
    .addToUi();

  saveSheetSnapshots(SpreadsheetApp.getActiveSpreadsheet());
}

function menuSpieltagFilterSetzen(): void {
  const ui = SpreadsheetApp.getUi();
  const saisonSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SAISON);
  if (!saisonSheet) {
    ui.alert('Fehler', 'Saison-Sheet nicht gefunden.', ui.ButtonSet.OK);
    return;
  }
  applySpieltagFilter(saisonSheet, SHEET_CONFIG);
  ui.alert('Fertig', 'Der Spieltag-Filter wurde gesetzt. Es werden nur noch Spieltage angezeigt.', ui.ButtonSet.OK);
}

function menuSpieltagFilterEntfernen(): void {
  const ui = SpreadsheetApp.getUi();
  const saisonSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SAISON);
  if (!saisonSheet) {
    ui.alert('Fehler', 'Saison-Sheet nicht gefunden.', ui.ButtonSet.OK);
    return;
  }
  removeSpieltagFilter(saisonSheet);
  ui.alert('Fertig', 'Der Spieltag-Filter wurde entfernt. Alle Tage sind wieder sichtbar.', ui.ButtonSet.OK);
}

function menuAutorisierungPruefen(): void {
  const ui = SpreadsheetApp.getUi();
  try {
    const email = Session.getActiveUser().getEmail();
    autorisiere();
    ui.alert(
      'Autorisierung OK',
      `Angemeldet als: ${email || 'Unbekannt'}\n\nDer installierbare onEdit-Trigger ist eingerichtet.\n\nHinweis: Die "Sicherheitswarnung" beim ersten Klick nach einem Deployment ist die normale Google-Autorisierung – einmal bestätigen, danach läuft alles ohne weitere Nachfrage.`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Fehler', `Autorisierung fehlgeschlagen:\n${e}`, ui.ButtonSet.OK);
  }
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
    props.deleteProperty('SUPPRESS_NOTIFICATION');
    buildAllSheets(SHEET_CONFIG);
    SpreadsheetApp.getActiveSpreadsheet().rename(SHEET_CONFIG.einstellungen.sheetTitel);
    saveSheetSnapshots(SpreadsheetApp.getActiveSpreadsheet());
    cleanupTriggers();
    ensureOnEditTrigger();
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
      'Die Rohdaten wurden als E-Mail verschickt. Die Tabellen können per Copy & Paste direkt in Google Sheets oder eine .tsv-Datei eingefügt werden.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Fehler', `Beim Export ist ein Fehler aufgetreten:\n${e}`, ui.ButtonSet.OK);
  }
}

function menuAufstellungenGenerieren(): void {
  try {
    const warnings = generateAufstellungen();
    const zusatz = warnings.length > 0 ? `\n\nHinweise:\n- ${warnings.join('\n- ')}` : '';
    SpreadsheetApp.getUi().alert('Fertig', `Die Aufstellungen wurden generiert.${zusatz}`, SpreadsheetApp.getUi().ButtonSet.OK);
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
  PropertiesService.getScriptProperties().setProperty('BULK_EDIT', 'true');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const saisonSheet = ss.getSheetByName(SHEET_NAMES.SAISON);
  if (!saisonSheet) {
    PropertiesService.getScriptProperties().deleteProperty('SUPPRESS_NOTIFICATION');
    PropertiesService.getScriptProperties().deleteProperty('BULK_EDIT');
    ui.alert('Fehler', 'Saison-Sheet nicht gefunden.', ui.ButtonSet.OK);
    return;
  }

  const lastRow = saisonSheet.getLastRow();
  let count = 0;
  for (let row = 2; row <= lastRow; row++) {
    const statusCell = saisonSheet.getRange(row, saisonStatusCol());
    if (String(statusCell.getValue()).trim() === 'Geplant') {
      const gegner = String(saisonSheet.getRange(row, saisonGegnerCol()).getValue() || '').trim();
      if (gegner) {
        statusCell.setValue('Final');
        count++;
      }
    }
  }

  PropertiesService.getScriptProperties().deleteProperty('BULK_EDIT');
  PropertiesService.getScriptProperties().deleteProperty('SUPPRESS_NOTIFICATION');
  resetDebounceTimer();

  try {
    sendEinsatzEmails();
    ui.alert('Fertig', `${count} Spieltermine finalisiert. Einsatz-Mails wurden versendet.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Fehler', `Finalisierung ok (${count} Termine), aber E-Mail-Versand fehlgeschlagen:\n${e}`, ui.ButtonSet.OK);
  }
}

function autorisiere(): void {
  ScriptApp.getProjectTriggers();
  cleanupTriggers();
  ensureOnEditTrigger();
}
