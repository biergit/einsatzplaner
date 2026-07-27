/// <reference path="ConfigTypes.ts" />

function generateAufstellungen(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  const abwesenheitenSheet = ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN);
  const saisonSheet = ss.getSheetByName(SHEET_NAMES.SAISON);

  if (!spielerSheet || !abwesenheitenSheet || !saisonSheet) {
    SpreadsheetApp.getUi().alert('Fehler', 'Nicht alle benötigten Sheets sind vorhanden.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const aktiveSpieler = readAktiveSpieler(spielerSheet);
  const lastRow = saisonSheet.getLastRow();
  const maxSlots = saisonMaxSlots();

  for (let row = 2; row <= lastRow; row++) {
    const gegner = String(saisonSheet.getRange(row, 2).getValue() || '').trim();
    if (!gegner) continue;

    const datum = saisonSheet.getRange(row, 1).getValue();
    if (!(datum instanceof Date) || isNaN(datum.getTime())) continue;

    const abwesendeNamen = new Set(
      readAbwesenheitenFuerDatum(abwesenheitenSheet, datum)
        .filter(a => !a.kommentar.toLowerCase().includes('muss'))
        .map(a => a.spieler)
    );

    const verfuegbare = aktiveSpieler
      .filter(s => !abwesendeNamen.has(s.name))
      .sort((a, b) => a.rang - b.rang);

    const einzahl = SHEET_CONFIG.einstellungen.spielformat.einzel;

    for (let slot = 0; slot < maxSlots; slot++) {
      const einsatzartCell = saisonSheet.getRange(row, saisonEinsatzartCol(slot));
      const spielerCell = saisonSheet.getRange(row, saisonSpielerCol(slot));

      const existingSpieler = String(spielerCell.getValue() || '').trim();
      const existingEinsatzart = String(einsatzartCell.getValue() || '').trim();

      if (existingSpieler && existingEinsatzart) continue;

      if (slot < einzahl) {
        if (!existingEinsatzart) einsatzartCell.setValue('Einzel+Doppel');
        if (!existingSpieler && slot < verfuegbare.length) {
          spielerCell.setValue(verfuegbare[slot].name);
        }
      } else {
        if (!existingEinsatzart) einsatzartCell.setValue('Doppel');
        if (!existingSpieler && slot - einzahl < verfuegbare.length) {
          spielerCell.setValue(verfuegbare[slot].name);
        }
      }
    }

    const statusCell = saisonSheet.getRange(row, saisonStatusCol());
    if (!String(statusCell.getValue()).trim()) {
      statusCell.setValue('Geplant');
    }
  }

  updateHinweisSpalte(saisonSheet, spielerSheet, abwesenheitenSheet, aktiveSpieler);
}

function updateHinweisSpalte(
  saisonSheet: GoogleAppsScript.Spreadsheet.Sheet,
  spielerSheet: GoogleAppsScript.Spreadsheet.Sheet,
  abwesenheitenSheet: GoogleAppsScript.Spreadsheet.Sheet,
  aktiveSpieler: Spieler[]
): void {
  const lastRow = saisonSheet.getLastRow();
  const maxSlots = saisonMaxSlots();

  for (let row = 2; row <= lastRow; row++) {
    const gegner = String(saisonSheet.getRange(row, 2).getValue() || '').trim();
    if (!gegner) {
      saisonSheet.getRange(row, saisonHinweisCol()).setValue('');
      continue;
    }

    const datum = saisonSheet.getRange(row, 1).getValue();
    if (!(datum instanceof Date) || isNaN(datum.getTime())) continue;

    const abwAlle = readAbwesenheitenFuerDatum(abwesenheitenSheet, datum);
    const abwMap = new Map<string, Abwesenheit>();
    for (const a of abwAlle) {
      abwMap.set(a.spieler, a);
    }

    const hinweise: string[] = [];

    for (let slot = 0; slot < maxSlots; slot++) {
      const spielerName = String(saisonSheet.getRange(row, saisonSpielerCol(slot)).getValue() || '').trim();
      if (!spielerName) continue;

      const sp = aktiveSpieler.find(s => s.name === spielerName);
      if (sp && !sp.aenderungenMelden) continue;

      const abw = abwMap.get(spielerName);
      if (abw) {
        hinweise.push(`${spielerName}: abwesend (${abw.kommentar || 'kein Kommentar'})`);
      } else if (!aktiveSpieler.some(s => s.name === spielerName) && spielerName.includes(' ')) {
        hinweise.push(`${spielerName}: nicht im Spieler-Sheet`);
      }
    }

    saisonSheet.getRange(row, saisonHinweisCol()).setValue(hinweise.join(' | '));
  }
}

function readAktiveSpieler(sheet: GoogleAppsScript.Spreadsheet.Sheet): Spieler[] {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  const data = sheet.getRange(2, 1, lastRow - 1, COL_SPIELER.Rolle).getValues();
  const spieler: Spieler[] = [];

  for (const row of data) {
    if (row[COL_SPIELER.Aktiv - 1] === true || row[COL_SPIELER.Aktiv - 1] === 'TRUE') {
      const name = String(row[COL_SPIELER.Name - 1]).trim();
      if (!name) continue;

      let rang = Number(row[COL_SPIELER.Rang - 1]);
      if (isNaN(rang) || rang <= 0) rang = 99;

      spieler.push({
        name,
        email: String(row[COL_SPIELER.Email - 1]),
        rang,
        aenderungenMelden: row[COL_SPIELER.AenderungenMelden - 1] === true,
        rolle: '',
      });
    }
  }

  return spieler;
}

function readAbwesenheitenFuerDatum(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  datum: Date
): Abwesenheit[] {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  const data = sheet.getRange(2, 1, lastRow - 1, COL_ABWESENHEITEN.Kommentar).getValues();
  const abwesenheiten: Abwesenheit[] = [];
  const d = new Date(datum.getFullYear(), datum.getMonth(), datum.getDate());

  for (const row of data) {
    const von = row[COL_ABWESENHEITEN.Von - 1] instanceof Date ? row[COL_ABWESENHEITEN.Von - 1] : new Date(row[COL_ABWESENHEITEN.Von - 1]);
    const bis = row[COL_ABWESENHEITEN.Bis - 1] instanceof Date ? row[COL_ABWESENHEITEN.Bis - 1] : new Date(row[COL_ABWESENHEITEN.Bis - 1]);

    if (isNaN(von.getTime()) || isNaN(bis.getTime())) continue;
    if (d < von || d > bis) continue;

    abwesenheiten.push({
      spieler: String(row[COL_ABWESENHEITEN.Spieler - 1]),
      von,
      bis,
      kommentar: String(row[COL_ABWESENHEITEN.Kommentar - 1]),
    });
  }

  return abwesenheiten;
}
