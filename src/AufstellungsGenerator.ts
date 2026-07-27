/// <reference path="ConfigTypes.ts" />

function generateAufstellungen(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  const abwesenheitenSheet = ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN);
  const termineSheet = ss.getSheetByName(SHEET_NAMES.SPIELTERMINE);
  const aufstellungenSheet = ss.getSheetByName(SHEET_NAMES.AUFSTELLUNGEN);

  if (!spielerSheet || !abwesenheitenSheet || !termineSheet || !aufstellungenSheet) {
    SpreadsheetApp.getUi().alert('Fehler', 'Nicht alle benötigten Sheets sind vorhanden.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const spieler = readAktiveSpieler(spielerSheet);
  const termine = readSpieltermine(termineSheet);

  if (spieler.length < SHEET_CONFIG.einstellungen.spielformat.einzel) {
    SpreadsheetApp.getUi().alert(
      'Warnung',
      `Nur ${spieler.length} aktive Spieler gefunden, aber ${SHEET_CONFIG.einstellungen.spielformat.einzel} werden benötigt.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }

  aufstellungenSheet.clearContents();
  formatGenHeader(aufstellungenSheet);

  const positionen = Array.from(
    { length: SHEET_CONFIG.einstellungen.spielformat.einzel },
    (_, i) => `Einzel ${i + 1}` as AufstellungsTyp
  );
  const doppelSlots = Array.from(
    { length: SHEET_CONFIG.einstellungen.spielformat.doppel },
    () => 'Doppel' as AufstellungsTyp
  );
  const allePositionen = [...positionen, ...doppelSlots];

  let currentRow = 2;

  for (const termin of termine) {
    const abwesende = readAbwesenheitenFuerDatum(abwesenheitenSheet, termin.datum);
    const abwesendeNamen = new Set(abwesende.map(a => a.spieler));

    const verfuegbare = spieler
      .filter(s => !abwesendeNamen.has(s.name))
      .sort((a, b) => a.rang - b.rang);

    const benoetigt = SHEET_CONFIG.einstellungen.spielformat.einzel;

    for (let i = 0; i < allePositionen.length; i++) {
      const pos = allePositionen[i];
      let spielerName = '';

      if (pos === 'Doppel') {
        const doppelIndex = i - positionen.length;
        if (verfuegbare.length >= benoetigt) {
          spielerName = verfuegbare[doppelIndex % verfuegbare.length].name;
        } else if (verfuegbare.length > doppelIndex) {
          spielerName = verfuegbare[doppelIndex].name;
        }
      } else {
        const einIndex = positionen.indexOf(pos);
        if (einIndex < verfuegbare.length) {
          spielerName = verfuegbare[einIndex].name;
        } else {
          spielerName = `⚠️ ZU WENIG SPIELER`;
        }
      }

      aufstellungenSheet.getRange(currentRow, COL_AUFSTELLUNGEN.Termin, 1, 3).setValues([[
        termin.datum,
        pos,
        spielerName,
      ]]);
      currentRow++;
    }

    if (abwesende.filter(a => a.kommentar.toLowerCase().includes('muss')).length > 0) {
      aufstellungenSheet.getRange(currentRow, COL_AUFSTELLUNGEN.Termin, 1, 3).setValues([[
        termin.datum,
        'ℹ️ Hinweis',
        'Einige Abwesenheiten sind mit "nur wenn es sein muss" markiert',
      ]]);
      currentRow++;
    }
  }
}

function formatGenHeader(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  const headers = ['Termin', 'Typ', 'Spieler'];
  sheet.getRange(1, COL_AUFSTELLUNGEN.Termin, 1, 3).setValues([headers]);
  const range = sheet.getRange(1, COL_AUFSTELLUNGEN.Termin, 1, 3);
  range.setFontWeight('bold');
  range.setBackground('#4A90D9');
  range.setFontColor('#FFFFFF');
  range.setHorizontalAlignment('center');
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

  for (const row of data) {
    const von = row[COL_ABWESENHEITEN.Von - 1] instanceof Date ? row[COL_ABWESENHEITEN.Von - 1] : new Date(row[COL_ABWESENHEITEN.Von - 1]);
    const bis = row[COL_ABWESENHEITEN.Bis - 1] instanceof Date ? row[COL_ABWESENHEITEN.Bis - 1] : new Date(row[COL_ABWESENHEITEN.Bis - 1]);

    if (isNaN(von.getTime()) || isNaN(bis.getTime())) {
      continue;
    }

    const d = new Date(datum.getFullYear(), datum.getMonth(), datum.getDate());

    if (d >= von && d <= bis) {
      abwesenheiten.push({
        spieler: String(row[COL_ABWESENHEITEN.Spieler - 1]),
        von,
        bis,
        kommentar: String(row[COL_ABWESENHEITEN.Kommentar - 1]),
      });
    }
  }

  return abwesenheiten;
}

function readSpieltermine(sheet: GoogleAppsScript.Spreadsheet.Sheet): Spieltermin[] {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  const data = sheet.getRange(2, 1, lastRow - 1, COL_SPIELTERMINE.Status).getValues();
  const termine: Spieltermin[] = [];

  for (const row of data) {
    const datum = row[COL_SPIELTERMINE.Datum - 1] instanceof Date ? row[COL_SPIELTERMINE.Datum - 1] : new Date(row[COL_SPIELTERMINE.Datum - 1]);
    if (isNaN(datum.getTime())) {
      continue;
    }

    const hg = String(row[COL_SPIELTERMINE.HeimGast - 1]);

    termine.push({
      datum,
      heimGast: (hg === 'Heim' || hg === 'Gast') ? hg : 'Heim',
      gegner: String(row[COL_SPIELTERMINE.Gegner - 1]),
      ort: String(row[COL_SPIELTERMINE.Ort - 1]),
      status: String(row[COL_SPIELTERMINE.Status - 1]) as Spieltermin['status'],
    });
  }

  return termine.sort((a, b) => a.datum.getTime() - b.datum.getTime());
}
