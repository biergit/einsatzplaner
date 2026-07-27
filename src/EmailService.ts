/// <reference path="ConfigTypes.ts" />

interface EinsatzInfo {
  spielterminDatum: string;
  gegner: string;
  heimGast: string;
  ort: string;
  typ: string;
}

function sendEinsatzEmails(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aufstellungenSheet = ss.getSheetByName(SHEET_NAMES.AUFSTELLUNGEN);
  const termineSheet = ss.getSheetByName(SHEET_NAMES.SPIELTERMINE);

  if (!aufstellungenSheet || !termineSheet) {
    return;
  }

  const termineMap = buildTerminMap(termineSheet);
  const einsaetze = groupEinsaetzeBySpieler(aufstellungenSheet, termineMap);
  const spielerMap = buildSpielerMap();

  const ohneEmail: string[] = [];

  for (const [spielerName, termine] of Object.entries(einsaetze)) {
    const spieler = spielerMap[spielerName];
    if (!spieler || !spieler.email || !spieler.email.includes('@')) {
      ohneEmail.push(spielerName);
      continue;
    }
    sendEinsatzplanEmail(spieler, termine);
  }

  if (ohneEmail.length > 0) {
    sendOhneEmailHinweis(ss, ohneEmail);
  }
}

function sendOhneEmailHinweis(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, namen: string[]): void {
  const kapitaenEmail = getKapitaenEmail(ss);
  if (!kapitaenEmail) {
    return;
  }

  const body = `Hallo,\n\n` +
    `Die Einsatzpläne wurden versendet. ` +
    `Folgende eingesetzte Spieler haben keine E-Mail-Adresse hinterlegt und sollten persönlich informiert werden:\n\n` +
    `${namen.map(n => `  – ${n}`).join('\n')}\n\n` +
    `Bitte kontaktiere sie direkt.\n\n` +
    `Viele Grüße,\nDein Einsatzplaner-Team`;

  MailApp.sendEmail({
    to: kapitaenEmail,
    subject: 'Einsatzplaner – Spieler ohne E-Mail-Adresse',
    body: body,
  });
}

function getKapitaenEmail(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): string {
  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (!spielerSheet) {
    return '';
  }

  const lastRow = spielerSheet.getLastRow();
  if (lastRow <= 1) {
    return '';
  }

  const data = spielerSheet.getRange(2, 1, lastRow - 1, COL_SPIELER.Rolle).getValues();

  for (const row of data) {
    const email = String(row[COL_SPIELER.Email - 1]).trim();
    const rolle = String(row[COL_SPIELER.Rolle - 1]).trim();
    if (email && email.includes('@') && rolle === 'Kapitän') {
      return email;
    }
  }

  return '';
}

function buildTerminMap(termineSheet: GoogleAppsScript.Spreadsheet.Sheet): Record<string, Spieltermin> {
  const map: Record<string, Spieltermin> = {};
  const lastRow = termineSheet.getLastRow();
  if (lastRow <= 1) {
    return map;
  }

  const data = termineSheet.getRange(2, 1, lastRow - 1, COL_SPIELTERMINE.Status).getValues();
  for (const row of data) {
    const datum = row[COL_SPIELTERMINE.Datum - 1];
    if (datum) {
      const datumStr = datum instanceof Date
        ? Utilities.formatDate(datum, 'Europe/Berlin', 'dd.MM.yyyy')
        : String(datum);
      const hg = String(row[COL_SPIELTERMINE.HeimGast - 1]);
      map[datumStr] = {
        datum: datum instanceof Date ? datum : new Date(datum),
        heimGast: (hg === 'Heim' || hg === 'Gast') ? hg : 'Heim',
        gegner: String(row[COL_SPIELTERMINE.Gegner - 1]),
        ort: String(row[COL_SPIELTERMINE.Ort - 1]),
        status: String(row[COL_SPIELTERMINE.Status - 1]) as Spieltermin['status'],
      };
    }
  }

  return map;
}

function groupEinsaetzeBySpieler(
  aufstellungenSheet: GoogleAppsScript.Spreadsheet.Sheet,
  termineMap: Record<string, Spieltermin>
): Record<string, EinsatzInfo[]> {
  const einsaetze: Record<string, EinsatzInfo[]> = {};
  const lastRow = aufstellungenSheet.getLastRow();
  if (lastRow <= 1) {
    return einsaetze;
  }

  const data = aufstellungenSheet.getRange(2, 1, lastRow - 1, COL_AUFSTELLUNGEN.Spieler).getValues();
  for (const row of data) {
    const termin = row[COL_AUFSTELLUNGEN.Termin - 1];
    const typ = String(row[COL_AUFSTELLUNGEN.Typ - 1]);
    const spieler = String(row[COL_AUFSTELLUNGEN.Spieler - 1]);

    if (!termin || !spieler || typ.startsWith('\u2139')) {
      continue;
    }

    const datumStr = termin instanceof Date
      ? Utilities.formatDate(termin, 'Europe/Berlin', 'dd.MM.yyyy')
      : String(termin);

    const terminInfo = termineMap[datumStr];

    if (!einsaetze[spieler]) {
      einsaetze[spieler] = [];
    }

    einsaetze[spieler].push({
      spielterminDatum: datumStr,
      gegner: terminInfo?.gegner ?? '?',
      heimGast: terminInfo?.heimGast ?? '?',
      ort: terminInfo?.ort ?? '?',
      typ: typ,
    });
  }

  return einsaetze;
}

function buildSpielerMap(): Record<string, Spieler> {
  const map: Record<string, Spieler> = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (!spielerSheet) {
    return map;
  }

  const lastRow = spielerSheet.getLastRow();
  if (lastRow <= 1) {
    return map;
  }

  const data = spielerSheet.getRange(2, 1, lastRow - 1, COL_SPIELER.Rolle).getValues();

  for (const row of data) {
    const name = String(row[COL_SPIELER.Name - 1]).trim();
    const email = String(row[COL_SPIELER.Email - 1]).trim();
    const rang = Number(row[COL_SPIELER.Rang - 1]) || 99;
    if (name) {
      map[name] = {
        name,
        email,
        rang,
        aenderungenMelden: row[COL_SPIELER.AenderungenMelden - 1] === true,
        rolle: '',
      };
    }
  }

  return map;
}

function sendEinsatzplanEmail(spieler: Spieler, einsaetze: EinsatzInfo[]): void {
  const termineStr = einsaetze
    .sort((a, b) => a.spielterminDatum.localeCompare(b.spielterminDatum))
    .map(e => {
      const hg = e.heimGast === 'Heim' ? 'Heimspiel' : 'Auswärtsspiel';
      const ort = e.ort ? ` – ${e.ort}` : '';
      return `  ${e.spielterminDatum} – ${hg} gegen ${e.gegner} (${e.typ})${ort}`;
    })
    .join('\n');

  const body = `Hallo ${spieler.name},\n\n` +
    `hier ist dein vorläufiger Einsatzplan für die Saison:\n\n` +
    `${termineStr}\n\n` +
    `Bitte gib Bescheid, falls du an einem der Termine doch nicht kannst.\n\n` +
    `Viele Grüße,\nDein Einsatzplaner-Team`;

  MailApp.sendEmail({
    to: spieler.email,
    subject: 'Dein Einsatzplan – Tischtennis',
    body: body,
  });
}
