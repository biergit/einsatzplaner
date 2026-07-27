/// <reference path="ConfigTypes.ts" />

interface EinsatzInfo {
  datum: string;
  gegner: string;
  startzeit: string;
  heimAuswaerts: string;
  einsatzart: string;
}

function sendEinsatzEmails(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const saisonSheet = ss.getSheetByName(SHEET_NAMES.SAISON);
  if (!saisonSheet) return;

  const lastRow = saisonSheet.getLastRow();
  const maxSlots = saisonMaxSlots();
  const heute = new Date();
  heute.setHours(0, 0, 0, 0);

  const einsaetzeProSpieler: Record<string, EinsatzInfo[]> = {};
  const alleAufgestellten = new Set<string>();

  for (let row = 2; row <= lastRow; row++) {
    const status = String(saisonSheet.getRange(row, saisonStatusCol()).getValue() || '').trim();
    if (status !== 'Final') continue;

    const datum = saisonSheet.getRange(row, 1).getValue();
    if (!(datum instanceof Date)) continue;
    if (datum < heute) continue;

    const gegner = String(saisonSheet.getRange(row, 2).getValue() || '').trim();
    const startzeit = String(saisonSheet.getRange(row, 3).getValue() || '').trim();
    const heimAuswaerts = String(saisonSheet.getRange(row, 4).getValue() || '').trim();
    const datumStr = Utilities.formatDate(datum, 'Europe/Berlin', 'dd.MM.yyyy');

    for (let slot = 0; slot < maxSlots; slot++) {
      const spielerName = String(saisonSheet.getRange(row, saisonSpielerCol(slot)).getValue() || '').trim();
      const einsatzart = String(saisonSheet.getRange(row, saisonEinsatzartCol(slot)).getValue() || '').trim();
      if (!spielerName || !einsatzart) continue;

      alleAufgestellten.add(spielerName);

      if (!einsaetzeProSpieler[spielerName]) {
        einsaetzeProSpieler[spielerName] = [];
      }
      einsaetzeProSpieler[spielerName].push({ datum: datumStr, gegner, startzeit, heimAuswaerts, einsatzart });
    }
  }

  const spielerMap = buildSpielerMap();
  const ohneEmail: string[] = [];

  for (const [spielerName, einsaetze] of Object.entries(einsaetzeProSpieler)) {
    const spieler = spielerMap[spielerName];
    if (!spieler || !spieler.email || !spieler.email.includes('@')) {
      ohneEmail.push(spielerName);
      continue;
    }
    sendEinsatzplanEmail(spieler, einsaetze);
  }

  if (ohneEmail.length > 0) {
    sendOhneEmailHinweis(ss, ohneEmail);
  }
}

function sendOhneEmailHinweis(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, namen: string[]): void {
  const kapitaenEmail = getKapitaenEmail(ss);
  if (!kapitaenEmail) return;

  const body = `Hallo,\n\n` +
    `Die Einsatzpläne wurden versendet. ` +
    `Folgende eingesetzte Spieler haben keine E-Mail-Adresse hinterlegt:\n\n` +
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
  if (!spielerSheet) return '';

  const lastRow = spielerSheet.getLastRow();
  if (lastRow <= 1) return '';

  const data = spielerSheet.getRange(2, 1, lastRow - 1, COL_SPIELER.Rolle).getValues();
  for (const row of data) {
    const email = String(row[COL_SPIELER.Email - 1]).trim();
    const rolle = String(row[COL_SPIELER.Rolle - 1]).trim();
    if (email && email.includes('@') && rolle === 'Kapitän') return email;
  }
  return '';
}

function buildSpielerMap(): Record<string, Spieler> {
  const map: Record<string, Spieler> = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (!spielerSheet) return map;

  const lastRow = spielerSheet.getLastRow();
  if (lastRow <= 1) return map;

  const data = spielerSheet.getRange(2, 1, lastRow - 1, COL_SPIELER.Rolle).getValues();
  for (const row of data) {
    const name = String(row[COL_SPIELER.Name - 1]).trim();
    const email = String(row[COL_SPIELER.Email - 1]).trim();
    const rang = Number(row[COL_SPIELER.Rang - 1]) || 99;
    if (name) {
      map[name] = {
        name, email, rang,
        aenderungenMelden: row[COL_SPIELER.AenderungenMelden - 1] === true,
        rolle: '',
      };
    }
  }
  return map;
}

function sendEinsatzplanEmail(spieler: Spieler, einsaetze: EinsatzInfo[]): void {
  const termineStr = einsaetze
    .sort((a, b) => a.datum.localeCompare(b.datum))
    .map(e => {
      const h = e.heimAuswaerts || '';
      const z = e.startzeit ? ` – ${e.startzeit}` : '';
      return `  ${e.datum}${z} – ${h} gegen ${e.gegner} (${e.einsatzart})`;
    })
    .join('\n');

  const body = `Hallo ${spieler.name},\n\n` +
    `hier ist dein Einsatzplan:\n\n${termineStr}\n\n` +
    `Viele Grüße,\nDein Einsatzplaner-Team`;

  MailApp.sendEmail({
    to: spieler.email,
    subject: 'Dein Einsatzplan – Tischtennis',
    body: body,
  });
}
