/// <reference path="ConfigTypes.ts" />

interface EinsatzInfo {
  datum: string;
  gegner: string;
  startzeit: string;
  heimAuswaerts: string;
  einsatzart: string;
  notTop4: boolean;
}

function sendEinsatzEmails(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const saisonSheet = ss.getSheetByName(SHEET_NAMES.SAISON);
  if (!saisonSheet) return;

  const lastRow = saisonSheet.getLastRow();
  const heute = new Date();
  heute.setHours(0, 0, 0, 0);

  const spielerNames = readSpielerNames(ss);
  const aktiveSpieler = readAktiveSpieler(ss.getSheetByName(SHEET_NAMES.SPIELER)!);
  const allAbw = buildAbwesenheitenIndex(ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN)!);
  const dates = readSaisonDates(saisonSheet, lastRow);

  const numPlayers = spielerNames.length;
  const allPlayerCols = saisonSheet.getRange(2, saisonSpielerCol(0), lastRow - 1, numPlayers).getValues() as string[][];

  const einsaetzeProSpieler: Record<string, EinsatzInfo[]> = {};
  const ohneEmail: string[] = [];

  const spielplan: string[] = [];

  for (let r = 0; r < dates.length; r++) {
    const datum = dates[r];
    if (!datum || datum < heute) continue;

    const row = r + 2;
    const status = String(saisonSheet.getRange(row, saisonStatusCol()).getValue() || '').trim();
    if (status !== 'Final') continue;

    const gegner = String(saisonSheet.getRange(row, saisonGegnerCol()).getValue() || '').trim();
    if (!gegner) continue;

    const startzeit = String(saisonSheet.getRange(row, saisonStartzeitCol()).getValue() || '').trim();
    const heimAuswaerts = String(saisonSheet.getRange(row, saisonHeimAuswaertsCol()).getValue() || '').trim();
    const datumStr = Utilities.formatDate(datum, 'Europe/Berlin', 'dd.MM.yyyy');

    const key = dateKey(datum);
    const dayMap = allAbw.get(key);
    const available = aktiveSpieler
      .filter(s => {
        const abw = dayMap?.get(s.name);
        return !abw || !abw.startsWith('✗');
      })
      .sort((a, b) => a.rang - b.rang);
    const top4Names = new Set(available.slice(0, 4).map(s => s.name));

    const z = startzeit ? ` – ${startzeit}` : '';
    spielplan.push(`${datumStr}${z}   ${heimAuswaerts}   ${gegner}`);

    for (let pi = 0; pi < numPlayers; pi++) {
      const val = String(allPlayerCols[r][pi] || '').trim();
      if (!val || val.startsWith('✗')) continue;

      const name = spielerNames[pi];
      const notTop4 = !top4Names.has(name);
      const marker = notTop4 ? ' ★' : '';
      spielplan.push(`  ${name}  →  ${val}${marker}`);

      if (!einsaetzeProSpieler[name]) einsaetzeProSpieler[name] = [];
      einsaetzeProSpieler[name].push({ datum: datumStr, gegner, startzeit, heimAuswaerts, einsatzart: val, notTop4 });
    }

    for (let ei = 0; ei < 3; ei++) {
      const eName = String(saisonSheet.getRange(row, saisonErsatzCol(ei)).getValue() || '').trim();
      if (!eName) continue;
      spielplan.push(`  ${eName}  →  Einzel+Doppel (Ersatz)`);
      if (!einsaetzeProSpieler[eName]) einsaetzeProSpieler[eName] = [];
      einsaetzeProSpieler[eName].push({ datum: datumStr, gegner, startzeit, heimAuswaerts, einsatzart: 'Einzel+Doppel', notTop4: false });
    }

    spielplan.push('');
  }

  const gesamtspielplan = spielplan.join('\n');

  const spielerMap = buildSpielerMap();

  for (const [spielerName, einsaetze] of Object.entries(einsaetzeProSpieler)) {
    const spieler = spielerMap[spielerName];
    if (!spieler || !spieler.email || !spieler.email.includes('@')) {
      ohneEmail.push(spielerName);
      continue;
    }
    sendEinsatzplanEmail(spieler, einsaetze, gesamtspielplan);
  }

  if (ohneEmail.length > 0) {
    sendOhneEmailHinweis(ss, ohneEmail);
  }
}

function readSpielerNames(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): string[] {
  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (!spielerSheet) return [];
  const lastRow = spielerSheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = spielerSheet.getRange(2, COL_SPIELER.Name, lastRow - 1, 1).getValues();
  return data.map((r: unknown[]) => String(r[0]).trim()).filter((n: string) => n);
}

function sendOhneEmailHinweis(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, namen: string[]): void {
  const kapitaenEmail = getKapitaenEmail(ss);
  if (!kapitaenEmail) return;

  MailApp.sendEmail({
    to: kapitaenEmail,
    subject: 'Einsatzplaner – Spieler ohne E-Mail-Adresse',
    body: `Hallo,\n\nDie Einsatzpläne wurden versendet. Folgende eingesetzte Spieler haben keine E-Mail-Adresse:\n\n${namen.map(n => `  – ${n}`).join('\n')}\n\nBitte kontaktiere sie direkt.\n\nDein Einsatzplaner-Team`,
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
    if (email && email.includes('@') && String(row[COL_SPIELER.Rolle - 1]).trim() === 'Kapitän') return email;
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
    if (name) {
      map[name] = {
        name, email, rang: Number(row[COL_SPIELER.Rang - 1]) || 99,
        aenderungenMelden: row[COL_SPIELER.AenderungenMelden - 1] === true,
        rolle: '',
      };
    }
  }
  return map;
}

function sendEinsatzplanEmail(spieler: Spieler, einsaetze: EinsatzInfo[], gesamtspielplan: string): void {
  const termineStr = einsaetze
    .sort((a, b) => a.datum.localeCompare(b.datum))
    .map(e => {
      const h = e.heimAuswaerts || '';
      const z = e.startzeit ? ` – ${e.startzeit}` : '';
      const marker = e.notTop4 ? ' ★' : '';
      return `  ${e.datum}${z} – ${h} gegen ${e.gegner} (${e.einsatzart})${marker}`;
    })
    .join('\n');

  const body = `Hallo ${spieler.name},\n\n` +
    `hier ist dein Einsatzplan (★ = außerhalb der Top 4):\n\n${termineStr}\n\n` +
    `─── GESAMTSPIELPLAN ───\n\n${gesamtspielplan}\n\n` +
    `Viele Grüße,\nDein Einsatzplaner-Team`;

  MailApp.sendEmail({
    to: spieler.email,
    subject: 'Dein Einsatzplan – Tischtennis',
    body: body,
  });
}
