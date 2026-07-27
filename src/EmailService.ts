/// <reference path="ConfigTypes.ts" />

interface EinsatzInfo {
  datum: string;
  gegner: string;
  startzeit: string;
  heimAuswaerts: string;
  einsatzart: string;
  besondereUnterstuetzung: boolean;
}

function sendEinsatzEmails(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const saisonSheet = ss.getSheetByName(SHEET_NAMES.SAISON);
  if (!saisonSheet) return;

  const lastRow = saisonSheet.getLastRow();
  const heute = new Date(); heute.setHours(0, 0, 0, 0);

  const spielerNames = readSpielerNames(ss);
  const aktiveSpieler = readAktiveSpieler(ss.getSheetByName(SHEET_NAMES.SPIELER)!);
  const allAbw = buildAbwesenheitenIndex(ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN)!);
  const dates = readSaisonDates(saisonSheet, lastRow);

  const numPlayers = spielerNames.length;
  const allPlayerCols = saisonSheet.getRange(2, saisonSpielerCol(0), lastRow - 1, numPlayers).getValues() as string[][];

  const einsaetzeProSpieler: Record<string, EinsatzInfo[]> = {};
  const ohneEmail: { name: string; einsaetze: EinsatzInfo[] }[] = [];

  const spielplanRows: string[][] = [];

  for (let r = 0; r < dates.length; r++) {
    const datum = dates[r];
    if (!datum || datum < heute) continue;
    const row = r + 2;

    const status = String(saisonSheet.getRange(row, saisonStatusCol()).getValue() || '').trim();
    if (status !== 'Final') continue;
    const gegner = String(saisonSheet.getRange(row, saisonGegnerCol()).getValue() || '').trim();
    if (!gegner) continue;

    const startVal = saisonSheet.getRange(row, saisonStartzeitCol()).getValue();
    const startzeit = formatStartzeit(startVal);

    const heimAuswaerts = String(saisonSheet.getRange(row, saisonHeimAuswaertsCol()).getValue() || '').trim();
    const datumStr = Utilities.formatDate(datum, 'Europe/Berlin', 'dd.MM.yyyy');

    const key = dateKey(datum);
    const dayMap = allAbw.get(key);
    const available = aktiveSpieler.filter(s => {
      const abw = dayMap?.get(s.name);
      return !abw || !abw.startsWith('✗');
    }).sort((a, b) => a.rang - b.rang);
    const top4Names = new Set(available.slice(0, 4).map(s => s.name));

    const matchPlayers: string[] = [];
    for (let pi = 0; pi < numPlayers; pi++) {
      const val = String(allPlayerCols[r][pi] || '').trim();
      if (!val || val.startsWith('✗')) continue;
      const name = spielerNames[pi];
      const bu = !top4Names.has(name);
      matchPlayers.push(`${name} → ${val}${bu ? ' ★' : ''}`);
      if (!einsaetzeProSpieler[name]) einsaetzeProSpieler[name] = [];
      einsaetzeProSpieler[name].push({ datum: datumStr, gegner, startzeit, heimAuswaerts, einsatzart: val, besondereUnterstuetzung: bu });
    }

    for (let ei = 0; ei < 3; ei++) {
      const eName = String(saisonSheet.getRange(row, saisonErsatzCol(ei)).getValue() || '').trim();
      if (!eName) continue;
      matchPlayers.push(`${eName} → Einzel+Doppel (Ersatz)`);
      if (!einsaetzeProSpieler[eName]) einsaetzeProSpieler[eName] = [];
      einsaetzeProSpieler[eName].push({ datum: datumStr, gegner, startzeit, heimAuswaerts, einsatzart: 'Einzel+Doppel', besondereUnterstuetzung: false });
    }

    spielplanRows.push([datumStr, startzeit, heimAuswaerts, gegner, matchPlayers.join('<br>')]);
  }

  const spielerMap = buildSpielerMap();

  for (const [spielerName, einsaetze] of Object.entries(einsaetzeProSpieler)) {
    const spieler = spielerMap[spielerName];
    if (!spieler || !spieler.email || !spieler.email.includes('@')) {
      ohneEmail.push({ name: spielerName, einsaetze });
      continue;
    }
    sendEinsatzplanEmail(spieler, einsaetze, spielplanRows);
  }

  if (ohneEmail.length > 0) sendOhneEmailHinweis(ss, ohneEmail);
}

function formatStartzeit(val: unknown): string {
  if (val instanceof Date) {
    const h = val.getHours();
    const m = val.getMinutes();
    if (h === 0 && m === 0 && val.getFullYear() === 1899) return '';
    return `${pad2(h)}:${pad2(m)}`;
  }
  return '';
}

function buildHtmlTable(rows: string[][], hasBu: boolean): string {
  if (rows.length === 0) return '';
  let html = '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:14px;width:100%">';
  html += '<tr style="background:#4A90D9;color:white"><th>Datum</th><th>Zeit</th><th>Ort</th><th>Gegner</th><th>Aufstellung</th></tr>';
  for (const r of rows) {
    const aufstellung = r[4]
      .replace(/ ★/g, ' <span style="color:#B8860B;font-size:12px">★</span>')
      .replace(/→/g, '—');
    html += `<tr><td>${r[0]}</td><td>${r[1]}</td><td>${r[2]}</td><td>${r[3]}</td><td style="font-size:14px">${aufstellung}</td></tr>`;
  }
  html += '</table>';
  if (hasBu) {
    html += '<p style="font-size:12px;color:#888">★ = besondere Unterstützung</p>';
  }
  return html;
}

function readSpielerNames(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): string[] {
  const s = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (!s) return [];
  const lr = s.getLastRow();
  if (lr <= 1) return [];
  return s.getRange(2, COL_SPIELER.Name, lr - 1, 1).getValues().map((r: unknown[]) => String(r[0]).trim()).filter((n: string) => n);
}

function sendOhneEmailHinweis(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, ohneEmail: { name: string; einsaetze: EinsatzInfo[] }[]): void {
  const kap = getKapitaenEmail(ss);
  if (!kap) return;

  let body = 'Hallo,\n\n';
  body += 'Es wurden soeben Spielpläne finalisiert und Einsatz-Mails verschickt.\n';
  body += 'Folgende eingesetzte Spieler haben keine E-Mail-Adresse und konnten nicht automatisch benachrichtigt werden:\n\n';

  for (const { name, einsaetze } of ohneEmail) {
    body += `${name}:\n`;
    for (const e of einsaetze.sort((a, b) => a.datum.localeCompare(b.datum))) {
      const z = e.startzeit ? ` ${e.startzeit}` : '';
      body += `  – ${e.datum}${z} (${e.heimAuswaerts ? e.heimAuswaerts + ' – ' : ''}${e.gegner})\n`;
    }
    body += '\n';
  }

  body += 'Bitte informiere sie persönlich über ihre Einsätze.\n\n';
  body += 'Dein Einsatzplaner-Team';

  MailApp.sendEmail({ to: kap, subject: 'Einsatzplaner – Spieler ohne E-Mail-Adresse', body });
}

function getKapitaenEmail(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): string {
  const s = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (!s) return '';
  const lr = s.getLastRow();
  if (lr <= 1) return '';
  const d = s.getRange(2, 1, lr - 1, COL_SPIELER.Rolle).getValues();
  for (const r of d) {
    if (String(r[COL_SPIELER.Email - 1]).includes('@') && String(r[COL_SPIELER.Rolle - 1]).trim() === 'Kapitän')
      return String(r[COL_SPIELER.Email - 1]).trim();
  }
  return '';
}

function buildSpielerMap(): Record<string, Spieler> {
  const m: Record<string, Spieler> = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (!s) return m;
  const lr = s.getLastRow();
  if (lr <= 1) return m;
  const d = s.getRange(2, 1, lr - 1, COL_SPIELER.Rolle).getValues();
  for (const r of d) {
    const name = String(r[COL_SPIELER.Name - 1]).trim();
    if (name) m[name] = { name, email: String(r[COL_SPIELER.Email - 1]).trim(),
      rang: Number(r[COL_SPIELER.Rang - 1]) || 99,
      aenderungenMelden: r[COL_SPIELER.AenderungenMelden - 1] === true, rolle: '' };
  }
  return m;
}

function sendEinsatzplanEmail(spieler: Spieler, einsaetze: EinsatzInfo[], spielplanRows: string[][]): void {
  const hasBu = einsaetze.some(e => e.besondereUnterstuetzung);
  const htmlTable = buildHtmlTable(spielplanRows, hasBu);

  let personliche = '';
  for (const e of einsaetze.sort((a, b) => a.datum.localeCompare(b.datum))) {
    const h = e.heimAuswaerts || '';
    const z = e.startzeit ? ` ${e.startzeit}` : '';
    const bu = e.besondereUnterstuetzung ? ' ★' : '';
    personliche += `<tr><td>${e.datum}${z}</td><td>${h}</td><td>${e.gegner}</td><td>${e.einsatzart}${bu}</td></tr>`;
  }

  let legende = '';
  if (hasBu) {
    legende = '<p style="font-size:12px;color:#888">★ = besondere Unterstützung</p>';
  }

  const html = `<html><body style="font-family:Arial,sans-serif;font-size:14px">
<p>Hallo ${spieler.name},</p>
<p>hier ist dein <b>persönlicher Einsatzplan</b>:</p>
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:14px">
<tr style="background:#4A90D9;color:white"><th>Datum / Zeit</th><th>Ort</th><th>Gegner</th><th>Einsatzart</th></tr>
${personliche}
</table>
${legende}
<p>─── <b>GESAMTSPIELPLAN</b> ───</p>
${htmlTable}
<p>Viele Grüße,<br>Dein Einsatzplaner-Team</p>
</body></html>`;

  MailApp.sendEmail({
    to: spieler.email,
    subject: 'Dein Einsatzplan – Tischtennis',
    htmlBody: html,
  });
}
