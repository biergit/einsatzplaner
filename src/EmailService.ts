/// <reference path="ConfigTypes.ts" />

interface EinsatzInfo {
  datum: string;
  gegner: string;
  startzeit: string;
  heimAuswaerts: string;
  einsatzart: string;
  besondereUnterstuetzung: boolean;
  hinweis: string;
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
    const hinweisText = String(saisonSheet.getRange(row, saisonHinweisCol()).getValue() || '').trim();
    const datumStr = Utilities.formatDate(datum, 'Europe/Berlin', 'dd.MM.yyyy');

    const key = dateKey(datum);
    const dayMap = allAbw.get(key);
    const available = aktiveSpieler.filter(s => {
      const abw = dayMap?.get(s.name);
      return !abw || !abw.startsWith('✗');
    }).sort((a, b) => a.rang - b.rang);
    const top4Names = new Set(available.slice(0, 4).map(s => s.name));

    const matchPlayers: string[] = [];
    let spEinzel = 0;
    let spDoppel = 0;
    for (let pi = 0; pi < numPlayers; pi++) {
      const val = String(allPlayerCols[r][pi] || '').trim();
      if (!val || val.startsWith('✗')) continue;
      const name = spielerNames[pi];
      const bu = !top4Names.has(name);
      if (bu) {
        matchPlayers.push(`<span style="background:#FFF7E0;padding:1px 4px">${name} → ${val}</span>`);
      } else {
        matchPlayers.push(`${name} → ${val}`);
      }
      if (!einsaetzeProSpieler[name]) einsaetzeProSpieler[name] = [];
      einsaetzeProSpieler[name].push({ datum: datumStr, gegner, startzeit, heimAuswaerts, einsatzart: val, besondereUnterstuetzung: bu, hinweis: hinweisText });
      if (val === 'Einzel+Doppel') { spEinzel++; spDoppel++; }
      else if (val === 'Einzel') spEinzel++;
      else if (val === 'Doppel') spDoppel++;
    }

    const f = SHEET_CONFIG.einstellungen.spielformat;
    let missEinzel = Math.max(0, f.einzel - spEinzel);
    let missDoppel = Math.max(0, f.doppel * 2 - spDoppel);

    for (let ei = 0; ei < 3; ei++) {
      const eName = String(saisonSheet.getRange(row, saisonErsatzCol(ei)).getValue() || '').trim();
      if (!eName) continue;
      let ersatzArt: string;
      if (missEinzel > 0) { ersatzArt = 'Einzel (Ersatz)'; missEinzel--; }
      else if (missDoppel > 0) { ersatzArt = 'Doppel (Ersatz)'; missDoppel--; }
      else { ersatzArt = 'Einzel+Doppel (Ersatz)'; }
      matchPlayers.push(`<span style="background:#FFF7E0;padding:1px 4px">${eName} → ${ersatzArt}</span>`);
      if (!einsaetzeProSpieler[eName]) einsaetzeProSpieler[eName] = [];
      einsaetzeProSpieler[eName].push({ datum: datumStr, gegner, startzeit, heimAuswaerts, einsatzart: ersatzArt, besondereUnterstuetzung: false, hinweis: hinweisText });
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
  if (typeof val === 'number' && val > 0 && val < 24) {
    const h = Math.floor(val);
    const m = Math.round((val - h) * 60);
    return `${pad2(h)}:${pad2(m)}`;
  }
  return '';
}

function buildHtmlTable(rows: string[][]): string {
  if (rows.length === 0) return '';
  let html = '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:14px;width:100%">';
  html += '<tr style="background:#4A90D9;color:white"><th>Datum</th><th>Zeit</th><th>Ort</th><th>Gegner</th><th>Aufstellung</th></tr>';
  for (const r of rows) {
    const aufstellung = r[4].replace(/→/g, '—');
    html += `<tr><td>${r[0]}</td><td>${r[1]}</td><td>${r[2]}</td><td>${r[3]}</td><td style="font-size:14px">${aufstellung}</td></tr>`;
  }
  html += '</table>';
  return html;
}

/** Gemeinsame CSS-Styles für alle HTML-Mails. */
function emailStyles() {
  return {
    body: 'font-family:Arial,sans-serif;font-size:14px;color:#333',
    table: 'border-collapse:collapse;font-size:13px',
    th: 'background:#4A90D9;color:#fff;text-align:left;padding:6px',
    td: 'padding:4px 6px;vertical-align:top',
    footer: 'margin-top:16px;color:#888',
    header: 'margin:12px 0 4px 0;font-weight:bold;font-size:15px',
    red: 'color:#c00',
    green: 'color:#080',
  };
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

  const s = emailStyles();

  let abschnitte = '';
  for (const { name, einsaetze } of ohneEmail) {
    let rows = '';
    for (const e of einsaetze.sort((a, b) => a.datum.localeCompare(b.datum))) {
      const h = e.heimAuswaerts || '';
      const z = e.startzeit ? ` ${e.startzeit}` : '';
      rows += `<tr>
<td style="${s.td}">${escapeHtml(e.datum)}${escapeHtml(z)}</td>
<td style="${s.td}">${escapeHtml(h)}</td>
<td style="${s.td}">${escapeHtml(e.gegner)}</td>
<td style="${s.td}">${escapeHtml(e.einsatzart)}</td>
<td style="${s.td}">${escapeHtml(e.hinweis)}</td>
</tr>`;
    }
    abschnitte += `<p style="${s.header}">${escapeHtml(name)}</p>
<table border="1" cellpadding="4" cellspacing="0" style="${s.table};width:100%">
<tr><th style="${s.th}">Datum / Zeit</th><th style="${s.th}">Ort</th><th style="${s.th}">Gegner</th><th style="${s.th}">Einsatzart</th><th style="${s.th}">Hinweis</th></tr>
${rows}
</table>`;
  }

  const html = `<html><body style="${s.body}">
<p>Hallo,</p>
<p>Es wurden soeben Spielpläne finalisiert und Einsatz-Mails verschickt.<br>
Folgende eingesetzte Spieler haben keine hinterlegte E-Mail-Adresse:</p>
${abschnitte}
<p style="margin-top:12px">Bitte informiere sie persönlich über ihre Einsätze.</p>
<p style="${s.footer}">Viele Grüße,<br>Dein Einsatzplaner-Team</p>
</body></html>`;

  MailApp.sendEmail({ to: kap, subject: 'Einsatzplaner – Spieler ohne hinterlegte E-Mail-Adresse', htmlBody: html });
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
  const htmlTable = buildHtmlTable(spielplanRows);

  let personliche = '';
  for (const e of einsaetze.sort((a, b) => a.datum.localeCompare(b.datum))) {
    const h = e.heimAuswaerts || '';
    const z = e.startzeit ? ` ${e.startzeit}` : '';
    const style = e.besondereUnterstuetzung ? ' style="background:#FFF7E0"' : '';
    personliche += `<tr${style}><td>${e.datum}${z}</td><td>${h}</td><td>${e.gegner}</td><td>${e.einsatzart}</td></tr>`;
  }

  const html = `<html><body style="font-family:Arial,sans-serif;font-size:14px">
<p>Hallo ${spieler.name},</p>
<p>hier ist dein <b>persönlicher Einsatzplan</b>:</p>
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:14px">
<tr style="background:#4A90D9;color:white"><th>Datum / Zeit</th><th>Ort</th><th>Gegner</th><th>Einsatzart</th></tr>
${personliche}
</table>
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
