/// <reference path="ConfigTypes.ts" />

interface EinsatzInfo {
  datum: string;
  gegner: string;
  startzeit: string;
  heimAuswaerts: string;
  einsatzart: string;
  besondereUnterstuetzung: boolean;
  kommentar: string;
}

function sendEinsatzEmails(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const saisonSheet = ss.getSheetByName(SHEET_NAMES.SAISON);
  if (!saisonSheet) return;

  const lastRow = saisonSheet.getLastRow();
  const heute = new Date(); heute.setHours(0, 0, 0, 0);

  const spielerNames = readSpielerNames(ss);
  const spieler = readSpieler(ss.getSheetByName(SHEET_NAMES.SPIELER)!);
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
    const kommentar = String(saisonSheet.getRange(row, saisonKommentarCol()).getValue() || '').trim();
    const datumStr = Utilities.formatDate(datum, 'Europe/Berlin', 'dd.MM.yyyy');

    const key = dateKey(datum);
    const dayMap = allAbw.get(key);
    const available = spieler.filter((s: Spieler) => {
      const abw = dayMap?.get(s.name);
      return !abw || !abw.startsWith('✗');
    }).sort((a: Spieler, b: Spieler) => a.rang - b.rang);
    const top4Names = new Set(available.slice(0, 4).map((s: Spieler) => s.name));

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
      einsaetzeProSpieler[name].push({ datum: datumStr, gegner, startzeit, heimAuswaerts, einsatzart: val, besondereUnterstuetzung: bu, kommentar });
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
      einsaetzeProSpieler[eName].push({ datum: datumStr, gegner, startzeit, heimAuswaerts, einsatzart: ersatzArt, besondereUnterstuetzung: false, kommentar });
    }

    spielplanRows.push([datumStr, startzeit, heimAuswaerts, gegner, matchPlayers.map(p => `<div style="line-height:1.6">${p}</div>`).join(''), kommentar]);
  }

  const spielerMap: Record<string, Spieler> = {};
  for (const s of spieler) spielerMap[s.name] = s;

  const kapEmail = getKapitaenEmail(ss);

  for (const [spielerName, einsaetze] of Object.entries(einsaetzeProSpieler)) {
    const s = spielerMap[spielerName];
    if (!s || !s.email || !s.email.includes('@')) {
      ohneEmail.push({ name: spielerName, einsaetze });
      continue;
    }
    // Kapitän separat am Ende (bekommt ggf. Ohne-Email-Tabelle)
    if (s.email === kapEmail) continue;
    sendEinsatzplanEmail(s, einsaetze, spielplanRows, '');
  }

  // Kapitän-Mail: persönlicher Plan + ggf. Ohne-Email-Tabelle
  if (kapEmail) {
    const kapSpieler = spieler.find((s: Spieler) => s.email === kapEmail);
    if (kapSpieler) {
      const kapEinsaetze = einsaetzeProSpieler[kapSpieler.name] || [];
      const ohneHtml = ohneEmail.length > 0 ? buildOhneEmailTabelle(ohneEmail) : '';
      sendEinsatzplanEmail(kapSpieler, kapEinsaetze, spielplanRows, ohneHtml);
      if (ohneEmail.length > 0) {
        Logger.log(`sendEinsatzEmails: Kapitän-Mail mit ${ohneEmail.length} Ohne-Email-Einträgen`);
      }
    }
  } else if (ohneEmail.length > 0) {
    Logger.log(`sendEinsatzEmails: ${ohneEmail.length} Spieler ohne E-Mail, aber kein Kapitän mit E-Mail gefunden`);
  }

  Logger.log(`sendEinsatzEmails: ${Object.keys(einsaetzeProSpieler).length} Spieler, ${ohneEmail.length} ohne E-Mail`);
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

/**
 * Baut eine HTML-Tabelle aus Ohne-Email-Daten (Spieler ohne hinterlegte
 * E-Mail-Adresse, die auf Final-Spieltagen aufgestellt sind).
 */
function buildOhneEmailTabelle(ohneEmail: { name: string; einsaetze: EinsatzInfo[] }[]): string {
  if (ohneEmail.length === 0) return '';

  const sortKey = (d: string) => {
    const p = d.split('.');
    return p.length === 3 ? p[2] + p[1] + p[0] : d;
  };

  let html = '<h1>Betroffene - nicht per E-Mail benachrichtigte Spieler</h1>';
    html += 'Folgende Spieler konnten nicht per E-Mail informiert werden, da keine E-Mail-Adresse hinterlegt war:'

  for (const { name, einsaetze } of ohneEmail) {
    const sorted = [...einsaetze].sort((a, b) => sortKey(a.datum).localeCompare(sortKey(b.datum)));

    let rows = '';
    for (const e of sorted) {
      rows += `<tr><td style="padding:2px 6px">${escapeHtml(e.datum)} — ${escapeHtml(e.gegner)}</td><td style="padding:2px 6px">${escapeHtml(e.einsatzart)}</td></tr>`;
    }

    html += `<h2>${escapeHtml(name)}</h2>
<table border="1" cellpadding="4" cellspacing="0" style="font-size:13px;border-collapse:collapse;margin-bottom:12px">
<tr style="background:#4A90D9;color:white"><th>Spieltag</th><th>Einsatz</th></tr>
${rows}
</table>`;
  }

  return html;
}

function buildHtmlTable(rows: string[][]): string {
  if (rows.length === 0) return '';
  const hasKommentar = rows.some(r => r[5]);
  let html = '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:14px;width:100%">';
  html += '<tr style="background:#4A90D9;color:white"><th>Datum</th><th>Zeit</th><th>Ort</th><th>Gegner</th><th>Aufstellung</th>';
  if (hasKommentar) html += '<th>Kommentar</th>';
  html += '</tr>';
  for (const r of rows) {
    const aufstellung = r[4].replace(/→/g, '—');
    const kmt = hasKommentar ? `<td>${escapeHtml(r[5] || '')}</td>` : '';
    html += `<tr><td>${r[0]}</td><td>${r[1]}</td><td>${r[2]}</td><td>${r[3]}</td><td style="font-size:14px;line-height:1.6">${aufstellung}</td>${kmt}</tr>`;
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

function getKapitaenEmail(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): string {
  const s = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (!s) return '';
  const lr = s.getLastRow();
  if (lr <= 1) return '';
  const d = s.getRange(2, 1, lr - 1, COL_SPIELER.Rolle).getValues();
  for (const r of d) {
    const name = String(r[COL_SPIELER.Name - 1] || '').trim();
    if (!name) continue;
    if (String(r[COL_SPIELER.Email - 1]).includes('@') && String(r[COL_SPIELER.Rolle - 1]).trim() === 'Kapitän')
      return String(r[COL_SPIELER.Email - 1]).trim();
  }
  return '';
}

function sendEinsatzplanEmail(spieler: Spieler, einsaetze: EinsatzInfo[], spielplanRows: string[][], ohneHtml: string): void {
  const htmlTable = buildHtmlTable(spielplanRows);

  const sorted = [...einsaetze].sort((a, b) => {
    const [da, ma, ya] = a.datum.split('.').map(Number);
    const [db, mb, yb] = b.datum.split('.').map(Number);
    return new Date(ya, ma - 1, da).getTime() - new Date(yb, mb - 1, db).getTime();
  });

  const hasKommentar = sorted.some(e => e.kommentar);

  let personliche = '';
  for (const e of sorted) {
    const h = e.heimAuswaerts || '';
    const z = e.startzeit ? ` ${e.startzeit}` : '';
    const style = e.besondereUnterstuetzung ? ' style="background:#FFF7E0"' : '';
    const kmt = hasKommentar ? `<td>${escapeHtml(e.kommentar)}</td>` : '';
    personliche += `<tr${style}><td>${e.datum}${z}</td><td>${h}</td><td>${e.gegner}</td><td>${e.einsatzart}</td>${kmt}</tr>`;
  }

  const kmtHead = hasKommentar ? '<th>Kommentar</th>' : '';

  const html = `<html><body style="font-family:Arial,sans-serif;font-size:14px">
<p>Hallo ${spieler.name},</p>
<p>hier ist dein persönlicher</p>
<h1>EINSATZPLAN</h1>
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:14px">
<tr style="background:#4A90D9;color:white"><th>Datum / Zeit</th><th>Ort</th><th>Gegner</th><th>Einsatzart</th>${kmtHead}</tr>
${personliche}
</table>
<p>und der</p>
<h1>GESAMTSPIELPLAN</h1>
${htmlTable}
${ohneHtml}
<p>Viele Grüße,<br>Dein Einsatzplaner-Team</p>
</body></html>`;

  MailApp.sendEmail({
    to: spieler.email,
    subject: `Neuer ${SHEET_CONFIG.einstellungen.teamName} - Einsatzplan`,
    htmlBody: html,
  });
}
