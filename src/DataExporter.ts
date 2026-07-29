/// <reference path="ConfigTypes.ts" />

function exportAllData(): void {
  const userEmail = Session.getActiveUser().getEmail();
  if (!userEmail || !userEmail.includes('@')) {
    SpreadsheetApp.getUi().alert(
      'Fehler',
      'Export per E-Mail benötigt ein Google-Konto. Deine E-Mail-Adresse konnte nicht ermittelt werden.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM.yyyy HH:mm');
  const s = emailStyles();
  const logo = getEmailLogo();

  let html = emailHeader() + (logo ? logo.html : '');
  html += `<p>Hallo,</p><p>hier der Datenexport vom ${timestamp}.</p>`;

  html += buildSpielerSection(ss, s);
  html += buildAbwesenheitenSection(ss, s);
  html += buildSaisonSection(ss, s);
  html += buildEinstellungenSection(s);

  html += `<p>Jede Sektion besteht aus einer HTML-Tabelle (für Copy & Paste nach Google Sheets) und einem
TSV-Block (für Copy & Paste in eine .tsv-Datei). Einstellungen liegen als JSON vor.</p>`;
  html += emailFooter();

  MailApp.sendEmail({
    to: userEmail,
    subject: `Einsatzplaner – Datenexport ${timestamp}`,
    htmlBody: html,
    ...(logo && { inlineImages: logo.inlineImages }),
  });
}

// ─── HTML-Builder ──────────────────────────────────────────────────────────

function htmlTable(headers: string[], rows: string[][], s: ReturnType<typeof emailStyles>): string {
  let h = `<table border="1" cellpadding="4" cellspacing="0" style="${s.table}">`;
  h += '<tr>';
  for (const th of headers) h += `<th style="${s.th}">${escapeHtml(th)}</th>`;
  h += '</tr>';
  for (const row of rows) {
    h += '<tr>';
    for (const cell of row) h += `<td style="${s.td}">${escapeHtml(cell)}</td>`;
    h += '</tr>';
  }
  h += '</table>';
  return h;
}

function tsvBlock(headers: string[], rows: string[][]): string {
  let tsv = headers.join('\t') + '\n';
  for (const row of rows) tsv += row.join('\t') + '\n';
  return `<p style="margin-top:4px;font-weight:bold">TSV (in .tsv-Datei einfügen):</p><pre style="font-size:12px;background:#f5f5f5;padding:8px;overflow:auto">${escapeHtml(tsv)}</pre>`;
}

function section(title: string, filename: string, htmlContent: string, tsvContent: string): string {
  return `<p style="margin-top:20px;font-weight:bold;font-size:15px">${escapeHtml(title)} (${escapeHtml(filename)})</p>${htmlContent}${tsvContent}`;
}

// ─── Spieler ───────────────────────────────────────────────────────────────

function buildSpielerSection(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, s: ReturnType<typeof emailStyles>): string {
  const headers = ['Name', 'Email', 'Rang', 'Änderungen melden', 'Rolle'];
  const sheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  const rows: string[][] = [];
  if (sheet) {
    const lr = sheet.getLastRow();
    if (lr > 1) {
      const data = sheet.getRange(2, 1, lr - 1, COL_SPIELER.Rolle).getValues();
      for (const row of data) {
        const name = String(row[COL_SPIELER.Name - 1] || '').trim();
        if (!name) continue;
        rows.push([
          name,
          String(row[COL_SPIELER.Email - 1] || '').trim(),
          String(row[COL_SPIELER.Rang - 1] || ''),
          row[COL_SPIELER.AenderungenMelden - 1] === true ? 'ja' : '',
          String(row[COL_SPIELER.Rolle - 1] || '').trim(),
        ]);
      }
    }
  }
  return section('Spieler', 'spieler.tsv', htmlTable(headers, rows, s), tsvBlock(headers, rows));
}

// ─── Abwesenheiten ─────────────────────────────────────────────────────────

function buildAbwesenheitenSection(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, s: ReturnType<typeof emailStyles>): string {
  const headers = ['Spieler/in', 'Abwesenheit von', 'Abwesenheit bis', 'Kommentar'];
  const sheet = ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN);
  const rows: string[][] = [];
  if (sheet) {
    const lr = sheet.getLastRow();
    if (lr > 1) {
      const data = sheet.getRange(2, 1, lr - 1, COL_ABWESENHEITEN.Kommentar).getValues();
      for (const row of data) {
        const spieler = String(row[COL_ABWESENHEITEN.Spieler - 1] || '').trim();
        if (!spieler) continue;
        const von = formatDate(row[COL_ABWESENHEITEN.Von - 1]);
        const bis = formatDate(row[COL_ABWESENHEITEN.Bis - 1]);
        if (!von || !bis) continue;
        rows.push([spieler, von, bis, String(row[COL_ABWESENHEITEN.Kommentar - 1] || '').trim()]);
      }
    }
  }
  return section('Abwesenheiten', 'abwesenheiten.tsv', htmlTable(headers, rows, s), tsvBlock(headers, rows));
}

// ─── Saison ────────────────────────────────────────────────────────────────

function buildSaisonSection(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, s: ReturnType<typeof emailStyles>): string {
  const playerNames = SHEET_CONFIG.spieler.map(p => p.name);
  const headers = ['Datum', 'Wochentag', 'Gegner', 'Startzeit', 'Heim/Auswärts', ...playerNames, 'Ersatz 1', 'Ersatz 2', 'Ersatz 3', 'Status', 'Kommentar'];
  const sheet = ss.getSheetByName(SHEET_NAMES.SAISON);
  const rows: string[][] = [];
  if (sheet) {
    const lr = sheet.getLastRow();
    if (lr > 1) {
      const numCols = saisonColCount();
      const data = sheet.getRange(2, 1, lr - 1, numCols).getValues();
      for (const row of data) {
        const rowArr: string[] = [];
        for (const cell of row) {
          if (cell instanceof Date) {
            rowArr.push(Utilities.formatDate(cell, 'Europe/Berlin', 'dd.MM.yyyy'));
          } else {
            rowArr.push(String(cell || '').trim());
          }
        }
        rows.push(rowArr);
      }
    }
  }
  return section('Saison', 'saison.tsv', htmlTable(headers, rows, s), tsvBlock(headers, rows));
}

// ─── Einstellungen ─────────────────────────────────────────────────────────

function buildEinstellungenSection(s: ReturnType<typeof emailStyles>): string {
  const json = JSON.stringify(SHEET_CONFIG.einstellungen, null, 2);
  return `<p style="margin-top:20px;font-weight:bold;font-size:15px">Einstellungen (einstellungen.json)</p>` +
    `<p style="font-weight:bold">JSON (in einstellungen.json einfügen):</p>` +
    `<pre style="font-size:12px;background:#f5f5f5;padding:8px;overflow:auto">${escapeHtml(json)}</pre>`;
}

// ─── Helper ────────────────────────────────────────────────────────────────

function formatDate(val: unknown): string {
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Europe/Berlin', 'dd.MM.yyyy');
  }
  return '';
}
