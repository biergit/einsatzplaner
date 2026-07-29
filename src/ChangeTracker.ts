/// <reference path="ConfigTypes.ts" />

interface ChangeEntry {
  /** Unix-Millisekunden (für JSON-Serialisierung in ScriptProperties) */
  timestamp: number;
  /** Sheet-Name ohne Blatt (z.B. "Saison") */
  sheetName: string;
  /** A1-Notation ohne Blatt-Präfix (z.B. "C5") */
  rangeA1: string;
  alterWert: string;
  neuerWert: string;
  bearbeiter: string;
}

interface SheetRowSnapshot {
  row: number;
  /** Inhaltlicher Key für Move-Erkennung (z.B. "Max|01.01.2026|15.01.2026|Urlaub") */
  key: string;
  /** Für das Log formatierter Anzeigetext */
  display: string;
}

interface AbwAffectedEntry {
  datumStr: string;
  startzeit: string;
  heimAuswaerts: string;
  gegner: string;
  spielerName: string;
  /** true wenn der Spieler eine Einsatzart (Einzel/Doppel/…) eingetragen hatte */
  warAufgestellt: boolean;
  /** true wenn der Spieltag-Status vor dem Revert "Final" war */
  warFinal: boolean;
  /** Validierungsmeldungen des Spieltags nach der Änderung (für Mail-Hinweis) */
  validierung: string;
}

interface SaisonRowSnapshot {
  datum: string;
  gegner: string;
  startzeit: string;
  heimAuswaerts: string;
  playerAssignments: string[];
  ersatz: string[];
  status: string;
  kommentar: string;
}

// ─── onEdit (nur installierbarer Trigger läuft durch) ─────────────────────

/**
 * Installierbarer onEdit-Handler.
 * Simple-Trigger mit authMode !== FULL werden ignoriert (haben keine ScriptApp-Rechte).
 */
function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  if (e.authMode !== ScriptApp.AuthMode.FULL) {
    Logger.log('onEdit: Simple-Trigger (LIMITED) — ignoriert');
    return;
  }

  Logger.log(`onEdit: Installierbarer-Trigger (FULL) — Sheet="${e.range.getSheet().getName()}" Row=${e.range.getRow()} Col=${e.range.getColumn()}`);

  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SHEET_BUILDER_RUNNING') === 'true') return;

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName === SHEET_NAMES.AENDERUNGSLOG || sheetName === SHEET_NAMES.DOKUMENTATION) return;

  const row = range.getRow();
  const col = range.getColumn();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Sheet-spezifische Aktionen (gekapselt, damit addPendingEdit + resetDebounceTimer
  // auch bei Fehlern in diesen Blöcken garantiert laufen)
  try {
    // Saison: Gegner → Status autom. setzen; Validierung aktualisieren
    if (sheetName === SHEET_NAMES.SAISON && row >= 2) {
      if (col === saisonGegnerCol()) {
        const newValue = e.value !== undefined ? String(e.value) : '';
        const statusCell = sheet.getRange(row, saisonStatusCol());
        if (newValue.trim()) {
          if (!String(statusCell.getValue() || '').trim()) statusCell.setValue('Geplant');
        } else {
          statusCell.setValue('');
        }
      }
      validateAndUpdateSaisonRow(range);
    }

    // Abwesenheiten: ✗-Markierungen im Saison-Sheet sofort neu aufbauen
    if (sheetName === SHEET_NAMES.ABWESENHEITEN && row >= 2) {
      Logger.log(`onEdit: Abwesenheiten-Edit Zeile ${row}, Spalte ${col} → rebuildAbwesenheitenInSaison`);
      rebuildAbwesenheitenInSaison(ss);
      restoreAbwesenheitenValidations(sheet);
      validateAbwesenheitRow(sheet, row);
    }

    // Spieler: Data-Validations (Checkboxen, Dropdowns) nach Cut+Paste wiederherstellen
    if (sheetName === SHEET_NAMES.SPIELER && row >= 2) {
      restoreSpielerValidations(sheet);
    }
  } catch (err) {
    Logger.log(`onEdit: Fehler bei Sheet-Aktionen — ${err}`);
  }

  // Änderung für den Debounce sammeln
  const isSingleCell = range.getNumRows() === 1 && range.getNumColumns() === 1;
  let oldValue = e.oldValue !== undefined ? String(e.oldValue) : '';
  let newValue = e.value !== undefined ? String(e.value) : '';

  if (sheetName === SHEET_NAMES.SAISON && col === saisonStartzeitCol()) {
    oldValue = formatTimeForLog(e.oldValue) || oldValue;
    const formatted = formatTimeForLog(e.value);
    if (formatted) newValue = formatted;
  }

  const userEmail = Session.getActiveUser().getEmail();
  const bearbeiter = userEmail && userEmail.includes('@') ? userEmail : 'Unbekannt';

  addPendingEdit({
    timestamp: Date.now(),
    sheetName,
    rangeA1: range.getA1Notation(),
    alterWert: oldValue,
    neuerWert: isSingleCell ? newValue : formatMultiCellEdit(sheet, range),
    bearbeiter,
  });

  resetDebounceTimer();

  // Fallback: Falls resetDebounceTimer keinen Trigger erstellen konnte
  // (z.B. Quota-Limit), wird DEBOUNCE_FAILED gesetzt. In diesem Fall
  // die Änderungen sofort verarbeiten, damit sie nicht verloren gehen.
  // NIEMALS sofort verarbeiten, nur weil PENDING_EDITS nicht leer ist —
  // das ist der Normalfall während der laufenden Debounce-Periode.
  // Der Debounce existiert, damit der Nutzer Zeit hat, die Aufstellung
  // zu korrigieren, bevor die Mail verschickt wird.
  if (props.getProperty('DEBOUNCE_FAILED') === 'true') {
    Logger.log('onEdit: DEBOUNCE_FAILED — sofortige Verarbeitung');
    props.deleteProperty('DEBOUNCE_FAILED');
    try {
      for (const t of ScriptApp.getProjectTriggers()) {
        if (t.getHandlerFunction() === 'onDebounceTimer') ScriptApp.deleteTrigger(t);
      }
      flushPendingChanges();
    } catch (err) {
      Logger.log(`onEdit: flushPendingChanges-Fallback fehlgeschlagen — ${err}`);
    }
  }
}

// ─── Abwesenheiten → Saison-Sheet sofort aktualisieren ─────────────────────

/**
 * Baut sämtliche ✗-Markierungen im Saison-Sheet aus den aktuellen Abwesenheiten
 * komplett neu auf. Setzt betroffene "Final"-Spieltage auf "Geplant" zurück.
 * Schreibt betroffene Spieltage als ABW_AFFECTED in ScriptProperties
 * (für die Mail-Anreicherung in sendChangeNotification).
 */
function rebuildAbwesenheitenInSaison(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): void {
  const saisonSheet = ss.getSheetByName(SHEET_NAMES.SAISON);
  if (!saisonSheet) return;
  const lastRow = saisonSheet.getLastRow();
  if (lastRow <= 1) return;
  const numRows = lastRow - 1;

  const spielerNames = readSpielerNames(ss);
  const numPlayers = spielerNames.length;
  if (numPlayers === 0) return;

  const allAbw = buildAbwesenheitenIndex(ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN)!);
  const dates = readSaisonDates(saisonSheet, lastRow);

  // Spieler-Zellen + Status als Batch lesen
  const allPlayerCols = saisonSheet.getRange(2, saisonSpielerCol(0), numRows, numPlayers).getValues() as string[][];
  const statusValues = saisonSheet.getRange(2, saisonStatusCol(), numRows, 1).getValues() as string[][];

  // Zeilen-Metadaten für ABW_AFFECTED (nur für betroffene Zeilen)
  const gegnerCols = saisonSheet.getRange(2, saisonGegnerCol(), numRows, 1).getValues();
  const heimCols = saisonSheet.getRange(2, saisonHeimAuswaertsCol(), numRows, 1).getValues();
  const startCols = saisonSheet.getRange(2, saisonStartzeitCol(), numRows, 1).getValues();

  let playerChanged = false;
  let statusChanged = false;
  const affected: AbwAffectedEntry[] = [];

  for (let r = 0; r < numRows; r++) {
    const d = dates[r];
    if (!d) continue;
    const key = dateKey(d);
    const dayMap = allAbw.get(key);

    const gegner = String(gegnerCols[r][0] || '').trim();
    const heimAw = String(heimCols[r][0] || '').trim();
    const isSpieltag = !!gegner;
    const wasFinalBefore = String(statusValues[r][0] || '').trim() === 'Final';

    for (let pi = 0; pi < numPlayers; pi++) {
      const current = String(allPlayerCols[r][pi] || '').trim();
      const abwDisplay = dayMap?.get(spielerNames[pi]);

      if (abwDisplay) {
        if (abwDisplay.startsWith('✗')) {
          if (!current || !current.startsWith('✗') || current !== abwDisplay) {
            const warAufgestellt = !!(current && !current.startsWith('✗'));
            if (warAufgestellt && wasFinalBefore) {
              statusValues[r][0] = 'Geplant';
              statusChanged = true;
            }
            allPlayerCols[r][pi] = abwDisplay;
            playerChanged = true;

            // Nur protokollieren wenn es ein Spieltag mit Gegner ist
            if (isSpieltag && warAufgestellt) {
              const datumStr = Utilities.formatDate(d, 'Europe/Berlin', 'dd.MM.yyyy');
              affected.push({
                datumStr,
                startzeit: formatStartzeit(startCols[r][0]),
                heimAuswaerts: heimAw,
                gegner,
                spielerName: spielerNames[pi],
                warAufgestellt,
                warFinal: wasFinalBefore && warAufgestellt,
                validierung: '',
              });
            }
          }
        }
      } else {
        if (current.startsWith('✗')) {
          allPlayerCols[r][pi] = '';
          playerChanged = true;
        }
      }
    }
  }

  if (playerChanged) {
    saisonSheet.getRange(2, saisonSpielerCol(0), numRows, numPlayers).setValues(allPlayerCols);
  }
  if (statusChanged) {
    saisonSheet.getRange(2, saisonStatusCol(), numRows, 1).setValues(statusValues);
  }

  if (playerChanged || statusChanged) {
    const refreshedAllAbw = buildAbwesenheitenIndex(ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN)!);
    const allSpieler: Spieler[] = spielerNames.map(name => ({
      name, email: '', rang: 99, aenderungenMelden: false, rolle: ''
    }));
    const ersatzData = saisonSheet.getRange(2, saisonErsatzCol(0), numRows, 3).getValues();
    const validierungValues: string[][] = [];
    const validierungPerRow = new Map<string, string>();
    for (let r = 0; r < numRows; r++) {
      const gegner = String(gegnerCols[r][0] || '').trim();
      if (!gegner) { validierungValues.push(['']); continue; }
      const v = computeValidierung(allPlayerCols[r], dates[r], allSpieler, refreshedAllAbw,
        [String(ersatzData[r][0] || ''), String(ersatzData[r][1] || ''), String(ersatzData[r][2] || '')]);
      validierungValues.push([v]);
      validierungPerRow.set(`${r}|${gegner}`, v);
    }
    saisonSheet.getRange(2, saisonValidierungCol(), numRows, 1).setValues(validierungValues);

    // Betroffenen Einträgen die Validierungsmeldungen zuweisen
    for (const a of affected) {
      for (let r = 0; r < numRows; r++) {
        const gegner = String(gegnerCols[r][0] || '').trim();
        if (!gegner) continue;
        const datumStr = dates[r] ? Utilities.formatDate(dates[r], 'Europe/Berlin', 'dd.MM.yyyy') : '';
        if (a.datumStr === datumStr && a.gegner === gegner) {
          a.validierung = validierungPerRow.get(`${r}|${gegner}`) || '';
          break;
        }
      }
    }
  }

  // Betroffene Spieltage für die spätere Mail-Anreicherung speichern
  const props = PropertiesService.getScriptProperties();
  props.setProperty('ABW_AFFECTED', JSON.stringify(affected));
  if (affected.length > 0) {
    const finalCount = affected.filter(a => a.warFinal).length;
    Logger.log(`rebuildAbwesenheitenInSaison: ${affected.length} betroffene Spieltage (davon ${finalCount} mit Final→Geplant-Revert)`);
  }
}

// ─── Formatierungs-Helfer ──────────────────────────────────────────────────

/**
 * Für Multi-Cell-Edits (Cut+Paste) einen lesbaren String aus der aktuellen
 * Zeilenbelegung bilden. Wichtig für den ersten Debounce-Durchlauf ohne Snapshot.
 */
function formatMultiCellEdit(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  range: GoogleAppsScript.Spreadsheet.Range
): string {
  if (sheet.getName() === SHEET_NAMES.ABWESENHEITEN) {
    const row = range.getRow();
    const data = sheet.getRange(row, 1, 1, COL_ABWESENHEITEN.Kommentar).getValues()[0];
    const spieler = String(data[COL_ABWESENHEITEN.Spieler - 1] || '').trim();
    const von = toDate(data[COL_ABWESENHEITEN.Von - 1]);
    const bis = toDate(data[COL_ABWESENHEITEN.Bis - 1]);
    const kommentar = String(data[COL_ABWESENHEITEN.Kommentar - 1] || '').trim();
    const vonStr = von ? Utilities.formatDate(von, 'Europe/Berlin', 'dd.MM.yyyy') : '-';
    const bisStr = bis ? Utilities.formatDate(bis, 'Europe/Berlin', 'dd.MM.yyyy') : '-';
    return `Spieler: ${spieler || '-'} | Von: ${vonStr} | Bis: ${bisStr} | Kommentar: ${kommentar || '-'}`;
  }
  return `(Mehrere Zellen, ${range.getNumRows()}×${range.getNumColumns()})`;
}

/**
 * Wandelt Rohwerte der Startzeit-Spalte in "HH:MM" um.
 * Unterstützt Date-Objekte und Zahlen (Bruchteil eines Tages oder ganze Stunden).
 */
function formatTimeForLog(val: unknown): string {
  if (val instanceof Date) {
    const h = val.getHours();
    const m = val.getMinutes();
    if (h === 0 && m === 0 && val.getFullYear() === 1899) return '';
    return `${pad2(h)}:${pad2(m)}`;
  }
  if (typeof val === 'number' && val > 0) {
    if (val < 1) {
      const totalMinutes = val * 24 * 60;
      const h = Math.floor(totalMinutes / 60);
      const m = Math.round(totalMinutes % 60);
      return `${pad2(h)}:${pad2(m)}`;
    }
    if (val < 24) {
      const h = Math.floor(val);
      const m = Math.round((val - h) * 60);
      return `${pad2(h)}:${pad2(m)}`;
    }
  }
  return '';
}

// ─── Abwesenheiten-Validierung ─────────────────────────────────────────────

/**
 * Prüft die Von/Bis-Daten einer Abwesenheitszeile und schreibt
 * eine Validierungsmeldung in die Validierungsspalte.
 */
function validateAbwesenheitRow(sheet: GoogleAppsScript.Spreadsheet.Sheet, row: number): void {
  const data = sheet.getRange(row, 1, 1, COL_ABWESENHEITEN.Validierung).getValues()[0];
  const spieler = String(data[COL_ABWESENHEITEN.Spieler - 1] || '').trim();
  const von = toDate(data[COL_ABWESENHEITEN.Von - 1]);
  const bis = toDate(data[COL_ABWESENHEITEN.Bis - 1]);

  // Nur validieren wenn mindestens ein Feld gefüllt ist
  if (!spieler && !von && !bis) {
    sheet.getRange(row, COL_ABWESENHEITEN.Validierung).setValue('');
    return;
  }

  const fehler: string[] = [];

  if (!von || !bis) {
    fehler.push('Von- und Bis-Datum erforderlich');
  } else if (von.getTime() > bis.getTime()) {
    fehler.push('Von-Datum darf nicht nach dem Bis-Datum liegen');
  }

  sheet.getRange(row, COL_ABWESENHEITEN.Validierung).setValue(fehler.join(' | '));
}

// ─── Data-Validation nach Cut+Paste wiederherstellen ───────────────────────

/** Stellt Spieler-Dropdown, Datumsformat und Kalender-Picker im Abwesenheiten-Sheet wieder her. */
function restoreAbwesenheitenValidations(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  const lastRow = Math.max(sheet.getMaxRows() - 1, 2);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (spielerSheet) {
    const spielerLastRow = spielerSheet.getLastRow();
    if (spielerLastRow > 1) {
      sheet.getRange(2, COL_ABWESENHEITEN.Spieler, lastRow, 1).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInRange(spielerSheet.getRange(2, COL_SPIELER.Name, spielerLastRow - 1, 1))
          .setAllowInvalid(true).build());
    }
  }
  sheet.getRange(2, COL_ABWESENHEITEN.Von, lastRow, 2).setNumberFormat('DD.MM.YYYY');
  // Kalender-Element (Date Picker) nach Cut+Paste wiederherstellen
  const dateValidation = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .setHelpText('Datum im Kalender auswählen (oder frei eingeben)')
    .build();
  sheet.getRange(2, COL_ABWESENHEITEN.Von, lastRow, 1).setDataValidation(dateValidation);
  sheet.getRange(2, COL_ABWESENHEITEN.Bis, lastRow, 1).setDataValidation(dateValidation);
}

/** Stellt Checkboxen und Dropdowns im Spieler-Sheet wieder her. */
function restoreSpielerValidations(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  const lastRow = Math.max(sheet.getMaxRows() - 1, 2);
  sheet.getRange(2, COL_SPIELER.Rang, lastRow, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberGreaterThan(0).setAllowInvalid(true).build());
  sheet.getRange(2, COL_SPIELER.AenderungenMelden, lastRow, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().build());
  sheet.getRange(2, COL_SPIELER.Rolle, lastRow, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Kapitän']).setAllowInvalid(true).build());
}

// ─── Snapshot-Fabriken (für Abwesenheiten- und Spieler-Diff) ───────────────

/**
 * Liest alle belegten Zeilen des Abwesenheiten-Sheets als Snapshot-Array.
 * Key = `spieler|von|bis|kommentar` für inhaltsbasierte Move-Erkennung.
 */
function readAbwSnapshot(sheet: GoogleAppsScript.Spreadsheet.Sheet | null): SheetRowSnapshot[] {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, COL_ABWESENHEITEN.Kommentar).getValues();
  const rows: SheetRowSnapshot[] = [];
  for (let r = 0; r < data.length; r++) {
    const spieler = String(data[r][COL_ABWESENHEITEN.Spieler - 1] || '').trim();
    const von = toDate(data[r][COL_ABWESENHEITEN.Von - 1]);
    const bis = toDate(data[r][COL_ABWESENHEITEN.Bis - 1]);
    const kommentar = String(data[r][COL_ABWESENHEITEN.Kommentar - 1] || '').trim();
    if (!spieler && !von && !bis && !kommentar) continue;
    const vonStr = von ? Utilities.formatDate(von, 'Europe/Berlin', 'dd.MM.yyyy') : '';
    const bisStr = bis ? Utilities.formatDate(bis, 'Europe/Berlin', 'dd.MM.yyyy') : '';
    rows.push({
      row: r + 2,
      key: `${spieler}|${vonStr}|${bisStr}|${kommentar}`,
      display: `Spieler: ${spieler} | Von: ${vonStr} | Bis: ${bisStr} | Kommentar: ${kommentar || '-'}`,
    });
  }
  return rows;
}

/**
 * Liest alle belegten Zeilen des Spieler-Sheets als Snapshot-Array.
 * Key = `name|email|rang|melden|rolle`.
 */
function readSpielerSnapshot(sheet: GoogleAppsScript.Spreadsheet.Sheet | null): SheetRowSnapshot[] {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, COL_SPIELER.Rolle).getValues();
  const rows: SheetRowSnapshot[] = [];
  for (let r = 0; r < data.length; r++) {
    const name = String(data[r][COL_SPIELER.Name - 1] || '').trim();
    if (!name) continue;
    const email = String(data[r][COL_SPIELER.Email - 1] || '').trim();
    const rang = String(data[r][COL_SPIELER.Rang - 1] || '').trim();
    const melden = data[r][COL_SPIELER.AenderungenMelden - 1] === true ? 'ja' : '';
    const rolle = String(data[r][COL_SPIELER.Rolle - 1] || '').trim();
    rows.push({
      row: r + 2,
      key: `${name}|${email}|${rang}|${melden}|${rolle}`,
      display: `Name: ${name} | Email: ${email || '-'} | Rang: ${rang} | Melden: ${melden || 'nein'} | Rolle: ${rolle || '-'}`,
    });
  }
  return rows;
}

/**
 * Liest alle Saison-Zeilen mit Gegner (= Spieltage) als Snapshot.
 * Für die zeilenweise Vorher-/Nachher-Darstellung in der Änderungsmail.
 */
function readSaisonSnapshot(sheet: GoogleAppsScript.Spreadsheet.Sheet | null): SaisonRowSnapshot[] {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const numCols = saisonColCount();
  const numPlayers = SHEET_CONFIG.spieler.length;
  const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  const rows: SaisonRowSnapshot[] = [];

  for (let r = 0; r < data.length; r++) {
    const gegner = String(data[r][saisonGegnerCol() - 1] || '').trim();
    if (!gegner) continue;

    const datumVal = data[r][saisonDatumCol() - 1];
    const datum = datumVal instanceof Date
      ? Utilities.formatDate(datumVal, 'Europe/Berlin', 'dd.MM.yyyy') : '';

    const assignments: string[] = [];
    for (let pi = 0; pi < numPlayers; pi++) {
      assignments.push(String(data[r][saisonSpielerCol(pi) - 1] || '').trim());
    }

    const ersatz: string[] = [];
    for (let ei = 0; ei < 3; ei++) {
      ersatz.push(String(data[r][saisonErsatzCol(ei) - 1] || '').trim());
    }

    rows.push({
      datum,
      gegner,
      startzeit: formatStartzeit(data[r][saisonStartzeitCol() - 1]),
      heimAuswaerts: String(data[r][saisonHeimAuswaertsCol() - 1] || '').trim(),
      playerAssignments: assignments,
      ersatz,
      status: String(data[r][saisonStatusCol() - 1] || '').trim(),
      kommentar: String(data[r][saisonKommentarCol() - 1] || '').trim(),
    });
  }
  return rows;
}

// ─── Debounce-Mechanismus ──────────────────────────────────────────────────

/**
 * Hängt einen ChangeEntry an die PENDING_EDITS-Liste in ScriptProperties an.
 * Die tatsächliche Verarbeitung erfolgt erst beim Ablauf des Debounce-Timers.
 */
function addPendingEdit(entry: ChangeEntry): void {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('PENDING_EDITS') || '[]';
  const edits: ChangeEntry[] = JSON.parse(raw);
  edits.push(entry);
  props.setProperty('PENDING_EDITS', JSON.stringify(edits));
}

/**
 * Stellt sicher, dass ein onDebounceTimer-Trigger existiert.
 * Während Bulk-Operationen (SHEET_BUILDER_RUNNING oder BULK_EDIT) wird
 * kein Timer angelegt — der Aufrufer muss das selbst am Ende tun.
 *
 * Falls die Trigger-Erstellung fehlschlägt (Quota, Berechtigungen), wird
 * ein DEBOUNCE_FAILED-Flag gesetzt. Der nächste onEdit-Aufruf erkennt das
 * und verarbeitet die ausstehenden Änderungen sofort.
 */
function resetDebounceTimer(): void {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SHEET_BUILDER_RUNNING') === 'true') return;
  if (props.getProperty('BULK_EDIT') === 'true') return;

  for (const t of ScriptApp.getProjectTriggers()) {
    if (t.getHandlerFunction() === 'onDebounceTimer') ScriptApp.deleteTrigger(t);
  }
  const minuten = SHEET_CONFIG.einstellungen.debounceMinuten;
  try {
    ScriptApp.newTrigger('onDebounceTimer')
      .timeBased().after(minuten * 60 * 1000).create();
    Logger.log(`resetDebounceTimer: Timer in ${minuten} Min. erstellt`);
  } catch (e) {
    Logger.log(`resetDebounceTimer: Trigger-Erstellung fehlgeschlagen — ${e}`);
    props.setProperty('DEBOUNCE_FAILED', 'true');
  }
}

/** Wird vom Debounce-Timer aufgerufen. Löscht den Timer und leert den Pending-Puffer. */
function onDebounceTimer(): void {
  Logger.log('onDebounceTimer: Timer ausgelöst');
  for (const t of ScriptApp.getProjectTriggers()) {
    if (t.getHandlerFunction() === 'onDebounceTimer') ScriptApp.deleteTrigger(t);
  }
  try {
    flushPendingChanges();
  } catch (e) {
    Logger.log(`onDebounceTimer error: ${e}`);
  }
}

// ─── Pending-Puffer verarbeiten und Log schreiben ──────────────────────────

/**
 * Zentrale Funktion, die beim Ablauf des Debounce-Timers aufgerufen wird.
 *
 * Ablauf:
 *   1. Lädt PENDING_EDITS, alle Snapshots, ABW_AFFECTED aus ScriptProperties
 *   2. Computed via Snapshot-Diff die Änderungen für Abw, Spieler und Saison
 *   3. Alle Einträge werden ins Änderungslog-Blatt geschrieben (permanentes Log)
 *   4. Speichert neue Snapshots für den nächsten Debounce-Durchlauf
 *   5. Sendet eine Mail-Benachrichtigung (in-memory, kein Sheet-Readback)
 *
 * Wichtig: Die Mail wird auch dann gesendet, wenn ABW_AFFECTED einen
 * Status-Revert (Final→Geplant) verzeichnet, selbst wenn die Snapshot-Diffs
 * keine Log-Einträge geliefert haben. So werden Kapitän und betroffene
 * Spieler immer über Aufstellungs-Änderungen durch Abwesenheiten informiert.
 */
function flushPendingChanges(): void {
  const props = PropertiesService.getScriptProperties();

  // ── 1. Alle relevanten Properties laden ──
  const pendingRaw = props.getProperty('PENDING_EDITS') || '[]';
  const abwSnapshotRaw = props.getProperty('ABW_SNAPSHOT');
  const spielerSnapshotRaw = props.getProperty('SPIELER_SNAPSHOT');
  const saisonSnapshotRaw = props.getProperty('SAISON_SNAPSHOT');
  const abwAffectedRaw = props.getProperty('ABW_AFFECTED');

  Logger.log(`flushPendingChanges: gestartet | PENDING_EDITS=${pendingRaw.length} Zch | SNAP=${!!saisonSnapshotRaw} | ABW_AFFECTED=${!!abwAffectedRaw}`);

  // Sofort löschen, damit bei einem Crash keine veralteten Daten liegen bleiben
  props.deleteProperty('PENDING_EDITS');
  props.deleteProperty('ABW_SNAPSHOT');
  props.deleteProperty('SPIELER_SNAPSHOT');
  props.deleteProperty('SAISON_SNAPSHOT');
  props.deleteProperty('ABW_AFFECTED');

  const pendingEdits: ChangeEntry[] = JSON.parse(pendingRaw);
  const hasAnyPending = pendingEdits.length > 0 || abwSnapshotRaw || spielerSnapshotRaw || saisonSnapshotRaw;
  if (!hasAnyPending) {
    Logger.log('flushPendingChanges: kein Pending + keine Snapshots → return');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logEntries: ChangeEntry[] = [];
  const bearbeiter = pendingEdits.length > 0 ? pendingEdits[0].bearbeiter : 'Unbekannt';

  // ── 2. Abwesenheiten-Diff via Snapshot ──
  if (abwSnapshotRaw) {
    const snapshot: SheetRowSnapshot[] = JSON.parse(abwSnapshotRaw);
    const current = readAbwSnapshot(ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN)!);
    computeSheetDiff(snapshot, current, SHEET_NAMES.ABWESENHEITEN, bearbeiter, logEntries);
  }

  // ── 3. Spieler-Diff via Snapshot ──
  if (spielerSnapshotRaw) {
    const snapshot: SheetRowSnapshot[] = JSON.parse(spielerSnapshotRaw);
    const current = readSpielerSnapshot(ss.getSheetByName(SHEET_NAMES.SPIELER)!);
    computeSheetDiff(snapshot, current, SHEET_NAMES.SPIELER, bearbeiter, logEntries);
  }

  // ── 4. Saison-Diff via Snapshot ──
  let saisonDiffs: SaisonDiffResult | null = null;
  if (saisonSnapshotRaw) {
    const snapshot: SaisonRowSnapshot[] = JSON.parse(saisonSnapshotRaw);
    const current = readSaisonSnapshot(ss.getSheetByName(SHEET_NAMES.SAISON)!);
    saisonDiffs = computeSaisonDiff(snapshot, current);
    // Saison-Änderungen als zeilenweise Log-Einträge schreiben
    for (const mod of saisonDiffs.modified) {
      logEntries.push({
        timestamp: Date.now(),
        sheetName: SHEET_NAMES.SAISON,
        rangeA1: mod.datum,
        alterWert: formatSaisonRowForLog(mod.oldRow),
        neuerWert: formatSaisonRowForLog(mod.newRow),
        bearbeiter,
      });
    }
    for (const n of saisonDiffs.added) {
      logEntries.push({
        timestamp: Date.now(),
        sheetName: SHEET_NAMES.SAISON,
        rangeA1: n.datum,
        alterWert: '(neuer Spieltag)',
        neuerWert: n.gegner,
        bearbeiter,
      });
    }
    for (const d of saisonDiffs.deleted) {
      logEntries.push({
        timestamp: Date.now(),
        sheetName: SHEET_NAMES.SAISON,
        rangeA1: d.datum,
        alterWert: d.gegner,
        neuerWert: '(gelöscht)',
        bearbeiter,
      });
    }
  }

  // ── 5. Ins Log-Blatt schreiben (permanent) ──
  if (logEntries.length > 0) {
    for (const entry of logEntries) {
      logChange(entry);
    }
  }

  // ── 6. Neue Snapshots für den nächsten Debounce ablegen ──
  saveSheetSnapshots(ss);

  // ── 7. Mail-Benachrichtigung (in-memory, kein Sheet-Readback) ──

  // ABW_AFFECTED vor dem Early-Return laden, damit Status-Reverts
  // (Final→Geplant) auch dann gemeldet werden, wenn die Snapshot-Diffs
  // keine Log-Einträge produziert haben.
  let abwAffected: AbwAffectedEntry[] | null = null;
  if (abwAffectedRaw) {
    abwAffected = JSON.parse(abwAffectedRaw);
    if (abwAffected && abwAffected.length === 0) abwAffected = null;
  }

  // Nachkontrolle: Wurde der Spieltag während der Debounce-Zeit repariert
  // (Status wieder auf Final) → keine "Nachplanung nötig"-Meldung mehr
  if (abwAffected && saisonDiffs) {
    const currentByKey = new Map<string, SaisonRowSnapshot>();
    for (const m of saisonDiffs.modified) {
      currentByKey.set(`${m.datum}|${m.gegner}`, m.newRow);
    }
    for (const a of abwAffected) {
      if (a.warFinal) {
        const cur = currentByKey.get(`${a.datumStr}|${a.gegner}`);
        if (cur && cur.status === 'Final') {
          a.warFinal = false;
          Logger.log(`flushPendingChanges: ${a.datumStr} ${a.gegner} wieder Final → warFinal zurückgesetzt`);
        }
      }
    }
  }

  const hasAffectedFinals = abwAffected ? abwAffected.some(a => a.warFinal) : false;

  if (logEntries.length === 0 && !hasAffectedFinals) {
    Logger.log('flushPendingChanges: keine Änderungen');
    return;
  }

  const saisonMod = saisonDiffs?.modified.length ?? 0;
  const saisonAdd = saisonDiffs?.added.filter(a => a.status === 'Final').length ?? 0;
  const saisonDel = saisonDiffs?.deleted.length ?? 0;
  const skippedAdd = (saisonDiffs?.added.length ?? 0) - saisonAdd;
  const abwCount = logEntries.filter(e => e.sheetName === SHEET_NAMES.ABWESENHEITEN).length;
  const spCount = logEntries.filter(e => e.sheetName === SHEET_NAMES.SPIELER).length;
  Logger.log(`flushPendingChanges: ${logEntries.length} Einträge ` +
    `(Saison: ${saisonMod} mod, ${saisonAdd} neu, ${saisonDel} del, ${skippedAdd} Geplant übersprungen | ` +
    `Abw: ${abwCount} | Spieler: ${spCount})` +
    (hasAffectedFinals ? ` | Final→Geplant Reverts: ja` : ''));

  sendChangeNotification(logEntries, abwAffected, saisonDiffs);
}

/** Speichert den aktuellen Stand von Abwesenheiten-, Spieler- und Saison-Sheet für den nächsten Diff. */
function saveSheetSnapshots(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): void {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('ABW_SNAPSHOT', JSON.stringify(readAbwSnapshot(ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN)!)));
  props.setProperty('SPIELER_SNAPSHOT', JSON.stringify(readSpielerSnapshot(ss.getSheetByName(SHEET_NAMES.SPIELER)!)));
  props.setProperty('SAISON_SNAPSHOT', JSON.stringify(readSaisonSnapshot(ss.getSheetByName(SHEET_NAMES.SAISON)!)));
}

// ─── Snapshot-Diff (generisch für Abwesenheiten + Spieler) ─────────────────

/**
 * Vergleicht zwei Snapshot-Arrays und produziert ChangeEntry-Objekte.
 *
 * Algorithmus:
 *   - Einträge mit gleichem Key (gleicher Inhalt) werden gepaart und verworfen
 *     → Cut+Paste erzeugt keinen Log-Eintrag
 *   - Ungepaarte Einträge werden als "gelöscht" / "neu" klassifiziert
 *   - Fallback: gleiche Zeilennummer, anderer Key → "modifiziert" (Alt → Neu)
 */
function computeSheetDiff(
  snapshot: SheetRowSnapshot[],
  current: SheetRowSnapshot[],
  sheetName: string,
  bearbeiter: string,
  out: ChangeEntry[]
): void {
  const now = Date.now();
  const snapByKey = new Map<string, SheetRowSnapshot[]>();
  const snapByRow = new Map<number, SheetRowSnapshot>();
  for (const r of snapshot) {
    if (!snapByKey.has(r.key)) snapByKey.set(r.key, []);
    snapByKey.get(r.key)!.push(r);
    snapByRow.set(r.row, r);
  }

  const currByKey = new Map<string, SheetRowSnapshot[]>();
  const currByRow = new Map<number, SheetRowSnapshot>();
  for (const r of current) {
    if (!currByKey.has(r.key)) currByKey.set(r.key, []);
    currByKey.get(r.key)!.push(r);
    currByRow.set(r.row, r);
  }

  const allKeys = new Set([...snapByKey.keys(), ...currByKey.keys()]);

  const unmatchedSnap: SheetRowSnapshot[] = [];
  const unmatchedCurr: SheetRowSnapshot[] = [];

  // Key-basierte Paarung
  for (const key of allKeys) {
    const snapRows = snapByKey.get(key) || [];
    const currRows = currByKey.get(key) || [];

    while (snapRows.length > 0 && currRows.length > 0) {
      snapRows.shift();
      currRows.shift();
    }

    for (const s of snapRows) unmatchedSnap.push(s);
    for (const c of currRows) unmatchedCurr.push(c);
  }

  // Zeilenbasierter Fallback: selbe Zeile, anderer Inhalt → modifiziert
  for (const s of [...unmatchedSnap]) {
    const sc = currByRow.get(s.row);
    if (sc && unmatchedCurr.includes(sc)) {
      unmatchedSnap.splice(unmatchedSnap.indexOf(s), 1);
      unmatchedCurr.splice(unmatchedCurr.indexOf(sc), 1);
      out.push({
        timestamp: now,
        sheetName,
        rangeA1: `A${s.row}`,
        alterWert: s.display,
        neuerWert: sc.display,
        bearbeiter,
      });
    }
  }

  // Übrige → gelöscht
  for (const s of unmatchedSnap) {
    out.push({
      timestamp: now,
      sheetName,
      rangeA1: `A${s.row}`,
      alterWert: s.display,
      neuerWert: '(gelöscht)',
      bearbeiter,
    });
  }

  // Übrige → neu
  for (const c of unmatchedCurr) {
    out.push({
      timestamp: now,
      sheetName,
      rangeA1: `A${c.row}`,
      alterWert: '(neu)',
      neuerWert: c.display,
      bearbeiter,
    });
  }
}

// ─── Saison-Diff (für Mail-Darstellung) ─────────────────────────────────────

interface ChangedCell {
  label: string;
  oldVal: string;
  newVal: string;
}

interface SaisonModifiedRow {
  datum: string;
  gegner: string;
  oldRow: SaisonRowSnapshot;
  newRow: SaisonRowSnapshot;
  changedCells: ChangedCell[];
}

interface SaisonDiffResult {
  modified: SaisonModifiedRow[];
  added: SaisonRowSnapshot[];
  deleted: SaisonRowSnapshot[];
}

/**
 * Vergleicht zwei Saison-Snapshots (Spieltag-Zeilen) und kategorisiert
 * die Änderungen in modified, added, deleted.
 * Matching erfolgt über Schlüssel `datum|gegner` (robust gegen Zeilenverschiebungen).
 */
function computeSaisonDiff(snapshot: SaisonRowSnapshot[], current: SaisonRowSnapshot[]): SaisonDiffResult {
  function key(r: SaisonRowSnapshot) { return `${r.datum}|${r.gegner}`; }

  const oldByKey = new Map<string, SaisonRowSnapshot>();
  for (const r of snapshot) oldByKey.set(key(r), r);
  const newByKey = new Map<string, SaisonRowSnapshot>();
  for (const r of current) newByKey.set(key(r), r);

  const allKeys = new Set([...oldByKey.keys(), ...newByKey.keys()]);
  const modified: SaisonModifiedRow[] = [];
  const added: SaisonRowSnapshot[] = [];
  const deleted: SaisonRowSnapshot[] = [];

  for (const k of allKeys) {
    const oldR = oldByKey.get(k);
    const newR = newByKey.get(k);

    if (oldR && newR) {
      const changes = diffSaisonRow(oldR, newR);
      if (changes.length > 0) {
        modified.push({ datum: oldR.datum, gegner: oldR.gegner, oldRow: oldR, newRow: newR, changedCells: changes });
      }
    } else if (!oldR && newR) {
      added.push(newR);
    } else if (oldR && !newR) {
      deleted.push(oldR);
    }
  }

  return { modified, added, deleted };
}

function diffSaisonRow(oldR: SaisonRowSnapshot, newR: SaisonRowSnapshot): ChangedCell[] {
  const changes: ChangedCell[] = [];

  if (oldR.gegner !== newR.gegner) changes.push({ label: 'Gegner', oldVal: oldR.gegner, newVal: newR.gegner });
  if (oldR.startzeit !== newR.startzeit) changes.push({ label: 'Startzeit', oldVal: oldR.startzeit, newVal: newR.startzeit });
  if (oldR.heimAuswaerts !== newR.heimAuswaerts) changes.push({ label: 'Ort', oldVal: oldR.heimAuswaerts, newVal: newR.heimAuswaerts });

  for (let pi = 0; pi < SHEET_CONFIG.spieler.length; pi++) {
    const ov = oldR.playerAssignments[pi];
    const nv = newR.playerAssignments[pi];
    if (ov !== nv) changes.push({ label: SHEET_CONFIG.spieler[pi].name, oldVal: ov, newVal: nv });
  }

  for (let ei = 0; ei < 3; ei++) {
    const ov = oldR.ersatz[ei];
    const nv = newR.ersatz[ei];
    if (ov !== nv) changes.push({ label: `Ersatz ${ei + 1}`, oldVal: ov, newVal: nv });
  }

  if (oldR.status !== newR.status) changes.push({ label: 'Status', oldVal: oldR.status, newVal: newR.status });
  if (oldR.kommentar !== newR.kommentar) changes.push({ label: 'Kommentar', oldVal: oldR.kommentar, newVal: newR.kommentar });

  return changes;
}

function formatSaisonRowForLog(row: SaisonRowSnapshot): string {
  const parts: string[] = [];
  if (row.gegner) parts.push(row.gegner);
  if (row.startzeit) parts.push(row.startzeit);
  if (row.heimAuswaerts) parts.push(row.heimAuswaerts);
  for (let pi = 0; pi < row.playerAssignments.length; pi++) {
    const v = row.playerAssignments[pi];
    if (v) parts.push(`${SHEET_CONFIG.spieler[pi].name}: ${v}`);
  }
  for (let ei = 0; ei < row.ersatz.length; ei++) {
    if (row.ersatz[ei]) parts.push(`Ersatz ${ei + 1}: ${row.ersatz[ei]}`);
  }
  parts.push(`Status: ${row.status}`);
  if (row.kommentar) parts.push(`Kommentar: ${row.kommentar}`);
  return parts.join(' | ');
}

// ─── Saison-Validierung (Validierungsspalte) ───────────────────────────────────

function computeValidierung(
  playerRow: string[],
  datum: Date,
  spieler: Spieler[],
  allAbw: Map<string, Map<string, string>>,
  ersatzVals: string[]
): string {
  const key = dateKey(datum);
  const dayMap = allAbw.get(key);
  let einzel = 0;
  let doppel = 0;
  const warnungen: string[] = [];

  for (let pi = 0; pi < playerRow.length; pi++) {
    const val = String(playerRow[pi] || '').trim();
    if (!val || val.startsWith('✗')) continue;

    if (val === 'Einzel+Doppel') { einzel++; doppel++; }
    else if (val === 'Einzel') einzel++;
    else if (val === 'Doppel') doppel++;

    const name = spieler[pi].name;
    const abwDisplay = dayMap?.get(name);
    if (abwDisplay) warnungen.push(`${name}: ${abwDisplay}`);
  }

  let ersatz = 0;
  for (const e of ersatzVals) {
    if (String(e || '').trim()) ersatz++;
  }

  const f = SHEET_CONFIG.einstellungen.spielformat;
  let missingEinzel = Math.max(0, f.einzel - einzel);
  let missingDoppel = Math.max(0, f.doppel * 2 - doppel);
  for (let i = 0; i < ersatz; i++) {
    if (missingEinzel > 0 && missingDoppel > 0) { einzel++; doppel++; missingEinzel--; missingDoppel--; }
    else if (missingEinzel > 0) { einzel++; missingEinzel--; }
    else if (missingDoppel > 0) { doppel++; missingDoppel--; }
    else { einzel++; doppel++; }
  }

  const gesamt = einzel > doppel ? einzel : doppel;
  if (gesamt > 6) warnungen.push(`Mehr als 6 Spieler aufgestellt (${gesamt})`);
  if (einzel < f.einzel) warnungen.push(`Nur ${einzel}/${f.einzel} Einzel-Spieler`);
  if (einzel > f.einzel) warnungen.push(`Mehr als ${f.einzel} Einzel-Spieler (${einzel})`);
  if (doppel < f.doppel * 2) warnungen.push(`Nur ${doppel}/${f.doppel * 2} Doppel-Spieler`);
  if (doppel > f.doppel * 2) warnungen.push(`Mehr als ${f.doppel * 2} Doppel-Spieler (${doppel})`);

  return warnungen.join(' | ');
}

function validateAndUpdateSaisonRow(editedRange: GoogleAppsScript.Spreadsheet.Range): void {
  const row = editedRange.getRow();
  const sheet = editedRange.getSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const gegner = String(sheet.getRange(row, saisonGegnerCol()).getValue() || '').trim();
  if (!gegner) {
    sheet.getRange(row, saisonValidierungCol()).setValue('');
    return;
  }

  const datum = toDate(sheet.getRange(row, saisonDatumCol()).getValue());
  if (!datum) return;

  const allAbw = buildAbwesenheitenIndex(ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN)!);
  const numPlayers = SHEET_CONFIG.spieler.length;
  const playerRow = sheet.getRange(row, saisonSpielerCol(0), 1, numPlayers).getValues()[0] as string[];
  const allSpieler: Spieler[] = SHEET_CONFIG.spieler.map(s => ({
    name: s.name, email: '', rang: s.rang, aenderungenMelden: false, rolle: ''
  }));
  const ersatzVals = [
    String(sheet.getRange(row, saisonErsatzCol(0)).getValue() || ''),
    String(sheet.getRange(row, saisonErsatzCol(1)).getValue() || ''),
    String(sheet.getRange(row, saisonErsatzCol(2)).getValue() || ''),
  ];
  const validierung = computeValidierung(playerRow, datum, allSpieler, allAbw, ersatzVals);
  sheet.getRange(row, saisonValidierungCol()).setValue(validierung);
}

// ─── Änderungslog-Blatt (unverändert) ──────────────────────────────────────

/**
 * Schreibt einen ChangeEntry ins Änderungslog-Blatt.
 * Fügt sheetName und rangeA1 zu einem Bereichs-String zusammen
 * (identisch mit dem bisherigen Log-Format, rückwärtskompatibel).
 */
function logChange(entry: ChangeEntry): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAMES.AENDERUNGSLOG);
  if (!logSheet) return;

  logSheet.showSheet();
  const lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow + 1, COL_AENDERUNGSLOG.Zeitstempel, 1, 5).setValues([[
    new Date(entry.timestamp),
    `${entry.sheetName}!${entry.rangeA1}`,
    entry.alterWert,
    entry.neuerWert,
    entry.bearbeiter,
  ]]);
  logSheet.hideSheet();
}

// ─── Mail-Benachrichtigung ─────────────────────────────────────────────────

/**
 * Baut die Ohne-Email-Sektion für den Kapitän.
 *
 * Iteriert über saisonDiffs und sammelt alle Spieler ohne hinterlegte
 * E-Mail-Adresse, deren Aufstellung sich geändert hat. Die Tabelle zeigt
 * Datum/Gegner, Spielername und die neue Einsatzart (bzw. „-" wenn der
 * Spieler nicht mehr aufgestellt ist).
 *
 * Ist als eigenständige Funktion ohne Abhängigkeit zum Mail-Versand
 * implementiert, um später problemlos auf eine separate Mail umstellbar zu sein.
 *
 * Gibt einen leeren String zurück, wenn keine Ohne-Email-Änderungen vorliegen.
 */
function buildOhneEmailSektion(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, saisonDiffs: SaisonDiffResult): string {
  const ohneEmailNamen = new Set<string>();
  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (spielerSheet) {
    const lr = spielerSheet.getLastRow();
    const d = spielerSheet.getRange(2, 1, lr - 1, COL_SPIELER.Rolle).getValues();
    for (const r of d) {
      const name = String(r[COL_SPIELER.Name - 1] || '').trim();
      if (!name) continue;
      const email = String(r[COL_SPIELER.Email - 1]).trim();
      if (!email || !email.includes('@')) ohneEmailNamen.add(name);
    }
  }
  if (ohneEmailNamen.size === 0) return '';

  interface OhneEmailRow {
    spieltag: string;
    spieler: string;
    einsatz: string;
  }
  const rows: OhneEmailRow[] = [];
  const spieltagKey = (r: SaisonRowSnapshot) => `${r.datum} — ${r.gegner}`;

  const addRow = (r: SaisonRowSnapshot, name: string, einsatz: string) => {
    if (ohneEmailNamen.has(name)) {
      rows.push({ spieltag: spieltagKey(r), spieler: name, einsatz });
    }
  };

  for (const mod of saisonDiffs.modified) {
    for (let pi = 0; pi < SHEET_CONFIG.spieler.length; pi++) {
      const ov = mod.oldRow.playerAssignments[pi];
      const nv = mod.newRow.playerAssignments[pi];
      if (ov === nv) continue;
      const name = SHEET_CONFIG.spieler[pi].name;
      if (nv && !nv.startsWith('✗')) {
        addRow(mod.newRow, name, nv);
      } else {
        addRow(mod.oldRow, name, '-');
      }
    }
  }
  for (const a of saisonDiffs.added) {
    for (let pi = 0; pi < SHEET_CONFIG.spieler.length; pi++) {
      const nv = a.playerAssignments[pi];
      if (nv && !nv.startsWith('✗')) {
        addRow(a, SHEET_CONFIG.spieler[pi].name, nv);
      }
    }
  }
  for (const d of saisonDiffs.deleted) {
    for (let pi = 0; pi < SHEET_CONFIG.spieler.length; pi++) {
      const ov = d.playerAssignments[pi];
      if (ov && !ov.startsWith('✗')) {
        addRow(d, SHEET_CONFIG.spieler[pi].name, '-');
      }
    }
  }

  if (rows.length === 0) return '';

  const sortKey = (d: string) => {
    const p = d.split('.');
    return p.length === 3 ? p[2] + p[1] + p[0] : d;
  };
  rows.sort((a, b) => sortKey(a.spieltag.split(' — ')[0]).localeCompare(sortKey(b.spieltag.split(' — ')[0])));

  const s = emailStyles();
  let html = '<p style="margin-top:20px"><b>Ohne E-Mail-Adresse:</b><br>Folgende Spieler konnten nicht per E-Mail informiert werden:</p>';
  html += '<table border="1" cellpadding="4" cellspacing="0" style="font-size:13px;border-collapse:collapse">';
  html += '<tr style="background:#4A90D9;color:white"><th>Spieltag</th><th>Spieler</th><th>Einsatz</th></tr>';
  for (const r of rows) {
    html += `<tr><td style="padding:2px 6px">${escapeHtml(r.spieltag)}</td><td style="padding:2px 6px">${escapeHtml(r.spieler)}</td><td style="padding:2px 6px">${escapeHtml(r.einsatz)}</td></tr>`;
  }
  html += '</table>';

  return html;
}

/**
 * Baut eine HTML-Änderungsmail und verschickt sie an alle Empfänger.
 *
 * Empfänger-Kreis (in getAenderungenMeldenEmpfaenger):
 *   - Kapitän (Rolle "Kapitän") immer – unabhängig von der Checkbox
 *   - Spieler, deren eigener Einsatz geändert wurde (auch durch
 *     Abwesenheits-Revert Final→Geplant) – unabhängig von der Checkbox
 *   - Alle Spieler mit gesetzter Checkbox "Aufstellungsänderungen melden"
 *   - Ausnahme: der Bearbeiter selbst wird nicht benachrichtigt
 *
 * Darstellung:
 *   - Saison-Änderungen nach Spieltag gruppiert mit Zeilenkontext
 *     (Gegner, Datum, Ort)
 *   - Abwesenheits-Änderungen mit Spieltags-Bezug, angereichert
 *     mit warAufgestellt/warFinal-Kontext („Final → Geplant“-Hinweis)
 *
 * Die Mail wird auch dann gesendet, wenn nur ein Abwesenheits-Revert
 * (Final→Geplant) vorliegt, aber keine Snapshot-Diff-Einträge – etwa
 * beim ersten Edit oder wenn nur der Status revertiert wurde.
 */
function sendChangeNotification(
  entries: ChangeEntry[],
  abwAffected: AbwAffectedEntry[] | null,
  saisonDiffs: SaisonDiffResult | null
): void {
  const props = PropertiesService.getScriptProperties();
  const suppress = props.getProperty('SUPPRESS_NOTIFICATION') === 'true';

  // SUPPRESS_NOTIFICATION: Nach "Finalisieren + Senden" keine redundante
  // Änderungsmail. Ohne-Email-Info wird bereits synchron in
  // menuFinalisierenUndSenden per sendOhneEmailFinalMail versendet.
  if (suppress) {
    Logger.log('sendChangeNotification: SUPPRESS_NOTIFICATION → unterdrückt');
    props.deleteProperty('SUPPRESS_NOTIFICATION');
    return;
  }

  // Nur relevant wenn sichtbare Saison- oder Abw-Änderungen vorliegen.
  // Zusätzlich: Mail auch senden, wenn ein Abwesenheits-Revert (Final→Geplant)
  // stattfand, selbst wenn der Saison-Diff oder Abw-Snapshot-Diff leer ist.
  const visibleAdded = saisonDiffs ? saisonDiffs.added.filter(a => a.status === 'Final').length : 0;
  const visibleSaisonCount = (saisonDiffs?.modified.length ?? 0) + visibleAdded + (saisonDiffs?.deleted.length ?? 0);
  const showAbw = abwAffected && abwAffected.length > 0;
  const abwEntries = showAbw ? entries.filter(e => e.sheetName === SHEET_NAMES.ABWESENHEITEN) : [];
  const hasAffectedFinals = abwAffected ? abwAffected.some(a => a.warFinal) : false;
  Logger.log(`sendChangeNotification: Saison=${visibleSaisonCount} | AbwEntries=${abwEntries.length} | AbwAffected=${abwAffected?.length ?? 0} | hasAffectedFinals=${hasAffectedFinals}`);
  if (visibleSaisonCount === 0 && abwEntries.length === 0 && !hasAffectedFinals) {
    Logger.log('sendChangeNotification: keine sichtbaren Änderungen → return');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Betroffene Spieler aus dem Saison-Diff und aus Abwesenheits-Reverts ermitteln
  const affectedNames = new Set<string>();
  if (saisonDiffs) {
    for (const mod of saisonDiffs.modified) {
      for (const c of mod.changedCells) {
        if (SHEET_CONFIG.spieler.some(s => s.name === c.label)) affectedNames.add(c.label);
      }
    }
    for (const a of saisonDiffs.added) {
      for (let pi = 0; pi < a.playerAssignments.length; pi++) {
        if (a.playerAssignments[pi] && !a.playerAssignments[pi].startsWith('✗')) {
          affectedNames.add(SHEET_CONFIG.spieler[pi].name);
        }
      }
      for (const n of a.ersatz) { if (n) affectedNames.add(n); }
    }
    for (const d of saisonDiffs.deleted) {
      for (let pi = 0; pi < d.playerAssignments.length; pi++) {
        if (d.playerAssignments[pi] && !d.playerAssignments[pi].startsWith('✗')) {
          affectedNames.add(SHEET_CONFIG.spieler[pi].name);
        }
      }
      for (const n of d.ersatz) { if (n) affectedNames.add(n); }
    }
  }
  // Zusätzlich: Spieler, deren Abwesenheit einen Status-Revert (Final→Geplant)
  // ausgelöst hat, unabhängig davon ob der Saison-Diff die Änderung erfasst hat.
  if (abwAffected) {
    for (const a of abwAffected) {
      if (a.warAufgestellt) {
        affectedNames.add(a.spielerName);
      }
    }
  }

  Logger.log(`sendChangeNotification: affectedNames=[${[...affectedNames].join(', ')}]`);

  const empfaenger = getAenderungenMeldenEmpfaenger(ss, affectedNames);
  Logger.log(`sendChangeNotification: ${empfaenger.length} Empfänger — [${empfaenger.join(', ')}]`);
  if (empfaenger.length === 0) {
    Logger.log('sendChangeNotification: keine Empfänger → return');
    return;
  }

  const subject = `Einsatzplaner – Änderungen vom ${Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM.yyyy HH:mm')}`;
  const s = emailStyles();
  let html = `<html><body style="${s.body}">
<p>Hallo,</p>`;

  if (visibleSaisonCount > 0 && saisonDiffs) {
    html += visibleSaisonCount === 1
      ? `<p>folgender Spieltag wurde aktualisiert:</p>`
      : `<p>folgende <b>${visibleSaisonCount}</b> Spieltage wurden aktualisiert:</p>`;

    // ── Geänderte Spieltage (2-Zeilen-Tabelle: alt/rot, neu/grün) ──
    for (const mod of saisonDiffs.modified) {
      const changedSet = new Set(mod.changedCells.map(c => c.label));
      const fullCols = buildFullColumnList(mod.oldRow, mod.newRow, changedSet);
      html += saisonRowTableHtml('bearbeitet', `${mod.datum} — ${mod.gegner}`, fullCols);
    }

    // ── Neue Spieltage (nur Final, nicht nur Geplant) ──
    for (const n of saisonDiffs.added) {
      if (n.status !== 'Final') continue;
      const cols = buildColumnListFromRow(n);
      html += saisonSingleRowTableHtml('neu', `${n.datum} — ${n.gegner}`, cols, true);
    }

    // ── Gelöschte Spieltage ──
    for (const d of saisonDiffs.deleted) {
      const cols = buildColumnListFromRow(d);
      html += saisonSingleRowTableHtml('gelöscht', `${d.datum} — ${d.gegner}`, cols, false);
    }
  }

  // ── Abwesenheits-Änderungen ──
  // Abwesenheits-sektion anzeigen wenn es entweder Snapshot-Diff-Einträge gibt
  // oder ein Status-Revert (Final→Geplant) durch eine Abwesenheit ausgelöst wurde.
  if (showAbw && (abwEntries.length > 0 || abwAffected!.some(a => a.warFinal))) {
    if (visibleSaisonCount > 0) html += `<p style="margin-top:20px">Abwesenheits-Änderungen mit Spieltags-Bezug:</p>`;
    else html += `<p>Abwesenheits-Änderungen mit Spieltags-Bezug:</p>`;

    html += `<table cellpadding="4" cellspacing="0" style="${s.table}">`;
    for (const e of abwEntries) {
      html += `<tr><td style="${s.red};padding-right:8px">${escapeHtml(e.alterWert)}</td><td>→</td><td style="${s.green};padding-left:8px">${escapeHtml(e.neuerWert)}</td></tr>`;
    }
    html += `</table>`;

    for (const a of abwAffected!) {
      const context = [a.datumStr, a.startzeit, a.heimAuswaerts, a.gegner ? `— ${a.gegner}` : ''].filter(Boolean).join(' ');
      let status: string;
      if (a.warAufgestellt) {
        if (a.warFinal) {
          // Status ist immer noch Geplant (wurde nicht korrigiert)
          if (a.validierung) {
            status = `War aufgestellt (Final → Geplant). Validierung: ${a.validierung}`;
          } else {
            status = 'War aufgestellt (Final → Geplant) — Bitte Spieltag prüfen und Status auf Final setzen.';
          }
        } else {
          // Tag wurde korrigiert oder war nie Final
          status = 'War aufgestellt.';
        }
      } else {
        status = 'Stand als Ersatz zur Verfügung';
      }
      html += `<p style="margin:2px 0 2px 20px;font-size:13px">${escapeHtml(context)}<br><span style="${s.red}">${escapeHtml(a.spielerName)}: ${status}</span></p>`;
    }
  }

  // ── Ohne-Email-Sektion (nur für den Kapitän) ──
  const kapEmail = getKapitaenEmail(ss);
  const ohneHtml = saisonDiffs ? buildOhneEmailSektion(ss, saisonDiffs) : '';

  const footerHtml = `<p style="${s.footer}">Viele Grüße,<br>Dein Einsatzplaner-Team</p></body></html>`;

  for (const email of empfaenger) {
    const isKapitaen = email === kapEmail;
    const body = isKapitaen && ohneHtml
      ? html + ohneHtml + footerHtml
      : html + footerHtml;
    MailApp.sendEmail({ to: email, subject, htmlBody: body });
  }
  Logger.log(`sendChangeNotification: Mail an ${empfaenger.length} Empfänger (${empfaenger.join(', ')})` +
    (ohneHtml ? ' — mit Ohne-Email-Sektion für Kapitän' : ''));
}

// ─── Saison-Tabellen-Builder für die Mail ───────────────────────────────────

function buildColumnListFromRow(row: SaisonRowSnapshot): { label: string; oldVal: string; newVal: string }[] {
  const cols: { label: string; oldVal: string; newVal: string }[] = [];
  if (row.gegner) cols.push({ label: 'Gegner', oldVal: '', newVal: row.gegner });
  if (row.startzeit) cols.push({ label: 'Startzeit', oldVal: '', newVal: row.startzeit });
  if (row.heimAuswaerts) cols.push({ label: 'Ort', oldVal: '', newVal: row.heimAuswaerts });
  for (let pi = 0; pi < row.playerAssignments.length; pi++) {
    const v = row.playerAssignments[pi];
    if (v && !v.startsWith('✗')) cols.push({ label: SHEET_CONFIG.spieler[pi].name, oldVal: '', newVal: v });
  }
  for (let ei = 0; ei < 3; ei++) {
    const v = row.ersatz[ei];
    if (v) cols.push({ label: `Ersatz ${ei + 1}`, oldVal: '', newVal: v });
  }
  cols.push({ label: 'Status', oldVal: '', newVal: row.status });
  if (row.kommentar) cols.push({ label: 'Kommentar', oldVal: '', newVal: row.kommentar });
  return cols;
}

/** Baut die vollständige Spaltenliste für einen geänderten Spieltag auf (alle Spalten, nicht nur geänderte). */
function buildFullColumnList(oldR: SaisonRowSnapshot, newR: SaisonRowSnapshot, changedSet: Set<string>): { label: string; oldVal: string; newVal: string }[] {
  const cols: { label: string; oldVal: string; newVal: string }[] = [];
  cols.push({ label: 'Gegner', oldVal: oldR.gegner, newVal: newR.gegner });
  if (oldR.startzeit || newR.startzeit) cols.push({ label: 'Startzeit', oldVal: oldR.startzeit, newVal: newR.startzeit });
  cols.push({ label: 'Ort', oldVal: oldR.heimAuswaerts, newVal: newR.heimAuswaerts });
  for (let pi = 0; pi < SHEET_CONFIG.spieler.length; pi++) {
    const ov = oldR.playerAssignments[pi];
    const nv = newR.playerAssignments[pi];
    const name = SHEET_CONFIG.spieler[pi].name;
    if (ov && !ov.startsWith('✗') || nv && !nv.startsWith('✗') || changedSet.has(name)) {
      cols.push({ label: name, oldVal: ov, newVal: nv });
    }
  }
  for (let ei = 0; ei < 3; ei++) {
    const ov = oldR.ersatz[ei];
    const nv = newR.ersatz[ei];
    if (ov || nv || changedSet.has(`Ersatz ${ei + 1}`)) {
      cols.push({ label: `Ersatz ${ei + 1}`, oldVal: ov, newVal: nv });
    }
  }
  cols.push({ label: 'Status', oldVal: oldR.status, newVal: newR.status });
  if (oldR.kommentar || newR.kommentar || changedSet.has('Kommentar')) {
    cols.push({ label: 'Kommentar', oldVal: oldR.kommentar, newVal: newR.kommentar });
  }
  return cols;
}

/** 2-Zeilen-Tabelle für geänderte Spieltage: alt (rot) / neu (grün). */
function saisonRowTableHtml(typ: string, datum: string, cols: { label: string; oldVal: string; newVal: string }[]): string {
  const s = emailStyles();
  let h = `<p style="${s.header}">${escapeHtml(datum)} (${typ})</p>`;
  h += `<table border="1" cellpadding="4" cellspacing="0" style="${s.table}">`;
  h += `<tr>`;
  for (const c of cols) h += `<th style="${s.th}">${escapeHtml(c.label)}</th>`;
  h += '</tr><tr>';
  for (const c of cols) {
    const changed = c.oldVal !== c.newVal;
    h += `<td style="${changed ? s.red + ';' : ''}${s.td}">${escapeHtml(c.oldVal || '-')}</td>`;
  }
  h += '</tr><tr>';
  for (const c of cols) {
    const changed = c.oldVal !== c.newVal;
    h += `<td style="${changed ? s.green + ';' : ''}${s.td}">${escapeHtml(c.newVal || '-')}</td>`;
  }
  h += '</tr></table>';
  return h;
}

/** Einzeilige Tabelle für neue oder gelöschte Spieltage. */
function saisonSingleRowTableHtml(typ: string, datum: string, cols: { label: string; oldVal: string; newVal: string }[], isNew: boolean): string {
  const displayCols = cols.filter(c => (isNew && c.newVal) || (!isNew && c.oldVal));
  if (displayCols.length === 0) return '';

  const s = emailStyles();
  const color = isNew ? s.green : s.red;
  let h = `<p style="${s.header}">${escapeHtml(datum)} (${typ})</p>`;
  h += `<table border="1" cellpadding="4" cellspacing="0" style="${s.table}">`;
  h += '<tr>';
  for (const c of displayCols) h += `<th style="${s.th}">${escapeHtml(c.label)}</th>`;
  h += '</tr><tr>';
  for (const c of displayCols) {
    const v = isNew ? c.newVal : c.oldVal;
    h += `<td style="${color};${s.td}">${escapeHtml(v || '-')}</td>`;
  }
  h += '</tr></table>';
  return h;
}

/** HTML-Escaping für Mail-Inhalte. */
function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ─── Empfänger-Ermittlung ──────────────────────────────────────────────────

/**
 * Sammelt E-Mail-Adressen der Benachrichtigungs-Empfänger.
 *
 * Immer dabei:
 *   - Kapitän (Rolle "Kapitän" und gültige E-Mail)
 *   - Spieler, deren eigener Einsatz geändert wurde (unabhängig von Checkbox)
 *
 * Optional (wenn Checkbox gesetzt):
 *   - Alle Spieler mit "Aufstellungsänderungen melden" = TRUE
 *
 * Der aktuelle Bearbeiter wird von der eigenen Benachrichtigung
 * ausgeschlossen (gilt auch für den Kapitän).
 */
function getAenderungenMeldenEmpfaenger(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  affectedNames: Set<string>
): string[] {
  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (!spielerSheet) return [];

  const lastRow = spielerSheet.getLastRow();
  if (lastRow <= 1) return [];

  const currentUser = Session.getActiveUser().getEmail();
  const data = spielerSheet.getRange(2, 1, lastRow - 1, COL_SPIELER.Rolle).getValues();
  const empfaenger: string[] = [];
  let skippedNoEmail = 0;
  let skippedSelfEdit = 0;

  for (const row of data) {
    const name = String(row[COL_SPIELER.Name - 1] || '').trim();
    if (!name) continue; // Leere Zeilen (kein Spieler eingetragen) überspringen

    const email = String(row[COL_SPIELER.Email - 1]).trim();
    if (!email || !email.includes('@')) { skippedNoEmail++; continue; }
    if (email === currentUser) { skippedSelfEdit++; continue; }

    const rolle = String(row[COL_SPIELER.Rolle - 1] || '').trim();
    const isKapitaen = rolle === 'Kapitän';

    const meldenRaw = row[COL_SPIELER.AenderungenMelden - 1];
    const melden = meldenRaw === true
      || String(meldenRaw).toUpperCase() === 'TRUE'
      || meldenRaw === 1;

    const isAffected = affectedNames.has(name);

    if (isKapitaen || isAffected || melden) {
      const reason = isKapitaen ? 'Kapitän' : isAffected ? 'betroffen' : 'melden';
      Logger.log(`  include ${name} <${email}> — ${reason}`);
      if (!empfaenger.includes(email)) empfaenger.push(email);
    }
  }

  if (skippedNoEmail > 0 || skippedSelfEdit > 0) {
    Logger.log(`getAenderungenMeldenEmpfaenger: ${empfaenger.length} Empfänger` +
      ` (${skippedNoEmail} ohne Email, ${skippedSelfEdit} Selbst-Edit übersprungen)`);
  } else {
    Logger.log(`getAenderungenMeldenEmpfaenger: ${empfaenger.length} Empfänger`);
  }
  return empfaenger;
}
