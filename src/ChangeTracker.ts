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
  if (e.authMode !== ScriptApp.AuthMode.FULL) return;

  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SHEET_BUILDER_RUNNING') === 'true') return;

  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName === SHEET_NAMES.AENDERUNGSLOG || sheetName === SHEET_NAMES.DOKUMENTATION) return;

  const row = range.getRow();
  const col = range.getColumn();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

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
    rebuildAbwesenheitenInSaison(ss);
    restoreAbwesenheitenValidations(sheet);
    validateAbwesenheitRow(sheet, row);
  }

  // Spieler: Data-Validations (Checkboxen, Dropdowns) nach Cut+Paste wiederherstellen
  if (sheetName === SHEET_NAMES.SPIELER && row >= 2) {
    restoreSpielerValidations(sheet);
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

  const aktiveSpieler = readAktiveSpieler(ss.getSheetByName(SHEET_NAMES.SPIELER)!);
  if (!aktiveSpieler || aktiveSpieler.length === 0) return;

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
          if (!current || !current.startsWith('✗')) {
            const warAufgestellt = !!(current && !current.startsWith('✗'));
            if (warAufgestellt && wasFinalBefore) {
              statusValues[r][0] = 'Geplant';
              statusChanged = true;
            }
            allPlayerCols[r][pi] = abwDisplay;
            playerChanged = true;

            // Nur protokollieren wenn es ein Spieltag mit Gegner ist
            if (isSpieltag) {
              const datumStr = Utilities.formatDate(d, 'Europe/Berlin', 'dd.MM.yyyy');
              affected.push({
                datumStr,
                startzeit: formatStartzeit(startCols[r][0]),
                heimAuswaerts: heimAw,
                gegner,
                spielerName: spielerNames[pi],
                warAufgestellt,
                warFinal: wasFinalBefore && warAufgestellt,
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
    const validierungValues: string[][] = [];
    for (let r = 0; r < numRows; r++) {
      validierungValues.push([computeValidierung(allPlayerCols[r], dates[r], aktiveSpieler, refreshedAllAbw, saisonSheet, r + 2)]);
    }
    saisonSheet.getRange(2, saisonValidierungCol(), numRows, 1).setValues(validierungValues);
  }

  // Betroffene Spieltage für die spätere Mail-Anreicherung speichern
  const props = PropertiesService.getScriptProperties();
  props.setProperty('ABW_AFFECTED', JSON.stringify(affected));
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
    const von = toDateSafe(data[COL_ABWESENHEITEN.Von - 1]);
    const bis = toDateSafe(data[COL_ABWESENHEITEN.Bis - 1]);
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
  const von = toDateSafe(data[COL_ABWESENHEITEN.Von - 1]);
  const bis = toDateSafe(data[COL_ABWESENHEITEN.Bis - 1]);

  // Nur validieren wenn mindestens ein Feld gefüllt ist
  if (!spieler && !von && !bis) {
    sheet.getRange(row, COL_ABWESENHEITEN.Validierung).setValue('');
    return;
  }

  const fehler: string[] = [];

  if (!von || !bis) {
    fehler.push('Von- und Bis-Datum erforderlich');
  } else if (von.getTime() >= bis.getTime()) {
    fehler.push('Von-Datum muss vor dem Bis-Datum liegen');
  }

  sheet.getRange(row, COL_ABWESENHEITEN.Validierung).setValue(fehler.join(' | '));
}

// ─── Data-Validation nach Cut+Paste wiederherstellen ───────────────────────

/** Stellt Spieler-Dropdown und Datumsformat im Abwesenheiten-Sheet wieder her. */
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
}

/** Stellt Checkboxen und Dropdowns im Spieler-Sheet wieder her. */
function restoreSpielerValidations(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  const lastRow = Math.max(sheet.getMaxRows() - 1, 2);
  sheet.getRange(2, COL_SPIELER.Rang, lastRow, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberGreaterThan(0).setAllowInvalid(true).build());
  sheet.getRange(2, COL_SPIELER.Aktiv, lastRow, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().build());
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
    const von = toDateSafe(data[r][COL_ABWESENHEITEN.Von - 1]);
    const bis = toDateSafe(data[r][COL_ABWESENHEITEN.Bis - 1]);
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
 * Key = `name|email|rang|aktiv|melden|rolle`.
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
    const aktiv = data[r][COL_SPIELER.Aktiv - 1] === true ? 'ja' : '';
    const melden = data[r][COL_SPIELER.AenderungenMelden - 1] === true ? 'ja' : '';
    const rolle = String(data[r][COL_SPIELER.Rolle - 1] || '').trim();
    rows.push({
      row: r + 2,
      key: `${name}|${email}|${rang}|${aktiv}|${melden}|${rolle}`,
      display: `Name: ${name} | Email: ${email || '-'} | Rang: ${rang} | Aktiv: ${aktiv || 'nein'} | Melden: ${melden || 'nein'} | Rolle: ${rolle || '-'}`,
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
 */
function resetDebounceTimer(): void {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('SHEET_BUILDER_RUNNING') === 'true') return;
    if (props.getProperty('BULK_EDIT') === 'true') return;

    for (const t of ScriptApp.getProjectTriggers()) {
      if (t.getHandlerFunction() === 'onDebounceTimer') ScriptApp.deleteTrigger(t);
    }
    const minuten = SHEET_CONFIG.einstellungen.debounceMinuten;
    ScriptApp.newTrigger('onDebounceTimer')
      .timeBased().after(minuten * 60 * 1000).create();
  } catch (_) {
  }
}

/** Wird vom Debounce-Timer aufgerufen. Löscht den Timer und leert den Pending-Puffer. */
function onDebounceTimer(): void {
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
 */
function flushPendingChanges(): void {
  const props = PropertiesService.getScriptProperties();

  // ── 1. Alle relevanten Properties laden ──
  const pendingRaw = props.getProperty('PENDING_EDITS') || '[]';
  const abwSnapshotRaw = props.getProperty('ABW_SNAPSHOT');
  const spielerSnapshotRaw = props.getProperty('SPIELER_SNAPSHOT');
  const saisonSnapshotRaw = props.getProperty('SAISON_SNAPSHOT');
  const abwAffectedRaw = props.getProperty('ABW_AFFECTED');

  // Sofort löschen, damit bei einem Crash keine veralteten Daten liegen bleiben
  props.deleteProperty('PENDING_EDITS');
  props.deleteProperty('ABW_SNAPSHOT');
  props.deleteProperty('SPIELER_SNAPSHOT');
  props.deleteProperty('SAISON_SNAPSHOT');
  props.deleteProperty('ABW_AFFECTED');

  const pendingEdits: ChangeEntry[] = JSON.parse(pendingRaw);
  const hasAnyPending = pendingEdits.length > 0 || abwSnapshotRaw || spielerSnapshotRaw || saisonSnapshotRaw;
  if (!hasAnyPending) return;

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
    // Saison-Änderungen als kompakte Log-Einträge schreiben
    for (const mod of saisonDiffs.modified) {
      const modChanges = mod.changedCells.map(c => `${c.label}: ${c.oldVal || '-'} → ${c.newVal || '-'}`).join(' | ');
      logEntries.push({
        timestamp: Date.now(),
        sheetName: SHEET_NAMES.SAISON,
        rangeA1: mod.datum,
        alterWert: mod.gegner,
        neuerWert: modChanges,
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
  if (logEntries.length === 0) return;

  let abwAffected: AbwAffectedEntry[] | null = null;
  if (abwAffectedRaw) {
    abwAffected = JSON.parse(abwAffectedRaw);
    if (abwAffected && abwAffected.length === 0) abwAffected = null;
  }

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

// ─── Saison-Validierung (Validierungsspalte) ───────────────────────────────────

function computeValidierung(
  playerRow: string[],
  datum: Date,
  aktiveSpieler: Spieler[],
  allAbw: Map<string, Map<string, string>>,
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number
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

    const name = aktiveSpieler[pi].name;
    const abwDisplay = dayMap?.get(name);
    if (abwDisplay) warnungen.push(`${name}: ${abwDisplay}`);
  }

  let ersatz = 0;
  for (let ei = 0; ei < 3; ei++) {
    if (String(sheet.getRange(row, saisonErsatzCol(ei)).getValue() || '').trim()) {
      ersatz++;
    }
  }

  const f = SHEET_CONFIG.einstellungen.spielformat;
  let missingEinzel = Math.max(0, f.einzel - einzel);
  let missingDoppel = Math.max(0, f.doppel * 2 - doppel);
  for (let i = 0; i < ersatz; i++) {
    if (missingEinzel > 0) { einzel++; missingEinzel--; }
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

function toDateSafe(val: unknown): Date | null {
  if (val instanceof Date) return val;
  if (typeof val === 'string' && val) {
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
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

  const datum = toDateSafe(sheet.getRange(row, saisonDatumCol()).getValue());
  if (!datum) return;

  const aktiveSpieler = readAktiveSpieler(ss.getSheetByName(SHEET_NAMES.SPIELER)!);
  if (!aktiveSpieler || aktiveSpieler.length === 0) return;

  const allAbw = buildAbwesenheitenIndex(ss.getSheetByName(SHEET_NAMES.ABWESENHEITEN)!);
  const playerRow = sheet.getRange(row, saisonSpielerCol(0), 1, aktiveSpieler.length).getValues()[0] as string[];
  const validierung = computeValidierung(playerRow, datum, aktiveSpieler, allAbw, sheet, row);
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
 * Baut eine HTML-Änderungsmail und verschickt sie an alle Empfänger.
 *
 * Empfänger-Kreis:
 *   - Kapitän (Rolle "Kapitän") immer – unabhängig von der Checkbox
 *   - Alle Spieler mit gesetzter Checkbox "Aufstellungsänderungen melden"
 *     (außer dem Bearbeiter selbst, um Selbstbenachrichtigungen zu vermeiden)
 *
 * Darstellung:
 *   - Saison-Änderungen nach Spieltag gruppiert mit Zeilenkontext (Gegner, Datum, Ort)
 *   - Abwesenheits-Änderungen nur wenn ABW_AFFECTED Spieltage listet,
 *     dann angereichert mit Spieltag-Kontext und warAufgestellt/warFinal
 *   - Spieler- und sonstige Änderungen in einer separaten Tabelle
 */
function sendChangeNotification(
  entries: ChangeEntry[],
  abwAffected: AbwAffectedEntry[] | null,
  saisonDiffs: SaisonDiffResult | null
): void {
  // Finalisierungs-Flag: nach "Finalisieren + Emails senden" keine redundante Änderungsmail
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SUPPRESS_NOTIFICATION') === 'true') {
    props.deleteProperty('SUPPRESS_NOTIFICATION');
    return;
  }

  // Nur relevant wenn sichtbare Saison- oder Abw-Änderungen vorliegen
  const visibleAdded = saisonDiffs ? saisonDiffs.added.filter(a => a.status === 'Final').length : 0;
  const visibleSaisonCount = (saisonDiffs?.modified.length ?? 0) + visibleAdded + (saisonDiffs?.deleted.length ?? 0);
  const showAbw = abwAffected && abwAffected.length > 0;
  const abwEntries = showAbw ? entries.filter(e => e.sheetName === SHEET_NAMES.ABWESENHEITEN) : [];
  if (visibleSaisonCount === 0 && abwEntries.length === 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const empfaenger = getAenderungenMeldenEmpfaenger(ss);
  if (empfaenger.length === 0) return;

  const subject = `Einsatzplaner – Änderungen vom ${Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM.yyyy HH:mm')}`;
  let html = `<html><body style="font-family:Arial,sans-serif;font-size:14px;color:#333">
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
  if (showAbw && abwEntries.length > 0) {
    if (visibleSaisonCount > 0) html += `<p style="margin-top:20px">Abwesenheits-Änderungen mit Spieltags-Bezug:</p>`;
    else html += `<p>Abwesenheits-Änderungen mit Spieltags-Bezug:</p>`;

    html += `<table cellpadding="4" cellspacing="0" style="border-collapse:collapse;font-size:13px">`;
    for (const e of abwEntries) {
      html += `<tr><td style="color:#c00;padding-right:8px">${escapeHtml(e.alterWert)}</td><td>→</td><td style="color:#080;padding-left:8px">${escapeHtml(e.neuerWert)}</td></tr>`;
    }
    html += `</table>`;

    for (const a of abwAffected!) {
      const context = [a.datumStr, a.startzeit, a.heimAuswaerts, a.gegner ? `— ${a.gegner}` : ''].filter(Boolean).join(' ');
      const status = a.warAufgestellt
        ? (a.warFinal ? 'War aufgestellt (Final → Geplant) — Nachplanung nötig!' : 'War aufgestellt — Nachplanung nötig!')
        : 'Stand als Ersatz zur Verfügung';
      html += `<p style="margin:2px 0 2px 20px;font-size:13px">${escapeHtml(context)}<br><span style="color:#c00">${escapeHtml(a.spielerName)}: ${status}</span></p>`;
    }
  }

  html += `<p style="margin-top:16px;color:#888">Viele Grüße,<br>Dein Einsatzplaner-Team</p></body></html>`;

  for (const email of empfaenger) {
    MailApp.sendEmail({ to: email, subject, htmlBody: html });
  }
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
  let h = `<p style="margin:12px 0 4px 0;font-weight:bold;font-size:15px">${escapeHtml(datum)} (${typ})</p>`;
  h += `<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;font-size:12px">`;
  h += '<tr style="background:#eee">';
  for (const c of cols) h += `<th>${escapeHtml(c.label)}</th>`;
  h += '</tr><tr>';
  for (const c of cols) {
    const changed = c.oldVal !== c.newVal;
    h += `<td${changed ? ' style="color:#c00"' : ''}>${escapeHtml(c.oldVal || '-')}</td>`;
  }
  h += '</tr><tr>';
  for (const c of cols) {
    const changed = c.oldVal !== c.newVal;
    h += `<td${changed ? ' style="color:#080"' : ''}>${escapeHtml(c.newVal || '-')}</td>`;
  }
  h += '</tr></table>';
  return h;
}

/** Einzeilige Tabelle für neue oder gelöschte Spieltage. */
function saisonSingleRowTableHtml(typ: string, datum: string, cols: { label: string; oldVal: string; newVal: string }[], isNew: boolean): string {
  const displayCols = cols.filter(c => (isNew && c.newVal) || (!isNew && c.oldVal));
  if (displayCols.length === 0) return '';

  const color = isNew ? '#080' : '#c00';
  let h = `<p style="margin:12px 0 4px 0;font-weight:bold;font-size:15px">${escapeHtml(datum)} (${typ})</p>`;
  h += `<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;font-size:12px">`;
  h += '<tr style="background:#eee">';
  for (const c of displayCols) h += `<th>${escapeHtml(c.label)}</th>`;
  h += '</tr><tr>';
  for (const c of displayCols) {
    const v = isNew ? c.newVal : c.oldVal;
    h += `<td style="color:${color}">${escapeHtml(v || '-')}</td>`;
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
 *
 * Optional (wenn Checkbox gesetzt):
 *   - Alle Spieler mit "Aufstellungsänderungen melden" = TRUE
 *   - Ausnahme: der aktuelle Bearbeiter selbst wird nicht benachrichtigt
 */
function getAenderungenMeldenEmpfaenger(ss: GoogleAppsScript.Spreadsheet.Spreadsheet): string[] {
  const spielerSheet = ss.getSheetByName(SHEET_NAMES.SPIELER);
  if (!spielerSheet) return [];

  const lastRow = spielerSheet.getLastRow();
  if (lastRow <= 1) return [];

  const currentUser = Session.getActiveUser().getEmail();
  const data = spielerSheet.getRange(2, 1, lastRow - 1, COL_SPIELER.Rolle).getValues();
  const empfaenger: string[] = [];

  for (const row of data) {
    const email = String(row[COL_SPIELER.Email - 1]).trim();
    if (!email || !email.includes('@') || email === currentUser) continue;

    const rolle = String(row[COL_SPIELER.Rolle - 1] || '').trim();
    const melden = row[COL_SPIELER.AenderungenMelden - 1] === true
      || row[COL_SPIELER.AenderungenMelden - 1] === 'TRUE';

    // Kapitän immer, andere nur mit Checkbox
    if (rolle === 'Kapitän' || melden) {
      if (!empfaenger.includes(email)) empfaenger.push(email);
    }
  }
  return empfaenger;
}
