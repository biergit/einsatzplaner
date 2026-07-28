# PLAN – Einsatzplaner

## Architektur

```
data/*.tsv,json       ──build.py──►  dist/Config.js (überschreibt Defaults)
test-data/*.tsv,json  ──build.py --data-dir test-data──►  dist/Config.js (Testdaten)
                                              │
src/*.ts  ──tsc──►  dist/*.js  ──clasp push──►  Google Apps Script
```

- **`src/`**: Handgeschriebener TypeScript-Code (in Git)
- **`dist/`**: Kompilierte + generierte JS-Dateien (gitignored)
- **`data/`**: Produktivdaten (gitignored – sensible Daten)
- **`test-data/`**: Testdaten (in Git – fiktive Daten)
- **`build.py`**: Liest `--data-dir` (default: `data/`) → generiert `dist/Config.js`

`src/Config.ts` enthält leere Defaults und ist in Git getrackt. Bei `npm run build` wird es zuerst von `tsc` nach `dist/Config.js` kompiliert, dann überschreibt `build.py` die Datei mit den echten Daten aus `--data-dir`.

## Deployment-Modell

Zwei Google-Scripts (Produktiv und Test) mit separaten Script-IDs, verwaltet über zwei `.clasp.json`-Varianten:

```
┌─ deploy ──────────────┐    ┌─ deploy:test ───────────┐
│ switch:prod            │    │ switch:test              │
│   cp .clasp.json.prod  │    │   cp .clasp.json.test    │
│   → .clasp.json        │    │   → .clasp.json          │
│ build (data/)          │    │ build:test (test-data/)  │
│ clasp push             │    │ clasp push               │
└────────────────────────┘    └──────────────────────────┘
```

**Scripts:**

| Script | Zweck |
|--------|-------|
| `npm run create` | Neues Sheet + Apps Script anlegen → `.clasp.json` |
| `npm run save:prod` | `.clasp.json` → `.clasp.json.prod` |
| `npm run save:test` | `.clasp.json` → `.clasp.json.test` |
| `npm run switch:prod` | `.clasp.json.prod` → `.clasp.json` |
| `npm run switch:test` | `.clasp.json.test` → `.clasp.json` |
| `npm run deploy` | `switch:prod` + build + push |
| `npm run deploy:test` | `switch:test` + build:test + push |

Die `.clasp.json*`-Dateien sind alle gitignored (enthalten Google Script-IDs).

## Datenmodell

### Spieler (`COL_SPIELER`)
Name, Email (optional), Rang (1-N), Aktiv (Checkbox), Aufstellungsänderungen melden (Checkbox), Rolle (`Kapitän` oder leer)

### Abwesenheiten (`COL_ABWESENHEITEN`)
Spieler (Dropdown aus Spieler), Von, Bis, Kommentar, Validierung (automatische Validierung)

### Spieltermine (`COL_SPIELTERMINE`)
Datum, Heim/Gast (Dropdown), Gegner, Ort (optional), Status (Geplant/Finalisiert/Versendet)

### Aufstellungen (`COL_AUFSTELLUNGEN`)
Termin (Dropdown aus Spieltermine), Typ (Einzel 1-4 / Doppel), Spieler (Dropdown aus Spieler)

### Änderungslog (`COL_AENDERUNGSLOG`)
Zeitstempel, Bereich, Alter Wert, Neuer Wert, Bearbeiter

## Aufstellungslogik

1. Für jeden Spieltermin: Alle aktiven Spieler laden (sortiert nach Rang)
2. Abwesenheiten für das Datum prüfen → verfügbare Spieler
3. Top-4 verfügbare → Positionen Einzel 1-4
4. Zwei Doppel-Zeilen mit Vorschlägen aus den verfügbaren Spielern
5. Der Kapitän kann die Aufstellung danach manuell anpassen

## Benachrichtigungslogik

### Änderungs-Benachrichtigung (ChangeTracker)

1. `onEdit` (nur installierbarer Trigger mit `authMode === FULL`) triggert bei jeder Bearbeitung
2. Falls `SHEET_BUILDER_RUNNING`-Flag gesetzt → Abbruch (keine Logs/Emails beim Sheet-Neuaufbau)
3. Bearbeitungen an Dokumentation und Änderungslog werden ignoriert
4. Änderungen werden im Debounce-Puffer (`PENDING_EDITS` in ScriptProperties) gesammelt
5. Debounce-Timer (konfigurierbar via `einstellungen.debounceMinuten`) wird zurückgesetzt
6. Nach Ablauf: Snapshot-Diff für Abwesenheiten und Spieler, Saison-Pending-Edits direkt
7. Alle Einträge werden ins Änderungslog geschrieben (permanent, ungefiltert)
8. HTML-Mail an Kapitän (immer) + Spieler mit Checkbox "Aufstellungsänderungen melden"
   - Außer dem Bearbeiter selbst
   - Abwesenheits-Änderungen nur wenn geplante/finale Spieltage betroffen sind
   - Saison-Änderungen nach Spieltag gruppiert mit Vorher-/Nachher-Vergleich

### Einsatzplan-E-Mails (EmailService)

1. `sendEinsatzEmails()` iteriert über alle eingesetzten Spieler
2. Spieler **mit** E-Mail → persönlicher Einsatzplan per Mail
3. Spieler **ohne** E-Mail → Name wird auf "ohneE-Mail"-Liste gesammelt
4. Falls Liste nicht leer: Kapitän (Rolle `Kapitän`) bekommt separate Mail mit den Namen

## Rollen

| Rolle | Beschreibung |
|-------|-------------|
| `Kapitän` | Erhält Änderungs-Mails immer (auch ohne Checkbox). Erhält Validierung über eingesetzte Spieler ohne hinterlegte E-Mail-Adresse |
| *(leer)* | Normales Teammitglied, erhält Änderungs-Mails nur bei gesetzter Checkbox |

Die Rolle ist unabhängig von der Checkbox "Änderungen melden" – ein Spieler kann beides haben, nur eines oder keines.

## Spalten-Enums

Alle Sheet-Zugriffe verwenden `COL_*`-Enums aus `ConfigTypes.ts` statt Magic Numbers:

```typescript
enum COL_SPIELER { Name = 1, Email = 2, Rang = 3, Aktiv = 4, AenderungenMelden = 5, Rolle = 6 }
enum COL_ABWESENHEITEN { Spieler = 1, Von = 2, Bis = 3, Kommentar = 4, Validierung = 5 }
enum COL_SPIELTERMINE { Datum = 1, HeimGast = 2, Gegner = 3, Ort = 4, Status = 5 }
enum COL_AUFSTELLUNGEN { Termin = 1, Typ = 2, Spieler = 3 }
enum COL_AENDERUNGSLOG { Zeitstempel = 1, Bereich = 2, AlterWert = 3, NeuerWert = 4, Bearbeiter = 5 }
```

## Datei-Organisation

| Datei | Zuständigkeit |
|-------|--------------|
| `src/ConfigTypes.ts` | Interfaces + Spalten-Enums |
| `src/Config.ts` | SHEET_CONFIG mit leeren Defaults (getrackt) |
| `src/SheetBuilder.ts` | Löscht alle Sheets und baut 6 Sheets neu auf (inkl. Dokumentation) |
| `src/ChangeTracker.ts` | onEdit-Trigger, Debounce-Timer, Änderungs-Benachrichtigung |
| `src/DataExporter.ts` | Export aller Sheet-Daten nach Google Drive |
| `src/AufstellungsGenerator.ts` | Automatische Aufstellungsberechnung |
| `src/EmailService.ts` | Einsatzplan-E-Mails + Kapitän-Validierung |
| `src/Main.ts` | onOpen-Menü, Menu-Handler |
| `build.py` | Liest Daten → generiert `dist/Config.js` |
| `test/build.test.ts` | Validiert Build-Pipeline und generierte Config.js |

## Roadmap

- [x] Sheet-Builder (alle Sheets programmatisch erzeugen)
- [x] Change-Tracking mit Debounce-Timer
- [x] E-Mail-Benachrichtigungen an interessierte Teammitglieder
- [x] Spieltermine mit Status-Modell
- [x] Automatische Aufstellungs-Generierung (Rang + Abwesenheiten)
- [x] Einsatzplan-E-Mails an Spieler
- [x] Kapitän-Validierung bei Spielern ohne E-Mail
- [x] Test-Build mit fiktiven Daten
- [ ] Daten-Import aus bestehendem Sheet
- [ ] Delta-Erkennung bei Aufstellungs-Änderungen (nur geänderte Spieler benachrichtigen)
- [ ] Saison-bezogene Auswertungen
- [ ] Fehlerlog-Blatt (500 Zeilen, Time-basiertes Trimming) für unbehandelte Exceptions in Mail-Versand, Timer, PropertyService
