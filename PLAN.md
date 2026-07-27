# PLAN – Einsatzplaner

## Architektur

```
data/*.tsv,json  ──build.py──►  dist/Config.js (überschreibt Defaults)
test/*.tsv,json  ──build.py --data-dir test──►  dist/Config.js (Testdaten)
                                      │
src/*.ts  ──tsc──►  dist/*.js  ──clasp push──►  Google Apps Script
```

- **`src/`**: Handgeschriebener TypeScript-Code (in Git)
- **`dist/`**: Kompilierte + generierte JS-Dateien (gitignored)
- **`data/`**: Produktivdaten (gitignored – sensible Daten)
- **`test/`**: Testdaten (in Git – fiktive Daten)
- **`build.py`**: Liest `--data-dir` (default: `data/`) → generiert `dist/Config.js`

`src/Config.ts` enthält leere Defaults und ist in Git getrackt. Bei `npm run build` wird es zuerst von `tsc` nach `dist/Config.js` kompiliert, dann überschreibt `build.py` die Datei mit den echten Daten aus `--data-dir`.

## Datenmodell

### Spieler (`COL_SPIELER`)
Name, Email (optional), Rang (1-N), Aktiv (Checkbox), Änderungen melden (Checkbox), Rolle (`Kapitän` oder leer)

### Abwesenheiten (`COL_ABWESENHEITEN`)
Spieler (Dropdown aus Spieler), Von, Bis, Kommentar

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
5. "nur wenn es sein muss"-Hinweise werden als separate Zeile ausgegeben
6. Der Kapitän kann die Aufstellung danach manuell anpassen

## Benachrichtigungslogik

### Änderungs-Benachrichtigung (ChangeTracker)

1. `onEdit` triggert bei jeder Bearbeitung (außer Aufstellungen und Änderungslog)
2. Änderung wird ins Änderungslog geschrieben
3. Debounce-Timer (5 min) wird zurückgesetzt
4. Nach Ablauf: Alle Benutzer mit Checkbox "Änderungen melden" bekommen E-Mail
5. Der Bearbeiter selbst wird aus der Empfängerliste ausgeschlossen (via Google-Account-E-Mail)

### Einsatzplan-E-Mails (EmailService)

1. `sendEinsatzEmails()` iteriert über alle eingesetzten Spieler
2. Spieler **mit** E-Mail → persönlicher Einsatzplan per Mail
3. Spieler **ohne** E-Mail → Name wird auf "ohneE-Mail"-Liste gesammelt
4. Falls Liste nicht leer: Kapitän (Rolle `Kapitän`) bekommt separate Mail mit den Namen

## Rollen

| Rolle | Beschreibung |
|-------|-------------|
| `Kapitän` | Erhält Hinweis über eingesetzte Spieler ohne E-Mail-Adresse |
| *(leer)* | Normales Teammitglied, keine Sonderfunktion |

Die Rolle ist unabhängig von der Checkbox "Änderungen melden" – ein Spieler kann beides haben, nur eines oder keines.

## Spalten-Enums

Alle Sheet-Zugriffe verwenden `COL_*`-Enums aus `ConfigTypes.ts` statt Magic Numbers:

```typescript
enum COL_SPIELER { Name = 1, Email = 2, Rang = 3, Aktiv = 4, AenderungenMelden = 5, Rolle = 6 }
enum COL_ABWESENHEITEN { Spieler = 1, Von = 2, Bis = 3, Kommentar = 4 }
enum COL_SPIELTERMINE { Datum = 1, HeimGast = 2, Gegner = 3, Ort = 4, Status = 5 }
enum COL_AUFSTELLUNGEN { Termin = 1, Typ = 2, Spieler = 3 }
enum COL_AENDERUNGSLOG { Zeitstempel = 1, Bereich = 2, AlterWert = 3, NeuerWert = 4, Bearbeiter = 5 }
```

## Datei-Organisation

| Datei | Zuständigkeit |
|-------|--------------|
| `src/ConfigTypes.ts` | Interfaces + Spalten-Enums |
| `src/Config.ts` | SHEET_CONFIG mit leeren Defaults (getrackt) |
| `src/SheetBuilder.ts` | Programmatische Erstellung aller 5 Sheets |
| `src/ChangeTracker.ts` | onEdit-Trigger, Debounce-Timer, Änderungs-Benachrichtigung |
| `src/DataExporter.ts` | Export aller Sheet-Daten nach Google Drive |
| `src/AufstellungsGenerator.ts` | Automatische Aufstellungsberechnung |
| `src/EmailService.ts` | Einsatzplan-E-Mails + Kapitän-Hinweis |
| `src/Main.ts` | onOpen-Menü, Menu-Handler |
| `build.py` | Liest Daten → generiert `dist/Config.js` |

## Roadmap

- [x] Sheet-Builder (alle Sheets programmatisch erzeugen)
- [x] Change-Tracking mit Debounce-Timer
- [x] E-Mail-Benachrichtigungen an interessierte Teammitglieder
- [x] Spieltermine mit Status-Modell
- [x] Automatische Aufstellungs-Generierung (Rang + Abwesenheiten)
- [x] Einsatzplan-E-Mails an Spieler
- [x] Kapitän-Hinweis bei Spielern ohne E-Mail
- [x] Test-Build mit fiktiven Daten
- [ ] Daten-Import aus bestehendem Sheet
- [ ] Delta-Erkennung bei Aufstellungs-Änderungen (nur geänderte Spieler benachrichtigen)
- [ ] Saison-bezogene Auswertungen
