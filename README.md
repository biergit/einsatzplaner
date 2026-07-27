# Einsatzplaner

Google Apps Script zur Verwaltung eines Tischtennis-Teams. Erfasst Abwesenheiten, Spieltermine und generiert automatisch Einsatzpläne basierend auf Spielerrängen und Verfügbarkeit.

## Setup

### 1. Abhängigkeiten installieren

```bash
npm install
```

### 2. Daten vorbereiten

Die Dateien im `data/`-Verzeichnis enthalten die Team-Daten (gitignored – kein Commit sensibler Daten):

| Datei | Format | Inhalt |
|-------|--------|--------|
| `spieler.tsv` | TSV | Name, Email, Rang, Änderungen melden, Rolle |
| `abwesenheiten.tsv` | TSV | Spieler, Von, Bis, Kommentar |
| `spieltermine.tsv` | TSV | Datum, Heim/Gast, Gegner, Ort |
| `einstellungen.json` | JSON | teamName, saison, spielformat |

Vorlagen liegen als `.sample`-Dateien im `data/`-Verzeichnis.

### 3. Mit Google verbinden

```bash
npm run login       # Google-Login im Browser
npm run create      # Erstellt Produktiv-Sheet + Apps Script
```

### 4. Deployen

```bash
npm run deploy      # Baut mit Produktivdaten + pushed zu Google
npm run open        # Öffnet das Sheet im Browser
```

## Test-Build

Für Tests ohne Produktivdaten liegt in `test-data/` ein Satz fiktiver Daten:

```bash
npm run create:test     # Einmalig: eigenes Test-Sheet anlegen
npm run deploy:test     # Baut mit Testdaten + pushed
```

**Achtung:** Nach `create:test` ist die `.clasp.json` auf das Test-Script umgebogen. Für das Produktiv-Sheet wieder `npm run create` ausführen (oder `.clasp.json` manuell anpassen).

## Entwicklung

```
src/           TypeScript-Quellcode (handgeschrieben)
dist/          Kompilierte JS-Dateien + generierte Config (gitignored)
data/          Produktivdaten (gitignored)
test-data/     Testdaten – fiktiv, in Git getrackt
build.py       Liest data/ oder test-data/ → generiert dist/Config.js
```

- `npm run build` – `tsc` + `build.py` (Produktivdaten)
- `npm run build:test` – `tsc` + `build.py --data-dir test-data` (Testdaten)
- `npm run watch` – TypeScript-Watch-Modus
- `npm test` – Build-Pipeline-Tests ausführen
- `npm run test:watch` – Tests im Watch-Modus

## Sheets

| Sheet | Zweck |
|-------|-------|
| **Spieler** | Name, Email (optional), Rang, Aktiv, Änderungen melden, Rolle |
| **Abwesenheiten** | Spieler, Von, Bis, Kommentar |
| **Spieltermine** | Datum, Heim/Gast, Gegner, Ort (optional), Status |
| **Aufstellungen** | Termin, Typ (Einzel/Doppel), Spieler |
| **Änderungslog** | Automatisches Protokoll aller Änderungen (versteckt) |

## Menü-Funktionen

| Menüpunkt | Beschreibung |
|-----------|-------------|
| Sheet neu aufbauen | Löscht alle Sheets und baut sie aus der Konfiguration neu auf |
| Daten exportieren | Sichert alle Daten als JSON nach Google Drive |
| Aufstellungen generieren | Ermittelt verfügbare Spieler pro Termin und füllt das Aufstellungs-Sheet |
| Aufstellungen finalisieren | Setzt Status aller geplanten Termine auf "Finalisiert" |
| Emails senden | Versendet Einsatzpläne an alle eingesetzten Spieler mit E-Mail und informiert den Kapitän über Spieler ohne E-Mail |

## Automatische Benachrichtigungen

- Bei jeder Bearbeitung durch ein Teammitglied wird die Änderung protokolliert
- Nach 5 Minuten Inaktivität erhalten alle Spieler mit Checkbox **Änderungen melden** eine E-Mail mit den letzten Änderungen – außer sie waren selbst die Bearbeiter
- Der Bearbeiter wird anhand seiner Google-Account-E-Mail identifiziert
- Änderungen am Aufstellungs-Sheet und am Änderungslog selbst lösen keine Benachrichtigung aus

## Rollen

| Rolle | Bedeutung |
|-------|-----------|
| `Kapitän` | Bekommt nach dem E-Mail-Versand einen Hinweis über eingesetzte Spieler ohne E-Mail-Adresse |
| *(leer)* | Normales Teammitglied |

## Spieler ohne E-Mail

E-Mail-Adressen sind optional. Hat ein eingesetzter Spieler keine E-Mail hinterlegt, bekommt der Kapitän nach dem Versand der Einsatzpläne eine separate Nachricht mit der Liste dieser Spieler – damit er sie persönlich kontaktieren kann.
