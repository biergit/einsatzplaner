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
| `einstellungen.json` | JSON | teamName, saison, debounceMinuten, spielformat |

Vorlagen liegen als `.sample`-Dateien im `data/`-Verzeichnis.

### 3. Google-Scripts einrichten

```bash
npm run login               # Einmalig: Google-Login im Browser

# Produktiv-Script
npm run create              # Erstellt Sheet + Apps Script
npm run save:prod           # Speichert Script-ID als Produktiv

# Test-Script
npm run create              # Erstellt zweites Sheet + Apps Script
npm run save:test           # Speichert Script-ID als Test
```

Die Script-IDs werden in `.clasp.json.prod` und `.clasp.json.test` abgelegt (beide gitignored).

### 4. Deployen

```bash
npm run deploy              # Schaltet auf Produktiv → baut → pushed
npm run deploy:test         # Schaltet auf Test → baut mit Testdaten → pushed
npm run open               # Öffnet das aktuell aktive Sheet im Browser
npm run open:script        # Öffnet den Apps Script Editor
```

`deploy` und `deploy:test` kümmern sich automatisch um das Umschalten des Script-Kontexts. Kein manuelles Kopieren nötig.

### 5. Berechtigungen erteilen

Nach dem ersten Deploy öffnest du das Sheet und der Code ist da, aber das Menü fehlt noch. Einmalig:

1. Im **Apps Script Editor** die Funktion `onOpen` aus dem Dropdown auswählen (oben mittig)
2. Auf **"Ausführen"** (▶️) klicken
3. Google fragt nach Berechtigungen → alles bestätigen
4. Sheet im Browser neu laden (F5)

Danach erscheint das Menü **Einsatzplaner** und alle Funktionen stehen bereit. Dieser Schritt ist nur einmal nötig – bei späteren `clasp push`-Deployments ist nichts weiter zu tun.

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

## Script-Verwaltung

| Befehl | Beschreibung |
|--------|-------------|
| `npm run create` | Neues Sheet + Apps Script bei Google anlegen |
| `npm run save:prod` | Aktuelle Script-ID als Produktiv sichern (`.clasp.json.prod`) |
| `npm run save:test` | Aktuelle Script-ID als Test sichern (`.clasp.json.test`) |
| `npm run switch:prod` | Auf Produktiv-Script umschalten |
| `npm run switch:test` | Auf Test-Script umschalten |
| `npm run open` | Aktuell aktives Sheet im Browser öffnen |
| `npm run open:script` | Apps Script Editor für aktuelles Projekt öffnen |

## Sheets

| Sheet | Zweck |
|-------|-------|
| **Dokumentation** | Erläuterung aller Spalten, Funktionen und des Workflows |
| **Spieler** | Name, Email (optional), Rang, Aktiv, Änderungen melden, Rolle |
| **Abwesenheiten** | Spieler, Von, Bis, Kommentar |
| **Spieltermine** | Datum, Heim/Gast, Gegner, Ort (optional), Status |
| **Aufstellungen** | Termin, Typ (Einzel/Doppel), Spieler |
| **Änderungslog** | Automatisches Protokoll aller Änderungen (versteckt) |

## Menü-Funktionen

| Menüpunkt | Beschreibung |
|-----------|-------------|
| Sheet neu aufbauen | Löscht alle Sheets und baut sie aus der Konfiguration neu auf. Löst **keine** E-Mail-Benachrichtigungen aus. |
| Daten exportieren | Sichert alle Daten als JSON nach Google Drive |
| Aufstellungen generieren | Ermittelt verfügbare Spieler pro Termin und füllt das Aufstellungs-Sheet |
| Aufstellungen finalisieren | Setzt Status aller geplanten Termine auf "Finalisiert" |
| Emails senden | Versendet Einsatzpläne an alle eingesetzten Spieler mit E-Mail und informiert den Kapitän über Spieler ohne E-Mail |

## Automatische Benachrichtigungen

- Bei jeder Bearbeitung durch ein Teammitglied wird die Änderung protokolliert
- Nach einer konfigurierbaren Zeitspanne ohne weitere Änderung (`debounceMinuten` in `einstellungen.json`) erhalten alle Spieler mit Checkbox **Änderungen melden** eine E-Mail mit den letzten Änderungen – außer sie waren selbst die Bearbeiter
- Der Bearbeiter wird anhand seiner Google-Account-E-Mail identifiziert
- Änderungen am Dokumentation-, Aufstellungs- und Änderungslog-Sheet lösen keine Benachrichtigung aus
- Der Neuaufbau des Sheets (Menü → Sheet neu aufbauen) löst keine Benachrichtigungen aus

## Rollen

| Rolle | Bedeutung |
|-------|-----------|
| `Kapitän` | Bekommt nach dem E-Mail-Versand einen Hinweis über eingesetzte Spieler ohne E-Mail-Adresse |
| *(leer)* | Normales Teammitglied |

## Spieler ohne E-Mail

E-Mail-Adressen sind optional. Hat ein eingesetzter Spieler keine E-Mail hinterlegt, bekommt der Kapitän nach dem Versand der Einsatzpläne eine separate Nachricht mit der Liste dieser Spieler – damit er sie persönlich kontaktieren kann.
