#!/usr/bin/env python3
"""Baut dist/Config.js aus den Datendateien im angegebenen data-Verzeichnis."""

import argparse
import base64
import csv
import json
import mimetypes
import os
from datetime import date, datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DATA_DIR = os.path.join(SCRIPT_DIR, "data")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "dist", "Config.js")


def parse_date(s: str) -> date:
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Konnte Datum nicht parsen: {s}")


def read_spieler(data_dir: str) -> list[dict]:
    path = os.path.join(data_dir, "spieler.tsv")
    if not os.path.exists(path):
        print(f"  [skip] {path} nicht gefunden")
        return []
    with open(path, encoding="utf-8") as f:
        reader = csv.DictReader(f, delimiter="\t")
        result = []
        for row in reader:
            name = (row.get("Name") or "").strip()
            if not name:
                continue
            melden_raw = (row.get("Änderungen melden") or "").strip().lower()
            melden = melden_raw in ("ja", "true", "1", "x")
            rolle = (row.get("Rolle") or "").strip()
            result.append({
                "name": name,
                "email": (row.get("Email") or "").strip(),
                "rang": int(row.get("Rang") or 99),
                "aenderungenMelden": "true" if melden else "false",
                "rolle": rolle if rolle == "Kapitän" else "",
            })
        return result


def read_abwesenheiten(data_dir: str) -> list[dict]:
    path = os.path.join(data_dir, "abwesenheiten.tsv")
    if not os.path.exists(path):
        print(f"  [skip] {path} nicht gefunden")
        return []
    with open(path, encoding="utf-8") as f:
        reader = csv.DictReader(f, delimiter="\t")
        result = []
        for row in reader:
            name = (row.get("Spieler/in") or "").strip()
            if not name:
                continue
            von = parse_date(row["Abwesenheit von"])
            bis = parse_date(row["Abwesenheit bis"])
            kommentar = (row.get("Kommentar") or "").strip()
            result.append({
                "name": name,
                "von": von,
                "bis": bis,
                "kommentar": kommentar,
            })
        return result


def read_einstellungen(data_dir: str) -> dict:
    path = os.path.join(data_dir, "einstellungen.json")
    if not os.path.exists(path):
        print(f"  [skip] {path} nicht gefunden, verwende defaults")
        return {
            "sheetTitel": "TT Einsatzplaner",
            "teamName": "TT Team",
            "saison": "2026/2027",
            "saisonBeginn": "2026-09-01",
            "saisonEnde": "2026-12-31",
            "debounceMinuten": 5,
            "spieltage": [6],
            "spielformat": {"einzel": 4, "doppel": 2, "system": "Bundessystem"},
        }
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def generate_js(spieler, abwesenheiten, einstellungen, logo_base64: str) -> str:
    lines = ["// GENERIERT von build.py – nicht manuell bearbeiten.", ""]
    if logo_base64:
        lines.append(f'const LOGO_BASE64 = "{logo_base64}";')
    else:
        lines.append('const LOGO_BASE64 = "";')
    lines.append("")
    lines.append("var SHEET_CONFIG = {")

    lines.append("  einstellungen: {")
    sheet_titel = einstellungen.get("sheetTitel", einstellungen["teamName"])
    lines.append(f"    sheetTitel: '{sheet_titel}',")
    lines.append(f"    teamName: '{einstellungen['teamName']}',")
    lines.append(f"    saison: '{einstellungen['saison']}',")
    lines.append(f"    saisonBeginn: new Date('{einstellungen['saisonBeginn']}'),")
    lines.append(f"    saisonEnde: new Date('{einstellungen['saisonEnde']}'),")
    debounce = einstellungen.get("debounceMinuten", 5)
    lines.append(f"    debounceMinuten: {debounce},")
    spieltage = einstellungen.get("spieltage", [6])
    lines.append(f"    spieltage: {json.dumps(spieltage)},")
    sf = einstellungen["spielformat"]
    lines.append("    spielformat: {")
    lines.append(f"      einzel: {sf['einzel']},")
    lines.append(f"      doppel: {sf['doppel']},")
    lines.append(f"      system: '{sf['system']}',")
    lines.append("    },")
    lines.append("  },")

    lines.append("  spieler: [")
    for s in spieler:
        escaped_email = s['email'].replace("'", "\\'")
        lines.append(f"    {{ name: '{s['name']}', email: '{escaped_email}', rang: {s['rang']}, aenderungenMelden: {s['aenderungenMelden']}, rolle: '{s['rolle']}' }},")
    lines.append("  ],")

    lines.append("  abwesenheiten: [")
    for a in abwesenheiten:
        kommentar = a["kommentar"].replace("'", "\\'").replace("\n", "\\n")
        lines.append("    {")
        lines.append(f"      spieler: '{a['name']}',")
        lines.append(f"      von: new Date('{a['von'].isoformat()}'),")
        lines.append(f"      bis: new Date('{a['bis'].isoformat()}'),")
        lines.append(f"      kommentar: '{kommentar}',")
        lines.append("    },")
    lines.append("  ],")

    lines.append("};")
    lines.append("")
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="Baue dist/Config.js aus Datendateien.")
    parser.add_argument("--data-dir", default=DEFAULT_DATA_DIR, help="Pfad zum data-Verzeichnis (default: data/)")
    args = parser.parse_args()

    data_dir = os.path.abspath(args.data_dir)
    print(f"Baue dist/Config.js aus {data_dir} ...")
    print("  Lese spieler.tsv ...")
    spieler = read_spieler(data_dir)
    print(f"    {len(spieler)} Spieler gelesen")

    print("  Lese abwesenheiten.tsv ...")
    abwesenheiten = read_abwesenheiten(data_dir)
    print(f"    {len(abwesenheiten)} Abwesenheiten gelesen")

    print("  Lese einstellungen.json ...")
    einstellungen = read_einstellungen(data_dir)

    logo_base64 = ""
    logo_path = os.path.join(data_dir, "logo.png")
    if os.path.exists(logo_path):
        print(f"  Lese logo.png ...")
        mime, _ = mimetypes.guess_type(logo_path)
        with open(logo_path, "rb") as f:
            encoded = base64.b64encode(f.read()).decode()
        logo_base64 = f"data:{mime or 'image/png'};base64,{encoded}"

    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    content = generate_js(spieler, abwesenheiten, einstellungen, logo_base64)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"  dist/Config.js geschrieben ({len(content)} Bytes)")
    print("Fertig. Führe 'clasp push' aus um ins Google Sheet zu deployen.")


if __name__ == "__main__":
    main()
