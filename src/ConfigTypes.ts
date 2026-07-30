declare const LOGO_BASE64: string;

enum COL_SPIELER {
  Name = 1,
  Email = 2,
  Rang = 3,
  AenderungenMelden = 4,
  Rolle = 5,
}

enum COL_ABWESENHEITEN {
  Spieler = 1,
  Von = 2,
  Bis = 3,
  Kommentar = 4,
  Validierung = 5,
}

enum COL_AENDERUNGSLOG {
  Zeitstempel = 1,
  Bereich = 2,
  AlterWert = 3,
  NeuerWert = 4,
  Bearbeiter = 5,
}

function saisonDatumCol(): number { return 1; }
function saisonWochentagCol(): number { return 2; }
function saisonGegnerCol(): number { return 3; }
function saisonStartzeitCol(): number { return 4; }
function saisonHeimAuswaertsCol(): number { return 5; }
function saisonSpielerColStart(): number { return 6; }

function saisonSpielerCol(playerIndex: number): number {
  return saisonSpielerColStart() + playerIndex;
}

function saisonErsatzColStart(): number {
  return saisonSpielerColStart() + SHEET_CONFIG.spieler.length;
}

function saisonErsatzCol(index: number): number {
  return saisonErsatzColStart() + index;
}

function saisonStatusCol(): number {
  return saisonErsatzColStart() + 3;
}

function saisonKommentarCol(): number {
  return saisonStatusCol() + 1;
}

function saisonValidierungCol(): number {
  return saisonStatusCol() + 2;
}

function saisonColCount(): number {
  return saisonValidierungCol();
}

interface Spieler {
  name: string;
  email: string;
  rang: number;
  aenderungenMelden: boolean;
  rolle: 'Kapitän' | '';
}

interface Abwesenheit {
  spieler: string;
  von: Date;
  bis: Date;
  kommentar: string;
}

type AufstellungsTyp = 'Einzel+Doppel' | 'Einzel' | 'Doppel';

const ALLE_AUFSTELLUNGS_TYPEN: AufstellungsTyp[] = ['Einzel+Doppel', 'Einzel', 'Doppel'];

interface Einstellungen {
  sheetTitel: string;
  teamName: string;
  saison: string;
  saisonBeginn: Date;
  saisonEnde: Date;
  debounceMinuten: number;
  spieltage: number[];
  spielformat: {
    einzel: number;
    doppel: number;
    system: string;
  };
}

interface SheetConfig {
  spieler: Spieler[];
  abwesenheiten: Abwesenheit[];
  einstellungen: Einstellungen;
}

type AufstellungsZelle = AufstellungsTyp | '';

interface SaisonValidierung {
  einzelCount: number;
  doppelCount: number;
  ersatzCount: number;
  gesamtCount: number;
  warnungen: string[];
}
