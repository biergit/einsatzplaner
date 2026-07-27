enum COL_SPIELER {
  Name = 1,
  Email = 2,
  Rang = 3,
  Aktiv = 4,
  AenderungenMelden = 5,
  Rolle = 6,
}

enum COL_ABWESENHEITEN {
  Spieler = 1,
  Von = 2,
  Bis = 3,
  Kommentar = 4,
}

enum COL_AENDERUNGSLOG {
  Zeitstempel = 1,
  Bereich = 2,
  AlterWert = 3,
  NeuerWert = 4,
  Bearbeiter = 5,
}

function saisonMaxSlots(): number {
  return SHEET_CONFIG.einstellungen.spielformat.einzel + SHEET_CONFIG.einstellungen.spielformat.doppel;
}

function saisonEinsatzartCol(slotIndex: number): number {
  return 5 + slotIndex * 2;
}

function saisonSpielerCol(slotIndex: number): number {
  return 6 + slotIndex * 2;
}

function saisonStatusCol(): number {
  return 5 + saisonMaxSlots() * 2;
}

function saisonHinweisCol(): number {
  return saisonStatusCol() + 1;
}

function saisonErsterSpielerCol(): number {
  return saisonHinweisCol() + 1;
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

type AufstellungStatus = 'Geplant' | 'Final';

interface Einstellungen {
  teamName: string;
  saison: string;
  saisonBeginn: Date;
  saisonEnde: Date;
  debounceMinuten: number;
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
