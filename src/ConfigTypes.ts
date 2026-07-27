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

enum COL_SPIELTERMINE {
  Datum = 1,
  HeimGast = 2,
  Gegner = 3,
  Ort = 4,
  Status = 5,
}

enum COL_AUFSTELLUNGEN {
  Termin = 1,
  Typ = 2,
  Spieler = 3,
}

enum COL_AENDERUNGSLOG {
  Zeitstempel = 1,
  Bereich = 2,
  AlterWert = 3,
  NeuerWert = 4,
  Bearbeiter = 5,
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

interface Spieltermin {
  datum: Date;
  heimGast: 'Heim' | 'Gast';
  gegner: string;
  ort: string;
  status: 'Geplant' | 'Finalisiert' | 'Versendet';
}

interface AufstellungsEintrag {
  spielterminDatum: Date;
  typ: AufstellungsTyp;
  spieler: string;
}

type AufstellungsTyp = 'Einzel 1' | 'Einzel 2' | 'Einzel 3' | 'Einzel 4' | 'Doppel';

interface Einstellungen {
  teamName: string;
  saison: string;
  spielformat: {
    einzel: number;
    doppel: number;
    system: string;
  };
}

interface SheetConfig {
  spieler: Spieler[];
  abwesenheiten: Abwesenheit[];
  spieltermine: Spieltermin[];
  einstellungen: Einstellungen;
}
