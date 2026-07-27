/// <reference path="ConfigTypes.ts" />

const SHEET_CONFIG: SheetConfig = {
  einstellungen: {
    teamName: 'TT Team',
    saison: '2026/2027',
    saisonBeginn: new Date('2026-09-01'),
    saisonEnde: new Date('2026-12-31'),
    debounceMinuten: 5,
    spieltage: [6],
    spielformat: {
      einzel: 4,
      doppel: 2,
      system: 'Bundessystem',
    },
  },
  spieler: [],
  abwesenheiten: [],
};
