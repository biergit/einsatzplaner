import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, it, expect } from 'vitest';

const ROOT = resolve(__dirname, '..');
const src = readFileSync(resolve(ROOT, 'src', 'SheetBuilder.ts'), 'utf-8');

/** Liefert die 1-basierten Zeilennummern der gefundenen Überschriften */
function findHeaderRows(): number[] {
  const rows: number[] = [];
  // Suche nach den Überschrifts-Labels in den rows-Array-Einträgen
  const labels = ['SHEETS IM ÜBERBLICK', 'SPIELER-SPALTEN', 'SAISON-SPALTEN', 'MENÜ', 'BENACHRICHTIGUNGEN'];
  // Extrahiere den rows-Block zwischen "const rows: string[][] = [" und "];"
  const rowsBlock = src.match(/const rows: string\[\]\[\] = \[([\s\S]*?)\];/)![1];

  // Finde für jeden Label die Zeilennummer (1-basiert) im Array
  let lineNum = 0;
  for (const line of rowsBlock.split('\n')) {
    const trimmed = line.trim();
    if (trimmed.startsWith('[') && !trimmed.startsWith('//')) {
      lineNum++;
      for (const label of labels) {
        if (trimmed.includes(`'${label}'`) || trimmed.includes(`"${label}"`)) {
          rows.push(lineNum);
        }
      }
    }
  }
  return rows;
}

function findFormatIndices(): number[] {
  const m = src.match(/for\s*\(\s*const\s+r\s+of\s+\[([^\]]+)\]/);
  if (!m) throw new Error('Formatierungs-Indizes nicht gefunden');
  return m[1].split(',').map(s => parseInt(s.trim(), 10));
}

describe('Dokumentation Sheet Formatierung', () => {
  it('Überschriften-Zeilen sind blau hinterlegt, andere nicht', () => {
    const headerRows = findHeaderRows();
    const formatIndices = findFormatIndices();
    expect(headerRows).toEqual(formatIndices);
  });
});
