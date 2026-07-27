import { execSync } from 'node:child_process';
import { readFileSync, existsSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, it, expect, beforeAll } from 'vitest';

const ROOT = resolve(__dirname, '..');
const CONFIG_JS = resolve(ROOT, 'dist', 'Config.js');

beforeAll(() => {
  execSync('npx tsc && cp appsscript.json dist/ && python3 build.py --data-dir test-data', {
    cwd: ROOT, stdio: 'pipe',
  });
});

describe('build pipeline', () => {
  it('generates dist/Config.js', () => {
    expect(existsSync(CONFIG_JS)).toBe(true);
  });

  it('compiles all TypeScript source files', () => {
    const files = ['AufstellungsGenerator.js', 'ChangeTracker.js', 'Config.js', 'ConfigTypes.js',
      'DataExporter.js', 'EmailService.js', 'Main.js', 'SheetBuilder.js'];
    for (const f of files) {
      expect(existsSync(resolve(ROOT, 'dist', f))).toBe(true);
    }
  });
});

describe('generated Config.js', () => {
  let content: string;
  beforeAll(() => { content = readFileSync(CONFIG_JS, 'utf-8'); });

  it('defines SHEET_CONFIG', () => {
    expect(content).toContain('var SHEET_CONFIG');
  });

  it('contains team and saison', () => {
    expect(content).toContain("teamName: 'TT Team Test'");
    expect(content).toContain("saisonBeginn: new Date('2026-09-01')");
    expect(content).toContain("saisonEnde: new Date('2026-12-31')");
  });

  it('contains debounceMinuten', () => {
    expect(content).toContain('debounceMinuten: 1');
  });

  it('contains all 5 test players', () => {
    for (const name of ['Max Mustermann', 'Anna Beispiel', 'John Doe', 'Erika Musterfrau', 'Peter Test']) {
      expect(content).toContain(`name: '${name}'`);
    }
  });

  it('marks the captain', () => {
    expect(content).toContain("name: 'Max Mustermann', email: 'max@example.com', rang: 1, aenderungenMelden: true, rolle: 'Kapitän'");
  });

  it('handles players without email', () => {
    expect(content).toContain("name: 'Erika Musterfrau', email: '', rang: 4, aenderungenMelden: false, rolle: ''");
  });

  it('contains 6 test abwesenheiten', () => {
    const count = (content.match(/spieler: '/g) || []).length;
    expect(count).toBeGreaterThanOrEqual(6);
  });

  it('contains spielformat config', () => {
    expect(content).toContain('einzel: 4');
    expect(content).toContain('doppel: 2');
  });

  it('contains spieltage config', () => {
    expect(content).toContain('spieltage: [6]');
  });
});
