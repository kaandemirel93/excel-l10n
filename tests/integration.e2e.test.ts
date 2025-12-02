import path from 'node:path';
import fs from 'node:fs';
import ExcelJS from 'exceljs';
import { extract, exportUnitsToXliff, parseTranslated, merge, parseConfig } from '../src/index';
import type { Config } from '../src/types';

async function createWorkbook(tmpPath: string): Promise<string> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.getCell('A1').value = 'Header';
  ws.getCell('A2').value = 'Hello {0}! This is v1.';
  ws.getCell('A3').value = 'Dr. Smith went home. It was late!';
  // comment
  ws.getCell('A2').note = 'greeting';
  const file = path.join(tmpPath, 'sample.xlsx');
  await wb.xlsx.writeFile(file);
  return file;
}

function cfgWithSrx(tmpPath: string): Config {
  const cfgPath = path.resolve(__dirname, '../examples/config.yml');
  const cfg = parseConfig(cfgPath);
  // Ensure translateComments is on for test
  cfg.workbook.sheets[0].translateComments = true;
  // Force source column to A for this test, regardless of what's in config.yml
  cfg.workbook.sheets[0].sourceColumns = ['A'];
  // Adjust target columns to match: B for fr (after A)
  cfg.workbook.sheets[0].targetColumns = { fr: 'B' };
  return cfg;
}

describe('integration: extract → SRX → XLIFF → parse → merge', () => {
  const tmpDir = path.join(process.cwd(), '.out');
  beforeAll(() => { if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true }); });

  it('roundtrips placeholders and segments, merges back to target cols', async () => {
    const xlsx = await createWorkbook(tmpDir);
    const cfg = cfgWithSrx(tmpDir);

    const units = await extract(xlsx, cfg);
    expect(units.length).toBeGreaterThan(0);
    // expect segmentation on sentences
    const hello = units.find(u => u.row === 2)!;
    expect(hello.segments && hello.segments.length).toBeGreaterThan(0);

    const xlf = await exportUnitsToXliff(units, cfg, { srcLang: 'en' });
    expect(xlf).toContain('<xliff');
    expect(xlf).toContain('<unit');
    expect(xlf).toContain('<ph'); // placeholder present

    // simulate translation: add targets
    const parsed = parseTranslated(xlf, 'xlf');
    for (const tu of parsed) {
      if (tu.segments) {
        // Translate by wrapping source and keeping placeholder markers in place
        tu.segments = tu.segments.map(s => ({ ...s, target: `[${s.source}]` }));
      }
    }

    const outXlsx = path.join(tmpDir, 'sample.translated.xlsx');
    await merge(xlsx, outXlsx, parsed, cfg);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outXlsx);
    const ws2 = wb2.getWorksheet('Sheet1')!;
    // Target fr in column B per examples/config.yml, placeholder rehydrated
    const b2 = String(ws2.getCell('B2').value);
    expect(b2 && b2.length).toBeGreaterThan(0);
    expect(b2.startsWith('[')).toBe(true);
    // placeholders should be rehydrated (no [[ph:...]] markers), and Hello present
    expect(b2).toContain('Hello');
    expect(b2).toContain('{0}');
    expect(b2).not.toContain('[[ph:');
    // comments captured in meta (not directly re-written to Excel for MVP)
    // ensure original A2 is untouched
    expect(String(ws2.getCell('A2').value)).toBe('Hello {0}! This is v1.');
  });
});
