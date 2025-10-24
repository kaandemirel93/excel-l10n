import path from 'node:path';
import fs from 'node:fs';
import ExcelJS from 'exceljs';
import { extract, exportUnitsToJson, merge } from '../src/index';
import type { Config, TranslationUnit } from '../src/types';

async function createWorkbook(tmpPath: string): Promise<string> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.getCell('A1').value = 'Header';
  ws.getCell('A2').value = 'Hello {0}! This is v1.';
  ws.getCell('A2').font = { bold: true, color: { argb: 'FF000000' } };
  ws.getCell('A2').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } } as any;
  const file = path.join(tmpPath, 'sample-json.xlsx');
  await wb.xlsx.writeFile(file);
  return file;
}

function inlineCfg(): Config {
  return {
    global: { srcLang: 'en', overwrite: true, insertTargetPlacement: 'insertAfterSource' },
    segmentation: { enabled: true, rules: 'builtin' },
    workbook: {
      sheets: [{
        namePattern: 'Sheet1',
        sourceColumns: ['A','B'],
        targetColumns: { fr: 'B' },
        createTargetIfMissing: true,
        headerRow: 1,
        valuesStartRow: 2,
        skipHiddenRows: true,
        skipHiddenColumns: true,
        extractFormulaResults: true,
        preserveStyles: true,
        treatMergedRegions: 'top-left',
        inlineCodeRegexes: ['\\{\\d+\\}'],
        sourceLocale: 'en',
      }],
    },
  };
}

test('JSON roundtrip: placeholders meta retained, merge rehydrates and preserves style', async () => {
  const tmpDir = path.join(process.cwd(), '.out');
  if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true });
  const input = await createWorkbook(tmpDir);
  const cfg = inlineCfg();

  const units = await extract(input, cfg);
  const json = await exportUnitsToJson(units, cfg, { fileName: path.basename(input) });
  const obj = JSON.parse(json);
  const mutated: TranslationUnit[] = obj.units;
  // simulate translation: add targets with placeholder markers
  for (const tu of mutated) {
    if (tu.segments && tu.segments.length) {
      tu.segments = tu.segments.map(s => ({ ...s, target: s.source.replace('{0}', '[[ph:ph1]]') }));
    }
  }

  const out = path.join(tmpDir, 'sample-json.translated.xlsx');
  await merge(input, out, mutated, cfg);

  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.readFile(out);
  const ws2 = wb2.getWorksheet('Sheet1')!;
  const val = String(ws2.getCell('B2').value);
  expect(val).toContain('Hello {0}! This is v1.');
  expect(ws2.getCell('B2').font?.bold).toBe(true);
});
