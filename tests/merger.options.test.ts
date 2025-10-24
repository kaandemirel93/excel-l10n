import path from 'node:path';
import fs from 'node:fs';
import ExcelJS from 'exceljs';
import { merge } from '../src/index';
import type { Config, TranslationUnit } from '../src/types';

async function createWB(tmp: string): Promise<string> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.getCell('A1').value = 'Header';
  ws.getCell('A2').value = 'Source';
  ws.getCell('C2').value = 'Existing';
  const file = path.join(tmp, 'merger-opt.xlsx');
  await wb.xlsx.writeFile(file);
  return file;
}

test('overwrite=false prevents writing, createTargetIfMissing inserts column after source', async () => {
  const tmpDir = path.join(process.cwd(), '.out');
  if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true });
  const input = await createWB(tmpDir);

  const cfg: Config = {
    global: { overwrite: false, insertTargetPlacement: 'insertAfterSource' },
    workbook: {
      sheets: [{ namePattern: 'Sheet1', sourceColumns: ['A'], targetColumns: { fr: '' }, createTargetIfMissing: true, headerRow: 1, valuesStartRow: 2, preserveStyles: true }],
    },
  };

  const tus: TranslationUnit[] = [
    { id: 'Sheet1::R2CA', sheetName: 'Sheet1', row: 2, col: 'A', colIndex: 1, source: 'Source', segments: [{ id: 'Sheet1::R2CA_s0', source: 'Source', target: 'Traduction' }] },
  ];

  const out = path.join(tmpDir, 'merger-opt.out.xlsx');
  await merge(input, out, tus, cfg);

  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.readFile(out);
  const ws2 = wb2.getWorksheet('Sheet1')!;
  // overwrite=false should keep newly created target empty (but column should exist)
  // With insertAfterSource, new column should be B, previously C becomes D
  expect(ws2.getCell('B2').value).toBe(null);
  expect(ws2.getCell('D2').value).toBe('Existing');
});
