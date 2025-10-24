import path from 'node:path';
import fs from 'node:fs';
import ExcelJS from 'exceljs';
import { extract } from '../src/index';
import type { Config } from '../src/types';

async function makeWB(tmp: string): Promise<string> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.getCell('A1').value = 'HeaderA';
  ws.getRow(1).hidden = true; // hidden header row
  ws.getCell('A2').value = 'Value';
  ws.getCell('A3').value = 'Meta Row';
  ws.getRow(3).hidden = true; // hidden metadata row
  const file = path.join(tmp, 'hidden-meta.xlsx');
  await wb.xlsx.writeFile(file);
  return file;
}

test('headerName and metadataRows captured even if rows are hidden', async () => {
  const tmpDir = path.join(process.cwd(), '.out');
  if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true });
  const input = await makeWB(tmpDir);

  const cfg: Config = {
    global: { srcLang: 'en' },
    workbook: {
      sheets: [{
        namePattern: 'Sheet1',
        sourceColumns: ['A'],
        headerRow: 1,
        valuesStartRow: 2,
        metadataRows: [3],
        skipHiddenRows: true, // should still allow reading header/meta rows for context
      }],
    },
  };

  const tus = await extract(input, cfg);
  const u = tus.find(t => t.row === 2 && t.col === 'A')!;
  expect(u.meta?.headerName).toBe('HeaderA');
  expect(u.meta?.metadataRows && u.meta?.metadataRows[3]).toBe('Meta Row');
});
