import path from 'node:path';
import fs from 'node:fs';
import ExcelJS from 'exceljs';
import { extract } from '../src/index';
import type { Config } from '../src/types';

async function makeBookHiddenAndColors(tmp: string): Promise<string> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  // header
  ws.getCell('A1').value = 'HeaderA';
  ws.getCell('B1').value = 'HeaderB';
  // values
  ws.getCell('A2').value = 'Visible';
  ws.getCell('A3').value = 'HiddenRow';
  ws.getRow(3).hidden = true;
  ws.getCell('B2').value = 'HiddenCol';
  ws.getColumn(2).hidden = true;
  ws.getCell('A4').value = 'RedText';
  ws.getCell('A4').font = { color: { argb: 'FFFF0000' } };
  // merged region A5:B5
  ws.mergeCells('A5:B5');
  ws.getCell('A5').value = 'MergedTopLeft';

  const file = path.join(tmp, 'extractor-opt.xlsx');
  await wb.xlsx.writeFile(file);
  return file;
}

test('extractor respects hidden rows/columns, color exclusion and top-left merged policy', async () => {
  const tmpDir = path.join(process.cwd(), '.out');
  if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true });
  const input = await makeBookHiddenAndColors(tmpDir);

  const cfg: Config = {
    global: { srcLang: 'en' },
    workbook: {
      sheets: [{
        namePattern: 'Sheet1',
        sourceColumns: ['A','B'],
        headerRow: 1,
        valuesStartRow: 2,
        skipHiddenRows: true,
        skipHiddenColumns: true,
        excludeColors: ['#FF0000'],
        treatMergedRegions: 'top-left',
      }],
    },
  };

  const tus = await extract(input, cfg);
  // Should include A2, exclude row 3 (hidden), exclude B2 (hidden col), exclude A4 (red)
  const ids = tus.map(t => t.id);
  expect(ids).toContain('Sheet1::R2CA');
  expect(ids).not.toContain('Sheet1::R3CA');
  expect(ids).not.toContain('Sheet1::R2CB');
  expect(ids).not.toContain('Sheet1::R4CA');
  // merged: only top-left A5 should be included (B5 is same region but not included)
  expect(ids).toContain('Sheet1::R5CA');
  expect(ids).not.toContain('Sheet1::R5CB');
});
