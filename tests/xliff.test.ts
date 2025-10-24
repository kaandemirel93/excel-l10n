import { exportToXliff } from '../src/exporter/xliff';
import type { Config, TranslationUnit } from '../src/types';

test('XLIFF export includes header/comments notes and placeholder note when configured', async () => {
  const cfg: Config = {
    global: { srcLang: 'en', exportComments: true },
    workbook: { sheets: [{ namePattern: 'Sheet1', sourceColumns: ['A'], inlineCodeRegexes: ['\\{\\d+\\}'] }] },
  };
  const units: TranslationUnit[] = [
    {
      id: 'Sheet1::R2CA', sheetName: 'Sheet1', row: 2, col: 'A', colIndex: 1,
      source: 'Hello {0}! This is v1.',
      segments: [{ id: 'Sheet1::R2CA_s0', source: 'Hello {0}! This is v1.' }],
      meta: { headerName: 'Header', comments: 'note' },
    },
  ];
  const xlf = await exportToXliff(units, cfg, { srcLang: 'en' });
  expect(xlf).toContain('<notes>');
  expect(xlf).toContain('category="header"');
  expect(xlf).toContain('category="comments"');
  expect(xlf).toContain('<ph');
  expect(xlf).toContain('category="ph"');
});
