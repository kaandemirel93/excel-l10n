import { segmentUnits } from '../src/segmenter/index';
import type { Config, TranslationUnit } from '../src/types';

test('SRX/builtin segmentation splits sentences', () => {
  const cfg: Config = {
    global: { srcLang: 'en' },
    segmentation: { enabled: true, rules: 'builtin' },
    workbook: { sheets: [{ namePattern: 'Sheet1', sourceColumns: ['A'] }] },
  };
  const units: TranslationUnit[] = [
    { id: 'Sheet1::R2CA', sheetName: 'Sheet1', row: 2, col: 'A', colIndex: 1, source: 'Hello world. How are you?', segments: [] },
  ];
  const out = segmentUnits(units, cfg);
  expect(out[0].segments && out[0].segments!.length).toBeGreaterThan(1);
});
