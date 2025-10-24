import { exportToXliff } from '../src/exporter/xliff.js';
import { htmlToSkeleton } from '../src/io/excel.js';
import { TranslationUnit } from '../src/types.js';

describe('HTML inline markers in skeleton and XLIFF', () => {
  test('skeleton contains io/ic with htmltxt tokens', async () => {
    const html = '<div>Hi, <b>bold</b> text here.</div>';
    const { skeleton, texts, inlineMap } = await htmlToSkeleton(html);
    expect(skeleton).toBe('<div>[[htmltxt:1]][[io:1]][[htmltxt:2]][[ic:1]][[htmltxt:3]]</div>');
    expect(texts).toEqual(['Hi, ', 'bold', ' text here.']);
    expect(inlineMap[1].open).toBe('<b>');
    expect(inlineMap[1].close).toBe('</b>');
  });

  test('XLIFF notes include htmlSkeleton with inline markers', async () => {
    const unit: TranslationUnit = {
      id: 'Sheet1::R4CB',
      sheetName: 'Sheet1',
      row: 4,
      col: 'B',
      colIndex: 2,
      source: 'Hi, [[io:1]]bold[[ic:1]] text here.',
      segments: [{ id: 'Sheet1::R4CB_s0', source: 'Hi, [[io:1]]bold[[ic:1]] text here.' }],
      meta: {
        htmlSkeleton: '<div>[[htmltxt:1]][[io:1]][[htmltxt:2]][[ic:1]][[htmltxt:3]]</div>',
        htmlInlineMap: { '1': { open: '<b>', close: '</b>' } }
      }
    } as any;
    const config: any = { workbook: { sheets: [{ namePattern: 'Sheet1' }] }, global: { srcLang: 'en' } };
    const xlf = await exportToXliff([unit], config);
    expect(xlf).toMatch(/<note category="htmlSkeleton">&lt;div&gt;\[\[htmltxt:1\]\]\[\[io:1\]\]\[\[htmltxt:2\]\]\[\[ic:1\]\]\[\[htmltxt:3\]\]&lt;\/div&gt;<\/note>/);
  });
});
