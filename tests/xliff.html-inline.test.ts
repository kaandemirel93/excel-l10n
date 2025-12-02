import { exportToXliff } from '../src/exporter/xliff.js';
import { htmlToSkeleton, htmlToXliffInline } from '../src/io/excel.js';
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

  test('htmlToXliffInline converts HTML to XLIFF 2.1 inline elements', async () => {
    const html = '<div>Hi, <b>bold</b> text here.</div>';
    const result = await htmlToXliffInline(html, { xliffVersion: '2.1' });
    expect(result).toContain('<pc id="1" dataRef="html_b">bold</pc>');
    expect(result).toContain('Hi, ');
    expect(result).toContain(' text here.');
  });

  test('htmlToXliffInline converts HTML to XLIFF 1.2 inline elements', async () => {
    const html = '<div>Hi, <b>bold</b> text here.</div>';
    const result = await htmlToXliffInline(html, { xliffVersion: '1.2' });
    expect(result).toContain('<g id="1" ctype="bold">bold</g>');
    expect(result).toContain('Hi, ');
    expect(result).toContain(' text here.');
  });

  test('XLIFF export includes inline elements directly in source', async () => {
    const unit: TranslationUnit = {
      id: 'Sheet1::R4CB',
      sheetName: 'Sheet1',
      row: 4,
      col: 'B',
      colIndex: 2,
      source: 'Hi, <pc id="1" dataRef="html_b">bold</pc> text here.',
      segments: [{ id: 'Sheet1::R4CB_s0', source: 'Hi, <pc id="1" dataRef="html_b">bold</pc> text here.' }],
      meta: {
        htmlOriginal: '<div>Hi, <b>bold</b> text here.</div>'
      }
    } as any;
    const config: any = { workbook: { sheets: [{ namePattern: 'Sheet1' }] }, global: { srcLang: 'en', xliffVersion: '2.1' } };
    const xlf = await exportToXliff([unit], config);

    // Verify XLIFF 2.1 version
    expect(xlf).toContain('version="2.1"');
    // Verify inline elements are in the source
    expect(xlf).toContain('<pc id="1" dataRef="html_b">bold</pc>');
    // Verify NO htmlSkeleton or htmlInlineMap notes
    expect(xlf).not.toContain('htmlSkeleton');
    expect(xlf).not.toContain('htmlInlineMap');
  });

  test('XLIFF 1.2 export uses <g> elements', async () => {
    const unit: TranslationUnit = {
      id: 'Sheet1::R4CB',
      sheetName: 'Sheet1',
      row: 4,
      col: 'B',
      colIndex: 2,
      source: 'Hi, <g id="1" ctype="bold">bold</g> text here.',
      segments: [{ id: 'Sheet1::R4CB_s0', source: 'Hi, <g id="1" ctype="bold">bold</g> text here.' }],
    } as any;
    const config: any = { workbook: { sheets: [{ namePattern: 'Sheet1' }] }, global: { srcLang: 'en', xliffVersion: '1.2' } };
    const xlf = await exportToXliff([unit], config);

    // Verify XLIFF 1.2 version
    expect(xlf).toContain('version="1.2"');
    // Verify inline elements are in the source
    expect(xlf).toContain('<g id="1" ctype="bold">bold</g>');
  });
});
