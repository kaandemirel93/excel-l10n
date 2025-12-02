import { exportToXliff, parseXliffToUnits } from '../src/exporter/xliff.js';
import { htmlToSkeleton, htmlToXliffInline, htmlToXliffWithSkeleton } from '../src/io/excel.js';
import { TranslationUnit } from '../src/types.js';

describe('HTML inline markers in skeleton and XLIFF', () => {
  test('legacy: skeleton contains io/ic with htmltxt tokens', async () => {
    const html = '<div>Hi, <b>bold</b> text here.</div>';
    const { skeleton, texts, inlineMap } = await htmlToSkeleton(html);
    expect(skeleton).toBe('<div>[[htmltxt:1]][[io:1]][[htmltxt:2]][[ic:1]][[htmltxt:3]]</div>');
    expect(texts).toEqual(['Hi, ', 'bold', ' text here.']);
    expect(inlineMap[1].open).toBe('<b>');
    expect(inlineMap[1].close).toBe('</b>');
  });

  test('legacy: htmlToXliffInline converts HTML to XLIFF 2.1 inline elements', async () => {
    const html = '<div>Hi, <b>bold</b> text here.</div>';
    const result = await htmlToXliffInline(html, { xliffVersion: '2.1' });
    expect(result).toContain('<pc id="1" dataRef="html_b">bold</pc>');
    expect(result).toContain('Hi, ');
    expect(result).toContain(' text here.');
  });

  test('legacy: htmlToXliffInline converts HTML to XLIFF 1.2 inline elements', async () => {
    const html = '<div>Hi, <b>bold</b> text here.</div>';
    const result = await htmlToXliffInline(html, { xliffVersion: '1.2' });
    expect(result).toContain('<g id="1" ctype="bold">bold</g>');
    expect(result).toContain('Hi, ');
    expect(result).toContain(' text here.');
  });

  test('htmlToXliffWithSkeleton: XLIFF 2.1 with equivStart/equivEnd for attributes', async () => {
    const html = '<div>Visit <a href="/shop" class="link">our shop</a> today.</div>';
    const { skeleton, xliffSource, inlineMap } = await htmlToXliffWithSkeleton(html, { xliffVersion: '2.1' });
    
    // Skeleton should have block tag with placeholder
    expect(skeleton).toBe('<div>[[CONTENT]]</div>');
    
    // XLIFF source should have <pc> with equivStart/equivEnd preserving attributes
    expect(xliffSource).toContain('<pc id="1" dataRef="html_a"');
    expect(xliffSource).toContain('equivStart="&lt;a href=&quot;/shop&quot; class=&quot;link&quot;&gt;"');
    expect(xliffSource).toContain('equivEnd="&lt;/a&gt;"');
    expect(xliffSource).toContain('our shop</pc>');
    
    // No inlineMap for XLIFF 2.1 (uses equivStart/equivEnd)
    expect(inlineMap).toBeUndefined();
  });

  test('htmlToXliffWithSkeleton: XLIFF 1.2 with inlineMap for attributes', async () => {
    const html = '<div>Visit <a href="/shop" class="link">our shop</a> today.</div>';
    const { skeleton, xliffSource, inlineMap } = await htmlToXliffWithSkeleton(html, { xliffVersion: '1.2' });
    
    // Skeleton should have block tag with placeholder
    expect(skeleton).toBe('<div>[[CONTENT]]</div>');
    
    // XLIFF source should have <g> element
    expect(xliffSource).toContain('<g id="1" ctype="link">our shop</g>');
    
    // InlineMap should preserve full HTML with attributes for XLIFF 1.2
    expect(inlineMap).toBeDefined();
    expect(inlineMap![1].open).toBe('<a href="/shop" class="link">');
    expect(inlineMap![1].close).toBe('</a>');
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

  test('XLIFF 2.1 roundtrip: attributes preserved via equivStart/equivEnd', async () => {
    const unit: TranslationUnit = {
      id: 'test1',
      sheetName: 'Sheet1',
      row: 1,
      col: 'A',
      colIndex: 1,
      source: 'Visit <pc id="1" dataRef="html_a" equivStart="&lt;a href=&quot;/shop&quot;&gt;" equivEnd="&lt;/a&gt;">our shop</pc> today.',
      segments: [{ 
        id: 'test1_s0', 
        source: 'Visit <pc id="1" dataRef="html_a" equivStart="&lt;a href=&quot;/shop&quot;&gt;" equivEnd="&lt;/a&gt;">our shop</pc> today.' 
      }],
      meta: {
        htmlSkeleton: '<div>[[CONTENT]]</div>'
      }
    } as any;
    
    const config: any = { 
      workbook: { sheets: [{ namePattern: 'Sheet1' }] }, 
      global: { srcLang: 'en', xliffVersion: '2.1' } 
    };
    
    // Export to XLIFF
    const xlf = await exportToXliff([unit], config);
    expect(xlf).toContain('equivStart="&lt;a href=&quot;/shop&quot;&gt;"');
    
    // Parse back
    const parsed = parseXliffToUnits(xlf);
    expect(parsed.length).toBe(1);
    expect(parsed[0].source).toContain('<a href="/shop">our shop</a>');
    expect(parsed[0].source).not.toContain('<pc');
  });

  test('XLIFF 1.2 roundtrip: attributes preserved via inlineMap', async () => {
    const unit: TranslationUnit = {
      id: 'test1',
      sheetName: 'Sheet1',
      row: 1,
      col: 'A',
      colIndex: 1,
      source: 'Visit <g id="1" ctype="link">our shop</g> today.',
      segments: [{ 
        id: 'test1_s0', 
        source: 'Visit <g id="1" ctype="link">our shop</g> today.' 
      }],
      meta: {
        htmlSkeleton: '<div>[[CONTENT]]</div>',
        htmlInlineMap: {
          '1': { open: '<a href="/shop">', close: '</a>' }
        }
      }
    } as any;
    
    const config: any = { 
      workbook: { sheets: [{ namePattern: 'Sheet1' }] }, 
      global: { srcLang: 'en', xliffVersion: '1.2' } 
    };
    
    // Export to XLIFF
    const xlf = await exportToXliff([unit], config);
    expect(xlf).toContain('<g id="1" ctype="link">our shop</g>');
    expect(xlf).toContain('category="htmlInlineMap"');
    
    // Parse back
    const parsed = parseXliffToUnits(xlf);
    expect(parsed.length).toBe(1);
    // XLIFF 1.2 parser converts <g> to <a> but without attributes (needs inlineMap for merge)
    expect(parsed[0].source).toContain('<a>our shop</a>');
    expect(parsed[0].meta?.htmlInlineMap).toBeDefined();
    expect((parsed[0].meta as any).htmlInlineMap['1'].open).toBe('<a href="/shop">');
  });

  test('Nested inline tags work in both XLIFF versions', async () => {
    const html = '<p>This is <b>bold and <i>italic</i> text</b> here.</p>';
    
    // Test XLIFF 2.1
    const result21 = await htmlToXliffWithSkeleton(html, { xliffVersion: '2.1' });
    expect(result21.xliffSource).toContain('<pc id="1"');
    expect(result21.xliffSource).toContain('<pc id="2"');
    expect(result21.skeleton).toBe('<p>[[CONTENT]]</p>');
    
    // Test XLIFF 1.2
    const result12 = await htmlToXliffWithSkeleton(html, { xliffVersion: '1.2' });
    expect(result12.xliffSource).toContain('<g id="1"');
    expect(result12.xliffSource).toContain('<g id="2"');
    expect(result12.skeleton).toBe('<p>[[CONTENT]]</p>');
    expect(result12.inlineMap).toBeDefined();
  });
});
