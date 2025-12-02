import { exportToXliff, parseXliffToUnits } from '../src/exporter/xliff.js';
import { htmlToXliffInline } from '../src/io/excel.js';
import { TranslationUnit } from '../src/types.js';

describe('XLIFF merge bugs', () => {
    test('should not include attributes in flattened text (bold1html bug)', () => {
        // Simulate the parsed structure that caused the bug
        // <pc id="1" dataRef="html_b">bold</pc>
        // parsed by fast-xml-parser with ignoreAttributes: false
        const node = {
            id: '1',
            dataRef: 'html_b',
            '#text': 'bold'
        };

        // We can't easily access the internal 'flatten' function, but we can test parseXliffToUnits
        // which uses it.
        const xlf = `<?xml version="1.0" encoding="UTF-8"?>
<xliff version="2.1" srcLang="en">
  <file id="f1">
    <unit id="u1">
      <segment>
        <source>Hi, <pc id="1" dataRef="html_b">bold</pc> text.</source>
      </segment>
    </unit>
  </file>
</xliff>`;

        const units = parseXliffToUnits(xlf);
        const source = units[0].source;

        // Before fix: "Hi, bold1html_b text." (or similar)
        // After fix: "Hi, <b>bold</b> text."
        expect(source).toBe('Hi, <b>bold</b> text.');
    });

    test('should preserve block tags using skeleton', async () => {
        const html = '<h1 class="title">Title</h1>';
        // htmlToXliffWithSkeleton preserves block-level tags in the skeleton
        const { htmlToXliffWithSkeleton } = await import('../src/io/excel.js');
        const { skeleton, xliffSource } = await htmlToXliffWithSkeleton(html, { xliffVersion: '2.1' });

        expect(skeleton).toContain('<h1 class="title">');
        expect(skeleton).toContain('[[CONTENT]]');
        expect(skeleton).toContain('</h1>');
        expect(xliffSource).toBe('Title');

        // When exported to XLIFF, the skeleton should be in a note
        const unit: TranslationUnit = {
            id: 'u1',
            sheetName: 's1',
            row: 1,
            col: 'A',
            colIndex: 1,
            source: xliffSource,
            segments: [{ id: 'u1_s0', source: xliffSource }],
            meta: {
                htmlSkeleton: skeleton
            }
        } as any;

        const config: any = { workbook: { sheets: [{ namePattern: 's1' }] }, global: { xliffVersion: '2.1' } };
        const xlf = await exportToXliff([unit], config);

        // The skeleton should be in a note with category="htmlSkeleton"
        expect(xlf).toContain('category="htmlSkeleton"');
        expect(xlf).toContain('[[CONTENT]]');
    });
});
