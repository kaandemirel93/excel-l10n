import { exportToXliff } from '../src/exporter/xliff.js';
import { TranslationUnit } from '../src/types.js';

describe('XLIFF whitespace handling', () => {
    test('should not add extra newlines in mixed content source', async () => {
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

        // Check that the source content is on a single line or at least doesn't have newlines around the inline tag
        // The issue reported is that it looks like:
        // <source>
        //   Hi,
        //   <pc ...>bold</pc>
        //   text here.
        // </source>

        // We want: <source>Hi, <pc ...>bold</pc> text here.</source>

        // Check for the specific unwanted pattern (newline + spaces + <pc)
        expect(xlf).not.toMatch(/Hi,\s*\n\s*<pc/);
        expect(xlf).not.toMatch(/<\/pc>\s*\n\s*text/);

        // It should match the compact form (with xml:space="preserve")
        expect(xlf).toContain('<source xml:space="preserve">Hi, <pc id="1" dataRef="html_b">bold</pc> text here.</source>');
    });
});
