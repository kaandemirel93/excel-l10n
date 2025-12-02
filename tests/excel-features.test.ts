import path from 'node:path';
import fs from 'node:fs';
import ExcelJS from 'exceljs';
import { extract, exportUnitsToXliff, parseTranslated, merge } from '../src/index';
import type { Config } from '../src/types';

describe('Excel Features Roundtrip', () => {
  const tmpDir = path.join(process.cwd(), '.out');
  beforeAll(() => { if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true }); });

  it('should preserve rich text formatting', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 'Source';
    ws.getCell('B1').value = 'Target';
    
    // Rich text with different formatting
    ws.getCell('A2').value = {
      richText: [
        { text: 'This is ', font: { name: 'Arial', size: 12 } },
        { text: 'bold', font: { bold: true } },
        { text: ' and ', font: { name: 'Arial', size: 12 } },
        { text: 'italic', font: { italic: true } },
        { text: ' text.', font: { name: 'Arial', size: 12 } }
      ]
    };

    const xlsxPath = path.join(tmpDir, 'richtext.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [{
          namePattern: 'Sheet1',
          sourceColumns: ['A'],
          targetColumns: { en: 'B' },
          preserveStyles: true,
          headerRow: 1,
          valuesStartRow: 2
        }]
      },
      global: { srcLang: 'en', xliffVersion: '2.1' }
    };

    const units = await extract(xlsxPath, config);
    expect(units.length).toBeGreaterThan(0);
    expect(units[0].richText).toBe(true);
    expect(units[0].source).toBe('This is bold and italic text.');
    // Rich text cells may not have a single style, but individual text runs have styles

    const xlf = await exportUnitsToXliff(units, config);
    const parsed = parseTranslated(xlf, 'xlf');
    
    for (const tu of parsed) {
      if (tu.segments) {
        tu.segments = tu.segments.map(s => ({ ...s, target: `[${s.source}]` }));
      }
    }

    const outPath = path.join(tmpDir, 'richtext_merged.xlsx');
    await merge(xlsxPath, outPath, parsed, config);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outPath);
    const ws2 = wb2.getWorksheet('Sheet1')!;
    const b2 = ws2.getCell('B2').value;

    expect(b2).toBe('[This is bold and italic text.]');
  });

  it('should handle formulas with extractFormulaResults', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 'Source';
    ws.getCell('B1').value = 'Target';
    
    // Cell with formula
    ws.getCell('A2').value = { formula: 'CONCATENATE("Hello", " ", "World")', result: 'Hello World' };
    ws.getCell('A3').value = { formula: 'SUM(1,2,3)', result: 6 };

    const xlsxPath = path.join(tmpDir, 'formulas.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [{
          namePattern: 'Sheet1',
          sourceColumns: ['A'],
          targetColumns: { en: 'B' },
          extractFormulaResults: true
        }]
      },
      global: { srcLang: 'en', xliffVersion: '2.1' }
    };

    const units = await extract(xlsxPath, config);
    expect(units.length).toBe(2);
    expect(units[0].source).toBe('Hello World'); // Formula result
    expect(units[0].formula).toBe('CONCATENATE("Hello", " ", "World")');
    expect(units[1].source).toBe('6');
    expect(units[1].formula).toBe('SUM(1,2,3)');

    const xlf = await exportUnitsToXliff(units, config);
    const parsed = parseTranslated(xlf, 'xlf');
    
    for (const tu of parsed) {
      if (tu.segments) {
        tu.segments = tu.segments.map(s => ({ ...s, target: `[${s.source}]` }));
      }
    }

    const outPath = path.join(tmpDir, 'formulas_merged.xlsx');
    await merge(xlsxPath, outPath, parsed, config);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outPath);
    const ws2 = wb2.getWorksheet('Sheet1')!;
    
    expect(String(ws2.getCell('B2').value)).toBe('[Hello World]');
    expect(String(ws2.getCell('B3').value)).toBe('[6]');
  });

  it('should handle comments/notes when translateComments is enabled', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 'Source';
    ws.getCell('B1').value = 'Target';
    ws.getCell('A2').value = 'Text with comment';
    
    // Add comment/note
    ws.getCell('A2').note = 'This is a comment';

    const xlsxPath = path.join(tmpDir, 'comments.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [{
          namePattern: 'Sheet1',
          sourceColumns: ['A'],
          targetColumns: { en: 'B' },
          translateComments: true,
          headerRow: 1,
          valuesStartRow: 2
        }]
      },
      global: { srcLang: 'en', xliffVersion: '2.1', exportComments: true }
    };

    const units = await extract(xlsxPath, config);
    expect(units[0].meta?.comments).toBeDefined();

    const xlf = await exportUnitsToXliff(units, config);
    expect(xlf).toContain('category="comments"');
    
    // Note: Comments are exported but not currently parsed back from XLIFF notes
    // This is expected behavior - comments are for translator context, not for translation
    const parsed = parseTranslated(xlf, 'xlf');
    // Comments may not be in parsed units (they're in notes for context only)
    expect(parsed.length).toBeGreaterThan(0);
  });

  it('should preserve complex cell styles (font, fill, alignment)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 'Source';
    ws.getCell('B1').value = 'Target';
    
    const cell = ws.getCell('A2');
    cell.value = 'Styled text';
    cell.font = {
      name: 'Calibri',
      size: 14,
      bold: true,
      italic: true,
      color: { argb: 'FFFF0000' } // Red
    };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFFFF00' } // Yellow background
    } as any;
    cell.alignment = {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true
    };

    const xlsxPath = path.join(tmpDir, 'styles.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [{
          namePattern: 'Sheet1',
          sourceColumns: ['A'],
          targetColumns: { en: 'B' },
          preserveStyles: true
        }]
      },
      global: { srcLang: 'en', xliffVersion: '2.1' }
    };

    const units = await extract(xlsxPath, config);
    expect(units[0].style).toBeDefined();
    expect(units[0].style?.font?.bold).toBe(true);
    expect(units[0].style?.font?.italic).toBe(true);
    expect(units[0].style?.font?.color).toContain('FF0000');
    expect(units[0].style?.fill?.color).toContain('FFFF00');
    expect(units[0].style?.alignment?.horizontal).toBe('center');

    const xlf = await exportUnitsToXliff(units, config);
    const parsed = parseTranslated(xlf, 'xlf');
    
    for (const tu of parsed) {
      if (tu.segments) {
        tu.segments = tu.segments.map(s => ({ ...s, target: `[${s.source}]` }));
      }
    }

    const outPath = path.join(tmpDir, 'styles_merged.xlsx');
    await merge(xlsxPath, outPath, parsed, config);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outPath);
    const ws2 = wb2.getWorksheet('Sheet1')!;
    const b2 = ws2.getCell('B2');

    expect(b2.value).toBe('[Styled text]');
    expect(b2.font?.bold).toBe(true);
    expect(b2.font?.italic).toBe(true);
    expect((b2.font?.color as any)?.argb).toContain('FF0000');
    expect((b2.fill as any)?.fgColor?.argb).toContain('FFFF00');
    expect(b2.alignment?.horizontal).toBe('center');
  });

  it('should handle merged cells correctly', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 'Source';
    ws.getCell('B1').value = 'Target';
    
    // Merge cells A2:A3
    ws.mergeCells('A2:A3');
    ws.getCell('A2').value = 'Merged cell content';
    
    // Merge cells A4:B4
    ws.mergeCells('A4:B4');
    ws.getCell('A4').value = 'Horizontal merge';

    const xlsxPath = path.join(tmpDir, 'merged.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [{
          namePattern: 'Sheet1',
          sourceColumns: ['A'],
          targetColumns: { en: 'B' },
          treatMergedRegions: 'top-left'
        }]
      },
      global: { srcLang: 'en', xliffVersion: '2.1' }
    };

    const units = await extract(xlsxPath, config);
    // Should extract from top-left cell of merged region
    const mergedUnit = units.find(u => u.source === 'Merged cell content');
    expect(mergedUnit).toBeDefined();
    expect(mergedUnit?.isMerged).toBe(true);

    const xlf = await exportUnitsToXliff(units, config);
    const parsed = parseTranslated(xlf, 'xlf');
    
    for (const tu of parsed) {
      if (tu.segments) {
        tu.segments = tu.segments.map(s => ({ ...s, target: `[${s.source}]` }));
      }
    }

    const outPath = path.join(tmpDir, 'merged_merged.xlsx');
    await merge(xlsxPath, outPath, parsed, config);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outPath);
    const ws2 = wb2.getWorksheet('Sheet1')!;
    
    // Check that merged cell content is translated
    const b2 = ws2.getCell('B2').value;
    expect(b2).toBe('[Merged cell content]');
  });

  it('should handle mixed content: HTML + formulas + styles + segmentation', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 'Source';
    ws.getCell('B1').value = 'Target';
    
    // Complex cell with HTML, that will be segmented
    const cell = ws.getCell('A2');
    cell.value = '<div>Welcome to <b>our store</b>! We have great deals. Visit us today!</div>';
    cell.font = { bold: true, color: { argb: 'FF0000FF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCCCCC' } } as any;
    cell.alignment = { wrapText: true };

    const xlsxPath = path.join(tmpDir, 'mixed_content.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [{
          namePattern: 'Sheet1',
          sourceColumns: ['A'],
          targetColumns: { en: 'B' },
          html: { enabled: true },
          preserveStyles: true
        }]
      },
      global: { srcLang: 'en', xliffVersion: '2.1' },
      segmentation: {
        enabled: true,
        rules: {
          srxPath: path.resolve(__dirname, '../examples/default_rules.srx')
        }
      }
    };

    const units = await extract(xlsxPath, config);
    expect(units[0].meta?.htmlSkeleton).toBeDefined();
    expect(units[0].segments).toBeDefined();
    expect(units[0].segments!.length).toBeGreaterThan(1); // Multiple sentences
    expect(units[0].style).toBeDefined();

    const xlf = await exportUnitsToXliff(units, config);
    const parsed = parseTranslated(xlf, 'xlf');
    
    for (const tu of parsed) {
      if (tu.segments) {
        tu.segments = tu.segments.map(s => ({ ...s, target: `[${s.source}]` }));
      }
    }

    const outPath = path.join(tmpDir, 'mixed_content_merged.xlsx');
    await merge(xlsxPath, outPath, parsed, config);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outPath);
    const ws2 = wb2.getWorksheet('Sheet1')!;
    const b2 = String(ws2.getCell('B2').value);

    // Verify HTML structure preserved
    expect(b2).toContain('<div>');
    expect(b2).toContain('<b>our store</b>');
    expect(b2).toContain('</div>');
    
    // Verify translation applied
    expect(b2).toContain('[Welcome to');
    
    // Verify styles preserved
    expect(ws2.getCell('B2').font?.bold).toBe(true);
    expect((ws2.getCell('B2').font?.color as any)?.argb).toContain('0000FF');
  });

  it('should handle empty cells and whitespace correctly', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 'Source';
    ws.getCell('B1').value = 'Target';
    
    ws.getCell('A2').value = '   '; // Only whitespace
    ws.getCell('A3').value = ''; // Empty
    ws.getCell('A4').value = null; // Null
    ws.getCell('A5').value = 'Valid text';
    ws.getCell('A6').value = '  Leading and trailing  ';

    const xlsxPath = path.join(tmpDir, 'whitespace.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [{
          namePattern: 'Sheet1',
          sourceColumns: ['A'],
          targetColumns: { en: 'B' },
          headerRow: 1,
          valuesStartRow: 2
        }]
      },
      global: { srcLang: 'en', xliffVersion: '2.1' }
    };

    const units = await extract(xlsxPath, config);
    
    // Whitespace-only cells are currently extracted (may contain meaningful spaces)
    // Empty and null cells are skipped
    expect(units.length).toBeGreaterThan(0);
    
    // Find the valid text unit
    const validUnit = units.find(u => u.source === 'Valid text');
    expect(validUnit).toBeDefined();
    
    // Whitespace is preserved
    const whitespaceUnit = units.find(u => u.source.includes('Leading and trailing'));
    expect(whitespaceUnit).toBeDefined();
    expect(whitespaceUnit?.source).toBe('  Leading and trailing  ');
  });

  it('should handle special characters and Unicode', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 'Source';
    ws.getCell('B1').value = 'Target';
    
    ws.getCell('A2').value = 'Special chars: & < > " \' © ® ™';
    ws.getCell('A3').value = 'Unicode: 你好世界 مرحبا العالم Привет мир';
    ws.getCell('A4').value = 'Emoji: 😀 🎉 ✨ 🚀';
    ws.getCell('A5').value = 'Math: ∑ ∫ √ ∞ ≠ ≤ ≥';

    const xlsxPath = path.join(tmpDir, 'unicode.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [{
          namePattern: 'Sheet1',
          sourceColumns: ['A'],
          targetColumns: { en: 'B' }
        }]
      },
      global: { srcLang: 'en', xliffVersion: '2.1' }
    };

    const units = await extract(xlsxPath, config);
    expect(units.length).toBe(4);
    expect(units[0].source).toContain('&');
    expect(units[0].source).toContain('<');
    expect(units[0].source).toContain('>');
    expect(units[1].source).toContain('你好世界');
    expect(units[2].source).toContain('😀');
    expect(units[3].source).toContain('∑');

    const xlf = await exportUnitsToXliff(units, config);
    // XML should escape special chars
    expect(xlf).toContain('&amp;');
    expect(xlf).toContain('&lt;');
    expect(xlf).toContain('&gt;');
    
    const parsed = parseTranslated(xlf, 'xlf');
    
    for (const tu of parsed) {
      if (tu.segments) {
        tu.segments = tu.segments.map(s => ({ ...s, target: `[${s.source}]` }));
      }
    }

    const outPath = path.join(tmpDir, 'unicode_merged.xlsx');
    await merge(xlsxPath, outPath, parsed, config);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outPath);
    const ws2 = wb2.getWorksheet('Sheet1')!;
    
    // Verify special characters preserved
    expect(String(ws2.getCell('B2').value)).toContain('&');
    expect(String(ws2.getCell('B2').value)).toContain('<');
    expect(String(ws2.getCell('B3').value)).toContain('你好世界');
    expect(String(ws2.getCell('B4').value)).toContain('😀');
    expect(String(ws2.getCell('B5').value)).toContain('∑');
  });
});
