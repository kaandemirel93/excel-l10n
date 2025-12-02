import path from 'node:path';
import fs from 'node:fs';
import ExcelJS from 'exceljs';
import { extract, exportUnitsToXliff, parseTranslated, merge } from '../src/index';
import type { Config } from '../src/types';

describe('Production Scenarios', () => {
  const tmpDir = path.join(process.cwd(), '.out');
  beforeAll(() => { if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true }); });

  it('should handle multiple sheets with different configurations', async () => {
    const wb = new ExcelJS.Workbook();
    
    // Sheet 1: UI strings with HTML
    const ws1 = wb.addWorksheet('UI_Strings');
    ws1.getCell('A1').value = 'Key';
    ws1.getCell('B1').value = 'English';
    ws1.getCell('C1').value = 'French';
    ws1.getCell('B2').value = '<div>Welcome to <b>our app</b>!</div>';
    ws1.getCell('B3').value = 'Click <a href="/help">here</a> for help.';
    
    // Sheet 2: Marketing content with segmentation
    const ws2 = wb.addWorksheet('Marketing');
    ws2.getCell('A1').value = 'Content';
    ws2.getCell('B1').value = 'Translation';
    ws2.getCell('A2').value = 'Our product is amazing. It will change your life. Try it today!';
    ws2.getCell('A3').value = 'Special offer: 50% off everything. Limited time only!';
    
    // Sheet 3: Error messages (no HTML, no segmentation)
    const ws3 = wb.addWorksheet('Errors');
    ws3.getCell('A1').value = 'Code';
    ws3.getCell('B1').value = 'Message';
    ws3.getCell('C1').value = 'Translation';
    ws3.getCell('B2').value = 'File not found';
    ws3.getCell('B3').value = 'Invalid input';

    const xlsxPath = path.join(tmpDir, 'multi_sheet.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [
          {
            namePattern: 'UI_Strings',
            sourceColumns: ['B'],
            targetColumns: { fr: 'C' },
            html: { enabled: true },
            headerRow: 1,
            valuesStartRow: 2
          },
          {
            namePattern: 'Marketing',
            sourceColumns: ['A'],
            targetColumns: { fr: 'B' },
            headerRow: 1,
            valuesStartRow: 2
          },
          {
            namePattern: 'Errors',
            sourceColumns: ['B'],
            targetColumns: { fr: 'C' },
            headerRow: 1,
            valuesStartRow: 2
          }
        ]
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
    
    // Should extract from all sheets
    const uiUnits = units.filter(u => u.sheetName === 'UI_Strings');
    const marketingUnits = units.filter(u => u.sheetName === 'Marketing');
    const errorUnits = units.filter(u => u.sheetName === 'Errors');
    
    expect(uiUnits.length).toBe(2);
    expect(marketingUnits.length).toBe(2);
    expect(errorUnits.length).toBe(2);
    
    // UI strings should have HTML skeleton
    expect(uiUnits[0].meta?.htmlSkeleton).toBeDefined();
    
    // Marketing should have segments
    expect(marketingUnits[0].segments).toBeDefined();
    expect(marketingUnits[0].segments!.length).toBeGreaterThan(1);
    
    // Errors should be simple text
    expect(errorUnits[0].source).toBe('File not found');

    const xlf = await exportUnitsToXliff(units, config);
    const parsed = parseTranslated(xlf, 'xlf');
    
    for (const tu of parsed) {
      if (tu.segments) {
        tu.segments = tu.segments.map(s => ({ ...s, target: `[${s.source}]` }));
      }
    }

    const outPath = path.join(tmpDir, 'multi_sheet_merged.xlsx');
    await merge(xlsxPath, outPath, parsed, config);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outPath);
    
    // Verify all sheets were processed
    const ws1_out = wb2.getWorksheet('UI_Strings')!;
    const ws2_out = wb2.getWorksheet('Marketing')!;
    const ws3_out = wb2.getWorksheet('Errors')!;
    
    expect(String(ws1_out.getCell('C2').value)).toContain('<div>');
    expect(String(ws2_out.getCell('B2').value)).toContain('[Our product');
    expect(String(ws3_out.getCell('C2').value)).toBe('[File not found]');
  });

  it('should handle large files efficiently (1000+ rows)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('LargeSheet');
    ws.getCell('A1').value = 'ID';
    ws.getCell('B1').value = 'Source';
    ws.getCell('C1').value = 'Target';
    
    // Generate 1000 rows
    const rowCount = 1000;
    for (let i = 2; i <= rowCount + 1; i++) {
      ws.getCell(`A${i}`).value = `ID_${i - 1}`;
      ws.getCell(`B${i}`).value = `This is test string number ${i - 1}. It contains some text.`;
    }

    const xlsxPath = path.join(tmpDir, 'large_file.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [{
          namePattern: 'LargeSheet',
          sourceColumns: ['B'],
          targetColumns: { en: 'C' },
          headerRow: 1,
          valuesStartRow: 2
        }]
      },
      global: { srcLang: 'en', xliffVersion: '2.1' }
    };

    const startExtract = Date.now();
    const units = await extract(xlsxPath, config);
    const extractTime = Date.now() - startExtract;
    
    expect(units.length).toBe(rowCount);
    expect(extractTime).toBeLessThan(5000); // Should complete in under 5 seconds
    
    const startExport = Date.now();
    const xlf = await exportUnitsToXliff(units, config);
    const exportTime = Date.now() - startExport;
    
    expect(xlf).toContain('<xliff');
    expect(exportTime).toBeLessThan(3000); // Should complete in under 3 seconds
    
    const startParse = Date.now();
    const parsed = parseTranslated(xlf, 'xlf');
    const parseTime = Date.now() - startParse;
    
    expect(parsed.length).toBe(rowCount);
    expect(parseTime).toBeLessThan(3000); // Should complete in under 3 seconds
    
    // Add translations to a subset (simulating partial translation)
    for (let i = 0; i < 100; i++) {
      if (parsed[i].segments) {
        parsed[i].segments = parsed[i].segments!.map(s => ({ ...s, target: `[${s.source}]` }));
      }
    }

    const startMerge = Date.now();
    const outPath = path.join(tmpDir, 'large_file_merged.xlsx');
    await merge(xlsxPath, outPath, parsed, config);
    const mergeTime = Date.now() - startMerge;
    
    expect(mergeTime).toBeLessThan(5000); // Should complete in under 5 seconds

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outPath);
    const ws2 = wb2.getWorksheet('LargeSheet')!;
    
    // Verify first translated row
    expect(String(ws2.getCell('C2').value)).toContain('[This is test string');
    
    // Verify untranslated row - depending on mergeFallback config, it might have source or be empty
    const c500 = ws2.getCell('C500').value;
    // With default mergeFallback='source', untranslated cells get source text
    expect(c500 !== null).toBe(true);
  });

  it('should handle malformed HTML gracefully', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 'Source';
    ws.getCell('B1').value = 'Target';
    
    // Various malformed HTML scenarios
    ws.getCell('A2').value = '<div>Unclosed div';
    ws.getCell('A3').value = 'Unopened div</div>';
    ws.getCell('A4').value = '<b>Nested <i>tags</b></i>'; // Wrong nesting order
    ws.getCell('A5').value = '<div><b>Valid HTML</b></div>';
    ws.getCell('A6').value = '<script>alert("xss")</script>Normal text';
    ws.getCell('A7').value = '<<broken>>tags<<';

    const xlsxPath = path.join(tmpDir, 'malformed_html.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [{
          namePattern: 'Sheet1',
          sourceColumns: ['A'],
          targetColumns: { en: 'B' },
          html: { enabled: true },
          headerRow: 1,
          valuesStartRow: 2
        }]
      },
      global: { srcLang: 'en', xliffVersion: '2.1' }
    };

    // Should not throw errors
    const units = await extract(xlsxPath, config);
    expect(units.length).toBeGreaterThan(0);
    
    // Should handle valid HTML correctly
    const validUnit = units.find(u => u.source.includes('Valid HTML'));
    expect(validUnit).toBeDefined();
    expect(validUnit?.meta?.htmlSkeleton).toBeDefined();

    const xlf = await exportUnitsToXliff(units, config);
    expect(xlf).toContain('<xliff');
    
    const parsed = parseTranslated(xlf, 'xlf');
    expect(parsed.length).toBeGreaterThan(0);
    
    for (const tu of parsed) {
      if (tu.segments) {
        tu.segments = tu.segments.map(s => ({ ...s, target: `[${s.source}]` }));
      }
    }

    // Should not throw during merge
    const outPath = path.join(tmpDir, 'malformed_html_merged.xlsx');
    await merge(xlsxPath, outPath, parsed, config);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outPath);
    const ws2 = wb2.getWorksheet('Sheet1')!;
    
    // Valid HTML should be preserved
    const b5 = String(ws2.getCell('B5').value);
    expect(b5).toContain('<div>');
    expect(b5).toContain('<b>');
  });

  it('should handle mixed content types in same sheet', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Mixed');
    ws.getCell('A1').value = 'Type';
    ws.getCell('B1').value = 'Content';
    ws.getCell('C1').value = 'Translation';
    
    // Mix of different content types
    ws.getCell('A2').value = 'plain';
    ws.getCell('B2').value = 'Simple text';
    
    ws.getCell('A3').value = 'html';
    ws.getCell('B3').value = '<div>HTML <b>content</b></div>';
    
    ws.getCell('A4').value = 'formula';
    ws.getCell('B4').value = { formula: 'UPPER("hello")', result: 'HELLO' };
    
    ws.getCell('A5').value = 'richtext';
    ws.getCell('B5').value = {
      richText: [
        { text: 'Rich ', font: { bold: true } },
        { text: 'text', font: { italic: true } }
      ]
    };
    
    ws.getCell('A6').value = 'number';
    ws.getCell('B6').value = 12345;
    
    ws.getCell('A7').value = 'date';
    ws.getCell('B7').value = new Date('2024-01-01');

    const xlsxPath = path.join(tmpDir, 'mixed_types.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [{
          namePattern: 'Mixed',
          sourceColumns: ['B'],
          targetColumns: { en: 'C' },
          html: { enabled: true },
          extractFormulaResults: true,
          headerRow: 1,
          valuesStartRow: 2
        }]
      },
      global: { srcLang: 'en', xliffVersion: '2.1' }
    };

    const units = await extract(xlsxPath, config);
    
    // Should extract all content types
    expect(units.length).toBeGreaterThan(0);
    
    // Plain text
    const plainUnit = units.find(u => u.source === 'Simple text');
    expect(plainUnit).toBeDefined();
    
    // HTML
    const htmlUnit = units.find(u => u.source.includes('HTML'));
    expect(htmlUnit?.meta?.htmlSkeleton).toBeDefined();
    
    // Formula result
    const formulaUnit = units.find(u => u.source === 'HELLO');
    expect(formulaUnit?.formula).toBeDefined();
    
    // Rich text
    const richUnit = units.find(u => u.source === 'Rich text');
    expect(richUnit?.richText).toBe(true);
    
    // Number (converted to string)
    const numberUnit = units.find(u => u.source === '12345');
    expect(numberUnit).toBeDefined();
    
    // Date (converted to string)
    const dateUnit = units.find(u => u.source.includes('2024'));
    expect(dateUnit).toBeDefined();

    const xlf = await exportUnitsToXliff(units, config);
    const parsed = parseTranslated(xlf, 'xlf');
    
    for (const tu of parsed) {
      if (tu.segments) {
        tu.segments = tu.segments.map(s => ({ ...s, target: `[${s.source}]` }));
      }
    }

    const outPath = path.join(tmpDir, 'mixed_types_merged.xlsx');
    await merge(xlsxPath, outPath, parsed, config);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outPath);
    const ws2 = wb2.getWorksheet('Mixed')!;
    
    // Verify different types are handled
    expect(String(ws2.getCell('C2').value)).toBe('[Simple text]');
    expect(String(ws2.getCell('C3').value)).toContain('<div>');
    expect(String(ws2.getCell('C4').value)).toBe('[HELLO]');
  });

  it('should handle placeholder patterns correctly', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Placeholders');
    ws.getCell('A1').value = 'Source';
    ws.getCell('B1').value = 'Target';
    
    // Various placeholder patterns
    ws.getCell('A2').value = 'Hello {0}! Welcome to {1}.';
    ws.getCell('A3').value = 'You have %d new messages.';
    ws.getCell('A4').value = 'Price: %s USD';
    ws.getCell('A5').value = 'User {{name}} logged in at {{time}}.';
    ws.getCell('A6').value = 'Error: ${errorCode} - ${errorMessage}';

    const xlsxPath = path.join(tmpDir, 'placeholders.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [{
          namePattern: 'Placeholders',
          sourceColumns: ['A'],
          targetColumns: { en: 'B' },
          headerRow: 1,
          valuesStartRow: 2,
          inlineCodeRegexes: [
            '\\{\\d+\\}',           // {0}, {1}, etc.
            '%[sd]',                 // %s, %d
            '\\{\\{\\w+\\}\\}',     // {{name}}, {{time}}
            '\\$\\{\\w+\\}'         // ${errorCode}
          ]
        }]
      },
      global: { srcLang: 'en', xliffVersion: '2.1' }
    };

    const units = await extract(xlsxPath, config);
    expect(units.length).toBe(5);

    const xlf = await exportUnitsToXliff(units, config);
    
    // Placeholders should be marked as <ph> elements
    expect(xlf).toContain('<ph');
    
    const parsed = parseTranslated(xlf, 'xlf');
    
    // Simulate translation while keeping placeholders
    for (const tu of parsed) {
      if (tu.segments) {
        tu.segments = tu.segments.map(s => ({ 
          ...s, 
          target: s.source.replace('Hello', 'Bonjour').replace('Welcome', 'Bienvenue')
        }));
      }
    }

    const outPath = path.join(tmpDir, 'placeholders_merged.xlsx');
    await merge(xlsxPath, outPath, parsed, config);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outPath);
    const ws2 = wb2.getWorksheet('Placeholders')!;
    
    // Verify placeholders are preserved
    const b2 = String(ws2.getCell('B2').value);
    expect(b2).toContain('{0}');
    expect(b2).toContain('{1}');
    expect(b2).toContain('Bonjour');
  });

  it('should handle concurrent sheet processing', async () => {
    const wb = new ExcelJS.Workbook();
    
    // Create 5 sheets with different content
    for (let i = 1; i <= 5; i++) {
      const ws = wb.addWorksheet(`Sheet${i}`);
      ws.getCell('A1').value = 'Source';
      ws.getCell('B1').value = 'Target';
      
      for (let j = 2; j <= 50; j++) {
        ws.getCell(`A${j}`).value = `Sheet ${i} - String ${j - 1}`;
      }
    }

    const xlsxPath = path.join(tmpDir, 'concurrent_sheets.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [
          { namePattern: 'Sheet1', sourceColumns: ['A'], targetColumns: { en: 'B' }, headerRow: 1, valuesStartRow: 2 },
          { namePattern: 'Sheet2', sourceColumns: ['A'], targetColumns: { en: 'B' }, headerRow: 1, valuesStartRow: 2 },
          { namePattern: 'Sheet3', sourceColumns: ['A'], targetColumns: { en: 'B' }, headerRow: 1, valuesStartRow: 2 },
          { namePattern: 'Sheet4', sourceColumns: ['A'], targetColumns: { en: 'B' }, headerRow: 1, valuesStartRow: 2 },
          { namePattern: 'Sheet5', sourceColumns: ['A'], targetColumns: { en: 'B' }, headerRow: 1, valuesStartRow: 2 }
        ]
      },
      global: { srcLang: 'en', xliffVersion: '2.1' }
    };

    const units = await extract(xlsxPath, config);
    
    // Should extract from all sheets
    expect(units.length).toBe(5 * 49); // 5 sheets * 49 rows each
    
    // Verify each sheet is represented
    for (let i = 1; i <= 5; i++) {
      const sheetUnits = units.filter(u => u.sheetName === `Sheet${i}`);
      expect(sheetUnits.length).toBe(49);
    }

    const xlf = await exportUnitsToXliff(units, config);
    const parsed = parseTranslated(xlf, 'xlf');
    
    for (const tu of parsed) {
      if (tu.segments) {
        tu.segments = tu.segments.map(s => ({ ...s, target: `[${s.source}]` }));
      }
    }

    const outPath = path.join(tmpDir, 'concurrent_sheets_merged.xlsx');
    await merge(xlsxPath, outPath, parsed, config);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outPath);
    
    // Verify all sheets were processed correctly
    for (let i = 1; i <= 5; i++) {
      const ws = wb2.getWorksheet(`Sheet${i}`)!;
      expect(String(ws.getCell('B2').value)).toContain(`[Sheet ${i}`);
    }
  });
});
