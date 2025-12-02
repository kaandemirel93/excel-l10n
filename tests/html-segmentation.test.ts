import path from 'node:path';
import fs from 'node:fs';
import ExcelJS from 'exceljs';
import { extract, exportUnitsToXliff, parseTranslated, merge, parseConfig } from '../src/index';
import type { Config } from '../src/types';

describe('HTML with inline tags and segmentation', () => {
  const tmpDir = path.join(process.cwd(), '.out');
  beforeAll(() => { if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true }); });

  it('should handle complex HTML with inline tags, block tags, and segmentation', async () => {
    // Create a workbook with complex HTML content
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 'Source';
    ws.getCell('B1').value = 'Target';
    
    // Complex HTML with:
    // - Block-level tags (div, h1, p)
    // - Inline tags (b, i, a, span)
    // - Multiple sentences (for segmentation)
    // - Attributes on tags
    const complexHtml = '<div class="container"><h1 id="title">Welcome to <b>Our Store</b>!</h1><p>We offer <i>amazing</i> products. Visit our <a href="/shop">shop page</a> today. You won\'t regret it!</p><p>Special offer: <span class="highlight">50% off</span> everything.</p></div>';
    ws.getCell('A2').value = complexHtml;

    const xlsxPath = path.join(tmpDir, 'complex_html.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    // Config with segmentation enabled
    const config: Config = {
      workbook: {
        sheets: [{
          namePattern: 'Sheet1',
          sourceColumns: ['A'],
          targetColumns: { fr: 'B' },
          html: { enabled: true }
        }]
      },
      global: {
        srcLang: 'en',
        xliffVersion: '2.1'
      },
      segmentation: {
        enabled: true,
        rules: {
          srxPath: path.resolve(__dirname, '../examples/default_rules.srx')
        }
      }
    };

    // Extract
    const units = await extract(xlsxPath, config);
    expect(units.length).toBe(1);
    
    const unit = units[0];
    
    // Should have skeleton for block-level tags
    expect(unit.meta?.htmlSkeleton).toBeDefined();
    expect(unit.meta?.htmlSkeleton).toContain('<div class="container">');
    expect(unit.meta?.htmlSkeleton).toContain('<h1 id="title">');
    expect(unit.meta?.htmlSkeleton).toContain('<p>');
    expect(unit.meta?.htmlSkeleton).toContain('[[CONTENT]]');
    
    // Should have multiple segments (due to sentence segmentation)
    expect(unit.segments).toBeDefined();
    expect(unit.segments!.length).toBeGreaterThan(1);
    
    // Segments should contain XLIFF inline elements for inline tags
    const allSegmentsText = unit.segments!.map(s => s.source).join(' ');
    expect(allSegmentsText).toContain('<pc'); // XLIFF 2.1 inline element
    expect(allSegmentsText).toContain('dataRef="html_b"'); // Bold tag
    expect(allSegmentsText).toContain('dataRef="html_i"'); // Italic tag
    expect(allSegmentsText).toContain('dataRef="html_a"'); // Link tag
    expect(allSegmentsText).toContain('dataRef="html_span"'); // Span tag
    
    // Should NOT contain block-level tags in segments (they're in skeleton)
    expect(allSegmentsText).not.toContain('<div');
    expect(allSegmentsText).not.toContain('<h1');
    expect(allSegmentsText).not.toContain('<p>');

    // Export to XLIFF
    const xlf = await exportUnitsToXliff(units, config);
    
    // XLIFF should have multiple segments
    const segmentMatches = xlf.match(/<segment/g);
    expect(segmentMatches).toBeDefined();
    expect(segmentMatches!.length).toBeGreaterThan(1);
    
    // XLIFF should have skeleton in notes
    expect(xlf).toContain('category="htmlSkeleton"');
    
    // XLIFF should have inline elements
    expect(xlf).toContain('<pc id=');
    expect(xlf).toContain('dataRef="html_b"');

    // Parse and simulate translation
    const parsed = parseTranslated(xlf, 'xlf');
    expect(parsed.length).toBe(1);
    expect(parsed[0].segments).toBeDefined();
    expect(parsed[0].segments!.length).toBeGreaterThan(1);
    
    // Verify parsed segments have HTML tags (converted from XLIFF inline elements)
    const parsedSegmentsText = parsed[0].segments!.map(s => s.source).join(' ');
    expect(parsedSegmentsText).toContain('<b>'); // Converted from <pc dataRef="html_b">
    expect(parsedSegmentsText).toContain('<i>');
    expect(parsedSegmentsText).toContain('<a');
    expect(parsedSegmentsText).toContain('<span');
    
    // Add translations (wrap in brackets to identify translated content)
    for (const tu of parsed) {
      if (tu.segments) {
        tu.segments = tu.segments.map(s => ({ 
          ...s, 
          target: `[${s.source}]` 
        }));
      }
    }

    // Merge back
    const outPath = path.join(tmpDir, 'complex_html_merged.xlsx');
    await merge(xlsxPath, outPath, parsed, config);

    // Read result
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outPath);
    const ws2 = wb2.getWorksheet('Sheet1')!;
    const b2 = String(ws2.getCell('B2').value);

    // Verify the result has:
    // 1. Block-level tags preserved
    expect(b2).toContain('<div class="container">');
    expect(b2).toContain('<h1 id="title">');
    expect(b2).toContain('</h1>');
    expect(b2).toContain('<p>');
    expect(b2).toContain('</p>');
    expect(b2).toContain('</div>');
    
    // 2. Inline tags preserved with correct structure
    expect(b2).toContain('<b>Our Store</b>');
    expect(b2).toContain('<i>amazing</i>');
    expect(b2).toContain('<a href="/shop">shop page</a>');
    expect(b2).toContain('<span class="highlight">50% off</span>');
    
    // 3. Translated content (wrapped in brackets)
    expect(b2).toContain('[Welcome to');
    expect(b2).toContain('[We offer');
    expect(b2).toContain('[Special offer:');
    
    // 4. No corrupted tags or attributes in text
    expect(b2).not.toContain('html_b');
    expect(b2).not.toContain('html_i');
    expect(b2).not.toContain('dataRef');
    expect(b2).not.toMatch(/<b><\/b>\s*$/); // No empty tags at end
    
    // 5. Proper nesting maintained
    const divStart = b2.indexOf('<div');
    const divEnd = b2.lastIndexOf('</div>');
    expect(divStart).toBeGreaterThan(-1);
    expect(divEnd).toBeGreaterThan(divStart);
    
    // The content between div tags should contain all the inner HTML
    const divContent = b2.substring(divStart, divEnd + 6);
    expect(divContent).toContain('<h1');
    expect(divContent).toContain('<p>');
    expect(divContent).toContain('<b>');
    expect(divContent).toContain('<i>');
    expect(divContent).toContain('<a');
    expect(divContent).toContain('<span');
  });

  it('should handle nested inline tags correctly', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 'Source';
    ws.getCell('B1').value = 'Target';
    
    // Nested inline tags
    const nestedHtml = '<p>This is <b>bold and <i>italic</i> text</b> here.</p>';
    ws.getCell('A2').value = nestedHtml;

    const xlsxPath = path.join(tmpDir, 'nested_inline.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [{
          namePattern: 'Sheet1',
          sourceColumns: ['A'],
          targetColumns: { fr: 'B' }
        }]
      },
      global: {
        srcLang: 'en',
        xliffVersion: '2.1'
      }
    };

    const units = await extract(xlsxPath, config);
    const xlf = await exportUnitsToXliff(units, config);
    const parsed = parseTranslated(xlf, 'xlf');
    
    // Add translation
    for (const tu of parsed) {
      if (tu.segments) {
        tu.segments = tu.segments.map(s => ({ ...s, target: `[${s.source}]` }));
      }
    }

    const outPath = path.join(tmpDir, 'nested_inline_merged.xlsx');
    await merge(xlsxPath, outPath, parsed, config);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outPath);
    const ws2 = wb2.getWorksheet('Sheet1')!;
    const b2 = String(ws2.getCell('B2').value);

    // Verify nested structure is preserved
    expect(b2).toContain('<p>');
    expect(b2).toContain('<b>bold and <i>italic</i> text</b>');
    expect(b2).toContain('</p>');
    expect(b2).toContain('[This is');
  });

  it('should work with XLIFF 1.2 as well', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 'Source';
    ws.getCell('B1').value = 'Target';
    
    const html = '<div>Visit <a href="/shop" class="link">our shop</a> today. We have <b>great</b> deals!</div>';
    ws.getCell('A2').value = html;

    const xlsxPath = path.join(tmpDir, 'xliff12_test.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [{
          namePattern: 'Sheet1',
          sourceColumns: ['A'],
          targetColumns: { fr: 'B' }
        }]
      },
      global: {
        srcLang: 'en',
        xliffVersion: '1.2' // Test XLIFF 1.2
      },
      segmentation: {
        enabled: true,
        rules: {
          srxPath: path.resolve(__dirname, '../examples/default_rules.srx')
        }
      }
    };

    const units = await extract(xlsxPath, config);
    expect(units[0].meta?.htmlSkeleton).toBeDefined();
    expect(units[0].meta?.htmlInlineMap).toBeDefined(); // XLIFF 1.2 uses inlineMap
    
    const xlf = await exportUnitsToXliff(units, config);
    expect(xlf).toContain('version="1.2"');
    expect(xlf).toContain('<g id='); // XLIFF 1.2 uses <g> elements
    expect(xlf).toContain('category="htmlInlineMap"'); // InlineMap exported for 1.2
    
    const parsed = parseTranslated(xlf, 'xlf');
    expect(parsed[0].meta?.htmlInlineMap).toBeDefined();
    
    for (const tu of parsed) {
      if (tu.segments) {
        tu.segments = tu.segments.map(s => ({ ...s, target: `[${s.source}]` }));
      }
    }

    const outPath = path.join(tmpDir, 'xliff12_merged.xlsx');
    await merge(xlsxPath, outPath, parsed, config);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outPath);
    const ws2 = wb2.getWorksheet('Sheet1')!;
    const b2 = String(ws2.getCell('B2').value);

    // Verify attributes are preserved even with XLIFF 1.2
    expect(b2).toContain('<div>');
    expect(b2).toContain('<a href="/shop" class="link">our shop</a>');
    expect(b2).toContain('<b>great</b>');
    expect(b2).toContain('[Visit');
  });

  it('should handle multiple block-level tags at same level', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    ws.getCell('A1').value = 'Source';
    ws.getCell('B1').value = 'Target';
    
    // Multiple paragraphs
    const multiBlockHtml = '<div><p>First paragraph.</p><p>Second paragraph.</p><p>Third paragraph.</p></div>';
    ws.getCell('A2').value = multiBlockHtml;

    const xlsxPath = path.join(tmpDir, 'multi_block.xlsx');
    await wb.xlsx.writeFile(xlsxPath);

    const config: Config = {
      workbook: {
        sheets: [{
          namePattern: 'Sheet1',
          sourceColumns: ['A'],
          targetColumns: { fr: 'B' }
        }]
      },
      global: {
        srcLang: 'en',
        xliffVersion: '2.1'
      },
      segmentation: {
        enabled: true,
        rules: {
          srxPath: path.resolve(__dirname, '../examples/default_rules.srx')
        }
      }
    };

    const units = await extract(xlsxPath, config);
    
    // Should have skeleton with all paragraphs
    expect(units[0].meta?.htmlSkeleton).toContain('<div>');
    expect(units[0].meta?.htmlSkeleton).toContain('<p>');
    
    const xlf = await exportUnitsToXliff(units, config);
    const parsed = parseTranslated(xlf, 'xlf');
    
    for (const tu of parsed) {
      if (tu.segments) {
        tu.segments = tu.segments.map(s => ({ ...s, target: `[${s.source}]` }));
      }
    }

    const outPath = path.join(tmpDir, 'multi_block_merged.xlsx');
    await merge(xlsxPath, outPath, parsed, config);

    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(outPath);
    const ws2 = wb2.getWorksheet('Sheet1')!;
    const b2 = String(ws2.getCell('B2').value);

    // All paragraphs should be preserved
    const pMatches = b2.match(/<p>/g);
    expect(pMatches).toBeDefined();
    expect(pMatches!.length).toBe(3);
    
    expect(b2).toContain('[First paragraph.]');
    expect(b2).toContain('[Second paragraph.]');
    expect(b2).toContain('[Third paragraph.]');
  });
});
