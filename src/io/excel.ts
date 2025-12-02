import ExcelJS from 'exceljs';
import { TranslationUnit, CellStyleSnapshot, Config } from '../types.js';
import { colLetterToIndex, colIndexToLetter, makeTuId } from '../utils/index.js';

function looksHtml(s: string): boolean {
  if (!s) return false;
  // Heuristic: has a tag-like sequence
  return /<\/?[a-z][\s\S]*?>/i.test(s);
}

// Build an HTML skeleton with placeholders for text nodes and inline tags
export async function htmlToSkeleton(html: string, opts?: { translatableTags?: string[] }): Promise<{
  skeleton: string;
  texts: string[];
  inlineMap: Record<number, { open: string; close: string }>
}> {
  const { parse } = await import('node-html-parser');
  const root = parse(String(html), { lowerCaseTagName: false });
  const texts: string[] = [];
  const inlineMap: Record<number, { open: string; close: string }> = {};
  const defaultInline = ['b', 'strong', 'i', 'em', 'u', 'span', 'small', 'sub', 'sup', 'mark', 'a', 'code'];
  const configured = (opts?.translatableTags || []).map(t => String(t).toLowerCase());
  const inlineTags = new Set<string>([...defaultInline, ...configured]);
  let textIdx = 1;
  let inlineIdx = 1;
  const inlineIndexByNode = new WeakMap<object, number>();

  // First pass: Extract text nodes and replace them with [[htmltxt:N]]
  function extractTextNodes(node: any) {
    if (!node) return;

    if (node.nodeType === 3) { // text node
      const content = String(node.rawText ?? '');
      // Ignore pure-whitespace nodes; otherwise preserve whitespace inside tokens
      if (content.trim() !== '') {
        const token = `[[htmltxt:${textIdx++}]]`;
        texts.push(content);
        node.rawText = token;
      }
      return;
    }

    // Process children
    if (node.childNodes) {
      for (const child of node.childNodes) {
        extractTextNodes(child);
      }
    }
  }

  // Build skeleton, emitting inline markers for configured inline elements, and preserving non-inline element tags
  function buildSkeleton(node: any): string {
    if (!node) return '';
    if (node.nodeType === 3) {
      return String(node.rawText ?? '');
    }
    if (node.nodeType === 1) {
      const tag = String(node.tagName || '').toLowerCase();
      if (inlineTags.has(tag)) {
        let n = inlineIndexByNode.get(node as object);
        if (!n) {
          n = inlineIdx++;
          const rawAttrs = node.rawAttrs ? ` ${node.rawAttrs}` : '';
          inlineIndexByNode.set(node as object, n);
          inlineMap[n] = { open: `<${tag}${rawAttrs}>`, close: `</${tag}>` };
        }
        let out = `[[io:${n}]]`;
        if (node.childNodes) {
          for (const child of node.childNodes) out += buildSkeleton(child);
        }
        out += `[[ic:${n}]]`;
        return out;
      }
      // Non-inline element: if tag is present, preserve original tag wrapper; otherwise just concatenate children
      const rawAttrs = node.rawAttrs ? ` ${node.rawAttrs}` : '';
      if (tag) {
        let out = `<${tag}${rawAttrs}>`;
        if (node.childNodes) {
          for (const child of node.childNodes) out += buildSkeleton(child);
        }
        out += `</${tag}>`;
        return out;
      }
    }
    let out = '';
    if (node.childNodes) {
      for (const child of node.childNodes) {
        out += buildSkeleton(child);
      }
    }
    return out;
  }

  // Process the HTML
  extractTextNodes(root);

  // Generate the skeleton: only markers and [[htmltxt:N]] tokens in logical order
  const skeleton = buildSkeleton(root);

  return { skeleton, texts, inlineMap };
}

// Hybrid approach: Convert HTML to XLIFF inline elements for inline tags, preserve block tags in skeleton
export async function htmlToXliffWithSkeleton(html: string, opts?: { translatableTags?: string[]; xliffVersion?: '1.2' | '2.1' }): Promise<{
  skeleton: string;
  xliffSource: string;
  inlineMap?: Record<number, { open: string; close: string }>; // Only for XLIFF 1.2
}> {
  const { parse } = await import('node-html-parser');
  const root = parse(String(html), { lowerCaseTagName: false });
  const defaultInline = ['b', 'strong', 'i', 'em', 'u', 'span', 'small', 'sub', 'sup', 'mark', 'a', 'code'];
  const configured = (opts?.translatableTags || []).map(t => String(t).toLowerCase());
  const inlineTags = new Set<string>([...defaultInline, ...configured]);
  const version = opts?.xliffVersion || '2.1';
  let inlineIdCounter = 1;
  const inlineMap: Record<number, { open: string; close: string }> = {}; // Only used for XLIFF 1.2

  // Map HTML tag names to XLIFF ctype values for XLIFF 1.2
  const ctypeMap: Record<string, string> = {
    'b': 'bold',
    'strong': 'bold',
    'i': 'italic',
    'em': 'italic',
    'u': 'underline',
    'a': 'link',
    'code': 'code',
    'span': 'x-span',
    'small': 'x-small',
    'sub': 'x-sub',
    'sup': 'x-sup',
    'mark': 'x-mark',
  };

  // Process node and generate both skeleton and XLIFF source
  function processNode(node: any, isInsideBlock: boolean = false): { skeleton: string; xliff: string } {
    if (!node) return { skeleton: '', xliff: '' };

    // Text node
    if (node.nodeType === 3) {
      const text = String(node.rawText ?? '');
      return { skeleton: '', xliff: text };
    }

    // Element node
    if (node.nodeType === 1) {
      const tag = String(node.tagName || '').toLowerCase();

      // Process inline tags as XLIFF inline elements
      if (inlineTags.has(tag)) {
        const id = inlineIdCounter++;
        const rawAttrs = node.rawAttrs ? ` ${node.rawAttrs}` : '';
        
        let xliffInner = '';
        if (node.childNodes) {
          for (const child of node.childNodes) {
            const result = processNode(child, isInsideBlock);
            xliffInner += result.xliff;
          }
        }
        
        let xliffOut = '';
        if (version === '2.1') {
          // XLIFF 2.1: use <pc> elements with equivStart/equivEnd to preserve original HTML
          const dataRef = tag;
          const equivStart = `&lt;${tag}${rawAttrs.replace(/"/g, '&quot;')}&gt;`;
          const equivEnd = `&lt;/${tag}&gt;`;
          xliffOut = `<pc id="${id}" dataRef="html_${dataRef}" equivStart="${equivStart}" equivEnd="${equivEnd}">${xliffInner}</pc>`;
        } else {
          // XLIFF 1.2: use <g> elements with ctype
          // Store original tags in inline map for 1.2
          inlineMap[id] = { 
            open: `<${tag}${rawAttrs}>`, 
            close: `</${tag}>` 
          };
          const ctype = ctypeMap[tag] || `x-${tag}`;
          xliffOut = `<g id="${id}" ctype="${ctype}">${xliffInner}</g>`;
        }

        // Inline tags don't appear in skeleton - they're part of the content
        return { skeleton: '', xliff: xliffOut };
      }

      // Block-level tags: preserve in skeleton
      const rawAttrs = node.rawAttrs ? ` ${node.rawAttrs}` : '';
      let skeletonOut = '';
      let xliffOut = '';
      
      if (node.childNodes) {
        for (const child of node.childNodes) {
          const result = processNode(child, true);
          skeletonOut += result.skeleton;
          xliffOut += result.xliff;
        }
      }
      
      if (tag) {
        // If there's a nested skeleton, preserve it; otherwise use [[CONTENT]] placeholder
        const innerSkeleton = skeletonOut || '[[CONTENT]]';
        return { skeleton: `<${tag}${rawAttrs}>${innerSkeleton}</${tag}>`, xliff: xliffOut };
      }

      // No tag (e.g., root wrapper node) - pass through children's skeleton
      return { skeleton: skeletonOut, xliff: xliffOut };
    }

    // Other node types (like root wrapper): recurse through children
    let skeletonOut = '';
    let xliffOut = '';
    if (node.childNodes) {
      for (const child of node.childNodes) {
        const result = processNode(child, isInsideBlock);
        skeletonOut += result.skeleton;
        xliffOut += result.xliff;
      }
    }

    return { skeleton: skeletonOut, xliff: xliffOut };
  }

  const result = processNode(root);
  // Only return inlineMap for XLIFF 1.2 (XLIFF 2.1 uses equivStart/equivEnd)
  return { 
    skeleton: result.skeleton, 
    xliffSource: result.xliff, 
    inlineMap: version === '1.2' && Object.keys(inlineMap).length > 0 ? inlineMap : undefined 
  };
}

// Convert HTML to XLIFF inline elements (XLIFF 2.1 <pc> or XLIFF 1.2 <g>)
export async function htmlToXliffInline(html: string, opts?: { translatableTags?: string[]; xliffVersion?: '1.2' | '2.1' }): Promise<string> {
  const { parse } = await import('node-html-parser');
  const root = parse(String(html), { lowerCaseTagName: false });
  const defaultInline = ['b', 'strong', 'i', 'em', 'u', 'span', 'small', 'sub', 'sup', 'mark', 'a', 'code'];
  const configured = (opts?.translatableTags || []).map(t => String(t).toLowerCase());
  const inlineTags = new Set<string>([...defaultInline, ...configured]);
  const version = opts?.xliffVersion || '2.1';
  let inlineIdCounter = 1;

  // Map HTML tag names to XLIFF ctype values for XLIFF 1.2
  const ctypeMap: Record<string, string> = {
    'b': 'bold',
    'strong': 'bold',
    'i': 'italic',
    'em': 'italic',
    'u': 'underline',
    'a': 'link',
    'code': 'code',
    'span': 'x-span',
    'small': 'x-small',
    'sub': 'x-sub',
    'sup': 'x-sup',
    'mark': 'x-mark',
  };

  function processNode(node: any): string {
    if (!node) return '';

    // Text node
    if (node.nodeType === 3) {
      const text = String(node.rawText ?? '');
      // Escape XML special characters
      return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    // Element node
    if (node.nodeType === 1) {
      const tag = String(node.tagName || '').toLowerCase();

      // Process inline tags as XLIFF inline elements
      if (inlineTags.has(tag)) {
        const id = inlineIdCounter++;
        let innerContent = '';
        if (node.childNodes) {
          for (const child of node.childNodes) {
            innerContent += processNode(child);
          }
        }

        if (version === '2.1') {
          // XLIFF 2.1: use <pc> elements with dataRef
          const dataRef = tag;
          return `<pc id="${id}" dataRef="html_${dataRef}">${innerContent}</pc>`;
        } else {
          // XLIFF 1.2: use <g> elements with ctype
          const ctype = ctypeMap[tag] || `x-${tag}`;
          return `<g id="${id}" ctype="${ctype}">${innerContent}</g>`;
        }
      }

      // Block-level tags: strip them but keep content
      let out = '';
      if (node.childNodes) {
        for (const child of node.childNodes) {
          out += processNode(child);
        }
      }
      return out;
    }

    // Other node types: recurse through children
    let out = '';
    if (node.childNodes) {
      for (const child of node.childNodes) {
        out += processNode(child);
      }
    }
    return out;
  }

  return processNode(root);
}

function htmlToText(html: string, opts?: { translatableTags?: string[] }): string {
  let s = String(html);
  // Remove script/style contents entirely
  s = s.replace(/<script[\s\S]*?>[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?>[\s\S]*?<\/style>/gi, '');
  // Convert common line-break tags to newlines
  s = s.replace(/<\s*br\s*\/?\s*>/gi, '\n');
  // Insert newlines after block-level closings to preserve structure
  const blockTags = ['p', 'div', 'li', 'ul', 'ol', 'section', 'article', 'header', 'footer', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'title'];
  for (const t of blockTags) {
    const re = new RegExp(`<\\/${t}\\s*>`, 'gi');
    s = s.replace(re, '\n');
  }
  // Strip remaining tags but keep inner text
  s = s.replace(/<[^>]+>/g, '');
  // Collapse excessive spaces while preserving newlines
  s = s.replace(/[\t\r]+/g, '')
    .replace(/\u00A0/g, ' ')
    .replace(/ +/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  return s;
}

export async function readWorkbook(filePath: string): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  return wb;
}

function takeStyleSnapshot(cell: ExcelJS.Cell): CellStyleSnapshot | undefined {
  const font = cell.font ? {
    name: cell.font.name,
    size: cell.font.size as number | undefined,
    bold: cell.font.bold as boolean | undefined,
    italic: cell.font.italic as boolean | undefined,
    color: (cell.font.color && (cell.font.color as any).argb) ? `#${(cell.font.color as any).argb.slice(2)}` : undefined,
  } : undefined;
  const alignment = cell.alignment ? {
    horizontal: cell.alignment.horizontal,
    vertical: cell.alignment.vertical,
    wrapText: cell.alignment.wrapText as boolean | undefined,
  } : undefined;
  const fill = cell.fill && (cell.fill as any).fgColor?.argb ? { type: (cell.fill as any).type, color: `#${(cell.fill as any).fgColor.argb.slice(2)}` } : undefined;
  if (!font && !alignment && !fill) return undefined;
  return { font, alignment, fill };
}

function normalizeSheetMatch(namePattern: string, sheetName: string): boolean {
  try {
    if (namePattern.startsWith('^') || namePattern.endsWith('$') || namePattern.includes('[') || namePattern.includes('(')) {
      return new RegExp(namePattern).test(sheetName);
    }
  } catch {
    // fallback to exact
  }
  return namePattern === sheetName;
}

export async function extractUnits(inputXlsxPath: string, config: Config): Promise<TranslationUnit[]> {
  const wb = await readWorkbook(inputXlsxPath);
  const units: TranslationUnit[] = [];

  for (const sheetCfg of config.workbook.sheets) {
    const sheets = wb.worksheets.filter(ws => normalizeSheetMatch(sheetCfg.namePattern, ws.name));
    for (const ws of sheets) {
      const startRow = sheetCfg.valuesStartRow ?? 2;
      const lastRow = ws.rowCount;
      const excludedRows = new Set(sheetCfg.excludedRows || []);
      const excludedCols = new Set((sheetCfg.excludedColumns || []).map(s => s.toUpperCase()));

      for (let r = startRow; r <= lastRow; r++) {
        const row = ws.getRow(r);
        if ((sheetCfg.skipHiddenRows && (row.hidden ?? false)) || excludedRows.has(r)) continue;

        for (const srcCol of sheetCfg.sourceColumns) {
          const colLetter = srcCol.toUpperCase();
          if (excludedCols.has(colLetter)) continue;
          const cidx = colLetterToIndex(colLetter);
          const col = ws.getColumn(cidx);
          if (sheetCfg.skipHiddenColumns && (col.hidden ?? false)) continue;

          const cell = ws.getCell(r, cidx);
          // merged handling: only top-left when configured
          if (sheetCfg.treatMergedRegions !== 'expand' && cell.isMerged) {
            const m = cell.master as any; // top-left cell
            if (sheetCfg.treatMergedRegions === 'skip') continue;
            const mRow = Number(m.row);
            const mCol = Number(m.col);
            if (sheetCfg.treatMergedRegions === 'top-left' && (mRow !== r || mCol !== cidx)) continue;
          }

          // color exclusion (font or fill)
          const exColors = (sheetCfg.excludeColors || []).map(s => s.toLowerCase());
          const fontColor = (cell.font as any)?.color?.argb ? `#${(cell.font as any).color.argb.slice(2)}`.toLowerCase() : undefined;
          const fillColor = (cell.fill as any)?.fgColor?.argb ? `#${(cell.fill as any).fgColor.argb.slice(2)}`.toLowerCase() : undefined;
          if ((fontColor && exColors.includes(fontColor)) || (fillColor && exColors.includes(fillColor))) continue;

          let text = '';
          let richText = false;
          let formula: string | null = null;
          const v: any = cell.value;
          if (v && typeof v === 'object' && 'richText' in v) {
            richText = true;
            text = (v.richText as any[]).map(rt => rt.text).join('');
          } else if (v && typeof v === 'object' && 'formula' in v) {
            formula = v.formula;
            if (sheetCfg.extractFormulaResults !== false) {
              text = typeof v.result === 'string' ? v.result : String(v.result ?? '');
            } else {
              text = String(v.formula ?? '');
            }
          } else if (v == null) {
            text = '';
          } else {
            text = typeof v === 'string' ? v : String(v);
          }

          // Optional HTML sub-filter
          const meta: Record<string, any> = {};
          let htmlDetected = false;

          if ((sheetCfg as any).html?.enabled !== false && looksHtml(text)) {
            htmlDetected = true;
            const originalHtml = text;
            const transTags = (sheetCfg as any).html?.translatableTags as string[] | undefined;
            const xliffVersion = config.global?.xliffVersion || '2.1';

            // Use hybrid approach: XLIFF inline elements for inline tags, skeleton for block tags
            const { skeleton, xliffSource, inlineMap } = await htmlToXliffWithSkeleton(originalHtml, {
              translatableTags: transTags,
              xliffVersion
            });

            // Store skeleton for reconstruction during merge
            meta.htmlSkeleton = skeleton;
            // Only store inlineMap for XLIFF 1.2 (2.1 uses equivStart/equivEnd in the XLIFF itself)
            if (inlineMap) {
              meta.htmlInlineMap = inlineMap;
            }
            meta.htmlOriginal = originalHtml;

            // The translatable text is the XLIFF source with inline elements
            text = xliffSource;
          }

          if (!htmlDetected && text === '') continue; // skip empty when not HTML; when HTML with zero texts, keep TU with skeleton

          const style = sheetCfg.preserveStyles ? takeStyleSnapshot(cell) : undefined;
          const id = makeTuId(ws.name, r, colLetter);
          // header text from configured headerRow (by column)
          if (sheetCfg.headerRow && sheetCfg.headerRow >= 1) {
            const hcell = ws.getCell(sheetCfg.headerRow, cidx);
            const hv: any = hcell.value;
            let headerText = '';
            if (hv && typeof hv === 'object' && 'richText' in hv) {
              headerText = (hv.richText as any[]).map(rt => rt.text).join('');
            } else if (hv && typeof hv === 'object' && 'formula' in hv) {
              headerText = typeof hv.result === 'string' ? hv.result : String(hv.result ?? '');
            } else if (hv != null) {
              headerText = typeof hv === 'string' ? hv : String(hv);
            }
            if (headerText) meta.headerName = headerText;
          }
          // metadataRows values for this column
          if (sheetCfg.metadataRows && sheetCfg.metadataRows.length) {
            const map: Record<number, string> = {};
            for (const mr of sheetCfg.metadataRows) {
              const mcell = ws.getCell(mr, cidx);
              const mv: any = mcell.value;
              let mtext = '';
              if (mv && typeof mv === 'object' && 'richText' in mv) {
                mtext = (mv.richText as any[]).map(rt => rt.text).join('');
              } else if (mv && typeof mv === 'object' && 'formula' in mv) {
                mtext = typeof mv.result === 'string' ? mv.result : String(mv.result ?? '');
              } else if (mv != null) {
                mtext = typeof mv === 'string' ? mv : String(mv);
              }
              if (mtext) map[mr] = mtext;
            }
            if (Object.keys(map).length) meta.metadataRows = map;
          }
          if (sheetCfg.translateComments) {
            const note = (cell as any).note || (cell as any).comments;
            if (note) meta.comments = note;
          }
          const tu: TranslationUnit = {
            id,
            sheetName: ws.name,
            row: r,
            col: colLetter,
            colIndex: cidx,
            source: text,
            richText,
            style,
            formula,
            isMerged: cell.isMerged,
            mergedRange: (cell as any)._mergeCount ? (cell as any)._mergeCount : null,
            meta: Object.keys(meta).length ? meta : undefined,
            segments: [{
              id: `${id}_s0`,
              // Segment source equals translator-facing text (HTML stripped, inline markers preserved)
              source: text,
              target: ''
            }]
          };

          // Do not override segments for HTML; keep plain concatenated text for translators
          units.push(tu);
        }
      }
    }
  }
  return units;
}
