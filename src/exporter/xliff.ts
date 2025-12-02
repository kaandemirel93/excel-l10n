import { create } from 'xmlbuilder2';
import { XMLParser } from 'fast-xml-parser';
import { Config, Segment, TranslationUnit } from '../types.js';

function regexesForUnit(u: TranslationUnit, config: Config): RegExp[] {
  const sheetCfg = config.workbook.sheets.find(s => {
    try { return new RegExp(s.namePattern).test(u.sheetName) || s.namePattern === u.sheetName; } catch { return s.namePattern === u.sheetName; }
  });
  const res = (sheetCfg?.inlineCodeRegexes || []).map(r => {
    try { return new RegExp(r, 'g'); } catch { return null; }
  }).filter((x): x is RegExp => !!x);
  return res;
}

function encodePlaceholders(text: string, regs: RegExp[]): { encoded: string; map: Record<string, string> } {
  const map: Record<string, string> = {};
  let idx = 1;
  // First, protect ICU inner texts by converting them to <pc> markers
  // Very simplified detector for {var, plural|select, ... {text}}
  const ICU_BLOCK = /\{[^{}]*,\s*(plural|select)\s*,[\s\S]*?\}/g;
  let icuIdx = 1;
  let encoded = text.replace(ICU_BLOCK, (m) => {
    return m.replace(/\{([^{}]*)\}/g, (_inner, grp) => {
      const id = `icu${icuIdx++}`;
      // encode as [[pc:id:text]]; writer will emit <pc id>text</pc>
      return `[[pc:${id}:${String(grp).replace(/\]/g, '')}]]`;
    });
  });
  let phIndex = 1;
  for (const re of regs) {
    encoded = encoded.replace(re, (m) => {
      const id = `ph${phIndex++}`;
      map[id] = m;
      return `[[ph:${id}]]`;
    });
  }
  return { encoded, map };
}

function writeWithPh(parent: any, encoded: string) {
  const MARK = /(\[\[(ph|pc):[^\]]+\]\])/g;
  const parts = encoded.split('\u0000').join('');

  // Check if the text already contains XLIFF inline elements (<pc>, <g>, etc.)
  if (/<(pc|g|sc|ec|bx|ex|bpt|ept)\s/.test(parts)) {
    // Parse inline XLIFF elements properly, handling nesting
    let i = 0;
    while (i < parts.length) {
      // Look for next opening tag
      const tagSearch = parts.substring(i).match(/<(pc|g|sc|ec|bx|ex|bpt|ept)(\s[^>]*)?>/);
      
      if (!tagSearch) {
        // No more tags, write remaining text
        const remaining = parts.substring(i);
        if (remaining) parent.txt(remaining);
        break;
      }
      
      const matchStart = i + (tagSearch.index || 0);
      
      // Write any text before this tag
      if (matchStart > i) {
        parent.txt(parts.substring(i, matchStart));
      }
      
      const tagName = tagSearch[1];
      const attrsStr = tagSearch[2] || '';
      const selfClosing = attrsStr.trim().endsWith('/');
      
      i = matchStart + tagSearch[0].length;
      
      if (selfClosing) {
        // Self-closing tag
        const attrObj = parseAttributes(attrsStr);
        parent.ele(tagName, attrObj);
        continue;
      }
      
      // Find matching closing tag, handling nesting
      const closeTag = `</${tagName}>`;
      let depth = 1;
      let contentStart = i;
      let j = i;
      
      while (j < parts.length && depth > 0) {
        const nextOpen = parts.indexOf(`<${tagName}`, j);
        const nextClose = parts.indexOf(closeTag, j);
        
        if (nextClose === -1) {
          // No closing tag found
          break;
        }
        
        if (nextOpen !== -1 && nextOpen < nextClose) {
          // Found nested opening tag
          depth++;
          j = nextOpen + tagName.length + 1;
        } else {
          // Found closing tag
          depth--;
          if (depth === 0) {
            // This is our closing tag
            const innerContent = parts.substring(contentStart, nextClose);
            const attrObj = parseAttributes(attrsStr);
            const elem = parent.ele(tagName, attrObj);
            // Recursively handle inner content
            writeWithPh(elem, innerContent);
            i = nextClose + closeTag.length;
            break;
          }
          j = nextClose + closeTag.length;
        }
      }
      
      if (depth > 0) {
        // Unclosed tag, treat as text
        parent.txt(parts.substring(matchStart));
        break;
      }
    }
    return;
  }

  // Otherwise, process placeholder markers
  for (const part of parts.split(MARK)) {
    if (!part) continue;
    const mph = /^\[\[ph:([^\]]+)\]\]$/.exec(part);
    const mpc = /^\[\[pc:([^:]+?):([^\]]*)\]\]$/.exec(part);
    if (mph) {
      parent.ele('ph', { id: mph[1] });
    } else if (mpc) {
      const pc = parent.ele('pc', { id: mpc[1] });
      pc.txt(mpc[2]);
    } else {
      parent.txt(part);
    }
  }
}

// Helper to parse XML attributes from a string
function parseAttributes(attrString: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  if (!attrString) return attrs;
  const attrPattern = /(\w+)="([^"]*)"/g;
  let match;
  while ((match = attrPattern.exec(attrString)) !== null) {
    attrs[match[1]] = match[2];
  }
  return attrs;
}

export async function exportToXliff(units: TranslationUnit[], config: Config, options?: { srcLang?: string; trgLang?: string; generator?: string }): Promise<string> {
  const srcLang = options?.srcLang ?? config.global?.srcLang ?? 'en';
  const xliffVersion = config.global?.xliffVersion || '2.1';
  const attrs: any = { version: xliffVersion, srcLang };
  if (options?.trgLang) attrs.trgLang = options.trgLang;
  const root = create({ version: '1.0', encoding: 'UTF-8' }).ele('xliff', attrs);
  const file = root.ele('file', { id: 'workbook', original: 'workbook.xlsx', 'tool-id': options?.generator ?? 'excel-l10n' });

  for (const u of units) {
    const regs = regexesForUnit(u, config);
    const unit = file.ele('unit', { id: u.id });
    const notes = unit.ele('notes');
    notes.ele('note').txt(`sheet=${u.sheetName};row=${u.row};col=${u.col}`);
    if (config.global?.exportComments) {
      if (u.meta?.headerName) notes.ele('note', { category: 'header' }).txt(String(u.meta.headerName));
      if (u.meta?.metadataRows) notes.ele('note', { category: 'metadataRows' }).txt(JSON.stringify(u.meta.metadataRows));
      if (u.meta?.comments) notes.ele('note', { category: 'comments' }).txt(typeof u.meta.comments === 'string' ? u.meta.comments : JSON.stringify(u.meta.comments));
    }
    // Export HTML skeleton and inline map for reconstruction during merge
    if (u.meta?.htmlSkeleton) {
      notes.ele('note', { category: 'htmlSkeleton' }).txt(String(u.meta.htmlSkeleton));
    }
    if (u.meta?.htmlInlineMap) {
      notes.ele('note', { category: 'htmlInlineMap' }).txt(JSON.stringify(u.meta.htmlInlineMap));
    }
    if (u.meta?.htmlTexts) {
      notes.ele('note', { category: 'htmlTexts' }).txt(JSON.stringify(u.meta.htmlTexts));
    }

    const segs = u.segments && u.segments.length ? u.segments : [{ id: `${u.id}_s0`, source: u.source } as Segment];
    // collect placeholder map for this unit
    const phMap: Record<string, Record<string, string>> = {};
    for (const s of segs) {
      const seg = unit.ele('segment', { id: s.id });
      const { encoded, map } = encodePlaceholders(s.source, regs);
      phMap[s.id] = map;
      const src = seg.ele('source').att('xml:space', 'preserve');
      writeWithPh(src, encoded);
      if (s.target != null) {
        const tgt = seg.ele('target').att('xml:space', 'preserve');
        writeWithPh(tgt, s.target);
      }
    }
    if (Object.keys(phMap).length) {
      notes.ele('note', { category: 'ph' }).txt(JSON.stringify(phMap));
    }
  }

  // Disable pretty printing to ensure whitespace around inline tags is preserved exactly as is.
  // Pretty printing often adds newlines/indentation in mixed content (e.g. <source>text <pc>...</pc></source>)
  // which can introduce unwanted spaces or break words.
  return root.end({ prettyPrint: false });
}

function extractNotes(notesArray: any): { sheetName: string; row: number; col: string; ph?: Record<string, Record<string, string>>; htmlSkeleton?: string; htmlInlineMap?: Record<string, { open: string; close: string }>; htmlTexts?: string[] } {
  let sheetName = '';
  let row = 0;
  let col = '';
  let ph: Record<string, Record<string, string>> | undefined;
  let htmlSkeleton: string | undefined;
  let htmlInlineMap: Record<string, { open: string; close: string }> | undefined;
  let htmlTexts: string[] | undefined;
  
  // With preserveOrder: true, notesArray is an array of objects
  if (!Array.isArray(notesArray)) return { sheetName, row, col, ph, htmlSkeleton, htmlInlineMap, htmlTexts };
  
  for (const item of notesArray) {
    if (item.note) {
      // item.note is an array
      const noteArray = Array.isArray(item.note) ? item.note : [item.note];
      
      // Get category from item's attributes (at same level as 'note')
      let category = '';
      if (item[':@']) {
        const attrs = item[':@'];
        category = attrs['@_category'] || attrs.category || '';
      }
      
      // Get text from note array
      let text = '';
      for (const n of noteArray) {
        if (n['#text']) {
          text = n['#text'];
          break;
        }
      }
      
      if (!category && text) {
        // Plain note without category - check for sheet/row/col pattern
        const m = /sheet=(.*?);row=(\d+);col=([A-Z]+)/.exec(text);
        if (m) { sheetName = m[1]; row = parseInt(m[2], 10); col = m[3]; }
      } else if (category === 'ph' && text) {
        try { ph = JSON.parse(text); } catch { /* ignore */ }
      } else if (category === 'htmlSkeleton' && text) {
        htmlSkeleton = text;
      } else if (category === 'htmlInlineMap' && text) {
        try { htmlInlineMap = JSON.parse(text); } catch { /* ignore */ }
      } else if (category === 'htmlTexts' && text) {
        try { htmlTexts = JSON.parse(text); } catch { /* ignore */ }
      }
    }
  }
  
  return { sheetName, row, col, ph, htmlSkeleton, htmlInlineMap, htmlTexts };
}

function flattenText(node: any): string {
  if (node == null) return '';
  if (typeof node === 'string') return node;
  if (typeof node === 'object') {
    if (node.ph) {
      const list = Array.isArray(node.ph) ? node.ph : [node.ph];
      // construct a string with markers in order of appearance by scanning through a pseudo order: text before/after not preserved by parser reliably
    }
    // fast-xml-parser gives either string or objects; simplest approach: JSON stringify and replace tags
    const s = JSON.stringify(node);
    return s
      .replace(/\{"ph":\{"id":"(.*?)"\}\}/g, '[[ph:$1]]')
      .replace(/\{"ph":\[\{\"id\":\"(.*?)\"\}\]\}/g, '[[ph:$1]]')
      .replace(/\{"#text":"(.*?)"\}/g, '$1')
      .replace(/\{"source":"(.*?)"\}/g, '$1')
      .replace(/\{"target":"(.*?)"\}/g, '$1');
  }
  return String(node);
}

export function parseXliffToUnits(xlf: string): TranslationUnit[] {
  // Use preserveOrder: true to maintain the order of mixed content (text and inline elements)
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '', preserveOrder: true, trimValues: false });
  const obj: any = parser.parse(xlf);
  const units: TranslationUnit[] = [];

  // With preserveOrder: true, the structure is an array of objects with single keys
  // Find the xliff element
  const xliffArray = obj;
  let xliff: any = null;
  for (const item of xliffArray) {
    if (item.xliff) {
      xliff = item.xliff;
      break;
    }
  }

  if (!xliff) return units;

  // Extract xliff attributes
  let xliffTrg = '';
  for (const item of xliffArray) {
    if (item.xliff && item[':@']) {
      const attrs = item[':@'];
      xliffTrg = attrs['@_trgLang'] || attrs.trgLang || '';
    }
  }

  // Find file elements
  const files: any[] = [];
  for (const item of xliff) {
    if (item.file) {
      files.push(item.file);
    }
  }

  // Flatten text with <ph id> and XLIFF inline elements recursively
  // With preserveOrder: true, nodes are arrays of objects
  const flatten = (nodeArray: any): string => {
    if (!nodeArray) return '';
    if (typeof nodeArray === 'string') return nodeArray;
    if (!Array.isArray(nodeArray)) {
      // Fallback for non-array (shouldn't happen with preserveOrder: true)
      if (typeof nodeArray === 'object' && nodeArray['#text']) {
        return String(nodeArray['#text']);
      }
      return '';
    }

    let out = '';
    for (const node of nodeArray) {
      // Text node
      if (node['#text'] !== undefined) {
        out += String(node['#text']);
        continue;
      }

      // Placeholder element
      if (node.ph) {
        const phArray = node.ph;
        const attrs = node[':@'] || {};
        const id = attrs['@_id'] || attrs.id;
        if (id) out += `[[ph:${id}]]`;
        continue;
      }

      // XLIFF 2.1 pc element
      if (node.pc) {
        const pcArray = node.pc;
        const attrs = node[':@'] || {};
        const dataRef = attrs['@_dataRef'] || attrs.dataRef || '';
        const equivStart = attrs['@_equivStart'] || attrs.equivStart;
        const equivEnd = attrs['@_equivEnd'] || attrs.equivEnd;
        
        const innerText = flatten(pcArray);
        
        // If equivStart/equivEnd are present, use them (they contain the original HTML with attributes)
        if (equivStart && equivEnd) {
          // Decode HTML entities
          const openTag = equivStart.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
          const closeTag = equivEnd.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
          out += `${openTag}${innerText}${closeTag}`;
        } else {
          // Fallback: use dataRef to determine tag name
          const tagName = dataRef.replace(/^html_/, '') || 'span';
          out += `<${tagName}>${innerText}</${tagName}>`;
        }
        continue;
      }

      // XLIFF 1.2 g element
      if (node.g) {
        const gArray = node.g;
        const attrs = node[':@'] || {};
        const ctype = attrs['@_ctype'] || attrs.ctype || '';
        let tagName = 'span';
        if (ctype === 'bold') tagName = 'b';
        else if (ctype === 'italic') tagName = 'i';
        else if (ctype === 'underline') tagName = 'u';
        else if (ctype === 'link') tagName = 'a';
        else if (ctype === 'code') tagName = 'code';
        else if (ctype.startsWith('x-')) tagName = ctype.substring(2);

        const innerText = flatten(gArray);
        out += `<${tagName}>${innerText}</${tagName}>`;
        continue;
      }

      // sc/ec elements
      if (node.sc || node.ec) {
        if (node.sc) out += flatten(node.sc);
        if (node.ec) out += flatten(node.ec);
        continue;
      }

      // bx/ex/bpt/ept elements
      if (node.bpt) {
        const bptArray = node.bpt;
        for (const bptNode of bptArray) {
          if (bptNode['#text']) {
            out += bptNode['#text'];
          }
        }
        continue;
      }
      if (node.ept) {
        const eptArray = node.ept;
        for (const eptNode of eptArray) {
          if (eptNode['#text']) {
            out += eptNode['#text'];
          }
        }
        continue;
      }

      // source/target (recurse)
      if (node.source) {
        out += flatten(node.source);
        continue;
      }
      if (node.target) {
        out += flatten(node.target);
        continue;
      }
    }
    return out;
  };

  for (const f of files) {
    // With preserveOrder: true, f is an array of objects
    const fileArray = Array.isArray(f) ? f : [f];
    let fileTrg = xliffTrg;
    
    // Extract file attributes
    for (const item of fileArray) {
      if (item[':@']) {
        const attrs = item[':@'];
        fileTrg = attrs['@_trgLang'] || attrs.trgLang || fileTrg;
      }
    }
    
    // Find unit elements - keep the parent item that has both 'unit' and ':@'
    const unitItems: any[] = [];
    for (const item of fileArray) {
      if (item.unit) {
        unitItems.push(item);
      }
    }
    
    for (const unitItem of unitItems) {
      // unitItem is the object that contains both 'unit' array and ':@' attributes
      if (!unitItem || typeof unitItem !== 'object') continue;
      
      // Extract unit attributes from the parent object
      let id = '';
      if (unitItem[':@']) {
        const attrs = unitItem[':@'];
        id = attrs['@_id'] || attrs.id || '';
      }
      
      if (!id) continue;
      
      // Get the actual unit array
      const unitArray = unitItem.unit;
      if (!unitArray || !Array.isArray(unitArray)) continue;
      
      // Extract notes
      let sheetName = '', col = ''; let row = 0;
      let phMap: Record<string, Record<string, string>> | undefined;
      let htmlSkeleton: string | undefined;
      let htmlInlineMap: Record<string, { open: string; close: string }> | undefined;
      let htmlTexts: string[] | undefined;
      
      for (const item of unitArray) {
        if (item.notes) {
          const { sheetName: sn, row: rr, col: cc, ph, htmlSkeleton: hs, htmlInlineMap: him, htmlTexts: htxt } = extractNotes(item.notes);
          if (sn) sheetName = sn;
          if (rr) row = rr;
          if (cc) col = cc;
          if (ph) phMap = ph;
          if (hs) htmlSkeleton = hs;
          if (him) htmlInlineMap = him;
          if (htxt) htmlTexts = htxt;
        }
      }

      // Extract segments
      const segmentElements: any[] = [];
      for (const item of unitArray) {
        if (item.segment) {
          segmentElements.push(item.segment);
        }
      }
      
      const segments: Segment[] = segmentElements.map((segArray: any, idx: number) => {
        let segId = `${id}_s${idx}`;
        let sourceArray: any = null;
        let targetArray: any = null;
        
        if (Array.isArray(segArray)) {
          for (const item of segArray) {
            if (item[':@']) {
              const attrs = item[':@'];
              segId = attrs['@_id'] || attrs.id || segId;
            }
            if (item.source) {
              sourceArray = item.source;
            }
            if (item.target) {
              targetArray = item.target;
            }
          }
        }
        
        const src = flatten(sourceArray);
        const tgt = targetArray ? flatten(targetArray) : undefined;
        return { id: segId, source: src, target: tgt };
      });

      const tu: TranslationUnit = { id, sheetName, row, col, colIndex: 0, source: segments.map(s => s.source).join(''), segments, meta: {} };
      if (fileTrg) (tu.meta as any).targetLocale = fileTrg;
      if (phMap) (tu.meta as any).placeholders = phMap;
      if (htmlSkeleton) (tu.meta as any).htmlSkeleton = htmlSkeleton;
      if (htmlInlineMap) (tu.meta as any).htmlInlineMap = htmlInlineMap;
      if (htmlTexts) (tu.meta as any).htmlTexts = htmlTexts;
      units.push(tu);
    }
  }
  return units;
}
