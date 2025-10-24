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

export async function exportToXliff(units: TranslationUnit[], config: Config, options?: { srcLang?: string; trgLang?: string; generator?: string }): Promise<string> {
  const srcLang = options?.srcLang ?? config.global?.srcLang ?? 'en';
  const attrs: any = { version: '2.1', srcLang };
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
    if (u.meta && (u.meta as any).htmlSkeleton) {
      notes.ele('note', { category: 'htmlSkeleton' }).txt(String((u.meta as any).htmlSkeleton));
      if ((u.meta as any).htmlInlineMap) {
        try {
          notes.ele('note', { category: 'htmlInlineMap' }).txt(JSON.stringify((u.meta as any).htmlInlineMap));
        } catch {}
      }
    }

    const segs = u.segments && u.segments.length ? u.segments : [{ id: `${u.id}_s0`, source: u.source } as Segment];
    // collect placeholder map for this unit
    const phMap: Record<string, Record<string, string>> = {};
    for (const s of segs) {
      const seg = unit.ele('segment', { id: s.id });
      const { encoded, map } = encodePlaceholders(s.source, regs);
      phMap[s.id] = map;
      const src = seg.ele('source');
      writeWithPh(src, encoded);
      if (s.target != null) {
        const tgt = seg.ele('target');
        writeWithPh(tgt, s.target);
      }
    }
    if (Object.keys(phMap).length) {
      notes.ele('note', { category: 'ph' }).txt(JSON.stringify(phMap));
    }
  }

  return root.end({ prettyPrint: true });
}

function extractNotes(u: any): { sheetName: string; row: number; col: string; ph?: Record<string, Record<string, string>>; htmlSkeleton?: string; htmlInlineMap?: Record<string, { open: string; close: string }>; htmlTexts?: string[] } {
  let sheetName = '';
  let row = 0;
  let col = '';
  let ph: Record<string, Record<string, string>> | undefined;
  let htmlSkeleton: string | undefined;
  let htmlInlineMap: Record<string, { open: string; close: string }> | undefined;
  let htmlTexts: string[] | undefined;
  const notes = u.notes?.note;
  const notesArr = Array.isArray(notes) ? notes : (notes ? [notes] : []);
  for (const n of notesArr) {
    if (typeof n === 'string') {
      const m = /sheet=(.*?);row=(\d+);col=([A-Z]+)/.exec(n);
      if (m) { sheetName = m[1]; row = parseInt(m[2], 10); col = m[3]; }
    } else if (typeof n === 'object' && n.category === 'ph' && typeof n['#text'] === 'string') {
      try { ph = JSON.parse(n['#text']); } catch { /* ignore */ }
    } else if (typeof n === 'object' && n.category === 'htmlSkeleton' && typeof n['#text'] === 'string') {
      htmlSkeleton = n['#text'];
    } else if (typeof n === 'object' && n.category === 'htmlInlineMap' && typeof n['#text'] === 'string') {
      try { htmlInlineMap = JSON.parse(n['#text']); } catch { /* ignore */ }
    } else if (typeof n === 'object' && n.category === 'htmlTexts' && typeof n['#text'] === 'string') {
      try { htmlTexts = JSON.parse(n['#text']); } catch { /* ignore */ }
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
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '', preserveOrder: false, trimValues: false });
  const obj: any = parser.parse(xlf);
  const units: TranslationUnit[] = [];
  const xliff = obj.xliff;
  if (!xliff) return units;
  const files = Array.isArray(xliff.file) ? xliff.file : [xliff.file];
  const xliffTrg = xliff.trgLang as string | undefined;

  // Flatten text with <ph id> markers recursively
  const flatten = (node: any): string => {
    if (node == null) return '';
    if (typeof node === 'string') return node;
    if (Array.isArray(node)) return node.map(flatten).join('');
    let out = '';
    if (typeof node['#text'] === 'string') out += node['#text'];
    if (node.ph) {
      const phs = Array.isArray(node.ph) ? node.ph : [node.ph];
      for (const ph of phs) {
        const id = ph.id;
        out += `[[ph:${id}]]`;
      }
    }
    if (node.pc) out += flatten(node.pc);
    if (node.source) out += flatten(node.source);
    if (node.target) out += flatten(node.target);
    // gather any nested strings
    for (const [k, v] of Object.entries(node)) {
      if (k !== '#text' && k !== 'ph' && k !== 'source' && k !== 'target' && typeof v === 'string') out += v;
    }
    return out;
  };

  for (const f of files) {
    const fileUnits = Array.isArray(f.unit) ? f.unit : [f.unit];
    const fileTrg = (f.trgLang as string | undefined) || xliffTrg;
    for (const u of fileUnits) {
      const id: string = u.id;
      // notes
      let sheetName = '', col = ''; let row = 0;
      let phMap: Record<string, Record<string, string>> | undefined;
      let htmlSkeleton: string | undefined;
      let htmlInlineMap: Record<string, { open: string; close: string }> | undefined;
      let htmlTexts: string[] | undefined;
      if (u.notes) {
        const { sheetName: sn, row: rr, col: cc, ph, htmlSkeleton: hs, htmlInlineMap: him, htmlTexts: htxt } = extractNotes(u);
        sheetName = sn; row = rr; col = cc; phMap = ph; htmlSkeleton = hs; htmlInlineMap = him; htmlTexts = htxt;
      }

      const segsArr = Array.isArray(u.segment) ? u.segment : (u.segment ? [u.segment] : []);
      const segments: Segment[] = segsArr.map((s: any, idx: number) => {
        const src = flatten(s.source);
        const tgt = s.target ? flatten(s.target) : undefined;
        const sid = s.id || `${id}_s${idx}`;
        return { id: sid, source: src, target: tgt };
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
