import fs from 'node:fs';
import { Config, Segment, TranslationUnit } from '../types.js';

function regexesForUnit(u: TranslationUnit, config: Config): RegExp[] {
  const sheetCfg = config.workbook.sheets.find(s => {
    try { return new RegExp(s.namePattern).test(u.sheetName) || s.namePattern === u.sheetName; } catch { return s.namePattern === u.sheetName; }
  });
  const res = (sheetCfg?.inlineCodeRegexes || []).map(r => {
    try { return new RegExp(r, 'g'); } catch { return null as any; }
  }).filter(Boolean) as RegExp[];
  return res;
}

const PH_MARK_RE = /(\[\[ph:[^\]]+\]\])/g;

function encodePlaceholders(text: string, regs: RegExp[]): { encoded: string; map: Record<string, string> } {
  const map: Record<string, string> = {};
  let phIndex = 1;
  // ICU protect (simple)
  const ICU_BLOCK = /\{[^{}]*,\s*(plural|select)\s*,[\s\S]*?\}/g;
  let icuIdx = 1;
  const icuMap: Record<string, string> = {};
  let encoded = text.replace(ICU_BLOCK, (m) => m.replace(/\{([^{}]*)\}/g, (_inner, grp) => {
    const token = `icu${icuIdx++}`;
    icuMap[token] = `{${grp}}`;
    return `[[ph:${token}]]`;
  }));
  for (const k of Object.keys(icuMap)) map[k] = icuMap[k];
  for (const re of regs) {
    encoded = encoded.replace(re, (m) => {
      const id = `ph${phIndex++}`;
      map[id] = m;
      return `[[ph:${id}]]`;
    });
  }
  return { encoded, map };
}

function writeWithPhStr(encoded: string): string {
  // split by markers and write <ph id>
  const parts = encoded.split(PH_MARK_RE);
  let out = '';
  for (const part of parts) {
    if (!part) continue;
    const m = /^\[\[ph:([^\]]+)\]\]$/.exec(part);
    if (m) out += `<ph id="${m[1]}"/>`;
    else out += escapeXml(part);
  }
  return out;
}

function escapeXml(s: string): string {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

export async function exportToXliffStreamFromIterator(
  iter: AsyncIterable<TranslationUnit>,
  config: Config,
  options: { srcLang?: string; generator?: string },
  outPath: string
): Promise<void> {
  const srcLang = options.srcLang || config.global?.srcLang || 'en';
  const ws = fs.createWriteStream(outPath, { encoding: 'utf-8' });
  ws.write(`<?xml version="1.0" encoding="UTF-8"?>\n`);
  ws.write(`<xliff version="2.1" srcLang="${srcLang}">\n`);
  ws.write(`  <file id="workbook" original="workbook.xlsx" tool-id="${options.generator || 'excel-l10n'}">\n`);

  for await (const u of iter) {
    const regs = regexesForUnit(u, config);
    ws.write(`    <unit id="${escapeXml(u.id)}">\n`);
    ws.write(`      <notes><note>sheet=${escapeXml(u.sheetName)};row=${u.row};col=${escapeXml(u.col)}</note></notes>\n`);
    const segs: Segment[] = (u.segments && u.segments.length) ? u.segments : [{ id: `${u.id}_s0`, source: u.source } as Segment];
    const phMap: Record<string, Record<string, string>> = {};
    for (const s of segs) {
      const { encoded, map } = encodePlaceholders(s.source, regs);
      phMap[s.id] = map;
      ws.write(`      <segment id="${escapeXml(s.id)}">\n`);
      ws.write(`        <source>${writeWithPhStr(encoded)}</source>\n`);
      if (s.target) ws.write(`        <target>${writeWithPhStr(s.target)}</target>\n`);
      ws.write(`      </segment>\n`);
    }
    if (Object.keys(phMap).length) {
      ws.write(`      <notes><note category="ph">${escapeXml(JSON.stringify(phMap))}</note></notes>\n`);
    }
    ws.write(`    </unit>\n`);
  }

  ws.write(`  </file>\n`);
  ws.write(`</xliff>\n`);
  await new Promise<void>((resolve, reject) => ws.end(resolve));
}
