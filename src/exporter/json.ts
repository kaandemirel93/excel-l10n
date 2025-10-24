import { Config, TranslationUnit } from '../types.js';

function regexesForUnit(u: TranslationUnit, config: Config): RegExp[] {
  const sheetCfg = config.workbook.sheets.find(s => {
    try { return new RegExp(s.namePattern).test(u.sheetName) || s.namePattern === u.sheetName; } catch { return s.namePattern === u.sheetName; }
  });
  const res = (sheetCfg?.inlineCodeRegexes || []).map(r => {
    try { return new RegExp(r, 'g'); } catch { return null; }
  }).filter((x): x is RegExp => !!x);
  return res;
}

function computePlaceholderMap(text: string, regs: RegExp[]): Record<string, string> {
  const map: Record<string, string> = {};
  let idx = 1;
  for (const re of regs) {
    text.replace(re, (m) => {
      const id = `ph${idx++}`;
      if (!map[id]) map[id] = m;
      return m;
    });
  }
  return map;
}

export async function exportToJson(units: TranslationUnit[], config?: Config, meta?: Record<string, any>): Promise<string> {
  let enriched = units;
  if (config) {
    enriched = units.map(u => {
      const regs = regexesForUnit(u, config);
      if (!regs.length) return u;
      const mapBySeg: Record<string, Record<string, string>> = {};
      const segs = (u.segments && u.segments.length) ? u.segments : [{ id: `${u.id}_s0`, source: u.source }];
      for (const s of segs) {
        mapBySeg[s.id] = computePlaceholderMap(s.source, regs);
      }
      const metaObj = { ...(u.meta || {}), placeholders: mapBySeg };
      return { ...u, meta: metaObj };
    });
  }
  return JSON.stringify({ meta: meta ?? {}, units: enriched }, null, 2);
}

export function parseJsonUnits(jsonStr: string): TranslationUnit[] {
  const obj = JSON.parse(jsonStr);
  return obj.units as TranslationUnit[];
}
