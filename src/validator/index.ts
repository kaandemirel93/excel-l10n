import fs from 'node:fs';
import path from 'node:path';
import { parseTranslated } from '../index.js';
import type { TranslationUnit } from '../types.js';

export type ValidateOptions = {
  lengthFactor?: number; // default 2
  formatOverride?: 'xlf' | 'json';
};

export type Finding = { level: 'error' | 'warn' | 'info'; message: string; unitId?: string; locale?: string };
export type Report = { items: Finding[] };

const PLACEHOLDER_RE = /(\{\d+\}|\[\[ph:[^\]]+\]\])/g;
const ICU_BLOCK_RE = /\{\s*[^{}]+,\s*(plural|select)\s*,[\s\S]*?\}/g;

function listPlaceholders(s: string): string[] {
  return (String(s).match(PLACEHOLDER_RE) || []).sort();
}

function listIcuCategories(s: string): string[] {
  const cats = new Set<string>();
  const blocks = String(s).match(ICU_BLOCK_RE) || [];
  for (const b of blocks) {
    const m = b.match(/\b(one|other|few|many|two|zero)\b/g);
    if (m) m.forEach(x => cats.add(x));
  }
  return Array.from(cats).sort();
}

function validateUnit(u: TranslationUnit, opts: ValidateOptions, locale?: string): Finding[] {
  const out: Finding[] = [];
  const segs = (u.segments && u.segments.length) ? u.segments : [{ id: `${u.id}_s0`, source: u.source, target: (u as any).target }];
  for (const s of segs) {
    const src = s.source ?? '';
    const tgt = s.target ?? '';
    // missing targets
    if (!tgt || String(tgt).trim().length === 0) {
      out.push({ level: 'error', message: `Missing target for segment ${s.id}`, unitId: u.id, locale });
    }
    // placeholder mismatch
    const srcPh = listPlaceholders(src);
    const tgtPh = listPlaceholders(tgt);
    if (srcPh.join('|') !== tgtPh.join('|')) {
      out.push({ level: 'error', message: `Placeholder mismatch in ${s.id}: src=${srcPh.join(',')} tgt=${tgtPh.join(',')}`, unitId: u.id, locale });
    }
    // length check
    const factor = opts.lengthFactor ?? 2;
    if (String(src).length > 0) {
      const ratio = String(tgt).length / String(src).length;
      if (ratio > factor) out.push({ level: 'warn', message: `Segment ${s.id} length ratio ${ratio.toFixed(2)} > ${factor}`, unitId: u.id, locale });
    }
    // ICU categories
    const srcCats = listIcuCategories(src);
    const tgtCats = listIcuCategories(tgt);
    if (srcCats.length && srcCats.join('|') !== tgtCats.join('|')) {
      out.push({ level: 'error', message: `ICU categories differ in ${s.id}: src=${srcCats.join(',')} tgt=${tgtCats.join(',')}`, unitId: u.id, locale });
    }
  }
  return out;
}

export async function validateFiles(files: string[], opts: ValidateOptions): Promise<Report> {
  const items: Finding[] = [];
  for (const fp of files) {
    try {
      const raw = fs.readFileSync(fp, 'utf-8');
      const fmt = opts.formatOverride || (/(\.json)$/i.test(fp) ? 'json' : 'xlf');
      const units = parseTranslated(raw, fmt);
      const locale = detectTargetLocale(raw, fmt);
      for (const u of units) {
        items.push(...validateUnit(u, opts, locale));
      }
    } catch (e: any) {
      items.push({ level: 'error', message: `Failed to validate ${path.basename(fp)}: ${e?.message || e}` });
    }
  }
  return { items };
}

function detectTargetLocale(raw: string, fmt: 'xlf' | 'json'): string | undefined {
  try {
    if (fmt === 'json') return undefined;
    const m = raw.match(/\btrgLang="([^"]+)"/);
    return m?.[1];
  } catch { return undefined; }
}
