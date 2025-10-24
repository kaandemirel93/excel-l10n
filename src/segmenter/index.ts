import fs from 'node:fs';
import path from 'node:path';
import { XMLParser } from 'fast-xml-parser';
import { Config, Segment, TranslationUnit } from '../types.js';

type SrxRule = { break: boolean; before: RegExp; after: RegExp };
type SrxBundle = { rules: SrxRule[] };

function builtinRulesFor(lang: string): SrxBundle {
  // Minimal pragmatic defaults: break after ., !, ?, ; when followed by space or EoT, avoiding common abbreviations
  const abbrev = /(Mr|Mrs|Ms|Dr|Prof|Sr|Jr|vs|etc)\.$/i;
  const before = /[.!?;]+\)?\”?\»?\s*$/; // punctuation possibly followed by closing token
  const after = /^\s*\(?\“?\«?[A-Z0-9]/; // next sentence likely starts with uppercase/number
  return {
    rules: [
      { break: true, before: new RegExp(before), after: new RegExp(after) },
      // no-break when abbreviation before
      { break: false, before: new RegExp(abbrev), after: new RegExp(after) },
    ],
  };
}

function loadSrxFromFile(srxPath: string): { map: { pattern: string; name: string }[]; rules: Record<string, SrxRule[]> } {
  const xml = fs.readFileSync(srxPath, 'utf-8');
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '', trimValues: false });
  const doc: any = parser.parse(xml);
  const body = doc?.srx?.body;
  const maprules = body?.maprules;
  const languagerules = body?.languagerules;
  const maps: { pattern: string; name: string }[] = [];
  const rules: Record<string, SrxRule[]> = {};

  const lrList = Array.isArray(languagerules?.languagerule) ? languagerules.languagerule : (languagerules?.languagerule ? [languagerules.languagerule] : []);
  for (const lr of lrList) {
    const name: string = lr.name;
    const ruleList = Array.isArray(lr.rule) ? lr.rule : (lr.rule ? [lr.rule] : []);
    const compiled: SrxRule[] = [];
    for (const r of ruleList) {
      const br = (r.break || r['break']) === 'yes';
      const before = r.beforebreak ?? '';
      const after = r.afterbreak ?? '';
      try {
        compiled.push({ break: br, before: new RegExp(before), after: new RegExp(after) });
      } catch {
        // ignore invalid rule
      }
    }
    rules[name] = compiled;
  }

  const mapList = Array.isArray(maprules?.languagemap) ? maprules.languagemap : (maprules?.languagemap ? [maprules.languagemap] : []);
  for (const m of mapList) {
    maps.push({ pattern: m.languagepattern, name: m.languagerulename });
  }
  return { map: maps, rules };
}

function pickRuleSetForLocale(locale: string, srx: { map: { pattern: string; name: string }[]; rules: Record<string, SrxRule[]> }): SrxRule[] | null {
  for (const m of srx.map) {
    try {
      const re = new RegExp(m.pattern, 'i');
      if (re.test(locale)) return srx.rules[m.name] || null;
    } catch {
      // bad map entry, ignore
    }
  }
  return null;
}

function segmentTextByRules(text: string, bundle: SrxBundle): Segment[] {
  if (!text) return [{ id: '', source: '' }];
  const cutPositions: number[] = [];
  // scan characters and check rule pairs around boundary i
  for (let i = 1; i < text.length; i++) {
    const left = text.slice(0, i);
    const right = text.slice(i);
    let decision: boolean | null = null; // true=break, false=no-break, null=no rule matched
    for (const rule of bundle.rules) {
      if (rule.before.test(left) && rule.after.test(right)) {
        decision = rule.break;
        // SRX applies in order; first matching rule decides
        break;
      }
    }
    if (decision === true) cutPositions.push(i);
  }
  const segments: Segment[] = [];
  let prev = 0;
  let idx = 0;
  for (const pos of [...cutPositions, text.length]) {
    const s = text.slice(prev, pos).trim();
    if (s) segments.push({ id: `s${idx++}`, source: s, start: prev, end: pos });
    prev = pos;
  }
  if (!segments.length) segments.push({ id: 's0', source: text, start: 0, end: text.length });
  return segments;
}

function getLocaleForUnit(u: TranslationUnit, config: Config): string {
  const sheetCfg = config.workbook.sheets.find(s => {
    try { return new RegExp(s.namePattern).test(u.sheetName) || s.namePattern === u.sheetName; } catch { return s.namePattern === u.sheetName; }
  });
  return sheetCfg?.sourceLocale || config.global?.srcLang || 'en';
}

export function segmentUnits(units: TranslationUnit[], config: Config): TranslationUnit[] {
  const enabled = config.segmentation?.enabled !== false;
  if (!enabled) {
    return units.map(u => ({ ...u, segments: [{ id: `${u.id}_s0`, source: u.source }] }));
  }

  let srxRules: { map: { pattern: string; name: string }[]; rules: Record<string, SrxRule[]> } | null = null;
  if (config.segmentation?.rules && typeof config.segmentation.rules === 'object' && config.segmentation.rules.srxPath) {
    const p = path.resolve(config.segmentation.rules.srxPath);
    try { srxRules = loadSrxFromFile(p); } catch { srxRules = null; }
  }

  return units.map(u => {
    const locale = getLocaleForUnit(u, config);
    let bundle: SrxBundle | null = null;
    if (srxRules) {
      const set = pickRuleSetForLocale(locale, srxRules);
      if (set && set.length) bundle = { rules: set };
    }
    if (!bundle) bundle = builtinRulesFor(locale);

    const segs = segmentTextByRules(u.source, bundle).map((s, i) => ({ id: `${u.id}_s${i}`, source: s.source, start: s.start, end: s.end }));
    return { ...u, segments: segs };
  });
}
