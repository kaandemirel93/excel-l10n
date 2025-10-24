import { TranslationUnit, Config } from '../types.js';

export type PseudoOptions = {
  wrap?: { left: string; right: string };
  expandRate?: number; // e.g., 0.3 for +30%
};

const defaultMap: Record<string, string> = {
  A: 'Â', B: 'Ɓ', C: 'Ç', D: 'Ð', E: 'Ë', F: 'Ƒ', G: 'Ğ', H: 'Ĥ', I: 'Ï', J: 'Ĵ', K: 'Ķ', L: 'Ļ', M: 'Ṁ', N: 'Ñ', O: 'Ø', P: 'Ṕ', Q: 'Ɋ', R: 'Ŕ', S: 'Š', T: 'Ŧ', U: 'Ū', V: 'Ṽ', W: 'Ŵ', X: 'Ẋ', Y: 'Ÿ', Z: 'Ž',
  a: 'ã', b: 'ƀ', c: 'ç', d: 'đ', e: 'ë', f: 'ƒ', g: 'ğ', h: 'ħ', i: 'ï', j: 'ĵ', k: 'ķ', l: 'ĺ', m: 'ɱ', n: 'ñ', o: 'ø', p: 'þ', q: 'ʠ', r: 'ř', s: 'š', t: 'ŧ', u: 'ü', v: 'ṽ', w: 'ŵ', x: 'ẋ', y: 'ÿ', z: 'ž',
};

const PLACEHOLDER_RE = /(\[\[ph:[^\]]+\]\]|\{\d+\}|%s)/g;

export function pseudoTransform(text: string, opts?: PseudoOptions): string {
  const wrap = opts?.wrap ?? { left: '⟦', right: '⟧' };
  const rate = typeof opts?.expandRate === 'number' ? opts!.expandRate! : 0.3;
  const parts = String(text).split(PLACEHOLDER_RE);
  let out = '';
  for (const p of parts) {
    if (!p) continue;
    if (PLACEHOLDER_RE.test(p)) {
      out += p; // preserve placeholders untouched
    } else {
      const mapped = p.split('').map(ch => defaultMap[ch] ?? ch).join('');
      const extra = Math.ceil(mapped.length * rate);
      out += mapped + (extra > 0 ? 'ø'.repeat(extra) : '');
    }
  }
  return `${wrap.left}${out}${wrap.right}`;
}

export function pseudoUnits(units: TranslationUnit[], _config?: Config, options?: PseudoOptions): TranslationUnit[] {
  return units.map(u => ({
    ...u,
    segments: (u.segments && u.segments.length ? u.segments : [{ id: `${u.id}_s0`, source: u.source }]).map(s => ({
      ...s,
      target: pseudoTransform(s.source, options),
    })),
  }));
}
