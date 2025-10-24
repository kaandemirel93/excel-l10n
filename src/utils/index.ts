import { TranslationUnit } from '../types.js';

export function colIndexToLetter(n: number): string {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function colLetterToIndex(s: string): number {
  let n = 0;
  for (const ch of s.toUpperCase()) {
    if (ch < 'A' || ch > 'Z') break;
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n;
}

export function makeTuId(sheetName: string, row: number, colLetter: string): string {
  return `${encodeURIComponent(sheetName)}::R${row}C${colLetter.toUpperCase()}`;
}

export function compactUnits(units: TranslationUnit[]): TranslationUnit[] {
  return units.map(u => ({ ...u, meta: undefined }));
}
