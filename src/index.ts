import { Config, TranslationUnit } from './types.js';
import { parseConfig } from './config/index.js';
import { extractUnits } from './io/excel.js';
import { segmentUnits } from './segmenter/index.js';
import { exportToXliff, parseXliffToUnits } from './exporter/xliff.js';
import { exportToJson, parseJsonUnits } from './exporter/json.js';
import { mergeWorkbook } from './merger/index.js';

export type { Config, TranslationUnit } from './types.js';
export { parseConfig };

export async function extract(inputXlsxPath: string, config: Config): Promise<TranslationUnit[]> {
  const units = await extractUnits(inputXlsxPath, config);
  const segmented = segmentUnits(units, config);
  return segmented;
}

export async function exportUnitsToXliff(
  units: TranslationUnit[],
  config: Config,
  options?: { srcLang?: string; trgLang?: string; generator?: string }
): Promise<string> {
  return exportToXliff(units, config, options);
}

export async function exportUnitsToJson(units: TranslationUnit[], config?: Config, meta?: Record<string, any>): Promise<string> {
  return exportToJson(units, config, meta);
}

export async function merge(inputXlsxPath: string, outputXlsxPath: string, translatedUnits: TranslationUnit[], config: Config): Promise<void> {
  await mergeWorkbook(inputXlsxPath, outputXlsxPath, translatedUnits, config);
}

export function parseTranslated(input: string, format: 'xlf' | 'json'): TranslationUnit[] {
  return format === 'xlf' ? parseXliffToUnits(input) : parseJsonUnits(input);
}

// Experimental streaming scaffold: currently yields units after a normal extract.
// Replace with a true exceljs.stream-based reader in a subsequent iteration.
export async function* extractStream(inputXlsxPath: string, config: Config): AsyncGenerator<TranslationUnit> {
  const all = await extract(inputXlsxPath, config);
  for (const u of all) yield u;
}
