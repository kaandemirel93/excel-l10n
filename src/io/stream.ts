import ExcelJS from 'exceljs';
import { Config, TranslationUnit } from '../types.js';
import { colLetterToIndex, colIndexToLetter } from '../utils/index.js';

export async function* extractStreamWorkbook(inputXlsxPath: string, config: Config): AsyncGenerator<TranslationUnit> {
  const reader = new (ExcelJS as any).stream.xlsx.WorkbookReader(inputXlsxPath, { entries: 'emit', sharedStrings: 'cache', styles: 'cache', hyperlinks: 'emit', worksheets: 'emit' });

  const sheets = config.workbook.sheets;
  const matchSheet = (name: string) => sheets.filter(s => {
    try { return new RegExp(s.namePattern).test(name) || s.namePattern === name; } catch { return s.namePattern === name; }
  });

  for await (const ws of reader as any) {
    const wsa: any = ws;
    const sheetName: string = wsa.name || wsa.id || 'Sheet';
    const matches = matchSheet(sheetName);
    if (matches.length === 0) continue;

    for (const sheetCfg of matches) {
      const sourceCols = (sheetCfg.sourceColumns || []).map(colLetterToIndex);
      const valuesStart = sheetCfg.valuesStartRow ?? 2;
      const headerRow = sheetCfg.headerRow ?? 1;
      const metaRows: number[] = Array.isArray((sheetCfg as any).metadataRows) ? (sheetCfg as any).metadataRows : [];

      // caches for meta capture
      const headerCache: Record<number, string> = {};
      const metaCache: Record<number, Record<number, string>> = {};

      let rowIndex = 0;
      for await (const row of wsa) {
        rowIndex = (row as any).number || rowIndex + 1;

        // Capture header row values for requested columns regardless of hidden state
        if (headerRow && rowIndex === headerRow) {
          for (const srcIdx of sourceCols) {
            const cell = (row as any).getCell ? (row as any).getCell(srcIdx) : undefined;
            const v: any = cell?.value;
            const text = v == null ? '' : (typeof v === 'string' ? v : String(v));
            if (text) headerCache[srcIdx] = text;
          }
          continue; // don't yield units from header row
        }

        // Capture metadata rows if configured
        if (metaRows.includes(rowIndex)) {
          const m: Record<number, string> = {};
          for (const srcIdx of sourceCols) {
            const cell = (row as any).getCell ? (row as any).getCell(srcIdx) : undefined;
            const v: any = cell?.value;
            const text = v == null ? '' : (typeof v === 'string' ? v : String(v));
            if (text) m[srcIdx] = text;
          }
          if (Object.keys(m).length) metaCache[rowIndex] = m;
          continue; // skip yielding for meta rows
        }

        // Apply row filters for value rows
        if (rowIndex < valuesStart) continue;
        if ((sheetCfg.skipHiddenRows && (row as any).hidden) || (sheetCfg.excludedRows && sheetCfg.excludedRows.includes(rowIndex))) continue;

        for (const srcIdx of sourceCols) {
          const cell = (row as any).getCell ? (row as any).getCell(srcIdx) : undefined;
          const v: any = cell?.value;
          const text = v == null ? '' : (typeof v === 'string' ? v : String(v));
          if (!text) continue;
          const colLetter = colIndexToLetter(srcIdx);
          const meta: Record<string, any> = {};
          if (headerCache[srcIdx]) meta.headerName = headerCache[srcIdx];
          if (metaRows.length) {
            const m: Record<number, string> = {};
            for (const mr of metaRows) {
              const map = metaCache[mr];
              if (map && map[srcIdx]) m[mr] = map[srcIdx];
            }
            if (Object.keys(m).length) meta.metadataRows = m as any;
          }
          const tu: TranslationUnit = {
            id: `${sheetName}::R${rowIndex}C${colLetter}`,
            sheetName,
            row: rowIndex,
            col: colLetter,
            colIndex: srcIdx,
            source: text,
            segments: [],
            meta,
          } as any;
          yield tu;
        }
      }
    }
  }
}
