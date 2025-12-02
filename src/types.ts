export type CellStyleSnapshot = {
  font?: { name?: string; size?: number; bold?: boolean; italic?: boolean; color?: string };
  alignment?: { horizontal?: string; vertical?: string; wrapText?: boolean };
  fill?: { type?: string; color?: string };
};

export type Segment = {
  id: string;
  source: string;
  target?: string;
  start?: number;
  end?: number;
  meta?: { [k: string]: any };
};

export type TranslationUnit = {
  id: string;
  sheetName: string;
  row: number;
  col: string; // letter
  colIndex: number; // 1-based
  source: string;
  segments?: Segment[];
  richText?: boolean;
  style?: CellStyleSnapshot;
  formula?: string | null;
  isMerged?: boolean;
  mergedRange?: string | null;
  meta?: { [k: string]: any };
};

export type SheetConfig = {
  namePattern: string; // exact or regex
  sourceColumns: string[]; // letters or header names (MVP: letters)
  targetColumns?: { [locale: string]: string | "" };
  createTargetIfMissing?: boolean;
  headerRow?: number;
  valuesStartRow?: number;
  skipHiddenRows?: boolean;
  skipHiddenColumns?: boolean;
  excludeColors?: string[];
  extractFormulaResults?: boolean;
  preserveStyles?: boolean;
  translateComments?: boolean;
  treatMergedRegions?: "top-left" | "expand" | "skip";
  maxCharsPerTarget?: { [locale: string]: number };
  metadataRows?: number[];
  excludedRows?: number[];
  excludedColumns?: string[];
  inlineCodeRegexes?: string[];
  sourceLocale?: string;
  html?: {
    enabled?: boolean; // default true: detect and filter HTML in cell content
    translatableTags?: string[]; // tags whose inner text is translatable; default ['title']
  };
};

export type WorkbookConfig = {
  sheets: SheetConfig[];
};

export type SegConfig = {
  enabled?: boolean;
  rules?: "builtin" | { srxPath: string };
};

export type GlobalConfig = {
  overwrite?: boolean;
  insertTargetPlacement?: "insertAfterSource" | "appendToSheetEnd";
  srcLang?: string;
  targetLocale?: string; // used during merge when not derived from units
  exportComments?: boolean; // include comments/header/metadata in XLIFF notes
  mergeFallback?: 'source' | 'empty'; // when a segment target is missing: use source or leave empty
  xliffVersion?: '1.2' | '2.1'; // XLIFF version for export, default 2.1
};

export type Config = { workbook: WorkbookConfig; segmentation?: SegConfig; global?: GlobalConfig };
