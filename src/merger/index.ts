import ExcelJS from 'exceljs';
import { Config, TranslationUnit } from '../types.js';
import { colLetterToIndex } from '../utils/index.js';

type InlineMap = Record<string, { open?: string; close?: string } | undefined>;
type PlaceholderMap = Record<string, Record<string, string>>;

function ensureTargetColumn(
  ws: ExcelJS.Worksheet,
  desiredLetter: string | '',
  placement: 'insertAfterSource' | 'appendToSheetEnd',
  sourceColIndex: number
): number {
  if (desiredLetter && desiredLetter.trim()) {
    return colLetterToIndex(desiredLetter);
  }
  if (placement === 'insertAfterSource') {
    const insertAt = sourceColIndex + 1;
    ws.spliceColumns(insertAt, 0, []);
    return insertAt;
  }
  const newIndex = ws.columnCount + 1;
  ws.spliceColumns(newIndex, 0, []);
  return newIndex;
}

// Safe replaceAll for older runtimes (split/join)
function rAll(haystack: string, needle: string, replacement: string): string {
  return haystack.split(needle).join(replacement);
}

function tokenizeSkeleton(skeleton: string): Array<{ type: 'tag' | 'text' | 'io' | 'ic' | 'htmltxt'; value: string; id?: string }> {
  const tokens: Array<{ type: 'tag' | 'text' | 'io' | 'ic' | 'htmltxt'; value: string; id?: string }> = [];
  const re = /(\[\[(?:htmltxt|io|ic):\d+\]\]|<[^>]+>)/g;
  let lastIdx = 0;
  let match: RegExpExecArray | null;
  while ((match = re.exec(skeleton)) !== null) {
    if (match.index > lastIdx) {
      tokens.push({ type: 'text', value: skeleton.slice(lastIdx, match.index) });
    }
    const token = match[1];
    if (token.startsWith('[[htmltxt:')) {
      const id = token.slice('[[htmltxt:'.length, -2);
      tokens.push({ type: 'htmltxt', value: token, id });
    } else if (token.startsWith('[[io:')) {
      const id = token.slice('[[io:'.length, -2);
      tokens.push({ type: 'io', value: token, id });
    } else if (token.startsWith('[[ic:')) {
      const id = token.slice('[[ic:'.length, -2);
      tokens.push({ type: 'ic', value: token, id });
    } else {
      tokens.push({ type: 'tag', value: token });
    }
    lastIdx = re.lastIndex;
  }
  if (lastIdx < skeleton.length) {
    tokens.push({ type: 'text', value: skeleton.slice(lastIdx) });
  }
  return tokens;
}

type TranslationToken =
  | { type: 'text'; value: string }
  | { type: 'placeholder'; value: string; kind: 'io' | 'ic'; id: string };

function tokenizeTranslation(text: string): TranslationToken[] {
  const tokens: TranslationToken[] = [];
  const re = /(\[\[(?:io|ic):\d+\]\])/g;
  let lastIdx = 0;
  let match: RegExpExecArray | null;
  while ((match = re.exec(text)) !== null) {
    if (match.index > lastIdx) {
      tokens.push({ type: 'text', value: text.slice(lastIdx, match.index) });
    }
    const token = match[1];
    const kind = token.includes('[[io:') ? 'io' : 'ic';
    const id = token.slice(kind === 'io' ? '[[io:'.length : '[[ic:'.length, -2);
    tokens.push({ type: 'placeholder', value: token, kind, id });
    lastIdx = re.lastIndex;
  }
  if (lastIdx < text.length) {
    tokens.push({ type: 'text', value: text.slice(lastIdx) });
  }
  return tokens;
}

function composeHtmlFromSkeleton(translated: string, skeleton: string, inlineMap?: InlineMap): string {
  const skeletonTokens = tokenizeSkeleton(skeleton);
  const translationTokens = tokenizeTranslation(translated);
  let txIdx = 0;
  let result = '';

  for (const token of skeletonTokens) {
    if (token.type === 'tag' || token.type === 'text') {
      result += token.value;
      continue;
    }
    if (token.type === 'io' || token.type === 'ic') {
      const entry = inlineMap?.[token.id ?? ''];
      const replacement = token.type === 'io' ? entry?.open ?? '' : entry?.close ?? '';
      result += replacement;
      const current = translationTokens[txIdx];
      if (current && current.type === 'placeholder' && current.kind === token.type && current.id === token.id) {
        txIdx++;
      }
      continue;
    }
    if (token.type === 'htmltxt') {
      let chunk = '';
      while (txIdx < translationTokens.length && translationTokens[txIdx].type === 'text') {
        chunk += translationTokens[txIdx].value;
        txIdx++;
      }
      result += chunk;
      continue;
    }
  }

  // append any residual text tokens (should normally not happen)
  while (txIdx < translationTokens.length) {
    const tail = translationTokens[txIdx++];
    if (tail.type === 'text') result += tail.value;
  }

  return result;
}

export async function mergeWorkbook(
  inputXlsxPath: string,
  outputXlsxPath: string,
  translatedUnits: TranslationUnit[],
  config: Config
): Promise<void> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(inputXlsxPath);

  const overwrite = config.global?.overwrite ?? true;
  const placement = config.global?.insertTargetPlacement ?? 'insertAfterSource';

  for (const sheetCfg of config.workbook.sheets) {
    for (const ws of wb.worksheets.filter(
      w => new RegExp(sheetCfg.namePattern).test(w.name) || sheetCfg.namePattern === w.name
    )) {
      const localeToCol = sheetCfg.targetColumns || {};
      const explicitCols = new Map<string, string>();
      let autoCreateCount = 0;

      for (const [loc, col] of Object.entries(localeToCol)) {
        const colLetter = (col || '').toUpperCase();
        if (!colLetter) { autoCreateCount++; continue; }
        if (explicitCols.has(colLetter)) {
          throw new Error(
            `Target column collision on sheet '${ws.name}': locales '${explicitCols.get(colLetter)}' and '${loc}' both map to column '${colLetter}'.`
          );
        }
        explicitCols.set(colLetter, loc);
      }
      if (autoCreateCount > 1 && sheetCfg.createTargetIfMissing) {
        throw new Error(
          `Multiple locales configured to auto-create target columns on sheet '${ws.name}'. Define explicit targetColumns to avoid collisions.`
        );
      }

      const unitsForSheet = translatedUnits.filter(u => u.sheetName === ws.name);
      for (const tu of unitsForSheet) {
        const srcColIdx = colLetterToIndex(tu.col);
        const preferredLocale = config.global?.targetLocale || (tu.meta as any)?.targetLocale;
        const entries = preferredLocale
          ? Object.entries(localeToCol).filter(([lc]) => lc === preferredLocale)
          : Object.entries(localeToCol);

        for (const [, colLetter] of entries) {
          const targetIdx = ensureTargetColumn(ws, colLetter, placement, srcColIdx);
          const cell = ws.getCell(tu.row, targetIdx);
          if (!overwrite) continue;

          const rebuiltSegments: string[] = [];
          let pos = 0;
          const src = tu.source || '';

          if (tu.segments && tu.segments.length) {
            const unitInlineMap = (tu.meta as any)?.htmlInlineMap as InlineMap | undefined;
            const placeholderMap = (tu.meta as any)?.placeholders as PlaceholderMap | undefined;
            for (let i = 0; i < tu.segments.length; i++) {
              const seg = tu.segments[i];
              const chosen =
                (seg.target && seg.target.length > 0)
                  ? seg.target
                  : (config.global?.mergeFallback ?? 'source') === 'source'
                    ? (seg.source ?? '')
                    : '';

              const skeleton = seg.meta?.htmlSkeleton;

              // 1) restore the gap between last pos and current segment's source
              const idx = seg.source ? src.indexOf(seg.source, pos) : -1;
              let gap = '';
              if (idx >= 0) {
                const between = src.slice(pos, idx);
                gap = between.length > 0 ? between : (pos > 0 ? ' ' : '');
                pos = idx + (seg.source?.length ?? 0);
              }

              let segBody = chosen;
              if (placeholderMap && seg.id && placeholderMap[seg.id]) {
                for (const [pid, original] of Object.entries(placeholderMap[seg.id])) {
                  segBody = rAll(segBody, `[[ph:${pid}]]`, original);
                }
              }

              // 2) inject into legacy segment skeleton if present (backward compatibility)
              if (skeleton) {
                const tokenRegex = /\[\[htmltxt:\d+\]\]/g;
                let expanded = segBody;
                if (unitInlineMap) {
                  for (const [n, t] of Object.entries(unitInlineMap)) {
                    expanded = rAll(expanded, `[[io:${n}]]`, t?.open ?? '');
                    expanded = rAll(expanded, `[[ic:${n}]]`, t?.close ?? '');
                  }
                }
                segBody = skeleton.replace(tokenRegex, expanded);
              }

              rebuiltSegments.push(gap + segBody);
            }

            // trailing text after last segment, if any
            rebuiltSegments.push(src.slice(pos));
          } else {
            // No segments: fall back to TU source (tu.target no longer exists)
            rebuiltSegments.push(tu.source ?? '');
          }

          let finalText = rebuiltSegments.join('');
          const htmlSkeleton = (tu.meta as any)?.htmlSkeleton as string | undefined;
          const htmlInlineMap = (tu.meta as any)?.htmlInlineMap as InlineMap | undefined;
          if (htmlSkeleton) {
            finalText = composeHtmlFromSkeleton(finalText, htmlSkeleton, htmlInlineMap);
          } else if ((tu.meta as any)?.htmlInlineMap && !(tu.meta as any)?.htmlSkeleton) {
            // Inline map without skeleton: expand placeholders globally
            for (const [n, t] of Object.entries((tu.meta as any).htmlInlineMap as InlineMap)) {
              finalText = rAll(finalText, `[[io:${n}]]`, t?.open ?? '');
              finalText = rAll(finalText, `[[ic:${n}]]`, t?.close ?? '');
            }
          }

          if (finalText !== '') {
            cell.value = finalText;
          }

          if (sheetCfg.preserveStyles) {
            if (tu.style) {
              const style: any = {};
              if (tu.style.font) {
                style.font = {
                  ...('name' in tu.style.font! ? { name: tu.style.font!.name } : {}),
                  ...('size' in tu.style.font! ? { size: tu.style.font!.size } : {}),
                  ...('bold' in tu.style.font! ? { bold: tu.style.font!.bold } : {}),
                  ...('italic' in tu.style.font! ? { italic: tu.style.font!.italic } : {}),
                  ...('color' in tu.style.font! && tu.style.font!.color
                    ? { color: { argb: `FF${tu.style.font!.color!.replace('#', '').toUpperCase()}` } }
                    : {})
                };
              }
              if (tu.style.alignment) {
                style.alignment = { ...tu.style.alignment };
              }
              if (tu.style.fill && tu.style.fill.color) {
                style.fill = {
                  type: tu.style.fill.type || 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: `FF${tu.style.fill.color.replace('#', '').toUpperCase()}` }
                };
              }
              cell.style = { ...cell.style, ...style };
            } else {
              const srcCell = ws.getCell(tu.row, srcColIdx);
              cell.style = { ...srcCell.style };
            }
          }
        }
      }
    }
  }

  await wb.xlsx.writeFile(outputXlsxPath);
}

export { composeHtmlFromSkeleton as _composeHtmlFromSkeleton };
