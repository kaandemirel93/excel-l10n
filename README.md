# excel-l10n

A configurable Excel (XLSX) extraction → segmentation → XLIFF/JSON export → merge tool for modern JS workflows. Inspired by Okapi's OpenXML filter, with native support for multi-lingual target columns per sheet and rich Excel filter options via a simple JSON/YAML configuration.

## Quick start

```bash
# install deps
npm install

# build
npm run build

# extract to XLIFF (uses SRX if enabled in config)
dist/cli/index.js extract -c examples/config.yml -i examples/sample.xlsx -o out.xlf --src-lang en

# merge translated file
dist/cli/index.js merge -c examples/config.yml -i examples/sample.xlsx -t out.translated.xlf -o examples/sample.translated.xlsx
```

## CLI

- `excel-l10n extract -c config.yml -i workbook.xlsx -o out.xlf`
- `excel-l10n extract -c config.yml -i workbook.xlsx -o out.json --format json`
- `excel-l10n merge -c config.yml -i workbook.xlsx -t translated.xlf -o workbook.translated.xlsx`

Advanced:

- Per-locale XLIFF export (one file per target language):

  ```bash
  excel-l10n extract -c config.yml -i workbook.xlsx -o out.xlf --per-locale
  # emits out.fr.xlf, out.de.xlf, ... (based on targetColumns in config)
  ```

- Single bilingual XLIFF with explicit target language:

  ```bash
  excel-l10n extract -c config.yml -i workbook.xlsx -o out.fr.xlf --target-lang fr
  ```

- Merge all translations in one run (auto-detect trgLang per file):

  ```bash
  # input is a directory containing .xlf/.xliff/.json files
  excel-l10n merge -c config.yml -i workbook.xlsx -t ./translated/ -o workbook.merged.xlsx

  # or a comma-separated list
  excel-l10n merge -c config.yml -i workbook.xlsx -t out.fr.xlf,out.de.xlf -o workbook.merged.xlsx
  ```

Quick inline (no config file):

```bash
# extract with inline flags
excel-l10n extract -i in.xlsx --sheet "Sheet1" --source A --target fr=B,de=C -o out.xlf --src-lang en
```

Run `excel-l10n --help` for details.

## Programmatic API

```ts
import { parseConfig, extract, exportUnitsToXliff, exportUnitsToJson, parseTranslated, merge } from 'excel-l10n';

const cfg = parseConfig('examples/config.yml');
const units = await extract('examples/sample.xlsx', cfg);
const xlf = await exportUnitsToXliff(units, cfg, { srcLang: 'en' });
const json = await exportUnitsToJson(units, cfg, { fileName: 'sample.xlsx' });

// translate, then merge
const translatedUnits = parseTranslated(xlf, 'xlf');
await merge('examples/sample.xlsx', 'examples/sample.translated.xlsx', translatedUnits, cfg);
```

## Segmentation (SRX)

- SRX rules are supported via `segmentation.rules.srxPath` (see `examples/default_rules.srx`).
- If no matching rule is found for the locale, a pragmatic built-in sentence splitter is used.
- The locale is derived from `sheet.sourceLocale` or `global.srcLang`.

## Placeholders and inline codes

- Configure `inlineCodeRegexes` per sheet to detect non-translatable tokens (e.g., `{0}`, `%s`).
- XLIFF export converts tokens to `<ph id="..."/>` with a per-segment placeholder map preserved in a `<note category="ph">` JSON payload for roundtrip.
- JSON export preserves a placeholder map under `unit.meta.placeholders` without altering source text.
- During merge, placeholder markers (e.g., `[[ph:ph1]]` in translated content) are rehydrated back into original tokens.

## XLIFF notes

If `global.exportComments` is true, XLIFF export includes extra `<note>` entries per unit:

- `category=header` — the header cell text for the source column (`headerRow`).
- `category=metadataRows` — a JSON object of metadata row values for this column.
- `category=comments` — cell notes/comments if `translateComments` is enabled.

These notes help maintain roundtrip context (sheet/row/col are always included as a base note).

## Style preservation

When `preserveStyles` is true:
- A minimal style snapshot (font name/size/bold/italic/color, alignment, fill color) is captured at extract time.
- During merge, the snapshot is reapplied to the target cell. If no snapshot exists, styles are copied from the source cell.

Rich text run-level formatting is not preserved in the MVP; this can be extended in future iterations.

## Config highlights

- Sheet selection via `namePattern`.
- `sourceColumns` and `targetColumns` (locale → column letter). Optional auto-create targets.
- Row/column filtering: `headerRow`, `valuesStartRow`, `skipHiddenRows`, `skipHiddenColumns`, `excludedRows/Columns`.
- Color exclusion via `excludeColors`.
- Formula handling via `extractFormulaResults`.
- Merged regions policy via `treatMergedRegions` (top-left | expand | skip).
- Comments via `translateComments`.
- Notes export via `global.exportComments`.
- Merge fallback via `global.mergeFallback` (default: `source`). When a segment lacks a `<target>`, choose to use its `<source>` or leave it empty (`empty`).

### Example CLI flows

- Extract per-locale XLIFFs, then merge all at once:

```bash
excel-l10n extract -c config.yml -i workbook.xlsx -o out.xlf --per-locale
# Edit out.fr.xlf, out.de.xlf ...
excel-l10n merge -c config.yml -i workbook.xlsx -t ./outdir -o workbook.merged.xlsx
```

- Extract a single bilingual XLIFF for French and merge only FR:

```bash
excel-l10n extract -c config.yml -i workbook.xlsx -o out.fr.xlf --target-lang fr
# Edit out.fr.xlf
excel-l10n merge -c config.yml -i workbook.xlsx -t out.fr.xlf -o workbook.fr.xlsx --target-lang fr
```

See `src/config/schema.json` for the full JSON Schema.

## Tests

- Unit tests cover config parsing, SRX segmentation, utilities and more.
- Integration testing can be added to validate end-to-end roundtrips (example scaffold included under `tests/`).

## Status

MVP implementation with SRX segmentation, placeholders, XLIFF notes, and style preservation. Further enhancements planned:
- Rich text run-level formatting support
- Streaming for very large workbooks
- Expanded Okapi option coverage

## Pseudo-translation

Generate fake translations to test UI expansion and encoding.

```bash
# XLSX → XLSX pseudo
excel-l10n pseudo -c config.yml -i workbook.xlsx -o pseudo.xlsx --target-lang fr

# XLIFF → XLIFF pseudo
excel-l10n extract -c config.yml -i workbook.xlsx -o out.xlf
excel-l10n pseudo -t out.xlf -o out.pseudo.xlf --expand 0.3 --wrap "⟦,⟧"
```

Behavior:
- Wrap text with markers (default ⟦ ⟧)
- Expand length by +30% (configurable)
- Replace characters with accented/uncommon variants
- Preserve placeholders: `{0}`, `%s`, `[[ph:ph1]]`

## Validate translations

Automatically check translated XLIFF/JSON for common issues.

```bash
excel-l10n validate -t translated/ --json --length-factor 2.5
```

Checks include:
- Missing targets
- Placeholder mismatches (`{0}`, `[[ph:ph1]]`)
- Length warnings (ratio > factor)
- ICU categories preserved (plural/select)

Exit code 0 = OK; 1 = findings.

## ICU handling (plural/select)

ICU plural/select blocks are protected during XLIFF export so structure is not accidentally broken. Inner texts are represented with placeholders to preserve logic while still surfacing translatable parts. Validation ensures ICU categories (e.g., one, other) are preserved between source and target.

Example:

```
{count, plural, one {1 file} other {# files}}
```

## Streaming mode (experimental)

For very large workbooks, you can enable streaming extraction.

```bash
excel-l10n extract -c config.yml -i huge.xlsx -o huge.xlf --stream
```

Note: the current release exposes the streaming flag and API; subsequent versions will wire a true streaming reader under the hood for constant-memory processing.
