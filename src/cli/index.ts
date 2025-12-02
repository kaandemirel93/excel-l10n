#!/usr/bin/env node
import { Command } from 'commander';
import fs from 'node:fs';
import path from 'node:path';
import { extract, exportUnitsToJson, exportUnitsToXliff, merge, parseConfig, parseTranslated, extractStream } from '../index.js';
import { exportToXliffStreamFromIterator } from '../exporter/xliff_stream.js';
import { extractStreamWorkbook } from '../io/stream.js';
import type { Config } from '../types.js';

const program = new Command();
program
  .name('excel-l10n')
  .description('Excel localization: extract → segment → export → merge')
  .version('0.1.0');

program.command('extract')
  .requiredOption('-i, --input <xlsx>', 'Input Excel file')
  .option('-c, --config <path>', 'Config file (json|yaml)')
  .requiredOption('-o, --output <path>', 'Output file (xlf|json)')
  .option('--format <fmt>', 'Output format: xlf|json', 'xlf')
  .option('--src-lang <lang>', 'Source language')
  .option('--target-lang <lang>', 'Target language for XLIFF; if omitted and multiple targets exist, use --per-locale')
  .option('--per-locale', 'For XLIFF, emit one file per target locale (suffix: .<locale>.xlf)', false)
  .option('--xliff-version <version>', 'XLIFF version: 1.2|2.1 (default: 2.1)', '2.1')
  .option('--stream', 'Experimental: stream extraction for very large files', false)
  // quick inline overrides / minimal config
  .option('--sheet <name>', 'Sheet name or regex')
  .option('--source <cols>', 'Source columns, e.g. A or A,B')
  .option('--target <map>', 'Target mapping locale=col[,locale=col], e.g. fr=B,de=C')
  .option('--verbose', 'Verbose logging', false)
  .action(async (opts) => {
    let cfg: Config;
    if (opts.config) {
      cfg = parseConfig(opts.config);
    } else {
      // Build minimal config from flags
      if (!opts.sheet || !opts.source || !opts.target) {
        console.error('Missing config and inline flags. Provide -c <config> or --sheet, --source, --target');
        process.exit(1);
      }
      const targetColumns: Record<string, string> = {};
      String(opts.target).split(',').forEach((pair: string) => {
        const [lc, col] = pair.split('=');
        if (lc && col) targetColumns[lc.trim()] = col.trim();
      });
      cfg = {
        global: { srcLang: opts.srcLang, overwrite: true, insertTargetPlacement: 'insertAfterSource', xliffVersion: (opts.xliffVersion === '1.2' ? '1.2' : '2.1') as '1.2' | '2.1' },
        segmentation: { enabled: true, rules: 'builtin' },
        workbook: {
          sheets: [
            {
              namePattern: String(opts.sheet),
              sourceColumns: String(opts.source).split(',').map((s: string) => s.trim()),
              targetColumns,
              createTargetIfMissing: true,
              headerRow: 1,
              valuesStartRow: 2,
              skipHiddenRows: true,
              skipHiddenColumns: true,
              extractFormulaResults: true,
              preserveStyles: true,
              treatMergedRegions: 'top-left' as const,
              inlineCodeRegexes: ['\\{\\d+\\}', '%s'],
              sourceLocale: opts.srcLang || 'en',
            },
          ],
        },
      };
    }
    // Apply CLI xliffVersion override if provided
    if (opts.xliffVersion && cfg.global) {
      cfg.global.xliffVersion = (opts.xliffVersion === '1.2' ? '1.2' : '2.1') as '1.2' | '2.1';
    } else if (opts.xliffVersion && !cfg.global) {
      cfg.global = { xliffVersion: (opts.xliffVersion === '1.2' ? '1.2' : '2.1') as '1.2' | '2.1' };
    }
    // Stream mode: write XLIFF progressively directly to file to avoid buffering
    if (opts.stream) {
      const fmt = (opts.format || 'xlf').toLowerCase();
      if (fmt !== 'xlf' && fmt !== 'xliff') {
        console.error('--stream currently supports XLIFF output only.');
        process.exit(1);
      }
      const iter = extractStreamWorkbook(opts.input, cfg);
      await exportToXliffStreamFromIterator(iter, cfg, { srcLang: opts.srcLang || cfg.global?.srcLang, generator: 'excel-l10n' }, opts.output);
      if (opts.verbose) console.log(`Wrote ${opts.output}`);
      return;
    }
    const units = await extract(opts.input, cfg);
    const fmt = (opts.format || 'xlf').toLowerCase();
    if (fmt === 'json') {
      const out = await exportUnitsToJson(units, cfg, { fileName: path.basename(opts.input), timestamp: new Date().toISOString() });
      fs.writeFileSync(opts.output, out, 'utf-8');
    } else {
      const targets = cfg.workbook.sheets.flatMap(s => Object.keys(s.targetColumns || {}));
      const uniqueTargets = Array.from(new Set(targets)).filter(Boolean);
      if (opts.perLocale && uniqueTargets.length > 0) {
        for (const lc of uniqueTargets) {
          const outFile = opts.output.replace(/\.xlf$/i, `.${lc}.xlf`);
          const out = await exportUnitsToXliff(units, cfg, { srcLang: (opts.srcLang || cfg.global?.srcLang), trgLang: lc, generator: 'excel-l10n' });
          fs.writeFileSync(outFile, out, 'utf-8');
          if (opts.verbose) console.log(`Wrote ${outFile}`);
        }
      } else {
        const out = await exportUnitsToXliff(units, cfg, { srcLang: opts.srcLang || cfg.global?.srcLang, trgLang: opts.targetLang, generator: 'excel-l10n' });
        fs.writeFileSync(opts.output, out, 'utf-8');
      }
    }
    if (opts.verbose) console.log(`Wrote ${opts.output}`);
  });

program.command('pseudo')
  .option('-c, --config <path>', 'Config file (json|yaml) for XLSX input')
  .option('-i, --input <xlsx>', 'Input Excel file (for XLSX->XLIFF or XLSX->XLSX pseudo)')
  .option('-t, --translated <xlf|json>', 'Input XLIFF/JSON file (for in-place pseudo localization)')
  .requiredOption('-o, --output <path>', 'Output file (.xlf or .xlsx)')
  .option('--expand <rate>', 'Expansion rate (e.g. 0.3 = +30%)', parseFloat, 0.3)
  .option('--wrap <chars>', 'Wrap markers as LEFT,RIGHT (default ⟦,⟧)', (v) => v, '⟦,⟧')
  .option('--target-lang <lang>', 'Target language column to write when output is .xlsx')
  .option('--verbose', 'Verbose logging', false)
  .action(async (opts) => {
    const [left, right] = String(opts.wrap || '⟦,⟧').split(',');
    const pseudoOpts = { expandRate: Number.isFinite(opts.expand) ? opts.expand : 0.3, wrap: { left, right } };
    if (opts.translated) {
      // Read XLIFF/JSON, pseudo targets, write same format
      const raw = fs.readFileSync(opts.translated, 'utf-8');
      const fmt = opts.translated.endsWith('.json') ? 'json' : 'xlf';
      const units = parseTranslated(raw, fmt as any);
      const pseudoed = (await import('../pseudo/index.js')).pseudoUnits(units, undefined, pseudoOpts);
      if (opts.output.endsWith('.json')) {
        const out = await exportUnitsToJson(pseudoed, undefined, { pseudo: true });
        fs.writeFileSync(opts.output, out, 'utf-8');
      } else {
        const out = await exportUnitsToXliff(pseudoed, { workbook: { sheets: [] } } as any, { srcLang: 'en', generator: 'excel-l10n' });
        fs.writeFileSync(opts.output, out, 'utf-8');
      }
      if (opts.verbose) console.log(`Wrote ${opts.output}`);
      return;
    }
    if (!opts.input || !opts.config) {
      console.error('For XLSX input, both --input and --config are required.');
      process.exit(1);
    }
    const cfg = parseConfig(opts.config);
    const units = await extract(opts.input, cfg);
    const pseudoed = (await import('../pseudo/index.js')).pseudoUnits(units, cfg, pseudoOpts);
    if (opts.output.endsWith('.xlsx')) {
      const targetLang = opts.targetLang || cfg.global?.targetLocale || Object.keys(cfg.workbook.sheets[0].targetColumns || {})[0];
      if (!targetLang) {
        console.error('No target language specified and none found in config. Use --target-lang or configure targetColumns.');
        process.exit(1);
      }
      (cfg as any).global = { ...(cfg.global || {}), targetLocale: targetLang };
      await merge(opts.input, opts.output, pseudoed, cfg);
      if (opts.verbose) console.log(`Wrote ${opts.output}`);
    } else {
      const out = await exportUnitsToXliff(pseudoed, cfg, { srcLang: cfg.global?.srcLang || 'en', generator: 'excel-l10n' });
      fs.writeFileSync(opts.output, out, 'utf-8');
      if (opts.verbose) console.log(`Wrote ${opts.output}`);
    }
  });

program.command('validate')
  .requiredOption('-t, --translated <paths>', 'Translated inputs: file, comma list, or directory')
  .option('--length-factor <x>', 'Warn if target > x * source length (default 2)', parseFloat, 2)
  .option('--format <fmt>', 'Format override: xlf|json', (v) => v, undefined)
  .option('--json', 'Output JSON report and exit 1 on findings', false)
  .option('--verbose', 'Verbose logging', false)
  .action(async (opts) => {
    const collectFiles = (p: string): string[] => {
      const abs = path.resolve(p);
      try {
        const st = fs.statSync(abs);
        if (st.isDirectory()) {
          return fs.readdirSync(abs)
            .filter(f => /(xlf|xliff|json)$/i.test(f))
            .map(f => path.join(abs, f));
        }
      } catch { /* ignore */ }
      return [abs];
    };
    const inputs: string[] = opts.translated.split(',').flatMap((p: string) => collectFiles(p.trim()));
    const { validateFiles } = await import('../validator/index.js');
    const report = await validateFiles(inputs, { lengthFactor: opts.length_factor || opts.lengthFactor || 2, formatOverride: opts.format });
    const hasFindings = report.items.some(it => it.level !== 'info');
    if (opts.json) {
      console.log(JSON.stringify(report, null, 2));
    } else {
      for (const it of report.items) {
        const icon = it.level === 'error' ? '❌' : it.level === 'warn' ? '⚠️' : 'ℹ️';
        console.log(`${icon} ${it.message}`);
      }
    }
    process.exit(hasFindings ? 1 : 0);
  });
program.command('merge')
  .requiredOption('-i, --input <xlsx>', 'Input Excel file')
  .requiredOption('-c, --config <path>', 'Config file (json|yaml)')
  .requiredOption('-t, --translated <path>', 'Translated input: file, comma-separated list, or directory (xlf/xliff/json)')
  .requiredOption('-o, --output <xlsx>', 'Output Excel file')
  .option('--format <fmt>', 'Format of translated file: xlf|json', (val) => val, undefined)
  .option('--dry-run', 'Validate config and show actions without writing', false)
  .option('--verbose', 'Verbose logging', false)
  .action(async (opts) => {
    const cfg = parseConfig(opts.config);
    if (opts.targetLang) {
      (cfg as any).global = { ...(cfg.global || {}), targetLocale: opts.targetLang };
    }
    const collectFiles = (p: string): string[] => {
      const abs = path.resolve(p);
      try {
        const st = fs.statSync(abs);
        if (st.isDirectory()) {
          return fs.readdirSync(abs)
            .filter(f => /\.(xlf|xliff|json)$/i.test(f))
            .map(f => path.join(abs, f));
        }
      } catch { /* ignore */ }
      return [abs];
    };

    const inputs: string[] = opts.translated.split(',').flatMap((p: string) => collectFiles(p.trim()));
    if (inputs.length === 0) {
      console.error('No translated inputs found.');
      process.exit(1);
    }

    const allUnits = inputs.flatMap((fp) => {
      const raw = fs.readFileSync(fp, 'utf-8');
      const fmt = opts.format || (fp.endsWith('.json') ? 'json' : 'xlf');
      if (opts.verbose) console.log(`Parsing ${fp} as ${fmt}`);
      return parseTranslated(raw, fmt === 'json' ? 'json' : 'xlf');
    });
    if (opts.verbose) console.log(`Parsed ${allUnits.length} units from ${inputs.length} input(s).`);
    if (opts.dry_run || opts['dry-run']) {
      console.log(`Would merge ${allUnits.length} units into ${opts.input} → ${opts.output}`);
      return;
    }
    await merge(opts.input, opts.output, allUnits, cfg);
    if (opts.verbose) console.log(`Wrote ${opts.output}`);
  });

program.parseAsync(process.argv);
