import path from 'node:path';
import { parseConfig } from '../src/config/index';

test('parses YAML config and applies defaults', () => {
  const cfgPath = path.resolve(__dirname, '../examples/config.yml');
  const cfg = parseConfig(cfgPath);
  expect(cfg.workbook.sheets.length).toBeGreaterThan(0);
  const s0 = cfg.workbook.sheets[0];
  expect(s0.createTargetIfMissing).toBe(true);
  expect(s0.treatMergedRegions).toBe('top-left');
});
