import fs from 'node:fs';
import path from 'node:path';
import YAML from 'yaml';
import AjvPkg from 'ajv';
import { Config } from '../types.js';

// Load JSON schema in a way that works for both Jest (CJS) and runtime (ESM)
let schema: any;
{
  const candidates = [
    // running from project root (ts-jest)
    path.resolve(process.cwd(), 'src/config/schema.json'),
    // running from built dist in project
    path.resolve(process.cwd(), 'dist/config/schema.json'),
    // relative to cwd/config
    path.resolve(process.cwd(), 'config/schema.json'),
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) {
      schema = JSON.parse(fs.readFileSync(p, 'utf-8'));
      break;
    }
  }
  if (!schema) {
    // final fallback: try sibling path assuming Node resolves this file from src/config
    try {
      const fallback = path.resolve('src/config/schema.json');
      if (fs.existsSync(fallback)) schema = JSON.parse(fs.readFileSync(fallback, 'utf-8'));
    } catch {
      /* ignore */
    }
  }
}

// Normalize Ajv constructor across CJS/ESM
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const Ajv: any = (AjvPkg as any).default || (AjvPkg as any);
const ajv = new Ajv({ allErrors: true, useDefaults: true, strict: false });
const validate = ajv.compile(schema as any);

export function parseConfig(configPathOrObject: string | Config): Config {
  let cfg: any;
  if (typeof configPathOrObject === 'string') {
    const abs = path.resolve(configPathOrObject);
    const content = fs.readFileSync(abs, 'utf-8');
    if (abs.endsWith('.yml') || abs.endsWith('.yaml')) {
      cfg = YAML.parse(content);
    } else {
      cfg = JSON.parse(content);
    }
  } else {
    cfg = configPathOrObject;
  }

  const valid = validate(cfg);
  if (!valid) {
    const errs = (validate.errors || []).map((e: any) => `${e.instancePath} ${e.message}`).join('\n');
    throw new Error(`Invalid config:\n${errs}`);
  }
  return cfg as Config;
}
