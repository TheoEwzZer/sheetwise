// Generates ThirdPartyNotices.txt from the packages actually bundled into dist/extension.js.
import * as esbuild from 'esbuild';
import fs from 'node:fs';
import path from 'node:path';

const result = await esbuild.build({
  entryPoints: ['src/extension.ts'],
  bundle: true,
  platform: 'node',
  format: 'cjs',
  external: ['vscode'],
  write: false,
  metafile: true,
  logLevel: 'silent',
});

const pkgs = new Map();
for (const input of Object.keys(result.metafile.inputs)) {
  const norm = input.split(path.sep).join('/');
  const idx = norm.lastIndexOf('node_modules/');
  if (idx < 0) continue;
  const rest = norm.slice(idx + 'node_modules/'.length).split('/');
  const name = rest[0].startsWith('@') ? `${rest[0]}/${rest[1]}` : rest[0];
  pkgs.set(norm.slice(0, idx) + 'node_modules/' + name, name);
}

let out = 'Sheetwise bundles the following third-party software:\n\n';
for (const [dir, name] of [...pkgs].sort((a, b) => a[1].localeCompare(b[1]))) {
  const pj = JSON.parse(fs.readFileSync(path.join(dir, 'package.json'), 'utf8'));
  const license = typeof pj.license === 'string' ? pj.license : JSON.stringify(pj.license ?? pj.licenses ?? 'unknown');
  const file = fs.readdirSync(dir).find((f) => /^(licen[cs]e|copying)/i.test(f));
  out += `${'='.repeat(78)}\n${name}@${pj.version} — ${license}\n${'='.repeat(78)}\n`;
  out += file ? fs.readFileSync(path.join(dir, file), 'utf8').trim() + '\n\n' : `License: ${license}\n\n`;
}
fs.writeFileSync('ThirdPartyNotices.txt', out);
console.log(`${pkgs.size} packages: ${[...pkgs.values()].sort().join(', ')}`);
