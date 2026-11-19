import * as esbuild from 'esbuild';

const production = process.argv.includes('--production');
const watch = process.argv.includes('--watch');

/** @type {import('esbuild').BuildOptions[]} */
const configs = [
  {
    entryPoints: ['src/extension.ts'],
    bundle: true,
    platform: 'node',
    format: 'cjs',
    target: 'node18',
    outfile: 'dist/extension.js',
    external: ['vscode'],
    sourcemap: !production,
    minify: production,
    logLevel: 'warning',
  },
  {
    entryPoints: ['webview/main.ts'],
    bundle: true,
    platform: 'browser',
    format: 'iife',
    target: 'es2022',
    outfile: 'media/webview.js',
    sourcemap: production ? false : 'inline',
    minify: production,
    logLevel: 'warning',
  },
];

if (watch) {
  for (const c of configs) {
    const ctx = await esbuild.context(c);
    await ctx.watch();
  }
  console.log('esbuild: watching for changes…');
} else {
  await Promise.all(configs.map((c) => esbuild.build(c)));
  console.log('esbuild: build complete' + (production ? ' (production)' : ''));
}
