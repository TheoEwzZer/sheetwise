// Pure functions that render cells to HTML/CSS.
import type { CellStyle } from '../src/protocol.ts';

export const MAX_ROW = 1048576;
export const MAX_COL = 16384;

export const KEY = (r: number, c: number) => r * 16385 + c;

export function esc(s: string): string {
  return s.replace(/[&<>"']/g, (ch) =>
    ch === '&' ? '&amp;' : ch === '<' ? '&lt;' : ch === '>' ? '&gt;' : ch === '"' ? '&quot;' : '&#39;',
  );
}

export function colLetter(n: number): string {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function colNumber(s: string): number {
  let n = 0;
  for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

const FALLBACKS: Record<string, string> = {
  calibri: "Calibri, Carlito, 'Segoe UI', Arial, sans-serif",
  aptos: "Aptos, 'Segoe UI', Calibri, Arial, sans-serif",
  'aptos narrow': "'Aptos Narrow', 'Arial Narrow', 'Segoe UI', sans-serif",
  cambria: 'Cambria, Caladea, Georgia, serif',
  arial: "Arial, Helvetica, 'Liberation Sans', sans-serif",
  'times new roman': "'Times New Roman', Times, 'Liberation Serif', serif",
  'courier new': "'Courier New', Courier, monospace",
  consolas: "Consolas, 'Courier New', monospace",
};

// Note: single quotes only, since these values are inserted into style="…" attributes.
export function fontFamily(name: string | undefined): string | undefined {
  if (!name) return undefined;
  const fb = FALLBACKS[name.toLowerCase()];
  if (fb) return fb;
  return `'${name.replace(/['"<>&]/g, '')}', Calibri, 'Segoe UI', sans-serif`;
}

/** Typographic properties of a style (font, size, weight, color…). */
export function textCss(s: CellStyle | undefined): string {
  if (!s) return '';
  let css = '';
  const ff = fontFamily(s.ff);
  if (ff) css += `font-family:${ff};`;
  if (s.fs) css += `font-size:${s.fs}pt;`;
  if (s.b) css += 'font-weight:bold;';
  if (s.i) css += 'font-style:italic;';
  const deco: string[] = [];
  if (s.u) deco.push('underline');
  if (s.st) deco.push('line-through');
  if (deco.length) css += `text-decoration:${deco.join(' ')};${s.u === 'double' ? 'text-decoration-style:double;' : ''}`;
  if (s.fc) css += `color:${s.fc};`;
  return css;
}

export function cssFont(s: CellStyle | undefined, defaultFont: string): string {
  const size = s?.fs ?? 11;
  const fam = fontFamily(s?.ff ?? defaultFont) ?? 'sans-serif';
  return `${s?.i ? 'italic ' : ''}${s?.b ? 'bold ' : ''}${size}pt ${fam}`;
}

export function hAlign(s: CellStyle | undefined, kind: string): 'left' | 'center' | 'right' | 'justify' {
  const ha = s?.ha;
  switch (ha) {
    case 'left':
    case 'fill':
      return 'left';
    case 'center':
    case 'centerContinuous':
      return 'center';
    case 'right':
      return 'right';
    case 'justify':
    case 'distributed':
      return 'justify';
  }
  if (kind === 'n' || kind === 'd') return 'right';
  if (kind === 'b' || kind === 'e') return 'center';
  return 'left';
}

export function vAlign(s: CellStyle | undefined): 'top' | 'middle' | 'bottom' {
  switch (s?.va) {
    case 'top':
      return 'top';
    case 'middle':
    case 'center':
    case 'justify':
    case 'distributed':
      return 'middle';
    default:
      return 'bottom';
  }
}

/** Parses pasted TSV text (handles Excel-style quoting). */
export function parseTsv(text: string): string[][] {
  const t = text.replace(/\r\n?/g, '\n').replace(/\n$/, '');
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let i = 0;
  let quoted = false;
  let fieldStart = true;
  while (i < t.length) {
    const ch = t[i];
    if (quoted) {
      if (ch === '"') {
        if (t[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        quoted = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (fieldStart && ch === '"') {
      quoted = true;
      fieldStart = false;
      i++;
      continue;
    }
    fieldStart = false;
    if (ch === '\t') {
      row.push(field);
      field = '';
      fieldStart = true;
    } else if (ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      fieldStart = true;
    } else field += ch;
    i++;
  }
  row.push(field);
  rows.push(row);
  return rows;
}

export function tsvField(s: string): string {
  return /[\t\n\r"]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}
