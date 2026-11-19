// Resolves Excel colors (ARGB, theme + tint, indexed palette) to CSS.

export interface ThemeInfo {
  /** Colors in Excel theme index order: lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink */
  colors: string[];
  minorFont: string;
}

export interface ExcelColor {
  argb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
}

// Legacy indexed palette (BIFF8 / SpreadsheetML).
const INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

// Default Office 2013-2022 theme, in Excel index order.
const DEFAULT_THEME_COLORS = [
  'FFFFFF', '000000', 'E7E6E6', '44546A',
  '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47',
  '0563C1', '954F72',
];

export const DEFAULT_THEME: ThemeInfo = { colors: DEFAULT_THEME_COLORS, minorFont: 'Calibri' };

/** Extracts the colors and minor font from a theme1.xml. */
export function parseTheme(xml: string | undefined): ThemeInfo {
  if (!xml) return DEFAULT_THEME;
  const scheme = /<a:clrScheme[^>]*>([\s\S]*?)<\/a:clrScheme>/.exec(xml)?.[1];
  const read = (tag: string, fallback: string): string => {
    if (!scheme) return fallback;
    const block = new RegExp(`<a:${tag}>([\\s\\S]*?)</a:${tag}>`).exec(scheme)?.[1];
    if (!block) return fallback;
    const srgb = /<a:srgbClr[^>]*val="([0-9A-Fa-f]{6})"/.exec(block)?.[1];
    if (srgb) return srgb.toUpperCase();
    const sys = /<a:sysClr[^>]*lastClr="([0-9A-Fa-f]{6})"/.exec(block)?.[1];
    if (sys) return sys.toUpperCase();
    return fallback;
  };
  // The XML order is dk1, lt1, dk2, lt2 but Excel indexes lt1=0, dk1=1, lt2=2, dk2=3.
  const colors = [
    read('lt1', 'FFFFFF'),
    read('dk1', '000000'),
    read('lt2', 'E7E6E6'),
    read('dk2', '44546A'),
    read('accent1', '4472C4'),
    read('accent2', 'ED7D31'),
    read('accent3', 'A5A5A5'),
    read('accent4', 'FFC000'),
    read('accent5', '5B9BD5'),
    read('accent6', '70AD47'),
    read('hlink', '0563C1'),
    read('folHlink', '954F72'),
  ];
  const minorFont =
    /<a:minorFont>\s*<a:latin[^>]*typeface="([^"]+)"/.exec(xml)?.[1] ?? DEFAULT_THEME.minorFont;
  return { colors, minorFont };
}

function hexToRgb(hex: string): [number, number, number] {
  const n = parseInt(hex, 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

function rgbToHex(r: number, g: number, b: number): string {
  const h = (x: number) => Math.max(0, Math.min(255, Math.round(x))).toString(16).padStart(2, '0');
  return (h(r) + h(g) + h(b)).toUpperCase();
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hue(p, q, h + 1 / 3) * 255, hue(p, q, h) * 255, hue(p, q, h - 1 / 3) * 255];
}

/** Applies an Excel tint (-1..1) to the luminance. */
export function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const [h, s, l] = rgbToHsl(...hexToRgb(hex));
  const nl = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  return rgbToHex(...hslToRgb(h, s, Math.max(0, Math.min(1, nl))));
}

/**
 * Converts an Excel color to "#RRGGBB".
 * Returns undefined for the automatic color (default text).
 */
export function resolveColor(color: ExcelColor | undefined, theme: ThemeInfo): string | undefined {
  if (!color) return undefined;
  let hex: string | undefined;
  if (typeof color.argb === 'string') {
    const a = color.argb.replace('#', '');
    hex = a.length === 8 ? a.slice(2) : a.length === 6 ? a : undefined;
  } else if (typeof color.theme === 'number') {
    hex = theme.colors[color.theme] ?? DEFAULT_THEME_COLORS[color.theme];
  } else if (typeof color.indexed === 'number') {
    if (color.indexed >= 64) return undefined; // system colors (auto)
    hex = INDEXED[color.indexed];
  }
  if (!hex || !/^[0-9A-Fa-f]{6}$/.test(hex)) return undefined;
  if (typeof color.tint === 'number') hex = applyTint(hex, color.tint);
  return '#' + hex.toUpperCase();
}

/** True if the color is the "automatic" text color (untinted theme black). */
export function isAutoTextColor(color: ExcelColor | undefined): boolean {
  if (!color) return true;
  if (typeof color.indexed === 'number' && color.indexed >= 64) return true;
  if (color.theme === 1 && !color.tint) return true;
  return false;
}

/** "#RRGGBB" -> "FFRRGGBB" for ExcelJS. */
export function cssToArgb(css: string): string {
  return 'FF' + css.replace('#', '').toUpperCase();
}
