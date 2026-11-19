// Number / date display according to the Excel format (via SSF) and the selected locale.
import SSF from 'ssf';

export type Locale = 'fr' | 'en';

export interface FmtCtx {
  locale: Locale;
  date1904: boolean;
}

const MS_PER_DAY = 86400000;

export function dateToSerial(d: Date, date1904 = false): number {
  return 25569 + d.getTime() / MS_PER_DAY - (date1904 ? 1462 : 0);
}

export function serialToDate(v: number, date1904 = false): Date {
  return new Date(Math.round((v - 25569 + (date1904 ? 1462 : 0)) * MS_PER_DAY));
}

/** "Short" formats that depend on the system language (notably ids 14 and 22). */
export function normalizeFormat(fmt: string | undefined, locale: Locale): string {
  if (!fmt) return 'General';
  const f = fmt.trim();
  if (f === 'mm-dd-yy' || f === 'm/d/yy' || f === 'm/d/yyyy') {
    return locale === 'fr' ? 'dd/mm/yyyy' : 'm/d/yyyy';
  }
  if (f === 'm/d/yy "h":mm' || f === 'm/d/yy h:mm') {
    return locale === 'fr' ? 'dd/mm/yyyy hh:mm' : 'm/d/yyyy h:mm';
  }
  return f;
}

export function shortDateFormat(locale: Locale): string {
  return locale === 'fr' ? 'dd/mm/yyyy' : 'm/d/yyyy';
}

const dateCache = new Map<string, boolean>();
export function isDateFormat(fmt: string | undefined): boolean {
  if (!fmt || fmt === 'General') return false;
  let r = dateCache.get(fmt);
  if (r === undefined) {
    try {
      r = SSF.is_date(fmt);
    } catch {
      r = false;
    }
    dateCache.set(fmt, r);
  }
  return r;
}

const MONTHS_EN = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const MONTHS_FR = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
const MONTHS_FR_SHORT = ['janv.', 'févr.', 'mars', 'avr.', 'mai', 'juin', 'juil.', 'août', 'sept.', 'oct.', 'nov.', 'déc.'];
const DAYS_EN = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const DAYS_FR = ['dimanche', 'lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi'];
const DAYS_FR_SHORT = ['dim.', 'lun.', 'mar.', 'mer.', 'jeu.', 'ven.', 'sam.'];

function frenchifyDate(s: string): string {
  let out = s;
  MONTHS_EN.forEach((m, i) => {
    out = out.replace(new RegExp(`\\b${m}\\b`, 'g'), MONTHS_FR[i]);
  });
  DAYS_EN.forEach((d, i) => {
    out = out.replace(new RegExp(`\\b${d}\\b`, 'g'), DAYS_FR[i]);
  });
  MONTHS_EN.forEach((m, i) => {
    out = out.replace(new RegExp(`\\b${m.slice(0, 3)}\\b`, 'g'), MONTHS_FR_SHORT[i]);
  });
  DAYS_EN.forEach((d, i) => {
    out = out.replace(new RegExp(`\\b${d.slice(0, 3)}\\b`, 'g'), DAYS_FR_SHORT[i]);
  });
  return out;
}

/** Replaces English number separators with French ones. */
function frenchifyNumber(s: string): string {
  return s.replace(/[.,]/g, (ch) => (ch === '.' ? ',' : ' '));
}

/** Splits a format into sections (positive;negative;zero;text), ignoring quotes and brackets. */
function splitSections(fmt: string): string[] {
  const out: string[] = [];
  let cur = '';
  let inQuote = false;
  let inBracket = false;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (ch === '\\' && !inQuote) {
      cur += ch + (fmt[i + 1] ?? '');
      i++;
      continue;
    }
    if (ch === '"') inQuote = !inQuote;
    else if (!inQuote && ch === '[') inBracket = true;
    else if (!inQuote && ch === ']') inBracket = false;
    if (ch === ';' && !inQuote && !inBracket) {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

const FMT_COLORS: Record<string, string> = {
  black: '#000000',
  blue: '#0000FF',
  cyan: '#00FFFF',
  green: '#00FF00',
  magenta: '#FF00FF',
  red: '#FF0000',
  white: '#FFFFFF',
  yellow: '#FFFF00',
};

/** Color imposed by the format for this value, e.g. "[Red]" in the negative section. */
export function formatColor(fmt: string, v: number): string | undefined {
  if (!fmt.includes('[')) return undefined;
  const sections = splitSections(fmt);
  let sec = sections[0];
  if (v < 0 && sections.length > 1) sec = sections[1];
  else if (v === 0 && sections.length > 2) sec = sections[2];
  const m = /\[(black|blue|cyan|green|magenta|red|white|yellow)\]/i.exec(sec);
  return m ? FMT_COLORS[m[1].toLowerCase()] : undefined;
}

export function formatNumber(v: number, rawFmt: string | undefined, ctx: FmtCtx): string {
  const fmt = normalizeFormat(rawFmt, ctx.locale);
  let s: string;
  try {
    s = SSF.format(fmt, v, { date1904: ctx.date1904 });
  } catch {
    s = String(v);
  }
  if (ctx.locale === 'fr') {
    s = isDateFormat(fmt) ? frenchifyDate(s) : frenchifyNumber(s);
  }
  return s;
}

/** Representation without floating-point artifacts (15 significant digits, like Excel). */
export function plainNumber(v: number, locale: Locale): string {
  if (!Number.isFinite(v)) return String(v);
  const s = String(Number(v.toPrecision(15)));
  return locale === 'fr' ? s.replace('.', ',') : s;
}

function pad(n: number): string {
  return String(n).padStart(2, '0');
}

export function dateEditText(serial: number, ctx: FmtCtx): string {
  const d = serialToDate(serial, ctx.date1904);
  const datePart =
    ctx.locale === 'fr'
      ? `${pad(d.getUTCDate())}/${pad(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`
      : `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`;
  const h = d.getUTCHours();
  const m = d.getUTCMinutes();
  const sec = d.getUTCSeconds();
  const hasTime = h || m || sec;
  const timePart = hasTime ? `${pad(h)}:${pad(m)}${sec ? ':' + pad(sec) : ''}` : '';
  if (serial < 1 && serial >= 0) return timePart || '00:00';
  return hasTime ? `${datePart} ${timePart}` : datePart;
}

/** Text shown to the user when editing a numeric cell. */
export function numberEditText(v: number, rawFmt: string | undefined, ctx: FmtCtx): string {
  const fmt = normalizeFormat(rawFmt, ctx.locale);
  if (isDateFormat(fmt)) return dateEditText(v, ctx);
  if (/%/.test(fmt.replace(/"[^"]*"/g, ''))) return plainNumber(v * 100, ctx.locale) + '%';
  return plainNumber(v, ctx.locale);
}

export function booleanText(b: boolean, locale: Locale): string {
  if (locale === 'fr') return b ? 'VRAI' : 'FAUX';
  return b ? 'TRUE' : 'FALSE';
}
