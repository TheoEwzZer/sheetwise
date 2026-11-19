// Interprets text typed into a cell (the way Excel does).
import { type Locale, shortDateFormat, isDateFormat, normalizeFormat } from './numfmt.ts';

export type InputValue =
  | null
  | number
  | string
  | boolean
  | Date
  | { formula: string }
  | { error: string };

export interface ParsedInput {
  value: InputValue;
  /** Format to apply if the cell uses the General format. */
  numFmt?: string;
}

/** Function names from the French version of Excel -> names stored in the file. */
const FR_FUNCTIONS: Record<string, string> = {
  SOMME: 'SUM', MOYENNE: 'AVERAGE', SI: 'IF', 'SI.CONDITIONS': 'IFS', SIERREUR: 'IFERROR', 'SI.NON.DISP': 'IFNA',
  ET: 'AND', OU: 'OR', NON: 'NOT', 'OUX': 'XOR', NB: 'COUNT', NBVAL: 'COUNTA', 'NB.VIDE': 'COUNTBLANK',
  'NB.SI': 'COUNTIF', 'NB.SI.ENS': 'COUNTIFS', 'SOMME.SI': 'SUMIF', 'SOMME.SI.ENS': 'SUMIFS',
  'MOYENNE.SI': 'AVERAGEIF', 'MOYENNE.SI.ENS': 'AVERAGEIFS', 'MAX.SI.ENS': 'MAXIFS', 'MIN.SI.ENS': 'MINIFS',
  RECHERCHEV: 'VLOOKUP', RECHERCHEH: 'HLOOKUP', RECHERCHEX: 'XLOOKUP', RECHERCHE: 'LOOKUP', EQUIV: 'MATCH',
  ARRONDI: 'ROUND', 'ARRONDI.SUP': 'ROUNDUP', 'ARRONDI.INF': 'ROUNDDOWN', ENT: 'INT', TRONQUE: 'TRUNC',
  ABS: 'ABS', MOD: 'MOD', PUISSANCE: 'POWER', RACINE: 'SQRT', PRODUIT: 'PRODUCT', 'SOMMEPROD': 'SUMPRODUCT',
  MAX: 'MAX', MIN: 'MIN', MEDIANE: 'MEDIAN', 'ECARTYPE': 'STDEV', 'ECARTYPE.STANDARD': 'STDEV.S', RANG: 'RANK',
  CONCATENER: 'CONCATENATE', CONCAT: 'CONCAT', 'JOINDRE.TEXTE': 'TEXTJOIN', GAUCHE: 'LEFT', DROITE: 'RIGHT',
  STXT: 'MID', NBCAR: 'LEN', MAJUSCULE: 'UPPER', MINUSCULE: 'LOWER', NOMPROPRE: 'PROPER', SUPPRESPACE: 'TRIM',
  SUBSTITUE: 'SUBSTITUTE', REMPLACER: 'REPLACE', CHERCHE: 'SEARCH', TROUVE: 'FIND', TEXTE: 'TEXT', CNUM: 'VALUE',
  REPT: 'REPT', EXACT: 'EXACT', AUJOURDHUI: 'TODAY', MAINTENANT: 'NOW', DATE: 'DATE', ANNEE: 'YEAR', MOIS: 'MONTH',
  JOUR: 'DAY', HEURE: 'HOUR', MINUTE: 'MINUTE', SECONDE: 'SECOND', JOURSEM: 'WEEKDAY', 'NO.SEMAINE': 'WEEKNUM',
  'FIN.MOIS': 'EOMONTH', MOIS_DECALER: 'EDATE', 'NB.JOURS.OUVRES': 'NETWORKDAYS', 'SERIE.JOUR.OUVRE': 'WORKDAY',
  DATEDIF: 'DATEDIF', JOURS: 'DAYS', ESTVIDE: 'ISBLANK', ESTNUM: 'ISNUMBER', ESTTEXTE: 'ISTEXT',
  ESTERREUR: 'ISERROR', ESTERR: 'ISERR', ESTNA: 'ISNA', INDEX: 'INDEX', DECALER: 'OFFSET', INDIRECT: 'INDIRECT',
  LIGNE: 'ROW', COLONNE: 'COLUMN', LIGNES: 'ROWS', COLONNES: 'COLUMNS', CHOISIR: 'CHOOSE', 'SI.MULTIPLE': 'SWITCH',
  ALEA: 'RAND', 'ALEA.ENTRE.BORNES': 'RANDBETWEEN', VRAI: 'TRUE', FAUX: 'FALSE', PI: 'PI', 'LN': 'LN', EXP: 'EXP',
  'SOUS.TOTAL': 'SUBTOTAL', 'GRANDE.VALEUR': 'LARGE', 'PETITE.VALEUR': 'SMALL', 'NA': 'NA',
};

/** Replaces French function names (followed by a parenthesis, outside strings). */
export function translateFunctionNames(f: string): string {
  let out = '';
  let i = 0;
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"' || ch === "'") {
      const end = f.indexOf(ch, i + 1);
      const stop = end < 0 ? f.length : end + 1;
      out += f.slice(i, stop);
      i = stop;
      continue;
    }
    const m = /^[A-Za-zÀ-ÿ_][A-Za-zÀ-ÿ0-9_.]*/.exec(f.slice(i));
    if (m && !/[A-Za-z0-9_.$]/.test(f[i - 1] ?? '')) {
      const name = m[0];
      const next = f[i + name.length];
      const en = next === '(' ? FR_FUNCTIONS[name.toUpperCase()] : undefined;
      out += en ?? name;
      i += name.length;
      continue;
    }
    out += ch;
    i++;
  }
  return out;
}

const ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'];

/**
 * Converts a formula typed French-style (";" separator and decimal comma)
 * into stored Excel syntax ("," separator and decimal point).
 */
export function normalizeFormulaInput(input: string): string {
  const f = translateFunctionNames(input);
  // Detect whether there are ";" outside strings
  let hasSemicolon = false;
  let inStr = false;
  for (const ch of f) {
    if (ch === '"') inStr = !inStr;
    else if (!inStr && ch === ';') hasSemicolon = true;
  }
  if (!hasSemicolon) return f;
  let out = '';
  inStr = false;
  let inQuote = false; // sheet name in single quotes
  for (let i = 0; i < f.length; i++) {
    const ch = f[i];
    if (ch === '"' && !inQuote) inStr = !inStr;
    else if (ch === "'" && !inStr) inQuote = !inQuote;
    if (!inStr && !inQuote) {
      if (ch === ';') {
        out += ',';
        continue;
      }
      if (ch === ',' && /\d/.test(f[i - 1] ?? '') && /\d/.test(f[i + 1] ?? '')) {
        out += '.';
        continue;
      }
    }
    out += ch;
  }
  return out;
}

function parseNumber(s: string, locale: Locale): number | undefined {
  const t = s.replace(/[  ]/g, ' ').trim();
  if (!t) return undefined;
  let normalized: string | undefined;
  if (locale === 'fr') {
    if (/^[+-]?(\d{1,3}( \d{3})+|\d+)(,\d*)?(e[+-]?\d+)?$/i.test(t)) {
      normalized = t.replace(/ /g, '').replace(',', '.');
    } else if (/^[+-]?\d*\.\d+(e[+-]?\d+)?$/i.test(t) || /^[+-]?\d+\.?\d*e[+-]?\d+$/i.test(t)) {
      normalized = t; // decimal point accepted too
    } else if (/^[+-]?,\d+$/.test(t)) {
      normalized = t.replace(',', '.');
    }
  } else if (/^[+-]?(\d{1,3}(,\d{3})+|\d+)?(\.\d*)?(e[+-]?\d+)?$/i.test(t) && /\d/.test(t)) {
    normalized = t.replace(/,/g, '');
  }
  if (normalized === undefined) return undefined;
  const n = Number(normalized);
  return Number.isFinite(n) ? n : undefined;
}

function decimalsOf(s: string): number {
  const m = /[.,](\d+)\s*%?\s*$/.exec(s.trim());
  return m ? m[1].length : 0;
}

function twoDigitYear(y: number): number {
  return y < 30 ? 2000 + y : 1900 + y;
}

function validDate(y: number, m: number, d: number): boolean {
  if (m < 1 || m > 12 || d < 1 || d > 31) return false;
  const dt = new Date(Date.UTC(y, m - 1, d));
  return dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
}

function parseTime(s: string): [number, number, number] | undefined {
  const m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s.trim());
  if (!m) return undefined;
  const h = +m[1];
  const mi = +m[2];
  const se = m[3] ? +m[3] : 0;
  if (h > 23 || mi > 59 || se > 59) return undefined;
  return [h, mi, se];
}

function parseDate(s: string, locale: Locale): ParsedInput | undefined {
  const t = s.trim();
  const time = parseTime(t);
  if (time) {
    return {
      value: new Date(Date.UTC(1899, 11, 30, time[0], time[1], time[2])),
      numFmt: time[2] ? 'hh:mm:ss' : 'hh:mm',
    };
  }
  const m =
    /^(\d{1,4})[/.-](\d{1,2})(?:[/.-](\d{1,4}))?(?:\s+(\d{1,2}:\d{2}(?::\d{2})?))?$/.exec(t);
  if (!m) return undefined;
  let y: number;
  let mo: number;
  let d: number;
  const a = +m[1];
  const b = +m[2];
  const c = m[3] !== undefined ? +m[3] : undefined;
  if (m[1].length === 4) {
    // ISO yyyy-mm-dd
    if (c === undefined) return undefined;
    y = a;
    mo = b;
    d = c;
  } else {
    if (locale === 'fr') {
      d = a;
      mo = b;
    } else {
      mo = a;
      d = b;
    }
    if (c === undefined) y = new Date().getFullYear();
    else y = m[3].length <= 2 ? twoDigitYear(c) : c;
  }
  if (!validDate(y, mo, d)) return undefined;
  let hh = 0;
  let mi = 0;
  let se = 0;
  if (m[4]) {
    const tt = parseTime(m[4]);
    if (!tt) return undefined;
    [hh, mi, se] = tt;
  }
  const base = shortDateFormat(locale);
  return {
    value: new Date(Date.UTC(y, mo - 1, d, hh, mi, se)),
    numFmt: m[4] ? `${base} hh:mm` : base,
  };
}

export function parseInput(text: string, opts: { locale: Locale; cellNumFmt?: string }): ParsedInput {
  const { locale } = opts;
  const fmt = normalizeFormat(opts.cellNumFmt, locale);
  if (text === '') return { value: null };
  if (fmt === '@') return { value: text };
  if (text.startsWith("'")) return { value: text.slice(1) };
  if (text.startsWith('=') && text.length > 1) {
    return { value: { formula: normalizeFormulaInput(text.slice(1)) } };
  }
  const t = text.trim();
  if (!t) return { value: text };

  const upper = t.toUpperCase();
  if (upper === 'TRUE' || upper === 'VRAI') return { value: true };
  if (upper === 'FALSE' || upper === 'FAUX') return { value: false };
  if (ERRORS.includes(upper)) return { value: { error: upper } };

  // Percentage
  if (t.endsWith('%')) {
    const n = parseNumber(t.slice(0, -1), locale);
    if (n !== undefined) {
      const dec = decimalsOf(t);
      return { value: n / 100, numFmt: dec ? '0.' + '0'.repeat(dec) + '%' : '0%' };
    }
  }

  // Currency
  const cur = /^(€|\$)?\s*([^€$]+?)\s*(€|\$)?$/.exec(t);
  if (cur && (cur[1] || cur[3]) && !(cur[1] && cur[3])) {
    const n = parseNumber(cur[2], locale);
    if (n !== undefined) {
      const sym = cur[1] || cur[3];
      return {
        value: n,
        numFmt: sym === '€' ? '#,##0.00 "€"' : '"$"#,##0.00',
      };
    }
  }

  const n = parseNumber(t, locale);
  if (n !== undefined) {
    if (/%/.test(fmt.replace(/"[^"]*"/g, '')) && !isDateFormat(fmt)) return { value: n / 100 };
    return { value: n };
  }

  const date = parseDate(t, locale);
  if (date) return date;

  return { value: text };
}
