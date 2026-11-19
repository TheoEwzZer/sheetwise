// Parsing and rewriting of A1 references in Excel formulas.

export const MAX_ROW = 1048576;
export const MAX_COL = 16384;

export interface Ref {
  /** sheet name (without quotes) if the reference is qualified */
  sheet?: string;
  kind: 'cell' | 'range' | 'cols' | 'rows';
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  ar1: boolean;
  ac1: boolean;
  ar2: boolean;
  ac2: boolean;
}

export function colToNum(s: string): number {
  let n = 0;
  for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numToCol(n: number): string {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function address(r: number, c: number): string {
  return numToCol(c) + r;
}

const SHEET_PART = String.raw`(?:('(?:[^']|'')+'|[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*)!)?`;
const CELL = String.raw`(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})`;
const REF_RE = new RegExp(
  SHEET_PART +
    '(?:' +
    CELL +
    '(?::' +
    CELL +
    ')?' +
    String.raw`|(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})` +
    String.raw`|(\$?)(\d{1,7}):(\$?)(\d{1,7})` +
    ')',
  'y',
);

/** Characters that, immediately before, prevent a reference from starting. */
const BLOCK_BEFORE = /[A-Za-z0-9_.$!\]À-￿]/;
/** Characters that, immediately after, invalidate the reference (function, name…). */
const BLOCK_AFTER = /[A-Za-z0-9_(.!À-￿]/;

function unquoteSheet(raw: string): string {
  if (raw.startsWith("'")) return raw.slice(1, -1).replace(/''/g, "'");
  return raw;
}

export function quoteSheet(name: string): string {
  const simple = /^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/.test(name);
  const looksLikeRef = /^[A-Za-z]{1,3}\d+$/.test(name) || /^R\d*C\d*$/i.test(name);
  if (simple && !looksLikeRef) return name;
  return "'" + name.replace(/'/g, "''") + "'";
}

interface Match {
  start: number;
  end: number;
  sheetRaw?: string;
  ref: Ref;
}

function buildRef(m: RegExpExecArray): Ref | undefined {
  const sheet = m[1] !== undefined ? unquoteSheet(m[1]) : undefined;
  if (m[3] !== undefined) {
    const c1 = colToNum(m[3]);
    const r1 = +m[5];
    if (c1 > MAX_COL || r1 < 1 || r1 > MAX_ROW) return undefined;
    if (m[7] !== undefined) {
      const c2 = colToNum(m[7]);
      const r2 = +m[9];
      if (c2 > MAX_COL || r2 < 1 || r2 > MAX_ROW) return undefined;
      return {
        sheet, kind: 'range', r1, c1, r2, c2,
        ac1: m[2] === '$', ar1: m[4] === '$', ac2: m[6] === '$', ar2: m[8] === '$',
      };
    }
    return { sheet, kind: 'cell', r1, c1, r2: r1, c2: c1, ac1: m[2] === '$', ar1: m[4] === '$', ac2: m[2] === '$', ar2: m[4] === '$' };
  }
  if (m[11] !== undefined) {
    const c1 = colToNum(m[11]);
    const c2 = colToNum(m[13]);
    if (c1 > MAX_COL || c2 > MAX_COL) return undefined;
    return { sheet, kind: 'cols', r1: 1, c1, r2: MAX_ROW, c2, ac1: m[10] === '$', ac2: m[12] === '$', ar1: false, ar2: false };
  }
  if (m[15] !== undefined) {
    const r1 = +m[15];
    const r2 = +m[17];
    if (r1 < 1 || r2 < 1 || r1 > MAX_ROW || r2 > MAX_ROW) return undefined;
    return { sheet, kind: 'rows', r1, c1: 1, r2, c2: MAX_COL, ar1: m[14] === '$', ar2: m[16] === '$', ac1: false, ac2: false };
  }
  return undefined;
}

/** Walks the formula and calls `onRef` for each reference; `onText` receives the rest. */
function scan(formula: string, onText: (s: string) => void, onRef: (m: Match) => void): void {
  let i = 0;
  const n = formula.length;
  let text = '';
  const flush = () => {
    if (text) onText(text);
    text = '';
  };
  while (i < n) {
    const ch = formula[i];
    if (ch === '"') {
      // string literal
      let j = i + 1;
      while (j < n) {
        if (formula[j] === '"') {
          if (formula[j + 1] === '"') j += 2;
          else break;
        } else j++;
      }
      text += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (ch === '[') {
      // structured references / external workbook
      let depth = 0;
      let j = i;
      while (j < n) {
        if (formula[j] === '[') depth++;
        else if (formula[j] === ']') {
          depth--;
          if (depth === 0) break;
        }
        j++;
      }
      text += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    const prev = i > 0 ? formula[i - 1] : '';
    if (!prev || !BLOCK_BEFORE.test(prev)) {
      REF_RE.lastIndex = i;
      const m = REF_RE.exec(formula);
      if (m && m[0].length > 0) {
        const after = formula[i + m[0].length] ?? '';
        const ref = !BLOCK_AFTER.test(after) ? buildRef(m) : undefined;
        if (ref) {
          flush();
          onRef({ start: i, end: i + m[0].length, sheetRaw: m[1], ref });
          i += m[0].length;
          continue;
        }
      }
      if (ch === "'") {
        // quoted sheet name not followed by a recognized reference
        let j = i + 1;
        while (j < n) {
          if (formula[j] === "'") {
            if (formula[j + 1] === "'") j += 2;
            else break;
          } else j++;
        }
        text += formula.slice(i, j + 1);
        i = j + 1;
        continue;
      }
    }
    text += ch;
    i++;
  }
  flush();
}

export function findRefs(formula: string): Ref[] {
  const out: Ref[] = [];
  scan(formula, () => {}, (m) => out.push(m.ref));
  return out;
}

export function refToString(ref: Ref, sheetText?: string): string {
  const prefix = sheetText !== undefined ? sheetText + '!' : ref.sheet !== undefined ? quoteSheet(ref.sheet) + '!' : '';
  const d = (b: boolean) => (b ? '$' : '');
  switch (ref.kind) {
    case 'cell':
      return `${prefix}${d(ref.ac1)}${numToCol(ref.c1)}${d(ref.ar1)}${ref.r1}`;
    case 'range':
      return `${prefix}${d(ref.ac1)}${numToCol(ref.c1)}${d(ref.ar1)}${ref.r1}:${d(ref.ac2)}${numToCol(ref.c2)}${d(ref.ar2)}${ref.r2}`;
    case 'cols':
      return `${prefix}${d(ref.ac1)}${numToCol(ref.c1)}:${d(ref.ac2)}${numToCol(ref.c2)}`;
    case 'rows':
      return `${prefix}${d(ref.ar1)}${ref.r1}:${d(ref.ar2)}${ref.r2}`;
  }
}

/**
 * Rewrites each reference. `fn` returns: undefined (unchanged), null (#REF!) or a new Ref.
 */
export function transformRefs(formula: string, fn: (ref: Ref) => Ref | null | undefined): string {
  let out = '';
  let changed = false;
  scan(
    formula,
    (s) => {
      out += s;
    },
    (m) => {
      const res = fn({ ...m.ref });
      const original = formula.slice(m.start, m.end);
      if (res === undefined) {
        out += original;
      } else if (res === null) {
        changed = true;
        out += '#REF!';
      } else {
        changed = true;
        const sameSheet = res.sheet === m.ref.sheet;
        // keep the original sheet name spelling when it has not changed
        const sheetText = sameSheet && m.sheetRaw !== undefined ? m.sheetRaw : undefined;
        out += refToString(res, sheetText);
      }
    },
  );
  return changed ? out : formula;
}

export type Axis = 'row' | 'col';

/**
 * Shifts a reference for an insertion (count > 0) or deletion (count < 0)
 * of rows/columns starting at index `at`. Returns undefined if unchanged, null if destroyed.
 */
export function spliceRef(ref: Ref, axis: Axis, at: number, count: number): Ref | null | undefined {
  if (axis === 'row' && ref.kind === 'cols') return undefined;
  if (axis === 'col' && ref.kind === 'rows') return undefined;
  const k1 = axis === 'row' ? 'r1' : 'c1';
  const k2 = axis === 'row' ? 'r2' : 'c2';
  const max = axis === 'row' ? MAX_ROW : MAX_COL;
  let a = ref[k1];
  let b = ref[k2];
  if (count > 0) {
    if (a >= at) a += count;
    if (b >= at && b < max) b += count;
    a = Math.min(a, max);
    b = Math.min(b, max);
  } else {
    const n = -count;
    const end = at + n - 1;
    if (ref.kind === 'cell') {
      if (a >= at && a <= end) return null;
      if (a > end) a -= n;
      b = a;
    } else {
      if (a >= at && b <= end) return null;
      const na = a > end ? a - n : a >= at ? at : a;
      const nb = b >= max ? b : b > end ? b - n : b >= at ? at - 1 : b;
      a = na;
      b = nb;
    }
  }
  if (a === ref[k1] && b === ref[k2]) return undefined;
  return { ...ref, [k1]: a, [k2]: b };
}

/** Shifts the relative parts of a reference (copy/paste of a formula). */
export function offsetRef(ref: Ref, dr: number, dc: number): Ref | null | undefined {
  if (!dr && !dc) return undefined;
  const out = { ...ref };
  if (ref.kind !== 'cols') {
    if (!ref.ar1) out.r1 += dr;
    if (!ref.ar2) out.r2 += dr;
  }
  if (ref.kind !== 'rows') {
    if (!ref.ac1) out.c1 += dc;
    if (!ref.ac2) out.c2 += dc;
  }
  if (ref.kind === 'cell') {
    out.r2 = out.r1;
    out.c2 = out.c1;
  }
  if (out.r1 < 1 || out.c1 < 1 || out.r2 < 1 || out.c2 < 1 || out.r1 > MAX_ROW || out.r2 > MAX_ROW || out.c1 > MAX_COL || out.c2 > MAX_COL) {
    return null;
  }
  return out;
}

/** Parses "A1", "$A$1:B5", "Sheet1!A1:B2" (bare, without a surrounding formula). */
export function parseRangeText(text: string): (Ref & { sheet?: string }) | undefined {
  const refs: Ref[] = [];
  let other = '';
  scan(text.trim(), (s) => (other += s), (m) => refs.push(m.ref));
  if (refs.length !== 1 || other.trim()) return undefined;
  return refs[0];
}

export function rangeToA1(r1: number, c1: number, r2: number, c2: number): string {
  const a = address(r1, c1);
  return r1 === r2 && c1 === c2 ? a : `${a}:${address(r2, c2)}`;
}
