import type {
  CellModel,
  CellRange,
  CellStyle,
  FormatPatch,
  SheetModel,
  ToHost,
  ToWebview,
  WorkbookMeta,
  Settings,
  StructureOp,
  BorderSide,
} from '../src/protocol.ts';
import {
  KEY,
  MAX_COL,
  MAX_ROW,
  colLetter,
  colNumber,
  cssFont,
  esc,
  fontFamily,
  hAlign,
  parseTsv,
  textCss,
  tsvField,
  vAlign,
} from './render.ts';

declare function acquireVsCodeApi(): {
  postMessage(m: unknown): void;
  getState(): unknown;
  setState(s: unknown): void;
};

const vscode = acquireVsCodeApi();
const post = (m: ToHost) => vscode.postMessage(m);

// =====================================================================================
// Skeleton
// =====================================================================================

const ICON = {
  left: '<svg viewBox="0 0 16 16"><path d="M2 3h12v1.2H2zM2 6h8v1.2H2zM2 9h12v1.2H2zM2 12h8v1.2H2z"/></svg>',
  center: '<svg viewBox="0 0 16 16"><path d="M2 3h12v1.2H2zM4 6h8v1.2H4zM2 9h12v1.2H2zM4 12h8v1.2H4z"/></svg>',
  right: '<svg viewBox="0 0 16 16"><path d="M2 3h12v1.2H2zM6 6h8v1.2H6zM2 9h12v1.2H2zM6 12h8v1.2H6z"/></svg>',
  top: '<svg viewBox="0 0 16 16"><path d="M2 2h12v1.2H2zM7.4 14V6.3L5 8.7l-.85-.85L8 4l3.85 3.85-.85.85-2.4-2.4V14z"/></svg>',
  middle: '<svg viewBox="0 0 16 16"><path d="M2 7.4h12v1.2H2zM7.4 1h1.2v4l1.6-1.6.8.8L8 7.2 5 4.2l.8-.8 1.6 1.6zM7.4 15h1.2v-4l1.6 1.6.8-.8L8 8.8 5 11.8l.8.8 1.6-1.6z"/></svg>',
  bottom: '<svg viewBox="0 0 16 16"><path d="M2 12.8h12V14H2zM7.4 2v7.7L5 7.3l-.85.85L8 12l3.85-3.85L11 7.3 8.6 9.7V2z"/></svg>',
  wrap: '<svg viewBox="0 0 16 16"><path d="M2 3h12v1.2H2zM2 12h5v1.2H2zM2 7.5h9.5a2.5 2.5 0 0 1 0 5H10v1.6L7.6 12 10 10v1.3h1.5a1.3 1.3 0 0 0 0-2.6H2z"/></svg>',
  merge: '<svg viewBox="0 0 16 16"><path d="M1.5 2.5h13v11h-13zm1.2 1.2v8.6h10.6V3.7zM4 8l2-2v1.4h4V6l2 2-2 2V8.6H6V10z"/></svg>',
  clear: '<svg viewBox="0 0 16 16"><path d="M9.6 2.2 14 6.6 8.3 12.3H13v1.2H4.9L2 10.6zM4.1 10.3l2 2h.6l2.4-2.4L5.8 6.6z"/></svg>',
  grid: '<svg viewBox="0 0 16 16"><path d="M2 2h12v12H2zm1.2 1.2v3.2h3.2V3.2zm4.4 0v3.2h3.2V3.2zm4.4 0v3.2h1.8V3.2zM3.2 7.6v3.2h3.2V7.6zm4.4 0v3.2h3.2V7.6zm4.4 0v3.2h1.8V7.6zM3.2 12v.8h3.2V12zm4.4 0v.8h3.2V12zm4.4 0v.8h1.8V12z"/></svg>',
};

const FONTS = [
  'Aptos', 'Aptos Narrow', 'Arial', 'Arial Narrow', 'Calibri', 'Cambria', 'Candara', 'Consolas',
  'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana',
];
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
const NUMFMTS: [string, string][] = [
  ['General', 'General'],
  ['0', 'Number (0)'],
  ['#,##0.00', 'Number (1,234.56)'],
  ['#,##0.00 "€"', 'Currency (€)'],
  ['0%', 'Percentage'],
  ['0.00%', 'Percentage (0.00%)'],
  ['dd/mm/yyyy', 'Short date'],
  ['dddd d mmmm yyyy', 'Long date'],
  ['hh:mm', 'Time'],
  ['dd/mm/yyyy hh:mm', 'Date and time'],
  ['0.00E+00', 'Scientific'],
  ['@', 'Text'],
];

const app = document.getElementById('app')!;
app.innerHTML = `
<div id="toolbar" role="toolbar">
  <select id="tb-font" title="Font">${FONTS.map((f) => `<option>${f}</option>`).join('')}</select>
  <select id="tb-size" title="Font size">${SIZES.map((s) => `<option value="${s}">${s}</option>`).join('')}</select>
  <span class="sep"></span>
  <button data-cmd="bold" title="Bold (Ctrl+B)"><b>B</b></button>
  <button data-cmd="italic" title="Italic (Ctrl+I)"><i>I</i></button>
  <button data-cmd="underline" title="Underline (Ctrl+U)"><u>U</u></button>
  <button data-cmd="strike" title="Strikethrough (Ctrl+5)"><s>ab</s></button>
  <span class="sep"></span>
  <label class="color" title="Font color"><span class="glyph">A</span><span class="bar" id="fc-bar"></span><input type="color" id="tb-fc" value="#c00000"></label>
  <label class="color" title="Fill color"><span class="glyph fill">&#9638;</span><span class="bar" id="bg-bar"></span><input type="color" id="tb-bg" value="#ffff00"></label>
  <button data-cmd="nofill" title="No fill">&#8709;</button>
  <span class="sep"></span>
  <button data-cmd="left" title="Align left">${ICON.left}</button>
  <button data-cmd="center" title="Center">${ICON.center}</button>
  <button data-cmd="right" title="Align right">${ICON.right}</button>
  <button data-cmd="top" title="Top align">${ICON.top}</button>
  <button data-cmd="middle" title="Middle align">${ICON.middle}</button>
  <button data-cmd="bottom" title="Bottom align">${ICON.bottom}</button>
  <button data-cmd="wrap" title="Wrap text">${ICON.wrap}</button>
  <button data-cmd="merge" title="Merge / unmerge cells">${ICON.merge}</button>
  <span class="sep"></span>
  <select id="tb-border" title="Borders">
    <option value="">Borders…</option>
    <option value="all">All borders</option>
    <option value="outer">Outside borders</option>
    <option value="thickOuter">Thick outside borders</option>
    <option value="bottom">Bottom border</option>
    <option value="top">Top border</option>
    <option value="left">Left border</option>
    <option value="right">Right border</option>
    <option value="none">No border</option>
  </select>
  <select id="tb-numfmt" title="Number format">${NUMFMTS.map(([v, l]) => `<option value="${esc(v)}">${l}</option>`).join('')}<option value="__custom" disabled>Custom</option></select>
  <button data-cmd="clearFormat" title="Clear formatting">${ICON.clear}</button>
  <span class="spacer"></span>
  <button data-cmd="grid" title="Show / hide gridlines">${ICON.grid}</button>
  <select id="tb-zoom" title="Zoom">${[50, 75, 90, 100, 110, 125, 150, 175, 200].map((z) => `<option value="${z}">${z}%</option>`).join('')}</select>
</div>
<div id="fbar">
  <input id="namebox" spellcheck="false" title="Name Box: type a reference (e.g. B12 or A1:C5)">
  <span class="fx">fx</span>
  <textarea id="finput" rows="1" spellcheck="false"></textarea>
</div>
<div id="banner" hidden></div>
<div id="grid-scroll">
  <div id="sheet">
    <table id="grid"><colgroup></colgroup><thead></thead><tbody></tbody></table>
    <div id="sel-range" class="overlay"></div>
    <div id="sel-active" class="overlay"></div>
    <div id="copy-ants" class="overlay" hidden></div>
    <button id="dv-btn" class="overlay" hidden tabindex="-1" title="Pick from list (Alt+Down)">&#9662;</button>
    <textarea id="editor" spellcheck="false" autocomplete="off"></textarea>
  </div>
</div>
<div id="bottom"><div id="tabs"></div><div id="status"></div></div>
<div id="menu" hidden></div>
<div id="toast" hidden></div>
<div id="loading">Loading workbook…</div>
`;

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const scroller = $<HTMLDivElement>('grid-scroll');
const sheetEl = $<HTMLDivElement>('sheet');
const table = $<HTMLTableElement>('grid');
const colgroup = table.querySelector('colgroup')!;
const thead = table.querySelector('thead')!;
const tbody = table.querySelector('tbody')!;
const selRangeEl = $<HTMLDivElement>('sel-range');
const selActiveEl = $<HTMLDivElement>('sel-active');
const antsEl = $<HTMLDivElement>('copy-ants');
const dvBtn = $<HTMLButtonElement>('dv-btn');
const editor = $<HTMLTextAreaElement>('editor');
const finput = $<HTMLTextAreaElement>('finput');
const namebox = $<HTMLInputElement>('namebox');
const tabsEl = $<HTMLDivElement>('tabs');
const statusEl = $<HTMLDivElement>('status');
const menuEl = $<HTMLDivElement>('menu');
const toastEl = $<HTMLDivElement>('toast');
const bannerEl = $<HTMLDivElement>('banner');
const loadingEl = $<HTMLDivElement>('loading');

// =====================================================================================
// State
// =====================================================================================

interface Sel {
  ar: number;
  ac: number;
  fr: number;
  fc: number;
}

interface Editing {
  r: number;
  c: number;
  mode: 'enter' | 'edit';
  original: string;
  via: 'cell' | 'bar';
}

interface SavedState {
  cur: number;
  sels: Record<string, Sel>;
  zooms: Record<string, number>;
}

let meta: WorkbookMeta | null = null;
let settings: Settings = { locale: 'en', gridTheme: 'excel' };
let cur = 0;
let sh: SheetModel | null = null;
let cells = new Map<number, CellModel>();
let merges = new Map<number, [number, number, number, number]>();
let maxRow = 60;
let maxCol = 26;
let renderedRows = 0;
let colLeft: number[] = [];
let rowHeaderW = 46;
const HEADER_H = 22;
let showGrid = true;
let zoom = 1;
let sel: Sel = { ar: 1, ac: 1, fr: 1, fc: 1 };
let editing: Editing | null = null;
let clip: { sheet: number; range: CellRange; tsv: string; cut: boolean } | null = null;
const rowEls = new Map<number, HTMLTableRowElement>();
const frozenTop = new Map<number, number>();
let frozenBottom = 0; // bottom of the frozen rows (content coordinates)
let frozenRight = 0;

const saved = (vscode.getState() as SavedState | undefined) ?? { cur: 0, sels: {}, zooms: {} };
cur = saved.cur ?? 0;

function persist(): void {
  if (sh) {
    saved.sels[sh.name] = sel;
    saved.zooms[sh.name] = zoom;
  }
  saved.cur = cur;
  vscode.setState(saved);
}

// =====================================================================================
// Geometry
// =====================================================================================

const colW = (c: number) => sh?.cols[c]?.w ?? sh?.defaultColW ?? 64;
const colHidden = (c: number) => !!sh?.cols[c]?.hd || colW(c) === 0;
const rowHidden = (r: number) => !!sh?.rows[r]?.hd;
const rowH = (r: number) => sh?.rows[r]?.h ?? sh?.defaultRowH ?? 20;

function nextVisibleCol(c: number): number {
  let n = c + 1;
  while (n <= MAX_COL && colHidden(n)) n++;
  return n;
}
function prevVisibleCol(c: number): number {
  let n = c - 1;
  while (n >= 1 && colHidden(n)) n--;
  return n;
}
function nextVisibleRow(r: number): number {
  let n = r + 1;
  while (n <= MAX_ROW && rowHidden(n)) n++;
  return n;
}
function prevVisibleRow(r: number): number {
  let n = r - 1;
  while (n >= 1 && rowHidden(n)) n--;
  return n;
}
const firstVisibleCol = () => (colHidden(1) ? nextVisibleCol(1) : 1);
const firstVisibleRow = () => (rowHidden(1) ? nextVisibleRow(1) : 1);

function computeColLeft(): void {
  colLeft = new Array(maxCol + 2);
  let x = rowHeaderW;
  for (let c = 1; c <= maxCol + 1; c++) {
    colLeft[c] = x;
    if (c <= maxCol && !colHidden(c)) x += colW(c);
  }
}

function rowTop(r: number): number {
  const tr = rowEls.get(r);
  if (tr) return tr.offsetTop;
  // hidden row: use the position of the next visible one
  const n = nextVisibleRow(r);
  const tr2 = rowEls.get(n);
  return tr2 ? tr2.offsetTop : tbody.offsetTop + tbody.offsetHeight;
}
function rowBottom(r: number): number {
  const tr = rowEls.get(r);
  if (tr) return tr.offsetTop + tr.offsetHeight;
  const p = prevVisibleRow(r);
  const tr2 = rowEls.get(p);
  return tr2 ? tr2.offsetTop + tr2.offsetHeight : HEADER_H;
}

function mergeAt(r: number, c: number): [number, number, number, number] {
  return merges.get(KEY(r, c)) ?? [r, c, r, c];
}

function rectOf(r1: number, c1: number, r2: number, c2: number) {
  ensureRendered(r2);
  const x = colLeft[c1] ?? rowHeaderW;
  const xr = colLeft[Math.min(c2 + 1, maxCol + 1)] ?? x;
  const y = rowTop(r1);
  const yb = rowBottom(r2);
  return { x, y, w: Math.max(0, xr - x), h: Math.max(0, yb - y) };
}

// =====================================================================================
// Effective styles and cell rendering
// =====================================================================================

function effStyle(r: number, c: number): CellStyle | undefined {
  const cell = cells.get(KEY(r, c));
  if (cell) return cell.s;
  return sh?.rows[r]?.s ?? sh?.cols[c]?.s;
}

function sideOf(r: number, c: number, side: 't' | 'r' | 'b' | 'l'): BorderSide | undefined {
  return effStyle(r, c)?.bd?.[side];
}

function isEmpty(r: number, c: number): boolean {
  const cell = cells.get(KEY(r, c));
  return !cell || (!cell.t && !cell.fp && !cell.rt);
}

function borderCss(b: BorderSide | undefined, fallbackGrid: boolean): string {
  if (b) return `${b.s} ${b.c}`;
  return fallbackGrid ? '1px solid var(--grid)' : '1px solid transparent';
}

function cellHtml(r: number, c: number): string {
  let r2 = r;
  let c2 = c;
  let span = '';
  const m = merges.get(KEY(r, c));
  if (m) {
    if (m[0] !== r || m[1] !== c) {
      if (!rowHidden(m[0]) && !colHidden(m[1])) return '';
    } else {
      r2 = m[2];
      c2 = m[3];
      let cs = 0;
      for (let x = c; x <= c2; x++) if (!colHidden(x)) cs++;
      let rs = 0;
      for (let y = r; y <= r2; y++) if (!rowHidden(y)) rs++;
      if (cs > 1) span += ` colspan="${cs}"`;
      if (rs > 1) span += ` rowspan="${rs}"`;
    }
  }
  const cell = cells.get(KEY(r, c));
  const st = effStyle(r, c);
  const kind = cell?.k ?? 'z';
  const cls: string[] = [];
  let css = '';
  if (st?.bg) css += `background:${st.bg};`;
  css += textCss(st);

  // Borders: each cell carries its right and bottom border (plus left/top on the edges).
  const nc = nextVisibleCol(c2);
  const nr = nextVisibleRow(r2);
  const right = sideOf(r, c2, 'r') ?? sideOf(r, nc, 'l');
  const bottom = sideOf(r2, c, 'b') ?? sideOf(nr, c, 't');
  const gridR = showGrid && !st?.bg && !effStyle(r, nc)?.bg;
  const gridB = showGrid && !st?.bg && !effStyle(nr, c)?.bg;
  css += `border-right:${borderCss(right, gridR)};border-bottom:${borderCss(bottom, gridB)};`;
  if (c === firstVisibleCol()) {
    const l = sideOf(r, c, 'l');
    if (l) css += `border-left:${borderCss(l, false)};`;
  }
  if (r === firstVisibleRow()) {
    const t = sideOf(r, c, 't');
    if (t) css += `border-top:${borderCss(t, false)};`;
  }

  const ha = hAlign(st, kind);
  const va = vAlign(st);
  if (ha !== 'left') css += `text-align:${ha};`;
  if (va !== 'bottom') css += `vertical-align:${va};`;
  if (st?.ind) css += ha === 'right' ? `padding-right:${st.ind * 9 + 3}px;` : `padding-left:${st.ind * 9 + 3}px;`;
  if (st?.wr) cls.push('wr');

  // Frozen columns / rows
  if (sh && c <= sh.frozen.cols) {
    cls.push('fz-c');
    css += `left:${colLeft[c]}px;`;
    if (c2 === sh.frozen.cols || nextVisibleCol(c2) > sh.frozen.cols) cls.push('fz-c-last');
  }
  if (sh && r <= sh.frozen.rows) {
    cls.push('fz-r');
    const top = frozenTop.get(r);
    if (top !== undefined) css += `top:${top}px;`;
    if (r2 === sh.frozen.rows || nextVisibleRow(r2) > sh.frozen.rows) cls.push('fz-r-last');
  }

  let text = cell?.t ?? '';
  let inner: string;
  if (cell?.fp) {
    cls.push('fp');
    text = cell.e ?? '';
  }
  if (!st?.wr) text = text.replace(/\r?\n/g, ' ');
  if (cell?.rt && !cell.fp) {
    inner = cell.rt.map((run) => `<span style="${textCss(run.s)}">${esc(st?.wr ? run.t : run.t.replace(/\r?\n/g, ' '))}</span>`).join('');
  } else {
    inner = esc(text);
  }
  if (cell?.link) cls.push('lnk');
  if (kind === 'e') cls.push('err');

  // Text overflow into empty cells to the right (like Excel)
  let innerCss = '';
  if (!st?.wr && !m && (kind === 's' || cell?.fp) && ha === 'left' && text) {
    let w = colW(c);
    let x = c;
    let steps = 0;
    while (steps < 40) {
      x = nextVisibleCol(x);
      if (x > maxCol || !isEmpty(r, x) || merges.has(KEY(r, x))) break;
      w += colW(x);
      steps++;
    }
    if (w > colW(c)) {
      cls.push('spill');
      innerCss = `width:${w - 6}px;`;
    }
  }
  if (sh?.rows[r]?.h && !m) innerCss += `max-height:${rowH(r)}px;`;
  if (st?.rot === 'v') innerCss += 'writing-mode:vertical-rl;text-orientation:upright;';
  else if (st?.rot === 90) innerCss += 'writing-mode:vertical-rl;transform:rotate(180deg);';
  else if (st?.rot === 180) innerCss += 'writing-mode:vertical-rl;';

  let extra = '';
  if (cell?.note) {
    extra += '<span class="note-mark"></span>';
    cls.push('has-note');
  }
  const title = cell?.note ? ` title="${esc(cell.note)}"` : cell?.link ? ` title="${esc(cell.link)} — Ctrl+click to open"` : '';
  return `<td data-c="${c}"${span}${cls.length ? ` class="${cls.join(' ')}"` : ''} style="${css}"${title}><div class="ct"${innerCss ? ` style="${innerCss}"` : ''}>${inner}</div>${extra}</td>`;
}

function rowHtml(r: number): string {
  if (rowHidden(r)) return '';
  let html = `<tr data-r="${r}" style="height:${rowH(r)}px"><th class="rh${sh && r <= sh.frozen.rows ? ' fz-r' : ''}" data-r="${r}"${
    frozenTop.has(r) ? ` style="top:${frozenTop.get(r)}px"` : ''
  }>${r}<div class="rzr" data-r="${r}"></div></th>`;
  for (let c = 1; c <= maxCol; c++) {
    if (colHidden(c)) continue;
    html += cellHtml(r, c);
  }
  return html + '</tr>';
}

// =====================================================================================
// Sheet rendering
// =====================================================================================

function renderAll(keepScroll = true): void {
  if (!sh) return;
  const st = scroller.scrollTop;
  const sl = scroller.scrollLeft;
  rowHeaderW = Math.max(40, String(maxRow).length * 8 + 18);
  computeColLeft();
  frozenTop.clear();
  let cg = `<col style="width:${rowHeaderW}px">`;
  let head = `<tr style="height:${HEADER_H}px"><th class="corner" title="Select all"></th>`;
  for (let c = 1; c <= maxCol; c++) {
    if (colHidden(c)) continue;
    cg += `<col style="width:${colW(c)}px">`;
    const fz = c <= sh.frozen.cols;
    head += `<th class="ch${fz ? ' fz-c' : ''}" data-c="${c}"${fz ? ` style="left:${colLeft[c]}px"` : ''}>${colLetter(c)}<div class="rz" data-c="${c}"></div></th>`;
  }
  colgroup.innerHTML = cg;
  thead.innerHTML = head + '</tr>';
  table.style.width = `${colLeft[maxCol + 1]}px`;
  tbody.innerHTML = '';
  rowEls.clear();
  renderedRows = 0;
  const visibleRows = Math.ceil(scroller.clientHeight / zoom / 18) + 1;
  const wanted = Math.max(Math.ceil((keepScroll ? st / zoom : 0) / 18) + visibleRows + 60, sel.fr + 20, sel.ar + 20);
  renderRowsUntil(Math.min(maxRow, wanted));
  applyFrozenRows();
  sheetEl.style.zoom = String(zoom);
  table.classList.toggle('nogrid', !showGrid);
  if (keepScroll) {
    scroller.scrollTop = st;
    scroller.scrollLeft = sl;
  }
  updateSelection(false);
}

function renderRowsUntil(target: number): void {
  target = Math.min(target, maxRow);
  if (target <= renderedRows) return;
  const start = tbody.rows.length;
  let html = '';
  for (let r = renderedRows + 1; r <= target; r++) html += rowHtml(r);
  tbody.insertAdjacentHTML('beforeend', html);
  for (let i = start; i < tbody.rows.length; i++) {
    const tr = tbody.rows[i];
    rowEls.set(Number(tr.dataset.r), tr);
  }
  renderedRows = target;
}

function ensureRendered(r: number): void {
  if (r > maxRow) {
    maxRow = Math.min(MAX_ROW, r + 50);
  }
  if (r > renderedRows) renderRowsUntil(Math.min(maxRow, r + 50));
}

function applyFrozenRows(): void {
  frozenBottom = HEADER_H;
  frozenRight = rowHeaderW;
  if (!sh) return;
  if (sh.frozen.cols) frozenRight = colLeft[Math.min(sh.frozen.cols + 1, maxCol + 1)];
  if (!sh.frozen.rows) return;
  let top = (thead.rows[0] as HTMLTableRowElement | undefined)?.offsetHeight ?? HEADER_H;
  for (let r = 1; r <= sh.frozen.rows; r++) {
    const tr = rowEls.get(r);
    if (!tr) continue;
    frozenTop.set(r, top);
    for (const cell of Array.from(tr.children) as HTMLElement[]) cell.style.top = `${top}px`;
    top += tr.offsetHeight;
  }
  frozenBottom = top;
}

/** Re-renders changed cells and their neighbours (borders, overflow). */
function rerenderCells(list: [number, number][]): void {
  const todo = new Map<number, [number, number]>();
  const add = (r: number, c: number) => {
    if (r < 1 || c < 1) return;
    const m = merges.get(KEY(r, c));
    if (m) {
      r = m[0];
      c = m[1];
    }
    todo.set(KEY(r, c), [r, c]);
  };
  for (const [r, c] of list) {
    add(r, c);
    add(prevVisibleRow(r), c);
    let x = c;
    for (let i = 0; i < 40; i++) {
      x = prevVisibleCol(x);
      if (x < 1) break;
      add(r, x);
      if (!isEmpty(r, x)) break;
    }
  }
  for (const [r, c] of todo.values()) {
    const tr = rowEls.get(r);
    if (!tr) continue;
    const td = tr.querySelector<HTMLTableCellElement>(`td[data-c="${c}"]`);
    if (!td) continue;
    const html = cellHtml(r, c);
    if (html) td.outerHTML = html;
  }
}

// =====================================================================================
// Selection
// =====================================================================================

function normSel(): CellRange {
  let r1 = Math.min(sel.ar, sel.fr);
  let r2 = Math.max(sel.ar, sel.fr);
  let c1 = Math.min(sel.ac, sel.fc);
  let c2 = Math.max(sel.ac, sel.fc);
  if (sh && sh.merges.length && r2 - r1 < 100000) {
    let changed = true;
    while (changed) {
      changed = false;
      for (const [a, b, c, d] of sh.merges) {
        if (c < r1 || a > r2 || d < c1 || b > c2) continue;
        if (a < r1) { r1 = a; changed = true; }
        if (c > r2) { r2 = c; changed = true; }
        if (b < c1) { c1 = b; changed = true; }
        if (d > c2) { c2 = d; changed = true; }
      }
    }
  }
  return { r1, c1, r2, c2 };
}

/** Selection clamped to the rendered dimensions (for operations). */
function opRange(): CellRange {
  const rg = normSel();
  return { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, Math.max(maxRow, sh?.rowCount ?? 0)), c2: Math.min(rg.c2, Math.max(maxCol, sh?.colCount ?? 0)) };
}

const isFullCols = (rg: CellRange) => rg.r1 === 1 && rg.r2 >= MAX_ROW;
const isFullRows = (rg: CellRange) => rg.c1 === 1 && rg.c2 >= MAX_COL;

function addressText(r: number, c: number): string {
  return colLetter(c) + r;
}

let prevHeaderSel: Element[] = [];

function placeOverlay(el: HTMLElement, rect: { x: number; y: number; w: number; h: number }): void {
  el.style.left = `${rect.x}px`;
  el.style.top = `${rect.y}px`;
  el.style.width = `${rect.w}px`;
  el.style.height = `${rect.h}px`;
}

function updateSelection(scroll = true): void {
  if (!sh) return;
  const rg = normSel();
  const vr2 = Math.min(rg.r2, maxRow);
  const vc2 = Math.min(rg.c2, maxCol);
  placeOverlay(selRangeEl, rectOf(rg.r1, rg.c1, vr2, vc2));
  const m = mergeAt(sel.ar, sel.ac);
  placeOverlay(selActiveEl, rectOf(m[0], m[1], m[2], m[3]));
  selRangeEl.hidden = rg.r1 === m[0] && rg.r2 === m[2] && rg.c1 === m[1] && rg.c2 === m[3];
  editor.style.left = selActiveEl.style.left;
  editor.style.top = selActiveEl.style.top;

  // headers
  for (const el of prevHeaderSel) el.classList.remove('sel', 'sel-full');
  prevHeaderSel = [];
  const fullC = isFullCols(rg);
  const fullR = isFullRows(rg);
  for (let c = rg.c1; c <= vc2; c++) {
    const th = thead.querySelector(`th[data-c="${c}"]`);
    if (th) {
      th.classList.add('sel');
      if (fullC) th.classList.add('sel-full');
      prevHeaderSel.push(th);
    }
  }
  const lastR = Math.min(vr2, renderedRows);
  for (let r = rg.r1; r <= lastR; r++) {
    const th = rowEls.get(r)?.firstElementChild;
    if (th) {
      th.classList.add('sel');
      if (fullR) th.classList.add('sel-full');
      prevHeaderSel.push(th);
    }
  }

  // pending copy
  if (clip && clip.sheet === cur) {
    antsEl.hidden = false;
    placeOverlay(antsEl, rectOf(clip.range.r1, clip.range.c1, Math.min(clip.range.r2, maxRow), Math.min(clip.range.c2, maxCol)));
  } else antsEl.hidden = true;

  const active = cells.get(KEY(m[0], m[1]));
  if (!editing) {
    namebox.value = addressText(sel.ar, sel.ac);
    finput.value = active?.e ?? '';
  }
  // drop-down list
  if (active?.dv && !editing) {
    const rect = rectOf(m[0], m[1], m[2], m[3]);
    dvBtn.hidden = false;
    dvBtn.style.left = `${rect.x + rect.w + 1}px`;
    dvBtn.style.top = `${rect.y + rect.h - 18}px`;
  } else dvBtn.hidden = true;

  updateToolbar();
  updateStatus(rg);
  updateOverlayClip();
  if (scroll) scrollIntoView(sel.fr, sel.fc);
  persist();
}

function updateOverlayClip(): void {
  if (!sh) return;
  const st = scroller.scrollTop / zoom;
  const sl = scroller.scrollLeft / zoom;
  const rg = normSel();
  const inFrozenRows = sh.frozen.rows > 0 && rg.r2 <= sh.frozen.rows;
  const inFrozenCols = sh.frozen.cols > 0 && rg.c2 <= sh.frozen.cols;
  for (const el of [selRangeEl, selActiveEl, antsEl, dvBtn, editor]) {
    const y = parseFloat(el.style.top) || 0;
    const x = parseFloat(el.style.left) || 0;
    const ty = inFrozenRows ? st : 0;
    const tx = inFrozenCols ? sl : 0;
    el.style.transform = tx || ty ? `translate(${tx}px, ${ty}px)` : '';
    if (el === editor) continue;
    const clipTop = inFrozenRows ? 0 : Math.max(0, frozenBottom + st - y);
    const clipLeft = inFrozenCols ? 0 : Math.max(0, frozenRight + sl - x);
    el.style.clipPath = clipTop || clipLeft ? `inset(${clipTop}px 0 0 ${clipLeft}px)` : '';
  }
}

function scrollIntoView(r: number, c: number): void {
  if (!sh) return;
  const m = mergeAt(r, c);
  const rect = rectOf(m[0], m[1], m[2], m[3]);
  const z = zoom;
  if (r > sh.frozen.rows) {
    const topBound = scroller.scrollTop + frozenBottom * z;
    const bottomBound = scroller.scrollTop + scroller.clientHeight;
    if (rect.y * z < topBound) scroller.scrollTop = rect.y * z - frozenBottom * z;
    else if ((rect.y + rect.h) * z > bottomBound) scroller.scrollTop = (rect.y + rect.h) * z - scroller.clientHeight + 2;
  }
  if (c > sh.frozen.cols) {
    const leftBound = scroller.scrollLeft + frozenRight * z;
    const rightBound = scroller.scrollLeft + scroller.clientWidth;
    if (rect.x * z < leftBound) scroller.scrollLeft = rect.x * z - frozenRight * z;
    else if ((rect.x + rect.w) * z > rightBound) scroller.scrollLeft = Math.min(rect.x * z - frozenRight * z, (rect.x + rect.w) * z - scroller.clientWidth + 2);
  }
}

function setSel(r: number, c: number, extend = false): void {
  if (extend) {
    sel = { ...sel, fr: r, fc: c };
  } else {
    sel = { ar: r, ac: c, fr: r, fc: c };
  }
}

function growTo(r: number, c: number): void {
  let rerender = false;
  if (r > maxRow - 5 && maxRow < MAX_ROW) {
    maxRow = Math.min(MAX_ROW, r + 60);
  }
  if (c > maxCol - 2 && maxCol < MAX_COL) {
    maxCol = Math.min(MAX_COL, c + 10);
    rerender = true;
  }
  if (rerender) renderAll(true);
  else ensureRendered(r);
}

function step(v: number, d: number, hidden: (n: number) => boolean, max: number): number {
  if (!d) return v;
  const sign = Math.sign(d);
  for (let i = 0; i < Math.abs(d); i++) {
    let n = v + sign;
    while (n >= 1 && n <= max && hidden(n)) n += sign;
    if (n < 1 || n > max) break;
    v = n;
  }
  return v;
}

function jump(r: number, c: number, dr: number, dc: number): [number, number] {
  const limitR = Math.max(sh?.rowCount ?? 1, 1);
  const limitC = Math.max(sh?.colCount ?? 1, 1);
  const next = (rr: number, cc: number): [number, number] => [rr + dr, cc + dc];
  const inside = (rr: number, cc: number) => rr >= 1 && cc >= 1 && rr <= Math.max(limitR, r) && cc <= Math.max(limitC, c);
  let [nr, nc] = next(r, c);
  if (!inside(nr, nc)) {
    return [dr < 0 ? 1 : dr > 0 ? Math.max(limitR, r) : r, dc < 0 ? 1 : dc > 0 ? Math.max(limitC, c) : c];
  }
  if (!isEmpty(r, c) && !isEmpty(nr, nc)) {
    // advance to the last non-empty cell
    while (inside(nr + dr, nc + dc) && !isEmpty(nr + dr, nc + dc)) [nr, nc] = next(nr, nc);
    return [nr, nc];
  }
  while (inside(nr, nc) && isEmpty(nr, nc)) [nr, nc] = next(nr, nc);
  if (!inside(nr, nc)) return [dr < 0 ? 1 : dr > 0 ? Math.max(limitR, r) : r, dc < 0 ? 1 : dc > 0 ? Math.max(limitC, c) : c];
  return [nr, nc];
}

function move(dr: number, dc: number, extend = false, ctrl = false): void {
  if (!sh) return;
  let r = extend ? sel.fr : sel.ar;
  let c = extend ? sel.fc : sel.ac;
  if (!extend) {
    const m = mergeAt(r, c);
    if (dr > 0) r = m[2];
    if (dc > 0) c = m[3];
  }
  if (ctrl) [r, c] = jump(r, c, Math.sign(dr), Math.sign(dc));
  else {
    r = step(r, dr, rowHidden, MAX_ROW);
    c = step(c, dc, colHidden, MAX_COL);
  }
  if (!extend) {
    const m2 = merges.get(KEY(r, c));
    if (m2) {
      r = m2[0];
      c = m2[1];
    }
  }
  growTo(r, c);
  setSel(r, c, extend);
  updateSelection();
}

/** Enter/Tab movement inside a selected range. */
function moveWithin(dr: number, dc: number): void {
  const rg = normSel();
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return move(dr, dc);
  let { ar: r, ac: c } = sel;
  if (dr) {
    r += dr;
    if (r > rg.r2) { r = rg.r1; c = c + 1 > rg.c2 ? rg.c1 : c + 1; }
    if (r < rg.r1) { r = rg.r2; c = c - 1 < rg.c1 ? rg.c2 : c - 1; }
  } else {
    c += dc;
    if (c > rg.c2) { c = rg.c1; r = r + 1 > rg.r2 ? rg.r1 : r + 1; }
    if (c < rg.c1) { c = rg.c2; r = r - 1 < rg.r1 ? rg.r2 : r - 1; }
  }
  sel = { ...sel, ar: r, ac: c };
  updateSelection();
}

// =====================================================================================
// Toolbar / status bar
// =====================================================================================

const tbFont = $<HTMLSelectElement>('tb-font');
const tbSize = $<HTMLSelectElement>('tb-size');
const tbNumfmt = $<HTMLSelectElement>('tb-numfmt');
const tbBorder = $<HTMLSelectElement>('tb-border');
const tbZoom = $<HTMLSelectElement>('tb-zoom');
const tbFc = $<HTMLInputElement>('tb-fc');
const tbBg = $<HTMLInputElement>('tb-bg');

function activeStyle(): CellStyle | undefined {
  const m = mergeAt(sel.ar, sel.ac);
  return effStyle(m[0], m[1]);
}

function setPressed(cmd: string, on: boolean): void {
  document.querySelector(`#toolbar button[data-cmd="${cmd}"]`)?.classList.toggle('on', on);
}

function ensureOption(select: HTMLSelectElement, value: string, label = value): void {
  if (!Array.from(select.options).some((o) => o.value === value)) {
    const o = document.createElement('option');
    o.value = value;
    o.textContent = label;
    select.insertBefore(o, select.firstChild);
  }
}

function updateToolbar(): void {
  const s = activeStyle();
  const font = s?.ff ?? meta?.defaultFont ?? 'Calibri';
  ensureOption(tbFont, font);
  tbFont.value = font;
  const size = String(s?.fs ?? 11);
  ensureOption(tbSize, size);
  tbSize.value = size;
  setPressed('bold', !!s?.b);
  setPressed('italic', !!s?.i);
  setPressed('underline', !!s?.u);
  setPressed('strike', !!s?.st);
  setPressed('left', s?.ha === 'left');
  setPressed('center', s?.ha === 'center' || s?.ha === 'centerContinuous');
  setPressed('right', s?.ha === 'right');
  setPressed('top', s?.va === 'top');
  setPressed('middle', s?.va === 'middle');
  setPressed('bottom', s?.va === 'bottom');
  setPressed('wrap', !!s?.wr);
  setPressed('merge', merges.has(KEY(sel.ar, sel.ac)));
  setPressed('grid', showGrid);
  const nf = s?.nf ?? 'General';
  tbNumfmt.value = NUMFMTS.some(([v]) => v === nf) ? nf : '__custom';
  tbNumfmt.title = `Number format: ${nf}`;
  if (s?.fc) $('fc-bar').style.background = s.fc;
  tbZoom.value = String(Math.round(zoom * 100));
}

function fmtNumber(n: number): string {
  return n.toLocaleString(settings.locale === 'fr' ? 'fr-FR' : 'en-US', { maximumFractionDigits: 10 });
}

function updateStatus(rg: CellRange): void {
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) {
    statusEl.textContent = '';
    return;
  }
  let sum = 0;
  let nums = 0;
  let filled = 0;
  for (const cell of cells.values()) {
    if (cell.r < rg.r1 || cell.r > rg.r2 || cell.c < rg.c1 || cell.c > rg.c2 || cell.m) continue;
    if (cell.t || cell.rt) filled++;
    if ((cell.k === 'n' || cell.k === 'd') && typeof cell.v === 'number' && cell.k === 'n') {
      sum += cell.v;
      nums++;
    }
  }
  const parts: string[] = [];
  if (nums) {
    parts.push(`Average: ${fmtNumber(sum / nums)}`);
    parts.push(`Numerical Count: ${nums}`);
  }
  parts.push(`Count: ${filled}`);
  if (nums) parts.push(`Sum: ${fmtNumber(sum)}`);
  statusEl.textContent = parts.join('    ');
}

function sendFormat(patch: FormatPatch): void {
  if (!sh) return;
  post({ type: 'format', sheet: cur, ranges: [opRange()], patch });
}

function toolbarCommand(cmd: string): void {
  const s = activeStyle();
  switch (cmd) {
    case 'bold':
      sendFormat({ bold: !s?.b });
      break;
    case 'italic':
      sendFormat({ italic: !s?.i });
      break;
    case 'underline':
      sendFormat({ underline: !s?.u });
      break;
    case 'strike':
      sendFormat({ strike: !s?.st });
      break;
    case 'nofill':
      sendFormat({ fill: null });
      break;
    case 'left':
    case 'center':
    case 'right': {
      const on = cmd === 'center' ? s?.ha === 'center' || s?.ha === 'centerContinuous' : s?.ha === cmd;
      sendFormat({ hAlign: on ? null : cmd });
      break;
    }
    case 'top':
    case 'middle':
    case 'bottom':
      sendFormat({ vAlign: s?.va === cmd ? null : cmd });
      break;
    case 'wrap':
      sendFormat({ wrap: !s?.wr });
      break;
    case 'merge': {
      const rg = normSel();
      const single = rg.r1 === rg.r2 && rg.c1 === rg.c2;
      const hasMerge = sh?.merges.some(([a, b, c, d]) => !(c < rg.r1 || a > rg.r2 || d < rg.c1 || b > rg.c2));
      if (single && !hasMerge) break;
      post({ type: 'merge', sheet: cur, range: opRange(), unmerge: !!hasMerge });
      break;
    }
    case 'clearFormat':
      post({ type: 'clear', sheet: cur, ranges: [opRange()], what: 'formats' });
      break;
    case 'grid':
      showGrid = !showGrid;
      renderAll(true);
      break;
  }
  focusGrid();
}

document.getElementById('toolbar')!.addEventListener('mousedown', (e) => {
  // keep focus on the grid when clicking buttons
  if ((e.target as HTMLElement).closest('button')) e.preventDefault();
});
document.getElementById('toolbar')!.addEventListener('click', (e) => {
  const btn = (e.target as HTMLElement).closest('button');
  if (btn?.dataset.cmd) {
    commitEdit();
    toolbarCommand(btn.dataset.cmd);
  }
});
tbFont.addEventListener('change', () => {
  sendFormat({ fontName: tbFont.value });
  focusGrid();
});
tbSize.addEventListener('change', () => {
  sendFormat({ fontSize: Number(tbSize.value) });
  focusGrid();
});
tbNumfmt.addEventListener('change', () => {
  if (tbNumfmt.value !== '__custom') sendFormat({ numFmt: tbNumfmt.value });
  focusGrid();
});
tbBorder.addEventListener('change', () => {
  if (tbBorder.value) sendFormat({ border: tbBorder.value as FormatPatch['border'] });
  tbBorder.value = '';
  focusGrid();
});
tbZoom.addEventListener('change', () => {
  setZoom(Number(tbZoom.value) / 100);
  focusGrid();
});
tbFc.addEventListener('change', () => {
  $('fc-bar').style.background = tbFc.value;
  sendFormat({ fontColor: tbFc.value });
  focusGrid();
});
tbBg.addEventListener('change', () => {
  $('bg-bar').style.background = tbBg.value;
  sendFormat({ fill: tbBg.value });
  focusGrid();
});
$('fc-bar').style.background = tbFc.value;
$('bg-bar').style.background = tbBg.value;

function setZoom(z: number): void {
  const centerR = scroller.scrollTop / zoom;
  const centerC = scroller.scrollLeft / zoom;
  zoom = Math.min(4, Math.max(0.25, z));
  sheetEl.style.zoom = String(zoom);
  scroller.scrollTop = centerR * zoom;
  scroller.scrollLeft = centerC * zoom;
  updateSelection(false);
}

scroller.addEventListener(
  'wheel',
  (e) => {
    if (e.ctrlKey) {
      e.preventDefault();
      setZoom(zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1));
    }
  },
  { passive: false },
);

// =====================================================================================
// Editing
// =====================================================================================

function focusGrid(): void {
  if (!editing || editing.via === 'cell') editor.focus({ preventScroll: true });
}

function beginEdit(mode: 'enter' | 'edit', initial?: string, via: 'cell' | 'bar' = 'cell'): void {
  if (!sh) return;
  const m = mergeAt(sel.ar, sel.ac);
  const cell = cells.get(KEY(m[0], m[1]));
  const original = cell?.e ?? '';
  editing = { r: m[0], c: m[1], mode, original, via };
  clearAnts();
  dvBtn.hidden = true;
  if (via === 'bar') return;
  if (initial !== undefined) editor.value = initial;
  else editor.value = original;
  const st = effStyle(m[0], m[1]);
  const rect = rectOf(m[0], m[1], m[2], m[3]);
  editor.classList.add('editing');
  editor.style.font = cssFont(st, meta?.defaultFont ?? 'Calibri');
  editor.style.color = st?.fc ?? '';
  editor.style.background = st?.bg && !st.bg.includes('gradient') ? st.bg : '';
  editor.style.textAlign = hAlign(st, cell?.k ?? 's') === 'right' && mode === 'edit' ? 'right' : 'left';
  editor.style.whiteSpace = st?.wr ? 'pre-wrap' : 'pre';
  editor.style.minWidth = `${rect.w}px`;
  editor.style.minHeight = `${rect.h}px`;
  editor.style.width = `${rect.w}px`;
  editor.style.height = `${rect.h}px`;
  editor.dataset.maxw = String(st?.wr ? rect.w : Math.max(rect.w, scroller.clientWidth / zoom - (rect.x - scroller.scrollLeft / zoom) - 4));
  autosizeEditor();
  finput.value = editor.value;
  const end = editor.value.length;
  editor.setSelectionRange(end, end);
  editor.focus({ preventScroll: true });
}

function autosizeEditor(): void {
  if (!editing || editing.via !== 'cell') return;
  editor.style.width = editor.style.minWidth;
  const maxw = Number(editor.dataset.maxw) || 800;
  editor.style.width = `${Math.min(maxw, Math.max(parseFloat(editor.style.minWidth), editor.scrollWidth + 4))}px`;
  editor.style.height = editor.style.minHeight;
  editor.style.height = `${Math.max(parseFloat(editor.style.minHeight), editor.scrollHeight)}px`;
}

function endEdit(): void {
  editing = null;
  editor.classList.remove('editing');
  editor.value = '';
  editor.removeAttribute('style');
  updateSelection(false);
}

function commitEdit(): void {
  if (!editing) return;
  const value = editing.via === 'bar' ? finput.value : editor.value;
  const { r, c, original } = editing;
  endEdit();
  if (value !== original) {
    const cell = cells.get(KEY(r, c));
    // show the value immediately while waiting for the extension to respond
    if (!value.startsWith('=')) {
      cells.set(KEY(r, c), { ...(cell ?? { r, c, k: 's' }), k: cell?.k === 'z' || !cell ? 's' : cell.k, t: value, e: value, rt: undefined, fp: undefined });
      rerenderCells([[r, c]]);
    }
    post({ type: 'setCells', sheet: cur, items: [{ r, c, input: value }] });
  }
}

function cancelEdit(): void {
  if (!editing) return;
  endEdit();
}

editor.addEventListener('input', () => {
  if (!editing) {
    const v = editor.value;
    beginEdit('enter', v);
    return;
  }
  finput.value = editor.value;
  autosizeEditor();
});
editor.addEventListener('compositionstart', () => {
  if (!editing) beginEdit('enter', '');
});

finput.addEventListener('focus', () => {
  if (editing?.via === 'cell') {
    // move editing over to the formula bar
    const v = editor.value;
    const prev = editing;
    editor.classList.remove('editing');
    editor.removeAttribute('style');
    editing = { ...prev, via: 'bar', mode: 'edit' };
    finput.value = v;
    return;
  }
  if (!editing) beginEdit('edit', undefined, 'bar');
});
finput.addEventListener('keydown', (e) => {
  e.stopPropagation();
  if (e.key === 'Enter' && !e.altKey) {
    e.preventDefault();
    commitEdit();
    focusGrid();
    move(e.shiftKey ? -1 : 1, 0);
  } else if (e.key === 'Enter' && e.altKey) {
    e.preventDefault();
    finput.setRangeText('\n', finput.selectionStart, finput.selectionEnd, 'end');
  } else if (e.key === 'Tab') {
    e.preventDefault();
    commitEdit();
    focusGrid();
    move(0, e.shiftKey ? -1 : 1);
  } else if (e.key === 'Escape') {
    e.preventDefault();
    cancelEdit();
    focusGrid();
  }
});
finput.addEventListener('blur', () => {
  if (editing?.via === 'bar') commitEdit();
});

namebox.addEventListener('keydown', (e) => {
  e.stopPropagation();
  if (e.key === 'Enter') {
    e.preventDefault();
    const m = /^\s*\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?\s*$/.exec(namebox.value);
    if (m) {
      const c1 = colNumber(m[2] ? m[1] : 'A');
      const r1 = Number(m[2]);
      const c2 = m[3] ? colNumber(m[3]) : c1;
      const r2 = m[4] ? Number(m[4]) : r1;
      if (r1 >= 1 && r2 <= MAX_ROW && c1 >= 1 && c2 <= MAX_COL) {
        growTo(Math.max(r1, r2), Math.max(c1, c2));
        sel = { ar: r1, ac: c1, fr: r2, fc: c2 };
        updateSelection();
      }
    }
    focusGrid();
  } else if (e.key === 'Escape') {
    focusGrid();
  }
});
namebox.addEventListener('focus', () => namebox.select());

function isCtrl(e: KeyboardEvent | MouseEvent): boolean {
  return e.ctrlKey || e.metaKey;
}

editor.addEventListener('keydown', (e) => {
  if (editing) {
    const k = e.key;
    if (isCtrl(e) && (k === 'z' || k === 'y' || k === 'Z' || k === 'Y')) {
      e.stopPropagation(); // native undo inside the input
      return;
    }
    if (k === 'Enter' && !e.altKey) {
      e.preventDefault();
      commitEdit();
      moveWithin(e.shiftKey ? -1 : 1, 0);
    } else if (k === 'Enter' && e.altKey) {
      e.preventDefault();
      editor.setRangeText('\n', editor.selectionStart, editor.selectionEnd, 'end');
      finput.value = editor.value;
      autosizeEditor();
    } else if (k === 'Tab') {
      e.preventDefault();
      commitEdit();
      moveWithin(0, e.shiftKey ? -1 : 1);
    } else if (k === 'Escape') {
      e.preventDefault();
      cancelEdit();
    } else if (k === 'F2') {
      e.preventDefault();
      editing.mode = editing.mode === 'enter' ? 'edit' : 'enter';
    } else if (editing.mode === 'enter' && k.startsWith('Arrow')) {
      e.preventDefault();
      commitEdit();
      const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k as 'ArrowUp']!;
      move(d[0], d[1]);
    }
    return;
  }
  onGridKey(e);
});

function onGridKey(e: KeyboardEvent): void {
  const k = e.key;
  const ctrl = isCtrl(e);
  let handled = true;
  closeMenu();
  switch (k) {
    case 'ArrowUp':
      if (e.altKey) handled = false;
      else move(-1, 0, e.shiftKey, ctrl);
      break;
    case 'ArrowDown':
      if (e.altKey) openList();
      else move(1, 0, e.shiftKey, ctrl);
      break;
    case 'ArrowLeft':
      move(0, -1, e.shiftKey, ctrl);
      break;
    case 'ArrowRight':
      move(0, 1, e.shiftKey, ctrl);
      break;
    case 'Tab':
      moveWithin(0, e.shiftKey ? -1 : 1);
      break;
    case 'Enter':
      if (ctrl || e.altKey) handled = false;
      else moveWithin(e.shiftKey ? -1 : 1, 0);
      break;
    case 'Home':
      if (ctrl) {
        setSel(firstVisibleRow(), firstVisibleCol(), e.shiftKey);
        updateSelection();
      } else move(0, -(sel.fc - 1), e.shiftKey);
      break;
    case 'End':
      if (ctrl && sh) {
        setSel(Math.max(1, sh.rowCount), Math.max(1, sh.colCount), e.shiftKey);
        updateSelection();
      } else handled = false;
      break;
    case 'PageDown':
    case 'PageUp': {
      const n = Math.max(1, Math.floor(scroller.clientHeight / zoom / (sh?.defaultRowH ?? 20)) - 1);
      if (e.altKey) move(0, k === 'PageDown' ? 8 : -8, e.shiftKey);
      else move(k === 'PageDown' ? n : -n, 0, e.shiftKey);
      break;
    }
    case 'F2':
      beginEdit('edit');
      break;
    case 'Delete':
      post({ type: 'clear', sheet: cur, ranges: [opRange()], what: 'contents' });
      break;
    case 'Backspace':
      beginEdit('enter', '');
      break;
    case 'Escape':
      if (clip) clearAnts();
      else handled = false;
      break;
    case ' ':
      if (ctrl) {
        sel = { ar: 1, ac: sel.ac, fr: MAX_ROW, fc: sel.fc };
        updateSelection(false);
      } else if (e.shiftKey) {
        sel = { ar: sel.ar, ac: 1, fr: sel.fr, fc: MAX_COL };
        updateSelection(false);
      } else handled = false;
      break;
    default:
      handled = false;
      if (ctrl && !e.altKey) {
        const lower = k.toLowerCase();
        handled = true;
        if (lower === 'a') {
          sel = { ar: 1, ac: 1, fr: MAX_ROW, fc: MAX_COL };
          updateSelection(false);
        } else if (lower === 'b') toolbarCommand('bold');
        else if (lower === 'i') toolbarCommand('italic');
        else if (lower === 'u') toolbarCommand('underline');
        else if (k === '5') toolbarCommand('strike');
        else if (lower === 'd') fillDown('down');
        else if (lower === 'r') fillDown('right');
        else handled = false;
      }
  }
  if (handled) {
    e.preventDefault();
    e.stopPropagation();
  }
}

function fillDown(dir: 'down' | 'right'): void {
  const rg = opRange();
  if (dir === 'down' && rg.r2 > rg.r1) {
    post({ type: 'paste', sheet: cur, target: { ...rg, r1: rg.r1 + 1 }, internal: { sheet: cur, range: { ...rg, r2: rg.r1 }, cut: false } });
  } else if (dir === 'right' && rg.c2 > rg.c1) {
    post({ type: 'paste', sheet: cur, target: { ...rg, c1: rg.c1 + 1 }, internal: { sheet: cur, range: { ...rg, c2: rg.c1 }, cut: false } });
  }
}

// =====================================================================================
// Clipboard
// =====================================================================================

function clearAnts(): void {
  clip = null;
  antsEl.hidden = true;
}

function selectionText(rg: CellRange): { tsv: string; html: string } {
  const r2 = Math.min(rg.r2, Math.max(sh?.rowCount ?? 1, rg.r1));
  const c2 = Math.min(rg.c2, Math.max(sh?.colCount ?? 1, rg.c1));
  const lines: string[] = [];
  let html = '<table style="border-collapse:collapse">';
  for (let r = rg.r1; r <= r2; r++) {
    if (rowHidden(r)) continue;
    const fields: string[] = [];
    html += '<tr>';
    for (let c = rg.c1; c <= c2; c++) {
      if (colHidden(c)) continue;
      const cell = cells.get(KEY(r, c));
      const t = cell?.m ? '' : cell?.fp ? cell.e ?? '' : cell?.t ?? '';
      fields.push(tsvField(t));
      const st = cell?.s;
      html += `<td style="${st?.bg ? `background:${st.bg};` : ''}${textCss(st)}">${esc(t)}</td>`;
    }
    html += '</tr>';
    lines.push(fields.join('\t'));
  }
  html += '</table>';
  return { tsv: lines.join('\r\n'), html };
}

function isGridFocus(): boolean {
  const a = document.activeElement;
  return a === editor || a === document.body || a === null;
}

function doCopy(e: ClipboardEvent, cut: boolean): void {
  if (editing || !isGridFocus() || !sh) return;
  e.preventDefault();
  const rg = normSel();
  const { tsv, html } = selectionText(rg);
  e.clipboardData?.setData('text/plain', tsv);
  e.clipboardData?.setData('text/html', html);
  clip = { sheet: cur, range: { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, Math.max(sh.rowCount, rg.r1)), c2: Math.min(rg.c2, Math.max(sh.colCount, rg.c1)) }, tsv, cut };
  updateSelection(false);
}

document.addEventListener('copy', (e) => doCopy(e, false));
document.addEventListener('cut', (e) => doCopy(e, true));
document.addEventListener('paste', (e) => {
  if (editing || !isGridFocus() || !sh) return;
  e.preventDefault();
  pasteText(e.clipboardData?.getData('text/plain') ?? '');
});

function pasteText(text: string): void {
  if (!sh) return;
  const norm = (s: string) => s.replace(/\r\n?/g, '\n').replace(/\n$/, '');
  const target = opRange();
  if (clip && norm(text) === norm(clip.tsv)) {
    post({ type: 'paste', sheet: cur, target, internal: { sheet: clip.sheet, range: clip.range, cut: clip.cut } });
    if (clip.cut) clearAnts();
  } else {
    if (!text) return;
    post({ type: 'paste', sheet: cur, target, rows: parseTsv(text) });
    clearAnts();
  }
}

// =====================================================================================
// Mouse
// =====================================================================================

let dragging: null | 'cells' | 'cols' | 'rows' = null;
let resizing: null | { kind: 'col' | 'row'; index: number; start: number; size: number; el: HTMLElement } = null;

function cellFromEvent(target: EventTarget | null): [number, number] | undefined {
  const td = (target as HTMLElement | null)?.closest?.('td');
  if (!td || !tbody.contains(td)) return undefined;
  const tr = td.parentElement as HTMLTableRowElement;
  return [Number(tr.dataset.r), Number((td as HTMLElement).dataset.c)];
}

table.addEventListener('mousedown', (e) => {
  if (!sh) return;
  closeMenu();
  const t = e.target as HTMLElement;
  if (t.classList.contains('rz') || t.classList.contains('rzr')) {
    e.preventDefault();
    commitEdit();
    const kind = t.classList.contains('rz') ? 'col' : 'row';
    const index = Number(kind === 'col' ? t.dataset.c : t.dataset.r);
    const el =
      kind === 'col'
        ? (colgroup.children[Array.from(thead.rows[0].cells).findIndex((th) => (th as HTMLElement).dataset.c === String(index))] as HTMLElement)
        : rowEls.get(index)!;
    resizing = { kind, index, start: kind === 'col' ? e.clientX : e.clientY, size: kind === 'col' ? colW(index) : rowEls.get(index)!.offsetHeight, el };
    document.body.classList.add(kind === 'col' ? 'resizing-col' : 'resizing-row');
    return;
  }
  if (e.button !== 0 && e.button !== 2) return;
  const wasEditing = !!editing;
  commitEdit();
  e.preventDefault();
  const th = t.closest('th');
  if (th) {
    if (th.classList.contains('corner')) {
      sel = { ar: 1, ac: 1, fr: MAX_ROW, fc: MAX_COL };
    } else if (th.classList.contains('ch')) {
      const c = Number(th.dataset.c);
      const rg = normSel();
      if (e.button === 2 && isFullCols(rg) && c >= rg.c1 && c <= rg.c2) {
        // right-click inside the selection: keep it
      } else if (e.shiftKey) sel = { ...sel, ar: 1, fr: MAX_ROW, fc: c };
      else sel = { ar: 1, ac: c, fr: MAX_ROW, fc: c };
      dragging = 'cols';
    } else if (th.classList.contains('rh')) {
      const r = Number(th.dataset.r);
      const rg = normSel();
      if (e.button === 2 && isFullRows(rg) && r >= rg.r1 && r <= rg.r2) {
        // keep the selection
      } else if (e.shiftKey) sel = { ...sel, ac: 1, fc: MAX_COL, fr: r };
      else sel = { ar: r, ac: 1, fr: r, fc: MAX_COL };
      dragging = 'rows';
    }
    if (e.button === 2) dragging = null;
    updateSelection(false);
    focusGrid();
    return;
  }
  const pos = cellFromEvent(t);
  if (!pos) return;
  const [r, c] = pos;
  if (e.button === 2) {
    const rg = normSel();
    if (!(r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2)) setSel(r, c);
    updateSelection(false);
    focusGrid();
    return;
  }
  if (isCtrl(e)) {
    const cell = cells.get(KEY(r, c));
    if (cell?.link) {
      post({ type: 'openLink', url: cell.link });
      return;
    }
  }
  setSel(r, c, e.shiftKey);
  dragging = 'cells';
  updateSelection(false);
  if (!wasEditing) focusGrid();
  else focusGrid();
});

document.addEventListener('mousemove', (e) => {
  if (resizing) {
    const delta = ((resizing.kind === 'col' ? e.clientX : e.clientY) - resizing.start) / zoom;
    const size = Math.max(resizing.kind === 'col' ? 0 : 4, Math.round(resizing.size + delta));
    if (resizing.kind === 'col') resizing.el.style.width = `${size}px`;
    else resizing.el.style.height = `${size}px`;
    resizing.el.dataset.size = String(size);
    return;
  }
  if (!dragging || !sh) return;
  const el = document.elementFromPoint(e.clientX, e.clientY) as HTMLElement | null;
  if (!el) return;
  if (dragging === 'cells') {
    const pos = cellFromEvent(el);
    if (pos && (pos[0] !== sel.fr || pos[1] !== sel.fc)) {
      const m = merges.get(KEY(pos[0], pos[1]));
      sel = { ...sel, fr: m ? m[2] : pos[0], fc: m ? m[3] : pos[1] };
      updateSelection(false);
    }
  } else if (dragging === 'cols') {
    const th = el.closest('th.ch') as HTMLElement | null;
    if (th && Number(th.dataset.c) !== sel.fc) {
      sel = { ...sel, fc: Number(th.dataset.c) };
      updateSelection(false);
    }
  } else if (dragging === 'rows') {
    const th = el.closest('th.rh') as HTMLElement | null;
    if (th && Number(th.dataset.r) !== sel.fr) {
      sel = { ...sel, fr: Number(th.dataset.r) };
      updateSelection(false);
    }
  }
});

document.addEventListener('mouseup', () => {
  dragging = null;
  if (resizing) {
    const { kind, index, el } = resizing;
    resizing = null;
    document.body.classList.remove('resizing-col', 'resizing-row');
    const size = Number(el.dataset.size);
    if (!Number.isFinite(size) || !el.dataset.size) return;
    delete el.dataset.size;
    const rg = normSel();
    if (kind === 'col') {
      const cols = isFullCols(rg) && index >= rg.c1 && index <= rg.c2 ? range(rg.c1, Math.min(rg.c2, maxCol)) : [index];
      post({ type: 'colWidth', sheet: cur, cols, px: size });
    } else {
      const rows = isFullRows(rg) && index >= rg.r1 && index <= rg.r2 ? range(rg.r1, Math.min(rg.r2, maxRow)) : [index];
      post({ type: 'rowHeight', sheet: cur, rows, px: size });
    }
  }
});

function range(a: number, b: number): number[] {
  const out: number[] = [];
  for (let i = a; i <= b; i++) out.push(i);
  return out;
}

table.addEventListener('dblclick', (e) => {
  const t = e.target as HTMLElement;
  if (t.classList.contains('rz')) {
    autofitColumn(Number(t.dataset.c));
    return;
  }
  const pos = cellFromEvent(t);
  if (pos) beginEdit('edit');
});

const measureCtx = document.createElement('canvas').getContext('2d')!;
function autofitColumn(c: number): void {
  let w = 0;
  const def = meta?.defaultFont ?? 'Calibri';
  for (const cell of cells.values()) {
    if (cell.c !== c || cell.m || !cell.t) continue;
    if (merges.has(KEY(cell.r, cell.c))) continue;
    measureCtx.font = cssFont(cell.s, def);
    const lines = cell.s?.wr ? cell.t.split('\n') : [cell.t];
    for (const line of lines) w = Math.max(w, measureCtx.measureText(line).width);
  }
  const px = Math.ceil(w + 10);
  post({ type: 'colWidth', sheet: cur, cols: [c], px: w ? Math.max(px, 20) : sh?.defaultColW ?? 64 });
}

// =====================================================================================
// Context menus
// =====================================================================================

interface MenuItem {
  label?: string;
  hint?: string;
  action?: () => void;
  disabled?: boolean;
  sep?: true;
}

function openMenu(x: number, y: number, items: MenuItem[]): void {
  menuEl.innerHTML = items
    .map((it, i) =>
      it.sep
        ? '<div class="msep"></div>'
        : `<div class="mi${it.disabled ? ' disabled' : ''}" data-i="${i}"><span>${esc(it.label ?? '')}</span><span class="hint">${esc(it.hint ?? '')}</span></div>`,
    )
    .join('');
  menuEl.hidden = false;
  const w = menuEl.offsetWidth;
  const h = menuEl.offsetHeight;
  menuEl.style.left = `${Math.min(x, window.innerWidth - w - 4)}px`;
  menuEl.style.top = `${Math.min(y, window.innerHeight - h - 4)}px`;
  menuEl.onclick = (e) => {
    const el = (e.target as HTMLElement).closest('.mi') as HTMLElement | null;
    if (!el || el.classList.contains('disabled')) return;
    const it = items[Number(el.dataset.i)];
    closeMenu();
    it.action?.();
    focusGrid();
  };
}

function closeMenu(): void {
  if (!menuEl.hidden) {
    menuEl.hidden = true;
    menuEl.innerHTML = '';
  }
}

document.addEventListener('mousedown', (e) => {
  if (!menuEl.contains(e.target as Node)) closeMenu();
});

function structure(op: StructureOp, at: number, count: number): void {
  if (sh?.hasTables && (op === 'insertCols' || op === 'deleteCols')) {
    // also checked on the extension side
  }
  post({ type: 'structure', sheet: cur, op, at, count });
}

table.addEventListener('contextmenu', (e) => {
  e.preventDefault();
  if (!sh) return;
  const rg = normSel();
  const nRows = Math.min(rg.r2, maxRow) - rg.r1 + 1;
  const nCols = Math.min(rg.c2, maxCol) - rg.c1 + 1;
  const rowsLabel = nRows > 1 ? `${nRows} rows` : 'row';
  const colsLabel = nCols > 1 ? `${nCols} columns` : 'column';
  const active = cells.get(KEY(...(mergeAt(sel.ar, sel.ac).slice(0, 2) as [number, number])));
  const hasMerge = sh.merges.some(([a, b, c, d]) => !(c < rg.r1 || a > rg.r2 || d < rg.c1 || b > rg.c2));
  const fullC = isFullCols(rg);
  const fullR = isFullRows(rg);
  const items: MenuItem[] = [
    { label: 'Cut', hint: 'Ctrl+X', action: () => document.execCommand('cut') },
    { label: 'Copy', hint: 'Ctrl+C', action: () => document.execCommand('copy') },
    {
      label: 'Paste',
      hint: 'Ctrl+V',
      action: () => {
        navigator.clipboard
          ?.readText()
          .then(pasteText)
          .catch(() => toast('Use Ctrl+V to paste.'));
      },
    },
    { sep: true },
  ];
  if (!fullC) {
    items.push(
      { label: `Insert ${rowsLabel} above`, action: () => structure('insertRows', rg.r1, nRows) },
      { label: `Insert ${rowsLabel} below`, action: () => structure('insertRows', Math.min(rg.r2, maxRow) + 1, nRows) },
      { label: `Delete ${nRows > 1 ? `${nRows} rows` : 'row'}`, action: () => structure('deleteRows', rg.r1, nRows) },
    );
  }
  if (!fullR) {
    items.push(
      { label: `Insert ${colsLabel} to the left`, action: () => structure('insertCols', rg.c1, nCols) },
      { label: `Insert ${colsLabel} to the right`, action: () => structure('insertCols', Math.min(rg.c2, maxCol) + 1, nCols) },
      { label: `Delete ${nCols > 1 ? `${nCols} columns` : 'column'}`, action: () => structure('deleteCols', rg.c1, nCols) },
    );
  }
  items.push(
    { sep: true },
    { label: 'Clear contents', hint: 'Delete', action: () => post({ type: 'clear', sheet: cur, ranges: [opRange()], what: 'contents' }) },
    { label: 'Clear formats', action: () => post({ type: 'clear', sheet: cur, ranges: [opRange()], what: 'formats' }) },
    { label: 'Clear all', action: () => post({ type: 'clear', sheet: cur, ranges: [opRange()], what: 'all' }) },
    { sep: true },
    hasMerge
      ? { label: 'Unmerge cells', action: () => post({ type: 'merge', sheet: cur, range: opRange(), unmerge: true }) }
      : {
          label: 'Merge cells',
          disabled: rg.r1 === rg.r2 && rg.c1 === rg.c2,
          action: () => post({ type: 'merge', sheet: cur, range: opRange(), unmerge: false }),
        },
  );
  if (fullC && nCols === 1) {
    items.push({ label: 'AutoFit column width', action: () => autofitColumn(rg.c1) });
  }
  if (active?.link) items.push({ sep: true }, { label: 'Open hyperlink', hint: 'Ctrl+click', action: () => post({ type: 'openLink', url: active.link! }) });
  if (active?.dv) items.push({ label: 'Pick from list…', hint: 'Alt+Down', action: openList });
  openMenu(e.clientX, e.clientY, items);
});

// Data validation drop-down lists
function openList(): void {
  const m = mergeAt(sel.ar, sel.ac);
  const cell = cells.get(KEY(m[0], m[1]));
  if (!cell?.dv) return;
  post({ type: 'getList', sheet: cur, r: m[0], c: m[1] });
}

dvBtn.addEventListener('mousedown', (e) => {
  e.preventDefault();
  e.stopPropagation();
  openList();
});

function showList(r: number, c: number, values: string[]): void {
  if (r !== mergeAt(sel.ar, sel.ac)[0] || c !== mergeAt(sel.ar, sel.ac)[1]) return;
  const rect = dvBtn.getBoundingClientRect();
  const cellRect = selActiveEl.getBoundingClientRect();
  if (!values.length) {
    toast('Validation list is empty or not supported.');
    return;
  }
  openMenu(
    cellRect.left,
    Math.max(rect.bottom, cellRect.bottom) + 2,
    values.map((v) => ({ label: v, action: () => post({ type: 'setCells', sheet: cur, items: [{ r, c, input: v }] }) })),
  );
  menuEl.classList.add('list');
  setTimeout(() => menuEl.classList.remove('list'), 0);
}

// =====================================================================================
// Sheet tabs
// =====================================================================================

function renderTabs(): void {
  if (!meta) return;
  let html = '<button class="tab-add" title="New sheet">+</button>';
  meta.sheets.forEach((s, i) => {
    html += `<div class="tab${i === cur ? ' active' : ''}${s.hidden ? ' hidden-sheet' : ''}" data-i="${i}" title="${esc(s.name)}${s.hidden ? ' (hidden sheet)' : ''}"${
      s.tab ? ` style="--tab:${s.tab}"` : ''
    }>${esc(s.name)}</div>`;
  });
  tabsEl.innerHTML = html;
  tabsEl.querySelector('.tab.active')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

tabsEl.addEventListener('mousedown', (e) => {
  const t = e.target as HTMLElement;
  if (t.closest('input')) return;
  e.preventDefault();
  commitEdit();
  if (t.closest('.tab-add')) {
    post({ type: 'addSheet' });
    return;
  }
  const tab = t.closest('.tab') as HTMLElement | null;
  if (tab && e.button === 0) switchSheet(Number(tab.dataset.i));
});
tabsEl.addEventListener('dblclick', (e) => {
  const tab = (e.target as HTMLElement).closest('.tab') as HTMLElement | null;
  if (tab) renameTab(Number(tab.dataset.i));
});
tabsEl.addEventListener('contextmenu', (e) => {
  e.preventDefault();
  const tab = (e.target as HTMLElement).closest('.tab') as HTMLElement | null;
  const i = tab ? Number(tab.dataset.i) : -1;
  const items: MenuItem[] = [{ label: 'Insert sheet', action: () => post({ type: 'addSheet' }) }];
  if (i >= 0) {
    items.push(
      { label: 'Rename…', action: () => renameTab(i) },
      { label: 'Delete', action: () => post({ type: 'deleteSheet', index: i }) },
    );
  }
  openMenu(e.clientX, e.clientY - 10 - items.length * 26, items);
});

function renameTab(i: number): void {
  const tab = tabsEl.querySelector(`.tab[data-i="${i}"]`) as HTMLElement | null;
  if (!tab || !meta) return;
  const input = document.createElement('input');
  input.className = 'tab-input';
  input.value = meta.sheets[i].name;
  input.maxLength = 31;
  tab.textContent = '';
  tab.appendChild(input);
  input.focus();
  input.select();
  let done = false;
  const finish = (ok: boolean) => {
    if (done) return;
    done = true;
    const name = input.value.trim();
    if (ok && name && meta && name !== meta.sheets[i].name) post({ type: 'renameSheet', index: i, name });
    renderTabs();
    focusGrid();
  };
  input.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Enter') finish(true);
    else if (e.key === 'Escape') finish(false);
  });
  input.addEventListener('blur', () => finish(true));
}

function switchSheet(i: number): void {
  if (!meta || i < 0 || i >= meta.sheets.length) return;
  if (i === cur && sh) return;
  persist();
  cur = i;
  renderTabs();
  loadingEl.hidden = false;
  post({ type: 'getSheet', index: i });
}

// =====================================================================================
// Messages from the extension
// =====================================================================================

function loadSheet(data: SheetModel): void {
  const same = sh?.index === data.index && sh?.name === data.name;
  sh = data;
  cells = new Map();
  for (const cell of data.cells) cells.set(KEY(cell.r, cell.c), cell);
  merges = new Map();
  for (const m of data.merges) {
    const rr = m[2] - m[0] + 1;
    const cc = m[3] - m[1] + 1;
    if (rr * cc > 200000) continue;
    for (let r = m[0]; r <= m[2]; r++) for (let c = m[1]; c <= m[3]; c++) merges.set(KEY(r, c), m);
  }
  if (!same) {
    const s = saved.sels[data.name];
    sel = s ?? { ar: 1, ac: 1, fr: 1, fc: 1 };
    zoom = saved.zooms[data.name] ?? data.zoom / 100;
    showGrid = data.grid;
    maxRow = Math.max(data.rowCount + 30, 60, sel.fr + 10, sel.ar + 10);
    maxCol = Math.max(data.colCount + 6, 26, sel.ac + 3);
    maxRow = Math.min(maxRow, MAX_ROW);
    maxCol = Math.min(maxCol, MAX_COL);
    scroller.scrollTop = 0;
    scroller.scrollLeft = 0;
  } else {
    maxRow = Math.min(MAX_ROW, Math.max(maxRow, data.rowCount + 30));
    maxCol = Math.min(MAX_COL, Math.max(maxCol, data.colCount + 6));
  }
  loadingEl.hidden = true;
  renderAll(same);
  if (!same) scrollIntoView(sel.ar, sel.ac);
  focusGrid();
}

function applyMeta(m: WorkbookMeta): void {
  meta = m;
  document.documentElement.style.setProperty('--default-font', fontFamily(m.defaultFont) ?? 'Calibri, sans-serif');
  if (m.warnings.length && !bannerEl.dataset.closed) {
    bannerEl.hidden = false;
    bannerEl.innerHTML = `<span>⚠ This workbook contains features this editor cannot preserve: <b>${esc(
      m.warnings.join(', '),
    )}</b>. They will be lost if you save from this editor.</span><button title="Close">✕</button>`;
    bannerEl.querySelector('button')!.onclick = () => {
      bannerEl.hidden = true;
      bannerEl.dataset.closed = '1';
    };
  } else if (!m.warnings.length) bannerEl.hidden = true;
}

function firstVisibleSheet(): number {
  const i = meta?.sheets.findIndex((s) => !s.hidden) ?? 0;
  return i < 0 ? 0 : i;
}

let toastTimer: number | undefined;
function toast(message: string): void {
  toastEl.textContent = message;
  toastEl.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => (toastEl.hidden = true), 5000);
}

window.addEventListener('message', (ev: MessageEvent<ToWebview>) => {
  const msg = ev.data;
  switch (msg.type) {
    case 'init': {
      settings = msg.settings;
      document.body.classList.toggle('grid-vscode', settings.gridTheme === 'vscode');
      applyMeta(msg.meta);
      if (cur >= msg.meta.sheets.length || (msg.meta.sheets[cur]?.hidden && !sh)) cur = firstVisibleSheet();
      sh = null;
      renderTabs();
      post({ type: 'getSheet', index: cur });
      break;
    }
    case 'sheet':
      if (msg.index === cur) loadSheet(msg.data);
      break;
    case 'cells': {
      if (msg.sheet !== cur || !sh) break;
      const list: [number, number][] = [];
      let grow = false;
      for (const cell of msg.cells) {
        cells.set(KEY(cell.r, cell.c), cell);
        list.push([cell.r, cell.c]);
        if (cell.r > sh.rowCount) sh.rowCount = cell.r;
        if (cell.c > sh.colCount) sh.colCount = cell.c;
        if (cell.c > maxCol) grow = true;
      }
      for (const [r, c] of msg.empty) {
        cells.delete(KEY(r, c));
        list.push([r, c]);
      }
      if (grow) {
        maxCol = Math.min(MAX_COL, sh.colCount + 6);
        renderAll(true);
      } else {
        rerenderCells(list);
        updateSelection(false);
      }
      break;
    }
    case 'reset': {
      applyMeta(msg.meta);
      if (msg.activeSheet !== undefined) cur = msg.activeSheet;
      if (cur >= msg.meta.sheets.length) cur = msg.meta.sheets.length - 1;
      if (sh && msg.meta.sheets[cur]?.name !== sh.name) sh = null;
      renderTabs();
      post({ type: 'getSheet', index: cur });
      break;
    }
    case 'list':
      showList(msg.r, msg.c, msg.values);
      break;
    case 'toast':
      toast(msg.message);
      break;
  }
});

// Scrolling: progressive rendering of rows / columns
let scrollRaf = 0;
scroller.addEventListener('scroll', () => {
  if (scrollRaf) return;
  scrollRaf = requestAnimationFrame(() => {
    scrollRaf = 0;
    if (!sh) return;
    updateOverlayClip();
    const nearBottom = scroller.scrollTop + scroller.clientHeight > scroller.scrollHeight - 400;
    if (nearBottom) {
      if (renderedRows < maxRow) renderRowsUntil(renderedRows + 200);
      else if (maxRow < MAX_ROW) {
        maxRow = Math.min(MAX_ROW, maxRow + 200);
        const digits = String(maxRow).length !== String(renderedRows).length;
        if (digits) renderAll(true);
        else renderRowsUntil(maxRow);
      }
    }
    const nearRight = scroller.scrollLeft + scroller.clientWidth > scroller.scrollWidth - 150;
    if (nearRight && maxCol < MAX_COL) {
      maxCol = Math.min(MAX_COL, maxCol + 10);
      renderAll(true);
    }
  });
});

window.addEventListener('focus', () => {
  if (document.activeElement === document.body) focusGrid();
});
window.addEventListener('resize', () => updateOverlayClip());
window.addEventListener('error', (e) => post({ type: 'error', message: String(e.message) }));

post({ type: 'ready' });
