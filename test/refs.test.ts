import { test } from 'node:test';
import assert from 'node:assert/strict';
import { findRefs, offsetRef, spliceRef, transformRefs, parseRangeText, quoteSheet } from '../src/refs.ts';

const splice = (f: string, axis: 'row' | 'col', at: number, count: number, sheet?: string) =>
  transformRefs(f, (ref) => (sheet === undefined || ref.sheet === undefined || ref.sheet === sheet ? spliceRef(ref, axis, at, count) : undefined));

test('detects cell refs, ranges, whole columns and whole rows', () => {
  const refs = findRefs("SUM(A1:B3)+'My sheet'!$C$4+Sheet2!D:D+3:5+LOG10(2)+\"A1\"");
  assert.deepEqual(
    refs.map((r) => [r.kind, r.sheet ?? '', r.r1, r.c1, r.r2, r.c2]),
    [
      ['range', '', 1, 1, 3, 2],
      ['cell', 'My sheet', 4, 3, 4, 3],
      ['cols', 'Sheet2', 1, 4, 1048576, 4],
      ['rows', '', 3, 1, 5, 16384],
    ],
  );
});

test('ignores functions, names and strings', () => {
  assert.equal(findRefs('LOG10(5)+ATAN2(1,2)').length, 0);
  assert.equal(findRefs('"B2"&"C3"').length, 0);
  assert.equal(findRefs('Table1[@Col1]').length, 0);
  assert.equal(findRefs('[1]Sheet1!A1').length, 0);
});

test('row insertion', () => {
  assert.equal(splice('A1+A5+SUM(A2:A10)', 'row', 3, 2), 'A1+A7+SUM(A2:A12)');
  assert.equal(splice('$B$4*2', 'row', 4, 1), '$B$5*2');
  assert.equal(splice('A:A', 'row', 1, 5), 'A:A');
});

test('row deletion', () => {
  assert.equal(splice('A1+A5', 'row', 5, -1), 'A1+#REF!');
  assert.equal(splice('SUM(A2:A10)', 'row', 3, -2), 'SUM(A2:A8)');
  assert.equal(splice('SUM(A3:A4)', 'row', 3, -2), 'SUM(#REF!)');
  assert.equal(splice('SUM(A1:A3)', 'row', 2, -5), 'SUM(A1:A1)');
});

test('column insertion / deletion', () => {
  assert.equal(splice('B1+D1', 'col', 3, 1), 'B1+E1');
  assert.equal(splice('SUM(B:D)', 'col', 3, 1), 'SUM(B:E)');
  assert.equal(splice('C1', 'col', 3, -1), '#REF!');
  assert.equal(splice('SUM(A1:E1)', 'col', 2, -2), 'SUM(A1:C1)');
});

test('only affects the target sheet', () => {
  assert.equal(splice("Sheet2!A5+'Other'!A5", 'row', 1, 1, 'Sheet2'), "Sheet2!A6+'Other'!A5");
});

test('relative copy', () => {
  const f = transformRefs('A1+$B$2+B$3+$C4', (ref) => offsetRef(ref, 2, 1));
  assert.equal(f, 'B3+$B$2+C$3+$C6');
  assert.equal(transformRefs('A1', (ref) => offsetRef(ref, -1, 0)), '#REF!');
});

test('parses ranges and quotes sheet names', () => {
  const r = parseRangeText("'Café 2024'!$A$1:$B$20");
  assert.equal(r?.sheet, 'Café 2024');
  assert.equal(r?.r2, 20);
  assert.equal(quoteSheet('Sheet1'), 'Sheet1');
  assert.equal(quoteSheet("Bob's"), "'Bob''s'");
  assert.equal(quoteSheet('AB12'), "'AB12'");
});
