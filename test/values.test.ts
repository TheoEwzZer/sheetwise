import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseInput, normalizeFormulaInput } from '../src/parseInput.ts';
import { formatNumber, numberEditText, dateToSerial, formatColor, normalizeFormat } from '../src/numfmt.ts';
import { applyTint, parseTheme, resolveColor, DEFAULT_THEME } from '../src/colors.ts';

const fr = { locale: 'fr' as const };
const en = { locale: 'en' as const };

test('input: French and English numbers', () => {
  assert.equal(parseInput('3,5', fr).value, 3.5);
  assert.equal(parseInput('1 234,5', fr).value, 1234.5);
  assert.equal(parseInput('1.5', fr).value, 1.5);
  assert.equal(parseInput('-12', fr).value, -12);
  assert.equal(parseInput('1,234.5', en).value, 1234.5);
  assert.equal(parseInput('3,5', en).value, '3,5');
});

test('input: percentages, currency, booleans, errors, forced text', () => {
  assert.deepEqual(parseInput('12%', fr), { value: 0.12, numFmt: '0%' });
  assert.deepEqual(parseInput('12,5 %', fr), { value: 0.125, numFmt: '0.0%' });
  assert.deepEqual(parseInput('12,50 €', fr), { value: 12.5, numFmt: '#,##0.00 "€"' });
  assert.equal(parseInput('vrai', fr).value, true);
  assert.equal(parseInput('FALSE', fr).value, false);
  assert.deepEqual(parseInput('#N/A', fr).value, { error: '#N/A' });
  assert.equal(parseInput("'0123", fr).value, '0123');
  assert.equal(parseInput('0123', { locale: 'fr', cellNumFmt: '@' }).value, '0123');
  assert.equal(parseInput('15', { locale: 'fr', cellNumFmt: '0%' }).value, 0.15);
  assert.equal(parseInput('', fr).value, null);
  assert.equal(parseInput('hello', fr).value, 'hello');
});

test('input: dates and times', () => {
  const d = parseInput('19/11/2026', fr);
  assert.ok(d.value instanceof Date);
  assert.equal((d.value as Date).toISOString(), '2026-11-19T00:00:00.000Z');
  assert.equal(d.numFmt, 'dd/mm/yyyy');
  assert.equal((parseInput('11/19/2026', en).value as Date).toISOString(), '2026-11-19T00:00:00.000Z');
  assert.equal((parseInput('2026-11-19', fr).value as Date).getUTCDate(), 19);
  assert.equal(parseInput('31/02/2026', fr).value, '31/02/2026');
  const t = parseInput('14:30', fr);
  assert.equal(t.numFmt, 'hh:mm');
  assert.ok(Math.abs(dateToSerial(t.value as Date) - 14.5 / 24) < 1e-9);
});

test('input: formulas', () => {
  assert.deepEqual(parseInput('=SUM(A1:A3)', fr).value, { formula: 'SUM(A1:A3)' });
  assert.equal(normalizeFormulaInput('SUM(A1;A2;0,5)'), 'SUM(A1,A2,0.5)');
  assert.equal(normalizeFormulaInput('IF(A1>1;"a;b";2)'), 'IF(A1>1,"a;b",2)');
  assert.equal(normalizeFormulaInput('SUM(A1,A2)'), 'SUM(A1,A2)');
  assert.equal(normalizeFormulaInput('SOMME(B3;B4)'), 'SUM(B3,B4)');
  assert.equal(normalizeFormulaInput('SI(NB.SI(A:A;"x")>0;"SOMME(";moyenne(A1:A2))'), 'IF(COUNTIF(A:A,"x")>0,"SOMME(",AVERAGE(A1:A2))');
  assert.equal(normalizeFormulaInput("'SI'!A1+SIERREUR(1/0;0)"), "'SI'!A1+IFERROR(1/0,0)");
});

test('number display by locale', () => {
  const ctx = { locale: 'fr' as const, date1904: false };
  assert.equal(formatNumber(1234.567, '#,##0.00', ctx), '1 234,57');
  assert.equal(formatNumber(0.125, '0.0%', ctx), '12,5%');
  assert.equal(formatNumber(46345, 'mm-dd-yy', ctx), '19/11/2026');
  assert.equal(formatNumber(46345, 'd mmmm yyyy', ctx), '19 novembre 2026');
  assert.equal(formatNumber(46345, 'dddd', ctx), 'jeudi');
  assert.equal(formatNumber(3.5, undefined, ctx), '3,5');
  assert.equal(formatNumber(3.5, undefined, { locale: 'en', date1904: false }), '3.5');
  assert.equal(numberEditText(0.1 + 0.2, undefined, ctx), '0,3');
  assert.equal(numberEditText(0.15, '0%', ctx), '15%');
  assert.equal(normalizeFormat('m/d/yy "h":mm', 'fr'), 'dd/mm/yyyy hh:mm');
  assert.equal(formatColor('0.00;[Red]-0.00', -2), '#FF0000');
  assert.equal(formatColor('0.00;[Red]-0.00', 2), undefined);
});

test('colors: theme, tint, indexed palette', () => {
  assert.equal(resolveColor({ argb: 'FFE7E6E6' }, DEFAULT_THEME), '#E7E6E6');
  assert.equal(resolveColor({ theme: 4 }, DEFAULT_THEME), '#4472C4');
  assert.equal(resolveColor({ indexed: 10 }, DEFAULT_THEME), '#FF0000');
  assert.equal(resolveColor({ indexed: 64 }, DEFAULT_THEME), undefined);
  assert.equal(applyTint('000000', 0.5), '808080');
  assert.equal(applyTint('FFFFFF', -0.5), '808080');
  const theme = parseTheme(
    '<a:theme><a:themeElements><a:clrScheme name="x"><a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1><a:lt1><a:srgbClr val="FEFEFE"/></a:lt1><a:accent1><a:srgbClr val="ABCDEF"/></a:accent1></a:clrScheme><a:fontScheme><a:minorFont><a:latin typeface="Aptos Narrow"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>',
  );
  assert.equal(theme.colors[0], 'FEFEFE');
  assert.equal(theme.colors[1], '111111');
  assert.equal(theme.colors[4], 'ABCDEF');
  assert.equal(theme.minorFont, 'Aptos Narrow');
});
