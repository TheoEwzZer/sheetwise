import { test } from 'node:test';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { Serializer } from '../src/serialize.ts';
import { Recalc, cellKey } from '../src/recalc.ts';
import { spliceSheet, unshareFormulas, renameSheet, deleteSheet } from '../src/structure.ts';
import { DEFAULT_THEME } from '../src/colors.ts';

const ctx = { locale: 'fr' as const, date1904: false };

async function roundtrip(wb: ExcelJS.Workbook): Promise<ExcelJS.Workbook> {
  const buf = await wb.xlsx.writeBuffer();
  const out = new ExcelJS.Workbook();
  await out.xlsx.load(buf);
  return out;
}

function sample(): ExcelJS.Workbook {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  ws.getColumn(1).width = 20;
  ws.getCell('A1').value = 'Title';
  ws.getCell('A1').style = {
    font: { bold: true, size: 14, color: { argb: 'FFFF0000' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { theme: 4, tint: 0.4 } as never },
    border: { bottom: { style: 'medium', color: { argb: 'FF0000FF' } } },
    alignment: { horizontal: 'center', wrapText: true },
  };
  ws.getCell('A2').value = 10;
  ws.getCell('A3').value = 20;
  ws.getCell('A4').value = { formula: 'SUM(A2:A3)', result: 30 } as ExcelJS.CellFormulaValue;
  ws.getCell('B2').value = new Date(Date.UTC(2026, 10, 19));
  ws.getCell('B2').numFmt = 'dd/mm/yyyy';
  ws.getCell('B3').value = 0.256;
  ws.getCell('B3').numFmt = '0.0%';
  ws.mergeCells('C1:D2');
  ws.getCell('C1').value = 'Merged';
  const ws2 = wb.addWorksheet('Summary');
  ws2.getCell('A1').value = { formula: "'Data'!A4*2", result: 60 } as ExcelJS.CellFormulaValue;
  ws2.getCell('A2').value = { formula: "SUM('Data'!A2:A10)" } as ExcelJS.CellFormulaValue;
  return wb;
}

test('serializes values, formats, styles and merges', async () => {
  const wb = await roundtrip(sample());
  const s = new Serializer(DEFAULT_THEME, ctx);
  const model = s.sheet(wb.worksheets[0], 0);
  const byAddr = new Map(model.cells.map((c) => [`${c.r}:${c.c}`, c]));
  const a1 = byAddr.get('1:1')!;
  assert.equal(a1.t, 'Title');
  assert.equal(a1.s?.b, 1);
  assert.equal(a1.s?.fs, 14);
  assert.equal(a1.s?.fc, '#FF0000');
  assert.equal(a1.s?.bg, '#8FAADC');
  assert.equal(a1.s?.bd?.b?.s, '2px solid');
  assert.equal(a1.s?.ha, 'center');
  assert.equal(byAddr.get('4:1')!.t, '30');
  assert.equal(byAddr.get('4:1')!.e, '=SUM(A2:A3)');
  assert.equal(byAddr.get('2:2')!.t, '19/11/2026');
  assert.equal(byAddr.get('2:2')!.k, 'd');
  assert.equal(byAddr.get('3:2')!.t, '25,6%');
  assert.deepEqual(model.merges, [[1, 3, 2, 4]]);
  assert.equal(model.cols[1].w, 145);
});

test('recalculates dependent formulas and formulas without a result', async () => {
  const wb = await roundtrip(sample());
  unshareFormulas(wb);
  const rc = new Recalc(wb, false);
  const missing = rc.computeMissing();
  assert.ok(missing.includes(cellKey(1, 2, 1)));
  assert.equal((wb.worksheets[1].getCell('A2').value as ExcelJS.CellFormulaValue).result, 60);

  wb.worksheets[0].getCell('A2').value = 100;
  const changed = rc.recalc([cellKey(0, 2, 1)]);
  assert.deepEqual(new Set(changed), new Set([cellKey(0, 4, 1), cellKey(1, 1, 1), cellKey(1, 2, 1)]));
  assert.equal((wb.worksheets[0].getCell('A4').value as ExcelJS.CellFormulaValue).result, 120);
  assert.equal((wb.worksheets[1].getCell('A1').value as ExcelJS.CellFormulaValue).result, 240);
  assert.equal((wb.worksheets[1].getCell('A2').value as ExcelJS.CellFormulaValue).result, 240);

  wb.worksheets[0].getCell('A5').value = { formula: 'A4/0' } as ExcelJS.CellFormulaValue;
  rc.reindex([cellKey(0, 5, 1)]);
  rc.recalc([cellKey(0, 5, 1)]);
  assert.deepEqual((wb.worksheets[0].getCell('A5').value as ExcelJS.CellFormulaValue).result, { error: '#DIV/0!' });
});

test('keeps the Excel value for a function unknown to the engine', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 1;
  ws.getCell('B1').value = { formula: 'UNKNOWNFUNCTION(A1)', result: 42 } as ExcelJS.CellFormulaValue;
  const rc = new Recalc(wb, false);
  ws.getCell('A1').value = 2;
  rc.recalc([cellKey(0, 1, 1)]);
  assert.equal((ws.getCell('B1').value as ExcelJS.CellFormulaValue).result, 42);
});

test('inserts and deletes rows, rewriting formulas and merges', async () => {
  const wb = await roundtrip(sample());
  unshareFormulas(wb);
  spliceSheet(wb, 0, 'row', 2, 2);
  const ws = wb.worksheets[0];
  assert.equal(ws.getCell('A4').value, 10);
  assert.equal(ws.getCell('A6').formula, 'SUM(A4:A5)');
  assert.equal(wb.worksheets[1].getCell('A1').formula, "'Data'!A6*2");
  assert.equal(wb.worksheets[1].getCell('A2').formula, "SUM('Data'!A4:A12)");
  // the C1:D2 merge grows because the insertion is inside it
  const merges = Object.values((ws as unknown as { _merges: Record<string, { range: string }> })._merges).map((m) => m.range);
  assert.deepEqual(merges, ['C1:D4']);
  // the style of the row above is copied
  assert.equal(ws.getCell('A2').style.font?.bold, true);

  spliceSheet(wb, 0, 'row', 4, -1);
  assert.equal(ws.getCell('A4').value, 20);
  assert.equal(ws.getCell('A5').formula, 'SUM(A4:A4)');

  const out = await roundtrip(wb);
  assert.equal(out.worksheets[0].getCell('A5').formula, 'SUM(A4:A4)');
  assert.equal(out.worksheets[0].getCell('A1').style.fill?.type, 'pattern');
});

test('inserts columns with widths and validations', async () => {
  const wb = await roundtrip(sample());
  const ws = wb.worksheets[0];
  (ws as unknown as { dataValidations: { add(a: string, v: object): void } }).dataValidations.add('B5', {
    type: 'list',
    allowBlank: true,
    formulae: ['$A$2:$A$3'],
  });
  spliceSheet(wb, 0, 'col', 1, 1);
  assert.equal(ws.getCell('B1').value, 'Title');
  assert.equal(ws.getColumn(2).width, 20);
  assert.equal(ws.getCell('B4').formula, 'SUM(B2:B3)');
  const dv = (ws as unknown as { dataValidations: { model: Record<string, { formulae: string[] }> } }).dataValidations.model;
  assert.deepEqual(Object.keys(dv), ['C5']);
  assert.deepEqual(dv.C5.formulae, ['$B$2:$B$3']);
  const out = await roundtrip(wb);
  assert.equal(out.worksheets[0].getCell('B1').value, 'Title');
});

test('renames and deletes a sheet', async () => {
  const wb = await roundtrip(sample());
  renameSheet(wb, 0, 'Base 2026');
  assert.equal(wb.worksheets[1].getCell('A1').formula, "'Base 2026'!A4*2");
  deleteSheet(wb, 0);
  assert.equal(wb.worksheets.length, 1);
  assert.equal(wb.worksheets[0].getCell('A1').formula, '#REF!*2');
});

test('converts shared formulas to explicit formulas', async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 1;
  ws.getCell('A2').value = 2;
  ws.getCell('B1').value = { formula: 'A1*2', result: 2, shareType: 'shared', ref: 'B1:B2' } as unknown as ExcelJS.CellValue;
  ws.getCell('B2').value = { sharedFormula: 'B1', result: 4 } as unknown as ExcelJS.CellValue;
  const out = await roundtrip(wb);
  unshareFormulas(out);
  assert.equal(out.worksheets[0].getCell('B2').formula, 'A2*2');
  assert.equal((out.worksheets[0].getCell('B2').value as { sharedFormula?: string }).sharedFormula, undefined);
});

test('added functions: COUNTIFS, SUMIFS, AVERAGEIFS, TEXTJOIN, XLOOKUP, SWITCH, empty cell = ""', () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  const rows = [
    [1, 5, 'Yes'],
    [2, 3, 'No'],
    [1, 7, 'Yes'],
    [1, 2, 'No'],
  ];
  rows.forEach((r, i) => r.forEach((v, j) => (ws.getCell(i + 1, j + 1).value = v)));
  const formulas: Record<string, string> = {
    E1: 'COUNTIFS(A1:A4,1,B1:B4,">=4")',
    E2: 'SUMIFS(B1:B4,A1:A4,1)',
    E3: 'AVERAGEIFS(B1:B4,A1:A4,1,C1:C4,"Yes")',
    E4: 'TEXTJOIN(";",TRUE,C1:C4)',
    E5: 'XLOOKUP(7,B1:B4,C1:C4)',
    E6: 'SWITCH(A2,1,"one",2,"two")',
    E7: 'IF(Z99="","empty","full")',
    E8: 'IFERROR(AVERAGEIFS(B1:B4,A1:A4,9),"")',
    E9: 'MAXIFS(B1:B4,C1:C4,"No")',
  };
  for (const [a, f] of Object.entries(formulas)) ws.getCell(a).value = { formula: f } as ExcelJS.CellFormulaValue;
  new Recalc(wb, false).computeMissing();
  const res = (a: string) => ws.getCell(a).result;
  assert.equal(res('E1'), 2);
  assert.equal(res('E2'), 14);
  assert.equal(res('E3'), 6);
  assert.equal(res('E4'), 'Yes;No;Yes;No');
  assert.equal(res('E5'), 'Yes');
  assert.equal(res('E6'), 'two');
  assert.equal(res('E7'), 'empty');
  assert.equal(res('E8'), '');
  assert.equal(res('E9'), 3);
});
