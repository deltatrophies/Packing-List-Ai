import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import { makeExcel, makePdf, parseFinishedWorkbook, planPages } from '../exports.js';

const list = {
  customer: 'Test Customer', packing_date: '08-10-2026', private_mark: 'VS', transport: '', warnings: [],
  boxes: Array.from({ length: 7 }, (_, i) => ({ number: i + 1, items: [
    { code: `PC-${400 + i}`, size: 'M', quantity: 4, note: '' },
    { code: `FC-${20 + i}`, size: 'S', quantity: 2, note: 'ONLY CUP' },
    { code: `W-${2500 + i}`, size: 'L', quantity: 10, note: '' },
  ] })),
};

test('print planner keeps every box in exactly one group', () => {
  const pages = planPages(list.boxes);
  assert.deepEqual(pages.flat(2).map(x => x.number), [1,2,3,4,5,6,7]);
  assert.ok(pages.every(page => page.length <= 2));
});

test('Excel prints five rows per short box and imports its actual entries', async () => {
  const blob = await makeExcel(list);
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(await blob.arrayBuffer());
  const sheet = book.worksheets[0];
  assert.equal(sheet.getCell('A5').value, 'BOX No.');
  assert.equal(sheet.getCell('C5').value, 1);
  assert.equal(sheet.getCell('A5').font.size, 28);
  assert.equal(sheet.getCell('A6').font.size, 14);
  assert.equal(sheet.getCell('A7').font.size, 14);
  assert.equal(sheet.getCell('B8').font.size, 12);
  assert.equal(sheet.getCell('B8').value, '  PC-400 M');
  assert.equal(sheet.getCell('C12').value, null);
  assert.equal(sheet.getCell('A14').value, 'BOX No.');
  assert.ok(sheet.getCell('B8').border.left);
  assert.equal(sheet.pageSetup.printArea, 'A5:G21');
  assert.equal(sheet.pageSetup.scale, 125);
  const imported = await parseFinishedWorkbook(new File([blob], 'Test Customer 08-10-2026.xlsx'));
  assert.equal(imported.boxes.length, 7);
  assert.equal(imported.boxes[0].items.length, 3);
});

test('an odd final box does not create a fictitious box', async () => {
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(await (await makeExcel({ ...list, boxes: list.boxes.slice(0, 3) })).arrayBuffer());
  assert.equal(book.worksheets[0].getCell('A14').value, 'BOX No.');
  assert.equal(book.worksheets[0].getCell('E14').value, null);
});

test('a long box remains on one Excel print page', async () => {
  const longList = { ...list, boxes: [{ number: 1, items: Array.from({ length: 35 }, (_, i) => ({ code: `PC-${i + 1}`, size: 'M', quantity: 1 })) }] };
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(await (await makeExcel(longList)).arrayBuffer());
  assert.equal(book.worksheets.length, 1);
  assert.equal(book.worksheets[0].pageSetup.paperSize, 9);
  assert.equal(book.worksheets[0].pageSetup.fitToHeight, 1);
  assert.equal(book.worksheets[0].getCell('B42').value, '  PC-35 M');
});

test('18-entry box pair exports on A4 in Excel and PDF', async () => {
  const eighteen = { ...list, boxes: [
    { number: 5, items: Array.from({ length: 18 }, (_, i) => ({ code: `PC-${i + 1}`, size: 'M', quantity: 2 })) },
    { number: 6, items: list.boxes[0].items },
  ] };
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(await (await makeExcel(eighteen)).arrayBuffer());
  assert.equal(book.worksheets.length, 1);
  assert.equal(book.worksheets[0].pageSetup.paperSize, 9);
  assert.equal(book.worksheets[0].pageSetup.scale, 100);
  assert.equal(book.worksheets[0].getCell('B25').value, '  PC-18 M');
  const pdf = new TextDecoder().decode(await makePdf(eighteen).arrayBuffer());
  assert.match(pdf, /\/MediaBox \[0 0 841\.8\d* 595\.2\d*/);
  assert.doesNotMatch(pdf, /\/MediaBox \[0 0 1190/);
});

test('PDF export produces a nonempty document', async () => {
  const blob = makePdf(list);
  assert.equal(blob.type, 'application/pdf');
  assert.ok(blob.size > 1000);
  const bytes = new Uint8Array(await blob.arrayBuffer());
  assert.equal(new TextDecoder().decode(bytes.subarray(0, 5)), '%PDF-');
});

test('existing finished workbook remains readable', async () => {
  const bytes = await readFile(new URL('../../sample-output/Sharda%20Sports%20Borsad.xlsx', import.meta.url));
  const imported = await parseFinishedWorkbook(new File([bytes], 'Sharda sports borsad 08-10-2026.xlsx'));
  assert.ok(imported.boxes.length >= 1);
});
