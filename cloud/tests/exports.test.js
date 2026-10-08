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
  assert.equal(book.worksheets[0].getCell('B5').value, 'BOX No.');
  assert.equal(book.worksheets[0].getCell('B8').value, 'PC-400 M');
  assert.equal(book.worksheets[0].getCell('C12').value, null);
  assert.equal(book.worksheets[0].pageSetup.fitToHeight, 1);
  const imported = await parseFinishedWorkbook(new File([blob], 'Test Customer 08-10-2026.xlsx'));
  assert.equal(imported.boxes.length, 7);
  assert.equal(imported.boxes[0].items.length, 3);
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
