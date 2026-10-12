import ExcelJS from 'exceljs';
import { jsPDF } from 'jspdf';

export const MIN_ROWS_PER_BOX = 5;
const ROW_PT = 22;
const BOX_TITLE_PT = 36;
const MARK_PT = 30.75;
const TABLE_HEADER_PT = 33;
const GROUP_GAP_PT = 15;
const GROUP_PT = BOX_TITLE_PT + MARK_PT + TABLE_HEADER_PT;
const A4_GROUP_BUDGET_PT = 500;
const A4_LARGE_PRINT_PT = 455;
const visibleRows = box => Math.max(MIN_ROWS_PER_BOX, box.items.length);
const groupHeight = pair => GROUP_PT + ROW_PT * Math.max(...pair.map(visibleRows)) + GROUP_GAP_PT;

export function validateList(data) {
  if (!data?.boxes?.length) throw new Error('Add at least one box before export');
  const used = new Set();
  for (const box of data.boxes) {
    if (!Number.isInteger(box.number) || box.number < 1 || used.has(box.number)) throw new Error('Box numbers must be unique positive numbers');
    used.add(box.number);
    for (const [i, item] of box.items.entries()) {
      if (!String(item.code || '').trim()) throw new Error(`Box ${box.number}, row ${i + 1}: item code is empty`);
      if (!Number.isInteger(item.quantity) || item.quantity < 1) throw new Error(`Box ${box.number}, row ${i + 1}: quantity must be at least 1`);
    }
  }
}
export function planPages(boxes) {
  const ordered = [...boxes].sort((a, b) => a.number - b.number);
  const pairs = [];
  for (let i = 0; i < ordered.length; i += 2) pairs.push(ordered.slice(i, i + 2));
  const pages = [];
  let page = [], height = 0;
  for (const pair of pairs) {
    const required = groupHeight(pair);
    if (page.length && (page.length === 2 || height + required > A4_GROUP_BUDGET_PT)) {
      pages.push(page); page = []; height = 0;
    }
    page.push(pair); height += required;
  }
  if (page.length) pages.push(page);
  return pages;
}
export const itemLabel = item => [String(item.code || '').trim().toUpperCase(), String(item.size || '').trim().toUpperCase()].filter(Boolean).join(' ') +
  (item.note ? ` (${String(item.note).trim().toUpperCase()})` : '');

function setCell(sheet, address, value, bold = false) {
  const cell = Array.isArray(address) ? sheet.getCell(address[0], address[1]) : sheet.getCell(address);
  cell.value = value;
  cell.font = { name: 'Arial', size: 10, bold };
  cell.alignment = { vertical: 'middle' };
  return cell;
}
function sheetBox(sheet, box, startColumn, topRow, rows, mark) {
  if (!box) return;
  const left = sheet.getColumn(startColumn).letter;
  const middle = sheet.getColumn(startColumn + 1).letter;
  sheet.mergeCells(`${left}${topRow}:${middle}${topRow}`);
  sheet.mergeCells(`${left}${topRow + 1}:${middle}${topRow + 1}`);
  for (const [row, label, number, fontSize] of [
    [topRow, 'BOX No.', box.number, 28],
    [topRow + 1, 'PVT. MARK', mark, 14],
  ]) {
    for (const [column, value] of [[startColumn, label], [startColumn + 2, number]]) {
      const cell = sheet.getCell(row, column);
      cell.value = value;
      cell.font = { name: 'Arial', size: fontSize, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = thinBorder();
    }
  }
  for (const [offset, label] of ['S. No.', 'ITEM CODE', 'QTY.'].entries()) {
    const cell = sheet.getCell(topRow + 2, startColumn + offset);
    cell.value = label; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9EDF2' } };
    cell.font = { name: 'Arial', size: 14, bold: true };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = thinBorder();
  }
  for (let i = 0; i < rows; i++) {
    const row = topRow + 3 + i, item = box.items[i];
    sheet.getCell(row, startColumn).value = i + 1;
    if (item) {
      sheet.getCell(row, startColumn + 1).value = `  ${itemLabel(item)}`;
      const qty = sheet.getCell(row, startColumn + 2);
      qty.value = item.quantity; qty.numFmt = '0" PCS"';
    }
    for (let j = 0; j < 3; j++) {
      const cell = sheet.getCell(row, startColumn + j);
      cell.font = { name: 'Arial', size: 12 };
      cell.alignment = { vertical: 'middle', horizontal: j === 1 ? 'left' : 'center' };
      cell.border = thinBorder();
    }
    sheet.getRow(row).height = ROW_PT;
  }
  sheet.getRow(topRow).height = BOX_TITLE_PT;
  sheet.getRow(topRow + 1).height = MARK_PT;
  sheet.getRow(topRow + 2).height = TABLE_HEADER_PT;
}

function thinBorder() {
  const edge = { style: 'thin', color: { argb: 'FF000000' } };
  return { top: edge, bottom: edge, left: edge, right: edge };
}

export async function makeExcel(data) {
  validateList(data);
  const book = new ExcelJS.Workbook();
  for (const [pageIndex, groups] of planPages(data.boxes).entries()) {
    const sheet = book.addWorksheet(`Page ${String(pageIndex + 1).padStart(2, '0')}`);
    sheet.views = [{ showGridLines: false }];
    for (const [col, width] of Object.entries({ A: 9.43, B: 34, C: 11, D: 3, E: 9.71, F: 34, G: 11 })) sheet.getColumn(col).width = width;
    sheet.mergeCells('A1:G1');
    const title = setCell(sheet, 'A1', 'PACKING LIST', true);
    title.font = { name: 'Arial', size: 16, bold: true };
    title.alignment = { horizontal: 'center', vertical: 'middle' };
    sheet.getRow(1).height = 28;
    setCell(sheet, 'A2', 'TO:', true); setCell(sheet, 'B2', data.customer);
    setCell(sheet, 'E2', 'DATE:', true); setCell(sheet, 'F2', data.packing_date);
    setCell(sheet, 'A3', 'TOTAL BOXES:', true); setCell(sheet, 'B3', data.boxes.length);
    setCell(sheet, 'E3', 'TRANSPORT:', true); setCell(sheet, 'F3', data.transport);
    sheet.getCell('A3').alignment = { vertical: 'middle', wrapText: true };
    sheet.getCell('E3').alignment = { vertical: 'middle', wrapText: true };
    sheet.getRow(2).height = sheet.getRow(3).height = 23;
    let nextRow = 5;
    for (const pair of groups) {
      const rows = Math.max(...pair.map(visibleRows));
      sheetBox(sheet, pair[0], 1, nextRow, rows, data.private_mark);
      sheetBox(sheet, pair[1] || null, 5, nextRow, rows, data.private_mark);
      nextRow += 3 + rows + 1;
    }
    const pageHeight = groups.reduce((n, pair) => n + groupHeight(pair), 0);
    const oversized = pageHeight - GROUP_GAP_PT > 790;
    sheet.pageSetup = {
      orientation: 'landscape', paperSize: pageHeight > A4_GROUP_BUDGET_PT ? 8 : 9,
      fitToPage: oversized,
      ...(oversized ? { fitToWidth: 1, fitToHeight: 1 } : { scale: pageHeight <= A4_LARGE_PRINT_PT ? 125 : 100 }),
      margins: { left: .236, right: .236, top: .315, bottom: .315, header: 0, footer: 0 },
      horizontalCentered: true,
    };
    sheet.pageSetup.printArea = `A5:G${nextRow - 2}`;
  }
  return new Blob([await book.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

function fitText(doc, text, maxWidth, start = 9.2) {
  let size = start, output = String(text || '');
  doc.setFontSize(size);
  while (size > 6.5 && doc.getTextWidth(output) > maxWidth) { size -= .3; doc.setFontSize(size); }
  if (doc.getTextWidth(output) > maxWidth) {
    while (output && doc.getTextWidth(`${output}…`) > maxWidth) output = output.slice(0, -1);
    output += '…';
  }
  return output;
}
function pdfBox(doc, box, x, top, width, rows, mark, scale) {
  if (!box) return;
  doc.setDrawColor(0); doc.setLineWidth(.5);
  const numberW = width * 9.43 / 54.43, qtyW = width * 11 / 54.43;
  const qtyLeft = x + width - qtyW;
  const titleBottom = top + BOX_TITLE_PT * scale;
  const markBottom = titleBottom + MARK_PT * scale;
  const headerBottom = markBottom + TABLE_HEADER_PT * scale;
  const tableBottom = headerBottom + ROW_PT * rows * scale;
  doc.setFillColor(233, 237, 242); doc.rect(x, markBottom, width, TABLE_HEADER_PT * scale, 'F');
  for (const y of [top, titleBottom, markBottom, headerBottom, ...Array.from({ length: rows }, (_, i) => headerBottom + (i + 1) * ROW_PT * scale)]) doc.line(x, y, x + width, y);
  for (const xpos of [x, qtyLeft, x + width]) doc.line(xpos, top, xpos, tableBottom);
  doc.line(x + numberW, markBottom, x + numberW, tableBottom);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(28);
  doc.text('BOX No.', x + (width - qtyW) / 2, top + 27 * scale, { align: 'center' });
  doc.setFontSize(28 * scale);
  doc.text(String(box.number), qtyLeft + qtyW / 2, top + 27 * scale, { align: 'center' });
  doc.setFontSize(14 * scale);
  doc.text('PVT. MARK', x + (width - qtyW) / 2, titleBottom + 21 * scale, { align: 'center' });
  doc.text(fitText(doc, mark, qtyW - 8, 14 * scale), qtyLeft + qtyW / 2, titleBottom + 21 * scale, { align: 'center' });
  doc.setFontSize(14 * scale);
  doc.text('S. No.', x + numberW / 2, markBottom + 22 * scale, { align: 'center' });
  doc.text('ITEM CODE', x + numberW + (qtyLeft - x - numberW) / 2, markBottom + 22 * scale, { align: 'center' });
  doc.text('QTY.', qtyLeft + qtyW / 2, markBottom + 22 * scale, { align: 'center' });
  doc.setFont('helvetica', 'normal'); doc.setFontSize(12 * scale);
  for (let i = 0; i < rows; i++) {
    doc.setFontSize(12 * scale);
    const y = headerBottom + ((i + 1) * ROW_PT - 6) * scale;
    doc.text(String(i + 1), x + numberW / 2, y, { align: 'center' });
    const item = box.items[i];
    if (item) {
      doc.text(fitText(doc, itemLabel(item), width - numberW - qtyW - 14, 12 * scale), x + numberW + 7, y);
      doc.setFontSize(12 * scale);
      if (item.quantity !== null) doc.text(`${item.quantity} PCS`, qtyLeft + qtyW / 2, y, { align: 'center' });
    }
  }
}
export function makePdf(data) {
  validateList(data);
  const pages = planPages(data.boxes);
  let doc;
  pages.forEach((groups, pageIndex) => {
    const heightNeeded = groups.reduce((n, pair) => n + groupHeight(pair), 0);
    const format = heightNeeded <= A4_GROUP_BUDGET_PT ? [841.89, 595.28] : [1190.55, Math.max(841.89, heightNeeded + 70)];
    if (!doc) doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format, compress: true });
    else doc.addPage(format, 'landscape');
    const [pageWidth] = format, margin = 25;
    const gap = 15, boxWidth = (pageWidth - 2 * margin - gap) / 2;
    const scale = heightNeeded <= A4_LARGE_PRINT_PT ? 1.25 : 1;
    let top = 25;
    for (const pair of groups) {
      const rows = Math.max(...pair.map(visibleRows));
      pdfBox(doc, pair[0], margin, top, boxWidth, rows, data.private_mark, scale);
      pdfBox(doc, pair[1] || null, margin + boxWidth + gap, top, boxWidth, rows, data.private_mark, scale);
      top += groupHeight(pair) * scale;
    }
  });
  return doc.output('blob');
}

function value(cell) {
  const v = cell?.value;
  if (v && typeof v === 'object') return v.text || v.result || v.richText?.map(x => x.text).join('') || '';
  return v ?? '';
}
function splitLabel(raw) {
  const label = String(raw || '').trim();
  const match = label.match(/^([A-Z]{1,4}\s*[- ]?\s*\d{1,5})\s*(XXL|XL|XS|S|M|L)?\s*(.*)$/i);
  if (!match) return { code: label, size: '', note: '' };
  return { code: match[1].toUpperCase().replace(/^([A-Z]{1,4})\s*[- ]?\s*(\d+)$/, '$1-$2'), size: (match[2] || '').toUpperCase(), note: match[3].trim().replace(/^[()\s]+|[()\s]+$/g, '') };
}
export async function parseFinishedWorkbook(file) {
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(await file.arrayBuffer());
  const sheets = book.worksheets.filter(s => /^PRINT/i.test(s.name.trim()) || /^Page /i.test(s.name));
  if (!sheets.length) throw new Error('No PRINT or Page sheet found');
  const found = new Map(); let privateMark = '';
  for (const sheet of sheets) {
    const headings = [];
    sheet.eachRow((row, rowNo) => row.eachCell((cell, col) => {
      if (cell.isMerged && cell.master !== cell) return;
      if (String(value(cell)).trim().toUpperCase().replace(/\s/g, '').startsWith('BOXNO')) {
        const adjacent = Number(value(sheet.getCell(rowNo, col + 1)));
        const merged = Number(value(sheet.getCell(rowNo, col + 2)));
        const number = Number.isInteger(adjacent) && adjacent > 0 ? adjacent : merged;
        if (Number.isInteger(number) && number > 0) headings.push({ row: rowNo, col, itemCol: number === adjacent ? col : col + 1, number });
      }
    }));
    for (const h of headings) {
      const next = headings.filter(x => x.itemCol === h.itemCol && x.row > h.row).map(x => x.row);
      const end = Math.min(...next, sheet.rowCount + 1);
      let header = null;
      for (let row = h.row + 1; row < Math.min(h.row + 9, end); row++) {
        if (String(value(sheet.getCell(row, h.itemCol))).trim().toUpperCase() === 'ITEM CODE') { header = row; break; }
      }
      if (!header) continue;
      if (!privateMark) privateMark = String(value(sheet.getCell(h.row + 1, h.itemCol + 1))).trim();
      const items = [];
      for (let row = header + 1; row < end; row++) {
        const label = value(sheet.getCell(row, h.itemCol));
        const quantity = String(value(sheet.getCell(row, h.itemCol + 1))).match(/\d+/)?.[0];
        if (!label || !quantity) continue;
        items.push({ ...splitLabel(label), quantity: Number(quantity), needs_review: false, source_text: '' });
      }
      if (items.length) {
        if (found.has(h.number)) throw new Error(`Box ${h.number} occurs twice`);
        found.set(h.number, { number: h.number, items, needs_review: false });
      }
    }
  }
  if (!found.size) throw new Error('No filled boxes found');
  const stem = file.name.replace(/\.xlsx$/i, '');
  const date = stem.match(/\b\d{2}[- ]\d{2}[- ]\d{4}\b/)?.[0]?.replace(/ /g, '-') || '';
  return { customer: stem.split(/\b\d{2}[- ]\d{2}[- ]\d{4}\b/)[0].replace(/[ _-]+$/, ''), packing_date: date, private_mark: privateMark, transport: '', boxes: [...found.values()].sort((a, b) => a.number - b.number), warnings: ['Imported from Excel. Compare with source photos before saving the verified example.'] };
}
