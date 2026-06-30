import ExcelJS from "exceljs";
import type { ExchangeRates, ImageAttachment, PaymentProof, ReceiptItem } from "./types";
import {
  KOREA_CATEGORY_COLUMNS,
  KOREA_COVER_ROWS,
  KOREA_TEMPLATE_URL,
  USA_CATEGORY_ROWS,
  USA_TEMPLATE_URL
} from "./constants";
import { makeContactSheet, preparedImageDataUrl } from "./imageUtils";
import { downloadBlob, formatAmount, safeNumber } from "./utils";

const KOREA_DETAIL_COLUMNS = "ABCDEFGHIJKLMNOPQRS".split("");

function dataUrlBase64(dataUrl: string): string {
  return dataUrl.split(",", 2)[1] ?? dataUrl;
}

async function fetchWorkbook(url: string): Promise<ExcelJS.Workbook> {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Could not load template: ${url}`);
  const buffer = await response.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook;
}

function clearCells(sheet: ExcelJS.Worksheet, rows: number[], cols: string[]): void {
  for (const row of rows) {
    for (const col of cols) {
      sheet.getCell(`${col}${row}`).value = null;
    }
  }
}

function categoryForUsa(item: ReceiptItem): string {
  if (item.category in USA_CATEGORY_ROWS) return item.category;
  if (item.category === "materials" || item.category === "consumables") return "office";
  return "other";
}

function categoryForKorea(item: ReceiptItem): string {
  if (item.category in KOREA_CATEGORY_COLUMNS) return item.category;
  if (item.category === "advertising" || item.category === "office") return "materials";
  if (item.category === "entertainment") return "meals";
  return "other";
}

function koreaBucket(category: string): keyof typeof KOREA_COVER_ROWS {
  if (category === "transportation") return "transportation";
  if (category === "lodging") return "lodging";
  if (category === "meals" || category === "entertainment") return "meals";
  if (category === "materials" || category === "consumables" || category === "office") return "consumables";
  return "other";
}

function koreaAmounts(item: ReceiptItem, rates: ExchangeRates): { krw?: number; rmb?: number; note: string } {
  const amount = safeNumber(item.amount);
  let krw = safeNumber(item.krwAmount);
  let rmb = safeNumber(item.rmbAmount);
  if (amount !== undefined) {
    if (item.currency === "USD") {
      rmb = amount * rates.usdToRmb;
      krw = rmb / rates.krwToRmb;
    } else if (item.currency === "KRW") {
      krw = amount;
      rmb = amount * rates.krwToRmb;
    } else {
      rmb = amount;
      krw = amount / rates.krwToRmb;
    }
  }
  if (krw === undefined && rmb !== undefined) krw = rmb / rates.krwToRmb;
  if (rmb === undefined && krw !== undefined) rmb = krw * rates.krwToRmb;
  return {
    krw,
    rmb,
    note: amount !== undefined && item.currency !== "KRW" ? `(${formatAmount(amount)} ${item.currency})` : ""
  };
}

export interface UsaExpenseRow {
  row: number;
  category: string;
  item: ReceiptItem;
}

export function mapUsaExpenseRows(items: ReceiptItem[]): UsaExpenseRow[] {
  const cursors: Record<string, number> = {};
  return items.map((item) => {
    const category = categoryForUsa(item);
    const rows = USA_CATEGORY_ROWS[category] ?? USA_CATEGORY_ROWS.other;
    const cursor = cursors[category] ?? 0;
    if (cursor >= rows.length) throw new Error(`Not enough USA rows for ${category}.`);
    cursors[category] = cursor + 1;
    return { row: rows[cursor], category, item };
  });
}

export interface KoreaDetailRow {
  row: number;
  category: string;
  categoryColumn: string;
  bucket: keyof typeof KOREA_COVER_ROWS;
  krw?: number;
  rmb?: number;
  note: string;
  item: ReceiptItem;
}

export function mapKoreaDetailRows(items: ReceiptItem[], rates: ExchangeRates): KoreaDetailRow[] {
  return items.map((item, index) => {
    const row = index + 3;
    if (row > 33) throw new Error("Korea template supports up to 31 detail rows.");
    const category = categoryForKorea(item);
    const amounts = koreaAmounts(item, rates);
    return {
      row,
      category,
      categoryColumn: KOREA_CATEGORY_COLUMNS[category],
      bucket: koreaBucket(category),
      krw: amounts.krw,
      rmb: amounts.rmb,
      note: amounts.note,
      item
    };
  });
}

async function addImageToCell(
  workbook: ExcelJS.Workbook,
  sheet: ExcelJS.Worksheet,
  attachment: ImageAttachment | ImageAttachment[] | undefined,
  cell: string,
  maxWidth: number,
  maxHeight: number
): Promise<void> {
  if (!attachment) return;
  const prepared = Array.isArray(attachment)
    ? await makeContactSheet(attachment, maxWidth, maxHeight)
    : await preparedImageDataUrl(attachment, maxWidth, maxHeight);
  if (!prepared) return;
  const imageId = workbook.addImage({
    base64: dataUrlBase64(prepared.dataUrl),
    extension: "png"
  });
  const match = /^([A-Z]+)(\d+)$/.exec(cell);
  if (!match) return;
  const colLetters = match[1];
  const row = Number(match[2]) - 1;
  let col = 0;
  for (const char of colLetters) {
    col = col * 26 + (char.charCodeAt(0) - 64);
  }
  sheet.addImage(imageId, {
    tl: { col: col - 1, row },
    ext: { width: prepared.width, height: prepared.height }
  });
}

export function koreaReceiptLastRow(itemCount: number): number {
  const pageHeight = 50;
  const receiptsPerPage = 4;
  const pageCount = Math.max(1, Math.ceil(itemCount / receiptsPerPage));
  return pageCount * pageHeight;
}

export interface KoreaReceiptImageSlot {
  labelRange: string;
  labelCell: string;
  imageCell: string;
  maxWidth: number;
  maxHeight: number;
}

export function koreaReceiptPaymentLabel(index: number, item: ReceiptItem): string {
  const details = [item.paymentMethod.trim(), item.date.trim(), item.amount.trim() ? `${item.amount.trim()} ${item.currency}` : ""].filter(Boolean);
  return details.length ? `Payment ${index + 1}: ${details.join(" | ")}` : `Payment ${index + 1}`;
}

export function koreaReceiptImageSlots(itemCount: number): KoreaReceiptImageSlot[] {
  const receiptsPerPage = 4;
  const pageHeight = 50;
  const anchors = ["A", "D"];
  const rowOffsets = [1, 25];
  return Array.from({ length: itemCount }, (_, index) => {
    const page = Math.floor(index / receiptsPerPage);
    const slot = index % receiptsPerPage;
    const rowGroup = Math.floor(slot / anchors.length);
    const col = anchors[slot % anchors.length];
    const labelRow = page * pageHeight + rowOffsets[rowGroup];
    const imageRow = labelRow + 1;
    return {
      labelRange: col === "A" ? `A${labelRow}:C${labelRow}` : `D${labelRow}:E${labelRow}`,
      labelCell: `${col}${labelRow}`,
      imageCell: `${col}${imageRow}`,
      maxWidth: 215,
      maxHeight: 280
    };
  });
}

function configureKoreaReceiptPage(sheet: ExcelJS.Worksheet, itemCount: number): number {
  const pageHeight = 50;
  const lastRow = koreaReceiptLastRow(itemCount);
  for (const col of "ABCDEFGH".split("")) {
    sheet.getColumn(col).width = 16;
  }
  sheet.pageSetup = {
    ...sheet.pageSetup,
    orientation: "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    printArea: `A1:E${lastRow}`,
    margins: {
      left: 0.45,
      right: 0.45,
      top: 0.45,
      bottom: 0.45,
      header: 0.3,
      footer: 0.3
    }
  };
  (sheet as unknown as { rowBreaks?: unknown[] }).rowBreaks = [];
  for (let row = pageHeight; row < lastRow; row += pageHeight) {
    sheet.getRow(row).addPageBreak();
  }
  return lastRow;
}

export async function exportUsaWorkbook(items: ReceiptItem[], proofs: PaymentProof[], rates: ExchangeRates): Promise<void> {
  const workbook = await fetchWorkbook(USA_TEMPLATE_URL);
  const expense = workbook.getWorksheet("Expense report");
  const receipts = workbook.getWorksheet("Receipt and Payment of expenses");
  if (!expense || !receipts) throw new Error("USA template is missing required sheets.");
  expense.getCell("A3").value = `Date / 填表日期： ${new Date().toLocaleDateString("en-US")}`;
  expense.getCell("A4").value = "Employee: / 申请人：";
  expense.getCell("J1").value = rates.usdToRmb;
  const allRows = Array.from(new Set(Object.values(USA_CATEGORY_ROWS).flat())).sort((a, b) => a - b);
  clearCells(expense, allRows, ["A", "B", "C", "D", "E", "F", "H"]);
  for (const row of allRows) {
    expense.getCell(`G${row}`).value = { formula: `F${row}*$J$1` };
  }
  for (const { item, row } of mapUsaExpenseRows(items)) {
    expense.getCell(`A${row}`).value = item.place;
    expense.getCell(`B${row}`).value = item.date;
    expense.getCell(`C${row}`).value = item.details || item.receiptLabel || item.filename;
    expense.getCell(`D${row}`).value = item.purpose;
    expense.getCell(`E${row}`).value = item.projectNumber;
    expense.getCell(`F${row}`).value = safeNumber(item.amount) ?? null;
    expense.getCell(`G${row}`).value = { formula: `F${row}*$J$1` };
  }
  for (let row = 2; row < Math.max(56, items.length + 3); row += 1) {
    for (const col of ["A", "B", "C", "D", "E"]) {
      receipts.getCell(`${col}${row}`).value = null;
    }
    receipts.getRow(row).height = 126;
  }
  receipts.getColumn("D").width = 38;
  receipts.getColumn("E").width = 26;
  for (let index = 0; index < items.length; index += 1) {
    const item = items[index];
    const row = index + 2;
    receipts.getCell(`A${row}`).value = index + 1;
    receipts.getCell(`B${row}`).value = item.date;
    receipts.getCell(`C${row}`).value = safeNumber(item.amount) ?? null;
    await addImageToCell(workbook, receipts, item.images, `D${row}`, 260, 150);
    const proofImages = proofs.filter((proof) => proof.matchedReceiptId === item.id).map((proof) => proof.image);
    await addImageToCell(workbook, receipts, proofImages, `E${row}`, 180, 150);
  }
  const buffer = await workbook.xlsx.writeBuffer();
  downloadBlob(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `reimbursement_usa_${dateStamp()}.xlsx`);
}

export async function exportKoreaWorkbook(items: ReceiptItem[], rates: ExchangeRates): Promise<void> {
  const workbook = await fetchWorkbook(KOREA_TEMPLATE_URL);
  const cover = workbook.worksheets[0];
  const details = workbook.getWorksheet("报销明细") ?? workbook.worksheets[1];
  const receipts = workbook.worksheets[2];
  if (!cover || !details || !receipts) throw new Error("Korea template is missing required sheets.");
  const now = new Date();
  cover.getCell("A2").value = `报销部门：  ${now.getFullYear()}年 ${now.getMonth() + 1}月 ${now.getDate()}日 填 单据及附件共  页`;
  cover.getCell("A11").value = "领导审批           会计主管              会计                  出纳                 报销人                   领款人 ";
  for (const { row, label } of Object.values(KOREA_COVER_ROWS)) {
    cover.getCell(`A${row}`).value = label;
    cover.getCell(`C${row}`).value = null;
    cover.getCell(`D${row}`).value = null;
  }
  cover.getCell("C9").value = { formula: "SUM(C4:C8)" };
  cover.getCell("D9").value = { formula: "SUM(D4:D8)" };
  cover.getCell("B10").value = { formula: "D9" };
  for (let row = 3; row <= 34; row += 1) {
    for (const col of KOREA_DETAIL_COLUMNS) {
      details.getCell(`${col}${row}`).value = null;
    }
  }
  details.getCell("A34").value = "合计（外币）\nTotal";
  details.getCell("Q34").value = { formula: "SUM(Q3:Q33)" };
  details.getCell("A35").value = "合计（人民币）\nTotal";
  details.getCell("R35").value = { formula: "SUM(R3:R34)" };
  const summary: Record<string, { krw: number; rmb: number }> = {};
  for (const key of Object.keys(KOREA_COVER_ROWS)) summary[key] = { krw: 0, rmb: 0 };
  for (const mapped of mapKoreaDetailRows(items, rates)) {
    const { row, item } = mapped;
    details.getCell(`A${row}`).value = item.date;
    details.getCell(`B${row}`).value = item.purpose || item.details;
    details.getCell(`C${row}`).value = item.place;
    details.getCell(`D${row}`).value = "";
    details.getCell(`E${row}`).value = item.projectNumber;
    details.getCell(`${mapped.categoryColumn}${row}`).value = mapped.krw ?? null;
    details.getCell(`P${row}`).value = mapped.note;
    details.getCell(`Q${row}`).value = mapped.krw ?? null;
    details.getCell(`R${row}`).value = mapped.rmb ?? null;
    details.getCell(`S${row}`).value = item.paymentMethod;
    summary[mapped.bucket].krw += mapped.krw ?? 0;
    summary[mapped.bucket].rmb += mapped.rmb ?? 0;
  }
  for (const [bucket, totals] of Object.entries(summary)) {
    const row = KOREA_COVER_ROWS[bucket].row;
    cover.getCell(`C${row}`).value = totals.krw ? Math.round(totals.krw * 100) / 100 : null;
    cover.getCell(`D${row}`).value = totals.rmb ? Math.round(totals.rmb * 100) / 100 : null;
  }
  const lastReceiptRow = configureKoreaReceiptPage(receipts, items.length);
  for (let row = 1; row <= Math.max(240, lastReceiptRow); row += 1) {
    for (const col of "ABCDEFGH".split("")) {
      receipts.getCell(`${col}${row}`).value = null;
    }
  }
  const slots = koreaReceiptImageSlots(items.length);
  for (let index = 0; index < items.length; index += 1) {
    const slot = slots[index];
    receipts.mergeCells(slot.labelRange);
    receipts.getCell(slot.labelCell).value = koreaReceiptPaymentLabel(index, items[index]);
    receipts.getCell(slot.labelCell).font = { bold: true, size: 10, color: { argb: "FF1F2937" } };
    receipts.getCell(slot.labelCell).alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    receipts.getCell(slot.labelCell).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE9EEF3" } };
    receipts.getRow(Number(slot.labelCell.replace(/^[A-Z]+/, ""))).height = 20;
    await addImageToCell(workbook, receipts, items[index].images, slot.imageCell, slot.maxWidth, slot.maxHeight);
  }
  const buffer = await workbook.xlsx.writeBuffer();
  downloadBlob(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `reimbursement_korea_${dateStamp()}.xlsx`);
}

function dateStamp(): string {
  return new Date().toISOString().slice(0, 10).replace(/-/g, "");
}
