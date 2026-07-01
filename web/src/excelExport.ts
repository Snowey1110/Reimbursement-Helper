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
import { downloadBlob, formatAmount, isKoreaOtherText, itemSearchText, koreaInvoiceKindForItem, reportCategoryForItem, safeNumber, sortReceiptsForReport } from "./utils";

const KOREA_DETAIL_COLUMNS = "ABCDEFGHIJKLMNOPQRS".split("");
const KRW_NUMBER_FORMAT = "₩#,##0";

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

function clearWorksheetPanes(sheet: ExcelJS.Worksheet): void {
  sheet.views = [{ state: "normal", activeCell: "A1", showGridLines: true } as ExcelJS.WorksheetView];
}

function categoryForUsa(item: ReceiptItem): string {
  return reportCategoryForItem(item, "USA");
}

function categoryForKorea(item: ReceiptItem): string {
  return reportCategoryForItem(item, "Korea");
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
      krw = amount * rates.usdToKrw;
      rmb = krw * rates.krwToRmb;
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
    note: amount !== undefined && item.currency === "USD" ? `${formatAmount(amount)} USD` : ""
  };
}

function truncateAmount(value: number | undefined): number | undefined {
  if (value === undefined) return undefined;
  return Math.trunc(value);
}

export interface UsaExpenseRow {
  row: number;
  category: string;
  item: ReceiptItem;
}

export function mapUsaExpenseRows(items: ReceiptItem[]): UsaExpenseRow[] {
  const cursors: Record<string, number> = {};
  return sortReceiptsForReport(items, "USA").map((item) => {
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
  return sortReceiptsForReport(items, "Korea").map((item, index) => {
    const row = index + 3;
    if (row > 32) throw new Error("Korea template supports up to 30 detail rows.");
    const category = categoryForKorea(item);
    const amounts = koreaAmounts(item, rates);
    return {
      row,
      category,
      categoryColumn: KOREA_CATEGORY_COLUMNS[category],
      bucket: koreaBucket(category),
      krw: truncateAmount(amounts.krw),
      rmb: amounts.rmb,
      note: amounts.note,
      item
    };
  });
}

export function koreaDetailContent(item: ReceiptItem): string {
  const text = itemSearchText(item);
  const category = categoryForKorea(item);
  const kind = koreaInvoiceKindForItem(item);
  if (kind === "car_rental") return "\u79df\u8f66\u8d39\u7528";
  if (/\b(fuel|gas|gasoline|petrol|shell|bp|chevron|exxon)\b/.test(text)) return "\u7528\u8f66\u52a0\u6cb9";
  if (/\b(parking|park)\b/.test(text)) return "\u505c\u8f66\u8d39";
  if (/\btoll\b/.test(text)) return "\u8fc7\u8def\u8d39";
  if (isKoreaOtherText(text)) return "ESIM\u6d41\u91cf\u865a\u62df\u5361";
  if (/\busb\b/.test(text)) return "\u8d2d\u4e70USB";
  if (category === "materials") return "\u7269\u6599\u91c7\u8d2d";
  if (category === "physical_exam") return "\u4f53\u68c0\u8d39";
  if (category === "lodging") return "\u4f4f\u5bbf\u8d39";
  if (category === "nucleic_test") return "\u6838\u9178\u68c0\u6d4b\u8d39";
  if (category === "meals") return "\u9910\u8d39";
  if (category === "courier") return "\u5feb\u9012\u8d39";
  if (category === "consumables") return "\u6d88\u8017\u54c1";
  if (category === "welfare") return "\u798f\u5229\u8d39";
  return item.purpose.trim() || item.details.trim() || item.receiptLabel.trim() || item.filename;
}

export function koreaDetailLocation(item: ReceiptItem): string {
  const text = itemSearchText(item);
  if (
    item.currency === "USD" ||
    ["car_rental", "fuel", "parking", "esim", "materials"].includes(koreaInvoiceKindForItem(item)) ||
    /\b(usa|u\.s\.a|united states|ga|mi|il|ca|ny|atlanta|duluth|utica|chicago)\b/.test(text)
  ) {
    return "\u7f8e\u56fd";
  }
  if (/korea|\ud55c\uad6d|\ub300\ud55c\ubbfc\uad6d/.test(text)) return "\u97e9\u56fd";
  if (/china|macao|macau|\u4e2d\u56fd|\u6fb3\u95e8/.test(text)) return "\u4e2d\u56fd";
  return item.place.trim();
}

export function koreaPaymentMethodText(item: ReceiptItem): string {
  const raw = item.paymentMethod.trim();
  const text = raw.toLowerCase();
  if (text.includes("master")) return "MASTERCARD";
  if (text.includes("visa")) return "VISA";
  if (text.includes("wechat") || text.includes("weixin") || text.includes("\u5fae\u4fe1")) return "wechat";
  if (text.includes("credit")) return "Credit Card";
  if (text === "card") return "Credit Card";
  return raw;
}

export function koreaReceiptShouldBeWide(item: ReceiptItem): boolean {
  return koreaInvoiceKindForItem(item) === "car_rental" || itemSearchText(item).includes("national") || item.images.length > 1;
}

async function addImageToCell(
  workbook: ExcelJS.Workbook,
  sheet: ExcelJS.Worksheet,
  attachment: ImageAttachment | ImageAttachment[] | undefined,
  cell: string,
  maxWidth: number,
  maxHeight: number,
  allowUpscale = false,
  stretchToFit = false,
  stretchTiles = false
): Promise<void> {
  if (!attachment) return;
  const prepared = Array.isArray(attachment)
    ? await makeContactSheet(attachment, maxWidth, maxHeight, allowUpscale, stretchTiles)
    : await preparedImageDataUrl(attachment, maxWidth, maxHeight, allowUpscale);
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
    ext: {
      width: stretchToFit ? maxWidth : prepared.displayWidth ?? prepared.width,
      height: stretchToFit ? maxHeight : prepared.displayHeight ?? prepared.height
    }
  });
}

export function koreaReceiptLastRow(itemCount: number): number {
  const pageHeight = 60;
  const blocksPerPage = 4;
  const pageCount = Math.max(1, Math.ceil(itemCount / blocksPerPage));
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
  const content = koreaDetailContent(item);
  const cost = item.amount.trim()
    ? `${item.amount.trim()} ${item.currency}`
    : item.krwAmount.trim()
      ? `${item.krwAmount.trim()} KRW`
      : item.rmbAmount.trim()
        ? `${item.rmbAmount.trim()} RMB`
        : "";
  const details = [item.date.trim(), content, cost].filter(Boolean);
  return details.length ? details.join(" | ") : item.filename || `Payment ${index + 1}`;
}

export function koreaReceiptImageSlots(itemCount: number): KoreaReceiptImageSlot[] {
  return Array.from({ length: itemCount }, (_, index) => koreaReceiptImageSlot(index));
}

export function koreaReceiptImageSlot(index: number): KoreaReceiptImageSlot {
  const blocksPerPage = 4;
  const pageHeight = 60;
  const rowOffsets = [1, 1, 31, 31];
  const colRanges = [
    ["A", "D"],
    ["E", "H"],
    ["A", "D"],
    ["E", "H"]
  ];
  const page = Math.floor(index / blocksPerPage);
  const slot = index % blocksPerPage;
  const labelRow = page * pageHeight + rowOffsets[slot];
  const imageRow = labelRow + 1;
  const [startCol, endCol] = colRanges[slot];
  return {
    labelRange: `${startCol}${labelRow}:${endCol}${labelRow}`,
    labelCell: `${startCol}${labelRow}`,
    imageCell: `${startCol}${imageRow}`,
    maxWidth: 370,
    maxHeight: 520
  };
}

export function koreaReceiptWideSlot(slotIndex: number): KoreaReceiptImageSlot {
  const pageHeight = 60;
  const page = Math.floor(slotIndex / 4);
  const slot = slotIndex % 4;
  const rowOffsets = [1, 1, 31, 31];
  const labelRow = page * pageHeight + rowOffsets[slot];
  return {
    labelRange: `A${labelRow}:H${labelRow}`,
    labelCell: `A${labelRow}`,
    imageCell: `A${labelRow + 1}`,
    maxWidth: 760,
    maxHeight: 520
  };
}

function koreaReceiptSlotCount(blocks: Array<{ wide: boolean }>): number {
  let slotIndex = 0;
  for (const block of blocks) {
    if (block.wide && slotIndex % 2) slotIndex += 1;
    slotIndex += block.wide ? 2 : 1;
  }
  return Math.max(1, slotIndex);
}

function configureKoreaReceiptPage(sheet: ExcelJS.Worksheet, slotCount: number): number {
  const pageHeight = 60;
  const lastRow = koreaReceiptLastRow(slotCount);
  for (const col of "ABCDEFGH".split("")) {
    sheet.getColumn(col).width = 14;
  }
  sheet.pageSetup = {
    ...sheet.pageSetup,
    orientation: "portrait",
    fitToPage: true,
    scale: 78,
    fitToWidth: undefined,
    fitToHeight: 0,
    printArea: `A1:H${lastRow}`,
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
  const sortedItems = sortReceiptsForReport(items, "USA");
  expense.getCell("A3").value = `Date / 填表日期： ${new Date().toLocaleDateString("en-US")}`;
  expense.getCell("A4").value = "Employee: / 申请人：";
  expense.getCell("J1").value = rates.usdToRmb;
  const allRows = Array.from(new Set(Object.values(USA_CATEGORY_ROWS).flat())).sort((a, b) => a - b);
  clearCells(expense, allRows, ["A", "B", "C", "D", "E", "F", "H"]);
  for (const row of allRows) {
    expense.getCell(`G${row}`).value = { formula: `F${row}*$J$1` };
  }
  for (const { item, row } of mapUsaExpenseRows(sortedItems)) {
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
  for (let index = 0; index < sortedItems.length; index += 1) {
    const item = sortedItems[index];
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

export async function exportKoreaWorkbook(items: ReceiptItem[], rates: ExchangeRates, exchangeRateImages: ImageAttachment[] = []): Promise<void> {
  const workbook = await fetchWorkbook(KOREA_TEMPLATE_URL);
  const cover = workbook.worksheets[0];
  const details = workbook.getWorksheet("报销明细") ?? workbook.worksheets[1];
  const receipts = workbook.worksheets[2];
  if (!cover || !details || !receipts) throw new Error("Korea template is missing required sheets.");
  const sortedItems = sortReceiptsForReport(items, "Korea");
  (details as unknown as { orderNo: number }).orderNo = 0;
  (cover as unknown as { orderNo: number }).orderNo = 1;
  (receipts as unknown as { orderNo: number }).orderNo = 2;
  clearWorksheetPanes(details);
  details.pageSetup = {
    ...details.pageSetup,
    orientation: "landscape",
    fitToPage: true,
    scale: 34,
    printArea: "A1:S34",
    fitToWidth: undefined,
    fitToHeight: undefined,
    margins: {
      left: 0.25,
      right: 0.25,
      top: 0.75,
      bottom: 0.75,
      header: 0.3,
      footer: 0.3
    }
  };
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
  for (let row = 3; row <= 35; row += 1) {
    for (const col of KOREA_DETAIL_COLUMNS) {
      details.getCell(`${col}${row}`).value = null;
    }
  }
  try {
    details.unMergeCells("A34:B34");
    details.unMergeCells("A35:B35");
  } catch {
    // Template may already use the target merged ranges.
  }
  try {
    details.mergeCells("A33:B33");
    details.mergeCells("A34:B34");
  } catch {
    // Ignore if already merged.
  }
  details.getCell("A34").value = "合计（外币）\nTotal";
  details.getCell("Q34").value = { formula: "SUM(Q3:Q33)" };
  details.getCell("Q34").numFmt = KRW_NUMBER_FORMAT;
  details.getCell("A35").value = "合计（人民币）\nTotal";
  details.getCell("R35").value = { formula: "SUM(R3:R34)" };
  details.getCell("A35").value = null;
  details.getCell("Q34").value = null;
  details.getCell("R35").value = null;
  details.getCell("A33").value = "合计（外币）\nTotal";
  details.getCell("Q33").value = { formula: "SUM(Q3:Q32)" };
  details.getCell("Q33").numFmt = KRW_NUMBER_FORMAT;
  details.getCell("A34").value = "合计（人民币）\nTotal";
  details.getCell("R34").value = { formula: "SUM(R3:R33)" };
  details.getRow(33).height = 35.4;
  details.getRow(34).height = 50.4;
  details.getRow(35).height = 15;
  const summary: Record<string, { krw: number; rmb: number }> = {};
  for (const key of Object.keys(KOREA_COVER_ROWS)) summary[key] = { krw: 0, rmb: 0 };
  for (const mapped of mapKoreaDetailRows(sortedItems, rates)) {
    const { row, item } = mapped;
    details.getCell(`A${row}`).value = item.date;
    details.getCell(`B${row}`).value = koreaDetailContent(item);
    details.getCell(`C${row}`).value = koreaDetailLocation(item);
    details.getCell(`D${row}`).value = "";
    details.getCell(`E${row}`).value = item.projectNumber;
    details.getCell(`${mapped.categoryColumn}${row}`).value = mapped.krw ?? null;
    details.getCell(`${mapped.categoryColumn}${row}`).numFmt = KRW_NUMBER_FORMAT;
    details.getCell(`P${row}`).value = mapped.note;
    details.getCell(`Q${row}`).value = mapped.krw ?? null;
    details.getCell(`Q${row}`).numFmt = KRW_NUMBER_FORMAT;
    details.getCell(`R${row}`).value = mapped.rmb ?? null;
    details.getCell(`S${row}`).value = koreaPaymentMethodText(item);
    summary[mapped.bucket].krw += mapped.krw ?? 0;
    summary[mapped.bucket].rmb += mapped.rmb ?? 0;
  }
  for (const [bucket, totals] of Object.entries(summary)) {
    const row = KOREA_COVER_ROWS[bucket].row;
    cover.getCell(`C${row}`).value = totals.krw ? Math.trunc(totals.krw) : null;
    cover.getCell(`C${row}`).numFmt = KRW_NUMBER_FORMAT;
    cover.getCell(`D${row}`).value = totals.rmb ? Math.round(totals.rmb * 100) / 100 : null;
  }
  cover.getCell("C9").numFmt = KRW_NUMBER_FORMAT;
  const blocks = [
    ...(exchangeRateImages.length ? [{ label: "汇率 / Exchange Rate", images: exchangeRateImages, wide: true, stretchTiles: false }] : []),
    ...sortedItems.map((item, index) => ({
      label: koreaReceiptPaymentLabel(index, item),
      images: item.images,
      wide: koreaReceiptShouldBeWide(item),
      stretchTiles: true
    }))
  ];
  const slotCount = koreaReceiptSlotCount(blocks);
  const lastReceiptRow = configureKoreaReceiptPage(receipts, slotCount);
  for (let row = 1; row <= Math.max(240, lastReceiptRow); row += 1) {
    for (const col of "ABCDEFGH".split("")) {
      receipts.getCell(`${col}${row}`).value = null;
    }
  }
  let slotIndex = 0;
  for (let index = 0; index < blocks.length; index += 1) {
    if (blocks[index].wide && slotIndex % 2) slotIndex += 1;
    const slot = blocks[index].wide ? koreaReceiptWideSlot(slotIndex) : koreaReceiptImageSlot(slotIndex);
    receipts.mergeCells(slot.labelRange);
    receipts.getCell(slot.labelCell).value = blocks[index].label;
    receipts.getCell(slot.labelCell).font = { bold: true, size: 10, color: { argb: "FF1F2937" } };
    receipts.getCell(slot.labelCell).alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    receipts.getCell(slot.labelCell).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE9EEF3" } };
    receipts.getRow(Number(slot.labelCell.replace(/^[A-Z]+/, ""))).height = 22;
    await addImageToCell(workbook, receipts, blocks[index].images, slot.imageCell, slot.maxWidth, slot.maxHeight, true, true, blocks[index].stretchTiles);
    slotIndex += blocks[index].wide ? 2 : 1;
  }
  const buffer = await workbook.xlsx.writeBuffer();
  downloadBlob(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `reimbursement_korea_${dateStamp()}.xlsx`);
}

function dateStamp(): string {
  return new Date().toISOString().slice(0, 10).replace(/-/g, "");
}
