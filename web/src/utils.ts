import {
  DEFAULT_KRW_TO_RMB,
  DEFAULT_USD_TO_KRW,
  DEFAULT_USD_TO_RMB,
  KOREA_CATEGORY_COLUMNS,
  KOREA_INVOICE_KIND_ORDER,
  USA_CATEGORY_ROWS,
  USA_REPORT_CATEGORY_ORDER
} from "./constants";
import type { Category, Currency, DocumentType, ExchangeRates, ImageAttachment, PaymentProof, ReceiptItem } from "./types";

export function uid(prefix = "id"): string {
  return `${prefix}_${crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2)}`;
}

export function safeNumber(value: unknown): number | undefined {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value !== "string") {
    return undefined;
  }
  const cleaned = value.replace(/[^0-9.-]/g, "");
  if (!cleaned) {
    return undefined;
  }
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : undefined;
}

export function formatAmount(value: number | undefined): string {
  if (value === undefined || !Number.isFinite(value)) {
    return "";
  }
  return value.toFixed(2).replace(/\.00$/, "");
}

export function normalizeCurrency(value: unknown, fallback: Currency): Currency {
  const raw = String(value ?? "").trim().toUpperCase();
  if (raw.includes("USD") || raw.includes("$")) return "USD";
  if (raw.includes("KRW") || raw.includes("WON") || raw.includes("\u20a9")) return "KRW";
  if (raw.includes("RMB") || raw.includes("CNY") || raw.includes("YUAN") || raw.includes("\u00a5")) return "RMB";
  return fallback;
}

export function normalizeDocumentType(value: unknown, fallback: DocumentType = "unknown"): DocumentType {
  const raw = String(value ?? "").trim().toLowerCase().replace(/[\s-]+/g, "_");
  if (!raw) return fallback;
  if (["receipt", "invoice", "sales_receipt", "ticket_receipt"].includes(raw)) return "receipt";
  if (
    [
      "payment_proof",
      "proof",
      "payment",
      "payment_detail",
      "payment_details",
      "bank",
      "bank_statement",
      "card_statement",
      "credit_card",
      "transaction",
      "chase"
    ].includes(raw)
  ) {
    return "payment_proof";
  }
  if (["unknown", "unclear", "manual_review", "review"].includes(raw)) return "unknown";
  return fallback;
}

export function normalizeCategory(value: unknown): Category {
  const raw = String(value ?? "").trim().toLowerCase();
  if (!raw) return "other";
  if (isKoreaOtherText(raw)) return "other";
  if (raw.includes("parking") || raw.includes("transport") || raw.includes("taxi") || raw.includes("toll") || raw.includes("fuel")) return "transportation";
  if (raw.includes("hotel") || raw.includes("lodging") || raw.includes("accommodation")) return "lodging";
  if (raw.includes("meal") || raw.includes("food") || raw.includes("restaurant")) return "meals";
  if (raw.includes("advert")) return "advertising";
  if (raw.includes("office")) return "office";
  if (raw.includes("material")) return "materials";
  if (raw.includes("consumable")) return "consumables";
  if (raw.includes("courier") || raw.includes("shipping")) return "courier";
  if (raw.includes("welfare")) return "welfare";
  if (raw.includes("exam")) return "physical_exam";
  if (raw.includes("nucleic")) return "nucleic_test";
  if (raw.includes("entertain")) return "entertainment";
  if (
    [
      "transportation",
      "lodging",
      "meals",
      "advertising",
      "office",
      "entertainment",
      "materials",
      "consumables",
      "physical_exam",
      "nucleic_test",
      "courier",
      "welfare",
      "other"
    ].includes(raw)
  ) {
    return raw as Category;
  }
  return "other";
}

export function isKoreaOtherText(value: string): boolean {
  const raw = value.toLowerCase();
  return /\b(e-?sim|sim card|data plan|internet access)\b/.test(raw) || /\u6d41\u91cf|\u865a\u62df\u5361|\uc720\uc2ec|\ub370\uc774\ud130/.test(raw);
}

export function itemSearchText(item: ReceiptItem): string {
  return [
    item.category,
    item.purpose,
    item.details,
    item.receiptLabel,
    item.place,
    item.filename,
    item.paymentMethod,
    item.bankName,
    item.cardLast4,
    item.transactionReference
  ].join(" ").toLowerCase();
}

export function looksLikeUsaParkingReceipt(item: ReceiptItem): boolean {
  const text = itemSearchText(item);
  return /\b(parking|garage|park(?:ing)? lots?|self park|valet|meter|mccormick place|metropolis|iparkit|parkwhiz|spothero|wabash randolph|millennium park|lakeside center)\b/.test(
    text
  );
}

export function looksLikeUsaTollReceipt(item: ReceiptItem): boolean {
  const text = itemSearchText(item);
  return /\b(toll|turnpike|i-pass|e-zpass|ez-pass|indiana toll road|itr concession|bridge fee)\b/.test(text);
}

export function looksLikeUsaFuelReceipt(item: ReceiptItem): boolean {
  const text = itemSearchText(item);
  return /\b(fuel|gasoline|gas station|shell|bp|chevron|exxon|mobil|speedway|marathon)\b/.test(text);
}

export function normalizeUsaTransportationFields(item: ReceiptItem): ReceiptItem {
  const parking = looksLikeUsaParkingReceipt(item);
  const toll = looksLikeUsaTollReceipt(item);
  const fuel = looksLikeUsaFuelReceipt(item);
  if (!parking && !toll && !fuel) return item;

  const mealLike = /\b(breakfast|lunch|dinner|meal|food|snack|restaurant)\b/i;
  const purpose = toll ? "Toll" : fuel ? "Gas" : "Parking";
  const place = item.place.trim() || item.receiptLabel.trim() || item.filename;
  const details =
    !item.details.trim() || mealLike.test(item.details)
      ? `${purpose} at ${place}.`
      : item.details;
  return {
    ...item,
    category: "transportation",
    purpose: !item.purpose.trim() || mealLike.test(item.purpose) ? purpose : item.purpose,
    details
  };
}

export function reportCategoryForItem(item: ReceiptItem, formVersion: "USA" | "Korea"): Category {
  if (formVersion === "Korea") {
    const text = itemSearchText(item);
    if (isKoreaOtherText(text)) return "other";
    if (item.category in KOREA_CATEGORY_COLUMNS) return item.category;
    if (item.category === "advertising" || item.category === "office") return "materials";
    if (item.category === "entertainment") return "meals";
    return "other";
  }
  if (looksLikeUsaParkingReceipt(item) || looksLikeUsaTollReceipt(item) || looksLikeUsaFuelReceipt(item)) return "transportation";
  if (item.category in USA_CATEGORY_ROWS) return item.category;
  if (item.category === "materials" || item.category === "consumables") return "office";
  return "other";
}

export function koreaInvoiceKindForItem(item: ReceiptItem): string {
  const text = itemSearchText(item);
  const category = reportCategoryForItem(item, "Korea");
  if (/\b(national|car rental|rental car|rent[- ]?a[- ]?car|rental|avis|hertz|enterprise|budget|alamo)\b/.test(text)) return "car_rental";
  if (/\b(fuel|gas|gasoline|petrol|shell|bp|chevron|exxon)\b/.test(text)) return "fuel";
  if (/\b(parking|park|toll)\b/.test(text)) return "parking";
  if (isKoreaOtherText(text)) return "esim";
  if (/\b(usb|office|supplies|supply|material|materials)\b/.test(text)) return "materials";
  if (KOREA_INVOICE_KIND_ORDER.includes(category as (typeof KOREA_INVOICE_KIND_ORDER)[number])) return category;
  return "other";
}

function dateSortValue(value: string): number {
  const key = parseDateKey(value);
  if (!key) return Number.POSITIVE_INFINITY;
  const iso = /^\d{4}-\d{2}-\d{2}$/.test(key) ? `${key}T00:00:00` : key;
  const parsed = new Date(iso);
  if (!Number.isNaN(parsed.getTime())) {
    parsed.setFullYear(new Date().getFullYear());
    return parsed.getTime();
  }
  const time = Date.parse(iso);
  return Number.isNaN(time) ? Number.POSITIVE_INFINITY : time;
}

export function dateYearNeedsReview(value: string): boolean {
  const key = parseDateKey(value);
  if (!key) return false;
  const iso = /^\d{4}-\d{2}-\d{2}$/.test(key) ? `${key}T00:00:00` : key;
  const parsed = new Date(iso);
  return !Number.isNaN(parsed.getTime()) && parsed.getFullYear() !== new Date().getFullYear();
}

function sameMonthDay(left: string, right: string): boolean {
  const leftKey = parseDateKey(left);
  const rightKey = parseDateKey(right);
  if (!leftKey || !rightKey) return false;
  const leftDate = new Date(/^\d{4}-\d{2}-\d{2}$/.test(leftKey) ? `${leftKey}T00:00:00` : leftKey);
  const rightDate = new Date(/^\d{4}-\d{2}-\d{2}$/.test(rightKey) ? `${rightKey}T00:00:00` : rightKey);
  if (!Number.isNaN(leftDate.getTime()) && !Number.isNaN(rightDate.getTime())) {
    return leftDate.getMonth() === rightDate.getMonth() && leftDate.getDate() === rightDate.getDate();
  }
  return leftKey === rightKey;
}

export function sortReceiptsForReview(items: ReceiptItem[]): ReceiptItem[] {
  return items
    .map((item, index) => ({ item, index, review: dateYearNeedsReview(item.date) ? 0 : 1 }))
    .sort((left, right) => left.review - right.review || left.index - right.index)
    .map(({ item }) => item);
}

export function sortReceiptsForReport(items: ReceiptItem[], formVersion: "USA" | "Korea"): ReceiptItem[] {
  const order = formVersion === "Korea" ? KOREA_INVOICE_KIND_ORDER : USA_REPORT_CATEGORY_ORDER;
  const orderIndex = new Map<string, number>(order.map((category, index) => [category, index]));
  return items
    .map((item, index) => ({
      item,
      index,
      category: formVersion === "Korea" ? koreaInvoiceKindForItem(item) : reportCategoryForItem(item, formVersion),
      date: dateSortValue(item.date)
    }))
    .sort((left, right) => {
      const categoryDiff = (orderIndex.get(left.category) ?? order.length) - (orderIndex.get(right.category) ?? order.length);
      if (categoryDiff) return categoryDiff;
      if (left.date !== right.date) return left.date - right.date;
      return left.index - right.index;
    })
    .map(({ item }) => item);
}

export function blankReceipt(formVersion: "USA" | "Korea", image: ImageAttachment): ReceiptItem {
  return {
    id: uid("receipt"),
    filename: image.filename,
    status: "Empty",
    date: "",
    place: "",
    amount: "",
    currency: formVersion === "USA" ? "USD" : "KRW",
    krwAmount: "",
    rmbAmount: "",
    purpose: "",
    details: "",
    projectNumber: "",
    category: "transportation",
    paymentMethod: "",
    receiptLabel: "",
    documentType: formVersion === "USA" ? "unknown" : "receipt",
    matchedReceiptId: "",
    images: [image]
  };
}

export function updateAmounts(item: ReceiptItem, rates: ExchangeRates, source: "amount" | "krw" | "rmb" | "currency" = "amount"): ReceiptItem {
  const usdRmbRate = rates.usdToRmb || DEFAULT_USD_TO_RMB;
  const usdKrwRate = rates.usdToKrw || DEFAULT_USD_TO_KRW;
  const krwRate = rates.krwToRmb || DEFAULT_KRW_TO_RMB;
  const next = { ...item };
  const amount = safeNumber(next.amount);
  let krw = safeNumber(next.krwAmount);
  let rmb = safeNumber(next.rmbAmount);
  if (next.currency === "USD" && amount !== undefined && source !== "krw" && source !== "rmb") {
    krw = amount * usdKrwRate;
    rmb = krw * krwRate;
  } else if (next.currency === "KRW" && amount !== undefined && source !== "rmb") {
    krw = amount;
    rmb = amount * krwRate;
  } else if ((next.currency === "RMB" || next.currency === "CNY") && amount !== undefined && source !== "krw") {
    rmb = amount;
    krw = amount / krwRate;
  } else if (source === "krw" && krw !== undefined) {
    rmb = krw * krwRate;
    if (next.currency === "KRW") next.amount = formatAmount(krw);
    if (next.currency === "USD" && !next.amount.trim()) next.amount = formatAmount(krw / usdKrwRate);
  } else if (source === "rmb" && rmb !== undefined) {
    krw = rmb / krwRate;
    if (next.currency === "RMB" || next.currency === "CNY") next.amount = formatAmount(rmb);
    if (next.currency === "USD" && !next.amount.trim()) next.amount = formatAmount(krw / usdKrwRate || rmb / usdRmbRate);
  }
  next.krwAmount = formatAmount(krw);
  next.rmbAmount = formatAmount(rmb);
  return next;
}

export function parseDateKey(value: string): string {
  const date = new Date(value);
  if (!Number.isNaN(date.getTime())) {
    return date.toISOString().slice(0, 10);
  }
  return value.trim();
}

function assumedYearDateKey(value: string): string {
  const key = parseDateKey(value);
  if (!key) return "";
  const parsed = new Date(/^\d{4}-\d{2}-\d{2}$/.test(key) ? `${key}T00:00:00` : key);
  if (Number.isNaN(parsed.getTime())) return key;
  parsed.setFullYear(new Date().getFullYear());
  return parsed.toISOString().slice(0, 10);
}

export function receiptMergeKey(item: ReceiptItem): string | undefined {
  const amount = safeNumber(item.amount);
  const date = assumedYearDateKey(item.date);
  if (amount === undefined || !date) return undefined;
  return `${date}|${amount.toFixed(2)}`;
}

export function isUsaReceiptItem(item: ReceiptItem): boolean {
  return item.documentType === "receipt";
}

export function isUsaPaymentProofItem(item: ReceiptItem): boolean {
  return item.documentType === "payment_proof";
}

function receiptSideImages(item: ReceiptItem): ImageAttachment[] {
  if (item.documentType === "payment_proof") return [];
  const receiptImages = item.images.filter((image) => image.supportRole !== "payment_proof" && image.supportRole !== "ignore");
  return sortSourcePages(receiptImages.length ? receiptImages : [...item.images]);
}

function proofSideImages(item: ReceiptItem): ImageAttachment[] {
  if (item.documentType === "payment_proof") {
    return sortSourcePages(item.images.filter((image) => image.supportRole !== "ignore"));
  }
  return sortSourcePages(item.images.filter((image) => image.supportRole === "payment_proof"));
}

function sortSourcePages(images: ImageAttachment[]): ImageAttachment[] {
  const sourceOrder = new Map<string, number>();
  for (const image of images) {
    const source = image.sourceName.trim().toLowerCase();
    if (!sourceOrder.has(source)) sourceOrder.set(source, sourceOrder.size);
  }
  return images
    .map((image, index) => ({ image, index }))
    .sort((left, right) => {
      const sourceDiff =
        (sourceOrder.get(left.image.sourceName.trim().toLowerCase()) ?? Number.MAX_SAFE_INTEGER) -
        (sourceOrder.get(right.image.sourceName.trim().toLowerCase()) ?? Number.MAX_SAFE_INTEGER);
      if (sourceDiff) return sourceDiff;
      const pageDiff = (left.image.sourcePage ?? Number.MAX_SAFE_INTEGER) - (right.image.sourcePage ?? Number.MAX_SAFE_INTEGER);
      return pageDiff || left.index - right.index;
    })
    .map(({ image }) => image);
}

export function receiptAttentionReason(item: ReceiptItem, formVersion: "USA" | "Korea"): string | undefined {
  if (item.status === "Failed") return "AI failed";
  if (item.status === "Multiple possible matches") return "Multiple possible matches";
  if (formVersion === "USA" && item.documentType === "unknown") return "Unknown document type";
  if (!item.date.trim()) return "Missing date";
  if (dateYearNeedsReview(item.date)) return "Wrong year";
  if (!item.amount.trim() && !item.krwAmount.trim() && !item.rmbAmount.trim()) return "Missing amount";
  return undefined;
}

export function receiptNeedsAi(item: ReceiptItem, formVersion: "USA" | "Korea"): boolean {
  if (item.status === "Remembered" || item.status === "Manual remembered") return false;
  if (item.status === "Empty" || item.status === "Failed") return true;
  if (!item.date.trim() || !item.amount.trim()) return true;
  if (formVersion === "USA") {
    if (item.documentType === "unknown") return true;
    if (item.documentType === "receipt" && !item.purpose.trim() && !item.details.trim() && !item.receiptLabel.trim()) return true;
  }
  return false;
}

function normalizedTextTokens(value: string): Set<string> {
  const stopWords = new Set([
    "receipt", "invoice", "ticket", "payment", "proof", "transaction", "details",
    "total", "card", "credit", "debit", "visa", "mastercard", "page", "jpg", "jpeg", "png", "pdf"
  ]);
  return new Set(
    (value.toLowerCase().match(/[a-z0-9]+/g) ?? [])
      .filter((token) => token.length >= 3 && !stopWords.has(token))
  );
}

function tokenOverlap(left: ReceiptItem, right: ReceiptItem): number {
  const leftTokens = normalizedTextTokens(itemSearchText(left));
  const rightTokens = normalizedTextTokens(itemSearchText(right));
  return [...leftTokens].filter((token) => rightTokens.has(token)).length;
}

function sameSourceFile(left: ReceiptItem, right: ReceiptItem): boolean {
  const leftSources = new Set(left.images.map((image) => image.sourceName.trim().toLowerCase()).filter(Boolean));
  return right.images.some((image) => leftSources.has(image.sourceName.trim().toLowerCase()));
}

function itemHasSupportRole(item: ReceiptItem, role: ImageAttachment["supportRole"]): boolean {
  return item.images.some((image) => image.supportRole === role);
}

function sameReceiptTransaction(left: ReceiptItem, right: ReceiptItem): boolean {
  const leftAmount = safeNumber(left.amount);
  const rightAmount = safeNumber(right.amount);
  if (leftAmount === undefined || rightAmount === undefined || Math.abs(leftAmount - rightAmount) >= 0.01) return false;
  if (!sameMonthDay(left.transactionDate || left.date, right.transactionDate || right.date)) return false;
  if (left.transactionReference && right.transactionReference) {
    return left.transactionReference.trim().toLowerCase() === right.transactionReference.trim().toLowerCase();
  }
  if (sameSourceFile(left, right)) return true;
  const overlap = tokenOverlap(left, right);
  const companion = itemHasSupportRole(left, "companion_ticket") || itemHasSupportRole(right, "companion_ticket");
  if (companion) return overlap >= 1;
  const leftVendor = left.place.trim().toLowerCase().replace(/[^a-z0-9]+/g, " ");
  const rightVendor = right.place.trim().toLowerCase().replace(/[^a-z0-9]+/g, " ");
  return Boolean(leftVendor && leftVendor === rightVendor && overlap >= 2);
}

export function mergeSameUsaReceipts(items: ReceiptItem[]): ReceiptItem[] {
  const merged: ReceiptItem[] = [];
  for (const item of items) {
    const normalizedItem = normalizeUsaTransportationFields(item);
    if (!isUsaReceiptItem(normalizedItem)) {
      merged.push({ ...normalizedItem, images: [...normalizedItem.images] });
      continue;
    }
    const target = merged.find((candidate) => isUsaReceiptItem(candidate) && sameReceiptTransaction(candidate, normalizedItem));
    if (!target) {
      const copy = { ...normalizedItem, images: sortSourcePages([...normalizedItem.images]) };
      merged.push(copy);
      continue;
    }
    const existing = new Set(target.images.map((image) => `${image.sourceName}|${image.filename}|${image.dataUrl.slice(0, 64)}`));
    for (const image of normalizedItem.images) {
      const imageKey = `${image.sourceName}|${image.filename}|${image.dataUrl.slice(0, 64)}`;
      if (!existing.has(imageKey)) {
        target.images.push(image);
        existing.add(imageKey);
      }
    }
    target.images = sortSourcePages(target.images);
    for (const keyName of ["place", "purpose", "details", "projectNumber", "paymentMethod", "receiptLabel"] as const) {
      if (!target[keyName] && normalizedItem[keyName]) {
        target[keyName] = normalizedItem[keyName];
      }
    }
    target.status = "Merged";
  }
  return merged;
}

export function matchUsaPaymentProofItems(items: ReceiptItem[]): ReceiptItem[] {
  const receipts = items.filter(isUsaReceiptItem);
  const seenProofKeys = new Map<string, string>();
  const matched = items.map((item) => {
    if (!isUsaPaymentProofItem(item)) return item;
    const proofAmount = safeNumber(item.amount);
    const proofDate = item.transactionDate || item.date;
    if (proofAmount === undefined || !proofDate) {
      return { ...item, matchedReceiptId: "", status: "Missing date/amount" };
    }
    const proofKey = paymentProofTransactionKey(item);
    if (proofKey && seenProofKeys.has(proofKey)) {
      return { ...item, matchedReceiptId: "", status: "Duplicate proof" };
    }
    if (proofKey) seenProofKeys.set(proofKey, item.id);
    const candidates = receipts.map((receipt) => {
      const itemAmount = safeNumber(receipt.amount);
      if (itemAmount === undefined || Math.abs(itemAmount - proofAmount) >= 0.01) return undefined;
      const days = dateDistanceDays(proofDate, receipt.transactionDate || receipt.date);
      if (days === undefined || days > 1) return undefined;
      let score = days === 0 ? 4 : 2;
      if (
        item.transactionReference &&
        receipt.transactionReference &&
        item.transactionReference.trim().toLowerCase() === receipt.transactionReference.trim().toLowerCase()
      ) score += 8;
      if (item.cardLast4 && receipt.cardLast4 && item.cardLast4 === receipt.cardLast4) score += 2;
      score += Math.min(4, tokenOverlap(item, receipt));
      return { receipt, score };
    }).filter((candidate): candidate is { receipt: ReceiptItem; score: number } => Boolean(candidate));
    candidates.sort((left, right) => right.score - left.score);
    if (candidates.length === 1 || (candidates.length > 1 && candidates[0].score > candidates[1].score)) {
      return { ...item, matchedReceiptId: candidates[0].receipt.id, status: "Matched" };
    }
    return { ...item, matchedReceiptId: "", status: candidates.length > 1 ? "Multiple possible matches" : "Payment proof" };
  });
  const assignmentCounts = new Map<string, number>();
  for (const item of matched) {
    if (isUsaPaymentProofItem(item) && item.status === "Matched" && item.matchedReceiptId) {
      assignmentCounts.set(item.matchedReceiptId, (assignmentCounts.get(item.matchedReceiptId) ?? 0) + 1);
    }
  }
  return matched.map((item) =>
    isUsaPaymentProofItem(item) &&
    item.status === "Matched" &&
    item.matchedReceiptId &&
    (assignmentCounts.get(item.matchedReceiptId) ?? 0) > 1
      ? { ...item, matchedReceiptId: "", status: "Multiple possible matches" }
      : item
  );
}

function dateDistanceDays(left: string, right: string): number | undefined {
  const leftKey = parseDateKey(left);
  const rightKey = parseDateKey(right);
  if (!leftKey || !rightKey) return undefined;
  const leftDate = new Date(/^\d{4}-\d{2}-\d{2}$/.test(leftKey) ? `${leftKey}T00:00:00` : leftKey);
  const rightDate = new Date(/^\d{4}-\d{2}-\d{2}$/.test(rightKey) ? `${rightKey}T00:00:00` : rightKey);
  if (Number.isNaN(leftDate.getTime()) || Number.isNaN(rightDate.getTime())) return undefined;
  leftDate.setFullYear(new Date().getFullYear());
  rightDate.setFullYear(new Date().getFullYear());
  return Math.round(Math.abs(leftDate.getTime() - rightDate.getTime()) / 86_400_000);
}

function paymentProofTransactionKey(item: ReceiptItem): string | undefined {
  const amount = safeNumber(item.amount);
  const date = parseDateKey(item.transactionDate || item.date);
  if (amount === undefined || !date) return undefined;
  const bank = (item.bankName ?? "").trim().toLowerCase();
  const card = (item.cardLast4 ?? "").trim().toLowerCase();
  const reference = (item.transactionReference ?? "").trim().toLowerCase();
  const vendor = [...normalizedTextTokens(item.place)].slice(0, 4).join("-");
  const fingerprint = item.images.map((image) => image.fingerprint || "").filter(Boolean).sort().join("-");
  if (reference) return `${date}|${amount.toFixed(2)}|${bank}|${card}|${reference}`;
  return `${date}|${amount.toFixed(2)}|${bank}|${card}|${vendor}|${fingerprint}`;
}

export interface UsaReportRow {
  id: string;
  item?: ReceiptItem;
  expenseItem?: ReceiptItem;
  date: string;
  amount: string;
  receiptImages: ImageAttachment[];
  proofImages: ImageAttachment[];
}

const USA_MATCH_STOP_WORDS = new Set([
  "receipt",
  "payment",
  "proof",
  "transaction",
  "details",
  "total",
  "sale",
  "card",
  "credit",
  "debit",
  "visa",
  "mastercard",
  "jpg",
  "jpeg",
  "png",
  "pdf"
]);

function usaMatchTokens(item: ReceiptItem): Set<string> {
  const tokens = new Set<string>();
  for (const token of itemSearchText(item).toLowerCase().match(/[a-z0-9]+/g) ?? []) {
    if (token.length < 3 || USA_MATCH_STOP_WORDS.has(token)) continue;
    tokens.add(token);
  }
  return tokens;
}

function sameDate(left: string, right: string): boolean {
  const leftKey = parseDateKey(left);
  const rightKey = parseDateKey(right);
  return Boolean(leftKey && rightKey && leftKey === rightKey);
}

function amountlessReceiptMatch(item: ReceiptItem, candidates: ReceiptItem[]): ReceiptItem | undefined {
  const sameDatePaid = candidates.filter((candidate) => safeNumber(candidate.amount) !== undefined && sameDate(item.date, candidate.date));
  if (sameDatePaid.length === 1) return sameDatePaid[0];
  if (!sameDatePaid.length) return undefined;
  const itemTokens = usaMatchTokens(item);
  if (!itemTokens.size) return undefined;
  const scored = sameDatePaid
    .map((candidate) => ({ candidate, score: [...usaMatchTokens(candidate)].filter((token) => itemTokens.has(token)).length }))
    .filter((entry) => entry.score > 0)
    .sort((left, right) => right.score - left.score);
  if (!scored.length || scored[0].score < 2) return undefined;
  if (scored.length > 1 && scored[0].score === scored[1].score) return undefined;
  return scored[0].candidate;
}

export function buildUsaReportRows(items: ReceiptItem[], proofs: PaymentProof[] = []): UsaReportRow[] {
  let mergedItems = mergeSameUsaReceipts(items);
  const paidReceipts = mergedItems.filter((item) => isUsaReceiptItem(item) && safeNumber(item.amount) !== undefined);
  const amountlessTargetById = new Map<string, string>();
  for (const item of mergedItems) {
    if (!isUsaReceiptItem(item) || safeNumber(item.amount) !== undefined) continue;
    const target = amountlessReceiptMatch(item, paidReceipts);
    if (!target) continue;
    const existing = new Set(target.images.map((image) => image.id));
    for (const image of item.images) {
      if (!existing.has(image.id)) {
        target.images.push(image);
        existing.add(image.id);
      }
    }
    for (const keyName of ["place", "purpose", "details", "projectNumber", "paymentMethod", "receiptLabel"] as const) {
      if (!target[keyName] && item[keyName]) target[keyName] = item[keyName];
    }
    target.status = "Merged";
    amountlessTargetById.set(item.id, target.id);
  }
  if (amountlessTargetById.size) {
    mergedItems = mergedItems
      .filter((item) => !amountlessTargetById.has(item.id))
      .map((item) => (item.matchedReceiptId && amountlessTargetById.has(item.matchedReceiptId) ? { ...item, matchedReceiptId: amountlessTargetById.get(item.matchedReceiptId)! } : item));
  }
  mergedItems = matchUsaPaymentProofItems(mergedItems);
  const receiptItems = sortReceiptsForReport(mergedItems.filter(isUsaReceiptItem), "USA");
  const rows: UsaReportRow[] = receiptItems.map((item) => ({
    id: item.id,
    item,
    expenseItem: item,
    date: item.date,
    amount: item.amount,
    receiptImages: receiptSideImages(item),
    proofImages: proofSideImages(item)
  }));
  const rowByReceiptId = new Map(rows.map((row) => [row.id, row]));
  const proofCandidates = [
    ...mergedItems.filter((item) => isUsaPaymentProofItem(item) && item.status !== "Duplicate proof").map((item) => ({
      id: item.id,
      date: item.date,
      amount: item.amount,
      matchedReceiptId: item.matchedReceiptId,
      expenseItem: item,
      images: proofSideImages(item)
    })),
    ...proofs.map((proof) => ({
      id: proof.id,
      date: proof.date,
      amount: proof.amount,
      matchedReceiptId: proof.matchedReceiptId,
      expenseItem: {
        ...blankReceipt("USA", proof.image),
        id: proof.id,
        filename: proof.filename,
        status: proof.status,
        date: proof.date,
        place: proof.place,
        amount: proof.amount,
        documentType: "payment_proof" as DocumentType,
        receiptLabel: proof.filename
      },
      images: [proof.image]
    }))
  ];
  const proofOnlyRows: UsaReportRow[] = [];
  for (const proof of proofCandidates) {
    let target = proof.matchedReceiptId ? rowByReceiptId.get(proof.matchedReceiptId) : undefined;
    if (!target) {
      const proofAmount = safeNumber(proof.amount);
      const proofDate = parseDateKey(proof.date);
      const matches = rows.filter((row) => {
        const rowAmount = safeNumber(row.amount);
        return Boolean(
          proofDate &&
            sameMonthDay(row.date, proof.date) &&
            proofAmount !== undefined &&
            rowAmount !== undefined &&
            Math.abs(rowAmount - proofAmount) < 0.01
        );
      });
      if (matches.length === 1) target = matches[0];
    }
    if (target) {
      target.proofImages = sortSourcePages([...target.proofImages, ...proof.images]);
    } else {
      proofOnlyRows.push({
        id: proof.id,
        expenseItem: proof.expenseItem,
        date: proof.date,
        amount: proof.amount,
        receiptImages: [],
        proofImages: [...proof.images]
      });
    }
  }
  return [...rows, ...proofOnlyRows]
    .map((row, index) => ({ row, index, date: dateSortValue(row.date) }))
    .sort((left, right) => {
      if (left.date !== right.date) return left.date - right.date;
      return left.index - right.index;
    })
    .map(({ row }) => row);
}

export function usaExportAuditIssues(items: ReceiptItem[]): Array<{ item: ReceiptItem; reason: string }> {
  const issues: Array<{ item: ReceiptItem; reason: string }> = [];
  const seenImageIds = new Map<string, ReceiptItem>();
  for (const item of items) {
    if (item.status === "Multiple possible matches") {
      issues.push({ item, reason: "Multiple possible matches" });
    }
    const roles = new Set(item.images.map((image) => image.supportRole ?? "unknown"));
    if (item.status === "Filename hint" && ![...roles].some((role) => role === "receipt" || role === "companion_ticket" || role === "payment_proof")) {
      issues.push({ item, reason: "Pages need AI review" });
    }
    if (roles.has("unknown") && (roles.has("receipt") || roles.has("payment_proof"))) {
      issues.push({ item, reason: "Unresolved page role" });
    }
    for (const image of item.images) {
      const previous = seenImageIds.get(image.id);
      if (previous && previous.id !== item.id) {
        issues.push({ item, reason: "Support page appears in multiple entries" });
      } else {
        seenImageIds.set(image.id, item);
      }
    }
  }
  return issues;
}

export function matchPaymentProofs(proofs: PaymentProof[], items: ReceiptItem[]): PaymentProof[] {
  return proofs.map((proof) => {
    const proofAmount = safeNumber(proof.amount);
    const proofDate = parseDateKey(proof.date);
    if (proofAmount === undefined || !proofDate) {
      return { ...proof, matchedReceiptId: "", status: "Missing date/amount" };
    }
    const matches = items.filter((item) => {
      const itemAmount = safeNumber(item.amount);
      return Boolean(
        proofDate &&
          sameMonthDay(item.date, proof.date) &&
          proofAmount !== undefined &&
          itemAmount !== undefined &&
          Math.abs(itemAmount - proofAmount) < 0.01
      );
    });
    if (matches.length === 1) {
      return { ...proof, matchedReceiptId: matches[0].id, status: "Matched" };
    }
    return { ...proof, matchedReceiptId: "", status: matches.length > 1 ? "Multiple possible matches" : "Payment proof" };
  });
}

export function swapProofForReceipt(
  proofs: PaymentProof[],
  receiptId: string,
  currentProofIds: string[]
): { proofs: PaymentProof[]; selectedProofId?: string } {
  const nextProof = proofs.find((proof) => !currentProofIds.includes(proof.id));
  if (!nextProof) {
    return { proofs };
  }
  return {
    selectedProofId: nextProof.id,
    proofs: proofs.map((proof) => {
      if (currentProofIds.includes(proof.id)) return { ...proof, matchedReceiptId: "", status: "Payment proof" };
      if (proof.id === nextProof.id) return { ...proof, matchedReceiptId: receiptId, status: "Matched manually" };
      return proof;
    })
  };
}

export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}
