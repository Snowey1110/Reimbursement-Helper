import type { FormVersion, ReceiptItem } from "./types";
import { normalizeCategory, normalizeCurrency, normalizeDocumentType, updateAmounts } from "./utils";
import { DEFAULT_KRW_TO_RMB, DEFAULT_USD_TO_KRW, DEFAULT_USD_TO_RMB } from "./constants";

const RECEIPT_MEMORY_STORAGE_KEY = "reimbursement-helper-receipt-memory-v1";
const RECEIPT_MEMORY_MARKER = /__RH1_[A-Za-z0-9._~-]+__/;
const MAX_MEMORY_RECORDS = 2000;

type ReviewedSource = "ai" | "manual";

interface ReceiptMemoryRecord {
  version?: number;
  reviewedSource: ReviewedSource;
  updatedAt: string;
  formVersion: FormVersion;
  documentType: string;
  lockedDocumentType?: string;
  date: string;
  place: string;
  amount: string;
  currency: string;
  krwAmount: string;
  rmbAmount: string;
  purpose: string;
  details: string;
  projectNumber: string;
  category: string;
  paymentMethod: string;
  receiptLabel: string;
  documentEvidence?: string;
  transactionDate?: string;
  postedDate?: string;
  bankName?: string;
  cardLast4?: string;
  transactionReference?: string;
  manualFields?: string[];
  pageStates?: Array<{
    sourcePage?: number;
    supportRole?: string;
    roleSource?: string;
    rotationDegrees?: number;
    rotationLocked?: boolean;
    documentEvidence?: string;
    transactionDate?: string;
    postedDate?: string;
    bankName?: string;
    cardLast4?: string;
    transactionReference?: string;
    pageAmount?: string;
    pageVendor?: string;
  }>;
}

interface ReceiptMemoryStore {
  version: 2;
  records: Record<string, ReceiptMemoryRecord>;
}

function loadStore(): ReceiptMemoryStore {
  try {
    const parsed = JSON.parse(localStorage.getItem(RECEIPT_MEMORY_STORAGE_KEY) || "");
    if (parsed && typeof parsed === "object" && parsed.records && typeof parsed.records === "object") {
      return { version: 2, records: parsed.records };
    }
  } catch {
    // Storage can be disabled in private browsing; ignore and behave as no memory.
  }
  return { version: 2, records: {} };
}

function saveStore(store: ReceiptMemoryStore) {
  try {
    const entries = Object.entries(store.records).sort((left, right) => String(right[1].updatedAt).localeCompare(String(left[1].updatedAt)));
    store.records = Object.fromEntries(entries.slice(0, MAX_MEMORY_RECORDS));
    localStorage.setItem(RECEIPT_MEMORY_STORAGE_KEY, JSON.stringify(store));
  } catch {
    // Ignore storage quota/private browsing failures.
  }
}

function safeToken(value: unknown, limit = 28): string {
  return String(value ?? "")
    .trim()
    .replace(/[^A-Za-z0-9]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, limit)
    .replace(/-$/g, "");
}

function stripMarker(filename: string): string {
  return filename.replace(RECEIPT_MEMORY_MARKER, "").replace(/[_\s-]+(\.[^.]+)?$/, "$1") || "receipt";
}

function splitFilename(filename: string): { stem: string; ext: string } {
  const match = /^(.*?)(\.[^.]+)?$/.exec(filename);
  return { stem: match?.[1] || "receipt", ext: match?.[2] || "" };
}

function memoryKeysForItem(item: ReceiptItem): string[] {
  const keys: string[] = [];
  for (const image of item.images) {
    if (!image.fingerprint) continue;
    const page = image.sourcePage ? String(image.sourcePage) : "";
    const key = `sha256:${image.fingerprint}:page:${page}`;
    if (!keys.includes(key)) keys.push(key);
    const fileKey = `sha256:${image.fingerprint}:page:`;
    if (!keys.includes(fileKey)) keys.push(fileKey);
  }
  return keys;
}

function markerFromRecord(record: ReceiptMemoryRecord, fingerprint = ""): string {
  const amountToken = record.amount.replace(/[^0-9.-]+/g, "").replace(".", "p");
  const parts = [
    record.documentType !== "unknown" ? `t-${safeToken(record.documentType)}` : "",
    record.date ? `d-${safeToken(record.date, 10)}` : "",
    record.amount ? `a-${safeToken(`${amountToken}${record.currency}`, 18)}` : "",
    record.category ? `c-${safeToken(record.category, 20)}` : "",
    record.place ? `v-${safeToken(record.place, 24)}` : "",
    record.reviewedSource ? `r-${safeToken(record.reviewedSource, 8)}` : "",
    fingerprint ? `f-${safeToken(fingerprint, 12)}` : ""
  ].filter(Boolean);
  return `__RH1_${parts.slice(0, 7).join("_")}__`;
}

function parseMarker(value: string): Partial<ReceiptMemoryRecord> | undefined {
  const match = RECEIPT_MEMORY_MARKER.exec(value);
  if (!match) return undefined;
  const body = match[0].slice("__RH1_".length, -2);
  const record: Partial<ReceiptMemoryRecord> = {};
  for (const piece of body.split("_")) {
    const [key, ...rest] = piece.split("-");
    const raw = rest.join("-");
    if (!key || !raw) continue;
    if (key === "t") record.documentType = normalizeDocumentType(raw, "unknown");
    if (key === "d") record.date = raw;
    if (key === "a") {
      const amount = /^(-?\d+(?:[.p-]\d+)?)([A-Z]+)?/.exec(raw);
      if (amount) {
        let amountText = amount[1].replace("p", ".");
        if (!amountText.includes(".") && amountText.slice(1).includes("-")) {
          const sign = amountText.startsWith("-") ? "-" : "";
          const body = sign ? amountText.slice(1) : amountText;
          amountText = `${sign}${body.replace("-", ".")}`;
        }
        record.amount = amountText;
        record.currency = normalizeCurrency(amount[2], "USD");
      }
    }
    if (key === "c") record.category = raw;
    if (key === "v") record.place = raw.replace(/-/g, " ");
    if (key === "r") record.reviewedSource = raw === "manual" ? "manual" : "ai";
  }
  return record;
}

function applyRecord(item: ReceiptItem, record: Partial<ReceiptMemoryRecord>, formVersion: FormVersion, fullRecord: boolean): ReceiptItem {
  const manual = record.reviewedSource === "manual";
  const currentRecord = fullRecord && Number(record.version ?? 1) >= 2;
  const lockedType = normalizeDocumentType(record.lockedDocumentType, "unknown");
  const pageStates = Array.isArray(record.pageStates) ? record.pageStates : [];
  const next: ReceiptItem = {
    ...item,
    date: record.date ?? item.date,
    place: record.place ?? item.place,
    amount: record.amount ?? item.amount,
    currency: normalizeCurrency(record.currency ?? item.currency, formVersion === "USA" ? "USD" : item.currency),
    krwAmount: record.krwAmount ?? item.krwAmount,
    rmbAmount: record.rmbAmount ?? item.rmbAmount,
    purpose: record.purpose ?? item.purpose,
    details: record.details ?? item.details,
    projectNumber: record.projectNumber ?? item.projectNumber,
    category: normalizeCategory(record.category ?? item.category),
    paymentMethod: record.paymentMethod ?? item.paymentMethod,
    receiptLabel: record.receiptLabel ?? item.receiptLabel,
    documentType: formVersion === "USA" ? normalizeDocumentType(record.documentType, item.documentType) : item.documentType,
    lockedDocumentType:
      formVersion === "USA" && currentRecord && lockedType !== "unknown"
        ? lockedType
        : undefined,
    manualFields: currentRecord ? record.manualFields ?? item.manualFields : [],
    documentEvidence: (record.documentEvidence as ReceiptItem["documentEvidence"]) ?? item.documentEvidence,
    transactionDate: record.transactionDate ?? item.transactionDate,
    postedDate: record.postedDate ?? item.postedDate,
    bankName: record.bankName ?? item.bankName,
    cardLast4: record.cardLast4 ?? item.cardLast4,
    transactionReference: record.transactionReference ?? item.transactionReference,
    images: item.images.map((image, index) => {
      const state =
        pageStates.find((candidate) => candidate.sourcePage !== undefined && candidate.sourcePage === image.sourcePage) ??
        pageStates[index];
      if (!state) {
        return image;
      }
      return {
        ...image,
        supportRole: currentRecord ? (state.supportRole as typeof image.supportRole) ?? image.supportRole : image.supportRole,
        roleSource: currentRecord
          ? (state.roleSource as typeof image.roleSource) ?? (manual ? "memory" : image.roleSource)
          : image.roleSource,
        rotationDegrees: state.rotationDegrees ?? image.rotationDegrees,
        rotationLocked: state.rotationLocked ?? image.rotationLocked,
        documentEvidence: currentRecord
          ? (state.documentEvidence as typeof image.documentEvidence) ?? image.documentEvidence
          : image.documentEvidence,
        transactionDate: currentRecord ? state.transactionDate ?? image.transactionDate : image.transactionDate,
        postedDate: currentRecord ? state.postedDate ?? image.postedDate : image.postedDate,
        bankName: currentRecord ? state.bankName ?? image.bankName : image.bankName,
        cardLast4: currentRecord ? state.cardLast4 ?? image.cardLast4 : image.cardLast4,
        transactionReference: currentRecord ? state.transactionReference ?? image.transactionReference : image.transactionReference,
        pageAmount: currentRecord ? state.pageAmount ?? image.pageAmount : image.pageAmount,
        pageVendor: currentRecord ? state.pageVendor ?? image.pageVendor : image.pageVendor
      };
    }),
    status: currentRecord
      ? manual ? "Manual remembered" : "Remembered"
      : "Needs regen"
  };
  return updateAmounts(next, { usdToRmb: DEFAULT_USD_TO_RMB, usdToKrw: DEFAULT_USD_TO_KRW, krwToRmb: DEFAULT_KRW_TO_RMB }, "amount");
}

export function applyStoredReceiptMemory(item: ReceiptItem, formVersion: FormVersion): ReceiptItem {
  const store = loadStore();
  for (const key of memoryKeysForItem(item)) {
    const record = store.records[key];
    if (record) return applyRecord(item, record, formVersion, true);
  }
  const values = [item.filename, ...item.images.flatMap((image) => [image.filename, image.sourceName])];
  for (const value of values) {
    const marker = parseMarker(value);
    if (marker) return applyRecord(item, marker, formVersion, false);
  }
  return item;
}

export function receiptMemoryRecordFromItem(item: ReceiptItem, formVersion: FormVersion, reviewedSource: ReviewedSource): ReceiptMemoryRecord {
  return {
    version: 2,
    reviewedSource,
    updatedAt: new Date().toISOString(),
    formVersion,
    documentType: normalizeDocumentType(item.documentType, "unknown"),
    lockedDocumentType: item.lockedDocumentType,
    date: item.date.trim(),
    place: item.place.trim(),
    amount: item.amount.trim(),
    currency: normalizeCurrency(item.currency, formVersion === "USA" ? "USD" : "KRW"),
    krwAmount: item.krwAmount.trim(),
    rmbAmount: item.rmbAmount.trim(),
    purpose: item.purpose.trim(),
    details: item.details.trim(),
    projectNumber: item.projectNumber.trim(),
    category: normalizeCategory(item.category),
    paymentMethod: item.paymentMethod.trim(),
    receiptLabel: item.receiptLabel.trim(),
    documentEvidence: item.documentEvidence,
    transactionDate: item.transactionDate?.trim(),
    postedDate: item.postedDate?.trim(),
    bankName: item.bankName?.trim(),
    cardLast4: item.cardLast4?.trim(),
    transactionReference: item.transactionReference?.trim(),
    manualFields: item.manualFields ?? [],
    pageStates: item.images.map((image) => ({
      sourcePage: image.sourcePage,
      supportRole: image.supportRole,
      roleSource: image.roleSource,
      rotationDegrees: image.rotationDegrees,
      rotationLocked: image.rotationLocked,
      documentEvidence: image.documentEvidence,
      transactionDate: image.transactionDate,
      postedDate: image.postedDate,
      bankName: image.bankName,
      cardLast4: image.cardLast4,
      transactionReference: image.transactionReference,
      pageAmount: image.pageAmount,
      pageVendor: image.pageVendor
    }))
  };
}

export function rememberReceiptItem(item: ReceiptItem, formVersion: FormVersion, reviewedSource: ReviewedSource): ReceiptItem {
  const record = receiptMemoryRecordFromItem(item, formVersion, reviewedSource);
  const keys = memoryKeysForItem(item);
  if (!keys.length) return item;
  const hasValue = Boolean(record.date || record.amount || record.krwAmount || record.rmbAmount || record.place || record.purpose || record.details || record.receiptLabel || record.documentType !== "unknown");
  if (!hasValue) return item;
  const store = loadStore();
  for (const key of keys) {
    store.records[key] = record;
  }
  saveStore(store);
  const firstFingerprint = item.images.find((image) => image.fingerprint)?.fingerprint?.slice(0, 12) ?? "";
  const { stem, ext } = splitFilename(item.filename);
  const marker = markerFromRecord(record, firstFingerprint);
  const filename = `${stripMarker(stem).slice(0, Math.max(12, 180 - marker.length - ext.length))}${marker}${ext}`;
  return { ...item, filename };
}
