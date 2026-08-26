import type { ExchangeRates, ImageAttachment, PaymentProof, ReceiptItem } from "./types";

const DATABASE_NAME = "reimbursement-helper-web";
const DATABASE_VERSION = 1;
const STORE_NAME = "report-drafts";
const CURRENT_DRAFT_KEY = "current";
const DRAFT_VERSION = 1;

export interface ReportDraft {
  version: typeof DRAFT_VERSION;
  savedAt: number;
  formVersion: "USA" | "Korea";
  items: ReceiptItem[];
  proofs: PaymentProof[];
  exchangeRateImages: ImageAttachment[];
  rates: ExchangeRates;
  readyForExport: boolean;
}

export type ReportDraftState = Omit<ReportDraft, "version" | "savedAt">;

let writeQueue: Promise<void> = Promise.resolve();

function indexedDbAvailable(): boolean {
  return typeof indexedDB !== "undefined";
}

function openDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    let request: IDBOpenDBRequest;
    try {
      request = indexedDB.open(DATABASE_NAME, DATABASE_VERSION);
    } catch (error) {
      reject(error);
      return;
    }
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(STORE_NAME)) request.result.createObjectStore(STORE_NAME);
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("Could not open local report storage."));
    request.onblocked = () => reject(new Error("Local report storage is blocked by another tab."));
  });
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object";
}

export function normalizeReportDraft(value: unknown): ReportDraft | undefined {
  if (!isRecord(value) || value.version !== DRAFT_VERSION) return undefined;
  if (value.formVersion !== "USA" && value.formVersion !== "Korea") return undefined;
  if (!Array.isArray(value.items) || !Array.isArray(value.proofs) || !Array.isArray(value.exchangeRateImages)) return undefined;
  if (!isRecord(value.rates)) return undefined;
  const usdToRmb = Number(value.rates.usdToRmb);
  const usdToKrw = Number(value.rates.usdToKrw);
  const krwToRmb = Number(value.rates.krwToRmb);
  if (![usdToRmb, usdToKrw, krwToRmb].every(Number.isFinite)) return undefined;
  return {
    version: DRAFT_VERSION,
    savedAt: Number(value.savedAt) || 0,
    formVersion: value.formVersion,
    items: value.items as ReceiptItem[],
    proofs: value.proofs as PaymentProof[],
    exchangeRateImages: value.exchangeRateImages as ImageAttachment[],
    rates: { usdToRmb, usdToKrw, krwToRmb },
    readyForExport: Boolean(value.readyForExport)
  };
}

export async function loadReportDraft(): Promise<ReportDraft | undefined> {
  if (!indexedDbAvailable()) return undefined;
  const database = await openDatabase();
  try {
    const value = await new Promise<unknown>((resolve, reject) => {
      const transaction = database.transaction(STORE_NAME, "readonly");
      const request = transaction.objectStore(STORE_NAME).get(CURRENT_DRAFT_KEY);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error ?? new Error("Could not read the local report draft."));
    });
    return normalizeReportDraft(value);
  } finally {
    database.close();
  }
}

function queueWrite(operation: () => Promise<void>): Promise<void> {
  const next = writeQueue.catch(() => undefined).then(operation);
  writeQueue = next;
  return next;
}

export function saveReportDraft(state: ReportDraftState): Promise<void> {
  if (!indexedDbAvailable()) return Promise.resolve();
  const draft: ReportDraft = { ...state, version: DRAFT_VERSION, savedAt: Date.now() };
  return queueWrite(async () => {
    const database = await openDatabase();
    try {
      await new Promise<void>((resolve, reject) => {
        const transaction = database.transaction(STORE_NAME, "readwrite");
        transaction.objectStore(STORE_NAME).put(draft, CURRENT_DRAFT_KEY);
        transaction.oncomplete = () => resolve();
        transaction.onerror = () => reject(transaction.error ?? new Error("Could not save the local report draft."));
        transaction.onabort = () => reject(transaction.error ?? new Error("Saving the local report draft was cancelled."));
      });
    } finally {
      database.close();
    }
  });
}

export function clearReportDraft(): Promise<void> {
  if (!indexedDbAvailable()) return Promise.resolve();
  return queueWrite(async () => {
    const database = await openDatabase();
    try {
      await new Promise<void>((resolve, reject) => {
        const transaction = database.transaction(STORE_NAME, "readwrite");
        transaction.objectStore(STORE_NAME).delete(CURRENT_DRAFT_KEY);
        transaction.oncomplete = () => resolve();
        transaction.onerror = () => reject(transaction.error ?? new Error("Could not clear the local report draft."));
        transaction.onabort = () => reject(transaction.error ?? new Error("Clearing the local report draft was cancelled."));
      });
    } finally {
      database.close();
    }
  });
}
