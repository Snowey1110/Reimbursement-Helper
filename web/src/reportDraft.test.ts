// @vitest-environment node

import { IDBFactory, IDBObjectStore } from "fake-indexeddb";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { ReportDraftState } from "./reportDraft";
import { image, proof, receipt } from "./test/factories";

function report(overrides: Partial<ReportDraftState> = {}): ReportDraftState {
  return {
    formVersion: "USA",
    items: [
      receipt({
        id: "synthetic-receipt",
        status: "Manual reviewed",
        place: "Example merchant",
        amount: "12.50",
        purpose: "Example trip",
        details: "Synthetic test receipt",
        projectNumber: "TEST-001",
        lockedDocumentType: "receipt",
        manualFields: ["amount", "place"],
        images: [
          image({
            id: "synthetic-receipt-front",
            dataUrl: "data:image/png;base64,cmVjZWlwdC1mcm9udA==",
            fingerprint: "synthetic-fingerprint",
            supportRole: "receipt",
            roleSource: "manual",
            documentEvidence: "photographed_paper",
            rotationDegrees: 90,
            rotationLocked: true,
            cropPoints: [{ x: 2, y: 3 }, { x: 90, y: 4 }, { x: 91, y: 180 }, { x: 1, y: 181 }]
          }),
          image({
            id: "synthetic-receipt-back",
            dataUrl: "data:image/png;base64,cmVjZWlwdC1iYWNr",
            supportRole: "companion_ticket",
            sourceName: "synthetic-receipt.pdf",
            sourcePage: 2,
            pageCount: 2,
            isPdfPage: true
          })
        ]
      })
    ],
    proofs: [
      proof({
        id: "synthetic-proof",
        amount: "12.50",
        matchedReceiptId: "synthetic-receipt",
        image: image({
          id: "synthetic-proof-image",
          dataUrl: "data:image/png;base64,cGF5bWVudC1wcm9vZg==",
          supportRole: "payment_proof",
          documentEvidence: "bank_portal",
          transactionDate: "2026-06-19",
          postedDate: "2026-06-20",
          bankName: "Example Bank",
          cardLast4: "0000",
          transactionReference: "SYNTHETIC-REFERENCE",
          pageAmount: "12.50",
          pageVendor: "Example merchant"
        })
      })
    ],
    exchangeRateImages: [image({
      id: "synthetic-exchange-rate",
      dataUrl: "data:image/png;base64,ZXhjaGFuZ2UtcmF0ZQ=="
    })],
    rates: { usdToRmb: 7.1, usdToKrw: 1400, krwToRmb: 0.0051 },
    readyForExport: true,
    ...overrides
  };
}

function storeRawDraft(value: unknown): Promise<void> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open("reimbursement-helper-web", 1);
    request.onupgradeneeded = () => request.result.createObjectStore("report-drafts");
    request.onerror = () => reject(request.error);
    request.onsuccess = () => {
      const database = request.result;
      const transaction = database.transaction("report-drafts", "readwrite");
      transaction.objectStore("report-drafts").put(value, "current");
      transaction.oncomplete = () => {
        database.close();
        resolve();
      };
      transaction.onabort = transaction.onerror = () => {
        database.close();
        reject(transaction.error);
      };
    };
  });
}

describe("report draft storage", () => {
  let storage: typeof import("./reportDraft");

  beforeEach(async () => {
    vi.stubGlobal("indexedDB", new IDBFactory());
    vi.resetModules();
    storage = await import("./reportDraft");
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("returns no draft for a new database", async () => {
    await expect(storage.loadReportDraft()).resolves.toBeUndefined();
  });

  it.each([
    ["unsupported version", { ...report(), version: 999, savedAt: 1 }],
    ["missing report fields", { version: 1, savedAt: 1, formVersion: "USA" }],
    ["invalid exchange rates", { ...report(), version: 1, savedAt: 1, rates: { usdToRmb: "invalid" } }],
    ["null stored value", null],
    ["null receipt", { ...report(), version: 1, savedAt: 1, items: [null] }],
    ["null proof image", { ...report(), version: 1, savedAt: 1, proofs: [{ ...proof(), image: null }] }],
    ["null exchange-rate image", { ...report(), version: 1, savedAt: 1, exchangeRateImages: [null] }]
  ])("rejects a stored %s instead of treating it as an absent draft", async (_description, value) => {
    await storeRawDraft(value);

    await expect(storage.loadReportDraft()).rejects.toThrow();
    // A failed recovery must not silently remove the unrecoverable stored data.
    await expect(storage.loadReportDraft()).rejects.toThrow();
  });

  it.each(["USA", "Korea"] as const)("persists the full %s report across a module reload", async (formVersion) => {
    const state = report({ formVersion });
    const savedAt = 1_780_000_000_000;
    vi.spyOn(Date, "now").mockReturnValue(savedAt);

    await storage.saveReportDraft(state);
    vi.resetModules();
    const reloadedStorage = await import("./reportDraft");
    const restored = await reloadedStorage.loadReportDraft();

    expect(restored).toEqual({ ...state, version: 1, savedAt });
    expect(restored?.items).not.toBe(state.items);
    expect(restored?.proofs[0].image).not.toBe(state.proofs[0].image);
    expect(restored?.exchangeRateImages).not.toBe(state.exchangeRateImages);

    // Reading or editing a restored object must not mutate the stored snapshot.
    restored!.items[0].amount = "999";
    restored!.items[0].images[0].dataUrl = "data:image/png;base64,Y2hhbmdlZA==";
    await expect(reloadedStorage.loadReportDraft()).resolves.toEqual({ ...state, version: 1, savedAt });
  });

  it("replaces the previous draft with the latest complete state", async () => {
    await storage.saveReportDraft(report());
    const replacement = report({
      formVersion: "Korea",
      items: [],
      proofs: [],
      exchangeRateImages: [],
      readyForExport: false,
      rates: { usdToRmb: 7.2, usdToKrw: 1450, krwToRmb: 0.005 }
    });

    await storage.saveReportDraft(replacement);

    await expect(storage.loadReportDraft()).resolves.toEqual({
      ...replacement,
      version: 1,
      savedAt: expect.any(Number)
    });
  });

  it("clears only after all previously queued saves finish", async () => {
    const events: string[] = [];
    const originalPut = IDBObjectStore.prototype.put;
    const originalDelete = IDBObjectStore.prototype.delete;
    vi.spyOn(IDBObjectStore.prototype, "put").mockImplementation(function (this: IDBObjectStore, value, key) {
      events.push("save started");
      this.transaction.addEventListener("complete", () => events.push("save committed"));
      return originalPut.call(this, value, key);
    });
    vi.spyOn(IDBObjectStore.prototype, "delete").mockImplementation(function (this: IDBObjectStore, key) {
      events.push("clear started");
      return originalDelete.call(this, key);
    });

    const firstSave = storage.saveReportDraft(report());
    const secondSave = storage.saveReportDraft(report({ formVersion: "Korea" }));
    const clear = storage.clearReportDraft();
    await Promise.all([firstSave, secondSave, clear]);

    expect(events).toEqual([
      "save started", "save committed", "save started", "save committed", "clear started"
    ]);
    await expect(storage.loadReportDraft()).resolves.toBeUndefined();
  });

  it("preserves a new report saved after a queued clear", async () => {
    const newReport = report({ formVersion: "Korea", readyForExport: false });

    const oldSave = storage.saveReportDraft(report());
    const clear = storage.clearReportDraft();
    const newSave = storage.saveReportDraft(newReport);
    await Promise.all([oldSave, clear, newSave]);

    await expect(storage.loadReportDraft()).resolves.toEqual({
      ...newReport,
      version: 1,
      savedAt: expect.any(Number)
    });
  });

  it("does not let an aborted write poison the next queued save", async () => {
    const originalPut = IDBObjectStore.prototype.put;
    vi.spyOn(IDBObjectStore.prototype, "put").mockImplementationOnce(function (this: IDBObjectStore, value, key) {
      const request = originalPut.call(this, value, key);
      this.transaction.abort();
      return request;
    });
    const nextReport = report({ formVersion: "Korea" });

    const failedSave = expect(storage.saveReportDraft(report())).rejects.toThrow();
    const nextSave = storage.saveReportDraft(nextReport);
    await Promise.all([failedSave, nextSave]);

    await expect(storage.loadReportDraft()).resolves.toEqual({
      ...nextReport,
      version: 1,
      savedAt: expect.any(Number)
    });
  });

  it("still clears an existing draft after a queued write cannot open storage", async () => {
    await storage.saveReportDraft(report());
    const failure = new DOMException("Storage is unavailable", "SecurityError");
    vi.spyOn(indexedDB, "open").mockImplementationOnce(() => { throw failure; });

    const failedSave = expect(storage.saveReportDraft(report({ formVersion: "Korea" }))).rejects.toBe(failure);
    const clear = storage.clearReportDraft();
    await Promise.all([failedSave, clear]);

    await expect(storage.loadReportDraft()).resolves.toBeUndefined();
  });

  it("gracefully skips persistence when IndexedDB is not available", async () => {
    vi.stubGlobal("indexedDB", undefined);

    await expect(storage.loadReportDraft()).resolves.toBeUndefined();
    await expect(storage.saveReportDraft(report())).resolves.toBeUndefined();
    await expect(storage.clearReportDraft()).resolves.toBeUndefined();
  });

  it.each(["load", "save", "clear"] as const)("reports storage access errors during %s", async (operation) => {
    const failure = new DOMException("Storage is disabled", "SecurityError");
    vi.spyOn(indexedDB, "open").mockImplementationOnce(() => { throw failure; });

    const result = operation === "load" ? storage.loadReportDraft()
      : operation === "save" ? storage.saveReportDraft(report())
      : storage.clearReportDraft();

    await expect(result).rejects.toBe(failure);
  });
});
