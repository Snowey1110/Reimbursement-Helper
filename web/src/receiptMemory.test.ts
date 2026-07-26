import { beforeEach, describe, expect, it } from "vitest";
import { applyStoredReceiptMemory, rememberReceiptItem } from "./receiptMemory";
import { image, receipt } from "./test/factories";

describe("receipt memory", () => {
  beforeEach(() => {
    localStorage.clear();
  });

  it("restores manually reviewed details for the same uploaded file fingerprint", () => {
    const fingerprint = "abc123";
    const reviewed = rememberReceiptItem(
      receipt({
        id: "reviewed",
        filename: "fuel.png",
        date: "2026-06-26",
        place: "Shell",
        amount: "40.00",
        purpose: "Gas",
        details: "Actual fuel purchase",
        documentType: "receipt",
        lockedDocumentType: "receipt",
        manualFields: ["amount"],
        images: [image({ id: "source", fingerprint })]
      }),
      "USA",
      "manual"
    );

    expect(reviewed.filename).toContain("__RH1_");
    const restored = applyStoredReceiptMemory(
      receipt({
        id: "new-upload",
        filename: "fuel.png",
        date: "",
        place: "",
        amount: "",
        purpose: "",
        details: "",
        documentType: "unknown",
        images: [image({ id: "source-again", fingerprint })]
      }),
      "USA"
    );

    expect(restored.status).toBe("Manual remembered");
    expect(restored.amount).toBe("40.00");
    expect(restored.place).toBe("Shell");
    expect(restored.documentType).toBe("receipt");
    expect(restored.lockedDocumentType).toBe("receipt");
    expect(restored.manualFields).toContain("amount");

    localStorage.clear();
    const hinted = applyStoredReceiptMemory(
      receipt({
        id: "filename-hint",
        filename: reviewed.filename,
        date: "",
        place: "",
        amount: "",
        purpose: "",
        details: "",
        documentType: "unknown"
      }),
      "USA"
    );

    expect(hinted.status).toBe("Needs regen");
    expect(hinted.amount).toBe("40.00");
    expect(hinted.documentType).toBe("receipt");
    expect(hinted.lockedDocumentType).toBeUndefined();
    expect(hinted.images[0]?.roleSource).not.toBe("manual");
  });

  it("does not restore legacy page roles as manual locks", () => {
    localStorage.setItem(
      "reimbursement-helper-receipt-memory-v1",
      JSON.stringify({
        version: 2,
        records: {
          "sha256:legacy-page:page:": {
            reviewedSource: "manual",
            updatedAt: "2026-07-23T08:25:33",
            formVersion: "USA",
            documentType: "payment_proof",
            date: "2026-07-06",
            amount: "9.00",
            currency: "USD",
            pageStates: [
              {
                supportRole: "payment_proof",
                roleSource: "manual",
                pageAmount: "9.00"
              }
            ]
          }
        }
      })
    );

    const restored = applyStoredReceiptMemory(
      receipt({
        documentType: "unknown",
        images: [image({ id: "legacy-image", fingerprint: "legacy-page" })]
      }),
      "USA"
    );

    expect(restored.status).toBe("Needs regen");
    expect(restored.lockedDocumentType).toBeUndefined();
    expect(restored.images[0].supportRole).toBeUndefined();
    expect(restored.images[0].roleSource).toBeUndefined();
    expect(restored.images[0].pageAmount).toBeUndefined();
  });
});
