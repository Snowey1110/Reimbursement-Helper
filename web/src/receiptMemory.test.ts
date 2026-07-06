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

    expect(hinted.status).toBe("Filename hint");
    expect(hinted.amount).toBe("40.00");
  });
});
