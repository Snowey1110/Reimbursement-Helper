import { describe, expect, it } from "vitest";
import { image, proof, receipt } from "./test/factories";
import {
  matchPaymentProofs,
  mergeSameUsaReceipts,
  normalizeCurrency,
  reportCategoryForItem,
  swapProofForReceipt,
  sortReceiptsForReport,
  updateAmounts
} from "./utils";

describe("currency conversion", () => {
  it("updates USD, KRW, and RMB amounts in real time", () => {
    const rates = { usdToRmb: 7, usdToKrw: 1550, krwToRmb: 0.005 };

    expect(updateAmounts(receipt({ amount: "10", currency: "USD" }), rates).krwAmount).toBe("15500");
    expect(updateAmounts(receipt({ amount: "10", currency: "USD" }), rates).rmbAmount).toBe("77.50");
    expect(updateAmounts(receipt({ amount: "1000", currency: "KRW" }), rates).rmbAmount).toBe("5");
    expect(updateAmounts(receipt({ amount: "70", currency: "RMB" }), rates).krwAmount).toBe("14000");
  });

  it("normalizes common currency symbols", () => {
    expect(normalizeCurrency("$13.99", "KRW")).toBe("USD");
    expect(normalizeCurrency("\u20a913990", "USD")).toBe("KRW");
    expect(normalizeCurrency("CNY 70", "USD")).toBe("RMB");
  });
});

describe("USA receipt merging", () => {
  it("merges same-date and same-amount receipts while keeping different screenshots", () => {
    const merged = mergeSameUsaReceipts([
      receipt({ id: "r1", date: "2026-06-19", amount: "27", images: [image({ id: "img-1" })] }),
      receipt({ id: "r2", date: "2026-06-19", amount: "27.00", images: [image({ id: "img-2" })] })
    ]);

    expect(merged).toHaveLength(1);
    expect(merged[0].images.map((entry) => entry.id)).toEqual(["img-1", "img-2"]);
    expect(merged[0].status).toBe("Merged");
  });
});

describe("payment proof matching", () => {
  it("matches one proof to one same-date and same-amount receipt", () => {
    const matched = matchPaymentProofs([proof({ id: "p1" })], [receipt({ id: "r1" })]);

    expect(matched[0].matchedReceiptId).toBe("r1");
    expect(matched[0].status).toBe("Matched");
  });

  it("flags ambiguous proof matches for manual review", () => {
    const matched = matchPaymentProofs([proof({ id: "p1" })], [receipt({ id: "r1" }), receipt({ id: "r2" })]);

    expect(matched[0].matchedReceiptId).toBe("");
    expect(matched[0].status).toBe("Needs manual review");
  });

  it("swaps the selected payment proof manually", () => {
    const result = swapProofForReceipt(
      [proof({ id: "p1", matchedReceiptId: "r1" }), proof({ id: "p2" })],
      "r1",
      ["p1"]
    );

    expect(result.selectedProofId).toBe("p2");
    expect(result.proofs.find((entry) => entry.id === "p1")?.matchedReceiptId).toBe("");
    expect(result.proofs.find((entry) => entry.id === "p2")?.matchedReceiptId).toBe("r1");
  });
});

describe("report sorting", () => {
  it("groups receipts by report category first and then by date", () => {
    const sorted = sortReceiptsForReport(
      [
        receipt({ id: "late-transport", category: "transportation", date: "2026-06-20" }),
        receipt({ id: "early-other", category: "other", date: "2026-06-01" }),
        receipt({ id: "early-transport", category: "transportation", date: "2026-06-10" }),
        receipt({ id: "lodging", category: "lodging", date: "2026-06-05" })
      ],
      "Korea"
    );

    expect(sorted.map((item) => item.id)).toEqual(["early-transport", "late-transport", "lodging", "early-other"]);
  });

  it("classifies eSIM and data plans as Korea other", () => {
    const item = receipt({ category: "transportation", purpose: "eSIM data plan", details: "internet access" });

    expect(reportCategoryForItem(item, "Korea")).toBe("other");
  });
});
