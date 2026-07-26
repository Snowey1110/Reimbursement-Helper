import { describe, expect, it } from "vitest";
import { image, proof, receipt } from "./test/factories";
import {
  buildUsaReportRows,
  matchPaymentProofs,
  mergeSameUsaReceipts,
  normalizeCurrency,
  normalizeUsaTransportationFields,
  receiptNeedsAi,
  reportCategoryForItem,
  swapProofForReceipt,
  sortReceiptsForReport,
  updateAmounts,
  usaExportAuditIssues
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

  it("keeps ambiguous proof matches visible without blocking export", () => {
    const matched = matchPaymentProofs([proof({ id: "p1" })], [receipt({ id: "r1" }), receipt({ id: "r2" })]);

    expect(matched[0].matchedReceiptId).toBe("");
    expect(matched[0].status).toBe("Multiple possible matches");
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

  it("builds USA rows with blank receipt or proof sides when one side is missing", () => {
    const rows = buildUsaReportRows([
      receipt({ id: "receipt-only", date: "2026-06-19", amount: "27", documentType: "receipt", images: [image({ id: "receipt-img" })] }),
      receipt({ id: "proof-only", date: "2026-06-20", amount: "13.99", documentType: "payment_proof", images: [image({ id: "proof-img" })] })
    ]);

    expect(rows).toHaveLength(2);
    expect(rows[0].receiptImages.map((entry) => entry.id)).toEqual(["receipt-img"]);
    expect(rows[0].proofImages).toEqual([]);
    expect(rows[1].receiptImages).toEqual([]);
    expect(rows[1].proofImages.map((entry) => entry.id)).toEqual(["proof-img"]);
    expect(rows[1].expenseItem?.amount).toBe("13.99");
  });

  it("groups amountless same-date receipt support with the paid receipt", () => {
    const rows = buildUsaReportRows([
      receipt({
        id: "paid",
        date: "2026-06-19",
        place: "McCormick Place Lot A",
        details: "Parking",
        amount: "27",
        documentType: "receipt",
        images: [image({ id: "paid-img" })]
      }),
      receipt({
        id: "ticket",
        date: "2026-06-19",
        place: "McCormick Place Lot A",
        details: "QR parking ticket",
        amount: "",
        documentType: "receipt",
        images: [image({ id: "ticket-img" })]
      })
    ]);

    expect(rows).toHaveLength(1);
    expect(rows[0].receiptImages.map((entry) => entry.id)).toEqual(["paid-img", "ticket-img"]);
  });

  it("matches imported payment proof rows to grouped USA receipts", () => {
    const rows = buildUsaReportRows([
      receipt({ id: "r1", date: "2026-06-19", amount: "27", documentType: "receipt", images: [image({ id: "receipt-img" })] }),
      receipt({ id: "p1", date: "2026-06-19", amount: "27.00", documentType: "payment_proof", images: [image({ id: "proof-img" })] })
    ]);

    expect(rows).toHaveLength(1);
    expect(rows[0].receiptImages.map((entry) => entry.id)).toEqual(["receipt-img"]);
    expect(rows[0].proofImages.map((entry) => entry.id)).toEqual(["proof-img"]);
  });

  it("does not merge same-date same-amount receipts without shared transaction evidence", () => {
    const merged = mergeSameUsaReceipts([
      receipt({ id: "north", date: "2026-06-19", amount: "27", place: "North Garage", details: "North entrance" }),
      receipt({ id: "south", date: "2026-06-19", amount: "27", place: "South Garage", details: "South entrance" })
    ]);

    expect(merged.map((item) => item.id)).toEqual(["north", "south"]);
  });

  it("merges a companion ticket into its paid toll receipt", () => {
    const merged = mergeSameUsaReceipts([
      receipt({
        id: "ticket",
        date: "2026-06-26",
        amount: "2.50",
        place: "Indiana Toll Road",
        details: "Entry Portage ticket",
        images: [image({ id: "ticket-image", supportRole: "companion_ticket" })]
      }),
      receipt({
        id: "paid",
        date: "2026-06-26",
        amount: "2.50",
        place: "Indiana Toll Road",
        details: "Toll from Portage to LaPorte",
        images: [image({ id: "paid-image", supportRole: "receipt" })]
      })
    ]);

    expect(merged).toHaveLength(1);
    expect(merged[0].images.map((entry) => entry.id)).toEqual(["ticket-image", "paid-image"]);
  });

  it("orders multipage proof images top-to-bottom by source page", () => {
    const rows = buildUsaReportRows([
      receipt({
        id: "proof",
        documentType: "payment_proof",
        date: "2026-06-26",
        amount: "2.50",
        images: [
          image({ id: "page-2", sourceName: "proof.pdf", sourcePage: 2, supportRole: "payment_proof" }),
          image({ id: "page-1", sourceName: "proof.pdf", sourcePage: 1, supportRole: "payment_proof" })
        ]
      })
    ]);

    expect(rows[0].proofImages.map((entry) => entry.id)).toEqual(["page-1", "page-2"]);
  });

  it("blocks an ambiguous proof match instead of guessing", () => {
    const proofItem = receipt({
      id: "proof",
      documentType: "payment_proof",
      date: "2026-06-19",
      amount: "30",
      place: "Chase"
    });
    const matched = buildUsaReportRows([
      receipt({ id: "north", date: "2026-06-19", amount: "30", place: "North Garage" }),
      receipt({ id: "south", date: "2026-06-19", amount: "30", place: "South Garage" }),
      proofItem
    ]);
    const itemsForAudit = [
      receipt({ id: "north", date: "2026-06-19", amount: "30", place: "North Garage" }),
      receipt({ id: "south", date: "2026-06-19", amount: "30", place: "South Garage" }),
      { ...proofItem, status: "Multiple possible matches" }
    ];

    expect(matched).toHaveLength(3);
    expect(usaExportAuditIssues(itemsForAudit).map(({ reason }) => reason)).toContain("Multiple possible matches");
  });

  it("splits mixed receipt entries into receipt and proof columns by page role", () => {
    const rows = buildUsaReportRows([
      receipt({
        id: "mixed",
        date: "2026-06-19",
        amount: "27",
        documentType: "receipt",
        images: [
          image({ id: "chase-page", supportRole: "payment_proof" }),
          image({ id: "receipt-page", supportRole: "receipt" })
        ]
      })
    ]);

    expect(rows).toHaveLength(1);
    expect(rows[0].receiptImages.map((entry) => entry.id)).toEqual(["receipt-page"]);
    expect(rows[0].proofImages.map((entry) => entry.id)).toEqual(["chase-page"]);
  });

  it("forces parking and toll text back to transportation even if AI says meals", () => {
    const item = receipt({
      id: "wabash",
      place: "Wabash Randolph",
      purpose: "Breakfast",
      details: "Breakfast at Wabash Randolph",
      category: "meals",
      receiptLabel: "Parking lots and garages"
    });

    expect(reportCategoryForItem(item, "USA")).toBe("transportation");
    const normalized = normalizeUsaTransportationFields(item);
    expect(normalized.category).toBe("transportation");
    expect(normalized.purpose).toBe("Parking");
    expect(normalized.details).toBe("Parking at Wabash Randolph.");
  });

  it("sorts USA receipt support rows by date after matching proofs", () => {
    const rows = buildUsaReportRows([
      receipt({ id: "late", date: "2026-06-22", amount: "22", documentType: "receipt" }),
      receipt({ id: "proof-only", date: "2026-06-18", amount: "18", documentType: "payment_proof" }),
      receipt({ id: "early", date: "2026-06-19", amount: "19", documentType: "receipt" })
    ]);

    expect(rows.map((row) => row.date)).toEqual(["2026-06-18", "2026-06-19", "2026-06-22"]);
  });

  it("identifies only unprocessed or incomplete rows for Generate Empty", () => {
    expect(receiptNeedsAi(receipt({ status: "AI filled", documentType: "receipt" }), "USA")).toBe(false);
    expect(receiptNeedsAi(receipt({ status: "AI filled", documentType: "unknown" }), "USA")).toBe(true);
    expect(receiptNeedsAi(receipt({ status: "Empty", documentType: "receipt" }), "USA")).toBe(true);
    expect(receiptNeedsAi(receipt({ status: "Needs regen", documentType: "receipt" }), "USA")).toBe(true);
  });

  it("blocks unknown pages and page identity mismatches before USA export", () => {
    const unknown = receipt({
      id: "unknown",
      images: [image({ id: "unknown-page", supportRole: "unknown" })]
    });
    const legacyKnown = receipt({
      id: "legacy-known",
      images: [image({ id: "legacy-page", supportRole: "receipt" })]
    });
    const mismatch = receipt({
      id: "mismatch",
      date: "2026-07-06",
      amount: "9.00",
      images: [
        image({
          id: "fuel-page",
          supportRole: "receipt",
          transactionDate: "2026-06-29",
          pageAmount: "48.18"
        })
      ]
    });

    const issues = usaExportAuditIssues([unknown, legacyKnown, mismatch]);

    expect(issues.filter(({ item }) => item.id === "unknown").map(({ reason }) => reason)).toContain("Pages need AI review");
    expect(issues.filter(({ item }) => item.id === "legacy-known").map(({ reason }) => reason)).toContain("Pages need AI review");
    expect(issues.filter(({ item }) => item.id === "mismatch").map(({ reason }) => reason)).toContain("Page amount does not match row");
    expect(issues.filter(({ item }) => item.id === "mismatch").map(({ reason }) => reason)).toContain("Page date does not match row");

    const manualIssues = usaExportAuditIssues([{ ...mismatch, manualFields: ["amount", "date"] }]);
    expect(manualIssues.map(({ reason }) => reason)).not.toContain("Page amount does not match row");
    expect(manualIssues.map(({ reason }) => reason)).not.toContain("Page date does not match row");
  });
});

describe("report sorting", () => {
  it("groups Korea receipts by invoice kind first and then by date", () => {
    const sorted = sortReceiptsForReport(
      [
        receipt({ id: "usb", category: "materials", purpose: "USB", date: "2026-06-28" }),
        receipt({ id: "fuel-late", category: "transportation", purpose: "Fuel", date: "2026-06-16" }),
        receipt({ id: "esim", category: "transportation", purpose: "eSIM data plan", date: "2026-06-25" }),
        receipt({ id: "fuel-early", category: "transportation", purpose: "Gas", date: "2026-06-10" }),
        receipt({ id: "rental", category: "transportation", purpose: "National car rental", date: "2026-06-16" })
      ],
      "Korea"
    );

    expect(sorted.map((item) => item.id)).toEqual(["rental", "fuel-early", "fuel-late", "esim", "usb"]);
  });

  it("classifies eSIM and data plans as Korea other", () => {
    const item = receipt({ category: "transportation", purpose: "eSIM data plan", details: "internet access" });

    expect(reportCategoryForItem(item, "Korea")).toBe("other");
  });
});
