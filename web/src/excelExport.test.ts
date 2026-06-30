import ExcelJS from "exceljs";
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { receipt } from "./test/factories";
import { koreaReceiptImageSlots, koreaReceiptLastRow, koreaReceiptPaymentLabel, mapKoreaDetailRows, mapUsaExpenseRows } from "./excelExport";

describe("Excel row mapping", () => {
  it("maps USA expenses into the existing category rows", () => {
    const rows = mapUsaExpenseRows([
      receipt({ id: "r1", category: "transportation" }),
      receipt({ id: "r2", category: "transportation" }),
      receipt({ id: "r3", category: "lodging" })
    ]);

    expect(rows.map((row) => row.row)).toEqual([7, 8, 49]);
    expect(rows.map((row) => row.category)).toEqual(["transportation", "transportation", "lodging"]);
  });

  it("maps Korea detail rows with converted KRW and RMB amounts", () => {
    const rows = mapKoreaDetailRows(
      [receipt({ id: "r1", amount: "10", currency: "USD", category: "transportation" })],
      { usdToRmb: 7, krwToRmb: 0.005 }
    );

    expect(rows[0].row).toBe(3);
    expect(rows[0].categoryColumn).toBe("F");
    expect(rows[0].krw).toBe(14000);
    expect(rows[0].rmb).toBe(70);
    expect(rows[0].note).toBe("(10 USD)");
  });

  it("keeps Korea meals in column K so later detail columns remain available", () => {
    const rows = mapKoreaDetailRows(
      [receipt({ id: "r1", amount: "13900", currency: "KRW", category: "meals", paymentMethod: "Visa" })],
      { usdToRmb: 7, krwToRmb: 0.005 }
    );

    expect(rows[0].categoryColumn).toBe("K");
    expect(rows[0].item.paymentMethod).toBe("Visa");
  });

  it("keeps Korea receipt payment labels next to their image slots", () => {
    expect(koreaReceiptLastRow(5)).toBe(150);
    expect(koreaReceiptImageSlots(5)).toEqual([
      { labelRange: "A1:E1", labelCell: "A1", imageCell: "A2", maxWidth: 520, maxHeight: 420 },
      { labelRange: "A26:E26", labelCell: "A26", imageCell: "A27", maxWidth: 520, maxHeight: 420 },
      { labelRange: "A51:E51", labelCell: "A51", imageCell: "A52", maxWidth: 520, maxHeight: 420 },
      { labelRange: "A76:E76", labelCell: "A76", imageCell: "A77", maxWidth: 520, maxHeight: 420 },
      { labelRange: "A101:E101", labelCell: "A101", imageCell: "A102", maxWidth: 520, maxHeight: 420 }
    ]);
  });

  it("builds Korea payment labels from payment method, date, and amount", () => {
    expect(koreaReceiptPaymentLabel(0, receipt({ paymentMethod: "Visa", date: "2026-06-19", amount: "27.00", currency: "USD" }))).toBe(
      "Payment 1: Visa | 2026-06-19 | 27.00 USD"
    );
  });

  it("ships the Korea web template without a frozen detail pane", async () => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(readFileSync("public/templates/korea_reimbursement_template.xlsx"));
    const details = workbook.worksheets[1];

    expect(details?.views.some((view) => view.state === "frozen")).toBe(false);
  });
});
