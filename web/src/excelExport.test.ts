import ExcelJS from "exceljs";
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { receipt } from "./test/factories";
import { koreaReceiptImageSlots, koreaReceiptLastRow, koreaReceiptPaymentLabel, koreaReceiptWideSlot, mapKoreaDetailRows, mapUsaExpenseRows } from "./excelExport";

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

  it("truncates Korea KRW values instead of rounding them", () => {
    const rows = mapKoreaDetailRows(
      [receipt({ id: "r1", amount: "1", currency: "USD", category: "transportation" })],
      { usdToRmb: 6.8175, krwToRmb: 0.004402959 }
    );

    expect(rows[0].krw).toBe(1548);
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
    expect(koreaReceiptLastRow(5)).toBe(120);
    expect(koreaReceiptImageSlots(5)).toEqual([
      { labelRange: "A1:D1", labelCell: "A1", imageCell: "A2", maxWidth: 370, maxHeight: 520 },
      { labelRange: "E1:H1", labelCell: "E1", imageCell: "E2", maxWidth: 370, maxHeight: 520 },
      { labelRange: "A31:D31", labelCell: "A31", imageCell: "A32", maxWidth: 370, maxHeight: 520 },
      { labelRange: "E31:H31", labelCell: "E31", imageCell: "E32", maxWidth: 370, maxHeight: 520 },
      { labelRange: "A61:D61", labelCell: "A61", imageCell: "A62", maxWidth: 370, maxHeight: 520 }
    ]);
    expect(koreaReceiptWideSlot(0)).toEqual({ labelRange: "A1:H1", labelCell: "A1", imageCell: "A2", maxWidth: 760, maxHeight: 520 });
  });

  it("builds Korea payment labels from only date, content, and amount", () => {
    expect(koreaReceiptPaymentLabel(0, receipt({ paymentMethod: "Visa", date: "2026-06-19", purpose: "Parking", amount: "27.00", currency: "USD" }))).toBe(
      "2026-06-19 | Parking | 27.00 USD"
    );
  });

  it("ships the Korea web template without a frozen detail pane", async () => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(readFileSync("public/templates/korea_reimbursement_template.xlsx"));
    const details = workbook.worksheets[1];

    expect(details?.views.some((view) => view.state === "frozen")).toBe(false);
  });
});
