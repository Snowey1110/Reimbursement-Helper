import { describe, expect, it } from "vitest";
import { receipt } from "./test/factories";
import { koreaReceiptImageSlots, koreaReceiptLastRow, mapKoreaDetailRows, mapUsaExpenseRows } from "./excelExport";

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

  it("keeps Korea receipt images inside page slots without label cells", () => {
    expect(koreaReceiptLastRow(5)).toBe(100);
    expect(koreaReceiptImageSlots(5)).toEqual([
      { cell: "A2", maxWidth: 215, maxHeight: 300 },
      { cell: "D2", maxWidth: 215, maxHeight: 300 },
      { cell: "A26", maxWidth: 215, maxHeight: 300 },
      { cell: "D26", maxWidth: 215, maxHeight: 300 },
      { cell: "A52", maxWidth: 215, maxHeight: 300 }
    ]);
  });
});
