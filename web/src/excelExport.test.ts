import ExcelJS from "exceljs";
import { readFileSync } from "node:fs";
import { afterEach, describe, expect, it, vi } from "vitest";
import { image, receipt } from "./test/factories";
import {
  buildKoreaReceiptBlocks,
  exportKoreaWorkbook,
  koreaDetailContent,
  koreaDetailLocation,
  koreaReceiptFullPageSlot,
  koreaPaymentMethodText,
  koreaReceiptImageSlots,
  koreaReceiptLastRow,
  koreaReceiptPaymentLabel,
  koreaReceiptShouldBeWide,
  koreaReceiptWideSlot,
  mapKoreaDetailRows,
  mapUsaExpenseRows,
  sanitizeDrawingXml,
  sanitizeWorkbookXml
} from "./excelExport";

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("Excel row mapping", () => {
  it("maps USA expenses into the existing category rows", () => {
    const rows = mapUsaExpenseRows([
      receipt({ id: "r1", category: "transportation" }),
      receipt({ id: "r2", category: "transportation" }),
      receipt({ id: "r3", category: "lodging", place: "Hotel", purpose: "Hotel", details: "Hotel stay", receiptLabel: "Hotel receipt" })
    ]);

    expect(rows.map((row) => row.row)).toEqual([7, 8, 49]);
    expect(rows.map((row) => row.category)).toEqual(["transportation", "transportation", "lodging"]);
  });

  it("maps obvious USA parking expenses to transportation even when AI picked meals", () => {
    const rows = mapUsaExpenseRows([receipt({ id: "wabash", place: "Wabash Randolph", category: "meals", purpose: "Breakfast", details: "Parking lots and garages" })]);

    expect(rows[0].row).toBe(7);
    expect(rows[0].category).toBe("transportation");
  });

  it("sanitizes ExcelJS drawing XML so images do not have zero-size transforms", () => {
    const xml =
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<xdr:oneCellAnchor editAs="oneCell"><xdr:from/><xdr:ext cx="12345" cy="67890"/><xdr:pic><xdr:spPr>' +
      '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:extLst><a:ext uri="{test}"/></a:extLst>' +
      "</xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>";

    const sanitized = sanitizeDrawingXml(xml);

    expect(sanitized).toContain('<a:ext cx="12345" cy="67890"/>');
    expect(sanitized).not.toContain("<a:extLst>");
  });

  it("sanitizes ExcelJS print-area defined names back to absolute cell references", () => {
    const xml =
      '<workbook><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">&apos;Expense report&apos;!$A1:$J70</definedName>' +
      '<definedName name="_xlnm.Print_Area" localSheetId="1">&apos;Receipt and Payment of expenses&apos;!$B1:$E27</definedName></definedNames></workbook>';

    const sanitized = sanitizeWorkbookXml(xml);

    expect(sanitized).toContain("&apos;Expense report&apos;!$A$1:$J$70");
    expect(sanitized).toContain("&apos;Receipt and Payment of expenses&apos;!$B$1:$E$27");
    expect(sanitized).not.toContain("$A1");
    expect(sanitized).not.toContain("$B1");
  });

  it("maps Korea detail rows with converted KRW and RMB amounts", () => {
    const rows = mapKoreaDetailRows(
      [receipt({ id: "r1", amount: "10", currency: "USD", category: "transportation" })],
      { usdToRmb: 7, usdToKrw: 1400, krwToRmb: 0.005 }
    );

    expect(rows[0].row).toBe(3);
    expect(rows[0].categoryColumn).toBe("F");
    expect(rows[0].krw).toBe(14000);
    expect(rows[0].rmb).toBe(70);
    expect(rows[0].note).toBe("10 USD");
  });

  it("truncates Korea KRW values instead of rounding them", () => {
    const rows = mapKoreaDetailRows(
      [receipt({ id: "r1", amount: "1", currency: "USD", category: "transportation" })],
      { usdToRmb: 6.8175, usdToKrw: 1548.86, krwToRmb: 0.004402959 }
    );

    expect(rows[0].krw).toBe(1548);
  });

  it("keeps Korea meals in column K so later detail columns remain available", () => {
    const rows = mapKoreaDetailRows(
      [receipt({ id: "r1", amount: "13900", currency: "KRW", category: "meals", paymentMethod: "Visa" })],
      { usdToRmb: 7, usdToKrw: 1400, krwToRmb: 0.005 }
    );

    expect(rows[0].categoryColumn).toBe("K");
    expect(rows[0].item.paymentMethod).toBe("Visa");
  });

  it("sorts Korea detail rows by invoice kind and then date", () => {
    const rows = mapKoreaDetailRows(
      [
        receipt({ id: "usb", category: "materials", purpose: "USB", date: "2026-06-28" }),
        receipt({ id: "fuel", category: "transportation", purpose: "Fuel", date: "2026-06-10" }),
        receipt({ id: "rental", category: "transportation", purpose: "National car rental", date: "2026-06-16" })
      ],
      { usdToRmb: 7, usdToKrw: 1400, krwToRmb: 0.005 }
    );

    expect(rows.map((row) => row.item.id)).toEqual(["rental", "fuel", "usb"]);
  });

  it("exports eSIM-like Korea expenses to other column O", () => {
    const rows = mapKoreaDetailRows(
      [receipt({ id: "esim", category: "transportation", purpose: "eSIM data plan", details: "internet access", amount: "1000", currency: "KRW" })],
      { usdToRmb: 7, usdToKrw: 1400, krwToRmb: 0.005 }
    );

    expect(rows[0].category).toBe("other");
    expect(rows[0].categoryColumn).toBe("O");
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
    expect(koreaReceiptFullPageSlot(0)).toEqual({ labelRange: "A1:H1", labelCell: "A1", imageCell: "A2", maxWidth: 760, maxHeight: 1040 });
  });

  it("uses full printable receipt pages for grouped PDF page images", async () => {
    const blocks = await buildKoreaReceiptBlocks([
      receipt({
        id: "pdf",
        images: [
          image({ id: "pdf-page-1", width: 1200, height: 1500, isPdfPage: true, sourcePage: 1, pageCount: 2 }),
          image({ id: "pdf-page-2", width: 1200, height: 1500, isPdfPage: true, sourcePage: 2, pageCount: 2 })
        ]
      })
    ]);

    expect(blocks.map((block) => block.mode)).toEqual(["full", "full"]);
    expect(blocks.map((block) => block.images[0].sourcePage)).toEqual([1, 2]);
  });

  it("builds Korea payment labels from only date, content, and amount", () => {
    expect(koreaReceiptPaymentLabel(0, receipt({ paymentMethod: "Visa", date: "2026-06-19", purpose: "Parking", amount: "27.00", currency: "USD" }))).toBe(
      "2026-06-19 | \u505c\u8f66\u8d39 | 27.00 USD"
    );
  });

  it("cleans common Korea detail labels, location, payment method, and wide receipt choice", () => {
    const rental = receipt({ purpose: "National car rental", place: "ATLANTA INTL ARPT, GA", paymentMethod: "MASTERCARD (1723)", currency: "USD" });
    const esim = receipt({ purpose: "eSIM data plan", details: "internet access", currency: "RMB" });

    expect(koreaDetailContent(rental)).toBe("\u79df\u8f66\u8d39\u7528");
    expect(koreaDetailLocation(esim)).toBe("\u7f8e\u56fd");
    expect(koreaPaymentMethodText(rental)).toBe("MASTERCARD");
    expect(koreaReceiptShouldBeWide(rental)).toBe(true);
  });

  it("ships the Korea web template without a frozen detail pane", async () => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(readFileSync("public/templates/korea_reimbursement_template.xlsx"));
    const details = workbook.worksheets[1];

    expect(details?.views.some((view) => view.state === "frozen")).toBe(false);
  });

  it("generates a Korea workbook download without ExcelJS write errors", async () => {
    const template = readFileSync("public/templates/korea_reimbursement_template.xlsx");
    class TestBlob {
      parts: Array<ArrayBuffer | ArrayBufferView | string>;
      type: string;

      constructor(parts: Array<ArrayBuffer | ArrayBufferView | string>, options: { type?: string } = {}) {
        this.parts = parts;
        this.type = options.type ?? "";
      }

      async arrayBuffer(): Promise<ArrayBuffer> {
        const part = this.parts[0];
        if (part instanceof ArrayBuffer) return part;
        if (ArrayBuffer.isView(part)) return part.buffer.slice(part.byteOffset, part.byteOffset + part.byteLength) as ArrayBuffer;
        return new TextEncoder().encode(String(part)).buffer as ArrayBuffer;
      }
    }
    vi.stubGlobal("Blob", TestBlob);
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => new Response(template))
    );
    let downloadedBlob: TestBlob | undefined;
    vi.stubGlobal("URL", {
      createObjectURL: vi.fn((blob: TestBlob) => {
        downloadedBlob = blob;
        return "blob:workbook";
      }),
      revokeObjectURL: vi.fn()
    });
    const clickSpy = vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(() => undefined);

    await expect(
      exportKoreaWorkbook(
        [receipt({ id: "r1", category: "transportation", date: "2026-06-19", amount: "27", currency: "USD", images: [] })],
        { usdToRmb: 6.8175, usdToKrw: 1548.86, krwToRmb: 0.004402959 },
        []
      )
    ).resolves.toBeUndefined();

    expect(clickSpy).toHaveBeenCalled();
    expect(downloadedBlob).toBeDefined();
    const workbook = new ExcelJS.Workbook();
    const rawWorkbook = await downloadedBlob!.arrayBuffer();
    await workbook.xlsx.load(Buffer.from(new Uint8Array(rawWorkbook)));
    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(["报销明细", "境外同事报销使用", "发票"]);
    expect(workbook.getWorksheet("发票")?.getCell("A1").value).toBe("2026-06-19 | \u505c\u8f66\u8d39 | 27 USD");
  });
});
