import { afterEach, describe, expect, it, vi } from "vitest";
import { applyExtraction, extractKoreaExchangeRatesWithOpenAI, extractReceiptWithOpenAI } from "./ai";
import { image, receipt } from "./test/factories";

describe("OpenAI extraction", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("sends the user-entered API key and image to the Responses API", async () => {
    const fetchMock = vi.fn(async () =>
      new Response(
        JSON.stringify({
          output_text: JSON.stringify({
            date: "2026-06-19",
            place: "McCormick Place",
            vendor: "McCormick Place",
            amount: "27",
            currency: "USD",
            krw_amount: "",
            rmb_amount: "",
            purpose: "Parking",
            details: "Parking fee",
            project_number: "",
            category: "transportation",
            payment_method: "Visa",
            receipt_label: "Parking Receipt",
            document_type: "receipt"
          })
        }),
        { status: 200 }
      )
    );
    vi.stubGlobal("fetch", fetchMock);

    const result = await extractReceiptWithOpenAI("test-user-key", "gpt-test", "USA", receipt());

    expect(result.amount).toBe("27");
    expect(fetchMock).toHaveBeenCalledTimes(1);
    const [url, options] = fetchMock.mock.calls[0];
    expect(url).toBe("https://api.openai.com/v1/responses");
    expect((options as RequestInit).headers).toMatchObject({
      Authorization: "Bearer test-user-key"
    });
    const body = JSON.parse(String((options as RequestInit).body));
    expect(body.model).toBe("gpt-test");
    expect(body.input[0].content[2]).toMatchObject({
      type: "input_image",
      image_url: receipt().images[0].dataUrl
    });
  });

  it("sends every uploaded page/image for a receipt entry", async () => {
    const fetchMock = vi.fn(async () =>
      new Response(
        JSON.stringify({
          output_text: JSON.stringify({
            date: "",
            place: "",
            vendor: "",
            amount: "",
            currency: "USD",
            krw_amount: "",
            rmb_amount: "",
            purpose: "",
            details: "",
            project_number: "",
            category: "transportation",
            payment_method: "",
            receipt_label: "",
            document_type: "unknown"
          })
        }),
        { status: 200 }
      )
    );
    vi.stubGlobal("fetch", fetchMock);

    await extractReceiptWithOpenAI(
      "test-user-key",
      "gpt-test",
      "USA",
      receipt({ images: [image({ id: "page-1", filename: "file.pdf page 1" }), image({ id: "page-2", filename: "file.pdf page 2" })] })
    );

    const body = JSON.parse(String((fetchMock.mock.calls[0][1] as RequestInit).body));
    expect(body.input[0].content.filter((part: any) => part.type === "input_image")).toHaveLength(2);
  });

  it("extracts KRW to RMB rate from multiple exchange-rate screenshots", async () => {
    const fetchMock = vi.fn(async () =>
      new Response(
        JSON.stringify({
          output_text: JSON.stringify({
            krw_to_rmb_rate: "0.0044029590",
            usd_to_krw_rate: "1548.86",
            krw_to_usd_rate: "",
            confidence_notes: "explicit app rate visible"
          })
        }),
        { status: 200 }
      )
    );
    vi.stubGlobal("fetch", fetchMock);

    const image = receipt().images[0];
    const rate = await extractKoreaExchangeRatesWithOpenAI("test-user-key", "gpt-test", [{ ...image, id: "rate-1" }, { ...image, id: "rate-2" }], 6.8175);

    expect(rate.krwToRmb).toBeCloseTo(0.004402959);
    expect(rate.usdToKrw).toBeCloseTo(1548.86);
    const body = JSON.parse(String((fetchMock.mock.calls[0][1] as RequestInit).body));
    expect(body.input[0].content.filter((part: any) => part.type === "input_image")).toHaveLength(2);
  });

  it("keeps merchant credit-card receipts as receipts, not payment proof", () => {
    const result = applyExtraction(
      receipt({ filename: "toll_250_receipt.png" }),
      {
        date: "2026-06-26",
        place: "Indiana Toll Road",
        vendor: "Indiana Toll Road",
        amount: "2.50",
        currency: "USD",
        purpose: "Breakfast",
        details: "Credit Card Receipt. Total $2.50.",
        category: "meals",
        payment_method: "VISA CREDIT",
        receipt_label: "Toll receipt",
        document_type: "payment_proof"
      },
      "USA"
    );

    expect(result.documentType).toBe("receipt");
    expect(result.category).toBe("transportation");
    expect(result.purpose).toBe("Toll");
  });

  it("corrects Chase transaction screenshots to payment proof even if AI says receipt", () => {
    const result = applyExtraction(
      receipt({ filename: "chase_transaction.png" }),
      {
        date: "2026-06-17",
        place: "CH WABASH RANDOLPH",
        vendor: "CH WABASH RANDOLPH",
        amount: "30.00",
        currency: "USD",
        purpose: "Parking",
        details: "CHASE transaction details. Printed from Chase Personal Online. Merchant type Parking lots and garages. Posted date Jun 18, 2026. Rewards earned.",
        category: "transportation",
        payment_method: "Sapphire Reserve",
        receipt_label: "Chase transaction details",
        krw_amount: "",
        rmb_amount: "",
        project_number: "",
        document_type: "receipt"
      },
      "USA"
    );

    expect(result.documentType).toBe("payment_proof");
  });
});
