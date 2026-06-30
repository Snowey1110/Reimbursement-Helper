import { afterEach, describe, expect, it, vi } from "vitest";
import { extractKrwToRmbRateWithOpenAI, extractReceiptWithOpenAI } from "./ai";
import { receipt } from "./test/factories";

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
            receipt_label: "Parking Receipt"
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
    expect(body.input[0].content[1]).toMatchObject({
      type: "input_image",
      image_url: receipt().images[0].dataUrl
    });
  });

  it("extracts KRW to RMB rate from multiple exchange-rate screenshots", async () => {
    const fetchMock = vi.fn(async () =>
      new Response(
        JSON.stringify({
          output_text: JSON.stringify({
            krw_to_rmb_rate: "0.0044029590",
            usd_to_krw_rate: "1548.86",
            confidence_notes: "explicit app rate visible"
          })
        }),
        { status: 200 }
      )
    );
    vi.stubGlobal("fetch", fetchMock);

    const image = receipt().images[0];
    const rate = await extractKrwToRmbRateWithOpenAI("test-user-key", "gpt-test", [{ ...image, id: "rate-1" }, { ...image, id: "rate-2" }], 6.8175);

    expect(rate).toBeCloseTo(0.004402959);
    const body = JSON.parse(String((fetchMock.mock.calls[0][1] as RequestInit).body));
    expect(body.input[0].content.filter((part: any) => part.type === "input_image")).toHaveLength(2);
  });
});
