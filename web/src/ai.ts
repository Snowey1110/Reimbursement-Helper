import type { FormVersion, ImageAttachment, ReceiptExtraction, ReceiptItem } from "./types";
import { normalizeCategory, normalizeCurrency, updateAmounts } from "./utils";
import { DEFAULT_KRW_TO_RMB, DEFAULT_USD_TO_RMB } from "./constants";

const RECEIPT_SCHEMA = {
  type: "object",
  additionalProperties: false,
  properties: {
    date: { type: "string" },
    place: { type: "string" },
    vendor: { type: "string" },
    amount: { type: "string" },
    currency: { type: "string" },
    krw_amount: { type: "string" },
    rmb_amount: { type: "string" },
    purpose: { type: "string" },
    details: { type: "string" },
    project_number: { type: "string" },
    category: { type: "string" },
    payment_method: { type: "string" },
    receipt_label: { type: "string" }
  },
  required: [
    "date",
    "place",
    "vendor",
    "amount",
    "currency",
    "krw_amount",
    "rmb_amount",
    "purpose",
    "details",
    "project_number",
    "category",
    "payment_method",
    "receipt_label"
  ]
};

const EXCHANGE_RATE_SCHEMA = {
  type: "object",
  additionalProperties: false,
  properties: {
    krw_to_rmb_rate: { type: "string" },
    usd_to_krw_rate: { type: "string" },
    confidence_notes: { type: "string" }
  },
  required: ["krw_to_rmb_rate", "usd_to_krw_rate", "confidence_notes"]
};

function parseRate(value: unknown): number | undefined {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const text = String(value ?? "").replace(/,/g, "");
  const match = /-?\d+(?:\.\d+)?/.exec(text);
  if (!match) return undefined;
  const parsed = Number(match[0]);
  return Number.isFinite(parsed) ? parsed : undefined;
}

export async function extractReceiptWithOpenAI(apiKey: string, model: string, formVersion: FormVersion, item: ReceiptItem): Promise<ReceiptExtraction> {
  const image = item.images[0];
  const prompt = [
    `Extract reimbursement details for the ${formVersion} form.`,
    "Return dates as YYYY-MM-DD when possible.",
    "Use the actual receipt currency. If the receipt is USD, currency must be USD even on the Korea form.",
    "Prefer charged/paid total amount over subtotal, tax-only, or authorization metadata.",
    "Choose category from transportation, lodging, meals, advertising, office, entertainment, materials, consumables, physical_exam, nucleic_test, courier, welfare, other.",
    "Keep applicant/personal identity fields blank."
  ].join("\n");
  const response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model,
      input: [
        {
          role: "user",
          content: [
            { type: "input_text", text: prompt },
            { type: "input_image", image_url: image.dataUrl }
          ]
        }
      ],
      text: {
        format: {
          type: "json_schema",
          name: "receipt_extraction",
          schema: RECEIPT_SCHEMA,
          strict: true
        }
      }
    })
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`OpenAI request failed (${response.status}): ${text.slice(0, 300)}`);
  }
  const data = await response.json();
  const text = data.output_text ?? data.output?.flatMap((part: any) => part.content ?? []).find((part: any) => part.type === "output_text")?.text;
  if (!text) {
    throw new Error("OpenAI response did not include JSON text.");
  }
  return JSON.parse(text) as ReceiptExtraction;
}

export async function extractKrwToRmbRateWithOpenAI(
  apiKey: string,
  model: string,
  images: ImageAttachment[],
  usdToRmb: number
): Promise<number> {
  if (!images.length) throw new Error("Select 汇率 image files first.");
  const prompt = [
    "Read these exchange-rate screenshots for the Korea reimbursement form.",
    "Return only JSON.",
    "Prefer an explicit KRW -> RMB or 汇率 value, for example 0.0044029590.",
    "If the screenshot only shows USD -> KRW, calculate KRW -> RMB as USD_TO_RMB / USD_TO_KRW.",
    `The current USD_TO_RMB value from the app is ${usdToRmb}.`,
    "Return numbers as plain decimals without currency symbols or commas."
  ].join("\n");
  const response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model,
      input: [
        {
          role: "user",
          content: [
            { type: "input_text", text: prompt },
            ...images.map((image) => ({ type: "input_image", image_url: image.dataUrl }))
          ]
        }
      ],
      text: {
        format: {
          type: "json_schema",
          name: "exchange_rate_extraction",
          schema: EXCHANGE_RATE_SCHEMA,
          strict: true
        }
      }
    })
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`OpenAI request failed (${response.status}): ${text.slice(0, 300)}`);
  }
  const data = await response.json();
  const text = data.output_text ?? data.output?.flatMap((part: any) => part.content ?? []).find((part: any) => part.type === "output_text")?.text;
  if (!text) {
    throw new Error("OpenAI response did not include JSON text.");
  }
  const parsed = JSON.parse(text);
  const explicitRate = parseRate(parsed.krw_to_rmb_rate);
  if (explicitRate !== undefined && explicitRate > 0 && explicitRate < 1) return explicitRate;
  const usdToKrw = parseRate(parsed.usd_to_krw_rate);
  if (usdToKrw !== undefined && usdToKrw > 1 && usdToRmb) return usdToRmb / usdToKrw;
  throw new Error("AI could not find a usable KRW -> RMB exchange rate.");
}

export function applyExtraction(item: ReceiptItem, extraction: ReceiptExtraction, formVersion: FormVersion): ReceiptItem {
  const currency = normalizeCurrency(extraction.currency, formVersion === "USA" ? "USD" : item.currency);
  const next: ReceiptItem = {
    ...item,
    date: String(extraction.date ?? item.date ?? ""),
    place: String(extraction.place ?? extraction.vendor ?? item.place ?? ""),
    amount: String(extraction.amount ?? item.amount ?? ""),
    currency: formVersion === "USA" ? "USD" : currency,
    krwAmount: String(extraction.krw_amount ?? item.krwAmount ?? ""),
    rmbAmount: String(extraction.rmb_amount ?? item.rmbAmount ?? ""),
    purpose: String(extraction.purpose ?? item.purpose ?? ""),
    details: String(extraction.details ?? item.details ?? ""),
    projectNumber: String(extraction.project_number ?? item.projectNumber ?? ""),
    category: normalizeCategory(extraction.category ?? item.category),
    paymentMethod: String(extraction.payment_method ?? item.paymentMethod ?? ""),
    receiptLabel: String(extraction.receipt_label ?? item.receiptLabel ?? item.filename),
    status: "AI filled"
  };
  return updateAmounts(next, { usdToRmb: DEFAULT_USD_TO_RMB, krwToRmb: DEFAULT_KRW_TO_RMB }, "amount");
}
