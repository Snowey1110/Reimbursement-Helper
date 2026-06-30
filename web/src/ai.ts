import type { FormVersion, ReceiptExtraction, ReceiptItem } from "./types";
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
