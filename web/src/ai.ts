import type {
  DocumentEvidence,
  FormVersion,
  ImageAttachment,
  PageExtraction,
  ReceiptExtraction,
  ReceiptItem,
  SupportRole
} from "./types";
import { orientedImageDataUrl } from "./imageUtils";
import {
  looksLikeUsaFuelReceipt,
  looksLikeUsaParkingReceipt,
  looksLikeUsaTollReceipt,
  normalizeCategory,
  normalizeCurrency,
  normalizeDocumentType,
  normalizeUsaTransportationFields,
  updateAmounts
} from "./utils";
import { DEFAULT_KRW_TO_RMB, DEFAULT_USD_TO_KRW, DEFAULT_USD_TO_RMB } from "./constants";

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
    receipt_label: { type: "string" },
    document_type: { type: "string" },
    receipt_image_indexes: { type: "string" },
    payment_proof_image_indexes: { type: "string" },
    image_roles: { type: "string" },
    document_evidence: { type: "string" },
    transaction_date: { type: "string" },
    posted_date: { type: "string" },
    bank_name: { type: "string" },
    card_last4: { type: "string" },
    transaction_reference: { type: "string" },
    page_results: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        properties: {
          page_index: { type: "number" },
          role: { type: "string" },
          rotation_degrees: { type: "number" },
          evidence: { type: "string" },
          transaction_date: { type: "string" },
          posted_date: { type: "string" },
          amount: { type: "string" },
          vendor: { type: "string" },
          bank_name: { type: "string" },
          card_last4: { type: "string" },
          transaction_reference: { type: "string" }
        },
        required: [
          "page_index",
          "role",
          "rotation_degrees",
          "evidence",
          "transaction_date",
          "posted_date",
          "amount",
          "vendor",
          "bank_name",
          "card_last4",
          "transaction_reference"
        ]
      }
    }
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
    "receipt_label",
    "document_type",
    "receipt_image_indexes",
    "payment_proof_image_indexes",
    "image_roles",
    "document_evidence",
    "transaction_date",
    "posted_date",
    "bank_name",
    "card_last4",
    "transaction_reference",
    "page_results"
  ]
};

const IMAGE_ROLE_SCHEMA = {
  type: "object",
  additionalProperties: false,
  properties: {
    image_roles: { type: "string" }
  },
  required: ["image_roles"]
};

const EXCHANGE_RATE_SCHEMA = {
  type: "object",
  additionalProperties: false,
  properties: {
    krw_to_rmb_rate: { type: "string" },
    usd_to_krw_rate: { type: "string" },
    krw_to_usd_rate: { type: "string" },
    confidence_notes: { type: "string" }
  },
  required: ["krw_to_rmb_rate", "usd_to_krw_rate", "krw_to_usd_rate", "confidence_notes"]
};

function joinedText(values: unknown[]): string {
  return values.join(" ").toLowerCase();
}

function usaPaymentProofScore(values: unknown[]): number {
  const text = joinedText(values);
  let score = 0;
  if (/\b(chase|jpmorgan|sapphire reserve|freedom)\b/.test(text)) score += 2;
  if (/\b(transaction details|transaction date|posted date|printed from chase personal online)\b/.test(text)) score += 2;
  if (/\b(merchant type|card number|account activity|card statement|bank statement|proof of payment|payment proof)\b/.test(text)) score += 2;
  if (/\b(rewards earned|ultimate rewards|sale)\b/.test(text)) score += 1;
  return score;
}

function merchantReceiptScore(values: unknown[]): number {
  const text = joinedText(values);
  let score = 0;
  if (/\b(credit card receipt|merchant receipt|sales receipt|receipt|invoice)\b/.test(text)) score += 2;
  if (/\b(ticket|entry|exit|qr code|barcode)\b/.test(text)) score += 1;
  if (/\b(total|subtotal|tax|amount due|paid|payment method|card type|authorization|terminal|receipt number|cashier|thank you)\b/.test(text)) score += 2;
  if (/\b(indiana toll road|itr concession|metropolis|iparkit|parking charges|visit details)\b/.test(text)) score += 2;
  return score;
}

function looksLikeUsaPaymentProof(values: unknown[]): boolean {
  return usaPaymentProofScore(values) >= 3 && merchantReceiptScore(values) < 3;
}

function looksLikeStrongMerchantReceipt(values: unknown[]): boolean {
  return merchantReceiptScore(values) >= 3;
}

function looksLikeStrongPaymentProof(values: unknown[]): boolean {
  return usaPaymentProofScore(values) >= 4 && merchantReceiptScore(values) < 3;
}

function looksLikeMerchantReceipt(values: unknown[]): boolean {
  const text = values.join(" ").toLowerCase();
  return /\b(receipt|ticket|invoice|total|subtotal|tax|authorization|terminal|thank you|cashier|merchant copy)\b/.test(text);
}

function parseImageIndexes(value: unknown, max: number): Set<number> {
  const result = new Set<number>();
  const text = String(value ?? "");
  for (const match of text.matchAll(/\d+/g)) {
    const index = Number(match[0]) - 1;
    if (Number.isInteger(index) && index >= 0 && index < max) result.add(index);
  }
  return result;
}

type ImageSupportRole = NonNullable<ImageAttachment["supportRole"]>;

function normalizeImageRole(value: unknown): ImageSupportRole | undefined {
  const raw = String(value ?? "").trim().toLowerCase().replace(/[\s-]+/g, "_");
  if (!raw) return undefined;
  if (raw.includes("payment_proof") || raw.includes("bank_proof") || raw.includes("card_proof") || raw === "proof") return "payment_proof";
  if (raw.includes("companion") || raw === "ticket") return "companion_ticket";
  if (raw.includes("receipt") || raw.includes("invoice")) return "receipt";
  if (raw.includes("ignore")) return "ignore";
  if (raw.includes("unknown") || raw.includes("unclear")) return "unknown";
  return undefined;
}

function normalizeEvidence(value: unknown): DocumentEvidence {
  const raw = String(value ?? "").trim().toLowerCase().replace(/[\s-]+/g, "_");
  if (raw.includes("bank") || raw.includes("card_portal") || raw.includes("statement")) return "bank_portal";
  if (raw.includes("photo") || raw.includes("paper") || raw.includes("scan")) return "photographed_paper";
  if (raw.includes("merchant") || raw.includes("invoice") || raw.includes("document")) return "merchant_document";
  return "unknown";
}

function normalizeRotationDelta(value: unknown): number {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 0;
  const rounded = Math.round(numeric / 90) * 90;
  return ((rounded % 360) + 360) % 360;
}

function validPageResults(value: unknown, max: number): PageExtraction[] {
  if (!Array.isArray(value)) return [];
  const result: PageExtraction[] = [];
  const seen = new Set<number>();
  for (const raw of value) {
    if (!raw || typeof raw !== "object") continue;
    const candidate = raw as PageExtraction;
    const index = Math.round(Number(candidate.page_index));
    if (!Number.isFinite(index) || index < 1 || index > max || seen.has(index)) continue;
    seen.add(index);
    result.push({ ...candidate, page_index: index });
  }
  return result.sort((left, right) => left.page_index - right.page_index);
}

async function aiImageDataUrl(attachment: ImageAttachment, maxWidth: number, maxHeight: number): Promise<string> {
  if (((attachment.rotationDegrees % 360) + 360) % 360 === 0) return attachment.dataUrl;
  try {
    return (await orientedImageDataUrl(attachment, maxWidth, maxHeight)).dataUrl;
  } catch {
    return attachment.dataUrl;
  }
}

function parseImageRoleMap(value: unknown, max: number): Map<number, ImageSupportRole> {
  const result = new Map<number, ImageSupportRole>();
  const text = String(value ?? "");
  const indexedRolePattern = /(?:image|page|img)?\s*(\d+)\s*[:=\-)]+\s*([a-zA-Z_ -]+)/g;
  for (const match of text.matchAll(indexedRolePattern)) {
    const index = Number(match[1]) - 1;
    const role = normalizeImageRole(match[2]);
    if (Number.isInteger(index) && index >= 0 && index < max && role) {
      result.set(index, role);
    }
  }
  const roleListPattern = /\b(payment[_\s-]*proof|bank[_\s-]*proof|card[_\s-]*proof|proof|companion[_\s-]*ticket|receipt|invoice|ticket|unknown|ignore)\b\s*[:=]\s*([0-9,\s]+)/gi;
  for (const match of text.matchAll(roleListPattern)) {
    const role = normalizeImageRole(match[1]);
    if (!role) continue;
    for (const numberMatch of match[2].matchAll(/\d+/g)) {
      const index = Number(numberMatch[0]) - 1;
      if (Number.isInteger(index) && index >= 0 && index < max) {
        result.set(index, role);
      }
    }
  }
  return result;
}

function parseRate(value: unknown): number | undefined {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const text = String(value ?? "").replace(/,/g, "");
  const match = /-?\d+(?:\.\d+)?/.exec(text);
  if (!match) return undefined;
  const parsed = Number(match[0]);
  return Number.isFinite(parsed) ? parsed : undefined;
}

async function readJsonText(response: Response): Promise<string> {
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`OpenAI request failed (${response.status}): ${text.slice(0, 300)}`);
  }
  const data = await response.json();
  const text = data.output_text ?? data.output?.flatMap((part: any) => part.content ?? []).find((part: any) => part.type === "output_text")?.text;
  if (!text) {
    throw new Error("OpenAI response did not include JSON text.");
  }
  return text;
}

async function extractImageRolesWithOpenAI(apiKey: string, model: string, item: ReceiptItem): Promise<string> {
  const aiImages = await Promise.all(item.images.map((entry) => aiImageDataUrl(entry, 1800, 3000)));
  const prompt = [
    "Classify each attached USA reimbursement image/page.",
    "Return every image number exactly once in image_roles, like: 1:payment_proof, 2:receipt, 3:ignore.",
    "Use only these roles: receipt, payment_proof, ignore.",
    "payment_proof means a bank, credit-card, or card transaction screenshot/statement such as Chase/JPMorgan transaction details.",
    "receipt means a merchant receipt, ticket, invoice, QR entry ticket, parking receipt, toll receipt, fuel receipt, hotel invoice, airline receipt, or any seller-provided purchase document.",
    "A merchant receipt that says paid by credit card, Visa, Mastercard, card, or authorization is still receipt, not payment_proof.",
    "If a PDF has a Chase/card transaction page and a merchant receipt page, mark the Chase/card page payment_proof and the merchant receipt page receipt.",
    `Entry filename: ${item.filename}`
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
            ...item.images.flatMap((entry, index) => [
              { type: "input_text", text: `Image ${index + 1}: ${entry.filename || entry.sourceName || item.filename}` },
              { type: "input_image", image_url: aiImages[index] }
            ])
          ]
        }
      ],
      text: {
        format: {
          type: "json_schema",
          name: "image_role_extraction",
          schema: IMAGE_ROLE_SCHEMA,
          strict: true
        }
      }
    })
  });
  const text = await readJsonText(response);
  return String((JSON.parse(text) as { image_roles?: string }).image_roles ?? "");
}

export async function extractReceiptWithOpenAI(apiKey: string, model: string, formVersion: FormVersion, item: ReceiptItem): Promise<ReceiptExtraction> {
  const aiImages = await Promise.all(item.images.map((entry) => aiImageDataUrl(entry, 2200, 3600)));
  const prompt = [
    `Extract reimbursement details for the ${formVersion} form.`,
    "Return dates as YYYY-MM-DD when possible.",
    "Use the actual receipt currency. If the receipt is USD, currency must be USD even on the Korea form.",
    "Prefer charged/paid total amount over subtotal, tax-only, or authorization metadata.",
    "Choose category from transportation, lodging, meals, advertising, office, entertainment, materials, consumables, physical_exam, nucleic_test, courier, welfare, other.",
    "For USA uploads, classify each page before classifying the whole item.",
    "For every original page, return one page_results object with its 1-based page_index, role, clockwise rotation needed from the supplied image, evidence, dates, amount, vendor, bank, card last four digits, and transaction reference.",
    "Use page role receipt, companion_ticket, payment_proof, or ignore.",
    "Use evidence photographed_paper, merchant_document, bank_portal, or unknown.",
    "Payment proof means a digital bank/card portal or statement such as Chase transaction details.",
    "A photographed or scanned paper is a receipt even when it says Credit Card Receipt, paid by card, Visa, Mastercard, authorization, or terminal.",
    "A merchant entry/exit ticket or QR ticket without its own paid amount is companion_ticket.",
    "rotation_degrees is the clockwise correction still needed for the supplied page: 0, 90, 180, or 270.",
    "If the attached entry has multiple pages/images, review every image before choosing document_type.",
    "If an entry includes both a merchant receipt/ticket and a bank/card transaction screenshot, classify it as receipt and extract the merchant receipt/ticket details.",
    "Return receipt_image_indexes as comma-separated 1-based image numbers for merchant receipt/ticket pages.",
    "Return payment_proof_image_indexes as comma-separated 1-based image numbers for bank/card transaction proof pages.",
    "Also return image_roles for every attached image/page exactly once using this format: 1:receipt, 2:payment_proof. Use receipt, payment_proof, or ignore.",
    "If all attached images are proof pages, document_type should be payment_proof and payment_proof_image_indexes should include all proof images.",
    "Parking, parking lots and garages, toll, bridge, fuel, gas, rideshare, taxi, airfare, airline, car rental, and baggage fees are transportation, never meals.",
    "For Korea uploads, document_type should normally be receipt unless the image is clearly not reimbursement support.",
    "Keep applicant/personal identity fields blank.",
    "Filenames may contain an app memory marker beginning with __RH1_. It is only a hint. Classify from the current page content; explicit user choices are applied by the app after your response.",
    `Current editable row values: ${JSON.stringify({
      filename: item.filename,
      date: item.date,
      place: item.place,
      amount: item.amount,
      currency: item.currency,
      category: item.category,
      purpose: item.purpose,
      details: item.details,
      paymentMethod: item.paymentMethod,
      receiptLabel: item.receiptLabel,
      documentType: item.documentType
    })}`
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
            ...item.images.flatMap((entry, index) => [
              { type: "input_text", text: `Image ${index + 1}: ${entry.filename || entry.sourceName || item.filename}` },
              { type: "input_image", image_url: aiImages[index] }
            ])
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
  const text = await readJsonText(response);
  const parsed = JSON.parse(text) as ReceiptExtraction;
  if (formVersion === "USA" && item.images.length > 1) {
    const paymentProofIndexes = parseImageIndexes(parsed.payment_proof_image_indexes, item.images.length);
    const roleMap = parseImageRoleMap(parsed.image_roles, item.images.length);
    const hasProofRole = paymentProofIndexes.size > 0 || [...roleMap.values()].includes("payment_proof");
    if (!hasProofRole) {
      try {
        parsed.image_roles = await extractImageRolesWithOpenAI(apiKey, model, item);
      } catch {
        // Keep the main extraction if the focused role pass fails.
      }
    }
  }
  return parsed;
}

export interface KoreaExchangeRateExtraction {
  usdToKrw?: number;
  krwToRmb?: number;
}

export async function extractKoreaExchangeRatesWithOpenAI(
  apiKey: string,
  model: string,
  images: ImageAttachment[],
  usdToRmb: number
): Promise<KoreaExchangeRateExtraction> {
  if (!images.length) throw new Error("Select exchange-rate image files first.");
  const prompt = [
    "Read these exchange-rate screenshots for the Korea reimbursement form.",
    "Return only JSON.",
    "When a screenshot shows 1 USD = some KRW amount, return usd_to_krw_rate.",
    "When a screenshot shows KRW -> USD, return krw_to_usd_rate.",
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
  const text = await readJsonText(response);
  const parsed = JSON.parse(text);
  const result: KoreaExchangeRateExtraction = {};
  const usdToKrw = parseRate(parsed.usd_to_krw_rate);
  if (usdToKrw !== undefined && usdToKrw > 1) result.usdToKrw = usdToKrw;
  const krwToUsd = parseRate(parsed.krw_to_usd_rate);
  if (result.usdToKrw === undefined && krwToUsd !== undefined && krwToUsd > 0 && krwToUsd < 1) result.usdToKrw = 1 / krwToUsd;
  const explicitRate = parseRate(parsed.krw_to_rmb_rate);
  if (explicitRate !== undefined && explicitRate > 0 && explicitRate < 1) result.krwToRmb = explicitRate;
  if (result.krwToRmb === undefined && result.usdToKrw !== undefined && usdToRmb) result.krwToRmb = usdToRmb / result.usdToKrw;
  if (result.usdToKrw === undefined && result.krwToRmb === undefined) {
    throw new Error("AI could not find a usable exchange rate.");
  }
  return result;
}

export function applyExtraction(item: ReceiptItem, extraction: ReceiptExtraction, formVersion: FormVersion): ReceiptItem {
  const manualFields = new Set(item.manualFields ?? []);
  const keepManual = (field: string, extracted: unknown, current: string): string =>
    manualFields.has(field) ? current : String(extracted ?? current ?? "");
  const currency = normalizeCurrency(extraction.currency, formVersion === "USA" ? "USD" : item.currency);
  let documentType = formVersion === "USA" ? normalizeDocumentType(extraction.document_type, item.documentType) : item.documentType;
  const lockedDocumentType = formVersion === "USA" ? normalizeDocumentType(item.lockedDocumentType, "unknown") : "unknown";
  const classificationValues = [
    item.filename,
    ...item.images.flatMap((image) => [image.filename, image.sourceName]),
    extraction.place,
    extraction.vendor,
    extraction.details,
    extraction.payment_method,
    extraction.receipt_label,
    extraction.category
  ];
  const bankProof = looksLikeUsaPaymentProof(classificationValues);
  const strongProof = looksLikeStrongPaymentProof(classificationValues);
  const strongReceipt = looksLikeStrongMerchantReceipt(classificationValues);
  if (
    formVersion === "USA" &&
    documentType === "unknown" &&
    bankProof
  ) {
    documentType = "payment_proof";
  }
  if (formVersion === "USA" && documentType === "receipt" && strongProof) {
    documentType = "payment_proof";
  }
  if (formVersion === "USA" && documentType === "payment_proof" && strongReceipt) {
    documentType = "receipt";
  }
  const pageResults = validPageResults(extraction.page_results, item.images.length);
  const pageResultByIndex = new Map(pageResults.map((result) => [result.page_index - 1, result]));
  const roleMap = parseImageRoleMap(extraction.image_roles, item.images.length);
  const receiptImageIndexes = parseImageIndexes(extraction.receipt_image_indexes, item.images.length);
  const paymentProofImageIndexes = parseImageIndexes(extraction.payment_proof_image_indexes, item.images.length);
  for (const result of pageResults) {
    const index = result.page_index - 1;
    const role = normalizeImageRole(result.role);
    if (role === "payment_proof") {
      paymentProofImageIndexes.add(index);
      receiptImageIndexes.delete(index);
    } else if (role === "receipt" || role === "companion_ticket") {
      receiptImageIndexes.add(index);
      paymentProofImageIndexes.delete(index);
    } else if (role === "ignore") {
      receiptImageIndexes.delete(index);
      paymentProofImageIndexes.delete(index);
    }
  }
  for (const [index, role] of roleMap) {
    if (role === "payment_proof") {
      paymentProofImageIndexes.add(index);
      receiptImageIndexes.delete(index);
    } else if (role === "receipt") {
      receiptImageIndexes.add(index);
      paymentProofImageIndexes.delete(index);
    } else {
      receiptImageIndexes.delete(index);
      paymentProofImageIndexes.delete(index);
    }
  }
  if (lockedDocumentType === "payment_proof") {
    receiptImageIndexes.clear();
    item.images.forEach((_, index) => paymentProofImageIndexes.add(index));
    documentType = "payment_proof";
  } else if (lockedDocumentType === "receipt") {
    paymentProofImageIndexes.clear();
    item.images.forEach((_, index) => receiptImageIndexes.add(index));
    documentType = "receipt";
  }
  if (!receiptImageIndexes.size && !paymentProofImageIndexes.size && !roleMap.size && formVersion === "USA") {
    if (documentType === "payment_proof") {
      item.images.forEach((_, index) => paymentProofImageIndexes.add(index));
    } else if (documentType === "receipt") {
      item.images.forEach((_, index) => receiptImageIndexes.add(index));
    }
  }
  if (receiptImageIndexes.size) {
    documentType = "receipt";
  } else if (paymentProofImageIndexes.size === item.images.length && item.images.length > 0) {
    documentType = "payment_proof";
  }
  if (pageResults.length && lockedDocumentType === "unknown") {
    const pageRoles = pageResults.map((result) => normalizeImageRole(result.role));
    if (pageRoles.some((role) => role === "receipt" || role === "companion_ticket")) {
      documentType = "receipt";
    } else if (pageRoles.some((role) => role === "payment_proof")) {
      documentType = "payment_proof";
    }
  }
  const firstPageEvidence = pageResults
    .map((result) => normalizeEvidence(result.evidence))
    .find((evidence) => evidence !== "unknown");
  const firstPageValue = (key: keyof PageExtraction): string =>
    String(pageResults.map((result) => result[key]).find((value) => String(value ?? "").trim()) ?? "");
  let next: ReceiptItem = {
    ...item,
    images:
      formVersion === "USA"
        ? item.images.map((image, index) => {
            const pageResult = pageResultByIndex.get(index);
            const pageRole = pageResult ? normalizeImageRole(pageResult.role) : undefined;
            const manualRole = image.roleSource === "manual" ? image.supportRole : undefined;
            const lockedRole: SupportRole | undefined =
              lockedDocumentType === "payment_proof"
                ? "payment_proof"
                : lockedDocumentType === "receipt"
                  ? pageRole === "companion_ticket"
                    ? "companion_ticket"
                    : "receipt"
                  : undefined;
            const supportRole =
              manualRole ??
              lockedRole ??
              pageRole ??
              (paymentProofImageIndexes.has(index)
                ? "payment_proof"
                : receiptImageIndexes.has(index)
                  ? "receipt"
                  : roleMap.get(index) ?? image.supportRole ?? "unknown");
            const rotationDelta =
              pageResult && !image.rotationLocked
                ? normalizeRotationDelta(pageResult.rotation_degrees)
                : 0;
            return {
              ...image,
              supportRole,
              roleSource: manualRole ? "manual" : pageResult || roleMap.has(index) ? "ai" : image.roleSource ?? "",
              documentEvidence: pageResult ? normalizeEvidence(pageResult.evidence) : image.documentEvidence ?? "unknown",
              transactionDate: String(pageResult?.transaction_date ?? image.transactionDate ?? ""),
              postedDate: String(pageResult?.posted_date ?? image.postedDate ?? ""),
              pageAmount: String(pageResult?.amount ?? image.pageAmount ?? ""),
              pageVendor: String(pageResult?.vendor ?? image.pageVendor ?? ""),
              bankName: String(pageResult?.bank_name ?? image.bankName ?? ""),
              cardLast4: String(pageResult?.card_last4 ?? image.cardLast4 ?? ""),
              transactionReference: String(pageResult?.transaction_reference ?? image.transactionReference ?? ""),
              rotationDegrees: image.rotationLocked
                ? image.rotationDegrees
                : (image.rotationDegrees + rotationDelta) % 360
            };
          })
        : item.images,
    date: keepManual("date", extraction.date, item.date),
    place: keepManual("place", extraction.place || extraction.vendor, item.place),
    amount: keepManual("amount", extraction.amount, item.amount),
    currency: manualFields.has("currency") ? item.currency : formVersion === "USA" ? "USD" : currency,
    krwAmount: keepManual("krwAmount", extraction.krw_amount, item.krwAmount),
    rmbAmount: keepManual("rmbAmount", extraction.rmb_amount, item.rmbAmount),
    purpose: keepManual("purpose", extraction.purpose, item.purpose),
    details: keepManual("details", extraction.details, item.details),
    projectNumber: keepManual("projectNumber", extraction.project_number, item.projectNumber),
    category: manualFields.has("category") ? item.category : normalizeCategory(extraction.category ?? item.category),
    paymentMethod: keepManual("paymentMethod", extraction.payment_method, item.paymentMethod),
    receiptLabel: keepManual("receiptLabel", extraction.receipt_label, item.receiptLabel || item.filename),
    documentType,
    lockedDocumentType: item.lockedDocumentType,
    documentEvidence: normalizeEvidence(extraction.document_evidence) !== "unknown"
      ? normalizeEvidence(extraction.document_evidence)
      : firstPageEvidence ?? item.documentEvidence ?? "unknown",
    transactionDate: String(extraction.transaction_date || firstPageValue("transaction_date") || item.transactionDate || ""),
    postedDate: String(extraction.posted_date || firstPageValue("posted_date") || item.postedDate || ""),
    bankName: String(extraction.bank_name || firstPageValue("bank_name") || item.bankName || ""),
    cardLast4: String(extraction.card_last4 || firstPageValue("card_last4") || item.cardLast4 || ""),
    transactionReference: String(
      extraction.transaction_reference || firstPageValue("transaction_reference") || item.transactionReference || ""
    ),
    status: "AI filled"
  };
  const identityRoles: SupportRole[] =
    next.documentType === "payment_proof"
      ? ["payment_proof"]
      : ["receipt", "companion_ticket"];
  const identityImages = identityRoles.flatMap((role) =>
    next.images.filter((image) => image.supportRole === role)
  );
  const pageIdentityDate = identityImages.find((image) => image.transactionDate?.trim())?.transactionDate?.trim();
  const pageIdentityAmount = identityImages.find((image) => image.pageAmount?.trim())?.pageAmount?.trim();
  const pageIdentityVendor = identityImages.find((image) => image.pageVendor?.trim())?.pageVendor?.trim();
  if (pageIdentityDate && !manualFields.has("date")) next.date = pageIdentityDate;
  if (pageIdentityAmount && !manualFields.has("amount")) next.amount = pageIdentityAmount;
  if (pageIdentityVendor && !manualFields.has("place")) next.place = pageIdentityVendor;
  const paymentProofShouldBeReceipt =
    formVersion === "USA" &&
    next.documentType === "payment_proof" &&
    lockedDocumentType !== "payment_proof" &&
    !pageResults.length &&
    !strongProof &&
    (looksLikeMerchantReceipt(classificationValues) || looksLikeUsaParkingReceipt(next) || looksLikeUsaTollReceipt(next) || looksLikeUsaFuelReceipt(next));
  if (paymentProofShouldBeReceipt) {
    next = {
      ...next,
      documentType: "receipt",
      images: next.images.map((image) => (image.supportRole === "payment_proof" ? { ...image, supportRole: "receipt" } : image))
    };
  }
  if (formVersion === "USA") {
    next = normalizeUsaTransportationFields(next);
  }
  return updateAmounts(next, { usdToRmb: DEFAULT_USD_TO_RMB, usdToKrw: DEFAULT_USD_TO_KRW, krwToRmb: DEFAULT_KRW_TO_RMB }, "amount");
}
