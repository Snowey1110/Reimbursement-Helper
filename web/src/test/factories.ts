import type { Category, Currency, ImageAttachment, PaymentProof, ReceiptItem } from "../types";

export function image(overrides: Partial<ImageAttachment> = {}): ImageAttachment {
  const id = overrides.id ?? "image-1";
  return {
    id,
    filename: overrides.filename ?? `${id}.png`,
    sourceName: overrides.sourceName ?? `${id}.png`,
    dataUrl: overrides.dataUrl ?? `data:image/png;base64,${id}`,
    width: overrides.width ?? 100,
    height: overrides.height ?? 200,
    rotationDegrees: overrides.rotationDegrees ?? 0,
    cropPoints: overrides.cropPoints,
    sourcePage: overrides.sourcePage,
    pageCount: overrides.pageCount,
    isPdfPage: overrides.isPdfPage
  };
}

export function receipt(overrides: Partial<ReceiptItem> = {}): ReceiptItem {
  const id = overrides.id ?? "receipt-1";
  return {
    id,
    filename: overrides.filename ?? `${id}.png`,
    status: overrides.status ?? "Empty",
    date: overrides.date ?? "2026-06-19",
    place: overrides.place ?? "Vendor",
    amount: overrides.amount ?? "10",
    currency: (overrides.currency ?? "USD") as Currency,
    krwAmount: overrides.krwAmount ?? "",
    rmbAmount: overrides.rmbAmount ?? "",
    purpose: overrides.purpose ?? "Parking",
    details: overrides.details ?? "Parking fee",
    projectNumber: overrides.projectNumber ?? "",
    category: (overrides.category ?? "transportation") as Category,
    paymentMethod: overrides.paymentMethod ?? "Visa",
    receiptLabel: overrides.receiptLabel ?? "Receipt",
    images: overrides.images ?? [image({ id: `${id}-image` })]
  };
}

export function proof(overrides: Partial<PaymentProof> = {}): PaymentProof {
  const id = overrides.id ?? "proof-1";
  return {
    id,
    filename: overrides.filename ?? `${id}.png`,
    status: overrides.status ?? "Needs match",
    date: overrides.date ?? "2026-06-19",
    amount: overrides.amount ?? "10",
    place: overrides.place ?? "Card statement",
    matchedReceiptId: overrides.matchedReceiptId ?? "",
    image: overrides.image ?? image({ id: `${id}-image` })
  };
}
