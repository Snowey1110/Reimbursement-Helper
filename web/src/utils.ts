import { DEFAULT_KRW_TO_RMB, DEFAULT_USD_TO_RMB } from "./constants";
import type { Category, Currency, ExchangeRates, ImageAttachment, PaymentProof, ReceiptItem } from "./types";

export function uid(prefix = "id"): string {
  return `${prefix}_${crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2)}`;
}

export function safeNumber(value: unknown): number | undefined {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value !== "string") {
    return undefined;
  }
  const cleaned = value.replace(/[^0-9.-]/g, "");
  if (!cleaned) {
    return undefined;
  }
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : undefined;
}

export function formatAmount(value: number | undefined): string {
  if (value === undefined || !Number.isFinite(value)) {
    return "";
  }
  return value.toFixed(2).replace(/\.00$/, "");
}

export function normalizeCurrency(value: unknown, fallback: Currency): Currency {
  const raw = String(value ?? "").trim().toUpperCase();
  if (raw.includes("USD") || raw.includes("$")) return "USD";
  if (raw.includes("KRW") || raw.includes("WON") || raw.includes("\u20a9")) return "KRW";
  if (raw.includes("RMB") || raw.includes("CNY") || raw.includes("YUAN") || raw.includes("\u00a5")) return "RMB";
  return fallback;
}

export function normalizeCategory(value: unknown): Category {
  const raw = String(value ?? "").trim().toLowerCase();
  if (!raw) return "other";
  if (raw.includes("parking") || raw.includes("transport") || raw.includes("taxi") || raw.includes("toll") || raw.includes("fuel")) return "transportation";
  if (raw.includes("hotel") || raw.includes("lodging") || raw.includes("accommodation")) return "lodging";
  if (raw.includes("meal") || raw.includes("food") || raw.includes("restaurant")) return "meals";
  if (raw.includes("advert")) return "advertising";
  if (raw.includes("office")) return "office";
  if (raw.includes("material")) return "materials";
  if (raw.includes("consumable")) return "consumables";
  if (raw.includes("courier") || raw.includes("shipping")) return "courier";
  if (raw.includes("welfare")) return "welfare";
  if (raw.includes("exam")) return "physical_exam";
  if (raw.includes("nucleic")) return "nucleic_test";
  if (raw.includes("entertain")) return "entertainment";
  if (
    [
      "transportation",
      "lodging",
      "meals",
      "advertising",
      "office",
      "entertainment",
      "materials",
      "consumables",
      "physical_exam",
      "nucleic_test",
      "courier",
      "welfare",
      "other"
    ].includes(raw)
  ) {
    return raw as Category;
  }
  return "other";
}

export function blankReceipt(formVersion: "USA" | "Korea", image: ImageAttachment): ReceiptItem {
  return {
    id: uid("receipt"),
    filename: image.filename,
    status: "Empty",
    date: "",
    place: "",
    amount: "",
    currency: formVersion === "USA" ? "USD" : "KRW",
    krwAmount: "",
    rmbAmount: "",
    purpose: "",
    details: "",
    projectNumber: "",
    category: "transportation",
    paymentMethod: "",
    receiptLabel: "",
    images: [image]
  };
}

export function updateAmounts(item: ReceiptItem, rates: ExchangeRates, source: "amount" | "krw" | "rmb" | "currency" = "amount"): ReceiptItem {
  const usdRate = rates.usdToRmb || DEFAULT_USD_TO_RMB;
  const krwRate = rates.krwToRmb || DEFAULT_KRW_TO_RMB;
  const next = { ...item };
  const amount = safeNumber(next.amount);
  let krw = safeNumber(next.krwAmount);
  let rmb = safeNumber(next.rmbAmount);
  if (next.currency === "USD" && amount !== undefined && source !== "krw" && source !== "rmb") {
    rmb = amount * usdRate;
    krw = rmb / krwRate;
  } else if (next.currency === "KRW" && amount !== undefined && source !== "rmb") {
    krw = amount;
    rmb = amount * krwRate;
  } else if ((next.currency === "RMB" || next.currency === "CNY") && amount !== undefined && source !== "krw") {
    rmb = amount;
    krw = amount / krwRate;
  } else if (source === "krw" && krw !== undefined) {
    rmb = krw * krwRate;
    if (next.currency === "KRW") next.amount = formatAmount(krw);
  } else if (source === "rmb" && rmb !== undefined) {
    krw = rmb / krwRate;
    if (next.currency === "RMB" || next.currency === "CNY") next.amount = formatAmount(rmb);
  }
  next.krwAmount = formatAmount(krw);
  next.rmbAmount = formatAmount(rmb);
  return next;
}

export function parseDateKey(value: string): string {
  const date = new Date(value);
  if (!Number.isNaN(date.getTime())) {
    return date.toISOString().slice(0, 10);
  }
  return value.trim();
}

export function receiptMergeKey(item: ReceiptItem): string | undefined {
  const amount = safeNumber(item.amount);
  const date = parseDateKey(item.date);
  if (amount === undefined || !date) return undefined;
  return `${date}|${amount.toFixed(2)}`;
}

export function mergeSameUsaReceipts(items: ReceiptItem[]): ReceiptItem[] {
  const seen = new Map<string, ReceiptItem>();
  const merged: ReceiptItem[] = [];
  for (const item of items) {
    const key = receiptMergeKey(item);
    if (!key || !seen.has(key)) {
      const copy = { ...item, images: [...item.images] };
      if (key) seen.set(key, copy);
      merged.push(copy);
      continue;
    }
    const target = seen.get(key)!;
    const existing = new Set(target.images.map((image) => `${image.sourceName}|${image.filename}|${image.dataUrl.slice(0, 64)}`));
    for (const image of item.images) {
      const imageKey = `${image.sourceName}|${image.filename}|${image.dataUrl.slice(0, 64)}`;
      if (!existing.has(imageKey)) {
        target.images.push(image);
        existing.add(imageKey);
      }
    }
    for (const keyName of ["place", "purpose", "details", "projectNumber", "paymentMethod", "receiptLabel"] as const) {
      if (!target[keyName] && item[keyName]) {
        target[keyName] = item[keyName];
      }
    }
    target.status = "Merged";
  }
  return merged;
}

export function matchPaymentProofs(proofs: PaymentProof[], items: ReceiptItem[]): PaymentProof[] {
  return proofs.map((proof) => {
    const proofAmount = safeNumber(proof.amount);
    const proofDate = parseDateKey(proof.date);
    const matches = items.filter((item) => {
      const itemAmount = safeNumber(item.amount);
      const itemDate = parseDateKey(item.date);
      return Boolean(
        proofDate &&
          itemDate &&
          proofDate === itemDate &&
          proofAmount !== undefined &&
          itemAmount !== undefined &&
          Math.abs(itemAmount - proofAmount) < 0.01
      );
    });
    if (matches.length === 1) {
      return { ...proof, matchedReceiptId: matches[0].id, status: "Matched" };
    }
    return { ...proof, matchedReceiptId: "", status: "Needs manual review" };
  });
}

export function swapProofForReceipt(
  proofs: PaymentProof[],
  receiptId: string,
  currentProofIds: string[]
): { proofs: PaymentProof[]; selectedProofId?: string } {
  const nextProof = proofs.find((proof) => !currentProofIds.includes(proof.id));
  if (!nextProof) {
    return { proofs };
  }
  return {
    selectedProofId: nextProof.id,
    proofs: proofs.map((proof) => {
      if (currentProofIds.includes(proof.id)) return { ...proof, matchedReceiptId: "", status: "Needs manual review" };
      if (proof.id === nextProof.id) return { ...proof, matchedReceiptId: receiptId, status: "Matched manually" };
      return proof;
    })
  };
}

export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}
