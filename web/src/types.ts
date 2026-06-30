export type FormVersion = "USA" | "Korea";
export type Currency = "USD" | "KRW" | "RMB" | "CNY";

export type Category =
  | "transportation"
  | "lodging"
  | "meals"
  | "advertising"
  | "office"
  | "entertainment"
  | "materials"
  | "consumables"
  | "physical_exam"
  | "nucleic_test"
  | "courier"
  | "welfare"
  | "other";

export interface CropPoint {
  x: number;
  y: number;
}

export interface ImageAttachment {
  id: string;
  filename: string;
  sourceName: string;
  dataUrl: string;
  width: number;
  height: number;
  cropPoints?: CropPoint[];
  rotationDegrees: number;
}

export interface ReceiptItem {
  id: string;
  filename: string;
  status: string;
  date: string;
  place: string;
  amount: string;
  currency: Currency;
  krwAmount: string;
  rmbAmount: string;
  purpose: string;
  details: string;
  projectNumber: string;
  category: Category;
  paymentMethod: string;
  receiptLabel: string;
  images: ImageAttachment[];
}

export interface PaymentProof {
  id: string;
  filename: string;
  status: string;
  date: string;
  amount: string;
  place: string;
  matchedReceiptId: string;
  image: ImageAttachment;
}

export interface ExchangeRates {
  usdToRmb: number;
  krwToRmb: number;
}

export interface ReceiptExtraction {
  date?: string;
  place?: string;
  vendor?: string;
  amount?: string | number;
  currency?: string;
  krw_amount?: string | number;
  rmb_amount?: string | number;
  purpose?: string;
  details?: string;
  project_number?: string;
  category?: string;
  payment_method?: string;
  receipt_label?: string;
}

export interface SelectedTile {
  kind: "receipt" | "proof";
  receiptId: string;
  imageId?: string;
  proofId?: string;
}
