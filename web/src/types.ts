export type FormVersion = "USA" | "Korea";
export type Currency = "USD" | "KRW" | "RMB" | "CNY";
export type DocumentType = "unknown" | "receipt" | "payment_proof";
export type SupportRole = "unknown" | "receipt" | "companion_ticket" | "payment_proof" | "ignore";
export type RoleSource = "manual" | "ai" | "memory" | "";
export type DocumentEvidence = "unknown" | "photographed_paper" | "merchant_document" | "bank_portal";

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
  fingerprint?: string;
  supportRole?: SupportRole;
  roleSource?: RoleSource;
  documentEvidence?: DocumentEvidence;
  transactionDate?: string;
  postedDate?: string;
  bankName?: string;
  cardLast4?: string;
  transactionReference?: string;
  pageAmount?: string;
  pageVendor?: string;
  cropPoints?: CropPoint[];
  rotationDegrees: number;
  rotationLocked?: boolean;
  sourcePage?: number;
  pageCount?: number;
  isPdfPage?: boolean;
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
  documentType: DocumentType;
  lockedDocumentType?: DocumentType;
  documentEvidence?: DocumentEvidence;
  transactionDate?: string;
  postedDate?: string;
  bankName?: string;
  cardLast4?: string;
  transactionReference?: string;
  manualFields?: string[];
  matchedReceiptId: string;
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
  usdToKrw: number;
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
  document_type?: string;
  receipt_image_indexes?: string;
  payment_proof_image_indexes?: string;
  image_roles?: string;
  document_evidence?: string;
  transaction_date?: string;
  posted_date?: string;
  bank_name?: string;
  card_last4?: string;
  transaction_reference?: string;
  page_results?: PageExtraction[];
}

export interface PageExtraction {
  page_index: number;
  role: string;
  rotation_degrees: number;
  evidence: string;
  transaction_date: string;
  posted_date: string;
  amount: string;
  vendor: string;
  bank_name: string;
  card_last4: string;
  transaction_reference: string;
}

export interface SelectedTile {
  kind: "receipt" | "proof";
  receiptId: string;
  imageId?: string;
  proofId?: string;
}
