import type { Category } from "./types";

export const DEFAULT_USD_TO_RMB = 6.8175;
export const DEFAULT_USD_TO_KRW = 1548.86;
export const DEFAULT_KRW_TO_RMB = 0.004433;
export const DEFAULT_MODEL = "gpt-5.4-mini";
export const ADVANCED_MODEL = "gpt-5.5";

export const USA_TEMPLATE_URL = "./templates/usa_expense_report_template.xlsx";
export const KOREA_TEMPLATE_URL = "./templates/korea_reimbursement_template.xlsx";
export const FORM_VERSION_STORAGE_KEY = "reimbursement-helper-web-form-version";
export const LANGUAGE_STORAGE_KEY = "reimbursement-helper-web-language";

export const USA_REPORT_CATEGORY_ORDER = ["transportation", "lodging", "meals", "advertising", "office", "entertainment", "other"] as const;
export const KOREA_REPORT_CATEGORY_ORDER = [
  "transportation",
  "physical_exam",
  "lodging",
  "nucleic_test",
  "materials",
  "meals",
  "courier",
  "consumables",
  "welfare",
  "other"
] as const;

export const KOREA_INVOICE_KIND_ORDER = [
  "car_rental",
  "fuel",
  "parking",
  "transportation",
  "esim",
  "materials",
  "physical_exam",
  "lodging",
  "nucleic_test",
  "meals",
  "courier",
  "consumables",
  "welfare",
  "other"
] as const;

export const USA_CATEGORY_LABELS: Record<Category, string> = {
  transportation: "Transportation",
  lodging: "Lodging",
  meals: "Meals",
  advertising: "Advertising",
  office: "Office",
  entertainment: "Entertainment",
  materials: "Materials",
  consumables: "Consumables",
  physical_exam: "Physical exam",
  nucleic_test: "Nucleic acid test",
  courier: "Courier",
  welfare: "Welfare",
  other: "Other"
};

export const KOREA_CATEGORY_LABELS: Record<Category, string> = {
  transportation: "Transportation / 交通费",
  physical_exam: "Physical exam / 入职体检费",
  lodging: "Accommodation / 住宿费",
  nucleic_test: "Nucleic acid test / 核酸检测费",
  materials: "Material / 物料费",
  meals: "Meals / 业务招待费-餐费",
  courier: "Courier / 快递费",
  consumables: "Consumables / 消耗品",
  welfare: "Welfare / 福利费",
  advertising: "Advertising",
  office: "Office",
  entertainment: "Entertainment",
  other: "Other / 其他"
};

export const USA_CATEGORY_ROWS: Record<string, number[]> = {
  transportation: Array.from({ length: 40 }, (_, index) => index + 7),
  lodging: [49, 50, 51],
  meals: [54],
  advertising: [57],
  office: [60, 61, 62, 63, 64],
  entertainment: [67, 68],
  other: [60, 61, 62, 63, 64]
};

export const KOREA_CATEGORY_COLUMNS: Record<string, string> = {
  transportation: "F",
  physical_exam: "G",
  lodging: "H",
  nucleic_test: "I",
  materials: "J",
  meals: "K",
  courier: "L",
  consumables: "M",
  welfare: "N",
  other: "O"
};

export const KOREA_COVER_ROWS: Record<string, { row: number; label: string }> = {
  transportation: { row: 4, label: "交通费" },
  consumables: { row: 5, label: "消耗品" },
  lodging: { row: 6, label: "住宿费" },
  meals: { row: 7, label: "业务招待费/餐费" },
  other: { row: 8, label: "其他" }
};
