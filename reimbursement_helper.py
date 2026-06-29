from __future__ import annotations

import base64
import copy
import json
import logging
import os
import re
import shutil
import sys
import threading
import traceback
import uuid
from dataclasses import asdict, dataclass
from datetime import date, datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple
from urllib import error, request

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except Exception:  # pragma: no cover - import guard for friendly CLI errors
    tk = None  # type: ignore[assignment]
    filedialog = None  # type: ignore[assignment]
    messagebox = None  # type: ignore[assignment]
    ttk = None  # type: ignore[assignment]

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.cell.cell import MergedCell
    from openpyxl.drawing.image import Image as XLImage
except Exception:  # pragma: no cover
    Workbook = None  # type: ignore[assignment]
    load_workbook = None  # type: ignore[assignment]
    MergedCell = None  # type: ignore[assignment]
    XLImage = None  # type: ignore[assignment]

try:
    from PIL import Image, ImageOps, ImageTk
except Exception:  # pragma: no cover
    Image = None  # type: ignore[assignment]
    ImageOps = None  # type: ignore[assignment]
    ImageTk = None  # type: ignore[assignment]


APP_DIR = Path(__file__).resolve().parent
CONFIG_DIR = APP_DIR / "config"
TEMPLATE_DIR = APP_DIR / "templates"
OUTPUT_DIR = APP_DIR / "outputs"
WORK_DIR = APP_DIR / "work"
LOG_DIR = APP_DIR / "logs"
USER_SETTINGS_FILE = CONFIG_DIR / "user_settings.json"
USER_API_KEY_FILE = CONFIG_DIR / "api_key.txt"
TEMPLATES_CONFIG_FILE = CONFIG_DIR / "templates.json"
SUPPORTED_IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif"}
OPENAI_CHAT_COMPLETIONS_URL = "https://api.openai.com/v1/chat/completions"
DEFAULT_OPENAI_MODEL = "gpt-4o-mini"
DEFAULT_USD_TO_RMB_RATE = "6.8175"
DEFAULT_KRW_TO_RMB_RATE = "0.004433"
LOGGER = logging.getLogger("reimbursement_helper")


USA_TEMPLATE_NAME = "usa_expense_report_template.xlsx"
KOREA_COVER_TEMPLATE_NAME = "korea_cover_receipts_template.xlsx"
KOREA_DETAILS_TEMPLATE_NAME = "korea_details_template.xlsx"


def setup_logging() -> None:
    if LOGGER.handlers:
        return
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    handler = RotatingFileHandler(
        LOG_DIR / "app.log",
        maxBytes=750_000,
        backupCount=5,
        encoding="utf-8",
    )
    handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    )
    LOGGER.setLevel(logging.INFO)
    LOGGER.addHandler(handler)
    LOGGER.propagate = False
    LOGGER.info("Reimbursement Helper logging started")


def log_exception(message: str) -> Path:
    setup_logging()
    LOGGER.exception(message)
    return LOG_DIR / "app.log"


USA_CATEGORY_ROWS: Dict[str, List[int]] = {
    "transportation": list(range(7, 47)),
    "lodging": list(range(49, 52)),
    "meals": [54],
    "advertising": [57],
    "office": list(range(60, 65)),
    "entertainment": list(range(67, 69)),
    "other": list(range(60, 65)),
}

KOREA_CATEGORY_COLUMNS: Dict[str, str] = {
    "transportation": "F",
    "physical_exam": "G",
    "lodging": "H",
    "nucleic_test": "I",
    "materials": "J",
    "meals": "K",
    "courier": "L",
    "consumables": "M",
    "welfare": "N",
    "other": "O",
}

KOREA_COVER_ROWS: Dict[str, Tuple[int, str]] = {
    "transportation": (4, "交通费"),
    "consumables": (5, "消耗品"),
    "lodging": (6, "住宿费"),
    "meals": (7, "业务招待费/餐费"),
    "other": (8, "其他"),
}

USA_CATEGORY_LABELS = {
    "transportation": "Transportation",
    "lodging": "Lodging",
    "meals": "Meals",
    "advertising": "Advertising",
    "office": "Office",
    "entertainment": "Entertainment",
    "other": "Other",
}

KOREA_CATEGORY_LABELS = {
    "transportation": "Transportation / 交通费",
    "physical_exam": "Physical exam / 入职体检费",
    "lodging": "Accommodation / 住宿费",
    "nucleic_test": "Nucleic acid test / 核酸检测费",
    "materials": "Material / 物料费",
    "meals": "Meals / 业务招待费-餐费",
    "courier": "Courier / 快递费",
    "consumables": "Consumables / 消耗品",
    "welfare": "Welfare / 福利费",
    "other": "Other / 其他",
}


KOREA_COVER_ROWS = {
    "transportation": (4, "\u4ea4\u901a\u8d39"),
    "consumables": (5, "\u6d88\u8017\u54c1"),
    "lodging": (6, "\u4f4f\u5bbf\u8d39"),
    "meals": (7, "\u4e1a\u52a1\u62db\u5f85\u8d39/\u9910\u8d39"),
    "other": (8, "\u5176\u4ed6"),
}

KOREA_CATEGORY_LABELS = {
    "transportation": "Transportation / \u4ea4\u901a\u8d39",
    "physical_exam": "Physical exam / \u5165\u804c\u4f53\u68c0\u8d39",
    "lodging": "Accommodation / \u4f4f\u5bbf\u8d39",
    "nucleic_test": "Nucleic acid test / \u6838\u9178\u68c0\u6d4b\u8d39",
    "materials": "Material / \u7269\u6599\u8d39",
    "meals": "Meals / \u4e1a\u52a1\u62db\u5f85\u8d39-\u9910\u8d39",
    "courier": "Courier / \u5feb\u9012\u8d39",
    "consumables": "Consumables / \u6d88\u8017\u54c1",
    "welfare": "Welfare / \u798f\u5229\u8d39",
    "other": "Other / \u5176\u4ed6",
}


@dataclass
class ReceiptItem:
    item_id: str
    path: str
    filename: str
    date: str = ""
    place: str = ""
    amount: str = ""
    currency: str = "USD"
    krw_amount: str = ""
    rmb_amount: str = ""
    purpose: str = ""
    details: str = ""
    project_number: str = ""
    category: str = "transportation"
    payment_method: str = ""
    receipt_label: str = ""
    status: str = "Empty"


def appdata_daily_logger_key_file() -> Optional[Path]:
    appdata = os.getenv("APPDATA", "").strip()
    if not appdata:
        return None
    return Path(appdata) / "DailyLogger" / "settings" / "daily_logger_api_key.txt"


def load_user_settings() -> Dict[str, Any]:
    if not USER_SETTINGS_FILE.exists():
        return {}
    try:
        data = json.loads(USER_SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def load_templates_config() -> Dict[str, Any]:
    if not TEMPLATES_CONFIG_FILE.exists():
        return {}
    try:
        data = json.loads(TEMPLATES_CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def parse_env_file_for_key(path: Path) -> str:
    if not path.exists():
        return ""
    try:
        for line in path.read_text(encoding="utf-8").splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if stripped.startswith("OPENAI_API_KEY="):
                return stripped.split("=", 1)[1].strip().strip('"').strip("'")
    except Exception:
        return ""
    return ""


def get_openai_api_key() -> str:
    key = os.getenv("OPENAI_API_KEY", "").strip()
    if key:
        return key
    key = parse_env_file_for_key(APP_DIR / ".env.local") or parse_env_file_for_key(APP_DIR / ".env")
    if key:
        return key
    if USER_API_KEY_FILE.exists():
        try:
            key = USER_API_KEY_FILE.read_text(encoding="utf-8").strip()
            if key:
                return key
        except Exception:
            pass
    daily_logger_key = appdata_daily_logger_key_file()
    if daily_logger_key and daily_logger_key.exists():
        try:
            key = daily_logger_key.read_text(encoding="utf-8").strip()
            if key:
                return key
        except Exception:
            pass
    return ""


def image_mime_for_path(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".jpg", ".jpeg"}:
        return "image/jpeg"
    if suffix == ".webp":
        return "image/webp"
    if suffix == ".gif":
        return "image/gif"
    if suffix == ".bmp":
        return "image/bmp"
    return "image/png"


def image_data_url(path: Path) -> str:
    raw = path.read_bytes()
    return f"data:{image_mime_for_path(path)};base64,{base64.b64encode(raw).decode('ascii')}"


def extract_json_object(text: str) -> Dict[str, Any]:
    cleaned = text.strip()
    cleaned = re.sub(r"^```(?:json)?", "", cleaned, flags=re.IGNORECASE).strip()
    cleaned = re.sub(r"```$", "", cleaned).strip()
    try:
        parsed = json.loads(cleaned)
        return parsed if isinstance(parsed, dict) else {}
    except Exception:
        pass
    match = re.search(r"\{.*\}", cleaned, flags=re.DOTALL)
    if not match:
        return {}
    try:
        parsed = json.loads(match.group(0))
        return parsed if isinstance(parsed, dict) else {}
    except Exception:
        return {}


def call_openai_receipt_extraction(
    image_path: Path,
    form_version: str,
    existing: ReceiptItem,
    progress: Optional[Callable[[str], None]] = None,
) -> Dict[str, Any]:
    api_key = get_openai_api_key()
    if not api_key:
        raise RuntimeError(
            "No OpenAI API key was found. Add one to .env.local, config/api_key.txt, "
            "or Daily Logger settings, then try again."
        )

    def notify(message: str) -> None:
        if progress:
            try:
                progress(message)
            except Exception:
                pass

    notify("Reading receipt image")
    data_url = image_data_url(image_path)
    form_hint = "USA VisionNav reimbursement form" if form_version == "USA" else "Korea VisionNav reimbursement form"
    prompt = f"""
Extract reimbursement fields from this receipt image for a {form_hint}.
Return only compact JSON with these keys:
date, place, amount, currency, krw_amount, rmb_amount, purpose, details,
project_number, category, payment_method, receipt_label, confidence_notes.

Rules:
- date must be yyyy-mm-dd when visible, otherwise empty string.
- amount should be the paid total. Preserve cents when visible.
- category must be one of:
  transportation, lodging, meals, advertising, office, entertainment, physical_exam,
  nucleic_test, materials, courier, consumables, welfare, other.
- project_number should stay empty unless printed on the receipt or obvious from context.
- purpose should be short, for example Parking, Gas, Hotel, Meal, Baggage, Supplies.
- details should be a natural one-line description.
- receipt_label should be short enough to fit above a receipt image in Excel.
- If a field is unclear, use an empty string rather than guessing.

Current editable row values, to keep if they look more specific than the image:
{json.dumps(asdict(existing), ensure_ascii=False)}
""".strip()
    payload = json.dumps(
        {
            "model": os.getenv("OPENAI_MODEL", DEFAULT_OPENAI_MODEL),
            "messages": [
                {
                    "role": "system",
                    "content": (
                        "You extract structured receipt data for reimbursement forms. "
                        "You never invent project numbers, names, or hidden values."
                    ),
                },
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": data_url}},
                    ],
                },
            ],
            "temperature": 0,
        }
    ).encode("utf-8")
    req = request.Request(
        OPENAI_CHAT_COMPLETIONS_URL,
        data=payload,
        method="POST",
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
    )
    notify("Contacting AI")
    try:
        with request.urlopen(req, timeout=90) as response:
            body = response.read().decode("utf-8")
    except error.HTTPError as exc:
        try:
            details = exc.read().decode("utf-8")
        except Exception:
            details = str(exc)
        raise RuntimeError(f"OpenAI API error ({exc.code}): {details}") from exc
    except Exception as exc:
        raise RuntimeError(f"Could not contact OpenAI: {exc}") from exc

    notify("Applying extracted fields")
    try:
        parsed = json.loads(body)
        content = parsed["choices"][0]["message"]["content"]
    except Exception as exc:
        raise RuntimeError("OpenAI returned an unexpected response format.") from exc
    data = extract_json_object(str(content))
    if not data:
        raise RuntimeError("AI response did not contain usable JSON.")
    return data


def ensure_dependencies() -> None:
    if tk is None or ttk is None:
        raise RuntimeError("Tkinter is required to run the desktop app.")
    if load_workbook is None or Workbook is None or XLImage is None:
        raise RuntimeError("openpyxl is required. Install dependencies from requirements.txt.")
    if Image is None or ImageTk is None:
        raise RuntimeError("Pillow is required. Install dependencies from requirements.txt.")


def safe_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    try:
        return float(match.group(0))
    except ValueError:
        return None


def format_amount(value: float) -> str:
    if abs(value) >= 1000:
        return f"{value:.2f}"
    return f"{value:.2f}"


def parse_date_value(value: str) -> Any:
    text = (value or "").strip()
    if not text:
        return None
    formats = [
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%Y.%m.%d",
        "%Y/%m/%d",
        "%m-%d-%Y",
        "%B %d, %Y",
        "%b %d, %Y",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return text


def excel_date_formula_or_value(value: str) -> Any:
    parsed = parse_date_value(value)
    if isinstance(parsed, date):
        return parsed
    return parsed


def normalized_category(category: str, form_version: str) -> str:
    raw = (category or "").strip().lower()
    lookup = {
        "transport": "transportation",
        "transportation": "transportation",
        "parking": "transportation",
        "gas": "transportation",
        "fuel": "transportation",
        "uber": "transportation",
        "taxi": "transportation",
        "lodging": "lodging",
        "hotel": "lodging",
        "baggage": "lodging",
        "meal": "meals",
        "meals": "meals",
        "food": "meals",
        "restaurant": "meals",
        "entertainment": "entertainment",
        "advertising": "advertising",
        "office": "office",
        "material": "materials" if form_version == "Korea" else "office",
        "materials": "materials" if form_version == "Korea" else "office",
        "consumable": "consumables" if form_version == "Korea" else "office",
        "consumables": "consumables" if form_version == "Korea" else "office",
        "courier": "courier",
        "physical_exam": "physical_exam",
        "physical exam": "physical_exam",
        "nucleic_test": "nucleic_test",
        "nucleic test": "nucleic_test",
        "welfare": "welfare",
        "other": "other",
    }
    if raw in lookup:
        return lookup[raw]
    if form_version == "Korea" and raw in KOREA_CATEGORY_COLUMNS:
        return raw
    if form_version == "USA" and raw in USA_CATEGORY_ROWS:
        return raw
    return "other"


def prepare_excel_image_file(path: Path) -> Path:
    if Image is None:
        return path
    target_dir = WORK_DIR / "excel_images"
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / f"{uuid.uuid4().hex}.png"
    with Image.open(path) as source:
        image = ImageOps.exif_transpose(source) if ImageOps is not None else source.copy()
        if image.mode not in {"RGB", "RGBA"}:
            image = image.convert("RGB")
        image.save(target, "PNG")
    return target


def resize_excel_image(path: Path, max_width: int, max_height: int) -> Any:
    excel_path = prepare_excel_image_file(path)
    img = XLImage(str(excel_path))
    if Image is None:
        img.width = max_width
        img.height = max_height
        return img
    with Image.open(excel_path) as pil_img:
        width, height = pil_img.size
    if width <= 0 or height <= 0:
        img.width = max_width
        img.height = max_height
        return img
    scale = min(max_width / width, max_height / height, 1.0)
    img.width = max(1, int(width * scale))
    img.height = max(1, int(height * scale))
    return img


def clear_images(sheet: Any) -> None:
    try:
        sheet._images = []
    except Exception:
        pass


def clear_cells(sheet: Any, rows: Iterable[int], cols: Sequence[str]) -> None:
    for row in rows:
        for col in cols:
            cell = sheet[f"{col}{row}"]
            if MergedCell is not None and isinstance(cell, MergedCell):
                continue
            cell.value = None


def fit_columns_for_receipts(sheet: Any) -> None:
    for col in "ABCDEFGH":
        sheet.column_dimensions[col].width = 16


def export_usa(items: List[ReceiptItem], output_path: Path, exchange_rate: float) -> None:
    template = TEMPLATE_DIR / USA_TEMPLATE_NAME
    if not template.exists():
        raise FileNotFoundError(f"Missing USA template: {template}")
    wb = load_workbook(template)
    ws = wb["Expense report"]
    receipts_ws = wb["Receipt and Payment of expenses"]

    ws["A3"] = f"Date / 填表日期： {date.today().strftime('%m/%d/%Y')}"
    ws["A4"] = "Employee: / 申请人："
    ws["J1"] = exchange_rate
    all_entry_rows = sorted({row for rows in USA_CATEGORY_ROWS.values() for row in rows})
    clear_cells(ws, all_entry_rows, ["A", "B", "C", "D", "E", "F", "H"])
    for row in all_entry_rows:
        ws[f"G{row}"] = f"=F{row}*$J$1"

    row_cursors = {cat: 0 for cat in USA_CATEGORY_ROWS}
    for item in items:
        category = normalized_category(item.category, "USA")
        if category not in USA_CATEGORY_ROWS:
            category = "other"
        rows = USA_CATEGORY_ROWS[category]
        cursor = row_cursors[category]
        if cursor >= len(rows):
            raise RuntimeError(f"Not enough USA template rows for {USA_CATEGORY_LABELS.get(category, category)}.")
        row = rows[cursor]
        row_cursors[category] += 1
        amount = safe_float(item.amount)
        ws[f"A{row}"] = item.place.strip()
        ws[f"B{row}"] = excel_date_formula_or_value(item.date)
        ws[f"C{row}"] = item.details.strip() or item.receipt_label.strip() or item.filename
        ws[f"D{row}"] = item.purpose.strip()
        ws[f"E{row}"] = item.project_number.strip()
        ws[f"F{row}"] = amount if amount is not None else None
        ws[f"G{row}"] = f"=F{row}*$J$1"

    clear_images(receipts_ws)
    for row in range(2, max(56, len(items) + 3)):
        for col in "ABCDE":
            receipts_ws[f"{col}{row}"] = None
    receipts_ws.column_dimensions["D"].width = 38
    receipts_ws.column_dimensions["E"].width = 26
    for index, item in enumerate(items, start=1):
        row = index + 1
        receipts_ws[f"A{row}"] = index
        receipts_ws[f"B{row}"] = excel_date_formula_or_value(item.date)
        receipts_ws[f"C{row}"] = safe_float(item.amount)
        receipts_ws.row_dimensions[row].height = 126
        image_path = Path(item.path)
        if image_path.exists():
            receipts_ws.add_image(resize_excel_image(image_path, 260, 150), f"D{row}")

    wb.save(output_path)


def clone_sheet(source_ws: Any, target_wb: Any, title: str, index: Optional[int] = None) -> Any:
    target_ws = target_wb.create_sheet(title=title, index=index)
    for row in source_ws.iter_rows():
        for source_cell in row:
            target_cell = target_ws[source_cell.coordinate]
            target_cell.value = source_cell.value
            if source_cell.has_style:
                target_cell._style = copy.copy(source_cell._style)
            if source_cell.number_format:
                target_cell.number_format = source_cell.number_format
            if source_cell.font:
                target_cell.font = copy.copy(source_cell.font)
            if source_cell.fill:
                target_cell.fill = copy.copy(source_cell.fill)
            if source_cell.border:
                target_cell.border = copy.copy(source_cell.border)
            if source_cell.alignment:
                target_cell.alignment = copy.copy(source_cell.alignment)
            if source_cell.protection:
                target_cell.protection = copy.copy(source_cell.protection)
            if source_cell.comment:
                target_cell.comment = copy.copy(source_cell.comment)
    for merged in source_ws.merged_cells.ranges:
        target_ws.merge_cells(str(merged))
    for key, dim in source_ws.column_dimensions.items():
        target_ws.column_dimensions[key] = copy.copy(dim)
    for key, dim in source_ws.row_dimensions.items():
        target_ws.row_dimensions[key] = copy.copy(dim)
    target_ws.freeze_panes = source_ws.freeze_panes
    target_ws.sheet_view.showGridLines = source_ws.sheet_view.showGridLines
    if source_ws.print_area:
        target_ws.print_area = source_ws.print_area
    return target_ws


def korea_amounts(item: ReceiptItem, krw_to_rmb_rate: float) -> Tuple[Optional[float], Optional[float], str]:
    amount = safe_float(item.amount)
    krw = safe_float(item.krw_amount)
    rmb = safe_float(item.rmb_amount)
    currency = (item.currency or "").strip().upper() or "KRW"
    if krw is None and rmb is None and amount is not None:
        if currency == "RMB" or currency == "CNY":
            rmb = amount
            krw = amount / krw_to_rmb_rate if krw_to_rmb_rate else None
        elif currency == "KRW":
            krw = amount
            rmb = amount * krw_to_rmb_rate if krw_to_rmb_rate else None
    elif krw is None and rmb is not None and krw_to_rmb_rate:
        krw = rmb / krw_to_rmb_rate
    elif rmb is None and krw is not None and krw_to_rmb_rate:
        rmb = krw * krw_to_rmb_rate
    note = ""
    if amount is not None and currency not in {"KRW", "RMB", "CNY"}:
        note = f"({amount:g} {currency})"
    return krw, rmb, note


def korea_cover_bucket(category: str) -> str:
    if category == "transportation":
        return "transportation"
    if category == "lodging":
        return "lodging"
    if category in {"meals", "entertainment"}:
        return "meals"
    if category in {"materials", "consumables", "office"}:
        return "consumables"
    return "other"


def export_korea(items: List[ReceiptItem], output_path: Path, krw_to_rmb_rate: float) -> None:
    cover_template = TEMPLATE_DIR / KOREA_COVER_TEMPLATE_NAME
    details_template = TEMPLATE_DIR / KOREA_DETAILS_TEMPLATE_NAME
    if not cover_template.exists():
        raise FileNotFoundError(f"Missing Korea cover template: {cover_template}")
    if not details_template.exists():
        raise FileNotFoundError(f"Missing Korea details template: {details_template}")

    wb = load_workbook(cover_template)
    detail_template_wb = load_workbook(details_template)
    cover_ws = wb.worksheets[0]
    receipts_ws = wb.worksheets[1]
    detail_ws = clone_sheet(detail_template_wb.worksheets[0], wb, "报销明细", index=1)

    cover_ws["A2"] = f"报销部门：  {date.today().year}年 {date.today().month}月 {date.today().day}日 填 单据及附件共  页"
    cover_ws["A11"] = "领导审批           会计主管              会计                  出纳                 报销人                   领款人 "
    for row, label in [(row, label) for row, label in KOREA_COVER_ROWS.values()]:
        cover_ws[f"A{row}"] = label
        cover_ws[f"C{row}"] = 0
        cover_ws[f"D{row}"] = 0
    cover_ws["C9"] = "=SUM(C4:C8)"
    cover_ws["D9"] = "=SUM(D4:D8)"
    cover_ws["B10"] = "=D9"

    for row in range(3, 34):
        for col in list("ABCDEFGHIJKLMNOPQRS"):
            detail_ws[f"{col}{row}"] = None
    detail_ws["A34"] = "合计（外币）\nTotal"
    detail_ws["Q34"] = "=SUM(Q3:Q33)"
    detail_ws["A35"] = "合计（人民币）\nTotal"
    detail_ws["R35"] = "=SUM(R3:R34)"

    summary: Dict[str, Tuple[float, float]] = {
        key: (0.0, 0.0) for key in KOREA_COVER_ROWS
    }
    for index, item in enumerate(items, start=3):
        if index > 33:
            raise RuntimeError("Korea template supports up to 31 detail rows in this version.")
        category = normalized_category(item.category, "Korea")
        if category not in KOREA_CATEGORY_COLUMNS:
            category = "other"
        krw, rmb, original_note = korea_amounts(item, krw_to_rmb_rate)
        detail_ws[f"A{index}"] = excel_date_formula_or_value(item.date)
        detail_ws[f"B{index}"] = item.purpose.strip() or item.details.strip()
        detail_ws[f"C{index}"] = item.place.strip()
        detail_ws[f"D{index}"] = ""
        detail_ws[f"E{index}"] = item.project_number.strip()
        category_col = KOREA_CATEGORY_COLUMNS[category]
        detail_ws[f"{category_col}{index}"] = (
            f"{original_note} {krw:.0f} KRW".strip()
            if original_note and krw is not None
            else krw
        )
        detail_ws[f"Q{index}"] = krw
        detail_ws[f"R{index}"] = rmb
        detail_ws[f"S{index}"] = item.payment_method.strip()
        bucket = korea_cover_bucket(category)
        cur_krw, cur_rmb = summary[bucket]
        summary[bucket] = (cur_krw + (krw or 0.0), cur_rmb + (rmb or 0.0))

    for bucket, (krw_total, rmb_total) in summary.items():
        row, _label = KOREA_COVER_ROWS[bucket]
        cover_ws[f"C{row}"] = round(krw_total, 2) if krw_total else None
        cover_ws[f"D{row}"] = round(rmb_total, 2) if rmb_total else None

    clear_images(receipts_ws)
    fit_columns_for_receipts(receipts_ws)
    for row in range(1, max(240, len(items) * 40 + 20)):
        for col in "ABCDEFGH":
            receipts_ws[f"{col}{row}"] = None
    anchors = ["A", "D", "G"]
    row_step = 34
    for idx, item in enumerate(items):
        band = idx // len(anchors)
        slot = idx % len(anchors)
        label_row = 1 + band * row_step
        image_row = label_row + 1
        col = anchors[slot]
        label = item.receipt_label.strip() or item.details.strip() or item.filename
        receipts_ws[f"{col}{label_row}"] = label[:60]
        image_path = Path(item.path)
        if image_path.exists():
            receipts_ws.add_image(resize_excel_image(image_path, 280, 390), f"{col}{image_row}")

    wb.save(output_path)


class ReimbursementHelperApp:
    def __init__(self, root: Any) -> None:
        self.root = root
        self.items: List[ReceiptItem] = []
        self.selected_index: Optional[int] = None
        self.photo: Optional[Any] = None
        self.settings = load_user_settings()
        self.form_version = tk.StringVar(value="USA")
        self.exchange_rate = tk.StringVar(value=str(self.settings.get("usa_exchange_rate", DEFAULT_USD_TO_RMB_RATE)))
        self.krw_to_rmb_rate = tk.StringVar(value=str(self.settings.get("krw_to_rmb_rate", DEFAULT_KRW_TO_RMB_RATE)))
        self.status_text = tk.StringVar(value="Ready")
        self.field_vars: Dict[str, Any] = {}
        self.field_labels: Dict[str, Any] = {}
        self.field_widgets: Dict[str, Any] = {}
        self.category_values: List[str] = []
        self.category_combo: Optional[Any] = None
        self.usa_rate_label: Optional[Any] = None
        self.usa_rate_entry: Optional[Any] = None
        self.krw_rate_label: Optional[Any] = None
        self.krw_rate_entry: Optional[Any] = None
        self._loading_fields = False
        self._syncing_amounts = False
        self._last_amount_source = "amount"
        self._build_ui()
        self._attach_amount_traces()
        self._on_form_version_changed()

    def _build_ui(self) -> None:
        self.root.title("Reimbursement Helper")
        self.root.geometry("1220x760")
        self.root.minsize(980, 640)
        self.root.configure(bg="#F7F4EC")
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background="#F7F4EC")
        style.configure("Panel.TFrame", background="#FFFFFF")
        style.configure("TLabel", background="#F7F4EC", foreground="#222222", font=("Segoe UI", 10))
        style.configure("Header.TLabel", background="#F7F4EC", foreground="#222222", font=("Segoe UI", 13, "bold"))
        style.configure("Muted.TLabel", background="#F7F4EC", foreground="#666666", font=("Segoe UI", 9))
        style.configure("TButton", font=("Segoe UI", 10))

        header = ttk.Frame(self.root, padding=(18, 16, 18, 10))
        header.pack(fill="x")
        ttk.Label(header, text="REIMBURSEMENT HELPER", style="Header.TLabel").pack(side="left")
        controls = ttk.Frame(header)
        controls.pack(side="right")
        ttk.Label(controls, text="Form").pack(side="left", padx=(0, 6))
        form_combo = ttk.Combobox(
            controls,
            textvariable=self.form_version,
            values=["USA", "Korea"],
            state="readonly",
            width=10,
        )
        form_combo.pack(side="left", padx=(0, 12))
        form_combo.bind("<<ComboboxSelected>>", lambda _event: self._on_form_version_changed())
        ttk.Button(controls, text="Upload receipts", command=self.upload_receipts).pack(side="left", padx=4)
        ttk.Button(controls, text="Generate Details", command=self.generate_selected_details).pack(side="left", padx=4)
        ttk.Button(controls, text="Generate All", command=self.generate_all_details).pack(side="left", padx=4)
        ttk.Button(controls, text="Generate Excel", command=self.generate_excel).pack(side="left", padx=4)

        body = ttk.Frame(self.root, padding=(14, 0, 14, 10))
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(0, minsize=650, weight=0)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        manager = ttk.Frame(body, style="Panel.TFrame", padding=12)
        manager.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        manager.grid_columnconfigure(0, minsize=270, weight=1)
        manager.grid_columnconfigure(1, minsize=350, weight=1)
        manager.grid_rowconfigure(1, weight=1)
        ttk.Label(manager, text="Inserted receipts and details", style="Header.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 8)
        )

        left = ttk.Frame(manager, style="Panel.TFrame", padding=(0, 0, 10, 0))
        left.grid(row=1, column=0, sticky="nsew")
        self.tree = ttk.Treeview(left, columns=("status", "date", "amount"), show="tree headings", height=20)
        self.tree.heading("#0", text="File")
        self.tree.heading("status", text="Status")
        self.tree.heading("date", text="Date")
        self.tree.heading("amount", text="Amount")
        self.tree.column("#0", width=140, stretch=True)
        self.tree.column("status", width=78, stretch=False)
        self.tree.column("date", width=78, stretch=False)
        self.tree.column("amount", width=70, stretch=False)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        left_buttons = ttk.Frame(left, style="Panel.TFrame")
        left_buttons.pack(fill="x", pady=(10, 0))
        ttk.Button(left_buttons, text="Remove", command=self.remove_selected).pack(side="left")
        ttk.Button(left_buttons, text="Clear", command=self.clear_all).pack(side="left", padx=(8, 0))

        middle = ttk.Frame(manager, style="Panel.TFrame", padding=(10, 0, 0, 0))
        middle.grid(row=1, column=1, sticky="nsew")
        ttk.Label(middle, text="Details", style="Header.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        rate_frame = ttk.Frame(middle, style="Panel.TFrame")
        rate_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 10))
        self.usa_rate_label = ttk.Label(rate_frame, text="USD -> RMB")
        self.usa_rate_label.pack(side="left")
        self.usa_rate_entry = ttk.Entry(rate_frame, textvariable=self.exchange_rate, width=9)
        self.usa_rate_entry.pack(side="left", padx=(6, 14))
        self.krw_rate_label = ttk.Label(rate_frame, text="KRW -> RMB")
        self.krw_rate_label.pack(side="left")
        self.krw_rate_entry = ttk.Entry(rate_frame, textvariable=self.krw_to_rmb_rate, width=9)
        self.krw_rate_entry.pack(side="left", padx=(6, 0))

        fields = [
            ("date", "Date"),
            ("place", "Place / Vendor"),
            ("amount", "USD amount"),
            ("currency", "Currency"),
            ("krw_amount", "KRW amount"),
            ("rmb_amount", "RMB amount"),
            ("purpose", "Purpose"),
            ("details", "Details"),
            ("project_number", "Project number"),
            ("category", "Category"),
            ("payment_method", "Payment method"),
            ("receipt_label", "Receipt label"),
        ]
        for row, (key, label) in enumerate(fields, start=2):
            label_widget = ttk.Label(middle, text=label)
            label_widget.grid(row=row, column=0, sticky="w", pady=4)
            if key == "category":
                var = tk.StringVar()
                widget = ttk.Combobox(middle, textvariable=var, state="readonly", width=28)
                self.category_combo = widget
            elif key == "currency":
                var = tk.StringVar(value="USD")
                widget = ttk.Combobox(middle, textvariable=var, values=["USD", "KRW", "RMB", "CNY"], width=28)
            else:
                var = tk.StringVar()
                widget = ttk.Entry(middle, textvariable=var, width=31)
            self.field_vars[key] = var
            self.field_labels[key] = label_widget
            self.field_widgets[key] = widget
            widget.grid(row=row, column=1, sticky="ew", pady=4, padx=(8, 0))
        middle.grid_columnconfigure(1, weight=1)

        right = ttk.Frame(body, style="Panel.TFrame", padding=12)
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)
        ttk.Label(right, text="Receipt preview", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        preview_frame = ttk.Frame(right, style="Panel.TFrame")
        preview_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_label = tk.Label(
            preview_frame,
            bg="#FFFFFF",
            fg="#666666",
            text="Upload receipts to preview them here.",
            anchor="center",
            justify="center",
            bd=1,
            relief="solid",
        )
        self.preview_label.grid(row=0, column=0, sticky="nsew")

        footer = ttk.Frame(self.root, padding=(18, 0, 18, 12))
        footer.pack(fill="x")
        ttk.Label(footer, textvariable=self.status_text, style="Muted.TLabel").pack(side="left")
        ttk.Button(footer, text="Open output folder", command=self.open_output_folder).pack(side="right")

    def _attach_amount_traces(self) -> None:
        for key in ("amount", "krw_amount", "rmb_amount"):
            var = self.field_vars.get(key)
            if var is not None:
                var.trace_add("write", lambda *_args, k=key: self._sync_amount_fields(k))
        self.exchange_rate.trace_add("write", lambda *_args: self._sync_amount_fields("rate"))
        self.krw_to_rmb_rate.trace_add("write", lambda *_args: self._sync_amount_fields("rate"))

    def _sync_amount_fields(self, changed_key: str) -> None:
        if self._loading_fields or self._syncing_amounts:
            return
        self._syncing_amounts = True
        try:
            version = self.form_version.get()
            amount_var = self.field_vars.get("amount")
            krw_var = self.field_vars.get("krw_amount")
            rmb_var = self.field_vars.get("rmb_amount")
            if amount_var is None or krw_var is None or rmb_var is None:
                return
            if changed_key in {"amount", "krw_amount", "rmb_amount"}:
                self._last_amount_source = changed_key

            if version == "USA":
                rate = safe_float(self.exchange_rate.get())
                if not rate:
                    return
                if changed_key == "rmb_amount":
                    rmb = safe_float(rmb_var.get())
                    if rmb is not None:
                        amount_var.set(format_amount(rmb / rate))
                    return
                usd = safe_float(amount_var.get())
                if usd is not None:
                    rmb_var.set(format_amount(usd * rate))
                return

            rate = safe_float(self.krw_to_rmb_rate.get())
            if not rate:
                return
            if changed_key == "rmb_amount":
                rmb = safe_float(rmb_var.get())
                if rmb is not None:
                    krw_var.set(format_amount(rmb / rate))
                return
            krw = safe_float(krw_var.get())
            if krw is not None:
                rmb_var.set(format_amount(krw * rate))
        finally:
            self._syncing_amounts = False

    def _set_field_visible(self, key: str, visible: bool) -> None:
        label = self.field_labels.get(key)
        widget = self.field_widgets.get(key)
        if visible:
            if label is not None:
                label.grid()
            if widget is not None:
                widget.grid()
        else:
            if label is not None:
                label.grid_remove()
            if widget is not None:
                widget.grid_remove()

    def _update_field_visibility(self) -> None:
        version = self.form_version.get()
        if version == "USA":
            self.field_labels["amount"].configure(text="USD amount")
            self.field_labels["rmb_amount"].configure(text="RMB amount")
            self.field_vars["currency"].set("USD")
            self._set_field_visible("currency", False)
            self._set_field_visible("krw_amount", False)
            self._set_field_visible("amount", True)
            self._set_field_visible("rmb_amount", True)
            if self.usa_rate_label is not None:
                self.usa_rate_label.pack(side="left")
            if self.usa_rate_entry is not None:
                self.usa_rate_entry.pack(side="left", padx=(6, 14))
            if self.krw_rate_label is not None:
                self.krw_rate_label.pack_forget()
            if self.krw_rate_entry is not None:
                self.krw_rate_entry.pack_forget()
        else:
            self.field_labels["amount"].configure(text="Original amount")
            self.field_labels["rmb_amount"].configure(text="RMB amount")
            self._set_field_visible("currency", True)
            self._set_field_visible("krw_amount", True)
            self._set_field_visible("amount", True)
            self._set_field_visible("rmb_amount", True)
            if self.usa_rate_label is not None:
                self.usa_rate_label.pack_forget()
            if self.usa_rate_entry is not None:
                self.usa_rate_entry.pack_forget()
            if self.krw_rate_label is not None:
                self.krw_rate_label.pack(side="left")
            if self.krw_rate_entry is not None:
                self.krw_rate_entry.pack(side="left", padx=(6, 0))

    def _on_form_version_changed(self) -> None:
        version = self.form_version.get()
        labels = KOREA_CATEGORY_LABELS if version == "Korea" else USA_CATEGORY_LABELS
        self.category_values = [f"{key} - {label}" for key, label in labels.items()]
        cat_var = self.field_vars.get("category")
        if self.category_combo is not None:
            self.category_combo.configure(values=self.category_values)
        cat = self.field_vars.get("category")
        if cat is not None and not cat.get():
            cat.set(self.category_values[0] if self.category_values else "transportation")
        self._update_field_visibility()
        if self.selected_index is not None:
            self.load_selected_into_fields()
        self._sync_amount_fields("rate")

    def upload_receipts(self) -> None:
        if filedialog is None:
            return
        paths = filedialog.askopenfilenames(
            title="Upload receipt images",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.webp *.bmp *.gif"),
                ("All files", "*.*"),
            ],
        )
        if not paths:
            return
        self.save_current_fields()
        added = 0
        for raw_path in paths:
            path = Path(raw_path)
            if path.suffix.lower() not in SUPPORTED_IMAGE_SUFFIXES:
                continue
            self.items.append(
                ReceiptItem(
                    item_id=uuid.uuid4().hex,
                    path=str(path),
                    filename=path.name,
                    currency="USD" if self.form_version.get() == "USA" else "KRW",
                    category="transportation",
                )
            )
            added += 1
        self.refresh_tree()
        if self.items and self.selected_index is None:
            self.select_index(0)
        self.status_text.set(f"Added {added} receipt image(s).")

    def refresh_tree(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for index, item in enumerate(self.items):
            values = (item.status, item.date, item.amount or item.krw_amount or item.rmb_amount)
            self.tree.insert("", "end", iid=str(index), text=item.filename, values=values)
        if self.selected_index is not None and self.selected_index < len(self.items):
            self.tree.selection_set(str(self.selected_index))

    def _on_tree_select(self, _event: Any = None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        try:
            index = int(selection[0])
        except Exception:
            return
        self.select_index(index)

    def select_index(self, index: int) -> None:
        if self.selected_index == index:
            return
        self.save_current_fields()
        self.selected_index = index
        self.load_selected_into_fields()
        self.update_preview()

    def selected_item(self) -> Optional[ReceiptItem]:
        if self.selected_index is None:
            return None
        if self.selected_index < 0 or self.selected_index >= len(self.items):
            return None
        return self.items[self.selected_index]

    def save_current_fields(self) -> None:
        item = self.selected_item()
        if item is None:
            return
        for key, var in self.field_vars.items():
            value = var.get()
            if key == "category":
                value = value.split(" - ", 1)[0]
            setattr(item, key, value)

    def load_selected_into_fields(self) -> None:
        item = self.selected_item()
        if item is None:
            self._loading_fields = True
            try:
                for var in self.field_vars.values():
                    var.set("")
            finally:
                self._loading_fields = False
            return
        self._loading_fields = True
        try:
            for key, var in self.field_vars.items():
                value = getattr(item, key)
                if key == "category":
                    labels = KOREA_CATEGORY_LABELS if self.form_version.get() == "Korea" else USA_CATEGORY_LABELS
                    value = f"{value} - {labels.get(value, value)}"
                var.set(value)
        finally:
            self._loading_fields = False
        self._sync_amount_fields("rate")

    def update_preview(self) -> None:
        item = self.selected_item()
        if item is None:
            self.preview_label.configure(image="", text="Upload receipts to preview them here.")
            self.photo = None
            return
        path = Path(item.path)
        if not path.exists() or Image is None or ImageTk is None:
            self.preview_label.configure(image="", text=str(path))
            self.photo = None
            return
        try:
            image = Image.open(path)
            image.thumbnail((610, 610))
            self.photo = ImageTk.PhotoImage(image)
            self.preview_label.configure(image=self.photo, text="")
        except Exception as exc:
            log_exception("Receipt preview failed")
            self.preview_label.configure(image="", text=f"Could not preview image:\n{exc}")
            self.photo = None

    def remove_selected(self) -> None:
        if self.selected_index is None:
            return
        del self.items[self.selected_index]
        self.selected_index = None
        self.refresh_tree()
        if self.items:
            self.select_index(0)
        else:
            self.load_selected_into_fields()
            self.update_preview()
        self.status_text.set("Removed selected receipt.")

    def clear_all(self) -> None:
        if not self.items:
            return
        if messagebox and not messagebox.askyesno("Clear receipts", "Remove all uploaded receipt rows?"):
            return
        self.items.clear()
        self.selected_index = None
        self.refresh_tree()
        self.load_selected_into_fields()
        self.update_preview()
        self.status_text.set("Cleared receipt list.")

    def apply_ai_data_to_item(self, item: ReceiptItem, data: Dict[str, Any]) -> None:
        field_map = {
            "date": "date",
            "place": "place",
            "vendor": "place",
            "amount": "amount",
            "currency": "currency",
            "krw_amount": "krw_amount",
            "amount_krw": "krw_amount",
            "rmb_amount": "rmb_amount",
            "amount_rmb": "rmb_amount",
            "purpose": "purpose",
            "details": "details",
            "project_number": "project_number",
            "category": "category",
            "payment_method": "payment_method",
            "receipt_label": "receipt_label",
        }
        for src, dest in field_map.items():
            value = data.get(src)
            if value is None:
                continue
            text = str(value).strip()
            if not text:
                continue
            if dest == "category":
                text = normalized_category(text, self.form_version.get())
            setattr(item, dest, text)
        if self.form_version.get() == "USA":
            item.currency = "USD"
        item.status = "AI filled"

    def generate_selected_details(self) -> None:
        self.save_current_fields()
        item = self.selected_item()
        if item is None:
            self.show_info("Upload and select a receipt first.")
            return
        self._run_ai_for_items([item], reload_selected=True)

    def generate_all_details(self) -> None:
        self.save_current_fields()
        if not self.items:
            self.show_info("Upload receipts first.")
            return
        self._run_ai_for_items(list(self.items), reload_selected=True)

    def _run_ai_for_items(self, items: List[ReceiptItem], reload_selected: bool) -> None:
        def worker() -> None:
            try:
                for idx, item in enumerate(items, start=1):
                    self.set_status(f"Generating details {idx}/{len(items)}: {item.filename}")
                    data = call_openai_receipt_extraction(
                        Path(item.path),
                        self.form_version.get(),
                        item,
                        progress=self.set_status,
                    )
                    self.apply_ai_data_to_item(item, data)
                self.root.after(0, lambda: self.after_ai_success(reload_selected))
            except Exception as exc:
                log_exception("AI extraction failed")
                message = str(exc)
                self.root.after(0, lambda: self.after_ai_error(message))

        threading.Thread(target=worker, daemon=True).start()

    def after_ai_success(self, reload_selected: bool) -> None:
        self.refresh_tree()
        if reload_selected:
            self.load_selected_into_fields()
        self.status_text.set("AI details generated. Review before exporting.")

    def after_ai_error(self, message: str) -> None:
        self.status_text.set("AI extraction failed.")
        if messagebox:
            messagebox.showerror("Generate Details", f"{message}\n\nLogged to:\n{LOG_DIR / 'app.log'}")

    def set_status(self, message: str) -> None:
        self.root.after(0, lambda: self.status_text.set(message))

    def generate_excel(self) -> None:
        self.save_current_fields()
        if not self.items:
            self.show_info("Upload at least one receipt before generating Excel.")
            return
        if filedialog is None:
            return
        default_name = f"reimbursement_{self.form_version.get().lower()}_{date.today().strftime('%Y%m%d')}.xlsx"
        output = filedialog.asksaveasfilename(
            title="Save reimbursement workbook",
            initialdir=str(OUTPUT_DIR),
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if not output:
            return
        output_path = Path(output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            if self.form_version.get() == "USA":
                exchange_rate = safe_float(self.exchange_rate.get()) or safe_float(DEFAULT_USD_TO_RMB_RATE) or 6.8
                export_usa(self.items, output_path, exchange_rate)
            else:
                krw_rate = safe_float(self.krw_to_rmb_rate.get()) or safe_float(DEFAULT_KRW_TO_RMB_RATE) or 0.0046
                export_korea(self.items, output_path, krw_rate)
        except Exception as exc:
            log_path = log_exception("Excel export failed")
            if messagebox:
                messagebox.showerror(
                    "Generate Excel",
                    f"Could not generate workbook:\n{exc}\n\nLogged to:\n{log_path}",
                )
            self.status_text.set("Export failed.")
            return
        self.status_text.set(f"Saved: {output_path}")
        if messagebox:
            messagebox.showinfo("Generate Excel", f"Workbook saved:\n{output_path}")

    def open_output_folder(self) -> None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        if sys.platform == "win32":
            os.startfile(str(OUTPUT_DIR))  # type: ignore[attr-defined]
        else:
            self.show_info(f"Output folder: {OUTPUT_DIR}")

    def show_info(self, message: str) -> None:
        if messagebox:
            messagebox.showinfo("Reimbursement Helper", message)
        else:
            print(message)


def smoke_test_export() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    sample = ReceiptItem(
        item_id="sample",
        path=str(APP_DIR / "examples" / "sample_receipt_placeholder.png"),
        filename="sample_receipt_placeholder.png",
        date=date.today().strftime("%Y-%m-%d"),
        place="Sample Vendor",
        amount="12.34",
        currency="USD",
        krw_amount="18000",
        purpose="Parking",
        details="Sample parking receipt",
        project_number="",
        category="transportation",
        payment_method="",
        receipt_label="Sample receipt",
        status="Sample",
    )
    if Image is not None:
        sample_path = Path(sample.path)
        sample_path.parent.mkdir(parents=True, exist_ok=True)
        if not sample_path.exists():
            img = Image.new("RGB", (360, 220), "#FFFFFF")
            img.save(sample_path)
    export_usa([sample], OUTPUT_DIR / "smoke_usa.xlsx", safe_float(DEFAULT_USD_TO_RMB_RATE) or 6.8)
    export_korea([sample], OUTPUT_DIR / "smoke_korea.xlsx", safe_float(DEFAULT_KRW_TO_RMB_RATE) or 0.0046)
    print("Smoke exports created.")


def main(argv: Optional[List[str]] = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    try:
        setup_logging()
        ensure_dependencies()
        if "--smoke-test" in args:
            smoke_test_export()
            return 0
        CONFIG_DIR.mkdir(exist_ok=True)
        OUTPUT_DIR.mkdir(exist_ok=True)
        root = tk.Tk()
        ReimbursementHelperApp(root)
        root.mainloop()
        return 0
    except Exception as exc:
        log_path = log_exception("Application startup failed")
        details = "".join(traceback.format_exception_only(type(exc), exc)).strip()
        if messagebox and tk is not None:
            try:
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror("Reimbursement Helper", f"{details}\n\nLogged to:\n{log_path}")
                root.destroy()
            except Exception:
                print(details)
        else:
            print(details)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
