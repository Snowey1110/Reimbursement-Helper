from __future__ import annotations

import base64
import copy
import json
import logging
import math
import os
import re
import shutil
import subprocess
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
UNPROCESSED_DIR = APP_DIR / "Unprocessed"
PROCESSED_DIR = APP_DIR / "Processed"
USER_SETTINGS_FILE = CONFIG_DIR / "user_settings.json"
USER_API_KEY_FILE = CONFIG_DIR / "api_key.txt"
TEMPLATES_CONFIG_FILE = CONFIG_DIR / "templates.json"
SESSION_FILE = CONFIG_DIR / "session_state.json"
SUPPORTED_IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif"}
SUPPORTED_UPLOAD_SUFFIXES = SUPPORTED_IMAGE_SUFFIXES | {".pdf"}
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


def ensure_runtime_folders() -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    WORK_DIR.mkdir(parents=True, exist_ok=True)
    UNPROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)


def supported_image_files(folder: Path) -> List[Path]:
    if not folder.exists():
        return []
    return sorted(
        [
            path
            for path in folder.iterdir()
            if path.is_file() and path.suffix.lower() in SUPPORTED_IMAGE_SUFFIXES
        ],
        key=lambda path: path.name.lower(),
    )


def render_pdf_pages(pdf_path: Path) -> List[Path]:
    try:
        import fitz  # type: ignore[import-not-found]
    except Exception as exc:
        raise RuntimeError(
            "PDF support requires PyMuPDF. Run `python -m pip install -r requirements.txt`, then try again."
        ) from exc

    target_dir = WORK_DIR / "pdf_pages" / f"{pdf_path.stem}_{uuid.uuid4().hex[:8]}"
    target_dir.mkdir(parents=True, exist_ok=True)
    rendered: List[Path] = []
    try:
        document = fitz.open(str(pdf_path))
    except Exception as exc:
        raise RuntimeError(f"Could not open PDF: {pdf_path}") from exc
    try:
        for page_index in range(document.page_count):
            page = document.load_page(page_index)
            pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            output = target_dir / f"{pdf_path.stem}_page_{page_index + 1}.png"
            pixmap.save(str(output))
            rendered.append(output)
    finally:
        document.close()
    return rendered


def selected_file_key(path: Path, source_path: str = "", source_page: str = "") -> Tuple[str, str]:
    key_path = source_path or str(path)
    try:
        key_path = str(Path(key_path).resolve()).lower()
    except Exception:
        key_path = str(key_path).lower()
    return key_path, str(source_page or "")


def unique_path(folder: Path, filename: str) -> Path:
    candidate = folder / filename
    if not candidate.exists():
        return candidate
    stem = candidate.stem
    suffix = candidate.suffix
    for index in range(2, 10_000):
        candidate = folder / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
    return folder / f"{stem}_{uuid.uuid4().hex}{suffix}"


def same_folder(path: Path, folder: Path) -> bool:
    try:
        return path.resolve().parent == folder.resolve()
    except Exception:
        return False


def open_path(path: Path) -> None:
    if sys.platform == "win32":
        os.startfile(str(path))  # type: ignore[attr-defined]
    else:
        raise RuntimeError(f"Opening files is only implemented for Windows. Path: {path}")


def reveal_in_file_explorer(path: Path) -> None:
    if sys.platform == "win32":
        subprocess.Popen(["explorer", f"/select,{path}"])
    else:
        open_path(path.parent)


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
    source_path: str = ""
    source_page: str = ""
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
    crop_box: Optional[List[float]] = None


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


def save_user_settings(settings: Dict[str, Any]) -> None:
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        USER_SETTINGS_FILE.write_text(
            json.dumps(settings, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        log_exception("Could not save user settings")


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
- currency must match the receipt symbol/text: USD for "$", KRW for won/KRW,
  and RMB/CNY for yuan/RMB/CNY. Do not default to KRW just because the form is Korea.
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


def normalize_currency(value: Any, default: str = "USD") -> str:
    text = str(value or "").strip()
    if not text:
        return default
    upper = text.upper()
    if "$" in upper or "USD" in upper or "US DOLLAR" in upper or "DOLLAR" in upper:
        return "USD"
    if "KRW" in upper or "WON" in upper or "\u20a9" in upper or "\uffe6" in upper:
        return "KRW"
    if "RMB" in upper or "CNY" in upper or "YUAN" in upper or "CN\u00a5" in upper:
        return "RMB"
    if upper in {"USD", "KRW", "RMB", "CNY"}:
        return upper
    if upper in {"\u00a5", "\uffe5"}:
        return "RMB"
    return default


def category_value_to_key(value: str, form_version: str) -> str:
    return normalized_category((value or "").split(" - ", 1)[0], form_version)


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


def default_crop_points(width: int, height: int) -> List[Tuple[float, float]]:
    return [
        (0.0, 0.0),
        (float(width), 0.0),
        (float(width), float(height)),
        (0.0, float(height)),
    ]


def normalized_crop_points(crop_box: Any, width: int, height: int) -> Optional[List[Tuple[float, float]]]:
    if width <= 0 or height <= 0:
        return None
    if not isinstance(crop_box, (list, tuple)) or len(crop_box) != 4:
        if not isinstance(crop_box, (list, tuple)) or len(crop_box) != 8:
            return None
    try:
        values = [float(value) for value in crop_box]
    except (TypeError, ValueError):
        return None
    if len(values) == 4:
        x1, y1, x2, y2 = values
        points = [(x1, y1), (x2, y1), (x2, y2), (x1, y2)]
    else:
        points = [
            (values[0], values[1]),
            (values[2], values[3]),
            (values[4], values[5]),
            (values[6], values[7]),
        ]
    clamped = [
        (max(0.0, min(float(width), x)), max(0.0, min(float(height), y)))
        for x, y in points
    ]
    top_width = math.dist(clamped[0], clamped[1])
    bottom_width = math.dist(clamped[3], clamped[2])
    left_height = math.dist(clamped[0], clamped[3])
    right_height = math.dist(clamped[1], clamped[2])
    if max(top_width, bottom_width) < 5 or max(left_height, right_height) < 5:
        return None
    default = default_crop_points(width, height)
    if all(math.dist(point, default_point) < 1.0 for point, default_point in zip(clamped, default)):
        return None
    return clamped


def normalized_crop_box(crop_box: Any, width: int, height: int) -> Optional[Tuple[int, int, int, int]]:
    points = normalized_crop_points(crop_box, width, height)
    if not points:
        return None
    xs = [point[0] for point in points]
    ys = [point[1] for point in points]
    return (
        int(round(max(0.0, min(xs)))),
        int(round(max(0.0, min(ys)))),
        int(round(min(float(width), max(xs)))),
        int(round(min(float(height), max(ys)))),
    )


def solve_linear_system(matrix: List[List[float]], vector: List[float]) -> List[float]:
    size = len(vector)
    rows = [matrix[index][:] + [vector[index]] for index in range(size)]
    for pivot_index in range(size):
        pivot_row = max(range(pivot_index, size), key=lambda row: abs(rows[row][pivot_index]))
        if abs(rows[pivot_row][pivot_index]) < 1e-9:
            raise ValueError("Perspective crop points are too close together.")
        rows[pivot_index], rows[pivot_row] = rows[pivot_row], rows[pivot_index]
        pivot = rows[pivot_index][pivot_index]
        rows[pivot_index] = [value / pivot for value in rows[pivot_index]]
        for row_index in range(size):
            if row_index == pivot_index:
                continue
            factor = rows[row_index][pivot_index]
            rows[row_index] = [
                value - factor * rows[pivot_index][col_index]
                for col_index, value in enumerate(rows[row_index])
            ]
    return [rows[index][-1] for index in range(size)]


def perspective_coefficients(
    source_points: List[Tuple[float, float]],
    width: int,
    height: int,
) -> List[float]:
    destination_points = [
        (0.0, 0.0),
        (float(width), 0.0),
        (float(width), float(height)),
        (0.0, float(height)),
    ]
    matrix: List[List[float]] = []
    vector: List[float] = []
    for (u, v), (x, y) in zip(destination_points, source_points):
        matrix.append([u, v, 1.0, 0.0, 0.0, 0.0, -u * x, -v * x])
        vector.append(x)
        matrix.append([0.0, 0.0, 0.0, u, v, 1.0, -u * y, -v * y])
        vector.append(y)
    return solve_linear_system(matrix, vector)


def perspective_crop_image(image: Any, crop_box: Any) -> Any:
    points = normalized_crop_points(crop_box, image.width, image.height)
    if not points:
        return image
    top_width = math.dist(points[0], points[1])
    bottom_width = math.dist(points[3], points[2])
    left_height = math.dist(points[0], points[3])
    right_height = math.dist(points[1], points[2])
    output_width = max(1, int(round(max(top_width, bottom_width))))
    output_height = max(1, int(round(max(left_height, right_height))))
    try:
        coefficients = perspective_coefficients(points, output_width, output_height)
    except Exception:
        setup_logging()
        LOGGER.exception("Could not solve perspective crop; using original image")
        return image
    transform_mode = getattr(Image, "Transform", Image).PERSPECTIVE
    resample = getattr(getattr(Image, "Resampling", Image), "BICUBIC", Image.BICUBIC)
    return image.transform((output_width, output_height), transform_mode, coefficients, resample)


def prepare_excel_image_file(path: Path, crop_box: Any = None) -> Path:
    if Image is None:
        return path
    target_dir = WORK_DIR / "excel_images"
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / f"{uuid.uuid4().hex}.png"
    with Image.open(path) as source:
        image = ImageOps.exif_transpose(source) if ImageOps is not None else source.copy()
        image = perspective_crop_image(image, crop_box)
        if image.mode not in {"RGB", "RGBA"}:
            image = image.convert("RGB")
        image.save(target, "PNG")
    return target


def resize_excel_image(path: Path, max_width: int, max_height: int, crop_box: Any = None) -> Any:
    excel_path = prepare_excel_image_file(path, crop_box)
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


def configure_korea_receipt_page(sheet: Any, item_count: int) -> None:
    page_height = 50
    receipts_per_page = 4
    page_count = max(1, (item_count + receipts_per_page - 1) // receipts_per_page)
    last_row = page_count * page_height
    sheet.print_area = f"A1:E{last_row}"
    if sheet.sheet_properties.pageSetUpPr is None:
        try:
            from openpyxl.worksheet.properties import PageSetupProperties

            sheet.sheet_properties.pageSetUpPr = PageSetupProperties()
        except Exception:
            pass
    if sheet.sheet_properties.pageSetUpPr is not None:
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "portrait"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.page_margins.left = 0.45
    sheet.page_margins.right = 0.45
    sheet.page_margins.top = 0.45
    sheet.page_margins.bottom = 0.45
    try:
        from openpyxl.worksheet.pagebreak import Break

        sheet.row_breaks.brk = []
        for row in range(page_height, last_row, page_height):
            sheet.row_breaks.append(Break(id=row))
    except Exception:
        pass


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
            receipts_ws.add_image(resize_excel_image(image_path, 260, 150, item.crop_box), f"D{row}")

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


def korea_amounts(
    item: ReceiptItem,
    krw_to_rmb_rate: float,
    usd_to_rmb_rate: float,
) -> Tuple[Optional[float], Optional[float], str]:
    amount = safe_float(item.amount)
    krw = safe_float(item.krw_amount)
    rmb = safe_float(item.rmb_amount)
    currency = normalize_currency(item.currency, "KRW")
    if amount is not None:
        if currency == "USD":
            rmb = amount * usd_to_rmb_rate if usd_to_rmb_rate else rmb
            krw = rmb / krw_to_rmb_rate if rmb is not None and krw_to_rmb_rate else krw
        elif currency == "RMB" or currency == "CNY":
            rmb = amount
            krw = amount / krw_to_rmb_rate if krw_to_rmb_rate else None
        elif currency == "KRW":
            krw = amount
            rmb = amount * krw_to_rmb_rate if krw_to_rmb_rate else None
    if krw is None and rmb is not None and krw_to_rmb_rate:
        krw = rmb / krw_to_rmb_rate
    if rmb is None and krw is not None and krw_to_rmb_rate:
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


def export_korea(
    items: List[ReceiptItem],
    output_path: Path,
    krw_to_rmb_rate: float,
    usd_to_rmb_rate: float,
) -> None:
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
        krw, rmb, original_note = korea_amounts(item, krw_to_rmb_rate, usd_to_rmb_rate)
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
    configure_korea_receipt_page(receipts_ws, len(items))
    page_height = 50
    receipts_per_page = 4
    page_count = max(1, (len(items) + receipts_per_page - 1) // receipts_per_page)
    last_receipt_row = page_count * page_height
    for row in range(1, max(240, last_receipt_row + 1)):
        for col in "ABCDEFGH":
            receipts_ws[f"{col}{row}"] = None
    anchors = ["A", "D"]
    row_offsets = [2, 26]
    for idx, item in enumerate(items):
        page = idx // receipts_per_page
        slot = idx % receipts_per_page
        row_group = slot // len(anchors)
        col = anchors[slot % len(anchors)]
        image_row = (page * page_height) + row_offsets[row_group]
        image_path = Path(item.path)
        if image_path.exists():
            receipts_ws.add_image(resize_excel_image(image_path, 215, 300, item.crop_box), f"{col}{image_row}")

    wb.save(output_path)


class ReimbursementHelperApp:
    def __init__(self, root: Any) -> None:
        self.root = root
        self.items: List[ReceiptItem] = []
        self.selected_index: Optional[int] = None
        self.photo: Optional[Any] = None
        self.preview_canvas: Optional[Any] = None
        self.revert_crop_btn: Optional[Any] = None
        self._preview_image_bounds: Tuple[int, int, int, int] = (0, 0, 0, 0)
        self._preview_original_size: Tuple[int, int] = (0, 0)
        self._crop_handle_centers: Dict[str, Tuple[float, float]] = {}
        self._dragging_crop_handle: Optional[str] = None
        self.settings = load_user_settings()
        self.form_version = tk.StringVar(value="USA")
        self.exchange_rate = tk.StringVar(value=str(self.settings.get("usa_exchange_rate", DEFAULT_USD_TO_RMB_RATE)))
        self.krw_to_rmb_rate = tk.StringVar(value=str(self.settings.get("krw_to_rmb_rate", DEFAULT_KRW_TO_RMB_RATE)))
        self.status_text = tk.StringVar(value="Ready")
        self.progress_text = tk.StringVar(value="")
        self.progress_value = tk.DoubleVar(value=0)
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
        self._last_form_version = self.form_version.get()
        self._busy = False
        self._build_ui()
        self._attach_amount_traces()
        self._on_form_version_changed()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(250, self.restore_previous_session_if_available)

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
        self.upload_folder_btn = ttk.Button(controls, text="Select Files", command=self.select_files)
        self.upload_folder_btn.pack(side="left", padx=4)
        self.generate_details_btn = ttk.Button(controls, text="Generate Details", command=self.generate_selected_details)
        self.generate_details_btn.pack(side="left", padx=4)
        self.generate_all_btn = ttk.Button(controls, text="Generate All", command=self.generate_all_details)
        self.generate_all_btn.pack(side="left", padx=4)
        self.generate_excel_btn = ttk.Button(controls, text="Generate Excel", command=self.generate_excel)
        self.generate_excel_btn.pack(side="left", padx=4)

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
        self.tree = ttk.Treeview(
            left,
            columns=("status", "date", "amount"),
            show="tree headings",
            height=20,
            selectmode="extended",
        )
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
        preview_header = ttk.Frame(right, style="Panel.TFrame")
        preview_header.grid(row=0, column=0, sticky="ew")
        preview_header.grid_columnconfigure(0, weight=1)
        ttk.Label(preview_header, text="Receipt preview", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        self.revert_crop_btn = ttk.Button(preview_header, text="Revert crop", command=self.revert_crop, state="disabled")
        self.revert_crop_btn.grid(row=0, column=1, sticky="e")
        preview_frame = ttk.Frame(right, style="Panel.TFrame")
        preview_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_canvas = tk.Canvas(
            preview_frame,
            bg="#FFFFFF",
            highlightthickness=1,
            highlightbackground="#222222",
            bd=1,
            relief="solid",
        )
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")
        self.preview_canvas.bind("<Configure>", lambda _event: self.update_preview())
        self.preview_canvas.bind("<ButtonPress-1>", self._on_crop_press)
        self.preview_canvas.bind("<B1-Motion>", self._on_crop_drag)
        self.preview_canvas.bind("<ButtonRelease-1>", self._on_crop_release)

        footer = ttk.Frame(self.root, padding=(18, 0, 18, 12))
        footer.pack(fill="x")
        ttk.Label(footer, textvariable=self.status_text, style="Muted.TLabel").pack(side="left")
        ttk.Label(footer, textvariable=self.progress_text, style="Muted.TLabel").pack(side="left", padx=(16, 6))
        self.progress_bar = ttk.Progressbar(
            footer,
            variable=self.progress_value,
            maximum=100,
            length=180,
            mode="determinate",
        )
        self.progress_bar.pack(side="left")
        ttk.Button(footer, text="Open output folder", command=self.open_output_folder).pack(side="right")

    def _attach_amount_traces(self) -> None:
        for key in ("amount", "currency", "krw_amount", "rmb_amount"):
            var = self.field_vars.get(key)
            if var is not None:
                var.trace_add("write", lambda *_args, k=key: self._sync_amount_fields(k))
        for key in (
            "date",
            "place",
            "purpose",
            "details",
            "project_number",
            "category",
            "payment_method",
            "receipt_label",
        ):
            var = self.field_vars.get(key)
            if var is not None:
                var.trace_add("write", lambda *_args, k=key: self._on_field_changed(k))
        self.exchange_rate.trace_add("write", lambda *_args: self._sync_amount_fields("rate"))
        self.krw_to_rmb_rate.trace_add("write", lambda *_args: self._sync_amount_fields("rate"))

    def _on_field_changed(self, key: str) -> None:
        if self._loading_fields or self._syncing_amounts:
            return
        self._save_field_to_selected(key, refresh=True)

    def _sync_amount_fields(self, changed_key: str) -> None:
        if self._loading_fields or self._syncing_amounts:
            return
        self._syncing_amounts = True
        should_save = False
        try:
            version = self.form_version.get()
            amount_var = self.field_vars.get("amount")
            currency_var = self.field_vars.get("currency")
            krw_var = self.field_vars.get("krw_amount")
            rmb_var = self.field_vars.get("rmb_amount")
            if amount_var is None or currency_var is None or krw_var is None or rmb_var is None:
                should_save = False
            if changed_key in {"amount", "currency", "krw_amount", "rmb_amount"}:
                self._last_amount_source = changed_key

            if amount_var is None or currency_var is None or krw_var is None or rmb_var is None:
                return

            if version == "USA":
                rate = safe_float(self.exchange_rate.get())
                if rate:
                    if currency_var.get() != "USD":
                        currency_var.set("USD")
                    source = self._last_amount_source if changed_key == "rate" else changed_key
                    if source == "rmb_amount":
                        rmb = safe_float(rmb_var.get())
                        if rmb is not None:
                            amount_var.set(format_amount(rmb / rate))
                    else:
                        usd = safe_float(amount_var.get())
                        if usd is not None:
                            rmb_var.set(format_amount(usd * rate))
                    should_save = True
            else:
                krw_rate = safe_float(self.krw_to_rmb_rate.get())
                usd_rate = safe_float(self.exchange_rate.get())
                if krw_rate:
                    currency = normalize_currency(currency_var.get(), "KRW")
                    if currency_var.get() != currency:
                        currency_var.set(currency)
                    source = self._last_amount_source if changed_key == "rate" else changed_key
                    amount = safe_float(amount_var.get())
                    krw = safe_float(krw_var.get())
                    rmb = safe_float(rmb_var.get())
                    if source in {"amount", "currency", "rate"} and amount is not None:
                        if currency == "USD":
                            if usd_rate:
                                rmb = amount * usd_rate
                                krw = rmb / krw_rate
                        elif currency in {"RMB", "CNY"}:
                            rmb = amount
                            krw = amount / krw_rate
                        else:
                            krw = amount
                            rmb = amount * krw_rate
                    elif source == "rmb_amount" and rmb is not None:
                        krw = rmb / krw_rate
                        if currency in {"RMB", "CNY"}:
                            amount_var.set(format_amount(rmb))
                        elif currency == "KRW":
                            amount_var.set(format_amount(krw))
                        elif currency == "USD" and usd_rate and not amount_var.get().strip():
                            amount_var.set(format_amount(rmb / usd_rate))
                    elif source == "krw_amount" and krw is not None:
                        rmb = krw * krw_rate
                        if currency == "KRW":
                            amount_var.set(format_amount(krw))
                        elif currency in {"RMB", "CNY"} and not amount_var.get().strip():
                            amount_var.set(format_amount(rmb))
                        elif currency == "USD" and usd_rate and not amount_var.get().strip():
                            amount_var.set(format_amount(rmb / usd_rate))

                    if krw is None and rmb is not None:
                        krw = rmb / krw_rate
                    if rmb is None and krw is not None:
                        rmb = krw * krw_rate
                    if krw is not None:
                        krw_var.set(format_amount(krw))
                    if rmb is not None:
                        rmb_var.set(format_amount(rmb))
                    should_save = True
        finally:
            self._syncing_amounts = False
        if should_save:
            self._save_amount_fields_to_selected(changed_key, refresh=True)

    def update_item_amount_fields(self, item: ReceiptItem, source: str = "amount") -> None:
        version = self.form_version.get()
        if version == "USA":
            item.currency = "USD"
            rate = safe_float(self.exchange_rate.get()) or safe_float(DEFAULT_USD_TO_RMB_RATE)
            if not rate:
                return
            amount = safe_float(item.amount)
            rmb = safe_float(item.rmb_amount)
            if source == "rmb_amount" and rmb is not None:
                item.amount = format_amount(rmb / rate)
            elif amount is not None:
                item.rmb_amount = format_amount(amount * rate)
            return

        krw_rate = safe_float(self.krw_to_rmb_rate.get()) or safe_float(DEFAULT_KRW_TO_RMB_RATE)
        usd_rate = safe_float(self.exchange_rate.get()) or safe_float(DEFAULT_USD_TO_RMB_RATE)
        if not krw_rate:
            return
        item.currency = normalize_currency(item.currency, "KRW")
        amount = safe_float(item.amount)
        krw = safe_float(item.krw_amount)
        rmb = safe_float(item.rmb_amount)
        item_source = "amount" if source in {"currency", "rate"} and amount is not None else source
        if item_source == "amount" and amount is not None:
            if item.currency == "USD":
                rmb = amount * usd_rate if usd_rate else rmb
                krw = rmb / krw_rate if rmb is not None else krw
            elif item.currency in {"RMB", "CNY"}:
                rmb = amount
                krw = amount / krw_rate
            else:
                krw = amount
                rmb = amount * krw_rate
        elif item_source == "rmb_amount" and rmb is not None:
            krw = rmb / krw_rate
            if item.currency in {"RMB", "CNY"}:
                item.amount = format_amount(rmb)
            elif item.currency == "KRW" and amount is None:
                item.amount = format_amount(krw)
            elif item.currency == "USD" and amount is None and usd_rate:
                item.amount = format_amount(rmb / usd_rate)
        elif item_source == "krw_amount" and krw is not None:
            rmb = krw * krw_rate
            if item.currency == "KRW":
                item.amount = format_amount(krw)
            elif item.currency in {"RMB", "CNY"} and amount is None:
                item.amount = format_amount(rmb)
            elif item.currency == "USD" and amount is None and usd_rate:
                item.amount = format_amount(rmb / usd_rate)

        if krw is None and rmb is not None:
            krw = rmb / krw_rate
        if rmb is None and krw is not None:
            rmb = krw * krw_rate
        if krw is not None:
            item.krw_amount = format_amount(krw)
        if rmb is not None:
            item.rmb_amount = format_amount(rmb)

    def _save_field_to_selected(self, key: str, refresh: bool = False) -> None:
        var = self.field_vars.get(key)
        if var is None:
            return
        value = var.get()
        version = self.form_version.get()
        for index in self.selected_indices():
            item = self.items[index]
            if key == "category":
                setattr(item, key, category_value_to_key(value, version))
            else:
                setattr(item, key, value)
        if refresh:
            self.refresh_tree()

    def _save_amount_fields_to_selected(self, source: str, refresh: bool = False) -> None:
        version = self.form_version.get()
        amount_value = self.field_vars["amount"].get()
        currency_value = normalize_currency(self.field_vars["currency"].get(), "USD" if version == "USA" else "KRW")
        krw_value = self.field_vars["krw_amount"].get()
        rmb_value = self.field_vars["rmb_amount"].get()
        indices = range(len(self.items)) if source == "rate" else self.selected_indices()
        for index in indices:
            item = self.items[index]
            if source == "rate":
                self.update_item_amount_fields(item, "amount")
                continue
            if source == "amount":
                item.amount = amount_value
                if version == "USA":
                    item.currency = "USD"
                self.update_item_amount_fields(item, "amount")
            elif source == "currency":
                item.currency = currency_value
                self.update_item_amount_fields(item, "currency")
            elif source == "krw_amount":
                item.krw_amount = krw_value
                self.update_item_amount_fields(item, "krw_amount")
            elif source == "rmb_amount":
                item.rmb_amount = rmb_value
                self.update_item_amount_fields(item, "rmb_amount")
        if refresh:
            self.refresh_tree()

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
        form_changed = version != self._last_form_version
        labels = KOREA_CATEGORY_LABELS if version == "Korea" else USA_CATEGORY_LABELS
        self.category_values = [f"{key} - {label}" for key, label in labels.items()]
        cat_var = self.field_vars.get("category")
        if self.category_combo is not None:
            self.category_combo.configure(values=self.category_values)
        cat = self.field_vars.get("category")
        if cat is not None and not cat.get():
            cat.set(self.category_values[0] if self.category_values else "transportation")
        for item in self.items:
            if version == "USA":
                item.currency = "USD"
            else:
                item.currency = normalize_currency(item.currency, "KRW")
            self.update_item_amount_fields(item, "amount")
            if form_changed and item.status == "AI filled":
                item.status = "Needs regen"
        self._update_field_visibility()
        if self.selected_index is not None:
            self.load_selected_into_fields()
        self._sync_amount_fields("rate")
        self.refresh_tree()
        self._last_form_version = version

    def select_files(self) -> None:
        ensure_runtime_folders()
        if filedialog is None:
            return
        selected = filedialog.askopenfilenames(
            title="Select receipt files",
            filetypes=[
                ("Receipt files", "*.png *.jpg *.jpeg *.webp *.bmp *.gif *.pdf"),
                ("Image files", "*.png *.jpg *.jpeg *.webp *.bmp *.gif"),
                ("PDF files", "*.pdf"),
                ("All files", "*.*"),
            ],
        )
        paths = [Path(path) for path in selected]
        if not paths:
            return
        self.save_current_fields()
        existing = {
            selected_file_key(Path(item.path), item.source_path, item.source_page)
            for item in self.items
            if item.path
        }
        added = 0
        failed: List[str] = []
        first_new_index = len(self.items)
        for path in paths:
            suffix = path.suffix.lower()
            if suffix not in SUPPORTED_UPLOAD_SUFFIXES:
                continue
            if suffix == ".pdf":
                try:
                    page_paths = render_pdf_pages(path)
                except Exception as exc:
                    log_exception(f"Could not import PDF: {path}")
                    failed.append(f"{path.name}: {exc}")
                    continue
                for page_index, page_path in enumerate(page_paths, start=1):
                    key = selected_file_key(page_path, str(path), str(page_index))
                    if key in existing:
                        continue
                    self.items.append(
                        ReceiptItem(
                            item_id=uuid.uuid4().hex,
                            path=str(page_path),
                            filename=f"{path.name} page {page_index}",
                            source_path=str(path),
                            source_page=str(page_index),
                            currency="USD" if self.form_version.get() == "USA" else "KRW",
                            category="transportation",
                        )
                    )
                    existing.add(key)
                    added += 1
                continue
            key = selected_file_key(path, str(path), "")
            if key in existing:
                continue
            self.items.append(
                ReceiptItem(
                    item_id=uuid.uuid4().hex,
                    path=str(path),
                    filename=path.name,
                    source_path=str(path),
                    currency="USD" if self.form_version.get() == "USA" else "KRW",
                    category="transportation",
                )
            )
            existing.add(key)
            added += 1
        self.refresh_tree()
        if added and self.items and self.selected_index is None:
            self.select_index(first_new_index)
        self.save_session()
        self.status_text.set(f"Added {added} receipt file(s).")
        if failed and messagebox:
            messagebox.showwarning(
                "Select Files",
                "Some files could not be imported.\n\n"
                + "\n".join(failed[:5])
                + f"\n\nLogged to:\n{LOG_DIR / 'app.log'}",
            )

    def upload_folder(self) -> None:
        self.select_files()

    def refresh_tree(self) -> None:
        current_selection = set(self.tree.selection())
        self.tree.delete(*self.tree.get_children())
        for index, item in enumerate(self.items):
            values = (item.status, item.date, item.amount or item.krw_amount or item.rmb_amount)
            self.tree.insert("", "end", iid=str(index), text=item.filename, values=values)
        valid_selection = [iid for iid in current_selection if iid.isdigit() and int(iid) < len(self.items)]
        if valid_selection:
            self.tree.selection_set(*valid_selection)
        elif self.selected_index is not None and self.selected_index < len(self.items):
            self.tree.selection_set(str(self.selected_index))
        if self.selected_index is not None and self.selected_index < len(self.items):
            self.tree.focus(str(self.selected_index))

    def _on_tree_select(self, _event: Any = None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        focus = self.tree.focus()
        chosen = focus if focus in selection else selection[-1]
        try:
            index = int(chosen)
        except Exception:
            return
        self.select_index(index)

    def select_index(self, index: int) -> None:
        if self.selected_index == index:
            return
        self.save_current_fields()
        self.selected_index = index
        self.tree.focus(str(index))
        if str(index) not in self.tree.selection():
            self.tree.selection_set(str(index))
        self.load_selected_into_fields()
        self.update_preview()

    def selected_indices(self) -> List[int]:
        indices: List[int] = []
        for iid in self.tree.selection():
            try:
                index = int(iid)
            except Exception:
                continue
            if 0 <= index < len(self.items):
                indices.append(index)
        if not indices and self.selected_index is not None and 0 <= self.selected_index < len(self.items):
            indices.append(self.selected_index)
        return sorted(set(indices))

    def selected_items(self) -> List[ReceiptItem]:
        return [self.items[index] for index in self.selected_indices()]

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
                value = category_value_to_key(value, self.form_version.get())
            elif key == "currency":
                value = normalize_currency(value, "USD" if self.form_version.get() == "USA" else "KRW")
            setattr(item, key, value)
        self.update_item_amount_fields(item, self._last_amount_source)

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

    def clear_preview(self, message: str) -> None:
        if self.preview_canvas is None:
            return
        self.preview_canvas.delete("all")
        width = max(1, self.preview_canvas.winfo_width())
        height = max(1, self.preview_canvas.winfo_height())
        self.preview_canvas.create_text(
            width / 2,
            height / 2,
            text=message,
            fill="#666666",
            width=max(160, width - 30),
            justify="center",
        )
        self.photo = None
        self._preview_image_bounds = (0, 0, 0, 0)
        self._preview_original_size = (0, 0)
        self._crop_handle_centers = {}
        if self.revert_crop_btn is not None:
            self.revert_crop_btn.configure(state="disabled")

    def update_preview(self) -> None:
        if self.preview_canvas is None:
            return
        item = self.selected_item()
        if item is None:
            self.clear_preview("Select receipt image or PDF files to begin.")
            return
        path = Path(item.path)
        if not path.exists() or Image is None or ImageTk is None:
            self.clear_preview(str(path))
            return
        try:
            with Image.open(path) as source:
                image = ImageOps.exif_transpose(source) if ImageOps is not None else source.copy()
            self._preview_original_size = (image.width, image.height)
            canvas_w = max(240, self.preview_canvas.winfo_width())
            canvas_h = max(240, self.preview_canvas.winfo_height())
            max_w = max(1, canvas_w - 20)
            max_h = max(1, canvas_h - 20)
            image.thumbnail((max_w, max_h))
            display_w, display_h = image.size
            x = max(0, (canvas_w - display_w) // 2)
            y = max(0, (canvas_h - display_h) // 2)
            self.photo = ImageTk.PhotoImage(image)
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(x, y, anchor="nw", image=self.photo)
            self._preview_image_bounds = (x, y, display_w, display_h)
            self.draw_crop_overlay()
        except Exception as exc:
            log_exception("Receipt preview failed")
            self.clear_preview(f"Could not preview image:\n{exc}")

    def current_preview_crop_points(self, item: ReceiptItem) -> List[Tuple[float, float]]:
        width, height = self._preview_original_size
        crop = normalized_crop_points(item.crop_box, width, height)
        if crop:
            return crop
        return default_crop_points(width, height)

    def original_to_canvas(self, x: float, y: float) -> Tuple[float, float]:
        image_x, image_y, display_w, display_h = self._preview_image_bounds
        original_w, original_h = self._preview_original_size
        if original_w <= 0 or original_h <= 0:
            return float(image_x), float(image_y)
        return image_x + (x / original_w) * display_w, image_y + (y / original_h) * display_h

    def canvas_to_original(self, x: float, y: float) -> Tuple[float, float]:
        image_x, image_y, display_w, display_h = self._preview_image_bounds
        original_w, original_h = self._preview_original_size
        if display_w <= 0 or display_h <= 0:
            return 0.0, 0.0
        clamped_x = max(image_x, min(image_x + display_w, x))
        clamped_y = max(image_y, min(image_y + display_h, y))
        original_x = ((clamped_x - image_x) / display_w) * original_w
        original_y = ((clamped_y - image_y) / display_h) * original_h
        return original_x, original_y

    def draw_crop_overlay(self) -> None:
        if self.preview_canvas is None:
            return
        item = self.selected_item()
        original_w, original_h = self._preview_original_size
        image_x, image_y, display_w, display_h = self._preview_image_bounds
        if item is None or original_w <= 0 or original_h <= 0 or display_w <= 0 or display_h <= 0:
            return
        self.preview_canvas.delete("crop")
        crop = self.current_preview_crop_points(item)
        canvas_points = [self.original_to_canvas(x, y) for x, y in crop]
        flattened_points = [coord for point in canvas_points for coord in point]
        self.preview_canvas.create_polygon(
            *flattened_points,
            outline="#0078D7",
            fill="",
            width=2,
            tags=("crop",),
        )
        handle_size = 8
        centers = {
            "nw": canvas_points[0],
            "ne": canvas_points[1],
            "se": canvas_points[2],
            "sw": canvas_points[3],
        }
        self._crop_handle_centers = centers
        for handle, (hx, hy) in centers.items():
            self.preview_canvas.create_oval(
                hx - handle_size,
                hy - handle_size,
                hx + handle_size,
                hy + handle_size,
                fill="#FFFFFF",
                outline="#0078D7",
                width=2,
                tags=("crop", f"crop_{handle}"),
            )
        if self.revert_crop_btn is not None:
            has_crop = normalized_crop_points(item.crop_box, original_w, original_h) is not None
            self.revert_crop_btn.configure(state="normal" if has_crop else "disabled")

    def nearest_crop_handle(self, x: float, y: float) -> Optional[str]:
        nearest: Optional[str] = None
        nearest_distance = 16.0
        for handle, (hx, hy) in self._crop_handle_centers.items():
            distance = ((hx - x) ** 2 + (hy - y) ** 2) ** 0.5
            if distance <= nearest_distance:
                nearest = handle
                nearest_distance = distance
        return nearest

    def _on_crop_press(self, event: Any) -> None:
        self._dragging_crop_handle = self.nearest_crop_handle(event.x, event.y)

    def _on_crop_drag(self, event: Any) -> None:
        item = self.selected_item()
        handle = self._dragging_crop_handle
        original_w, original_h = self._preview_original_size
        if item is None or handle is None or original_w <= 0 or original_h <= 0:
            return
        points = self.current_preview_crop_points(item)
        x, y = self.canvas_to_original(event.x, event.y)
        handle_indexes = {"nw": 0, "ne": 1, "se": 2, "sw": 3}
        points[handle_indexes[handle]] = (x, y)
        item.crop_box = [
            round(value, 2)
            for point in points
            for value in point
        ]
        self.draw_crop_overlay()

    def _on_crop_release(self, _event: Any) -> None:
        if self._dragging_crop_handle is None:
            return
        self._dragging_crop_handle = None
        self.save_session()
        self.status_text.set("Crop updated for selected receipt.")

    def revert_crop(self) -> None:
        item = self.selected_item()
        if item is None:
            return
        item.crop_box = None
        self.draw_crop_overlay()
        self.save_session()
        self.status_text.set("Crop reverted for selected receipt.")

    def remove_selected(self) -> None:
        indices = self.selected_indices()
        if not indices:
            return
        for index in sorted(indices, reverse=True):
            del self.items[index]
        self.selected_index = None
        self.refresh_tree()
        if self.items:
            self.select_index(min(indices[0], len(self.items) - 1))
        else:
            self.load_selected_into_fields()
            self.update_preview()
        self.status_text.set(f"Removed {len(indices)} selected receipt(s).")
        self.save_session()

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
        self.save_session()

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
            elif dest == "currency":
                text = normalize_currency(text, "USD" if self.form_version.get() == "USA" else item.currency or "KRW")
            elif dest == "amount":
                hinted_currency = normalize_currency(text, "")
                if hinted_currency:
                    item.currency = hinted_currency
            setattr(item, dest, text)
        if self.form_version.get() == "USA":
            item.currency = "USD"
        else:
            item.currency = normalize_currency(item.currency, "KRW")
        self.update_item_amount_fields(item, "amount")
        item.status = "AI filled"

    def generate_selected_details(self) -> None:
        self.save_current_fields()
        items = self.selected_items()
        if not items:
            self.show_info("Upload and select a receipt first.")
            return
        self._run_ai_for_items(items, reload_selected=True)

    def generate_all_details(self) -> None:
        self.save_current_fields()
        if not self.items:
            self.show_info("Upload receipts first.")
            return
        self._run_ai_for_items(list(self.items), reload_selected=True)

    def _run_ai_for_items(self, items: List[ReceiptItem], reload_selected: bool) -> None:
        if self._busy:
            self.show_info("A batch is already running.")
            return
        total = len(items)
        if total <= 0:
            return
        form_version = self.form_version.get()
        self.set_busy(True)
        self.set_progress(0, total, f"0/{total}")
        self.status_text.set(f"Generating details for {total} receipt(s)...")

        def worker() -> None:
            successes = 0
            failures: List[str] = []
            for idx, item in enumerate(items, start=1):
                self.root.after(
                    0,
                    lambda i=item, n=idx: self.mark_item_status(i, f"Processing {n}/{total}", n - 1, total),
                )
                self.set_status(f"Generating details {idx}/{total}: {item.filename}")
                try:
                    data = call_openai_receipt_extraction(
                        Path(item.path),
                        form_version,
                        item,
                        progress=self.set_status,
                    )
                    self.apply_ai_data_to_item(item, data)
                    self.move_item_to_processed(item)
                    successes += 1
                    self.root.after(
                        0,
                        lambda i=item, n=idx: self.mark_item_status(i, "AI filled", n, total),
                    )
                except Exception as exc:
                    setup_logging()
                    LOGGER.exception("AI extraction failed for %s", item.filename)
                    failures.append(f"{item.filename}: {exc}")
                    self.root.after(
                        0,
                        lambda i=item, n=idx: self.mark_item_status(i, "Failed", n, total),
                    )
                    continue
            self.root.after(0, lambda: self.after_ai_complete(successes, failures, total, reload_selected))

        threading.Thread(target=worker, daemon=True).start()

    def after_ai_complete(
        self,
        successes: int,
        failures: List[str],
        total: int,
        reload_selected: bool,
    ) -> None:
        self.set_busy(False)
        self.set_progress(total, total, f"{total}/{total}")
        self.refresh_tree()
        if reload_selected:
            self.load_selected_into_fields()
            self.update_preview()
        self.save_session()
        if failures:
            self.status_text.set(f"AI finished: {successes} succeeded, {len(failures)} failed.")
            if messagebox:
                shown = "\n".join(failures[:5])
                if len(failures) > 5:
                    shown += f"\n...and {len(failures) - 5} more."
                messagebox.showwarning(
                    "Generate Details",
                    f"Finished with failures.\n\n{shown}\n\nLogged to:\n{LOG_DIR / 'app.log'}",
                )
        else:
            self.status_text.set("AI details generated. Review before exporting.")

    def set_status(self, message: str) -> None:
        self.root.after(0, lambda: self.status_text.set(message))

    def set_busy(self, busy: bool) -> None:
        self._busy = busy
        state = "disabled" if busy else "normal"
        for button in (
            self.upload_folder_btn,
            self.generate_details_btn,
            self.generate_all_btn,
            self.generate_excel_btn,
        ):
            try:
                button.configure(state=state)
            except Exception:
                pass

    def set_progress(self, done: int, total: int, text: str = "") -> None:
        if total <= 0:
            self.progress_value.set(0)
            self.progress_text.set(text)
            return
        pct = max(0.0, min(100.0, (done / total) * 100.0))
        self.progress_value.set(pct)
        self.progress_text.set(text or f"{done}/{total}")

    def mark_item_status(self, item: ReceiptItem, status: str, done: int, total: int) -> None:
        item.status = status
        self.refresh_tree()
        self.set_progress(done, total, f"{done}/{total}")
        if self.selected_item() is item:
            self.load_selected_into_fields()
            self.update_preview()

    def move_item_to_processed(self, item: ReceiptItem) -> bool:
        source = Path(item.path)
        if not source.exists() or not same_folder(source, UNPROCESSED_DIR):
            return False
        try:
            ensure_runtime_folders()
            destination = unique_path(PROCESSED_DIR, source.name)
            shutil.move(str(source), str(destination))
            item.path = str(destination)
            item.filename = destination.name
            LOGGER.info("Moved processed receipt to %s", destination)
            return True
        except Exception:
            log_exception(f"Could not move receipt to Processed: {source}")
            return False

    def move_all_loaded_to_processed(self) -> None:
        moved = False
        for item in self.items:
            moved = self.move_item_to_processed(item) or moved
        if moved:
            self.refresh_tree()
            self.update_preview()

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
                usd_rate = safe_float(self.exchange_rate.get()) or safe_float(DEFAULT_USD_TO_RMB_RATE) or 6.8
                export_korea(self.items, output_path, krw_rate, usd_rate)
        except Exception as exc:
            log_path = log_exception("Excel export failed")
            if messagebox:
                messagebox.showerror(
                    "Generate Excel",
                    f"Could not generate workbook:\n{exc}\n\nLogged to:\n{log_path}",
                )
            self.status_text.set("Export failed.")
            return
        self.move_all_loaded_to_processed()
        self.save_session()
        self.status_text.set(f"Saved: {output_path}")
        if messagebox:
            choice = messagebox.askyesnocancel(
                "Generate Excel",
                "Workbook saved.\n\n"
                "Yes: open the Excel file\n"
                "No: show it in the folder\n"
                "Cancel: stay here",
            )
            try:
                if choice is True:
                    open_path(output_path)
                elif choice is False:
                    reveal_in_file_explorer(output_path)
            except Exception:
                log_exception("Could not open generated workbook or folder")

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

    def session_payload(self) -> Dict[str, Any]:
        self.save_current_fields()
        return {
            "saved_at": datetime.now().isoformat(timespec="seconds"),
            "form_version": self.form_version.get(),
            "usa_exchange_rate": self.exchange_rate.get(),
            "krw_to_rmb_rate": self.krw_to_rmb_rate.get(),
            "selected_index": self.selected_index,
            "items": [asdict(item) for item in self.items],
        }

    def save_session(self) -> None:
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            SESSION_FILE.write_text(
                json.dumps(self.session_payload(), ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            log_exception("Could not save session state")

    def restore_previous_session_if_available(self) -> None:
        if not SESSION_FILE.exists():
            return
        try:
            data = json.loads(SESSION_FILE.read_text(encoding="utf-8"))
        except Exception:
            log_exception("Could not read previous session state")
            return
        raw_items = data.get("items")
        if not isinstance(raw_items, list) or not raw_items:
            return
        should_restore = True
        if messagebox:
            saved_at = str(data.get("saved_at") or "last time")
            should_restore = messagebox.askyesno(
                "Restore Previous File",
                f"A previous reimbursement session was saved at {saved_at}.\n\nRestore it?",
            )
        if not should_restore:
            self.status_text.set("Started a fresh reimbursement session.")
            return
        fields = set(ReceiptItem.__dataclass_fields__.keys())
        restored: List[ReceiptItem] = []
        for raw in raw_items:
            if not isinstance(raw, dict):
                continue
            values = {key: raw.get(key, "") for key in fields}
            if not values.get("item_id"):
                values["item_id"] = uuid.uuid4().hex
            if not isinstance(values.get("crop_box"), list):
                values["crop_box"] = None
            restored.append(ReceiptItem(**values))
        if not restored:
            return
        self.form_version.set(str(data.get("form_version") or "USA"))
        self.exchange_rate.set(str(data.get("usa_exchange_rate") or DEFAULT_USD_TO_RMB_RATE))
        self.krw_to_rmb_rate.set(str(data.get("krw_to_rmb_rate") or DEFAULT_KRW_TO_RMB_RATE))
        self.items = restored
        self._last_form_version = self.form_version.get()
        try:
            selected = int(data.get("selected_index", 0) or 0)
        except (TypeError, ValueError):
            selected = 0
        self.selected_index = None
        self._on_form_version_changed()
        self.refresh_tree()
        self.select_index(max(0, min(selected, len(self.items) - 1)))
        self.status_text.set("Previous reimbursement session restored.")

    def on_close(self) -> None:
        self.settings["usa_exchange_rate"] = self.exchange_rate.get()
        self.settings["krw_to_rmb_rate"] = self.krw_to_rmb_rate.get()
        save_user_settings(self.settings)
        self.save_session()
        self.root.destroy()


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
    export_korea(
        [sample],
        OUTPUT_DIR / "smoke_korea.xlsx",
        safe_float(DEFAULT_KRW_TO_RMB_RATE) or 0.0046,
        safe_float(DEFAULT_USD_TO_RMB_RATE) or 6.8,
    )
    print("Smoke exports created.")


def main(argv: Optional[List[str]] = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    try:
        setup_logging()
        ensure_dependencies()
        if "--smoke-test" in args:
            smoke_test_export()
            return 0
        ensure_runtime_folders()
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
