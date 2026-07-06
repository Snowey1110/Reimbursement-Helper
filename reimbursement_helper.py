from __future__ import annotations

import base64
from concurrent.futures import ThreadPoolExecutor, as_completed
import copy
import hashlib
import json
import logging
import math
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import traceback
import uuid
from dataclasses import asdict, dataclass, field
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
    from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
    from openpyxl.drawing.xdr import XDRPositiveSize2D
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils.cell import column_index_from_string
    from openpyxl.utils.units import pixels_to_EMU
    from openpyxl.worksheet.views import Selection
except Exception:  # pragma: no cover
    Workbook = None  # type: ignore[assignment]
    load_workbook = None  # type: ignore[assignment]
    MergedCell = None  # type: ignore[assignment]
    XLImage = None  # type: ignore[assignment]
    AnchorMarker = None  # type: ignore[assignment]
    OneCellAnchor = None  # type: ignore[assignment]
    XDRPositiveSize2D = None  # type: ignore[assignment]
    Alignment = None  # type: ignore[assignment]
    Font = None  # type: ignore[assignment]
    PatternFill = None  # type: ignore[assignment]
    column_index_from_string = None  # type: ignore[assignment]
    pixels_to_EMU = None  # type: ignore[assignment]
    Selection = None  # type: ignore[assignment]

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
RECEIPT_MEMORY_FILE = CONFIG_DIR / "receipt_memory.json"
SUPPORTED_IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif"}
SUPPORTED_UPLOAD_SUFFIXES = SUPPORTED_IMAGE_SUFFIXES | {".pdf"}
OPENAI_CHAT_COMPLETIONS_URL = "https://api.openai.com/v1/chat/completions"
DEFAULT_OPENAI_MODEL = "gpt-4o-mini"
DEFAULT_USD_TO_RMB_RATE = "6.8175"
DEFAULT_USD_TO_KRW_RATE = "1548.86"
DEFAULT_KRW_TO_RMB_RATE = "0.004433"
PDF_RENDER_SCALE = 3
KRW_NUMBER_FORMAT = "\u20a9#,##0"
USA_SUPPORT_DATE_NUMBER_FORMAT = "m/d/yyyy"
MAX_AI_RECEIPT_IMAGES = 8
AI_AUTO_ROTATE_LANDSCAPE_RATIO = 1.15
EXPORT_AUTO_ROTATE_LANDSCAPE_RATIO = 1.15
DEFAULT_AI_REQUEST_CONCURRENCY = 4
AI_REQUEST_CONCURRENCY_WARNING_THRESHOLD = 12
LOGGER = logging.getLogger("reimbursement_helper")
ExportProgressCallback = Callable[[int, int, str], None]
EXCEL_CONTACT_SHEET_CACHE: Dict[Tuple[Any, ...], Path] = {}
EXCEL_CONTACT_SHEET_CACHE_LOCK = threading.Lock()


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


def emit_export_progress(
    progress: Optional[ExportProgressCallback],
    done: int,
    total: int,
    message: str,
) -> None:
    if progress is None:
        return
    try:
        progress(done, max(1, total), message)
    except Exception:
        pass


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
            pixmap = page.get_pixmap(matrix=fitz.Matrix(PDF_RENDER_SCALE, PDF_RENDER_SCALE), alpha=False)
            output = target_dir / f"{pdf_path.stem}_page_{page_index + 1}.png"
            pixmap.save(str(output))
            rendered.append(output)
    finally:
        document.close()
    return rendered


def render_pdf_to_merged_image(pdf_path: Path) -> Path:
    page_paths = render_pdf_pages(pdf_path)
    if not page_paths:
        raise RuntimeError(f"PDF has no pages: {pdf_path}")
    if Image is None:
        raise RuntimeError(
            "PDF merge requires Pillow. Run `python -m pip install -r requirements.txt`, then try again."
        )
    target_dir = WORK_DIR / "pdf_pages" / f"{pdf_path.stem}_{uuid.uuid4().hex[:8]}"
    target_dir.mkdir(parents=True, exist_ok=True)
    images = []
    try:
        for page_path in page_paths:
            with Image.open(page_path) as image:
                images.append(image.convert("RGB").copy())
        max_width = max(image.width for image in images)
        gutter = 24
        total_height = sum(image.height for image in images) + gutter * (len(images) - 1)
        merged = Image.new("RGB", (max_width, total_height), "#FFFFFF")
        y = 0
        for image in images:
            x = (max_width - image.width) // 2
            merged.paste(image, (x, y))
            y += image.height + gutter
        output = target_dir / f"{pdf_path.stem}_merged.png"
        merged.save(output, "PNG")
        return output
    finally:
        for image in images:
            try:
                image.close()
            except Exception:
                pass


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
        os.startfile(str(path.resolve()))  # type: ignore[attr-defined]
    else:
        raise RuntimeError(f"Opening files is only implemented for Windows. Path: {path}")


def reveal_in_file_explorer(path: Path) -> None:
    if sys.platform == "win32":
        resolved = path.resolve()
        if resolved.exists():
            subprocess.Popen(["explorer.exe", "/select,", str(resolved)])
        else:
            open_path(resolved.parent)
    else:
        open_path(path.parent)


class ToolTip:
    def __init__(self, widget: Any, text: str) -> None:
        self.widget = widget
        self.text = text
        self.tip_window: Optional[Any] = None
        self.after_id: Optional[str] = None
        widget.bind("<Enter>", self.schedule)
        widget.bind("<Leave>", self.hide)
        widget.bind("<ButtonPress>", self.hide)

    def schedule(self, _event: Any = None) -> None:
        self.cancel()
        try:
            self.after_id = self.widget.after(450, self.show)
        except Exception:
            self.after_id = None

    def cancel(self) -> None:
        if self.after_id:
            try:
                self.widget.after_cancel(self.after_id)
            except Exception:
                pass
            self.after_id = None

    def show(self) -> None:
        if self.tip_window is not None or tk is None:
            return
        try:
            x = self.widget.winfo_rootx() + 18
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        except Exception:
            return
        self.tip_window = tk.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)
        self.tip_window.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            self.tip_window,
            text=self.text,
            background="#111827",
            foreground="#F8FAFC",
            borderwidth=0,
            padx=8,
            pady=4,
            font=("Segoe UI", 9),
        )
        label.pack()

    def hide(self, _event: Any = None) -> None:
        self.cancel()
        if self.tip_window is not None:
            try:
                self.tip_window.destroy()
            except Exception:
                pass
            self.tip_window = None


USA_CATEGORY_ROWS: Dict[str, List[int]] = {
    "transportation": list(range(7, 47)),
    "lodging": list(range(49, 52)),
    "meals": [54],
    "advertising": [57],
    "office": list(range(60, 65)),
    "entertainment": list(range(67, 69)),
    "other": list(range(60, 65)),
}

USA_EXPENSE_REPORT_BLOCKS: List[Tuple[str, int, List[int], int]] = [
    ("transportation", 6, list(range(7, 47)), 47),
    ("lodging", 48, list(range(49, 52)), 52),
    ("meals", 53, [54], 55),
    ("advertising", 56, [57], 58),
    ("office", 59, list(range(60, 65)), 65),
    ("entertainment", 66, list(range(67, 69)), 69),
]

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

USA_REPORT_CATEGORY_ORDER = [
    "transportation",
    "lodging",
    "meals",
    "advertising",
    "office",
    "entertainment",
    "other",
]

KOREA_REPORT_CATEGORY_ORDER = [
    "transportation",
    "physical_exam",
    "lodging",
    "nucleic_test",
    "materials",
    "meals",
    "courier",
    "consumables",
    "welfare",
    "other",
]

KOREA_INVOICE_KIND_ORDER = [
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
    "other",
]

DOCUMENT_TYPE_VALUES = {"unknown", "receipt", "payment_proof"}
DOCUMENT_TYPE_LABEL_KEYS = {
    "unknown": "document_type_unknown",
    "receipt": "document_type_receipt",
    "payment_proof": "document_type_payment_proof",
}

KOREA_RECEIPT_PAGE_HEIGHT = 60
KOREA_RECEIPT_SLOTS_PER_PAGE = 4
KOREA_RECEIPT_HALF_WIDTH = 370
KOREA_RECEIPT_WIDE_WIDTH = 760
KOREA_RECEIPT_HALF_HEIGHT = 520
KOREA_RECEIPT_FULL_HEIGHT = 1040
KOREA_FULL_PAGE_SPLIT_ASPECT = 1.75
USA_SUPPORT_ROW_HEIGHT = 540
USA_RECEIPT_COLUMN_WIDTH = 92.7
USA_PROOF_COLUMN_WIDTH = 78
USA_RECEIPT_IMAGE_MAX_WIDTH = 700
USA_RECEIPT_IMAGE_MAX_HEIGHT = 720
USA_PROOF_IMAGE_MAX_WIDTH = 590
USA_PROOF_IMAGE_MAX_HEIGHT = 720
USA_SUPPORT_IMAGE_PADDING_PX = 12
EXCEL_IMAGE_RENDER_SCALE = 2
EXCEL_IMAGE_JPEG_QUALITY = 92

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

CATEGORY_LABELS_BY_LANGUAGE: Dict[str, Dict[str, str]] = {
    "en": {
        "transportation": "Transportation",
        "lodging": "Lodging",
        "meals": "Meals",
        "advertising": "Advertising",
        "office": "Office",
        "entertainment": "Entertainment",
        "materials": "Materials",
        "consumables": "Consumables",
        "physical_exam": "Physical exam",
        "nucleic_test": "Nucleic acid test",
        "courier": "Courier",
        "welfare": "Welfare",
        "other": "Other",
    },
    "zh": {
        "transportation": "\u4ea4\u901a\u8d39",
        "lodging": "\u4f4f\u5bbf\u8d39",
        "meals": "\u4e1a\u52a1\u62db\u5f85\u8d39-\u9910\u8d39",
        "advertising": "\u5e7f\u544a\u8d39",
        "office": "\u529e\u516c\u8d39",
        "entertainment": "\u4e1a\u52a1\u62db\u5f85\u8d39",
        "materials": "\u7269\u6599\u8d39",
        "consumables": "\u6d88\u8017\u54c1",
        "physical_exam": "\u5165\u804c\u4f53\u68c0\u8d39",
        "nucleic_test": "\u6838\u9178\u68c0\u6d4b\u8d39",
        "courier": "\u5feb\u9012\u8d39",
        "welfare": "\u798f\u5229\u8d39",
        "other": "\u5176\u4ed6",
    },
    "ko": {
        "transportation": "\uad50\ud1b5\ube44",
        "lodging": "\uc219\ubc15\ube44",
        "meals": "\uc2dd\ub300",
        "advertising": "\uad11\uace0\ube44",
        "office": "\uc0ac\ubb34\ube44",
        "entertainment": "\uc811\ub300\ube44",
        "materials": "\uc790\uc7ac\ube44",
        "consumables": "\uc18c\ubaa8\ud488",
        "physical_exam": "\uac74\uac15\uac80\uc9c4\ube44",
        "nucleic_test": "\ud575\uc0b0 \uac80\uc0ac\ube44",
        "courier": "\ud0dd\ubc30\ube44",
        "welfare": "\ubcf5\ub9ac\ud6c4\uc0dd\ube44",
        "other": "\uae30\ud0c0",
    },
}

LANGUAGE_LABELS = {
    "en": "English",
    "zh": "\u4e2d\u6587",
    "ko": "\ud55c\uad6d\uc5b4",
}

UI_TEXT: Dict[str, Dict[str, str]] = {
    "en": {
        "app_title": "Reimbursement Helper",
        "header_title": "REIMBURSEMENT HELPER",
        "form": "Form",
        "language": "Language",
        "ai_batch": "AI batch",
        "select_files": "Select Files",
        "select_receipts": "Select Receipts",
        "select_payment_proof": "Select Payment Proof",
        "upload_mode": "Import as",
        "upload_mode_auto": "Auto detect",
        "upload_mode_receipt": "Receipts only",
        "upload_mode_payment_proof": "Payment proof only",
        "select_exchange_rate": "Select Exchange Rate Images",
        "generate_details": "Generate Details",
        "generate_empty": "Generate Empty",
        "generate_all": "Generate All",
        "cancel_generation": "Cancel",
        "cancel_requested": "Cancel requested. Finishing the current file, then stopping...",
        "generation_cancelled": "Generation cancelled. Processed {done}/{total}.",
        "generate_excel": "Generate Excel",
        "inserted_receipts": "Inserted receipts and details",
        "file": "File",
        "status": "Status",
        "date": "Date",
        "amount": "Amount",
        "remove": "Remove",
        "clear": "Clear",
        "details": "Details",
        "receipt_preview": "Receipt preview",
        "receipt_screenshots": "Receipt screenshots",
        "payment_proof_section": "Payment proof",
        "zoom_in": "Zoom in",
        "zoom_out": "Zoom out",
        "rotate_left": "Rotate left",
        "rotate_right": "Rotate right",
        "revert": "Revert",
        "delete_screenshot": "Delete selected screenshot",
        "open_output_folder": "Open output folder",
        "ready": "Ready",
        "clipboard_no_source": "Select a screenshot or receipt row to copy.",
        "clipboard_copied": "Copied {count} screenshot(s).",
        "clipboard_cut": "Cut {count} screenshot(s). Select a receipt row and press Ctrl+V to merge.",
        "clipboard_empty": "No copied screenshot. Press Ctrl+C or Ctrl+X on a screenshot first.",
        "clipboard_no_target": "Select a receipt row before pasting screenshots.",
        "clipboard_needs_receipt_target": "Select a receipt row, not a payment proof row, before pasting.",
        "clipboard_pasted": "Merged {count} screenshot(s).",
        "place_vendor": "Place / Vendor",
        "usd_amount": "USD amount",
        "original_amount": "Original amount",
        "currency": "Currency",
        "krw_amount": "KRW amount",
        "rmb_amount": "RMB amount",
        "purpose": "Purpose",
        "project_number": "Project number",
        "category": "Category",
        "payment_method": "Payment method",
        "receipt_label": "Receipt label",
        "document_type": "Document type",
        "document_type_unknown": "Unknown",
        "document_type_receipt": "Receipt",
        "document_type_payment_proof": "Payment proof",
        "select_receipt_prompt": "Select receipt image or PDF files to begin.",
        "proof_prompt": "Select Receipts, then mark payment proof rows if needed.",
        "no_receipt_screenshot": "No receipt screenshot.",
        "exchange_only_korea": "Exchange-rate images are only used for the Korea form.",
        "selected_exchange_files": "Selected {count} exchange-rate image file(s).",
        "selected_exchange_no_key": "Selected exchange-rate image(s). Add an API key to auto-read the rate.",
        "reading_exchange": "Reading exchange-rate image...",
        "reading_exchange_short": "Reading exchange rate",
        "selected_exchange_failed": "Selected exchange-rate image(s), but could not auto-read rate: {message}",
        "updated_exchange_rate": "Updated rate from exchange-rate image: {updates}",
        "no_empty_rows": "No empty or unprocessed rows need AI.",
        "high_ai_concurrency_title": "High parallel request count",
        "high_ai_concurrency_warning": "You set AI batch to {requested}. This run may send up to {active} OpenAI requests at the same time. A high value can hit rate limits or slow this computer. Continue?",
        "unknown_document_type": "{count} row(s) still have unknown document type: {files}. Choose Receipt or Payment proof before exporting.",
        "manual_review_title": "Manual review needed",
        "unknown_document_type_detail": "{count} row(s) need a file type before Excel can be generated.\n\nSet each one to Receipt or Payment proof:\n{files}",
        "attention_ai_failed": "AI failed",
        "attention_unknown_type": "Missing file type",
        "attention_missing_date": "Missing date",
        "attention_missing_amount": "Missing amount",
        "attention_missing_vendor": "Missing vendor",
        "attention_missing_details": "Missing details",
        "attention_missing_date_amount": "Missing date/amount",
        "attention_wrong_year": "Check year",
    },
    "zh": {
        "app_title": "\u62a5\u9500\u52a9\u624b",
        "header_title": "\u62a5\u9500\u52a9\u624b",
        "form": "\u8868\u683c",
        "language": "\u8bed\u8a00",
        "ai_batch": "\u540c\u65f6\u5904\u7406",
        "select_files": "\u9009\u62e9\u6587\u4ef6",
        "select_receipts": "\u9009\u62e9\u53d1\u7968",
        "select_payment_proof": "\u9009\u62e9\u4ed8\u6b3e\u51ed\u8bc1",
        "upload_mode": "\u5bfc\u5165\u4e3a",
        "upload_mode_auto": "\u81ea\u52a8\u8bc6\u522b",
        "upload_mode_receipt": "\u4ec5\u53d1\u7968",
        "upload_mode_payment_proof": "\u4ec5\u4ed8\u6b3e\u51ed\u8bc1",
        "select_exchange_rate": "\u9009\u62e9\u6c47\u7387\u56fe\u7247",
        "generate_details": "\u751f\u6210\u660e\u7ec6",
        "generate_empty": "\u7a7a\u767d\u586b\u7a7a",
        "generate_all": "\u5168\u90e8\u751f\u6210",
        "cancel_generation": "\u53d6\u6d88",
        "cancel_requested": "\u5df2\u8bf7\u6c42\u53d6\u6d88\u3002\u5f53\u524d\u6587\u4ef6\u5b8c\u6210\u540e\u4f1a\u505c\u6b62...",
        "generation_cancelled": "\u5df2\u53d6\u6d88\u751f\u6210\u3002\u5df2\u5904\u7406 {done}/{total}\u3002",
        "generate_excel": "\u751f\u6210 Excel",
        "inserted_receipts": "\u5df2\u63d2\u5165\u53d1\u7968\u548c\u660e\u7ec6",
        "file": "\u6587\u4ef6",
        "status": "\u72b6\u6001",
        "date": "\u65e5\u671f",
        "amount": "\u91d1\u989d",
        "remove": "\u5220\u9664",
        "clear": "\u6e05\u7a7a",
        "details": "\u660e\u7ec6",
        "receipt_preview": "\u53d1\u7968\u9884\u89c8",
        "receipt_screenshots": "\u53d1\u7968\u622a\u56fe",
        "payment_proof_section": "\u4ed8\u6b3e\u51ed\u8bc1",
        "zoom_in": "\u653e\u5927",
        "zoom_out": "\u7f29\u5c0f",
        "rotate_left": "\u5411\u5de6\u65cb\u8f6c",
        "rotate_right": "\u5411\u53f3\u65cb\u8f6c",
        "revert": "\u8fd8\u539f",
        "delete_screenshot": "\u5220\u9664\u9009\u4e2d\u622a\u56fe",
        "open_output_folder": "\u6253\u5f00\u8f93\u51fa\u6587\u4ef6\u5939",
        "ready": "\u5c31\u7eea",
        "clipboard_no_source": "\u8bf7\u5148\u9009\u4e2d\u4e00\u5f20\u622a\u56fe\u6216\u4e00\u6761\u53d1\u7968\u8bb0\u5f55\u518d\u590d\u5236\u3002",
        "clipboard_copied": "\u5df2\u590d\u5236 {count} \u5f20\u622a\u56fe\u3002",
        "clipboard_cut": "\u5df2\u526a\u5207 {count} \u5f20\u622a\u56fe\u3002\u8bf7\u9009\u4e2d\u76ee\u6807\u53d1\u7968\u540e\u6309 Ctrl+V \u5408\u5e76\u3002",
        "clipboard_empty": "\u6ca1\u6709\u5df2\u590d\u5236\u7684\u622a\u56fe\u3002\u8bf7\u5148\u5728\u622a\u56fe\u4e0a\u6309 Ctrl+C \u6216 Ctrl+X\u3002",
        "clipboard_no_target": "\u7c98\u8d34\u524d\u8bf7\u5148\u9009\u4e2d\u4e00\u6761\u53d1\u7968\u8bb0\u5f55\u3002",
        "clipboard_needs_receipt_target": "\u7c98\u8d34\u524d\u8bf7\u9009\u4e2d\u53d1\u7968\u8bb0\u5f55\uff0c\u4e0d\u8981\u9009\u4ed8\u6b3e\u51ed\u8bc1\u8bb0\u5f55\u3002",
        "clipboard_pasted": "\u5df2\u5408\u5e76 {count} \u5f20\u622a\u56fe\u3002",
        "place_vendor": "\u5730\u70b9 / \u5546\u5bb6",
        "usd_amount": "\u7f8e\u5143\u91d1\u989d",
        "original_amount": "\u539f\u5e01\u91d1\u989d",
        "currency": "\u5e01\u79cd",
        "krw_amount": "\u97e9\u5143\u91d1\u989d",
        "rmb_amount": "\u4eba\u6c11\u5e01\u91d1\u989d",
        "purpose": "\u7528\u9014",
        "project_number": "\u9879\u76ee\u53f7",
        "category": "\u7c7b\u522b",
        "payment_method": "\u4ed8\u6b3e\u65b9\u5f0f",
        "receipt_label": "\u53d1\u7968\u6807\u7b7e",
        "document_type": "\u6587\u4ef6\u7c7b\u578b",
        "document_type_unknown": "\u672a\u77e5",
        "document_type_receipt": "\u53d1\u7968",
        "document_type_payment_proof": "\u4ed8\u6b3e\u51ed\u8bc1",
        "select_receipt_prompt": "\u8bf7\u9009\u62e9\u53d1\u7968\u56fe\u7247\u6216 PDF \u6587\u4ef6\u3002",
        "proof_prompt": "\u8bf7\u9009\u62e9\u53d1\u7968\uff0c\u5fc5\u8981\u65f6\u5c06\u4ed8\u6b3e\u51ed\u8bc1\u884c\u6807\u8bb0\u51fa\u6765\u3002",
        "no_receipt_screenshot": "\u6ca1\u6709\u53d1\u7968\u622a\u56fe\u3002",
        "exchange_only_korea": "\u6c47\u7387\u56fe\u7247\u4ec5\u7528\u4e8e\u97e9\u56fd\u8868\u683c\u3002",
        "selected_exchange_files": "\u5df2\u9009\u62e9 {count} \u4e2a\u6c47\u7387\u56fe\u7247\u6587\u4ef6\u3002",
        "selected_exchange_no_key": "\u5df2\u9009\u62e9\u6c47\u7387\u56fe\u7247\u3002\u6dfb\u52a0 API \u5bc6\u94a5\u540e\u53ef\u81ea\u52a8\u8bfb\u53d6\u6c47\u7387\u3002",
        "reading_exchange": "\u6b63\u5728\u8bfb\u53d6\u6c47\u7387\u56fe\u7247...",
        "reading_exchange_short": "\u6b63\u5728\u8bfb\u53d6\u6c47\u7387",
        "selected_exchange_failed": "\u5df2\u9009\u62e9\u6c47\u7387\u56fe\u7247\uff0c\u4f46\u65e0\u6cd5\u81ea\u52a8\u8bfb\u53d6\u6c47\u7387\uff1a{message}",
        "updated_exchange_rate": "\u5df2\u6839\u636e\u6c47\u7387\u56fe\u7247\u66f4\u65b0\uff1a{updates}",
        "no_empty_rows": "\u6ca1\u6709\u9700\u8981 AI \u5904\u7406\u7684\u7a7a\u767d\u6216\u672a\u5904\u7406\u8bb0\u5f55\u3002",
        "high_ai_concurrency_title": "\u5e76\u884c\u8bf7\u6c42\u8f83\u9ad8",
        "high_ai_concurrency_warning": "\u4f60\u8bbe\u7f6e\u4e86\u540c\u65f6\u5904\u7406 {requested} \u4e2a\uff0c\u672c\u6b21\u5c06\u6700\u591a\u540c\u65f6\u53d1\u9001 {active} \u4e2a OpenAI \u8bf7\u6c42\u3002\u6570\u503c\u8fc7\u9ad8\u53ef\u80fd\u89e6\u53d1\u9650\u901f\u6216\u8ba9\u7535\u8111\u53d8\u6162\u3002\u8981\u7ee7\u7eed\u5417\uff1f",
        "unknown_document_type": "\u8fd8\u6709 {count} \u6761\u8bb0\u5f55\u7684\u6587\u4ef6\u7c7b\u578b\u662f\u672a\u77e5\uff1a{files}\u3002\u5bfc\u51fa\u524d\u8bf7\u9009\u62e9\u53d1\u7968\u6216\u4ed8\u6b3e\u51ed\u8bc1\u3002",
        "manual_review_title": "\u9700\u8981\u4eba\u5de5\u5904\u7406",
        "unknown_document_type_detail": "\u6709 {count} \u6761\u8bb0\u5f55\u9700\u8981\u5148\u9009\u62e9\u6587\u4ef6\u7c7b\u578b\uff0c\u624d\u80fd\u751f\u6210 Excel\u3002\n\n\u8bf7\u628a\u8fd9\u4e9b\u6587\u4ef6\u8bbe\u4e3a\u201c\u53d1\u7968\u201d\u6216\u201c\u4ed8\u6b3e\u51ed\u8bc1\u201d\uff1a\n{files}",
        "attention_ai_failed": "AI \u5931\u8d25",
        "attention_unknown_type": "\u7f3a\u5c11\u6587\u4ef6\u7c7b\u578b",
        "attention_missing_date": "\u7f3a\u5c11\u65e5\u671f",
        "attention_missing_amount": "\u7f3a\u5c11\u91d1\u989d",
        "attention_missing_vendor": "\u7f3a\u5c11\u5546\u5bb6",
        "attention_missing_details": "\u7f3a\u5c11\u660e\u7ec6",
        "attention_missing_date_amount": "\u7f3a\u5c11\u65e5\u671f/\u91d1\u989d",
        "attention_wrong_year": "\u68c0\u67e5\u5e74\u4efd",
    },
    "ko": {
        "app_title": "\ud658\uae09 \ub3c4\uc6b0\ubbf8",
        "header_title": "\ud658\uae09 \ub3c4\uc6b0\ubbf8",
        "form": "\uc591\uc2dd",
        "language": "\uc5b8\uc5b4",
        "ai_batch": "\ub3d9\uc2dc \ucc98\ub9ac",
        "select_files": "\ud30c\uc77c \uc120\ud0dd",
        "select_receipts": "\uc601\uc218\uc99d \uc120\ud0dd",
        "select_payment_proof": "\uacb0\uc81c \uc99d\ube59 \uc120\ud0dd",
        "upload_mode": "\uac00\uc838\uc624\uae30 \uc720\ud615",
        "upload_mode_auto": "\uc790\ub3d9 \uac10\uc9c0",
        "upload_mode_receipt": "\uc601\uc218\uc99d\ub9cc",
        "upload_mode_payment_proof": "\uacb0\uc81c \uc99d\ube59\ub9cc",
        "select_exchange_rate": "\ud658\uc728 \uc774\ubbf8\uc9c0 \uc120\ud0dd",
        "generate_details": "\uc138\ubd80 \uc815\ubcf4 \uc0dd\uc131",
        "generate_empty": "\ube48 \ud56d\ubaa9 \uc0dd\uc131",
        "generate_all": "\uc804\uccb4 \uc0dd\uc131",
        "cancel_generation": "\ucde8\uc18c",
        "cancel_requested": "\ucde8\uc18c\ub97c \uc694\uccad\ud588\uc2b5\ub2c8\ub2e4. \ud604\uc7ac \ud30c\uc77c\uc774 \ub05d\ub098\uba74 \uc911\uc9c0\ud569\ub2c8\ub2e4...",
        "generation_cancelled": "\uc0dd\uc131\uc774 \ucde8\uc18c\ub418\uc5c8\uc2b5\ub2c8\ub2e4. {done}/{total} \ucc98\ub9ac\ub428.",
        "generate_excel": "Excel \uc0dd\uc131",
        "inserted_receipts": "\ucd94\uac00\ub41c \uc601\uc218\uc99d \ubc0f \uc138\ubd80 \uc815\ubcf4",
        "file": "\ud30c\uc77c",
        "status": "\uc0c1\ud0dc",
        "date": "\ub0a0\uc9dc",
        "amount": "\uae08\uc561",
        "remove": "\uc0ad\uc81c",
        "clear": "\ucd08\uae30\ud654",
        "details": "\uc138\ubd80 \uc815\ubcf4",
        "receipt_preview": "\uc601\uc218\uc99d \ubbf8\ub9ac\ubcf4\uae30",
        "receipt_screenshots": "\uc601\uc218\uc99d \uc2a4\ud06c\ub9b0\uc0f7",
        "payment_proof_section": "\uacb0\uc81c \uc99d\ube59",
        "zoom_in": "\ud655\ub300",
        "zoom_out": "\ucd95\uc18c",
        "rotate_left": "\uc67c\ucabd \ud68c\uc804",
        "rotate_right": "\uc624\ub978\ucabd \ud68c\uc804",
        "revert": "\ub418\ub3cc\ub9ac\uae30",
        "delete_screenshot": "\uc120\ud0dd\ud55c \uc2a4\ud06c\ub9b0\uc0f7 \uc0ad\uc81c",
        "open_output_folder": "\ucd9c\ub825 \ud3f4\ub354 \uc5f4\uae30",
        "ready": "\uc900\ube44\ub428",
        "clipboard_no_source": "\uc2a4\ud06c\ub9b0\uc0f7 \ub610\ub294 \uc601\uc218\uc99d \ud589\uc744 \uba3c\uc800 \uc120\ud0dd\ud558\uc138\uc694.",
        "clipboard_copied": "\uc2a4\ud06c\ub9b0\uc0f7 {count}\uac1c\ub97c \ubcf5\uc0ac\ud588\uc2b5\ub2c8\ub2e4.",
        "clipboard_cut": "\uc2a4\ud06c\ub9b0\uc0f7 {count}\uac1c\ub97c \uc798\ub790\uc2b5\ub2c8\ub2e4. \ub300\uc0c1 \uc601\uc218\uc99d \ud589\uc744 \uc120\ud0dd\ud55c \ub4a4 Ctrl+V\ub97c \ub204\ub974\uc138\uc694.",
        "clipboard_empty": "\ubcf5\uc0ac\ub41c \uc2a4\ud06c\ub9b0\uc0f7\uc774 \uc5c6\uc2b5\ub2c8\ub2e4. \uba3c\uc800 \uc2a4\ud06c\ub9b0\uc0f7\uc5d0\uc11c Ctrl+C \ub610\ub294 Ctrl+X\ub97c \ub204\ub974\uc138\uc694.",
        "clipboard_no_target": "\ubd99\uc5ec\ub123\uae30 \uc804\uc5d0 \uc601\uc218\uc99d \ud589\uc744 \uc120\ud0dd\ud558\uc138\uc694.",
        "clipboard_needs_receipt_target": "\ubd99\uc5ec\ub123\uae30 \uc804\uc5d0 \uacb0\uc81c \uc99d\ube59 \ud589\uc774 \uc544\ub2cc \uc601\uc218\uc99d \ud589\uc744 \uc120\ud0dd\ud558\uc138\uc694.",
        "clipboard_pasted": "\uc2a4\ud06c\ub9b0\uc0f7 {count}\uac1c\ub97c \ud569\ucce4\uc2b5\ub2c8\ub2e4.",
        "place_vendor": "\uc7a5\uc18c / \uc5c5\uccb4",
        "usd_amount": "USD \uae08\uc561",
        "original_amount": "\uc6d0\ud654 \uae08\uc561",
        "currency": "\ud1b5\ud654",
        "krw_amount": "KRW \uae08\uc561",
        "rmb_amount": "RMB \uae08\uc561",
        "purpose": "\ubaa9\uc801",
        "project_number": "\ud504\ub85c\uc81d\ud2b8 \ubc88\ud638",
        "category": "\ubd84\ub958",
        "payment_method": "\uacb0\uc81c \uc218\ub2e8",
        "receipt_label": "\uc601\uc218\uc99d \ub77c\ubca8",
        "document_type": "\ubb38\uc11c \uc720\ud615",
        "document_type_unknown": "\uc54c \uc218 \uc5c6\uc74c",
        "document_type_receipt": "\uc601\uc218\uc99d",
        "document_type_payment_proof": "\uacb0\uc81c \uc99d\ube59",
        "select_receipt_prompt": "\uc601\uc218\uc99d \uc774\ubbf8\uc9c0 \ub610\ub294 PDF \ud30c\uc77c\uc744 \uc120\ud0dd\ud558\uc138\uc694.",
        "proof_prompt": "\uc601\uc218\uc99d\uc744 \uc120\ud0dd\ud558\uace0 \ud544\uc694\ud558\uba74 \uacb0\uc81c \uc99d\ube59 \ud589\uc744 \ud45c\uc2dc\ud558\uc138\uc694.",
        "no_receipt_screenshot": "\uc601\uc218\uc99d \uc2a4\ud06c\ub9b0\uc0f7\uc774 \uc5c6\uc2b5\ub2c8\ub2e4.",
        "exchange_only_korea": "\ud658\uc728 \uc774\ubbf8\uc9c0\ub294 \ud55c\uad6d \uc591\uc2dd\uc5d0\ub9cc \uc0ac\uc6a9\ub429\ub2c8\ub2e4.",
        "selected_exchange_files": "\ud658\uc728 \uc774\ubbf8\uc9c0 \ud30c\uc77c {count}\uac1c\ub97c \uc120\ud0dd\ud588\uc2b5\ub2c8\ub2e4.",
        "selected_exchange_no_key": "\ud658\uc728 \uc774\ubbf8\uc9c0\ub97c \uc120\ud0dd\ud588\uc2b5\ub2c8\ub2e4. API \ud0a4\ub97c \ucd94\uac00\ud558\uba74 \uc790\ub3d9\uc73c\ub85c \ud658\uc728\uc744 \uc77d\uc2b5\ub2c8\ub2e4.",
        "reading_exchange": "\ud658\uc728 \uc774\ubbf8\uc9c0\ub97c \uc77d\ub294 \uc911...",
        "reading_exchange_short": "\ud658\uc728 \uc77d\ub294 \uc911",
        "selected_exchange_failed": "\ud658\uc728 \uc774\ubbf8\uc9c0\ub97c \uc120\ud0dd\ud588\uc9c0\ub9cc \uc790\ub3d9\uc73c\ub85c \uc77d\uc9c0 \ubabb\ud588\uc2b5\ub2c8\ub2e4: {message}",
        "updated_exchange_rate": "\ud658\uc728 \uc774\ubbf8\uc9c0\uc5d0\uc11c \ud658\uc728\uc744 \uc5c5\ub370\uc774\ud2b8\ud588\uc2b5\ub2c8\ub2e4: {updates}",
        "no_empty_rows": "AI\uac00 \ucc98\ub9ac\ud560 \ube48 \ud56d\ubaa9\uc774\ub098 \ubbf8\ucc98\ub9ac \ud589\uc774 \uc5c6\uc2b5\ub2c8\ub2e4.",
        "high_ai_concurrency_title": "\ub3d9\uc2dc \uc694\uccad \uc218\uac00 \ub192\uc74c",
        "high_ai_concurrency_warning": "\ub3d9\uc2dc \ucc98\ub9ac\ub97c {requested}\uac1c\ub85c \uc124\uc815\ud588\uc2b5\ub2c8\ub2e4. \uc774\ubc88 \uc2e4\ud589\uc5d0\uc11c OpenAI \uc694\uccad\uc744 \ucd5c\ub300 {active}\uac1c\uae4c\uc9c0 \ub3d9\uc2dc\uc5d0 \ubcf4\ub0bc \uc218 \uc788\uc2b5\ub2c8\ub2e4. \ub108\ubb34 \ub192\uc740 \uac12\uc740 \uc18d\ub3c4 \uc81c\ud55c\uc774\ub098 PC \uc9c0\uc5f0\uc744 \uc77c\uc73c\ud0ac \uc218 \uc788\uc2b5\ub2c8\ub2e4. \uacc4\uc18d\ud560\uae4c\uc694?",
        "unknown_document_type": "{count}\uac1c \ud589\uc758 \ubb38\uc11c \uc720\ud615\uc774 \uc54c \uc218 \uc5c6\uc74c\uc785\ub2c8\ub2e4: {files}. \ub0b4\ubcf4\ub0b4\uae30 \uc804\uc5d0 \uc601\uc218\uc99d \ub610\ub294 \uacb0\uc81c \uc99d\ube59\uc744 \uc120\ud0dd\ud558\uc138\uc694.",
        "manual_review_title": "\uc218\ub3d9 \ud655\uc778 \ud544\uc694",
        "unknown_document_type_detail": "Excel\uc744 \uc0dd\uc131\ud558\uae30 \uc804\uc5d0 {count}\uac1c \ud589\uc758 \ubb38\uc11c \uc720\ud615\uc744 \uc9c0\uc815\ud574\uc57c \ud569\ub2c8\ub2e4.\n\n\uac01 \ud30c\uc77c\uc744 \uc601\uc218\uc99d \ub610\ub294 \uacb0\uc81c \uc99d\ube59\uc73c\ub85c \uc124\uc815\ud558\uc138\uc694:\n{files}",
        "attention_ai_failed": "AI \uc2e4\ud328",
        "attention_unknown_type": "\ubb38\uc11c \uc720\ud615 \ub204\ub77d",
        "attention_missing_date": "\ub0a0\uc9dc \ub204\ub77d",
        "attention_missing_amount": "\uae08\uc561 \ub204\ub77d",
        "attention_missing_vendor": "\uc0c1\ud638 \ub204\ub77d",
        "attention_missing_details": "\uc138\ubd80 \ub0b4\uc6a9 \ub204\ub77d",
        "attention_missing_date_amount": "\ub0a0\uc9dc/\uae08\uc561 \ub204\ub77d",
        "attention_wrong_year": "\uc5f0\ub3c4 \ud655\uc778",
    },
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
    document_type: str = "receipt"
    locked_document_type: str = ""
    matched_receipt_id: str = ""
    status: str = "Empty"
    crop_box: Optional[List[float]] = None
    rotation_degrees: int = 0
    receipt_images: List[Dict[str, Any]] = field(default_factory=list)


@dataclass
class BankStatementItem:
    item_id: str
    path: str
    filename: str
    source_path: str = ""
    source_page: str = ""
    date: str = ""
    amount: str = ""
    place: str = ""
    matched_receipt_id: str = ""
    status: str = "Needs AI"
    crop_box: Optional[List[float]] = None
    rotation_degrees: int = 0


def attachment_from_item(item: Any) -> Dict[str, Any]:
    return {
        "id": getattr(item, "item_id", uuid.uuid4().hex),
        "path": getattr(item, "path", ""),
        "filename": getattr(item, "filename", ""),
        "source_path": getattr(item, "source_path", ""),
        "source_page": getattr(item, "source_page", ""),
        "crop_box": copy.deepcopy(getattr(item, "crop_box", None)),
        "rotation_degrees": normalize_rotation(getattr(item, "rotation_degrees", 0)),
    }


def normalize_attachment(raw: Any) -> Optional[Dict[str, Any]]:
    if not isinstance(raw, dict):
        return None
    path = str(raw.get("path") or "").strip()
    if not path:
        return None
    return {
        "id": str(raw.get("id") or uuid.uuid4().hex),
        "path": path,
        "filename": str(raw.get("filename") or Path(path).name),
        "source_path": str(raw.get("source_path") or ""),
        "source_page": str(raw.get("source_page") or ""),
        "crop_box": raw.get("crop_box") if isinstance(raw.get("crop_box"), list) else None,
        "rotation_degrees": normalize_rotation(raw.get("rotation_degrees", 0)),
    }


def clone_attachment_for_manual_paste(raw: Any) -> Optional[Dict[str, Any]]:
    attachment = normalize_attachment(raw)
    if attachment is None:
        return None
    attachment["id"] = uuid.uuid4().hex
    attachment["crop_box"] = copy.deepcopy(attachment.get("crop_box"))
    return attachment


def fill_blank_receipt_fields(target: ReceiptItem, source: Optional[ReceiptItem]) -> None:
    if source is None:
        return
    for attr in (
        "date",
        "place",
        "amount",
        "krw_amount",
        "rmb_amount",
        "purpose",
        "details",
        "project_number",
        "payment_method",
        "receipt_label",
    ):
        current = str(getattr(target, attr, "") or "").strip()
        incoming = str(getattr(source, attr, "") or "").strip()
        if not current and incoming:
            setattr(target, attr, incoming)


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


def normalized_form_version(value: Any) -> str:
    return "Korea" if str(value) == "Korea" else "USA"


def normalize_document_type(value: Any, default: str = "unknown") -> str:
    text = str(value or "").strip().lower().replace("-", "_").replace(" ", "_")
    if text in {"receipt", "invoice", "merchant_receipt"}:
        return "receipt"
    if text in {
        "payment_proof",
        "proof",
        "payment",
        "bank_statement",
        "card_statement",
        "credit_card",
        "transaction",
        "chase",
    }:
        return "payment_proof"
    if text in {"unknown", "unclear", "manual_review", "needs_review"}:
        return "unknown"
    return default if default in DOCUMENT_TYPE_VALUES else "unknown"


def normalize_usa_upload_mode(value: Any) -> str:
    text = str(value or "").strip().lower().replace("-", "_").replace(" ", "_")
    return text if text in {"receipt", "payment_proof"} else "auto"


def normalize_language(value: Any) -> str:
    return str(value) if str(value) in UI_TEXT else "en"


def ui_text(language: str, key: str, **values: Any) -> str:
    template = UI_TEXT.get(normalize_language(language), UI_TEXT["en"]).get(key, UI_TEXT["en"].get(key, key))
    for name, value in values.items():
        template = template.replace("{" + name + "}", str(value))
    return template


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


def ai_receipt_image_variants(prepared_path: Path, label: str) -> List[Tuple[Path, str]]:
    if Image is None:
        return [(prepared_path, label)]
    try:
        with Image.open(prepared_path) as source:
            width, height = source.size
            if width <= max(1, int(height * AI_AUTO_ROTATE_LANDSCAPE_RATIO)):
                return [(prepared_path, label)]
            image = source.copy()
    except Exception:
        return [(prepared_path, label)]
    target_dir = WORK_DIR / "prepared_images"
    target_dir.mkdir(parents=True, exist_ok=True)
    variants: List[Tuple[Path, str]] = []
    for degrees, suffix in ((90, "rotated left"), (-90, "rotated right")):
        target = target_dir / f"{uuid.uuid4().hex}_ai_{suffix.replace(' ', '_')}.png"
        image.rotate(degrees, expand=True).save(target, "PNG")
        variants.append((target, f"{label} ({suffix})"))
    return variants


def receipt_ai_image_parts(image_path: Path, existing: ReceiptItem) -> List[Dict[str, Any]]:
    attachments = [attachment for attachment in (normalize_attachment(raw) for raw in existing.receipt_images) if attachment]
    if not attachments:
        attachments = [
            {
                "id": existing.item_id,
                "path": str(image_path),
                "filename": existing.filename or image_path.name,
                "source_path": existing.source_path,
                "source_page": existing.source_page,
                "crop_box": existing.crop_box,
                "rotation_degrees": normalize_rotation(existing.rotation_degrees),
            }
        ]
    parts: List[Dict[str, Any]] = []
    for attachment in attachments:
        if len(parts) >= MAX_AI_RECEIPT_IMAGES:
            break
        path = Path(str(attachment.get("path") or ""))
        if not path.exists():
            continue
        prepared = prepare_attachment_image_file(attachment)
        label = str(attachment.get("filename") or path.name)
        source_page = str(attachment.get("source_page") or "").strip()
        if source_page and f"page {source_page}" not in label.lower():
            label = f"{label} (page {source_page})"
        for variant_path, variant_label in ai_receipt_image_variants(prepared, label):
            if len(parts) >= MAX_AI_RECEIPT_IMAGES:
                break
            parts.append({"label": variant_label, "path": variant_path, "index": len(parts) + 1})
    if parts:
        return parts
    prepared_image = prepare_receipt_image_file(image_path, existing.crop_box, existing.rotation_degrees)
    return [{"label": existing.filename or image_path.name, "path": prepared_image, "index": 1}]


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
    image_parts = receipt_ai_image_parts(image_path, existing)
    form_hint = "USA VisionNav reimbursement form" if form_version == "USA" else "Korea VisionNav reimbursement form"
    page_count_hint = len(image_parts)
    prompt = f"""
Extract reimbursement fields from the attached receipt image(s) for a {form_hint}.
Return only compact JSON with these keys:
date, place, amount, currency, krw_amount, rmb_amount, purpose, details,
project_number, category, payment_method, receipt_label, document_type,
rotation_degrees, confidence_notes.

Rules:
- date must be yyyy-mm-dd when visible, otherwise empty string.
- amount should be the paid total. Preserve cents when visible.
- If the receipt text is sideways or upside down, mentally rotate the image upright
  before reading the fields.
- rotation_degrees should be 0, 90, 180, or 270 clockwise degrees needed to rotate
  the original primary image upright. If you used a candidate labeled rotated right,
  return 90. If you used a candidate labeled rotated left, return 270.
- Some uploads are multi-page PDFs. Review all {page_count_hint} attached image(s);
  the paid total may be on a later page, payment info panel, or receipt summary page.
  Use the page that contains the final Total/Payment info/Paid amount. Ignore page
  numbers, rewards points, ticket numbers, order IDs, and QR-code values.
- If an image label says rotated left or rotated right, it is an orientation candidate
  for the same page. Use the candidate where text is upright and readable; do not
  treat orientation candidates as separate receipts.
- For USA bank/card payment proof screenshots, date should be the Transaction date
  line and amount should be the Sale/charge amount. Ignore print/export time and
  use Posted date only when Transaction date is not present.
- currency must match the receipt symbol/text: USD for "$", KRW for won/KRW,
  and RMB/CNY for yuan/RMB/CNY. Do not default to KRW just because the form is Korea.
- category must be one of:
  transportation, lodging, meals, advertising, office, entertainment, physical_exam,
  nucleic_test, materials, courier, consumables, welfare, other.
- project_number should stay empty unless printed on the receipt or obvious from context.
- purpose should be short, for example Parking, Gas, Hotel, Meal, Baggage, Supplies.
- For USA gas-station receipts, including 7-Eleven, Shell, BP, Exxon, Mobil,
  Chevron, Speedway, Circle K, Marathon, Sunoco, Valero, Pilot, Love's, Wawa,
  QuikTrip, Casey's, Murphy, Costco Gas, or Meijer Gas, use purpose Gas and
  category transportation. Do not classify those as snacks or office purchases.
- For USA parking or garage charges, including merchant type "Parking lots and
  garages", Wabash Randolph, McCormick Place Lot, iParkit, Metropolis, ParkWhiz,
  SpotHero, or any parking/toll/garage wording, use purpose Parking or Toll and
  category transportation. Do not classify parking as breakfast, meals, food, or
  entertainment.
- details should be a natural one-line description.
- receipt_label should be short enough to fit above a receipt image in Excel.
- For USA, document_type must be receipt, payment_proof, or unknown. Use payment_proof
  only for bank/card transaction screenshots or statements, not merchant receipts.
  A merchant receipt that says Credit Card Receipt, Payment Method, Card Type,
  authorization, terminal ID, total, or receipt number is still a receipt.
  If the attached images include both a merchant receipt/ticket and a bank/card
  transaction screenshot, classify the entry as receipt and extract the merchant
  receipt/ticket details.
- For Korea, use receipt unless the image is clearly not reimbursement support.
- If a field is unclear, use an empty string rather than guessing.
- Image labels or filenames may contain an app memory marker beginning with __RH1_.
  Treat that marker as prior reimbursement memory. Prefer manual-reviewed marker values
  for document_type, date, amount, currency, category, and vendor unless the visible
  receipt clearly contradicts them.

Current editable row values, to keep if they look more specific than the image:
{json.dumps(asdict(existing), ensure_ascii=False)}
""".strip()
    user_content: List[Dict[str, Any]] = [{"type": "text", "text": prompt}]
    for part in image_parts:
        user_content.append({"type": "text", "text": f"Image {part['index']}: {part['label']}"})
        user_content.append({"type": "image_url", "image_url": {"url": image_data_url(Path(part["path"]))}})
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
                    "content": user_content,
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


def exchange_rates_from_data(data: Dict[str, Any], usd_to_rmb_rate: float) -> Dict[str, float]:
    rates: Dict[str, float] = {}
    for key in ("usd_to_krw_rate", "usd_to_krw"):
        value = safe_float(data.get(key))
        if value is not None and value > 1:
            rates["usd_to_krw_rate"] = value
            break
    krw_to_usd = safe_float(data.get("krw_to_usd_rate")) or safe_float(data.get("krw_to_usd"))
    if "usd_to_krw_rate" not in rates and krw_to_usd and 0 < krw_to_usd < 1:
        rates["usd_to_krw_rate"] = 1 / krw_to_usd
    for key in ("krw_to_rmb_rate", "krw_to_rmb", "rate", "exchange_rate"):
        value = safe_float(data.get(key))
        if value is not None and 0 < value < 1:
            rates["krw_to_rmb_rate"] = value
            break
    usd_to_krw = rates.get("usd_to_krw_rate")
    if "krw_to_rmb_rate" not in rates and usd_to_krw and usd_to_rmb_rate:
        rates["krw_to_rmb_rate"] = usd_to_rmb_rate / usd_to_krw
    return rates


def format_exchange_rate(rate: float) -> str:
    return f"{rate:.10f}".rstrip("0").rstrip(".")


def call_openai_exchange_rate_extraction(
    image_paths: Sequence[Path],
    usd_to_rmb_rate: float,
    progress: Optional[Callable[[str], None]] = None,
) -> Dict[str, float]:
    api_key = get_openai_api_key()
    if not api_key:
        raise RuntimeError("No OpenAI API key was found, so the exchange rate image could not be read.")
    existing_paths = [path for path in image_paths if path.exists()]
    if not existing_paths:
        raise RuntimeError("No readable exchange rate images were found.")

    def notify(message: str) -> None:
        if progress:
            try:
                progress(message)
            except Exception:
                pass

    notify("Reading exchange rate image")
    prompt = f"""
Read these exchange-rate screenshots for the Korea reimbursement form.
Return only compact JSON with these keys:
usd_to_krw_rate, krw_to_usd_rate, krw_to_rmb_rate, confidence_notes.

Rules:
- Prefer an explicit USD -> KRW value, for example 1548.86, when a screenshot shows
  1 USD = 1548.86 KRW.
- If a screenshot shows KRW -> USD instead, return krw_to_usd_rate and leave
  usd_to_krw_rate empty.
- Prefer an explicit KRW -> RMB or 汇率 value, for example 0.0044029590.
- If the images only show USD -> KRW, for example 1 USD = 1548.86 KRW,
  calculate KRW -> RMB as USD_TO_RMB / USD_TO_KRW.
- The current USD_TO_RMB value from the app is {usd_to_rmb_rate:g}.
- Return numbers as plain decimals without currency symbols or commas.
- If unsure, leave krw_to_rmb_rate empty and explain briefly in confidence_notes.
""".strip()
    content: List[Dict[str, Any]] = [{"type": "text", "text": prompt}]
    for path in existing_paths:
        content.append({"type": "image_url", "image_url": {"url": image_data_url(path)}})
    payload = json.dumps(
        {
            "model": os.getenv("OPENAI_MODEL", DEFAULT_OPENAI_MODEL),
            "messages": [
                {
                    "role": "system",
                    "content": "You extract numeric exchange rates from screenshots for reimbursement forms.",
                },
                {"role": "user", "content": content},
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

    try:
        parsed = json.loads(body)
        response_text = parsed["choices"][0]["message"]["content"]
    except Exception as exc:
        raise RuntimeError("OpenAI returned an unexpected exchange-rate response format.") from exc
    data = extract_json_object(str(response_text))
    rates = exchange_rates_from_data(data, usd_to_rmb_rate)
    if not rates:
        raise RuntimeError("AI could not find a usable exchange rate.")
    return rates


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
    raw = (value or "").strip()
    if not raw:
        return "other"
    if " - " in raw:
        raw = raw.split(" - ", 1)[0].strip()
    normalized = normalized_category(raw, form_version)
    if normalized != "other" or raw.lower() == "other":
        return normalized
    for labels in CATEGORY_LABELS_BY_LANGUAGE.values():
        for key, label in labels.items():
            if raw == label:
                return normalized_category(key, form_version)
    return normalized


def format_amount(value: float) -> str:
    if abs(value) >= 1000:
        return f"{value:.2f}"
    return f"{value:.2f}"


def truncate_amount(value: Optional[float]) -> Optional[int]:
    if value is None:
        return None
    return math.trunc(value)


def clamp_ai_concurrency(value: Any) -> int:
    try:
        parsed = int(round(float(value)))
    except (TypeError, ValueError):
        parsed = DEFAULT_AI_REQUEST_CONCURRENCY
    return max(1, parsed)


def parse_date_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = (str(value) if value is not None else "").strip()
    if not text:
        return None
    formats = [
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
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


def excel_report_date_value(value: Any, expected_year: Optional[int] = None) -> Any:
    parsed = parse_date_value(value)
    if isinstance(parsed, date):
        year = expected_year or expected_report_year()
        try:
            return parsed.replace(year=year)
        except ValueError:
            return parsed
    return parsed


def normalized_category(category: str, form_version: str) -> str:
    raw = (category or "").strip().lower()
    if is_korea_other_text(raw):
        return "other"
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


RECEIPT_MEMORY_MARKER_RE = re.compile(r"__RH1_[A-Za-z0-9._~-]+__")
RECEIPT_MEMORY_SCHEMA_VERSION = 1


def strip_receipt_memory_marker(value: str) -> str:
    cleaned = RECEIPT_MEMORY_MARKER_RE.sub("", value or "")
    cleaned = re.sub(r"[_\s-]+$", "", cleaned).strip()
    return cleaned or "receipt"


def filename_safe_token(value: Any, limit: int = 28) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    text = re.sub(r"[^A-Za-z0-9]+", "-", text)
    text = re.sub(r"-+", "-", text).strip("-")
    return text[:limit].strip("-")


def file_sha256_short(path: Path, length: int = 12) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()[:length]


def receipt_memory_key_for_path(path: Path, source_page: str = "") -> str:
    if not path.exists() or not path.is_file():
        return ""
    try:
        digest = file_sha256_short(path, 24)
    except Exception:
        return ""
    page = str(source_page or "").strip()
    return f"sha256:{digest}:page:{page}"


def receipt_source_path_for_memory(path: str, source_path: str) -> Path:
    source = Path(str(source_path or "").strip()) if str(source_path or "").strip() else Path(str(path or ""))
    return source


def receipt_memory_keys_for_item(item: Any) -> List[str]:
    keys: List[str] = []
    source = receipt_source_path_for_memory(getattr(item, "path", ""), getattr(item, "source_path", ""))
    key = receipt_memory_key_for_path(source, getattr(item, "source_page", ""))
    if key:
        keys.append(key)
    for raw in getattr(item, "receipt_images", []) or []:
        if not isinstance(raw, dict):
            continue
        attachment_source = receipt_source_path_for_memory(str(raw.get("path") or ""), str(raw.get("source_path") or ""))
        attachment_key = receipt_memory_key_for_path(attachment_source, str(raw.get("source_page") or ""))
        if attachment_key and attachment_key not in keys:
            keys.append(attachment_key)
    return keys


def load_receipt_memory() -> Dict[str, Any]:
    if not RECEIPT_MEMORY_FILE.exists():
        return {"version": RECEIPT_MEMORY_SCHEMA_VERSION, "records": {}}
    try:
        data = json.loads(RECEIPT_MEMORY_FILE.read_text(encoding="utf-8"))
    except Exception:
        log_exception("Could not read receipt memory")
        return {"version": RECEIPT_MEMORY_SCHEMA_VERSION, "records": {}}
    if not isinstance(data, dict):
        return {"version": RECEIPT_MEMORY_SCHEMA_VERSION, "records": {}}
    records = data.get("records")
    if not isinstance(records, dict):
        data["records"] = {}
    data["version"] = RECEIPT_MEMORY_SCHEMA_VERSION
    return data


def save_receipt_memory(data: Dict[str, Any]) -> None:
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        RECEIPT_MEMORY_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        log_exception("Could not save receipt memory")


def receipt_memory_record_from_item(item: Any, form_version: str, reviewed_source: str) -> Dict[str, Any]:
    document_type = normalize_document_type(getattr(item, "document_type", "unknown"), "unknown")
    currency = normalize_currency(getattr(item, "currency", ""), "USD" if normalized_form_version(form_version) == "USA" else "KRW")
    return {
        "version": RECEIPT_MEMORY_SCHEMA_VERSION,
        "reviewed_source": "manual" if reviewed_source == "manual" else "ai",
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "form_version": normalized_form_version(form_version),
        "document_type": document_type,
        "date": str(getattr(item, "date", "") or "").strip(),
        "place": str(getattr(item, "place", "") or "").strip(),
        "amount": str(getattr(item, "amount", "") or "").strip(),
        "currency": currency,
        "krw_amount": str(getattr(item, "krw_amount", "") or "").strip(),
        "rmb_amount": str(getattr(item, "rmb_amount", "") or "").strip(),
        "purpose": str(getattr(item, "purpose", "") or "").strip(),
        "details": str(getattr(item, "details", "") or "").strip(),
        "project_number": str(getattr(item, "project_number", "") or "").strip(),
        "category": normalized_category(str(getattr(item, "category", "") or ""), normalized_form_version(form_version)),
        "payment_method": str(getattr(item, "payment_method", "") or "").strip(),
        "receipt_label": str(getattr(item, "receipt_label", "") or "").strip(),
    }


def receipt_memory_record_has_values(record: Dict[str, Any]) -> bool:
    for key in ("date", "amount", "krw_amount", "rmb_amount", "place", "purpose", "details", "receipt_label"):
        if str(record.get(key) or "").strip():
            return True
    return normalize_document_type(record.get("document_type"), "unknown") != "unknown"


def build_receipt_memory_marker(record: Dict[str, Any], fingerprint: str = "") -> str:
    parts: List[str] = []
    doc_type = normalize_document_type(record.get("document_type"), "unknown")
    if doc_type != "unknown":
        parts.append(f"t-{doc_type}")
    date_text = filename_safe_token(record.get("date"), 10)
    if date_text:
        parts.append(f"d-{date_text}")
    amount_number = re.sub(r"[^0-9.-]+", "", str(record.get("amount") or "").strip()).replace(".", "p")
    amount_text = filename_safe_token(f"{amount_number}{record.get('currency', '')}", 18)
    if amount_text:
        parts.append(f"a-{amount_text}")
    category = filename_safe_token(record.get("category"), 20)
    if category:
        parts.append(f"c-{category}")
    vendor = filename_safe_token(record.get("place"), 24)
    if vendor:
        parts.append(f"v-{vendor}")
    reviewed = filename_safe_token(record.get("reviewed_source"), 8)
    if reviewed:
        parts.append(f"r-{reviewed}")
    if fingerprint:
        parts.append(f"f-{filename_safe_token(fingerprint, 12)}")
    return "__RH1_" + "_".join(parts[:7]) + "__"


def parse_receipt_memory_marker(value: str) -> Dict[str, Any]:
    match = RECEIPT_MEMORY_MARKER_RE.search(value or "")
    if not match:
        return {}
    body = match.group(0)[len("__RH1_") : -2]
    record: Dict[str, Any] = {"version": RECEIPT_MEMORY_SCHEMA_VERSION}
    for piece in body.split("_"):
        if "-" not in piece:
            continue
        key, raw_value = piece.split("-", 1)
        value_text = raw_value.replace("-", " ").strip()
        if key == "t":
            record["document_type"] = normalize_document_type(raw_value, "unknown")
        elif key == "d":
            record["date"] = raw_value
        elif key == "a":
            amount_match = re.match(r"(-?\d+(?:[.p-]\d+)?)([A-Z]+)?", raw_value)
            if amount_match:
                amount_text = amount_match.group(1).replace("p", ".")
                if "." not in amount_text and "-" in amount_text[1:]:
                    sign = "-" if amount_text.startswith("-") else ""
                    body = amount_text[1:] if sign else amount_text
                    amount_text = sign + body.replace("-", ".", 1)
                record["amount"] = amount_text
                if amount_match.group(2):
                    record["currency"] = normalize_currency(amount_match.group(2), "USD")
        elif key == "c":
            record["category"] = raw_value
        elif key == "v":
            record["place"] = value_text
        elif key == "r":
            record["reviewed_source"] = raw_value
    return record


def apply_receipt_memory_record_to_item(item: Any, record: Dict[str, Any], form_version: str, full_record: bool = True) -> bool:
    if not isinstance(record, dict) or not record:
        return False
    changed = False
    field_map = {
        "date": "date",
        "place": "place",
        "amount": "amount",
        "currency": "currency",
        "krw_amount": "krw_amount",
        "rmb_amount": "rmb_amount",
        "purpose": "purpose",
        "details": "details",
        "project_number": "project_number",
        "category": "category",
        "payment_method": "payment_method",
        "receipt_label": "receipt_label",
        "document_type": "document_type",
    }
    for source, dest in field_map.items():
        value = record.get(source)
        if value is None:
            continue
        text = str(value).strip()
        if not text and dest not in {"project_number"}:
            continue
        if dest == "category":
            text = normalized_category(text, normalized_form_version(form_version))
        elif dest == "currency":
            text = normalize_currency(text, "USD" if normalized_form_version(form_version) == "USA" else "KRW")
        elif dest == "document_type":
            text = normalize_document_type(text, "unknown")
        if getattr(item, dest, "") != text:
            setattr(item, dest, text)
            changed = True
    if full_record:
        reviewed = str(record.get("reviewed_source") or "").strip()
        item.status = "Manual remembered" if reviewed == "manual" else "Remembered"
    return changed


def receipt_item_signature(item: Any) -> Tuple[Any, ...]:
    return (
        getattr(item, "date", ""),
        getattr(item, "place", ""),
        getattr(item, "amount", ""),
        getattr(item, "currency", ""),
        getattr(item, "krw_amount", ""),
        getattr(item, "rmb_amount", ""),
        getattr(item, "purpose", ""),
        getattr(item, "details", ""),
        getattr(item, "project_number", ""),
        getattr(item, "category", ""),
        getattr(item, "payment_method", ""),
        getattr(item, "receipt_label", ""),
        getattr(item, "document_type", ""),
    )


USA_FUEL_VENDOR_PATTERN = re.compile(
    r"\b("
    r"7[\s-]?eleven|shell|bp|exxon|mobil|chevron|speedway|circle\s*k|marathon|"
    r"sunoco|valero|pilot|love'?s|wawa|quiktrip|quicktrip|casey'?s|murphy|"
    r"costco\s*gas|meijer\s*gas|gas\s*station"
    r")\b",
    re.IGNORECASE,
)


USA_PARKING_VENDOR_PATTERN = re.compile(
    r"\b("
    r"parking|park(?:ing)?\s+lot|parking\s+lots?\s+and\s+garages?|garage|self\s+park|"
    r"valet|meter|metropolis|iparkit|parkwhiz|spot\s*hero|wabash\s+randolph|"
    r"mccormick\s+place\s+lot|millennium\s+garages?|interpark|laz\s+parking|"
    r"sp\+\s*parking|airport\s+parking"
    r")\b",
    re.IGNORECASE,
)


def looks_like_usa_fuel_receipt(item: Any, *extra_values: Any) -> bool:
    text = " ".join(
        str(value or "")
        for value in (
            getattr(item, "place", ""),
            getattr(item, "purpose", ""),
            getattr(item, "details", ""),
            getattr(item, "receipt_label", ""),
            getattr(item, "filename", ""),
            *extra_values,
        )
    )
    lower = text.lower()
    if re.search(r"\b(fuel|gasoline|gas purchase|petrol|diesel|unleaded|gallons?|pump)\b", lower):
        return True
    return bool(USA_FUEL_VENDOR_PATTERN.search(text))


def normalize_usa_fuel_receipt_fields(item: Any, *extra_values: Any) -> None:
    if not looks_like_usa_fuel_receipt(item, *extra_values):
        return
    item.category = "transportation"
    item.purpose = "Gas"
    place = str(getattr(item, "place", "") or "").strip()
    current_details = str(getattr(item, "details", "") or "").strip()
    if not current_details or re.search(r"\b(snacks?|convenience|purchase at|purchasing)\b", current_details, re.IGNORECASE):
        item.details = f"Fuel purchase at {place}." if place else "Fuel purchase."
    label = str(getattr(item, "receipt_label", "") or "").strip()
    if not label or re.search(r"\b(snacks?|convenience|purchase)\b", label, re.IGNORECASE):
        item.receipt_label = "Gas Receipt"


def looks_like_usa_parking_receipt(item: Any, *extra_values: Any) -> bool:
    text = " ".join(
        str(value or "")
        for value in (
            getattr(item, "place", ""),
            getattr(item, "purpose", ""),
            getattr(item, "details", ""),
            getattr(item, "receipt_label", ""),
            getattr(item, "filename", ""),
            getattr(item, "payment_method", ""),
            *extra_values,
        )
    )
    lower = text.lower()
    if re.search(r"\b(toll|bridge fee|road charge|turnpike)\b", lower):
        return True
    return bool(USA_PARKING_VENDOR_PATTERN.search(text))


def normalize_usa_parking_receipt_fields(item: Any, *extra_values: Any) -> None:
    if not looks_like_usa_parking_receipt(item, *extra_values):
        return
    item.category = "transportation"
    place = str(getattr(item, "place", "") or "").strip()
    text = " ".join(str(value or "") for value in extra_values).lower()
    is_toll = bool(
        re.search(
            r"\b(toll|bridge fee|road charge|turnpike)\b",
            " ".join(
                str(value or "")
                for value in (
                    getattr(item, "place", ""),
                    getattr(item, "purpose", ""),
                    getattr(item, "details", ""),
                    getattr(item, "receipt_label", ""),
                    getattr(item, "filename", ""),
                    text,
                )
            ).lower(),
        )
    )
    corrected_purpose = "Toll" if is_toll else "Parking"
    current_purpose = str(getattr(item, "purpose", "") or "").strip()
    if not current_purpose or re.search(r"\b(breakfast|lunch|dinner|meal|meals|food|restaurant|snack)\b", current_purpose, re.IGNORECASE):
        item.purpose = corrected_purpose
    current_details = str(getattr(item, "details", "") or "").strip()
    if not current_details or re.search(r"\b(breakfast|lunch|dinner|meal|meals|food|restaurant|snack)\b", current_details, re.IGNORECASE):
        if is_toll:
            item.details = f"Toll charge at {place}." if place else "Toll charge."
        else:
            item.details = f"Parking at {place}." if place else "Parking fee."
    current_label = str(getattr(item, "receipt_label", "") or "").strip()
    if not current_label or re.search(r"\b(breakfast|lunch|dinner|meal|meals|food|restaurant|snack)\b", current_label, re.IGNORECASE):
        item.receipt_label = "Toll Receipt" if is_toll else "Parking Receipt"


def is_korea_other_text(value: str) -> bool:
    raw = (value or "").lower()
    return bool(re.search(r"\b(e-?sim|sim card|data plan|internet access)\b", raw)) or any(
        token in raw for token in ("\u6d41\u91cf", "\u865a\u62df\u5361", "\uc720\uc2ec", "\ub370\uc774\ud130")
    )


def item_search_text(item: Any) -> str:
    return " ".join(
        str(getattr(item, attr, "") or "")
        for attr in ("category", "purpose", "details", "receipt_label", "place", "filename", "payment_method")
    ).lower()


def korea_invoice_kind_for_item(item: Any) -> str:
    text = item_search_text(item)
    category = report_category_for_item(item, "Korea")
    if re.search(r"\b(national|car rental|rental car|rent[- ]?a[- ]?car|rental|avis|hertz|enterprise|budget|alamo)\b", text):
        return "car_rental"
    if re.search(r"\b(fuel|gas|gasoline|petrol|shell|bp|chevron|exxon)\b", text):
        return "fuel"
    if re.search(r"\b(parking|park|toll)\b", text):
        return "parking"
    if is_korea_other_text(text):
        return "esim"
    if re.search(r"\b(usb|office|supplies|supply|material|materials)\b", text):
        return "materials"
    if category in KOREA_INVOICE_KIND_ORDER:
        return category
    return "other"


def korea_detail_content(item: "ReceiptItem") -> str:
    text = item_search_text(item)
    category = report_category_for_item(item, "Korea")
    if korea_invoice_kind_for_item(item) == "car_rental":
        return "\u79df\u8f66\u8d39\u7528"
    if re.search(r"\b(fuel|gas|gasoline|petrol|shell|bp|chevron|exxon)\b", text):
        return "\u7528\u8f66\u52a0\u6cb9"
    if re.search(r"\b(parking|park)\b", text):
        return "\u505c\u8f66\u8d39"
    if re.search(r"\btoll\b", text):
        return "\u8fc7\u8def\u8d39"
    if is_korea_other_text(text):
        return "ESIM\u6d41\u91cf\u865a\u62df\u5361"
    if re.search(r"\busb\b", text):
        return "\u8d2d\u4e70USB"
    if category == "materials":
        return "\u7269\u6599\u91c7\u8d2d"
    if category == "physical_exam":
        return "\u4f53\u68c0\u8d39"
    if category == "lodging":
        return "\u4f4f\u5bbf\u8d39"
    if category == "nucleic_test":
        return "\u6838\u9178\u68c0\u6d4b\u8d39"
    if category == "meals":
        return "\u9910\u8d39"
    if category == "courier":
        return "\u5feb\u9012\u8d39"
    if category == "consumables":
        return "\u6d88\u8017\u54c1"
    if category == "welfare":
        return "\u798f\u5229\u8d39"
    return item.purpose.strip() or item.details.strip() or item.receipt_label.strip() or item.filename


def korea_detail_location(item: "ReceiptItem") -> str:
    text = item_search_text(item)
    currency = normalize_currency(item.currency, "KRW")
    if (
        currency == "USD"
        or korea_invoice_kind_for_item(item) in {"car_rental", "fuel", "parking", "esim", "materials"}
        or re.search(r"\b(usa|u\.s\.a|united states|ga|mi|il|ca|ny|atlanta|duluth|utica|chicago)\b", text)
    ):
        return "\u7f8e\u56fd"
    if any(token in text for token in ("korea", "\ud55c\uad6d", "\ub300\ud55c\ubbfc\uad6d")):
        return "\u97e9\u56fd"
    if any(token in text for token in ("china", "macao", "macau", "\u4e2d\u56fd", "\u6fb3\u95e8")):
        return "\u4e2d\u56fd"
    return item.place.strip()


def korea_payment_method_text(item: "ReceiptItem") -> str:
    raw = (item.payment_method or "").strip()
    text = raw.lower()
    if "master" in text:
        return "MASTERCARD"
    if "visa" in text:
        return "VISA"
    if "wechat" in text or "weixin" in text or "\u5fae\u4fe1" in text:
        return "wechat"
    if "credit" in text:
        return "Credit Card"
    if text == "card":
        return "Credit Card"
    return raw


def korea_receipt_block_is_wide(item: "ReceiptItem") -> bool:
    text = item_search_text(item)
    return korea_invoice_kind_for_item(item) == "car_rental" or "national" in text or len(ensure_receipt_images(item)) > 1


def report_category_for_item(item: Any, form_version: str) -> str:
    version = normalized_form_version(form_version)
    category = category_value_to_key(str(getattr(item, "category", "") or ""), version)
    if version == "Korea":
        searchable = item_search_text(item)
        if is_korea_other_text(searchable):
            return "other"
        if category in KOREA_CATEGORY_COLUMNS:
            return category
        if category in {"advertising", "office"}:
            return "materials"
        if category == "entertainment":
            return "meals"
        return "other"
    if category in USA_CATEGORY_ROWS:
        return category
    if category in {"materials", "consumables"}:
        return "office"
    return "other"


def usa_export_category_for_item(item: Any) -> str:
    category = report_category_for_item(item, "USA")
    text = item_search_text(item)
    if looks_like_usa_fuel_receipt(item):
        return "transportation"
    if looks_like_usa_parking_receipt(item):
        return "transportation"
    if category == "other" and re.search(r"\b(parking|park|toll|ticket|garage|transport|fuel|gas|flight|airline|airport|ride|taxi|lyft|uber)\b", text):
        return "transportation"
    if category == "other":
        return "transportation"
    return category


def expected_report_year() -> int:
    return date.today().year


def date_year_needs_review(value: Any, expected_year: Optional[int] = None) -> bool:
    parsed = parse_date_value(value)
    if isinstance(parsed, date):
        return parsed.year != (expected_year or expected_report_year())
    return False


def date_sort_value(value: str, expected_year: Optional[int] = None) -> float:
    parsed = parse_date_value(value)
    if isinstance(parsed, date):
        year = expected_year or expected_report_year()
        try:
            return float(parsed.replace(year=year).toordinal())
        except ValueError:
            return float((parsed.month * 32) + parsed.day)
    return float("inf")


def assumed_report_date_key(value: Any, expected_year: Optional[int] = None) -> str:
    parsed = parse_date_value(value)
    if isinstance(parsed, date):
        year = expected_year or expected_report_year()
        return f"{year:04d}-{parsed.month:02d}-{parsed.day:02d}"
    return str(value or "").strip()


def same_month_day_values(left: Any, right: Any) -> bool:
    left_value = parse_date_value(left)
    right_value = parse_date_value(right)
    if isinstance(left_value, date) and isinstance(right_value, date):
        return (left_value.month, left_value.day) == (right_value.month, right_value.day)
    return bool(left and right and str(left).strip() == str(right).strip())


def sort_receipt_list_for_review(items: List[Any]) -> List[Any]:
    indexed = [
        (
            0 if date_year_needs_review(getattr(item, "date", "")) else 1,
            index,
            item,
        )
        for index, item in enumerate(items)
    ]
    return [entry[2] for entry in sorted(indexed, key=lambda entry: (entry[0], entry[1]))]


def sort_receipts_for_report(items: List[Any], form_version: str) -> List[Any]:
    version = normalized_form_version(form_version)
    order = KOREA_INVOICE_KIND_ORDER if version == "Korea" else USA_REPORT_CATEGORY_ORDER
    order_index = {category: index for index, category in enumerate(order)}
    indexed = [
        (
            order_index.get(korea_invoice_kind_for_item(item) if version == "Korea" else report_category_for_item(item, version), len(order)),
            date_sort_value(str(getattr(item, "date", "") or "")),
            index,
            item,
        )
        for index, item in enumerate(items)
    ]
    return [entry[3] for entry in sorted(indexed, key=lambda entry: (entry[0], entry[1], entry[2]))]


def is_usa_receipt_item(item: "ReceiptItem") -> bool:
    return normalize_document_type(getattr(item, "document_type", "receipt"), "receipt") == "receipt"


def is_usa_payment_proof_item(item: "ReceiptItem") -> bool:
    return normalize_document_type(getattr(item, "document_type", "unknown"), "unknown") == "payment_proof"


def looks_like_usa_merchant_receipt_text(*values: Any) -> bool:
    text = " ".join(str(value or "") for value in values).lower()
    if re.search(r"\b(credit card receipt|sales receipt|merchant receipt|customer receipt)\b", text):
        return True
    if re.search(r"\b(payment method|card type|authorization|terminal id|receipt\s*:|invoice\s*:)\b", text) and re.search(
        r"\b(total|subtotal|tax|amount|paid)\b",
        text,
    ):
        return True
    if re.search(r"\b(receipt|invoice)\b", text) and re.search(r"\b(total|payment method|card type|thank you)\b", text):
        return True
    return False


def usa_payment_proof_text_score(*values: Any) -> int:
    text = " ".join(str(value or "") for value in values).lower()
    score = 0
    if re.search(r"\b(chase|jpmorgan|sapphire reserve|freedom)\b", text):
        score += 2
    if re.search(r"\b(transaction details|transaction date|posted date|printed from chase personal online)\b", text):
        score += 2
    if re.search(r"\b(merchant type|card number|account activity|card statement|bank statement|proof of payment|payment proof)\b", text):
        score += 2
    if re.search(r"\b(rewards earned|ultimate rewards|sale)\b", text):
        score += 1
    return score


def looks_like_usa_payment_proof_text(*values: Any) -> bool:
    if looks_like_usa_merchant_receipt_text(*values):
        return False
    return usa_payment_proof_text_score(*values) >= 3


def looks_like_usa_workbook_proof_name(*values: Any) -> bool:
    text = " ".join(str(value or "") for value in values).lower()
    return bool(
        re.search(r"\bworkbook[_\s-]*row\d+[_\s-]*proof[_\s-]*e[_\s-]", text)
        or re.search(r"[_\s-]proof[_\s-]*e[_\s-]*20\d{2}[-_]\d{1,2}[-_]\d{1,2}", text)
    )


def parse_usa_support_metadata_from_text(*values: Any) -> Tuple[str, str]:
    text = " ".join(str(value or "") for value in values)
    patterns = [
        r"(20\d{2})[-_](\d{1,2})[-_](\d{1,2})[_\s-]+(-?\d+(?:\.\d+)?)",
        r"(20\d{2})(\d{2})(\d{2})[_\s-]+(-?\d+(?:\.\d+)?)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if not match:
            continue
        year, month, day, amount = match.groups()
        try:
            parsed_date = date(int(year), int(month), int(day)).isoformat()
        except ValueError:
            parsed_date = ""
        return parsed_date, format_amount(safe_float(amount))
    return "", ""


def apply_usa_workbook_proof_metadata(item: Any, *extra_values: Any) -> None:
    values = (
        getattr(item, "filename", ""),
        getattr(item, "path", ""),
        getattr(item, "source_path", ""),
        getattr(item, "source_page", ""),
        *extra_values,
    )
    if not looks_like_usa_workbook_proof_name(*values):
        return
    parsed_date, parsed_amount = parse_usa_support_metadata_from_text(*values)
    if parsed_date:
        item.date = parsed_date
    if parsed_amount:
        item.amount = parsed_amount
    item.document_type = "payment_proof"


def looks_like_usa_receipt_support_text(*values: Any) -> bool:
    text = " ".join(str(value or "") for value in values).lower()
    if looks_like_usa_payment_proof_text(text):
        return False
    return bool(
        re.search(
            r"\b(receipt|ticket|invoice|merchant|parking|park|toll|gas|fuel|hotel|restaurant|meal|entry|exit|qr|barcode)\b",
            text,
        )
    )


def looks_like_usa_amountless_support_text(*values: Any) -> bool:
    text = " ".join(str(value or "") for value in values).lower()
    if looks_like_usa_payment_proof_text(text):
        return False
    return bool(
        re.search(
            r"\b(ticket|entry|exit|qr|barcode|validation|information|scan|lost ticket|parking time)\b",
            text,
        )
    )


def usa_receipt_support_without_amount(item: "ReceiptItem") -> bool:
    if not is_usa_receipt_item(item):
        return False
    if safe_float(item.amount) is not None or safe_float(item.krw_amount) is not None or safe_float(item.rmb_amount) is not None:
        return False
    return looks_like_usa_amountless_support_text(
        item.filename,
        item.place,
        item.purpose,
        item.details,
        item.payment_method,
        item.receipt_label,
    )


def auto_classify_usa_document_type(item: "ReceiptItem", *extra_values: Any) -> None:
    current = normalize_document_type(item.document_type, "unknown")
    values = (
        item.filename,
        item.place,
        item.purpose,
        item.details,
        item.payment_method,
        item.receipt_label,
        *extra_values,
    )
    if looks_like_usa_workbook_proof_name(*values):
        apply_usa_workbook_proof_metadata(item, *extra_values)
        return
    if looks_like_usa_merchant_receipt_text(*values):
        item.document_type = "receipt"
        return
    if looks_like_usa_payment_proof_text(*values):
        item.document_type = "payment_proof"
    elif current != "unknown":
        item.document_type = current
    elif looks_like_usa_receipt_support_text(*values):
        item.document_type = "receipt"
    else:
        item.document_type = "unknown"


def receipt_attention_reasons(item: "ReceiptItem", form_version: str) -> List[str]:
    version = normalized_form_version(form_version)
    reasons: List[str] = []
    if item.status == "Failed":
        reasons.append("AI failed")
    if version == "USA" and normalize_document_type(item.document_type, "unknown") == "unknown":
        reasons.append("Unknown document type")
    if not item.date.strip():
        reasons.append("Missing date")
    elif date_year_needs_review(item.date):
        reasons.append("Wrong year")
    if not (item.amount.strip() or item.krw_amount.strip() or item.rmb_amount.strip()) and not (
        version == "USA" and usa_receipt_support_without_amount(item)
    ):
        reasons.append("Missing amount")
    if is_usa_receipt_item(item) or version == "Korea":
        if not (item.details.strip() or item.purpose.strip() or item.receipt_label.strip()):
            reasons.append("Missing details")
        if not item.place.strip():
            reasons.append("Missing vendor")
    return reasons


def receipt_attention_reason(item: "ReceiptItem", form_version: str) -> Optional[str]:
    reasons = receipt_attention_reasons(item, form_version)
    return reasons[0] if reasons else None


def payment_proof_missing_status(item: Any) -> Optional[str]:
    missing_date = not str(getattr(item, "date", "") or "").strip()
    missing_amount = safe_float(getattr(item, "amount", "")) is None
    if missing_date and missing_amount:
        return "Missing date/amount"
    if missing_date:
        return "Missing date"
    if missing_amount:
        return "Missing amount"
    return None


def receipt_needs_ai(item: "ReceiptItem", form_version: str) -> bool:
    version = normalized_form_version(form_version)
    if item.status in {"Remembered", "Manual remembered"}:
        return False
    if item.status in {"Empty", "Failed", "Needs regen"}:
        return True
    if not item.date.strip():
        return True
    if not item.amount.strip() and not (version == "USA" and usa_receipt_support_without_amount(item)):
        return True
    if version == "USA":
        document_type = normalize_document_type(item.document_type, "unknown")
        if document_type == "unknown":
            return True
        if document_type == "receipt" and not (
            item.purpose.strip() or item.details.strip() or item.receipt_label.strip()
        ):
            return True
    return False


def usa_receipt_merge_key(item: "ReceiptItem") -> Optional[Tuple[str, float]]:
    amount = safe_float(item.amount)
    if amount is None:
        return None
    date_key = assumed_report_date_key(item.date)
    if not date_key:
        return None
    return date_key, round(amount, 2)


def usa_receipt_date_score(item: "ReceiptItem") -> int:
    parsed = parse_date_value(item.date)
    if isinstance(parsed, date):
        return 2 if parsed.year == date.today().year else 1
    return 0


def usa_same_transaction_despite_bad_year(left: "ReceiptItem", right: "ReceiptItem") -> bool:
    left_amount = safe_float(left.amount)
    right_amount = safe_float(right.amount)
    if left_amount is None or right_amount is None or abs(left_amount - right_amount) > 0.01:
        return False
    left_date = parse_date_value(left.date)
    right_date = parse_date_value(right.date)
    if not isinstance(left_date, date) or not isinstance(right_date, date):
        return False
    current_year = date.today().year
    if not ({left_date.year, right_date.year} & {current_year}):
        return False
    if left_date.year == right_date.year:
        return False
    try:
        normalized_left = left_date.replace(year=current_year)
        normalized_right = right_date.replace(year=current_year)
    except ValueError:
        return False
    if abs((normalized_left - normalized_right).days) > 3:
        return False
    left_tokens = usa_match_tokens(left)
    right_tokens = usa_match_tokens(right)
    if len(left_tokens & right_tokens) >= 1 and (
        looks_like_usa_fuel_receipt(left) or looks_like_usa_fuel_receipt(right)
    ):
        return True
    left_place = re.sub(r"[^a-z0-9]+", "", str(left.place or "").lower())
    right_place = re.sub(r"[^a-z0-9]+", "", str(right.place or "").lower())
    return bool(left_place and right_place and left_place == right_place)


def merge_usa_receipt_item_into(target: "ReceiptItem", source: "ReceiptItem") -> None:
    merge_attachment_lists(ensure_receipt_images(target), ensure_receipt_images(source))
    for attr in (
        "place",
        "purpose",
        "details",
        "project_number",
        "category",
        "payment_method",
        "receipt_label",
        "rmb_amount",
    ):
        if not str(getattr(target, attr, "") or "").strip() and str(getattr(source, attr, "") or "").strip():
            setattr(target, attr, getattr(source, attr))
    if target.status not in {"Failed", "Processing"}:
        target.status = "Merged"


def same_expense_date_values(left: str, right: str) -> bool:
    return same_month_day_values(left, right)


USA_MATCH_STOP_WORDS = {
    "all",
    "category",
    "central",
    "class",
    "date",
    "deck",
    "drive",
    "entry",
    "exit",
    "from",
    "information",
    "initial",
    "indiana",
    "lane",
    "lower",
    "method",
    "number",
    "page",
    "paid",
    "receipt",
    "payment",
    "proof",
    "transaction",
    "details",
    "total",
    "sale",
    "card",
    "credit",
    "debit",
    "visa",
    "mastercard",
    "road",
    "safely",
    "save",
    "terminal",
    "ticket",
    "toll",
    "to",
    "transportation",
    "using",
    "vehicle",
    "website",
    "jpg",
    "jpeg",
    "png",
    "pdf",
}


def usa_match_tokens(item: "ReceiptItem") -> set[str]:
    text = item_search_text(item)
    tokens = set()
    for token in re.findall(r"[a-z0-9]+", text.lower()):
        if len(token) < 3 or token in USA_MATCH_STOP_WORDS:
            continue
        if token.isdigit():
            continue
        digits = re.sub(r"\D", "", token)
        if digits:
            continue
        tokens.add(token)
    return tokens


def usa_identifier_is_date_or_time(digits: str) -> bool:
    if len(digits) == 8:
        if digits[:4] in {"2025", "2026", "2027"} or digits[-4:] in {"2025", "2026", "2027"}:
            return True
    if len(digits) == 6:
        first = int(digits[:2])
        second = int(digits[2:4])
        third = int(digits[4:6])
        if first <= 23 and second <= 59 and third <= 59:
            return True
    return False


def usa_identifier_tokens(item: "ReceiptItem") -> set[str]:
    text = item_search_text(item).lower()
    identifiers = set()
    for token in re.findall(r"[a-z0-9]*\d[a-z0-9]*", text):
        digits = re.sub(r"\D", "", token)
        if len(digits) >= 5 and not usa_identifier_is_date_or_time(digits):
            identifiers.add(token)
    return identifiers


def usa_amountless_receipt_match(item: "ReceiptItem", candidates: List["ReceiptItem"]) -> Optional["ReceiptItem"]:
    same_date = [
        candidate
        for candidate in candidates
        if is_usa_receipt_item(candidate)
        and safe_float(candidate.amount) is not None
        and same_expense_date_values(item.date, candidate.date)
    ]
    if not same_date:
        return None
    item_identifiers = usa_identifier_tokens(item)
    if item_identifiers:
        identifier_matches = [
            candidate
            for candidate in same_date
            if item_identifiers & usa_identifier_tokens(candidate)
        ]
        if len(identifier_matches) == 1:
            return identifier_matches[0]
    item_tokens = usa_match_tokens(item)
    if not item_tokens:
        return None
    scored: List[Tuple[int, ReceiptItem]] = []
    for candidate in same_date:
        overlap = len(item_tokens & usa_match_tokens(candidate))
        if overlap:
            scored.append((overlap, candidate))
    if not scored:
        return None
    scored.sort(key=lambda entry: entry[0], reverse=True)
    if scored[0][0] < 2 and not (scored[0][0] == 1 and len(scored) == 1):
        return None
    if len(scored) > 1 and scored[0][0] == scored[1][0]:
        return None
    return scored[0][1]


def attachment_looks_like_usa_payment_proof(attachment: Dict[str, Any]) -> bool:
    values = (
        attachment.get("filename", ""),
        attachment.get("path", ""),
        attachment.get("source_path", ""),
        attachment.get("source_page", ""),
    )
    return looks_like_usa_workbook_proof_name(*values) or looks_like_usa_payment_proof_text(*values)


def receipt_item_from_payment_proof_attachment(
    attachment: Dict[str, Any],
    parent: "ReceiptItem",
) -> "ReceiptItem":
    parsed_date, parsed_amount = parse_usa_support_metadata_from_text(
        attachment.get("filename", ""),
        attachment.get("path", ""),
        attachment.get("source_path", ""),
        attachment.get("source_page", ""),
    )
    path = str(attachment.get("path") or "")
    filename = str(attachment.get("filename") or Path(path).name or parent.filename)
    proof = ReceiptItem(
        item_id=str(attachment.get("id") or uuid.uuid4().hex),
        path=path,
        filename=filename,
        source_path=str(attachment.get("source_path") or path),
        source_page=str(attachment.get("source_page") or ""),
        date=parsed_date,
        place="",
        amount=parsed_amount,
        currency="USD",
        purpose="",
        details="",
        project_number=getattr(parent, "project_number", ""),
        category=getattr(parent, "category", "transportation") or "transportation",
        payment_method=getattr(parent, "payment_method", ""),
        receipt_label=filename,
        document_type="payment_proof",
        status="Payment proof",
        crop_box=copy.deepcopy(attachment.get("crop_box")),
        rotation_degrees=normalize_rotation(attachment.get("rotation_degrees", 0)),
        receipt_images=[copy.deepcopy(attachment)],
    )
    return proof


def split_embedded_usa_payment_proofs(items: List["ReceiptItem"]) -> List["ReceiptItem"]:
    split_items: List[ReceiptItem] = []
    for item in items:
        auto_classify_usa_document_type(item)
        images = ensure_receipt_images(item)
        if is_usa_payment_proof_item(item):
            split_items.append(item)
            continue
        proof_images = [attachment for attachment in images if attachment_looks_like_usa_payment_proof(attachment)]
        if not proof_images:
            split_items.append(item)
            continue
        receipt_images = [
            attachment
            for attachment in images
            if not attachment_looks_like_usa_payment_proof(attachment)
        ]
        if receipt_images:
            item.receipt_images = receipt_images
            ensure_receipt_images(item)
            split_items.append(item)
        else:
            item.receipt_images = proof_images
            ensure_receipt_images(item)
            apply_usa_workbook_proof_metadata(item)
            parsed_date, parsed_amount = parse_usa_support_metadata_from_text(
                proof_images[0].get("filename", ""),
                proof_images[0].get("path", ""),
                proof_images[0].get("source_path", ""),
                proof_images[0].get("source_page", ""),
            )
            if parsed_date:
                item.date = parsed_date
            if parsed_amount:
                item.amount = parsed_amount
            item.document_type = "payment_proof"
            item.status = payment_proof_missing_status(item) or "Payment proof"
            split_items.append(item)
            continue
        for attachment in proof_images:
            split_items.append(receipt_item_from_payment_proof_attachment(attachment, item))
    return split_items


def merge_same_usa_receipt_items(items: List["ReceiptItem"]) -> Tuple[List["ReceiptItem"], Dict[str, str], int]:
    merged_count = 0
    merged_items: List[ReceiptItem] = []
    seen: Dict[Tuple[str, float], ReceiptItem] = {}
    removed_to_target: Dict[str, str] = {}
    for item in items:
        auto_classify_usa_document_type(item)
        normalize_usa_fuel_receipt_fields(item)
        normalize_usa_parking_receipt_fields(item)
        if not is_usa_receipt_item(item):
            merged_items.append(item)
            continue
        auto_rotate_landscape_receipt_images(item)
        ensure_receipt_images(item)
        key = usa_receipt_merge_key(item)
        if key is None or key not in seen:
            if key is not None:
                seen[key] = item
            merged_items.append(item)
            continue
        target = seen[key]
        merge_usa_receipt_item_into(target, item)
        removed_to_target[item.item_id] = target.item_id
        merged_count += 1
    if merged_items:
        retained_exact: List[ReceiptItem] = []
        for item in merged_items:
            if not is_usa_receipt_item(item) or safe_float(item.amount) is None:
                retained_exact.append(item)
                continue
            target_index = next(
                (
                    index
                    for index, candidate in enumerate(retained_exact)
                    if is_usa_receipt_item(candidate)
                    and usa_same_transaction_despite_bad_year(candidate, item)
                ),
                None,
            )
            if target_index is None:
                retained_exact.append(item)
                continue
            target = retained_exact[target_index]
            keeper = item if usa_receipt_date_score(item) > usa_receipt_date_score(target) else target
            duplicate = target if keeper is item else item
            if keeper is item:
                retained_exact[target_index] = keeper
            merge_usa_receipt_item_into(keeper, duplicate)
            removed_to_target[duplicate.item_id] = keeper.item_id
            merged_count += 1
        merged_items = retained_exact
    if merged_items:
        retained: List[ReceiptItem] = []
        paid_receipts = [
            item
            for item in merged_items
            if is_usa_receipt_item(item) and safe_float(item.amount) is not None
        ]
        for item in merged_items:
            if not is_usa_receipt_item(item) or safe_float(item.amount) is not None:
                retained.append(item)
                continue
            target = usa_amountless_receipt_match(item, paid_receipts)
            if target is None:
                retained.append(item)
                continue
            merge_usa_receipt_item_into(target, item)
            removed_to_target[item.item_id] = target.item_id
            merged_count += 1
        merged_items = retained
    if removed_to_target:
        for item in merged_items:
            if getattr(item, "matched_receipt_id", "") in removed_to_target:
                item.matched_receipt_id = removed_to_target[item.matched_receipt_id]
    return merged_items, removed_to_target, merged_count


def match_usa_payment_proof_items(items: List["ReceiptItem"]) -> None:
    receipts = [item for item in items if is_usa_receipt_item(item)]
    for proof_item in items:
        if not is_usa_payment_proof_item(proof_item):
            continue
        proof_amount = safe_float(proof_item.amount)
        missing_status = payment_proof_missing_status(proof_item)
        if missing_status:
            proof_item.matched_receipt_id = ""
            proof_item.status = missing_status
            continue
        matches: List[ReceiptItem] = []
        for receipt in receipts:
            receipt_amount = safe_float(receipt.amount)
            if receipt_amount is None:
                continue
            if abs(receipt_amount - proof_amount) > 0.01:
                continue
            if not same_expense_date_values(receipt.date, proof_item.date):
                continue
            matches.append(receipt)
        if len(matches) == 1:
            proof_item.matched_receipt_id = matches[0].item_id
            proof_item.status = "Matched"
        else:
            proof_item.matched_receipt_id = ""
            proof_item.status = "Multiple possible matches" if len(matches) > 1 else "Payment proof"


def match_bank_statement_to_receipt(bank_item: "BankStatementItem", receipts: List["ReceiptItem"]) -> None:
    bank_amount = safe_float(bank_item.amount)
    missing_status = payment_proof_missing_status(bank_item)
    if missing_status:
        bank_item.matched_receipt_id = ""
        bank_item.status = missing_status
        return
    matches: List[ReceiptItem] = []
    for receipt in receipts:
        receipt_amount = safe_float(receipt.amount)
        if receipt_amount is None:
            continue
        if abs(receipt_amount - bank_amount) > 0.01:
            continue
        if not same_expense_date_values(receipt.date, bank_item.date):
            continue
        matches.append(receipt)
    if len(matches) == 1:
        bank_item.matched_receipt_id = matches[0].item_id
        bank_item.status = "Matched"
    else:
        bank_item.matched_receipt_id = ""
        bank_item.status = "Multiple possible matches" if len(matches) > 1 else "Payment proof"


def build_usa_report_rows(
    items: List["ReceiptItem"],
    bank_items: Optional[List["BankStatementItem"]] = None,
) -> List[Dict[str, Any]]:
    working_items, removed_to_target, _merged_count = merge_same_usa_receipt_items(
        split_embedded_usa_payment_proofs(copy.deepcopy(list(items)))
    )
    for bank_item in bank_items or []:
        if bank_item.matched_receipt_id in removed_to_target:
            bank_item.matched_receipt_id = removed_to_target[bank_item.matched_receipt_id]
    match_usa_payment_proof_items(working_items)
    receipt_items = sort_receipts_for_report(
        [item for item in working_items if is_usa_receipt_item(item)],
        "USA",
    )
    rows: List[Dict[str, Any]] = []
    for item in receipt_items:
        rows.append(
            {
                "id": item.item_id,
                "item": item,
                "expense_item": item,
                "date": item.date,
                "amount": item.amount,
                "receipt_images": ensure_receipt_images(item),
                "proof_images": [],
            }
        )
    row_by_receipt_id = {str(row["id"]): row for row in rows}
    proof_candidates: List[Dict[str, Any]] = []
    for item in working_items:
        if is_usa_payment_proof_item(item):
            proof_candidates.append(
                {
                    "id": item.item_id,
                    "date": item.date,
                    "amount": item.amount,
                    "matched_receipt_id": item.matched_receipt_id,
                    "expense_item": item,
                    "images": ensure_receipt_images(item),
                }
            )
    for bank_item in bank_items or []:
        if not bank_item.matched_receipt_id:
            match_bank_statement_to_receipt(bank_item, receipt_items)
        proof_candidates.append(
            {
                "id": bank_item.item_id,
                "date": bank_item.date,
                "amount": bank_item.amount,
                "matched_receipt_id": bank_item.matched_receipt_id,
                "expense_item": ReceiptItem(
                    item_id=bank_item.item_id,
                    path=bank_item.path,
                    filename=bank_item.filename,
                    source_path=bank_item.source_path,
                    source_page=bank_item.source_page,
                    date=bank_item.date,
                    place=bank_item.place,
                    amount=bank_item.amount,
                    currency="USD",
                    purpose="",
                    details=bank_item.place,
                    category="transportation",
                    payment_method="",
                    receipt_label=bank_item.filename,
                    document_type="payment_proof",
                    status=bank_item.status,
                ),
                "images": [attachment_from_item(bank_item)],
            }
        )

    def proof_source_key(images: List[Dict[str, Any]]) -> str:
        if not images:
            return uuid.uuid4().hex
        first = images[0]
        source = str(first.get("source_path") or first.get("path") or "")
        try:
            return str(Path(source).resolve()).lower()
        except Exception:
            return source.lower()

    proof_only_rows: List[Dict[str, Any]] = []
    for proof in proof_candidates:
        matched_id = str(proof.get("matched_receipt_id") or "")
        target_row = row_by_receipt_id.get(matched_id)
        proof_images = list(proof.get("images") or [])
        if target_row is not None:
            source_key = proof_source_key(proof_images)
            existing_source_key = str(target_row.get("_proof_source_key") or "")
            if not existing_source_key or existing_source_key == source_key:
                target_row["_proof_source_key"] = source_key
                target_row["proof_images"].extend(proof_images)
                continue
            proof_only_rows.append(
                {
                    "id": str(proof.get("id") or uuid.uuid4().hex),
                    "item": None,
                    "expense_item": None,
                    "date": str(proof.get("date") or ""),
                    "amount": str(proof.get("amount") or ""),
                    "receipt_images": [],
                    "proof_images": proof_images,
                }
            )
            continue
        proof_only_rows.append(
            {
                "id": str(proof.get("id") or uuid.uuid4().hex),
                "item": None,
                "expense_item": proof.get("expense_item"),
                "date": str(proof.get("date") or ""),
                "amount": str(proof.get("amount") or ""),
                "receipt_images": [],
                "proof_images": proof_images,
            }
        )
    return rows + proof_only_rows


def sort_usa_support_rows_by_date(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    indexed = [
        (
            date_sort_value(str(row.get("date") or "")),
            index,
            row,
        )
        for index, row in enumerate(rows)
    ]
    return [entry[2] for entry in sorted(indexed, key=lambda entry: (entry[0], entry[1]))]


def normalize_rotation(value: Any) -> int:
    try:
        degrees = int(float(value))
    except (TypeError, ValueError):
        degrees = 0
    return degrees % 360 // 90 * 90


def apply_primary_receipt_rotation_hint(item: "ReceiptItem", value: Any) -> bool:
    rotation = normalize_rotation(value)
    if not rotation:
        return False
    images = ensure_receipt_images(item)
    if not images:
        return False
    primary = images[0]
    if normalize_rotation(primary.get("rotation_degrees", 0)):
        return False
    primary["rotation_degrees"] = rotation
    item.rotation_degrees = rotation
    item.crop_box = primary.get("crop_box")
    return True


def attachment_needs_landscape_rotation(attachment: Dict[str, Any]) -> bool:
    if Image is None:
        return False
    if normalize_rotation(attachment.get("rotation_degrees", 0)):
        return False
    if attachment.get("crop_box"):
        return False
    path = Path(str(attachment.get("path") or ""))
    if not path.exists():
        return False
    try:
        with Image.open(path) as source:
            image = ImageOps.exif_transpose(source) if ImageOps is not None else source
            width, height = image.size
    except Exception:
        return False
    return width > max(1, int(height * EXPORT_AUTO_ROTATE_LANDSCAPE_RATIO))


def auto_rotate_landscape_receipt_images(item: "ReceiptItem") -> bool:
    if not is_usa_receipt_item(item):
        return False
    changed = False
    images = ensure_receipt_images(item)
    for attachment in images:
        if attachment_needs_landscape_rotation(attachment):
            attachment["rotation_degrees"] = 90
            changed = True
    if changed and images:
        primary = images[0]
        item.rotation_degrees = normalize_rotation(primary.get("rotation_degrees", 0))
        item.crop_box = primary.get("crop_box")
    return changed


def oriented_image_from_path(path: Path, rotation_degrees: Any = 0) -> Any:
    with Image.open(path) as source:
        image = ImageOps.exif_transpose(source) if ImageOps is not None else source.copy()
    rotation = normalize_rotation(rotation_degrees)
    if rotation:
        image = image.rotate(-rotation, expand=True)
    return image


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


def flatten_crop_points(points: List[Tuple[float, float]]) -> List[float]:
    return [round(value, 2) for point in points for value in point]


def rotated_crop_points(
    crop_box: Any,
    width: int,
    height: int,
    delta_degrees: int,
) -> Optional[List[float]]:
    points = normalized_crop_points(crop_box, width, height)
    if not points:
        return None
    delta = normalize_rotation(delta_degrees)
    if delta == 90:
        transformed = [(height - y, x) for x, y in points]
        ordered = [transformed[index] for index in (3, 0, 1, 2)]
    elif delta == 270:
        transformed = [(y, width - x) for x, y in points]
        ordered = [transformed[index] for index in (1, 2, 3, 0)]
    elif delta == 180:
        transformed = [(width - x, height - y) for x, y in points]
        ordered = [transformed[index] for index in (2, 3, 0, 1)]
    else:
        ordered = points
    return flatten_crop_points(ordered)


def prepare_receipt_image_file(path: Path, crop_box: Any = None, rotation_degrees: Any = 0) -> Path:
    if Image is None:
        return path
    target_dir = WORK_DIR / "prepared_images"
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / f"{uuid.uuid4().hex}.png"
    image = oriented_image_from_path(path, rotation_degrees)
    image = perspective_crop_image(image, crop_box)
    if image.mode not in {"RGB", "RGBA"}:
        image = image.convert("RGB")
    image.save(target, "PNG")
    return target


def prepare_excel_image_file(path: Path, crop_box: Any = None, rotation_degrees: Any = 0) -> Path:
    return prepare_receipt_image_file(path, crop_box, rotation_degrees)


def resize_excel_image(
    path: Path,
    max_width: int,
    max_height: int,
    crop_box: Any = None,
    rotation_degrees: Any = 0,
) -> Any:
    excel_path = prepare_excel_image_file(path, crop_box, rotation_degrees)
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


def attachment_group_filename(item: ReceiptItem, attachments: List[Dict[str, Any]]) -> str:
    if not attachments:
        return item.filename
    source_paths = {
        str(attachment.get("source_path") or "").strip()
        for attachment in attachments
        if str(attachment.get("source_path") or "").strip()
    }
    if len(source_paths) == 1:
        source_path = next(iter(source_paths))
        if source_path.lower().endswith(".pdf"):
            page_count = len(attachments)
            source_name = Path(source_path).name
            if page_count > 1:
                return f"{source_name} ({page_count} pages)"
            source_page = str(attachments[0].get("source_page") or "").strip()
            return f"{source_name} page {source_page or '1'}"
    return str(attachments[0].get("filename") or item.filename or Path(str(attachments[0].get("path") or "")).name)


def ensure_receipt_images(item: ReceiptItem) -> List[Dict[str, Any]]:
    normalized = [attachment for attachment in (normalize_attachment(raw) for raw in item.receipt_images) if attachment]
    if not normalized and item.path:
        normalized = [attachment_from_item(item)]
    item.receipt_images = normalized
    if normalized:
        primary = normalized[0]
        item.path = primary["path"]
        item.filename = attachment_group_filename(item, normalized)
        item.source_path = primary.get("source_path", "")
        item.source_page = primary.get("source_page", "")
        item.crop_box = primary.get("crop_box")
        item.rotation_degrees = normalize_rotation(primary.get("rotation_degrees", 0))
    return item.receipt_images


def attachment_file_key(attachment: Dict[str, Any]) -> Tuple[str, str]:
    return selected_file_key(
        Path(str(attachment.get("path") or "")),
        str(attachment.get("source_path") or ""),
        str(attachment.get("source_page") or ""),
    )


def merge_attachment_lists(
    target: List[Dict[str, Any]],
    source: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    seen = {attachment_file_key(attachment) for attachment in target}
    for attachment in source:
        key = attachment_file_key(attachment)
        if key in seen:
            continue
        target.append(copy.deepcopy(attachment))
        seen.add(key)
    return target


def prepare_attachment_image_file(attachment: Dict[str, Any]) -> Path:
    return prepare_receipt_image_file(
        Path(str(attachment.get("path") or "")),
        attachment.get("crop_box"),
        attachment.get("rotation_degrees", 0),
    )


def excel_ready_rgb(image: Any) -> Any:
    if Image is None:
        return image
    if image.mode == "RGB":
        return image
    if image.mode == "RGBA":
        background = Image.new("RGB", image.size, "white")
        background.paste(image, mask=image.getchannel("A"))
        return background
    return image.convert("RGB")


def save_excel_ready_image(image: Any, target_dir: Path, suffix: str = "contact") -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    target = target_dir / f"{uuid.uuid4().hex}_{suffix}.jpg"
    output = excel_ready_rgb(image)
    output.save(target, "JPEG", quality=EXCEL_IMAGE_JPEG_QUALITY, subsampling=0)
    return target


def attachment_image_for_export(attachment: Dict[str, Any]) -> Optional[Any]:
    if Image is None:
        return None
    path = Path(str(attachment.get("path") or ""))
    if not path.exists():
        return None
    try:
        image = oriented_image_from_path(path, attachment.get("rotation_degrees", 0))
        image = perspective_crop_image(image, attachment.get("crop_box"))
        return excel_ready_rgb(image)
    except Exception:
        log_exception(f"Could not prepare image for Excel export: {path}")
        return None


def downscale_for_excel_processing(
    image: Any,
    max_width: int,
    max_height: int,
    allow_upscale: bool,
    compact_whitespace: bool,
    max_columns: int,
) -> Any:
    if Image is None:
        return image
    render_scale = EXCEL_IMAGE_RENDER_SCALE if allow_upscale else 1
    resampling = getattr(Image, "Resampling", None)
    resample = getattr(resampling, "LANCZOS", getattr(Image, "LANCZOS", Image.BICUBIC))
    if compact_whitespace and max_columns == 1:
        target_width = max(1, int(max_width * render_scale))
        if image.width > target_width:
            scale = target_width / max(1, image.width)
            return image.resize((target_width, max(1, int(image.height * scale))), resample)
        return image
    target_width = max(1, int(max_width * render_scale))
    target_height = max(1, int(max_height * render_scale))
    if image.width <= target_width and image.height <= target_height:
        return image
    copy_image = image.copy()
    copy_image.thumbnail((target_width, target_height), resample)
    return copy_image


def trim_white_image(image: Any, padding: int = 18, threshold: int = 245) -> Any:
    if Image is None:
        return image
    rgb = image.convert("RGB")
    pixels = rgb.load()
    left, top = rgb.width, rgb.height
    right, bottom = -1, -1
    for y in range(rgb.height):
        for x in range(rgb.width):
            r, g, b = pixels[x, y]
            if min(r, g, b) < threshold:
                left = min(left, x)
                top = min(top, y)
                right = max(right, x)
                bottom = max(bottom, y)
    if right < left or bottom < top:
        return image
    left = max(0, left - padding)
    top = max(0, top - padding)
    right = min(rgb.width - 1, right + padding)
    bottom = min(rgb.height - 1, bottom + padding)
    if left == 0 and top == 0 and right == rgb.width - 1 and bottom == rgb.height - 1:
        return image
    return image.crop((left, top, right + 1, bottom + 1))


def compact_vertical_whitespace(image: Any, padding: int = 10, gap: int = 12) -> Any:
    if Image is None:
        return image
    rgb = image.convert("RGB")
    pixels = rgb.load()
    min_dark_pixels = max(4, int(rgb.width * 0.0025))
    active_rows: List[int] = []
    for y in range(rgb.height):
        dark_count = 0
        for x in range(rgb.width):
            r, g, b = pixels[x, y]
            if min(r, g, b) < 245:
                dark_count += 1
                if dark_count >= min_dark_pixels:
                    active_rows.append(y)
                    break
    if not active_rows:
        return image
    bands: List[Tuple[int, int]] = []
    start = active_rows[0]
    previous = active_rows[0]
    merge_gap = max(12, int(rgb.height * 0.012))
    for y in active_rows[1:]:
        if y - previous > merge_gap:
            bands.append((max(0, start - padding), min(rgb.height, previous + padding + 1)))
            start = y
        previous = y
    bands.append((max(0, start - padding), min(rgb.height, previous + padding + 1)))
    merged: List[Tuple[int, int]] = []
    for top, bottom in bands:
        if merged and top - merged[-1][1] <= gap:
            merged[-1] = (merged[-1][0], bottom)
        else:
            merged.append((top, bottom))
    new_height = sum(bottom - top for top, bottom in merged) + gap * max(0, len(merged) - 1)
    if new_height >= int(rgb.height * 0.92):
        return image
    compacted = Image.new("RGB", (rgb.width, max(1, new_height)), "white")
    y_out = 0
    for index, (top, bottom) in enumerate(merged):
        segment = rgb.crop((0, top, rgb.width, bottom))
        compacted.paste(segment, (0, y_out))
        y_out += segment.height
        if index < len(merged) - 1:
            y_out += gap
    return compacted


def attachment_contact_sheet_token(attachment: Dict[str, Any]) -> Optional[Tuple[Any, ...]]:
    normalized = normalize_attachment(attachment)
    if not normalized:
        return None
    path = Path(str(normalized.get("path") or ""))
    try:
        resolved_path = str(path.resolve()).lower()
    except Exception:
        resolved_path = str(path).lower()
    try:
        stat = path.stat()
        file_state: Tuple[Any, ...] = (stat.st_mtime_ns, stat.st_size)
    except OSError:
        file_state = (None, None)
    crop_box = normalized.get("crop_box")
    crop_token: Optional[Tuple[float, ...]] = None
    if isinstance(crop_box, (list, tuple)):
        try:
            crop_token = tuple(round(float(value), 4) for value in crop_box)
        except (TypeError, ValueError):
            crop_token = None
    return (
        resolved_path,
        file_state,
        str(normalized.get("source_path") or "").lower(),
        str(normalized.get("source_page") or ""),
        crop_token,
        normalize_rotation(normalized.get("rotation_degrees", 0)),
    )


def attachment_contact_sheet_cache_key(
    attachments: List[Dict[str, Any]],
    max_width: int,
    max_height: int,
    allow_upscale: bool,
    stretch_tiles: bool,
    trim_whitespace: bool,
    compact_whitespace: bool,
    max_columns: int,
) -> Optional[Tuple[Any, ...]]:
    tokens = [attachment_contact_sheet_token(attachment) for attachment in attachments]
    if not tokens or any(token is None for token in tokens):
        return None
    return (
        "contact-sheet",
        tuple(tokens),
        int(max_width),
        int(max_height),
        bool(allow_upscale),
        bool(stretch_tiles),
        bool(trim_whitespace),
        bool(compact_whitespace),
        int(max_columns),
        EXCEL_IMAGE_RENDER_SCALE,
        EXCEL_IMAGE_JPEG_QUALITY,
    )


def get_cached_contact_sheet(cache_key: Optional[Tuple[Any, ...]]) -> Optional[Path]:
    if cache_key is None:
        return None
    with EXCEL_CONTACT_SHEET_CACHE_LOCK:
        cached = EXCEL_CONTACT_SHEET_CACHE.get(cache_key)
    if cached is None:
        return None
    if cached.exists():
        return cached
    with EXCEL_CONTACT_SHEET_CACHE_LOCK:
        if EXCEL_CONTACT_SHEET_CACHE.get(cache_key) == cached:
            EXCEL_CONTACT_SHEET_CACHE.pop(cache_key, None)
    return None


def remember_contact_sheet(cache_key: Optional[Tuple[Any, ...]], path: Optional[Path]) -> Optional[Path]:
    if cache_key is not None and path is not None:
        with EXCEL_CONTACT_SHEET_CACHE_LOCK:
            EXCEL_CONTACT_SHEET_CACHE[cache_key] = path
    return path


def prepare_attachment_contact_sheet(
    attachments: List[Dict[str, Any]],
    max_width: int,
    max_height: int,
    allow_upscale: bool = False,
    stretch_tiles: bool = False,
    trim_whitespace: bool = False,
    compact_whitespace: bool = False,
    max_columns: int = 2,
) -> Optional[Path]:
    if Image is None:
        return None
    normalized = [attachment for attachment in (normalize_attachment(raw) for raw in attachments) if attachment]
    if not normalized:
        return None
    cache_key = attachment_contact_sheet_cache_key(
        normalized,
        max_width,
        max_height,
        allow_upscale,
        stretch_tiles,
        trim_whitespace,
        compact_whitespace,
        max_columns,
    )
    cached = get_cached_contact_sheet(cache_key)
    if cached is not None:
        return cached
    if len(normalized) == 1 and not allow_upscale:
        return remember_contact_sheet(cache_key, prepare_attachment_image_file(normalized[0]))
    target_dir = WORK_DIR / "prepared_images"
    target_dir.mkdir(parents=True, exist_ok=True)
    prepared_images = []
    for attachment in normalized:
        output = attachment_image_for_export(attachment)
        if output is None:
            continue
        output = downscale_for_excel_processing(
            output,
            max_width,
            max_height,
            allow_upscale,
            compact_whitespace,
            max_columns,
        )
        if trim_whitespace:
            output = trim_white_image(output)
        if compact_whitespace:
            output = compact_vertical_whitespace(output)
        prepared_images.append(output)
    if not prepared_images:
        return None
    target_dir = WORK_DIR / "prepared_images"
    target_dir.mkdir(parents=True, exist_ok=True)
    if len(prepared_images) == 1 and not stretch_tiles:
        source = prepared_images[0]
        render_scale = EXCEL_IMAGE_RENDER_SCALE if allow_upscale else 1
        render_width = max(1, int(max_width * render_scale))
        render_height = max(1, int(max_height * render_scale))
        scale = min(render_width / max(1, source.width), render_height / max(1, source.height))
        if not allow_upscale:
            scale = min(scale, 1.0)
        resampling = getattr(Image, "Resampling", None)
        resample = getattr(resampling, "LANCZOS", getattr(Image, "LANCZOS", Image.BICUBIC))
        resized = source.resize((max(1, int(source.width * scale)), max(1, int(source.height * scale))), resample)
        return remember_contact_sheet(cache_key, save_excel_ready_image(resized, target_dir))
    if max_columns == 1 and compact_whitespace and not stretch_tiles:
        render_scale = 2 if allow_upscale else 1
        render_width = max(1, int(max_width * render_scale))
        gutter = 8 * render_scale
        resampling = getattr(Image, "Resampling", None)
        resample = getattr(resampling, "LANCZOS", getattr(Image, "LANCZOS", Image.BICUBIC))
        resized_tiles = []
        for source in prepared_images:
            scale = render_width / max(1, source.width)
            if not allow_upscale:
                scale = min(scale, 1.0)
            resized_tiles.append(source.resize((max(1, int(source.width * scale)), max(1, int(source.height * scale))), resample))
        render_height = sum(tile.height for tile in resized_tiles) + gutter * max(0, len(resized_tiles) - 1)
        sheet = Image.new("RGB", (render_width, max(1, render_height)), "white")
        y = 0
        for index, tile in enumerate(resized_tiles):
            x = max(0, (render_width - tile.width) // 2)
            sheet.paste(tile, (x, y))
            y += tile.height
            if index < len(resized_tiles) - 1:
                y += gutter
        return remember_contact_sheet(cache_key, save_excel_ready_image(sheet, target_dir))
    cols = max(1, min(max_columns, len(prepared_images)))
    rows = (len(prepared_images) + cols - 1) // cols
    render_scale = 2 if allow_upscale else 1
    render_width = max(1, int(max_width * render_scale))
    render_height = max(1, int(max_height * render_scale))
    gutter = 8 * render_scale
    cell_width = max(1, (render_width - gutter * (cols - 1)) // cols)
    cell_height = max(1, (render_height - gutter * (rows - 1)) // rows)
    sheet = Image.new("RGB", (render_width, render_height), "white")
    for index, img in enumerate(prepared_images):
        tile = img.copy()
        if stretch_tiles:
            resampling = getattr(Image, "Resampling", None)
            resample = getattr(resampling, "LANCZOS", getattr(Image, "LANCZOS", Image.BICUBIC))
            tile = tile.resize((cell_width, cell_height), resample)
        elif allow_upscale:
            scale = min(cell_width / max(1, tile.width), cell_height / max(1, tile.height))
            resampling = getattr(Image, "Resampling", None)
            resample = getattr(resampling, "LANCZOS", getattr(Image, "LANCZOS", Image.BICUBIC))
            tile = tile.resize((max(1, int(tile.width * scale)), max(1, int(tile.height * scale))), resample)
        else:
            tile.thumbnail((cell_width, cell_height))
        col = index % cols
        row = index // cols
        x = col * (cell_width + gutter) + max(0, (cell_width - tile.width) // 2)
        y = row * (cell_height + gutter) + max(0, (cell_height - tile.height) // 2)
        sheet.paste(tile, (x, y))
    return remember_contact_sheet(cache_key, save_excel_ready_image(sheet, target_dir))


def resize_attachment_contact_sheet(
    attachments: List[Dict[str, Any]],
    max_width: int,
    max_height: int,
    allow_upscale: bool = False,
    stretch_tiles: bool = False,
    trim_whitespace: bool = False,
    compact_whitespace: bool = False,
    max_columns: int = 2,
) -> Optional[Any]:
    contact_path = prepare_attachment_contact_sheet(
        attachments,
        max_width,
        max_height,
        allow_upscale,
        stretch_tiles,
        trim_whitespace,
        compact_whitespace,
        max_columns,
    )
    if contact_path is None:
        return None
    img = XLImage(str(contact_path))
    with Image.open(contact_path) as pil_img:
        width, height = pil_img.size
    scale = min(max_width / max(1, width), max_height / max(1, height))
    if not allow_upscale:
        scale = min(scale, 1.0)
    img.width = max(1, int(width * scale))
    img.height = max(1, int(height * scale))
    return img


def prewarm_attachment_contact_sheet(
    attachments: List[Dict[str, Any]],
    max_width: int,
    max_height: int,
    allow_upscale: bool = False,
    stretch_tiles: bool = False,
    trim_whitespace: bool = False,
    compact_whitespace: bool = False,
    max_columns: int = 2,
) -> None:
    try:
        prepare_attachment_contact_sheet(
            attachments,
            max_width,
            max_height,
            allow_upscale,
            stretch_tiles,
            trim_whitespace,
            compact_whitespace,
            max_columns,
        )
    except Exception:
        LOGGER.debug("Background Excel image preparation failed.", exc_info=True)


def add_xl_image_with_padding(sheet: Any, image: Any, cell_ref: str, padding_px: int = 0) -> None:
    if padding_px <= 0 or AnchorMarker is None or OneCellAnchor is None or XDRPositiveSize2D is None or pixels_to_EMU is None:
        sheet.add_image(image, cell_ref)
        return
    match = re.fullmatch(r"([A-Z]+)(\d+)", cell_ref)
    if not match:
        sheet.add_image(image, cell_ref)
        return
    try:
        col_letters, row_text = match.groups()
        if column_index_from_string is not None:
            col_index = column_index_from_string(col_letters) - 1
        else:
            col_index = 0
            for char in col_letters:
                col_index = col_index * 26 + (ord(char) - 64)
            col_index -= 1
        row_index = int(row_text) - 1
        image.anchor = OneCellAnchor(
            _from=AnchorMarker(
                col=col_index,
                row=row_index,
                colOff=pixels_to_EMU(padding_px),
                rowOff=pixels_to_EMU(padding_px),
            ),
            ext=XDRPositiveSize2D(
                cx=pixels_to_EMU(max(1, int(getattr(image, "width", 1)))),
                cy=pixels_to_EMU(max(1, int(getattr(image, "height", 1)))),
            ),
        )
        sheet.add_image(image)
    except Exception:
        sheet.add_image(image, cell_ref)


def korea_attachment_is_pdf_page(attachment: Dict[str, Any]) -> bool:
    source_path = str(attachment.get("source_path") or "")
    source_page = str(attachment.get("source_page") or "")
    return source_path.lower().endswith(".pdf") and bool(source_page)


def korea_attachment_needs_full_page(attachment: Dict[str, Any], force_full: bool = False) -> bool:
    if force_full or korea_attachment_is_pdf_page(attachment):
        return True
    if Image is None:
        return False
    try:
        path = Path(str(attachment.get("path") or ""))
        with Image.open(path) as img:
            width, height = img.size
    except Exception:
        return False
    if width <= 0 or height <= 0:
        return False
    aspect = height / width
    return height >= 1400 or aspect >= 1.65 or width >= 1400


def split_korea_attachment_for_full_page(attachment: Dict[str, Any]) -> List[Dict[str, Any]]:
    if Image is None:
        return [attachment]
    try:
        prepared = prepare_attachment_image_file(attachment)
        with Image.open(prepared) as source:
            image = source.convert("RGB").copy()
    except Exception:
        return [attachment]
    width, height = image.size
    if width <= 0 or height <= 0:
        return [attachment]
    max_slice_height = max(1, int(width * KOREA_FULL_PAGE_SPLIT_ASPECT))
    if height <= max_slice_height:
        return [attachment]
    target_dir = WORK_DIR / "prepared_images"
    target_dir.mkdir(parents=True, exist_ok=True)
    overlap = min(60, max(0, max_slice_height // 20))
    slices: List[Dict[str, Any]] = []
    y = 0
    part = 1
    while y < height:
        bottom = min(height, y + max_slice_height)
        if height - bottom and height - bottom < max_slice_height * 0.25:
            bottom = height
        cropped = image.crop((0, y, width, bottom))
        output = target_dir / f"{uuid.uuid4().hex}_part_{part}.png"
        cropped.save(output, "PNG")
        next_attachment = copy.deepcopy(attachment)
        base_name = str(attachment.get("filename") or Path(str(attachment.get("path") or "")).name)
        source_page = str(attachment.get("source_page") or "")
        next_attachment.update(
            {
                "id": uuid.uuid4().hex,
                "path": str(output),
                "filename": f"{base_name} part {part}",
                "source_page": f"{source_page}-part-{part}" if source_page else f"part-{part}",
                "crop_box": None,
                "rotation_degrees": 0,
            }
        )
        slices.append(next_attachment)
        if bottom >= height:
            break
        y = max(0, bottom - overlap)
        part += 1
    return slices or [attachment]


def korea_receipt_slot_count_for_blocks(blocks: List[Dict[str, Any]]) -> int:
    slot_index = 0
    for block in blocks:
        mode = str(block.get("mode") or "half")
        if mode == "full":
            remainder = slot_index % KOREA_RECEIPT_SLOTS_PER_PAGE
            if remainder:
                slot_index += KOREA_RECEIPT_SLOTS_PER_PAGE - remainder
            slot_index += KOREA_RECEIPT_SLOTS_PER_PAGE
        elif mode == "wide":
            if slot_index % 2:
                slot_index += 1
            slot_index += 2
        else:
            slot_index += 1
    return max(1, slot_index)


def build_korea_receipt_blocks(
    items: List["ReceiptItem"],
    exchange_rate_attachments: Optional[List[Dict[str, Any]]] = None,
) -> List[Dict[str, Any]]:
    blocks: List[Dict[str, Any]] = []
    if exchange_rate_attachments:
        blocks.append(
            {
                "label": "汇率 / Exchange Rate",
                "attachments": exchange_rate_attachments,
                "mode": "wide",
                "stretch_tiles": False,
            }
        )
    for idx, item in enumerate(items):
        label = korea_receipt_payment_label(idx + 1, item)
        attachments = ensure_receipt_images(item)
        force_full = len(attachments) > 1
        for attachment in attachments:
            if korea_attachment_needs_full_page(attachment, force_full):
                for part_attachment in split_korea_attachment_for_full_page(attachment):
                    blocks.append(
                        {
                            "label": label,
                            "attachments": [part_attachment],
                            "mode": "full",
                            "stretch_tiles": False,
                        }
                    )
            else:
                blocks.append(
                    {
                        "label": label,
                        "attachments": [attachment],
                        "mode": "wide" if korea_receipt_block_is_wide(item) else "half",
                        "stretch_tiles": False,
                    }
                )
    return blocks


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
        sheet.column_dimensions[col].width = 14


def configure_korea_receipt_page(sheet: Any, slot_count: int) -> None:
    page_height = KOREA_RECEIPT_PAGE_HEIGHT
    slots_per_page = KOREA_RECEIPT_SLOTS_PER_PAGE
    page_count = max(1, (slot_count + slots_per_page - 1) // slots_per_page)
    last_row = page_count * page_height
    sheet.print_area = f"A1:H{last_row}"
    if sheet.sheet_properties.pageSetUpPr is None:
        try:
            from openpyxl.worksheet.properties import PageSetupProperties

            sheet.sheet_properties.pageSetUpPr = PageSetupProperties()
        except Exception:
            pass
    if sheet.sheet_properties.pageSetUpPr is not None:
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "portrait"
    sheet.page_setup.scale = 78
    sheet.page_setup.fitToWidth = None
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


def korea_receipt_slot(index: int, wide: bool = False, full_page: bool = False) -> Dict[str, Any]:
    page_height = KOREA_RECEIPT_PAGE_HEIGHT
    slots_per_page = KOREA_RECEIPT_SLOTS_PER_PAGE
    row_offsets = [1, 1, 31, 31]
    col_ranges = [("A", "D"), ("E", "H"), ("A", "D"), ("E", "H")]
    page = index // slots_per_page
    slot = index % slots_per_page
    label_row = page * page_height + row_offsets[slot]
    image_row = label_row + 1
    if full_page:
        label_row = page * page_height + 1
        return {
            "label_range": f"A{label_row}:H{label_row}",
            "label_cell": f"A{label_row}",
            "image_cell": f"A{label_row + 1}",
            "max_width": KOREA_RECEIPT_WIDE_WIDTH,
            "max_height": KOREA_RECEIPT_FULL_HEIGHT,
        }
    if wide:
        return {
            "label_range": f"A{label_row}:H{label_row}",
            "label_cell": f"A{label_row}",
            "image_cell": f"A{image_row}",
            "max_width": KOREA_RECEIPT_WIDE_WIDTH,
            "max_height": KOREA_RECEIPT_HALF_HEIGHT,
        }
    start_col, end_col = col_ranges[slot]
    return {
        "label_range": f"{start_col}{label_row}:{end_col}{label_row}",
        "label_cell": f"{start_col}{label_row}",
        "image_cell": f"{start_col}{image_row}",
        "max_width": KOREA_RECEIPT_HALF_WIDTH,
        "max_height": KOREA_RECEIPT_HALF_HEIGHT,
    }


def korea_receipt_cost_text(item: "ReceiptItem") -> str:
    amount = item.amount.strip()
    if amount:
        return f"{amount} {normalize_currency(item.currency, 'KRW')}".strip()
    if item.krw_amount.strip():
        return f"{item.krw_amount.strip()} KRW"
    if item.rmb_amount.strip():
        return f"{item.rmb_amount.strip()} RMB"
    return ""


def korea_receipt_payment_label(index: int, item: "ReceiptItem") -> str:
    details: List[str] = []
    if item.date.strip():
        details.append(item.date.strip())
    content = korea_detail_content(item)
    if content:
        details.append(content)
    cost = korea_receipt_cost_text(item)
    if cost:
        details.append(cost)
    return " | ".join(details) if details else item.filename or f"Payment {index}"


def set_korea_receipt_label(sheet: Any, cell_ref: str, label: str) -> None:
    cell = sheet[cell_ref]
    cell.value = label
    if Font is not None:
        cell.font = Font(bold=True, size=10, color="1F2937")
    if Alignment is not None:
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    if PatternFill is not None:
        cell.fill = PatternFill(fill_type="solid", fgColor="E9EEF3")


def usa_report_section_for_category(category: str) -> str:
    if category in {block[0] for block in USA_EXPENSE_REPORT_BLOCKS}:
        return category
    return "office"


def copy_excel_cell(source_cell: Any, target_cell: Any, copy_value: bool = True) -> None:
    if MergedCell is not None and isinstance(source_cell, MergedCell):
        return
    if copy_value:
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


def copy_excel_row(source_ws: Any, target_ws: Any, source_row: int, target_row: int, max_col: int) -> None:
    for col in range(1, max_col + 1):
        copy_excel_cell(source_ws.cell(source_row, col), target_ws.cell(target_row, col))
    source_dim = source_ws.row_dimensions[source_row]
    target_dim = target_ws.row_dimensions[target_row]
    target_dim.height = source_dim.height
    target_dim.hidden = source_dim.hidden
    target_dim.outlineLevel = source_dim.outlineLevel
    target_dim.collapsed = source_dim.collapsed


def set_cell_alignment(
    cell: Any,
    horizontal: Optional[str] = None,
    vertical: Optional[str] = None,
    wrap_text: Optional[bool] = None,
) -> None:
    if Alignment is None:
        return
    existing = cell.alignment
    cell.alignment = Alignment(
        horizontal=horizontal if horizontal is not None else existing.horizontal,
        vertical=vertical if vertical is not None else existing.vertical,
        text_rotation=existing.textRotation,
        wrap_text=wrap_text if wrap_text is not None else existing.wrap_text,
        shrink_to_fit=existing.shrink_to_fit,
        indent=existing.indent,
    )


def configure_usa_expense_header_row(sheet: Any) -> None:
    sheet.row_dimensions[5].height = max(sheet.row_dimensions[5].height or 0, 32)
    for col in "ABCDEFGH":
        cell = sheet[f"{col}5"]
        if MergedCell is not None and isinstance(cell, MergedCell):
            continue
        set_cell_alignment(
            cell,
            horizontal=cell.alignment.horizontal or "center",
            vertical="center",
            wrap_text=True,
        )


def configure_usa_expense_entry_row(sheet: Any, row: int) -> None:
    sheet.row_dimensions[row].height = max(sheet.row_dimensions[row].height or 0, 18)
    set_cell_alignment(sheet[f"A{row}"], horizontal="center", vertical="center", wrap_text=True)
    set_cell_alignment(sheet[f"B{row}"], horizontal="center", vertical="center")


def estimate_wrapped_lines(text: Any, width_chars: float) -> int:
    cleaned = str(text or "").strip()
    if not cleaned:
        return 1
    usable_width = max(8, int(width_chars * 0.82))
    total = 0
    for raw_line in cleaned.splitlines() or [""]:
        line = raw_line.strip()
        total += max(1, math.ceil(len(line) / usable_width))
    return total


def adjust_usa_expense_entry_row_height(sheet: Any, row: int) -> None:
    width = sheet.column_dimensions["A"].width or 24
    line_count = estimate_wrapped_lines(sheet[f"A{row}"].value, width)
    target_height = 18 if line_count <= 1 else min(72, 18 * line_count)
    sheet.row_dimensions[row].height = max(sheet.row_dimensions[row].height or 0, target_height)
    set_cell_alignment(sheet[f"A{row}"], horizontal="center", vertical="center", wrap_text=True)


def prepare_usa_expense_report_sheet(
    wb: Any,
    template_ws: Any,
    section_counts: Dict[str, int],
) -> Tuple[Any, Dict[str, List[int]]]:
    template_index = wb.worksheets.index(template_ws)
    template_title = template_ws.title
    generated_ws = wb.create_sheet(f"{template_title} generated", template_index)
    max_col = max(template_ws.max_column or 1, 10)

    for key, dim in template_ws.column_dimensions.items():
        generated_ws.column_dimensions[key] = copy.copy(dim)
    generated_ws.freeze_panes = template_ws.freeze_panes
    generated_ws.sheet_view.showGridLines = template_ws.sheet_view.showGridLines
    generated_ws.sheet_format = copy.copy(template_ws.sheet_format)
    generated_ws.sheet_properties = copy.copy(template_ws.sheet_properties)
    generated_ws.page_margins = copy.copy(template_ws.page_margins)
    generated_ws.page_setup = copy.copy(template_ws.page_setup)
    generated_ws.print_options = copy.copy(template_ws.print_options)

    for row in range(1, 6):
        copy_excel_row(template_ws, generated_ws, row, row, max_col)
    configure_usa_expense_header_row(generated_ws)

    category_rows: Dict[str, List[int]] = {}
    subtotal_rows: List[int] = []
    target_row = 6
    for section, label_row, template_entry_rows, subtotal_row in USA_EXPENSE_REPORT_BLOCKS:
        copy_excel_row(template_ws, generated_ws, label_row, target_row, max_col)
        label_target_row = target_row
        generated_ws[f"G{label_target_row}"] = f"=F{label_target_row}*$J$1"
        target_row += 1

        required_entries = max(len(template_entry_rows), section_counts.get(section, 0))
        entry_rows: List[int] = []
        for index in range(required_entries):
            source_row = template_entry_rows[min(index, len(template_entry_rows) - 1)]
            copy_excel_row(template_ws, generated_ws, source_row, target_row, max_col)
            current_height = generated_ws.row_dimensions[target_row].height
            if current_height is None or current_height < 18:
                generated_ws.row_dimensions[target_row].height = 18
            configure_usa_expense_entry_row(generated_ws, target_row)
            generated_ws[f"G{target_row}"] = f"=F{target_row}*$J$1"
            entry_rows.append(target_row)
            target_row += 1

        copy_excel_row(template_ws, generated_ws, subtotal_row, target_row, max_col)
        generated_ws.row_dimensions[target_row].height = max(
            generated_ws.row_dimensions[target_row].height or 0,
            30,
        )
        generated_ws[f"F{target_row}"] = f"=SUM(F{label_target_row}:F{entry_rows[-1]})"
        generated_ws[f"G{target_row}"] = f"=F{target_row}*$J$1"
        generated_ws.merge_cells(f"A{target_row}:C{target_row}")
        category_rows[section] = entry_rows
        subtotal_rows.append(target_row)
        target_row += 1

    copy_excel_row(template_ws, generated_ws, 70, target_row, max_col)
    generated_ws.row_dimensions[target_row].height = max(
        generated_ws.row_dimensions[target_row].height or 0,
        30,
    )
    generated_ws[f"F{target_row}"] = f"=SUM({','.join(f'F{row}' for row in subtotal_rows)})"
    generated_ws[f"G{target_row}"] = f"=F{target_row}*$J$1"
    generated_ws.merge_cells(f"A{target_row}:C{target_row}")
    generated_ws.merge_cells("A1:H1")
    try:
        generated_ws.merge_cells(f"H5:H{max(5, target_row - 1)}")
    except Exception:
        pass
    try:
        generated_ws.print_area = f"$A$1:$H${target_row}"
    except Exception:
        pass

    wb.remove(template_ws)
    generated_ws.title = template_title
    return generated_ws, category_rows


def add_korea_receipt_block(
    sheet: Any,
    index: int,
    label: str,
    attachments: List[Dict[str, Any]],
    wide: bool = False,
    full_page: bool = False,
    stretch_tiles: bool = True,
) -> None:
    slot = korea_receipt_slot(index, wide=wide, full_page=full_page)
    label_row = int(re.sub(r"^[A-Z]+", "", str(slot["label_cell"])))
    try:
        sheet.merge_cells(str(slot["label_range"]))
    except Exception:
        pass
    sheet.row_dimensions[label_row].height = 22
    set_korea_receipt_label(sheet, str(slot["label_cell"]), label)
    receipt_image = resize_attachment_contact_sheet(
        attachments,
        int(slot["max_width"]),
        int(slot["max_height"]),
        allow_upscale=True,
        stretch_tiles=stretch_tiles,
    )
    if receipt_image is not None:
        receipt_image.width = int(slot["max_width"])
        receipt_image.height = int(slot["max_height"])
        sheet.add_image(receipt_image, str(slot["image_cell"]))


def export_usa(
    items: List[ReceiptItem],
    output_path: Path,
    exchange_rate: float,
    bank_items: Optional[List[BankStatementItem]] = None,
    progress: Optional[ExportProgressCallback] = None,
) -> None:
    setup_logging()
    export_started = time.perf_counter()
    emit_export_progress(progress, 0, 1, "Loading USA template...")
    template = TEMPLATE_DIR / USA_TEMPLATE_NAME
    if not template.exists():
        raise FileNotFoundError(f"Missing USA template: {template}")
    wb = load_workbook(template)
    ws = wb["Expense report"]
    receipts_ws = wb["Receipt and Payment of expenses"]
    report_rows = build_usa_report_rows(items, bank_items)
    total_steps = max(4, len(report_rows) + 4)
    emit_export_progress(progress, 1, total_steps, "Filling USA expense report...")
    sorted_items = [
        row.get("expense_item") or row.get("item")
        for row in report_rows
        if (row.get("expense_item") is not None or row.get("item") is not None)
        and safe_float(getattr(row.get("expense_item") or row.get("item"), "amount", "")) is not None
    ]
    section_counts = {block[0]: 0 for block in USA_EXPENSE_REPORT_BLOCKS}
    for item in sorted_items:
        category = usa_export_category_for_item(item)
        section_counts[usa_report_section_for_category(category)] += 1
    ws, category_rows = prepare_usa_expense_report_sheet(wb, ws, section_counts)

    ws["A3"] = f"Date / 填表日期： {date.today().strftime('%m/%d/%Y')}"
    ws["A4"] = "Employee: / 申请人："
    ws["J1"] = exchange_rate
    all_entry_rows = sorted({row for rows in category_rows.values() for row in rows})
    clear_cells(ws, all_entry_rows, ["A", "B", "C", "D", "E", "F", "H"])
    for row in all_entry_rows:
        ws[f"G{row}"] = f"=F{row}*$J$1"

    row_cursors = {cat: 0 for cat in category_rows}
    for item in sorted_items:
        category = usa_export_category_for_item(item)
        section = usa_report_section_for_category(category)
        rows = category_rows[section]
        cursor = row_cursors[section]
        row = rows[cursor]
        row_cursors[section] += 1
        amount = safe_float(item.amount)
        ws[f"A{row}"] = item.place.strip()
        ws[f"B{row}"] = excel_report_date_value(item.date)
        ws[f"C{row}"] = item.details.strip() or item.receipt_label.strip() or item.filename
        ws[f"D{row}"] = item.purpose.strip()
        ws[f"E{row}"] = item.project_number.strip()
        ws[f"F{row}"] = amount if amount is not None else None
        ws[f"G{row}"] = f"=F{row}*$J$1"
        adjust_usa_expense_entry_row_height(ws, row)

    clear_images(receipts_ws)
    for row in range(2, max(56, len(report_rows) + 3)):
        for col in "ABCDE":
            receipts_ws[f"{col}{row}"] = None
    receipts_ws.column_dimensions["A"].width = 11
    receipts_ws.column_dimensions["B"].width = 25.9
    receipts_ws.column_dimensions["C"].width = 17
    receipts_ws.column_dimensions["D"].width = USA_RECEIPT_COLUMN_WIDTH
    receipts_ws.column_dimensions["E"].width = USA_PROOF_COLUMN_WIDTH
    image_padding = USA_SUPPORT_IMAGE_PADDING_PX
    receipt_image_width = max(1, USA_RECEIPT_IMAGE_MAX_WIDTH - image_padding * 2)
    receipt_image_height = max(1, USA_RECEIPT_IMAGE_MAX_HEIGHT - image_padding * 2)
    proof_image_width = max(1, USA_PROOF_IMAGE_MAX_WIDTH - image_padding * 2)
    proof_image_height = max(1, USA_PROOF_IMAGE_MAX_HEIGHT - image_padding * 2)
    support_rows = sort_usa_support_rows_by_date(report_rows)
    images_started = time.perf_counter()
    embedded_images = 0
    emit_export_progress(progress, 2, total_steps, "Preparing receipt and payment proof images...")
    for index, row_data in enumerate(support_rows, start=1):
        row = index + 1
        receipts_ws[f"A{row}"] = index
        receipts_ws[f"B{row}"] = excel_report_date_value(str(row_data.get("date") or ""))
        receipts_ws[f"C{row}"] = safe_float(row_data.get("amount"))
        receipts_ws[f"B{row}"].number_format = USA_SUPPORT_DATE_NUMBER_FORMAT
        receipts_ws[f"C{row}"].number_format = '$#,##0.00'
        for col in ("A", "B", "C"):
            set_cell_alignment(receipts_ws[f"{col}{row}"], horizontal="center", vertical="center")
        receipts_ws.row_dimensions[row].height = USA_SUPPORT_ROW_HEIGHT
        receipt_image = resize_attachment_contact_sheet(
            list(row_data.get("receipt_images") or []),
            receipt_image_width,
            receipt_image_height,
            allow_upscale=True,
        )
        if receipt_image is not None:
            add_xl_image_with_padding(receipts_ws, receipt_image, f"D{row}", image_padding)
            embedded_images += 1
        proof_image = resize_attachment_contact_sheet(
            list(row_data.get("proof_images") or []),
            proof_image_width,
            proof_image_height,
            allow_upscale=True,
            trim_whitespace=True,
            compact_whitespace=True,
            max_columns=1,
        )
        if proof_image is not None:
            add_xl_image_with_padding(receipts_ws, proof_image, f"E{row}", image_padding)
            embedded_images += 1
        emit_export_progress(
            progress,
            min(total_steps - 2, index + 2),
            total_steps,
            f"Processing Excel images {index}/{len(support_rows)}...",
        )

    save_started = time.perf_counter()
    emit_export_progress(progress, total_steps - 1, total_steps, "Saving Excel workbook...")
    wb.save(output_path)
    export_finished = time.perf_counter()
    emit_export_progress(progress, total_steps, total_steps, "Excel workbook saved.")
    LOGGER.info(
        "USA export completed: rows=%s images=%s image_time=%.2fs save_time=%.2fs total=%.2fs output=%s",
        len(report_rows),
        embedded_images,
        save_started - images_started,
        export_finished - save_started,
        export_finished - export_started,
        output_path,
    )


def clone_sheet(
    source_ws: Any,
    target_wb: Any,
    title: str,
    index: Optional[int] = None,
    min_rows: int = 0,
    min_cols: int = 0,
) -> Any:
    target_ws = target_wb.create_sheet(title=title, index=index)
    max_row = max(source_ws.max_row or 1, min_rows)
    max_col = max(source_ws.max_column or 1, min_cols)
    for row in source_ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
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
    target_ws.sheet_format = copy.copy(source_ws.sheet_format)
    target_ws.sheet_properties = copy.copy(source_ws.sheet_properties)
    target_ws.page_margins = copy.copy(source_ws.page_margins)
    target_ws.page_setup = copy.copy(source_ws.page_setup)
    target_ws.print_options = copy.copy(source_ws.print_options)
    if source_ws.print_area:
        target_ws.print_area = source_ws.print_area
    return target_ws


def clear_freeze_panes(sheet: Any) -> None:
    sheet.freeze_panes = None
    sheet.sheet_view.pane = None
    if Selection is not None:
        sheet.sheet_view.selection = [Selection(activeCell="A1", sqref="A1")]


def korea_amounts(
    item: ReceiptItem,
    krw_to_rmb_rate: float,
    usd_to_rmb_rate: float,
    usd_to_krw_rate: float,
) -> Tuple[Optional[float], Optional[float], str]:
    amount = safe_float(item.amount)
    krw = safe_float(item.krw_amount)
    rmb = safe_float(item.rmb_amount)
    currency = normalize_currency(item.currency, "KRW")
    if amount is not None:
        if currency == "USD":
            krw = amount * usd_to_krw_rate if usd_to_krw_rate else krw
            if krw is None and usd_to_rmb_rate and krw_to_rmb_rate:
                krw = (amount * usd_to_rmb_rate) / krw_to_rmb_rate
            rmb = krw * krw_to_rmb_rate if krw is not None and krw_to_rmb_rate else (amount * usd_to_rmb_rate if usd_to_rmb_rate else rmb)
        elif currency == "RMB" or currency == "CNY":
            rmb = amount
            krw = amount / krw_to_rmb_rate if krw_to_rmb_rate else (
                (amount / usd_to_rmb_rate) * usd_to_krw_rate if usd_to_rmb_rate and usd_to_krw_rate else None
            )
        elif currency == "KRW":
            krw = amount
            rmb = amount * krw_to_rmb_rate if krw_to_rmb_rate else None
    if krw is None and rmb is not None and krw_to_rmb_rate:
        krw = rmb / krw_to_rmb_rate
    if rmb is None and krw is not None and krw_to_rmb_rate:
        rmb = krw * krw_to_rmb_rate
    note = ""
    if amount is not None and currency == "USD":
        note = f"{amount:g} USD"
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


def korea_receipt_slot_count(items: List["ReceiptItem"], has_exchange_rate: bool = False) -> int:
    slot_index = 2 if has_exchange_rate else 0
    for item in items:
        wide = korea_receipt_block_is_wide(item)
        if wide and slot_index % 2:
            slot_index += 1
        slot_index += 2 if wide else 1
    return max(1, slot_index)


def export_korea(
    items: List[ReceiptItem],
    output_path: Path,
    krw_to_rmb_rate: float,
    usd_to_rmb_rate: float,
    usd_to_krw_rate: float,
    exchange_rate_items: Optional[List[BankStatementItem]] = None,
    progress: Optional[ExportProgressCallback] = None,
) -> None:
    setup_logging()
    export_started = time.perf_counter()
    emit_export_progress(progress, 0, 1, "Loading Korea templates...")
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
    detail_ws = clone_sheet(detail_template_wb.worksheets[0], wb, "报销明细", index=0, min_rows=35, min_cols=20)
    sorted_items = sort_receipts_for_report(items, "Korea")
    emit_export_progress(progress, 1, max(4, len(sorted_items) + 4), "Filling Korea reimbursement details...")
    try:
        wb._sheets = [detail_ws, cover_ws, receipts_ws]
    except Exception:
        pass
    clear_freeze_panes(detail_ws)
    detail_ws.print_area = "A1:S34"
    if detail_ws.sheet_properties.pageSetUpPr is not None:
        detail_ws.sheet_properties.pageSetUpPr.fitToPage = True
    detail_ws.page_setup.orientation = "landscape"
    detail_ws.page_setup.scale = 34
    detail_ws.page_setup.fitToWidth = None
    detail_ws.page_setup.fitToHeight = None
    detail_ws.page_margins.left = 0.25
    detail_ws.page_margins.right = 0.25
    detail_ws.page_margins.top = 0.75
    detail_ws.page_margins.bottom = 0.75
    for col in "KLMNOPQRS":
        detail_ws.column_dimensions[col].hidden = False

    cover_ws["A2"] = f"报销部门：  {date.today().year}年 {date.today().month}月 {date.today().day}日 填 单据及附件共  页"
    cover_ws["A11"] = "领导审批           会计主管              会计                  出纳                 报销人                   领款人 "
    for row, label in [(row, label) for row, label in KOREA_COVER_ROWS.values()]:
        cover_ws[f"A{row}"] = label
        cover_ws[f"C{row}"] = 0
        cover_ws[f"D{row}"] = 0
    cover_ws["C9"] = "=SUM(C4:C8)"
    cover_ws["D9"] = "=SUM(D4:D8)"
    cover_ws["B10"] = "=D9"

    for merged_range in ("A33:B33", "A34:B34", "A35:B35"):
        try:
            detail_ws.unmerge_cells(merged_range)
        except Exception:
            pass
    clear_cells(detail_ws, range(3, 36), list("ABCDEFGHIJKLMNOPQRS"))
    try:
        detail_ws.merge_cells("A33:B33")
        detail_ws.merge_cells("A34:B34")
    except Exception:
        pass
    detail_ws["A33"] = "合计（外币）\nTotal"
    detail_ws["Q33"] = "=SUM(Q3:Q32)"
    detail_ws["Q33"].number_format = KRW_NUMBER_FORMAT
    detail_ws["A34"] = "合计（人民币）\nTotal"
    detail_ws["R34"] = "=SUM(R3:R33)"
    detail_ws.row_dimensions[33].height = 35.4
    detail_ws.row_dimensions[34].height = 50.4
    detail_ws.row_dimensions[35].height = None

    summary: Dict[str, Tuple[float, float]] = {
        key: (0.0, 0.0) for key in KOREA_COVER_ROWS
    }
    for index, item in enumerate(sorted_items, start=3):
        if index > 32:
            raise RuntimeError("Korea template supports up to 30 detail rows in this version.")
        category = report_category_for_item(item, "Korea")
        if category not in KOREA_CATEGORY_COLUMNS:
            category = "other"
        krw, rmb, original_note = korea_amounts(item, krw_to_rmb_rate, usd_to_rmb_rate, usd_to_krw_rate)
        krw_whole = truncate_amount(krw)
        detail_ws[f"A{index}"] = excel_date_formula_or_value(item.date)
        detail_ws[f"B{index}"] = korea_detail_content(item)
        detail_ws[f"C{index}"] = korea_detail_location(item)
        detail_ws[f"D{index}"] = ""
        detail_ws[f"E{index}"] = item.project_number.strip()
        category_col = KOREA_CATEGORY_COLUMNS[category]
        detail_ws[f"{category_col}{index}"] = krw_whole
        detail_ws[f"{category_col}{index}"].number_format = KRW_NUMBER_FORMAT
        detail_ws[f"P{index}"] = original_note
        detail_ws[f"Q{index}"] = krw_whole
        detail_ws[f"Q{index}"].number_format = KRW_NUMBER_FORMAT
        detail_ws[f"R{index}"] = rmb
        detail_ws[f"S{index}"] = korea_payment_method_text(item)
        bucket = korea_cover_bucket(category)
        cur_krw, cur_rmb = summary[bucket]
        summary[bucket] = (cur_krw + (krw_whole or 0), cur_rmb + (rmb or 0.0))

    for bucket, (krw_total, rmb_total) in summary.items():
        row, _label = KOREA_COVER_ROWS[bucket]
        cover_ws[f"C{row}"] = truncate_amount(krw_total) if krw_total else None
        cover_ws[f"C{row}"].number_format = KRW_NUMBER_FORMAT
        cover_ws[f"D{row}"] = round(rmb_total, 2) if rmb_total else None
    cover_ws["C9"].number_format = KRW_NUMBER_FORMAT

    clear_images(receipts_ws)
    fit_columns_for_receipts(receipts_ws)
    exchange_rate_attachments = [attachment_from_item(item) for item in exchange_rate_items or []]
    receipt_blocks = build_korea_receipt_blocks(sorted_items, exchange_rate_attachments)
    total_steps = max(4, len(receipt_blocks) + 4)
    emit_export_progress(progress, 2, total_steps, "Preparing Korea invoice pages...")
    slot_count = korea_receipt_slot_count_for_blocks(receipt_blocks)
    configure_korea_receipt_page(receipts_ws, slot_count)
    page_height = KOREA_RECEIPT_PAGE_HEIGHT
    slots_per_page = KOREA_RECEIPT_SLOTS_PER_PAGE
    page_count = max(1, (slot_count + slots_per_page - 1) // slots_per_page)
    last_receipt_row = page_count * page_height
    for row in range(1, max(240, last_receipt_row + 1)):
        for col in "ABCDEFGH":
            receipts_ws[f"{col}{row}"] = None
    block_index = 0
    images_started = time.perf_counter()
    embedded_blocks = 0
    for block in receipt_blocks:
        mode = str(block.get("mode") or "half")
        if mode == "full":
            remainder = block_index % KOREA_RECEIPT_SLOTS_PER_PAGE
            if remainder:
                block_index += KOREA_RECEIPT_SLOTS_PER_PAGE - remainder
        elif mode == "wide" and block_index % 2:
            block_index += 1
        add_korea_receipt_block(
            receipts_ws,
            block_index,
            str(block.get("label") or ""),
            list(block.get("attachments") or []),
            wide=mode == "wide",
            full_page=mode == "full",
            stretch_tiles=bool(block.get("stretch_tiles", False)),
        )
        embedded_blocks += 1
        emit_export_progress(
            progress,
            min(total_steps - 2, embedded_blocks + 2),
            total_steps,
            f"Processing invoice images {embedded_blocks}/{len(receipt_blocks)}...",
        )
        block_index += KOREA_RECEIPT_SLOTS_PER_PAGE if mode == "full" else (2 if mode == "wide" else 1)

    save_started = time.perf_counter()
    emit_export_progress(progress, total_steps - 1, total_steps, "Saving Excel workbook...")
    wb.save(output_path)
    export_finished = time.perf_counter()
    emit_export_progress(progress, total_steps, total_steps, "Excel workbook saved.")
    LOGGER.info(
        "Korea export completed: rows=%s receipt_blocks=%s image_time=%.2fs save_time=%.2fs total=%.2fs output=%s",
        len(sorted_items),
        embedded_blocks,
        save_started - images_started,
        export_finished - save_started,
        export_finished - export_started,
        output_path,
    )


class ReimbursementHelperApp:
    def __init__(self, root: Any) -> None:
        self.root = root
        self.items: List[ReceiptItem] = []
        self.bank_items: List[BankStatementItem] = []
        self.exchange_rate_items: List[BankStatementItem] = []
        self.selected_index: Optional[int] = None
        self.selected_bank_index: Optional[int] = None
        self.bank_tree: Optional[Any] = None
        self.photo: Optional[Any] = None
        self.preview_canvas: Optional[Any] = None
        self.revert_crop_btn: Optional[Any] = None
        self.zoom_out_btn: Optional[Any] = None
        self.zoom_in_btn: Optional[Any] = None
        self.zoom_label: Optional[Any] = None
        self.rotate_left_btn: Optional[Any] = None
        self.rotate_right_btn: Optional[Any] = None
        self.delete_screenshot_btn: Optional[Any] = None
        self.swap_proof_btn: Optional[Any] = None
        self.unlink_proof_btn: Optional[Any] = None
        self.upload_folder_btn: Optional[Any] = None
        self.select_payment_proof_btn: Optional[Any] = None
        self.generate_details_btn: Optional[Any] = None
        self.generate_all_btn: Optional[Any] = None
        self.generate_excel_btn: Optional[Any] = None
        self.selected_attachment_kind = "receipt"
        self.selected_attachment_index = 0
        self.image_clipboard: Optional[Dict[str, Any]] = None
        self._selection_anchor_index: Optional[int] = None
        self.preview_tiles: List[Dict[str, Any]] = []
        self.preview_photos: List[Any] = []
        self.preview_action_regions: Dict[str, Dict[str, Any]] = {}
        self.proof_drop_region: Optional[Tuple[int, int, int, int]] = None
        self._pressed_preview_tile: Optional[Dict[str, Any]] = None
        self._dragging_preview_tile: Optional[Dict[str, Any]] = None
        self._drag_start: Tuple[int, int] = (0, 0)
        self._drag_outline_id: Optional[int] = None
        self._preview_image_bounds: Tuple[int, int, int, int] = (0, 0, 0, 0)
        self._preview_original_size: Tuple[int, int] = (0, 0)
        self._crop_handle_centers: Dict[str, Tuple[float, float]] = {}
        self._dragging_crop_handle: Optional[str] = None
        self.settings = load_user_settings()
        self.form_version = tk.StringVar(value=normalized_form_version(self.settings.get("last_form_version", "USA")))
        self.language = tk.StringVar(value=normalize_language(self.settings.get("language", "en")))
        self.language_display = tk.StringVar(value=LANGUAGE_LABELS.get(self.language.get(), "English"))
        self.exchange_rate = tk.StringVar(value=str(self.settings.get("usa_exchange_rate", DEFAULT_USD_TO_RMB_RATE)))
        self.usd_to_krw_rate = tk.StringVar(value=str(self.settings.get("usd_to_krw_rate", DEFAULT_USD_TO_KRW_RATE)))
        self.krw_to_rmb_rate = tk.StringVar(value=str(self.settings.get("krw_to_rmb_rate", DEFAULT_KRW_TO_RMB_RATE)))
        self.ai_concurrency = tk.StringVar(value=str(clamp_ai_concurrency(self.settings.get("ai_concurrency", DEFAULT_AI_REQUEST_CONCURRENCY))))
        self.usa_upload_mode = tk.StringVar(value=normalize_usa_upload_mode(self.settings.get("usa_upload_mode", "auto")))
        self.usa_upload_mode_display = tk.StringVar(value="")
        self.preview_zoom = tk.DoubleVar(value=1.0)
        self.status_text = tk.StringVar(value=self.tr("ready"))
        self.progress_text = tk.StringVar(value="")
        self.progress_value = tk.DoubleVar(value=0)
        self.field_vars: Dict[str, Any] = {}
        self.field_labels: Dict[str, Any] = {}
        self.field_widgets: Dict[str, Any] = {}
        self.category_values: List[str] = []
        self.category_label_to_key: Dict[str, str] = {}
        self.category_combo: Optional[Any] = None
        self.document_type_combo: Optional[Any] = None
        self.header_label: Optional[Any] = None
        self.form_label: Optional[Any] = None
        self.language_label: Optional[Any] = None
        self.ai_concurrency_label: Optional[Any] = None
        self.ai_concurrency_entry: Optional[Any] = None
        self.upload_mode_label: Optional[Any] = None
        self.upload_mode_combo: Optional[Any] = None
        self.manager_label: Optional[Any] = None
        self.details_label: Optional[Any] = None
        self.preview_label: Optional[Any] = None
        self.remove_btn: Optional[Any] = None
        self.clear_btn: Optional[Any] = None
        self.open_output_btn: Optional[Any] = None
        self.usa_rate_label: Optional[Any] = None
        self.usa_rate_entry: Optional[Any] = None
        self.usd_krw_rate_label: Optional[Any] = None
        self.usd_krw_rate_entry: Optional[Any] = None
        self.krw_rate_label: Optional[Any] = None
        self.krw_rate_entry: Optional[Any] = None
        self._loading_fields = False
        self._syncing_amounts = False
        self._last_amount_source = "amount"
        self._last_form_version = self.form_version.get()
        self._busy = False
        self._ai_cancel_event: Optional[threading.Event] = None
        self._details_ready_for_export = False
        self._excel_prewarm_executor = ThreadPoolExecutor(max_workers=2)
        self._excel_prewarm_futures: List[Any] = []
        self._excel_prewarm_keys: set = set()
        self._excel_prewarm_lock = threading.Lock()
        self._receipt_import_active = False
        self._receipt_memory = load_receipt_memory()
        self.last_generated_output_path: Optional[Path] = None
        self._build_ui()
        self._attach_amount_traces()
        self._on_form_version_changed()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(250, self.restore_previous_session_if_available)

    def tr(self, key: str, **values: Any) -> str:
        return ui_text(self.language.get(), key, **values)

    def _build_ui(self) -> None:
        self.root.title(self.tr("app_title"))
        self._set_initial_window_geometry()
        app_bg = "#E9EEF3"
        panel_bg = "#F8FAFC"
        canvas_bg = "#F5F7FA"
        border = "#CBD5E1"
        text = "#17202A"
        muted = "#64748B"
        accent = "#2563EB"
        button_bg = "#E6EDF5"
        button_active = "#D6E2EE"
        button_disabled = "#EEF3F8"
        button_border = "#94A3B8"
        control_bg = "#F1F5F9"
        self.root.configure(bg=app_bg)
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            ".",
            background=app_bg,
            foreground=text,
            bordercolor=border,
            lightcolor="#F8FAFC",
            darkcolor="#AAB8C7",
            troughcolor="#D8E1EC",
            focuscolor=app_bg,
        )
        style.configure("TFrame", background=app_bg)
        style.configure("Panel.TFrame", background=panel_bg)
        style.configure("TLabel", background=app_bg, foreground=text, font=("Segoe UI", 10))
        style.configure("Header.TLabel", background=app_bg, foreground=text, font=("Segoe UI", 13, "bold"))
        style.configure("Panel.TLabel", background=panel_bg, foreground=text, font=("Segoe UI", 10))
        style.configure("PanelHeader.TLabel", background=panel_bg, foreground=text, font=("Segoe UI", 13, "bold"))
        style.configure("Muted.TLabel", background=app_bg, foreground=muted, font=("Segoe UI", 9))
        style.configure(
            "TButton",
            font=("Segoe UI", 10),
            padding=(10, 5),
            background=button_bg,
            foreground=text,
            bordercolor=button_border,
            lightcolor=button_bg,
            darkcolor=button_border,
            focuscolor=button_bg,
            relief="flat",
        )
        style.map(
            "TButton",
            background=[
                ("pressed", "#C6D7E8"),
                ("active", button_active),
                ("disabled", button_disabled),
            ],
            foreground=[("disabled", "#94A3B8")],
            bordercolor=[("disabled", "#C9D4E0"), ("active", "#7C93AA")],
            lightcolor=[("disabled", button_disabled), ("active", button_active)],
            darkcolor=[("disabled", "#C9D4E0"), ("active", "#7C93AA")],
        )
        style.configure(
            "Recommended.TButton",
            font=("Segoe UI", 10),
            padding=(10, 5),
            background=accent,
            foreground="#FFFFFF",
            bordercolor="#1D4ED8",
            lightcolor=accent,
            darkcolor="#1D4ED8",
            focuscolor=accent,
            relief="flat",
        )
        style.map(
            "Recommended.TButton",
            background=[
                ("pressed", "#1E40AF"),
                ("active", "#1D4ED8"),
                ("disabled", button_disabled),
            ],
            foreground=[("disabled", "#94A3B8"), ("!disabled", "#FFFFFF")],
            bordercolor=[("disabled", "#C9D4E0"), ("active", "#1E40AF")],
            lightcolor=[("disabled", button_disabled), ("active", "#1D4ED8")],
            darkcolor=[("disabled", "#C9D4E0"), ("active", "#1E40AF")],
        )
        style.configure(
            "Icon.TButton",
            font=("Segoe UI Symbol", 13),
            width=3,
            padding=(4, 3),
            background=button_bg,
            bordercolor=button_border,
            lightcolor=button_bg,
            darkcolor=button_border,
            focuscolor=button_bg,
            relief="flat",
        )
        style.configure(
            "TEntry",
            fieldbackground=control_bg,
            background=control_bg,
            foreground=text,
            insertcolor=text,
            bordercolor=border,
            lightcolor=control_bg,
            darkcolor=border,
            focuscolor=control_bg,
            selectbackground="#CFE1F6",
            selectforeground=text,
            relief="flat",
        )
        style.configure(
            "Warning.TEntry",
            fieldbackground="#FFF1F2",
            background="#FFF1F2",
            foreground="#7F1D1D",
            insertcolor=text,
            bordercolor="#DC2626",
            lightcolor="#FFF1F2",
            darkcolor="#DC2626",
            focuscolor="#DC2626",
            selectbackground="#FECACA",
            selectforeground="#7F1D1D",
            relief="flat",
        )
        style.configure(
            "TCombobox",
            fieldbackground=control_bg,
            background=button_bg,
            foreground=text,
            arrowcolor=text,
            bordercolor=border,
            lightcolor=control_bg,
            darkcolor=border,
            focuscolor=control_bg,
            selectbackground="#CFE1F6",
            selectforeground=text,
            relief="flat",
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", control_bg), ("!disabled", control_bg), ("disabled", button_disabled)],
            background=[("readonly", button_bg), ("active", button_active), ("disabled", button_disabled)],
            selectbackground=[("readonly", control_bg), ("!disabled", control_bg), ("disabled", button_disabled)],
            selectforeground=[("readonly", text), ("!disabled", text)],
            bordercolor=[("focus", "#7EA5D3"), ("disabled", "#C9D4E0")],
            lightcolor=[("focus", control_bg), ("disabled", button_disabled)],
            darkcolor=[("focus", "#7EA5D3"), ("disabled", "#C9D4E0")],
        )
        style.configure("Treeview", background=panel_bg, fieldbackground=panel_bg, foreground=text, rowheight=24, bordercolor=border)
        style.configure(
            "Treeview.Heading",
            background="#E2EAF3",
            foreground=text,
            font=("Segoe UI", 9, "bold"),
            bordercolor=button_border,
            lightcolor="#E2EAF3",
            darkcolor=button_border,
            relief="flat",
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#D6E2EE"), ("pressed", "#C6D7E8")],
            bordercolor=[("active", "#7C93AA")],
            lightcolor=[("active", "#D6E2EE")],
            darkcolor=[("active", "#7C93AA")],
        )
        style.map("Treeview", background=[("selected", "#2F6F9F")], foreground=[("selected", "#FFFFFF")])
        style.configure(
            "Horizontal.TProgressbar",
            background=accent,
            troughcolor="#D8E1EC",
            bordercolor="#D8E1EC",
            lightcolor=accent,
            darkcolor=accent,
        )

        header = ttk.Frame(self.root, padding=(18, 16, 18, 10))
        header.pack(fill="x")
        self.header_label = ttk.Label(header, text=self.tr("header_title"), style="Header.TLabel")
        self.header_label.pack(side="left")
        controls = ttk.Frame(header)
        controls.pack(side="right")
        self.form_label = ttk.Label(controls, text=self.tr("form"))
        self.form_label.pack(side="left", padx=(0, 6))
        form_combo = ttk.Combobox(
            controls,
            textvariable=self.form_version,
            values=["USA", "Korea"],
            state="readonly",
            width=10,
        )
        form_combo.pack(side="left", padx=(0, 12))
        form_combo.bind("<<ComboboxSelected>>", lambda _event: self._on_form_version_changed())
        self.language_label = ttk.Label(controls, text=self.tr("language"))
        self.language_label.pack(side="left", padx=(0, 6))
        language_combo = ttk.Combobox(
            controls,
            textvariable=self.language_display,
            values=list(LANGUAGE_LABELS.values()),
            state="readonly",
            width=10,
        )
        language_combo.pack(side="left", padx=(0, 12))
        language_combo.bind("<<ComboboxSelected>>", lambda _event: self._on_language_changed())
        self.ai_concurrency_label = ttk.Label(controls, text=self.tr("ai_batch"))
        self.ai_concurrency_label.pack(side="left", padx=(0, 6))
        self.ai_concurrency_entry = ttk.Entry(controls, textvariable=self.ai_concurrency, width=4)
        self.ai_concurrency_entry.pack(side="left", padx=(0, 12))
        self.ai_concurrency_entry.bind("<FocusOut>", lambda _event: self.ai_concurrency.set(str(self.ai_concurrency_value())))
        self.ai_concurrency_entry.bind("<Return>", lambda _event: self.ai_concurrency.set(str(self.ai_concurrency_value())))
        self.upload_mode_label = ttk.Label(controls, text=self.tr("upload_mode"))
        self.upload_mode_label.pack(side="left", padx=(0, 6))
        self.upload_mode_combo = ttk.Combobox(
            controls,
            textvariable=self.usa_upload_mode_display,
            values=self.upload_mode_values(),
            state="readonly",
            width=16,
        )
        self.upload_mode_combo.pack(side="left", padx=(0, 12))
        self.upload_mode_combo.bind("<<ComboboxSelected>>", self._on_upload_mode_changed)
        self.update_upload_mode_options()
        self.upload_folder_btn = ttk.Button(controls, text=self.tr("select_files"), command=self.select_files)
        self.upload_folder_btn.pack(side="left", padx=4)
        self.select_payment_proof_btn = ttk.Button(
            controls,
            text=self.tr("select_payment_proof"),
            command=self.select_support_files,
        )
        self.select_payment_proof_btn.pack(side="left", padx=4)
        self.generate_details_btn = ttk.Button(controls, text=self.tr("generate_empty"), command=self.generate_empty_details)
        self.generate_details_btn.pack(side="left", padx=4)
        self.generate_all_btn = ttk.Button(controls, text=self.tr("generate_all"), command=self.generate_all_or_cancel)
        self.generate_all_btn.pack(side="left", padx=4)
        self.generate_excel_btn = ttk.Button(controls, text=self.tr("generate_excel"), command=self.generate_excel)
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
        self.manager_label = ttk.Label(manager, text=self.tr("inserted_receipts"), style="PanelHeader.TLabel")
        self.manager_label.grid(
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
        self.tree.heading("#0", text=self.tr("file"))
        self.tree.heading("status", text=self.tr("status"))
        self.tree.heading("date", text=self.tr("date"))
        self.tree.heading("amount", text=self.tr("amount"))
        self.tree.column("#0", width=140, stretch=True)
        self.tree.column("status", width=118, stretch=False)
        self.tree.column("date", width=78, stretch=False)
        self.tree.column("amount", width=70, stretch=False)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Shift-Button-1>", self._on_tree_shift_click)
        self.tree.bind("<Delete>", self.remove_selected)
        self.tree.bind("<BackSpace>", self.remove_selected)
        self.tree.bind("<Control-a>", self.select_all_receipts)
        self.tree.bind("<Control-A>", self.select_all_receipts)
        self.tree.tag_configure("danger", background="#FEE2E2", foreground="#991B1B")
        self.tree.tag_configure("review", background="#FEF3C7", foreground="#92400E")
        left_buttons = ttk.Frame(left, style="Panel.TFrame")
        left_buttons.pack(fill="x", pady=(10, 0))
        self.remove_btn = ttk.Button(left_buttons, text=self.tr("remove"), command=self.remove_selected)
        self.remove_btn.pack(side="left")
        self.clear_btn = ttk.Button(left_buttons, text=self.tr("clear"), command=self.clear_all)
        self.clear_btn.pack(side="left", padx=(8, 0))

        middle = ttk.Frame(manager, style="Panel.TFrame", padding=(10, 0, 0, 0))
        middle.grid(row=1, column=1, sticky="nsew")
        self.details_label = ttk.Label(middle, text=self.tr("details"), style="PanelHeader.TLabel")
        self.details_label.grid(row=0, column=0, columnspan=2, sticky="w")
        rate_frame = ttk.Frame(middle, style="Panel.TFrame")
        rate_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 10))
        self.usa_rate_label = ttk.Label(rate_frame, text="USD -> RMB", style="Panel.TLabel")
        self.usa_rate_label.pack(side="left")
        self.usa_rate_entry = ttk.Entry(rate_frame, textvariable=self.exchange_rate, width=9)
        self.usa_rate_entry.pack(side="left", padx=(6, 14))
        self.usd_krw_rate_label = ttk.Label(rate_frame, text="USD -> KRW", style="Panel.TLabel")
        self.usd_krw_rate_label.pack(side="left")
        self.usd_krw_rate_entry = ttk.Entry(rate_frame, textvariable=self.usd_to_krw_rate, width=9)
        self.usd_krw_rate_entry.pack(side="left", padx=(6, 14))
        self.krw_rate_label = ttk.Label(rate_frame, text="KRW -> RMB", style="Panel.TLabel")
        self.krw_rate_label.pack(side="left")
        self.krw_rate_entry = ttk.Entry(rate_frame, textvariable=self.krw_to_rmb_rate, width=9)
        self.krw_rate_entry.pack(side="left", padx=(6, 0))

        fields = [
            ("date", self.tr("date")),
            ("place", self.tr("place_vendor")),
            ("amount", self.tr("usd_amount")),
            ("currency", self.tr("currency")),
            ("krw_amount", self.tr("krw_amount")),
            ("rmb_amount", self.tr("rmb_amount")),
            ("purpose", self.tr("purpose")),
            ("details", self.tr("details")),
            ("project_number", self.tr("project_number")),
            ("category", self.tr("category")),
            ("payment_method", self.tr("payment_method")),
            ("receipt_label", self.tr("receipt_label")),
            ("document_type", self.tr("document_type")),
        ]
        for row, (key, label) in enumerate(fields, start=2):
            label_widget = ttk.Label(middle, text=label, style="Panel.TLabel")
            label_widget.grid(row=row, column=0, sticky="w", pady=4)
            if key == "category":
                var = tk.StringVar()
                widget = ttk.Combobox(middle, textvariable=var, state="readonly", width=28)
                self.category_combo = widget
            elif key == "document_type":
                var = tk.StringVar()
                widget = ttk.Combobox(middle, textvariable=var, state="readonly", width=28)
                self.document_type_combo = widget
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
        self.preview_label = ttk.Label(preview_header, text=self.tr("receipt_preview"), style="PanelHeader.TLabel")
        self.preview_label.grid(row=0, column=0, sticky="w")
        self.zoom_out_btn = ttk.Button(
            preview_header,
            text="-",
            command=lambda: self.zoom_preview(-1),
            state="disabled",
            style="Icon.TButton",
        )
        self.zoom_out_btn.grid(row=0, column=1, sticky="e", padx=(0, 6))
        ToolTip(self.zoom_out_btn, self.tr("zoom_out"))
        self.zoom_label = ttk.Label(preview_header, text="100%", style="Muted.TLabel", width=5, anchor="center")
        self.zoom_label.grid(row=0, column=2, sticky="e", padx=(0, 6))
        self.zoom_in_btn = ttk.Button(
            preview_header,
            text="+",
            command=lambda: self.zoom_preview(1),
            state="disabled",
            style="Icon.TButton",
        )
        self.zoom_in_btn.grid(row=0, column=3, sticky="e", padx=(0, 6))
        ToolTip(self.zoom_in_btn, self.tr("zoom_in"))
        self.rotate_left_btn = ttk.Button(
            preview_header,
            text="\u27F2",
            command=lambda: self.rotate_selected(-90),
            state="disabled",
            style="Icon.TButton",
        )
        self.rotate_left_btn.grid(row=0, column=4, sticky="e", padx=(0, 6))
        ToolTip(self.rotate_left_btn, self.tr("rotate_left"))
        self.rotate_right_btn = ttk.Button(
            preview_header,
            text="\u27F3",
            command=lambda: self.rotate_selected(90),
            state="disabled",
            style="Icon.TButton",
        )
        self.rotate_right_btn.grid(row=0, column=5, sticky="e", padx=(0, 6))
        ToolTip(self.rotate_right_btn, self.tr("rotate_right"))
        self.revert_crop_btn = ttk.Button(
            preview_header,
            text=self.tr("revert"),
            command=self.revert_crop,
            state="disabled",
        )
        self.revert_crop_btn.grid(row=0, column=6, sticky="e")
        ToolTip(self.revert_crop_btn, self.tr("revert"))
        self.delete_screenshot_btn = ttk.Button(
            preview_header,
            text="\u2715",
            command=self.delete_selected_screenshot,
            state="disabled",
            style="Icon.TButton",
        )
        self.delete_screenshot_btn.grid(row=0, column=7, sticky="e", padx=(6, 0))
        ToolTip(self.delete_screenshot_btn, self.tr("delete_screenshot"))
        preview_frame = ttk.Frame(right, style="Panel.TFrame")
        preview_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        self.preview_canvas = tk.Canvas(
            preview_frame,
            bg=canvas_bg,
            highlightthickness=1,
            highlightbackground=border,
            bd=1,
            relief="solid",
        )
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")
        preview_v_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_canvas.yview)
        preview_v_scroll.grid(row=0, column=1, sticky="ns")
        preview_h_scroll = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.preview_canvas.xview)
        preview_h_scroll.grid(row=1, column=0, sticky="ew")
        self.preview_canvas.configure(
            yscrollcommand=preview_v_scroll.set,
            xscrollcommand=preview_h_scroll.set,
        )
        self.preview_canvas.bind("<Configure>", lambda _event: self.update_preview())
        self.preview_canvas.bind("<ButtonPress-1>", self._on_crop_press)
        self.preview_canvas.bind("<B1-Motion>", self._on_crop_drag)
        self.preview_canvas.bind("<ButtonRelease-1>", self._on_crop_release)
        self.preview_canvas.bind("<MouseWheel>", self._on_preview_mousewheel)
        for sequence in ("<Control-c>", "<Control-C>"):
            self.root.bind_all(sequence, lambda event, mode="copy": self.copy_screenshot_shortcut(event, mode), add="+")
        for sequence in ("<Control-x>", "<Control-X>"):
            self.root.bind_all(sequence, lambda event, mode="cut": self.copy_screenshot_shortcut(event, mode), add="+")
        for sequence in ("<Control-v>", "<Control-V>"):
            self.root.bind_all(sequence, self.paste_screenshot_shortcut, add="+")

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
        self.open_output_btn = ttk.Button(footer, text=self.tr("open_output_folder"), command=self.open_output_folder)
        self.open_output_btn.pack(side="right")

    def _set_initial_window_geometry(self) -> None:
        try:
            screen_width = max(900, int(self.root.winfo_screenwidth()))
            screen_height = max(600, int(self.root.winfo_screenheight()))
        except Exception:
            screen_width, screen_height = 1280, 800
        width = min(1220, max(900, screen_width - 32))
        height = min(760, max(560, screen_height - 100))
        x = max(0, (screen_width - width) // 2)
        y = max(0, min((screen_height - height) // 2, screen_height - height - 40))
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.minsize(min(980, width), min(640, height))

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
            "document_type",
        ):
            var = self.field_vars.get(key)
            if var is not None:
                var.trace_add("write", lambda *_args, k=key: self._on_field_changed(k))
        self.exchange_rate.trace_add("write", lambda *_args: self._sync_amount_fields("rate"))
        self.usd_to_krw_rate.trace_add("write", lambda *_args: self._sync_amount_fields("rate"))
        self.krw_to_rmb_rate.trace_add("write", lambda *_args: self._sync_amount_fields("rate"))

    def category_display_label(self, category: str) -> str:
        labels = CATEGORY_LABELS_BY_LANGUAGE.get(self.language.get(), CATEGORY_LABELS_BY_LANGUAGE["en"])
        return labels.get(category, category)

    def category_key_from_display(self, value: str) -> str:
        return self.category_label_to_key.get(value, category_value_to_key(value, self.form_version.get()))

    def document_type_display_label(self, document_type: str) -> str:
        key = normalize_document_type(document_type, "unknown")
        return self.tr(DOCUMENT_TYPE_LABEL_KEYS.get(key, "document_type_unknown"))

    def document_type_key_from_display(self, value: str) -> str:
        text = str(value or "").strip()
        for key, text_key in DOCUMENT_TYPE_LABEL_KEYS.items():
            if text == self.tr(text_key):
                return key
        return normalize_document_type(text, "unknown")

    def document_type_values(self) -> List[str]:
        return [self.document_type_display_label(key) for key in ("unknown", "receipt", "payment_proof")]

    def upload_mode_display_label(self, mode: str) -> str:
        key = normalize_usa_upload_mode(mode)
        text_key = {
            "auto": "upload_mode_auto",
            "receipt": "upload_mode_receipt",
            "payment_proof": "upload_mode_payment_proof",
        }.get(key, "upload_mode_auto")
        return self.tr(text_key)

    def upload_mode_key_from_display(self, value: str) -> str:
        text = str(value or "").strip()
        for key in ("auto", "receipt", "payment_proof"):
            if text == self.upload_mode_display_label(key):
                return key
        return normalize_usa_upload_mode(text)

    def upload_mode_values(self) -> List[str]:
        return [self.upload_mode_display_label(key) for key in ("auto", "receipt", "payment_proof")]

    def update_upload_mode_options(self) -> None:
        if self.upload_mode_combo is None:
            return
        current = normalize_usa_upload_mode(self.usa_upload_mode.get())
        self.upload_mode_combo.configure(values=self.upload_mode_values())
        self.usa_upload_mode_display.set(self.upload_mode_display_label(current))

    def _on_upload_mode_changed(self, _event: Any = None) -> None:
        self.usa_upload_mode.set(self.upload_mode_key_from_display(self.usa_upload_mode_display.get()))
        self.settings["usa_upload_mode"] = self.usa_upload_mode.get()
        save_user_settings(self.settings)
        self.update_upload_mode_options()

    def attention_reason_label(self, reason: Optional[str]) -> str:
        mapping = {
            "AI failed": "attention_ai_failed",
            "Unknown document type": "attention_unknown_type",
            "Missing date": "attention_missing_date",
            "Missing amount": "attention_missing_amount",
            "Missing vendor": "attention_missing_vendor",
            "Missing details": "attention_missing_details",
            "Missing date/amount": "attention_missing_date_amount",
            "Wrong year": "attention_wrong_year",
        }
        key = mapping.get(reason or "")
        return self.tr(key) if key else str(reason or "")

    def update_document_type_options(self) -> None:
        if self.document_type_combo is None:
            return
        item = self.selected_item()
        if item is not None:
            current_key = normalize_document_type(item.document_type, "unknown")
        else:
            doc_var = self.field_vars.get("document_type")
            current_key = self.document_type_key_from_display(doc_var.get()) if doc_var is not None else "unknown"
        values = self.document_type_values()
        self.document_type_combo.configure(values=values)
        doc_var = self.field_vars.get("document_type")
        if doc_var is not None:
            was_loading = self._loading_fields
            self._loading_fields = True
            try:
                doc_var.set(self.document_type_display_label(current_key))
            finally:
                self._loading_fields = was_loading

    def _on_language_changed(self) -> None:
        chosen_label = self.language_display.get()
        for code, label in LANGUAGE_LABELS.items():
            if label == chosen_label:
                self.language.set(code)
                break
        else:
            self.language.set("en")
            self.language_display.set(LANGUAGE_LABELS["en"])
        self.settings["language"] = self.language.get()
        save_user_settings(self.settings)
        self.refresh_language_texts()

    def refresh_language_texts(self) -> None:
        self.root.title(self.tr("app_title"))
        for widget, key in (
            (self.header_label, "header_title"),
            (self.form_label, "form"),
            (self.language_label, "language"),
            (self.ai_concurrency_label, "ai_batch"),
            (self.upload_mode_label, "upload_mode"),
            (self.upload_folder_btn, "select_files"),
            (self.generate_details_btn, "generate_empty"),
            (self.generate_all_btn, "generate_all"),
            (self.generate_excel_btn, "generate_excel"),
            (self.manager_label, "inserted_receipts"),
            (self.details_label, "details"),
            (self.preview_label, "receipt_preview"),
            (self.remove_btn, "remove"),
            (self.clear_btn, "clear"),
            (self.revert_crop_btn, "revert"),
            (self.open_output_btn, "open_output_folder"),
        ):
            if widget is not None:
                try:
                    widget.configure(text=self.tr(key))
                except Exception:
                    pass
        if self._ai_cancel_event is not None and self.generate_all_btn is not None:
            try:
                self.generate_all_btn.configure(text=self.tr("cancel_generation"))
            except Exception:
                pass
        self.tree.heading("#0", text=self.tr("file"))
        self.tree.heading("status", text=self.tr("status"))
        self.tree.heading("date", text=self.tr("date"))
        self.tree.heading("amount", text=self.tr("amount"))
        for key, text_key in {
            "date": "date",
            "place": "place_vendor",
            "currency": "currency",
            "krw_amount": "krw_amount",
            "rmb_amount": "rmb_amount",
            "purpose": "purpose",
            "details": "details",
            "project_number": "project_number",
            "category": "category",
            "payment_method": "payment_method",
            "receipt_label": "receipt_label",
            "document_type": "document_type",
        }.items():
            label = self.field_labels.get(key)
            if label is not None:
                label.configure(text=self.tr(text_key))
        self.update_document_type_options()
        self.update_upload_mode_options()
        self._update_field_visibility()
        self._on_form_version_changed(refresh_language_only=True)
        if self.status_text.get() in {ui_text(code, "ready") for code in UI_TEXT}:
            self.status_text.set(self.tr("ready"))
        self.update_preview()

    def ai_concurrency_value(self) -> int:
        value = clamp_ai_concurrency(self.ai_concurrency.get())
        if self.ai_concurrency.get() != str(value):
            self.ai_concurrency.set(str(value))
        return value

    def confirm_ai_concurrency(self, requested: int, receipt_count: int, bank_count: int = 0) -> bool:
        active = min(requested, max(1, receipt_count, bank_count))
        if active <= AI_REQUEST_CONCURRENCY_WARNING_THRESHOLD:
            return True
        message = self.tr("high_ai_concurrency_warning", requested=requested, active=active)
        if messagebox:
            return bool(messagebox.askyesno(self.tr("high_ai_concurrency_title"), message))
        self.status_text.set(message)
        return True

    def remembered_file_dialog_dir(self) -> str:
        candidates = [
            self.settings.get("last_file_dialog_dir"),
            self.settings.get("last_open_dir"),
            str(APP_DIR),
        ]
        for candidate in candidates:
            if not candidate:
                continue
            try:
                path = Path(str(candidate)).expanduser()
                if path.exists() and path.is_dir():
                    return str(path)
            except Exception:
                continue
        return str(APP_DIR)

    def remember_file_dialog_dir(self, paths: Sequence[Path]) -> None:
        for path in paths:
            try:
                folder = path.parent if path.is_file() or path.suffix else path
                if folder.exists() and folder.is_dir():
                    self.settings["last_file_dialog_dir"] = str(folder)
                    save_user_settings(self.settings)
                    return
            except Exception:
                continue

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
                usd_rmb_rate = safe_float(self.exchange_rate.get())
                usd_krw_rate = safe_float(self.usd_to_krw_rate.get())
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
                            if usd_krw_rate:
                                krw = amount * usd_krw_rate
                                rmb = krw * krw_rate
                            elif usd_rmb_rate:
                                rmb = amount * usd_rmb_rate
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
                        elif currency == "USD" and usd_krw_rate and not amount_var.get().strip():
                            amount_var.set(format_amount(krw / usd_krw_rate))
                        elif currency == "USD" and usd_rmb_rate and not amount_var.get().strip():
                            amount_var.set(format_amount(rmb / usd_rmb_rate))
                    elif source == "krw_amount" and krw is not None:
                        rmb = krw * krw_rate
                        if currency == "KRW":
                            amount_var.set(format_amount(krw))
                        elif currency in {"RMB", "CNY"} and not amount_var.get().strip():
                            amount_var.set(format_amount(rmb))
                        elif currency == "USD" and usd_krw_rate and not amount_var.get().strip():
                            amount_var.set(format_amount(krw / usd_krw_rate))
                        elif currency == "USD" and usd_rmb_rate and not amount_var.get().strip():
                            amount_var.set(format_amount(rmb / usd_rmb_rate))

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

    def preview_zoom_value(self) -> float:
        try:
            value = float(self.preview_zoom.get())
        except Exception:
            value = 1.0
        return max(0.5, min(3.0, value))

    def update_zoom_label(self) -> None:
        if self.zoom_label is not None:
            self.zoom_label.configure(text=f"{int(round(self.preview_zoom_value() * 100))}%")

    def set_preview_zoom(self, value: float) -> None:
        old_value = self.preview_zoom_value()
        new_value = max(0.5, min(3.0, value))
        if abs(old_value - new_value) < 0.001:
            return
        self.preview_zoom.set(new_value)
        self.update_zoom_label()
        self.update_preview()

    def zoom_preview(self, direction: int) -> None:
        factor = 1.25 if direction > 0 else 0.8
        self.set_preview_zoom(self.preview_zoom_value() * factor)

    def preview_event_coords(self, event: Any) -> Tuple[float, float]:
        if self.preview_canvas is None:
            return float(getattr(event, "x", 0)), float(getattr(event, "y", 0))
        return (
            float(self.preview_canvas.canvasx(getattr(event, "x", 0))),
            float(self.preview_canvas.canvasy(getattr(event, "y", 0))),
        )

    def _on_preview_mousewheel(self, event: Any) -> str:
        if self.preview_canvas is None:
            return "break"
        if getattr(event, "state", 0) & 0x0004:
            self.zoom_preview(1 if getattr(event, "delta", 0) > 0 else -1)
            return "break"
        units = -1 * int(getattr(event, "delta", 0) / 120)
        self.preview_canvas.yview_scroll(units, "units")
        return "break"

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
        usd_rmb_rate = safe_float(self.exchange_rate.get()) or safe_float(DEFAULT_USD_TO_RMB_RATE)
        usd_krw_rate = safe_float(self.usd_to_krw_rate.get()) or safe_float(DEFAULT_USD_TO_KRW_RATE)
        if not krw_rate:
            return
        item.currency = normalize_currency(item.currency, "KRW")
        amount = safe_float(item.amount)
        krw = safe_float(item.krw_amount)
        rmb = safe_float(item.rmb_amount)
        item_source = "amount" if source in {"currency", "rate"} and amount is not None else source
        if item_source == "amount" and amount is not None:
            if item.currency == "USD":
                krw = amount * usd_krw_rate if usd_krw_rate else krw
                rmb = krw * krw_rate if krw is not None else (amount * usd_rmb_rate if usd_rmb_rate else rmb)
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
            elif item.currency == "USD" and amount is None and usd_krw_rate:
                item.amount = format_amount(krw / usd_krw_rate)
            elif item.currency == "USD" and amount is None and usd_rmb_rate:
                item.amount = format_amount(rmb / usd_rmb_rate)
        elif item_source == "krw_amount" and krw is not None:
            rmb = krw * krw_rate
            if item.currency == "KRW":
                item.amount = format_amount(krw)
            elif item.currency in {"RMB", "CNY"} and amount is None:
                item.amount = format_amount(rmb)
            elif item.currency == "USD" and amount is None and usd_krw_rate:
                item.amount = format_amount(krw / usd_krw_rate)
            elif item.currency == "USD" and amount is None and usd_rmb_rate:
                item.amount = format_amount(rmb / usd_rmb_rate)

        if krw is None and rmb is not None:
            krw = rmb / krw_rate
        if rmb is None and krw is not None:
            rmb = krw * krw_rate
        if krw is not None:
            item.krw_amount = format_amount(krw)
        if rmb is not None:
            item.rmb_amount = format_amount(rmb)

    def clear_attention_status_if_ready(self, item: ReceiptItem) -> None:
        if receipt_attention_reason(item, self.form_version.get()) is not None:
            return
        if item.status in {"Empty", "Failed", "Missing date/amount", "Missing date", "Missing amount", "Missing details", "Missing vendor", "Wrong year"}:
            item.status = "Payment proof" if is_usa_payment_proof_item(item) else "Manual filled"

    def update_date_field_attention(self) -> None:
        widget = self.field_widgets.get("date")
        if widget is None:
            return
        item = self.selected_item()
        style = "Warning.TEntry" if item is not None and date_year_needs_review(item.date) else "TEntry"
        try:
            widget.configure(style=style)
        except Exception:
            pass

    def _save_field_to_selected(self, key: str, refresh: bool = False) -> None:
        var = self.field_vars.get(key)
        if var is None:
            return
        value = var.get()
        version = self.form_version.get()
        for index in self.selected_indices():
            item = self.items[index]
            before = receipt_item_signature(item)
            if key == "category":
                setattr(item, key, self.category_key_from_display(value))
            elif key == "document_type":
                item.document_type = normalize_document_type(self.document_type_key_from_display(value), "unknown")
                item.locked_document_type = "" if item.document_type == "unknown" else item.document_type
                if item.document_type != "payment_proof":
                    item.matched_receipt_id = ""
            else:
                setattr(item, key, value)
            self.update_item_amount_fields(item, self._last_amount_source)
            self.clear_attention_status_if_ready(item)
            if receipt_item_signature(item) != before:
                self.remember_receipt_item(item, "manual")
        if refresh:
            self.refresh_tree()
        if key == "date":
            self.update_date_field_attention()

    def _save_amount_fields_to_selected(self, source: str, refresh: bool = False) -> None:
        version = self.form_version.get()
        amount_value = self.field_vars["amount"].get()
        currency_value = normalize_currency(self.field_vars["currency"].get(), "USD" if version == "USA" else "KRW")
        krw_value = self.field_vars["krw_amount"].get()
        rmb_value = self.field_vars["rmb_amount"].get()
        indices = range(len(self.items)) if source == "rate" else self.selected_indices()
        for index in indices:
            item = self.items[index]
            before = receipt_item_signature(item)
            if source == "rate":
                self.update_item_amount_fields(item, "amount")
                if receipt_item_signature(item) != before:
                    self.remember_receipt_item(item, "manual")
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
            self.clear_attention_status_if_ready(item)
            if receipt_item_signature(item) != before:
                self.remember_receipt_item(item, "manual")
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
            if self.upload_folder_btn is not None:
                self.upload_folder_btn.configure(text=self.tr("select_receipts"))
            if self.upload_mode_label is not None:
                try:
                    self.upload_mode_label.pack(side="left", padx=(0, 6), before=self.upload_folder_btn)
                except Exception:
                    self.upload_mode_label.pack(side="left", padx=(0, 6))
            if self.upload_mode_combo is not None:
                try:
                    self.upload_mode_combo.pack(side="left", padx=(0, 12), before=self.upload_folder_btn)
                except Exception:
                    self.upload_mode_combo.pack(side="left", padx=(0, 12))
            self.field_labels["amount"].configure(text=self.tr("usd_amount"))
            self.field_labels["rmb_amount"].configure(text=self.tr("rmb_amount"))
            self.field_vars["currency"].set("USD")
            self._set_field_visible("currency", False)
            self._set_field_visible("krw_amount", False)
            self._set_field_visible("amount", True)
            self._set_field_visible("rmb_amount", True)
            self._set_field_visible("document_type", True)
            if self.usa_rate_label is not None:
                self.usa_rate_label.pack(side="left")
            if self.usa_rate_entry is not None:
                self.usa_rate_entry.pack(side="left", padx=(6, 14))
            if self.usd_krw_rate_label is not None:
                self.usd_krw_rate_label.pack_forget()
            if self.usd_krw_rate_entry is not None:
                self.usd_krw_rate_entry.pack_forget()
            if self.krw_rate_label is not None:
                self.krw_rate_label.pack_forget()
            if self.krw_rate_entry is not None:
                self.krw_rate_entry.pack_forget()
            if self.select_payment_proof_btn is not None:
                self.select_payment_proof_btn.pack_forget()
        else:
            if self.upload_folder_btn is not None:
                self.upload_folder_btn.configure(text=self.tr("select_files"))
            if self.upload_mode_label is not None:
                self.upload_mode_label.pack_forget()
            if self.upload_mode_combo is not None:
                self.upload_mode_combo.pack_forget()
            self.field_labels["amount"].configure(text=self.tr("original_amount"))
            self.field_labels["rmb_amount"].configure(text=self.tr("rmb_amount"))
            self._set_field_visible("currency", True)
            self._set_field_visible("krw_amount", True)
            self._set_field_visible("amount", True)
            self._set_field_visible("rmb_amount", True)
            self._set_field_visible("document_type", False)
            if self.usa_rate_label is not None:
                self.usa_rate_label.pack_forget()
            if self.usa_rate_entry is not None:
                self.usa_rate_entry.pack_forget()
            if self.usd_krw_rate_label is not None:
                self.usd_krw_rate_label.pack(side="left")
            if self.usd_krw_rate_entry is not None:
                self.usd_krw_rate_entry.pack(side="left", padx=(6, 14))
            if self.krw_rate_label is not None:
                self.krw_rate_label.pack(side="left")
            if self.krw_rate_entry is not None:
                self.krw_rate_entry.pack(side="left", padx=(6, 0))
            if self.select_payment_proof_btn is not None:
                self.select_payment_proof_btn.configure(text=self.tr("select_exchange_rate"), command=self.select_support_files, state="normal")
                try:
                    self.select_payment_proof_btn.pack(side="left", padx=4, before=self.generate_details_btn)
                except Exception:
                    self.select_payment_proof_btn.pack(side="left", padx=4)

    def suggested_toolbar_action(self) -> str:
        if self._busy:
            return ""
        if self.items and self._details_ready_for_export:
            return "generate_excel"
        if not self.items:
            return "select_files"
        if self.form_version.get() == "Korea" and not self.exchange_rate_items:
            return "select_payment_proof"
        if self.form_version.get() == "USA":
            needs_ai = any(receipt_needs_ai(item, "USA") for item in self.items)
            has_finished_rows = any(not receipt_needs_ai(item, "USA") for item in self.items)
            return "generate_empty" if needs_ai and has_finished_rows else "generate_all"
        if any(receipt_needs_ai(item, self.form_version.get()) for item in self.items):
            return "generate_empty"
        return "generate_all"

    def update_toolbar_recommendation(self) -> None:
        suggested = self.suggested_toolbar_action()
        buttons = {
            "select_files": self.upload_folder_btn,
            "select_payment_proof": self.select_payment_proof_btn,
            "generate_empty": self.generate_details_btn,
            "generate_all": self.generate_all_btn,
            "generate_excel": self.generate_excel_btn,
        }
        for action, button in buttons.items():
            if button is None:
                continue
            try:
                button.configure(style="Recommended.TButton" if action == suggested else "TButton")
            except Exception:
                pass

    def _on_form_version_changed(self, refresh_language_only: bool = False) -> None:
        version = normalized_form_version(self.form_version.get())
        if version != self.form_version.get():
            self.form_version.set(version)
        form_changed = version != self._last_form_version and not refresh_language_only
        if form_changed:
            self._details_ready_for_export = False
            self.settings["last_form_version"] = version
            save_user_settings(self.settings)
        cat_var = self.field_vars.get("category")
        current_key = self.category_key_from_display(cat_var.get()) if cat_var is not None else ""
        order = KOREA_REPORT_CATEGORY_ORDER if version == "Korea" else USA_REPORT_CATEGORY_ORDER
        self.category_values = [self.category_display_label(key) for key in order]
        self.category_label_to_key = {
            self.category_display_label(key): key
            for key in order
        }
        if self.category_combo is not None:
            self.category_combo.configure(values=self.category_values)
        self.update_document_type_options()
        cat = self.field_vars.get("category")
        if cat is not None:
            if current_key in order:
                cat.set(self.category_display_label(current_key))
            elif not cat.get():
                cat.set(self.category_values[0] if self.category_values else "transportation")
        for item in self.items:
            if version == "USA":
                item.currency = "USD"
                item.document_type = normalize_document_type(item.document_type, "receipt")
                locked = normalize_document_type(item.locked_document_type, "unknown")
                item.locked_document_type = locked if locked != "unknown" else ""
            else:
                item.currency = normalize_currency(item.currency, "KRW")
                item.document_type = "receipt"
                item.locked_document_type = ""
                item.matched_receipt_id = ""
            self.update_item_amount_fields(item, "amount")
            if form_changed and item.status == "AI filled":
                item.status = "Needs regen"
        self._update_field_visibility()
        if self.selected_index is not None:
            self.load_selected_into_fields()
        self._sync_amount_fields("rate")
        self.refresh_tree()
        self.refresh_bank_tree()
        self._last_form_version = version
        self.update_toolbar_recommendation()

    def select_files(self) -> None:
        ensure_runtime_folders()
        if filedialog is None:
            return
        if self._receipt_import_active:
            self.show_info("Receipt files are still being added. Please wait a moment.")
            return
        selected = filedialog.askopenfilenames(
            title="Select receipt files",
            initialdir=self.remembered_file_dialog_dir(),
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
        self.remember_file_dialog_dir(paths)
        self.save_current_fields()
        self._details_ready_for_export = False
        upload_mode = normalize_usa_upload_mode(self.usa_upload_mode.get())
        self.settings["usa_upload_mode"] = upload_mode
        save_user_settings(self.settings)
        self.start_incremental_receipt_import(paths, upload_mode)

    def apply_usa_upload_mode_to_item(self, item: ReceiptItem, upload_mode: str) -> None:
        if self.form_version.get() != "USA":
            return
        mode = normalize_usa_upload_mode(upload_mode)
        if mode == "auto":
            return
        item.document_type = mode
        item.locked_document_type = mode
        if mode == "payment_proof":
            item.status = "Payment proof"
        else:
            item.matched_receipt_id = ""

    def receipt_memory_records_for_item(self, item: ReceiptItem) -> Tuple[Optional[Dict[str, Any]], Optional[Dict[str, Any]]]:
        records = self._receipt_memory.get("records")
        if not isinstance(records, dict):
            records = {}
            self._receipt_memory["records"] = records
        for key in receipt_memory_keys_for_item(item):
            record = records.get(key)
            if isinstance(record, dict):
                return record, None
        marker_sources = [item.filename, item.source_path, item.path]
        for attachment in item.receipt_images:
            marker_sources.extend(
                [
                    str(attachment.get("filename") or ""),
                    str(attachment.get("source_path") or ""),
                    str(attachment.get("path") or ""),
                ]
            )
        for value in marker_sources:
            marker = parse_receipt_memory_marker(Path(str(value)).name)
            if marker:
                return None, marker
        return None, None

    def apply_stored_receipt_memory(self, item: ReceiptItem) -> bool:
        full_record, marker_record = self.receipt_memory_records_for_item(item)
        if full_record:
            changed = apply_receipt_memory_record_to_item(item, full_record, self.form_version.get(), full_record=True)
            self.update_item_amount_fields(item, "amount")
            return changed
        if marker_record:
            changed = apply_receipt_memory_record_to_item(item, marker_record, self.form_version.get(), full_record=False)
            self.update_item_amount_fields(item, "amount")
            if item.status == "Empty":
                item.status = "Filename hint"
            return changed
        return False

    def source_paths_for_receipt_item(self, item: ReceiptItem) -> List[Path]:
        paths: List[Path] = []
        for raw_path in (item.source_path, item.path):
            if raw_path:
                paths.append(Path(raw_path))
        for attachment in ensure_receipt_images(item):
            for raw_path in (attachment.get("source_path"), attachment.get("path")):
                if raw_path:
                    paths.append(Path(str(raw_path)))
        unique: List[Path] = []
        seen: set = set()
        for path in paths:
            try:
                resolved = str(path.resolve()).lower()
            except Exception:
                resolved = str(path).lower()
            if resolved in seen:
                continue
            seen.add(resolved)
            try:
                in_work_dir = path.resolve().is_relative_to(WORK_DIR.resolve())
            except Exception:
                in_work_dir = False
            if path.exists() and path.is_file() and not in_work_dir:
                unique.append(path)
        return unique

    def tagged_receipt_source_name(self, source: Path, record: Dict[str, Any], fingerprint: str) -> str:
        marker = build_receipt_memory_marker(record, fingerprint)
        base_stem = strip_receipt_memory_marker(source.stem)
        max_stem_len = max(12, 180 - len(marker) - len(source.suffix))
        base_stem = base_stem[:max_stem_len].rstrip(" ._-") or "receipt"
        return f"{base_stem}{marker}{source.suffix}"

    def apply_source_rename_to_item(self, item: ReceiptItem, old_path: Path, new_path: Path) -> None:
        old_text = str(old_path)
        new_text = str(new_path)

        def same_path_text(value: Any) -> bool:
            try:
                return str(Path(str(value)).resolve()).lower() == str(old_path.resolve()).lower()
            except Exception:
                return str(value or "").lower() == old_text.lower()

        if same_path_text(item.path):
            item.path = new_text
        if same_path_text(item.source_path):
            item.source_path = new_text
        for attachment in item.receipt_images:
            if same_path_text(attachment.get("path")):
                attachment["path"] = new_text
            source_changed = same_path_text(attachment.get("source_path"))
            if source_changed:
                attachment["source_path"] = new_text
            if same_path_text(attachment.get("path")) or source_changed:
                page = str(attachment.get("source_page") or "").strip()
                attachment["filename"] = f"{new_path.name} page {page}" if page else new_path.name
        ensure_receipt_images(item)

    def rename_receipt_sources_with_memory(self, item: ReceiptItem, record: Dict[str, Any]) -> None:
        renamed_any = False
        for source in self.source_paths_for_receipt_item(item):
            try:
                fingerprint = file_sha256_short(source, 12)
                new_name = self.tagged_receipt_source_name(source, record, fingerprint)
                if source.name == new_name:
                    continue
                destination = source.with_name(new_name)
                if destination.exists() and str(destination.resolve()).lower() != str(source.resolve()).lower():
                    destination = unique_path(source.parent, new_name)
                source.rename(destination)
                self.apply_source_rename_to_item(item, source, destination)
                renamed_any = True
            except Exception:
                LOGGER.warning("Could not rename remembered receipt source: %s", source, exc_info=True)
        if renamed_any:
            ensure_receipt_images(item)

    def remember_receipt_item(self, item: ReceiptItem, reviewed_source: str) -> None:
        record = receipt_memory_record_from_item(item, self.form_version.get(), reviewed_source)
        if not receipt_memory_record_has_values(record):
            return
        keys = receipt_memory_keys_for_item(item)
        records = self._receipt_memory.setdefault("records", {})
        if not isinstance(records, dict):
            records = {}
            self._receipt_memory["records"] = records
        for key in keys:
            records[key] = dict(record)
        self.rename_receipt_sources_with_memory(item, record)
        for key in receipt_memory_keys_for_item(item):
            records[key] = dict(record)
        save_receipt_memory(self._receipt_memory)

    def import_receipt_file(self, path: Path, existing: set, upload_mode: str = "auto") -> Tuple[bool, Optional[str]]:
        suffix = path.suffix.lower()
        if suffix not in SUPPORTED_UPLOAD_SUFFIXES:
            return False, None
        if suffix == ".pdf":
            try:
                page_paths = render_pdf_pages(path)
            except Exception as exc:
                log_exception(f"Could not import PDF: {path}")
                return False, f"{path.name}: {exc}"
            if not page_paths:
                return False, f"{path.name}: PDF has no pages"
            key = selected_file_key(page_paths[0], str(path), "1")
            if key in existing:
                return False, None
            filename = f"{path.name} ({len(page_paths)} pages)"
            receipt_images = []
            for page_index, page_path in enumerate(page_paths, start=1):
                receipt_images.append(
                    {
                        "id": uuid.uuid4().hex,
                        "path": str(page_path),
                        "filename": f"{path.name} page {page_index}",
                        "source_path": str(path),
                        "source_page": str(page_index),
                        "crop_box": None,
                        "rotation_degrees": 0,
                    }
                )
            item = ReceiptItem(
                item_id=uuid.uuid4().hex,
                path=str(page_paths[0]),
                filename=filename,
                source_path=str(path),
                source_page="1",
                currency="USD" if self.form_version.get() == "USA" else "KRW",
                category="transportation",
                document_type="unknown" if self.form_version.get() == "USA" else "receipt",
                receipt_images=receipt_images,
            )
            self.apply_stored_receipt_memory(item)
            self.apply_usa_upload_mode_to_item(item, upload_mode)
            self.items.append(item)
            existing.add(key)
            return True, None
        key = selected_file_key(path, str(path), "")
        if key in existing:
            return False, None
        item = ReceiptItem(
            item_id=uuid.uuid4().hex,
            path=str(path),
            filename=path.name,
            source_path=str(path),
            currency="USD" if self.form_version.get() == "USA" else "KRW",
            category="transportation",
            document_type="unknown" if self.form_version.get() == "USA" else "receipt",
            receipt_images=[
                {
                    "id": uuid.uuid4().hex,
                    "path": str(path),
                    "filename": path.name,
                    "source_path": str(path),
                    "source_page": "",
                    "crop_box": None,
                    "rotation_degrees": 0,
                }
            ],
        )
        self.apply_stored_receipt_memory(item)
        self.apply_usa_upload_mode_to_item(item, upload_mode)
        self.items.append(item)
        existing.add(key)
        return True, None

    def start_incremental_receipt_import(self, paths: List[Path], upload_mode: str = "auto") -> None:
        existing = {
            selected_file_key(Path(item.path), item.source_path, item.source_page)
            for item in [*self.items, *self.bank_items]
            if item.path
        }
        total = len(paths)
        if total <= 0:
            return
        state = {"index": 0, "added": 0}
        failed: List[str] = []
        first_new_index = len(self.items)
        self._receipt_import_active = True
        self.status_text.set(f"Adding receipt files 0/{total}...")
        self.set_progress(0, total, f"0/{total}")

        def step() -> None:
            processed = 0
            started = time.perf_counter()
            while state["index"] < total and processed < 4 and time.perf_counter() - started < 0.08:
                path = paths[state["index"]]
                state["index"] += 1
                added, error_text = self.import_receipt_file(path, existing, upload_mode)
                if added:
                    state["added"] += 1
                if error_text:
                    failed.append(error_text)
                processed += 1
            self.refresh_tree()
            self.set_progress(state["index"], total, f"{state['index']}/{total}")
            self.status_text.set(f"Adding receipt files {state['index']}/{total}...")
            if state["index"] < total:
                self.root.after(10, step)
                return
            self._receipt_import_active = False
            if state["added"] and self.items and self.selected_index is None:
                self.select_index(first_new_index)
            self.save_session()
            self.status_text.set(f"Added {state['added']} receipt file(s).")
            self.update_toolbar_recommendation()
            if failed and messagebox:
                messagebox.showwarning(
                    "Select Files",
                    "Some files could not be imported.\n\n"
                    + "\n".join(failed[:5])
                    + f"\n\nLogged to:\n{LOG_DIR / 'app.log'}",
                )

        self.root.after(0, step)

    def upload_folder(self) -> None:
        self.select_files()

    def select_support_files(self) -> None:
        if self.form_version.get() == "USA":
            self.select_payment_proofs()
        else:
            self.select_exchange_rate_images()

    def select_payment_proofs(self) -> None:
        if self.form_version.get() != "USA":
            self.show_info("Payment proof images are only used for the USA form.")
            return
        ensure_runtime_folders()
        if filedialog is None:
            return
        selected = filedialog.askopenfilenames(
            title="Select payment proof files",
            initialdir=self.remembered_file_dialog_dir(),
            filetypes=[
                ("Payment proof files", "*.png *.jpg *.jpeg *.webp *.bmp *.gif *.pdf"),
                ("Image files", "*.png *.jpg *.jpeg *.webp *.bmp *.gif"),
                ("PDF files", "*.pdf"),
                ("All files", "*.*"),
            ],
        )
        paths = [Path(path) for path in selected]
        if not paths:
            return
        self.remember_file_dialog_dir(paths)
        self._details_ready_for_export = False
        existing = {
            selected_file_key(Path(item.path), item.source_path, item.source_page)
            for item in self.bank_items
            if item.path
        }
        added = 0
        failed: List[str] = []
        for path in paths:
            suffix = path.suffix.lower()
            if suffix not in SUPPORTED_UPLOAD_SUFFIXES:
                continue
            if suffix == ".pdf":
                try:
                    merged_path = render_pdf_to_merged_image(path)
                except Exception as exc:
                    log_exception(f"Could not import payment proof PDF: {path}")
                    failed.append(f"{path.name}: {exc}")
                    continue
                key = selected_file_key(merged_path, str(path), "merged")
                if key in existing:
                    continue
                self.bank_items.append(
                    BankStatementItem(
                        item_id=uuid.uuid4().hex,
                        path=str(merged_path),
                        filename=f"{path.name} (all pages)",
                        source_path=str(path),
                        source_page="merged",
                    )
                )
                existing.add(key)
                added += 1
                continue
            key = selected_file_key(path, str(path), "")
            if key in existing:
                continue
            self.bank_items.append(
                BankStatementItem(
                    item_id=uuid.uuid4().hex,
                    path=str(path),
                    filename=path.name,
                    source_path=str(path),
                )
            )
            existing.add(key)
            added += 1
        self.save_session()
        self.status_text.set(f"Added {added} payment proof file(s).")
        self.update_toolbar_recommendation()
        if self.selected_item() is not None:
            self.update_preview()
        if failed and messagebox:
            messagebox.showwarning(
                "Select Payment Proof",
                "Some files could not be imported.\n\n"
                + "\n".join(failed[:5])
                + f"\n\nLogged to:\n{LOG_DIR / 'app.log'}",
            )

    def select_exchange_rate_images(self) -> None:
        if self.form_version.get() != "Korea":
            self.show_info(self.tr("exchange_only_korea"))
            return
        ensure_runtime_folders()
        if filedialog is None:
            return
        selected = filedialog.askopenfilenames(
            title=self.tr("select_exchange_rate"),
            initialdir=self.remembered_file_dialog_dir(),
            filetypes=[
                ("Exchange rate image files", "*.png *.jpg *.jpeg *.webp *.bmp *.gif *.pdf"),
                ("Image files", "*.png *.jpg *.jpeg *.webp *.bmp *.gif"),
                ("PDF files", "*.pdf"),
                ("All files", "*.*"),
            ],
        )
        paths = [Path(path) for path in selected]
        if not paths:
            return
        self.remember_file_dialog_dir(paths)
        self._details_ready_for_export = False
        self.exchange_rate_items.clear()
        added = 0
        failed: List[str] = []
        for path in paths:
            suffix = path.suffix.lower()
            if suffix not in SUPPORTED_UPLOAD_SUFFIXES:
                continue
            if suffix == ".pdf":
                try:
                    image_path = render_pdf_to_merged_image(path)
                except Exception as exc:
                    log_exception(f"Could not import 汇率 PDF: {path}")
                    failed.append(f"{path.name}: {exc}")
                    continue
                filename = f"{path.name} (all pages)"
                source_page = "merged"
            else:
                image_path = path
                filename = path.name
                source_page = ""
            self.exchange_rate_items.append(
                BankStatementItem(
                    item_id=uuid.uuid4().hex,
                    path=str(image_path),
                    filename=filename,
                    source_path=str(path),
                    source_page=source_page,
                    status="Exchange rate",
                )
            )
            added += 1
        self.save_session()
        self.status_text.set(self.tr("selected_exchange_files", count=added))
        self.update_toolbar_recommendation()
        if added:
            self.read_exchange_rate_from_images()
        if failed and messagebox:
            messagebox.showwarning(
                self.tr("select_exchange_rate"),
                "Some files could not be imported.\n\n"
                + "\n".join(failed[:5])
                + f"\n\nLogged to:\n{LOG_DIR / 'app.log'}",
            )

    def read_exchange_rate_from_images(self) -> None:
        if not self.exchange_rate_items:
            return
        if not get_openai_api_key():
            self.status_text.set(self.tr("selected_exchange_no_key"))
            return
        image_paths = [Path(item.path) for item in self.exchange_rate_items]
        usd_to_rmb_rate = safe_float(self.exchange_rate.get()) or safe_float(DEFAULT_USD_TO_RMB_RATE) or 6.8175
        self.set_busy(True)
        self.set_progress(0, 1, self.tr("reading_exchange_short"))
        self.status_text.set(self.tr("reading_exchange"))

        def worker() -> None:
            try:
                rates = call_openai_exchange_rate_extraction(
                    image_paths,
                    usd_to_rmb_rate,
                    progress=lambda message: self.set_status(message),
                )
            except Exception as exc:
                log_exception("Could not read exchange rate images")
                message = str(exc)

                def fail() -> None:
                    self.set_busy(False)
                    self.set_progress(0, 1, "")
                    self.status_text.set(self.tr("selected_exchange_failed", message=message))

                self.root.after(0, fail)
                return

            def done() -> None:
                updates: List[str] = []
                if "usd_to_krw_rate" in rates:
                    self.usd_to_krw_rate.set(format_exchange_rate(rates["usd_to_krw_rate"]))
                    self.settings["usd_to_krw_rate"] = self.usd_to_krw_rate.get()
                    updates.append(f"USD -> KRW {self.usd_to_krw_rate.get()}")
                if "krw_to_rmb_rate" in rates:
                    self.krw_to_rmb_rate.set(format_exchange_rate(rates["krw_to_rmb_rate"]))
                    self.settings["krw_to_rmb_rate"] = self.krw_to_rmb_rate.get()
                    updates.append(f"KRW -> RMB {self.krw_to_rmb_rate.get()}")
                save_user_settings(self.settings)
                self.save_session()
                self.set_busy(False)
                self.set_progress(1, 1, "1/1")
                self.status_text.set(self.tr("updated_exchange_rate", updates=", ".join(updates)))

            self.root.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    def refresh_tree(self) -> None:
        current_selection = set(self.tree.selection())
        self.tree.delete(*self.tree.get_children())
        for index, item in enumerate(self.items):
            if self.form_version.get() == "USA" and item.status != "Manual filled":
                auto_classify_usa_document_type(item)
            attention_reason = receipt_attention_reason(item, self.form_version.get())
            if attention_reason is None and item.status in {
                "Missing date/amount",
                "Missing date",
                "Missing amount",
                "Missing details",
                "Missing vendor",
            }:
                item.status = "Payment proof" if is_usa_payment_proof_item(item) else "Manual filled"
            status_text = self.attention_reason_label(attention_reason) if attention_reason else item.status
            values = (status_text, item.date, item.amount or item.krw_amount or item.rmb_amount)
            if attention_reason in {"AI failed", "Wrong year"}:
                tags = ("danger",)
            elif attention_reason:
                tags = ("review",)
            else:
                tags = ()
            self.tree.insert("", "end", iid=str(index), text=item.filename, values=values, tags=tags)
        valid_selection = [iid for iid in current_selection if iid.isdigit() and int(iid) < len(self.items)]
        if valid_selection:
            self.tree.selection_set(*valid_selection)
        elif self.selected_index is not None and self.selected_index < len(self.items):
            self.tree.selection_set(str(self.selected_index))
        if self.selected_index is not None and self.selected_index < len(self.items):
            self.tree.focus(str(self.selected_index))

    def receipt_label_by_id(self, item_id: str) -> str:
        for item in self.items:
            if item.item_id == item_id:
                return item.receipt_label.strip() or item.filename
        return ""

    def refresh_bank_tree(self) -> None:
        if self.bank_tree is None:
            return
        current_selection = set(self.bank_tree.selection())
        self.bank_tree.delete(*self.bank_tree.get_children())
        for index, item in enumerate(self.bank_items):
            values = (
                item.status,
                item.date,
                item.amount,
                self.receipt_label_by_id(item.matched_receipt_id),
            )
            self.bank_tree.insert("", "end", iid=str(index), text=item.filename, values=values)
        valid_selection = [iid for iid in current_selection if iid.isdigit() and int(iid) < len(self.bank_items)]
        if valid_selection:
            self.bank_tree.selection_set(*valid_selection)
        elif self.selected_bank_index is not None and self.selected_bank_index < len(self.bank_items):
            self.bank_tree.selection_set(str(self.selected_bank_index))

    def _on_bank_tree_select(self, _event: Any = None) -> None:
        if self.bank_tree is None:
            return
        selection = self.bank_tree.selection()
        if not selection:
            self.selected_bank_index = None
            return
        try:
            self.selected_bank_index = int(selection[0])
        except Exception:
            self.selected_bank_index = None
        if self.selected_bank_index is not None:
            self.selected_attachment_kind = "proof"
            self.selected_attachment_index = 0
            self.update_preview()

    def selected_bank_item(self) -> Optional[BankStatementItem]:
        if self.selected_bank_index is None:
            return None
        if self.selected_bank_index < 0 or self.selected_bank_index >= len(self.bank_items):
            return None
        return self.bank_items[self.selected_bank_index]

    def _on_receipt_drag_start(self, event: Any) -> None:
        return

    def _is_widget_inside_bank_section(self, widget: Any) -> bool:
        return False

    def _on_receipt_drag_release(self, event: Any) -> None:
        return

    def move_selected_receipts_to_bank(self) -> None:
        self.show_info("Use Document type to mark an imported USA row as Payment proof.")

    def move_receipts_to_bank(self, indices: List[int]) -> None:
        self.show_info("Use Document type to mark imported USA rows as Payment proof.")

    def remove_selected_bank(self) -> None:
        item = self.selected_bank_item()
        if item is None:
            return
        del self.bank_items[self.selected_bank_index]
        self.selected_bank_index = None
        self.refresh_bank_tree()
        self.selected_attachment_kind = "receipt"
        self.selected_attachment_index = 0
        self.update_preview()
        self.save_session()
        self.status_text.set("Removed selected payment proof image.")
        self._details_ready_for_export = False
        self.update_toolbar_recommendation()

    def link_selected_bank_to_receipt(self) -> None:
        bank_item = self.selected_bank_item()
        receipt = self.selected_item()
        if bank_item is None or receipt is None:
            self.show_info("Select a payment proof image and the matching receipt row first.")
            return
        bank_item.matched_receipt_id = receipt.item_id
        bank_item.status = "Matched manually"
        self.refresh_bank_tree()
        self.save_session()
        self.status_text.set("Linked payment proof to selected receipt.")
        self._details_ready_for_export = False
        self.update_toolbar_recommendation()

    def _on_tree_select(self, _event: Any = None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        self.selected_attachment_kind = "receipt"
        self.selected_attachment_index = 0
        focus = self.tree.focus()
        chosen = focus if focus in selection else selection[-1]
        try:
            index = int(chosen)
        except Exception:
            return
        self.select_index(index)
        if len(selection) == 1:
            self._selection_anchor_index = index

    def _on_tree_shift_click(self, event: Any) -> str:
        iid = self.tree.identify_row(event.y)
        if not iid:
            return ""
        try:
            index = int(iid)
        except Exception:
            return "break"
        if index < 0 or index >= len(self.items):
            return "break"
        anchor = self._selection_anchor_index
        if anchor is None or anchor < 0 or anchor >= len(self.items):
            anchor = self.selected_index if self.selected_index is not None else index
        if anchor is None or anchor < 0 or anchor >= len(self.items):
            anchor = index
        start, end = sorted((anchor, index))
        selection = [str(row) for row in range(start, end + 1)]
        self.save_current_fields()
        self.tree.selection_set(selection)
        self.tree.focus(str(index))
        self.selected_index = index
        self.selected_attachment_kind = "receipt"
        self.selected_attachment_index = 0
        self.load_selected_into_fields()
        self.update_preview()
        return "break"

    def select_index(self, index: int) -> None:
        if self.selected_index == index:
            return
        self.save_current_fields()
        self.selected_index = index
        self._selection_anchor_index = index
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

    def select_all_receipts(self, _event: Any = None) -> str:
        if not self.items:
            return "break"
        self.tree.selection_set([str(index) for index in range(len(self.items))])
        self.tree.focus(str(0))
        self.selected_index = 0
        self._selection_anchor_index = 0
        self.load_selected_into_fields()
        self.update_preview()
        self.status_text.set(f"Selected all {len(self.items)} receipt row(s).")
        return "break"

    def selected_items(self) -> List[ReceiptItem]:
        return [self.items[index] for index in self.selected_indices()]

    def shortcut_focus_is_text_editor(self) -> bool:
        try:
            focus = self.root.focus_get()
            widget_class = str(focus.winfo_class()) if focus is not None else ""
        except Exception:
            return False
        return widget_class in {
            "Entry",
            "TEntry",
            "Text",
            "Combobox",
            "TCombobox",
            "Spinbox",
            "TSpinbox",
        }

    def screenshot_clipboard_entries_from_tile(self) -> List[Dict[str, Any]]:
        tile = self.selected_preview_tile()
        receipt = self.selected_item()
        if tile is None or receipt is None:
            return []
        kind = str(tile.get("kind") or "receipt")
        if kind == "proof":
            proof_item = tile.get("proof_item")
            if not isinstance(proof_item, BankStatementItem):
                return []
            return [
                {
                    "kind": "proof",
                    "attachment": attachment_from_item(proof_item),
                    "source_proof_id": proof_item.item_id,
                    "source_proof": copy.deepcopy(proof_item),
                }
            ]
        images = ensure_receipt_images(receipt)
        index = int(tile.get("index") or 0)
        if index < 0 or index >= len(images):
            return []
        attachment = copy.deepcopy(images[index])
        return [
            {
                "kind": "proof" if self.form_version.get() == "USA" and is_usa_payment_proof_item(receipt) else "receipt",
                "attachment": attachment,
                "source_item_id": receipt.item_id,
                "source_image_id": str(attachment.get("id") or ""),
                "source_item": copy.deepcopy(receipt),
            }
        ]

    def screenshot_clipboard_entries_from_rows(self) -> List[Dict[str, Any]]:
        entries: List[Dict[str, Any]] = []
        for item in self.selected_items():
            images = ensure_receipt_images(item)
            kind = "proof" if self.form_version.get() == "USA" and is_usa_payment_proof_item(item) else "receipt"
            for attachment in images:
                entries.append(
                    {
                        "kind": kind,
                        "attachment": copy.deepcopy(attachment),
                        "source_item_id": item.item_id,
                        "source_image_id": str(attachment.get("id") or ""),
                        "source_item": copy.deepcopy(item),
                    }
                )
        return entries

    def copy_screenshot_shortcut(self, event: Any = None, mode: str = "copy") -> Optional[str]:
        if self.shortcut_focus_is_text_editor():
            return None
        self.save_current_fields()
        focus = None
        try:
            focus = self.root.focus_get()
        except Exception:
            pass
        entries = (
            self.screenshot_clipboard_entries_from_rows()
            if focus is self.tree
            else self.screenshot_clipboard_entries_from_tile()
        )
        if not entries and focus is not self.tree:
            entries = self.screenshot_clipboard_entries_from_rows()
        if not entries:
            self.status_text.set(self.tr("clipboard_no_source"))
            return "break"
        clipboard_mode = "cut" if mode == "cut" else "copy"
        self.image_clipboard = {"mode": clipboard_mode, "entries": entries}
        self.status_text.set(self.tr("clipboard_cut" if clipboard_mode == "cut" else "clipboard_copied", count=len(entries)))
        return "break"

    def paste_screenshot_shortcut(self, event: Any = None) -> Optional[str]:
        if self.shortcut_focus_is_text_editor():
            return None
        clipboard = self.image_clipboard if isinstance(self.image_clipboard, dict) else None
        entries = list(clipboard.get("entries") or []) if clipboard else []
        if not entries:
            self.status_text.set(self.tr("clipboard_empty"))
            return "break"
        self.save_current_fields()
        target = self.selected_item()
        if target is None:
            self.status_text.set(self.tr("clipboard_no_target"))
            return "break"
        if self.form_version.get() == "USA" and is_usa_payment_proof_item(target):
            self.status_text.set(self.tr("clipboard_needs_receipt_target"))
            return "break"

        mode = "cut" if clipboard.get("mode") == "cut" else "copy"
        target_id = target.item_id
        pasted_receipt_ids: List[str] = []
        pasted_proof_ids: List[str] = []
        cut_source_images: Dict[str, set] = {}
        cut_source_proofs: set = set()

        target_images = ensure_receipt_images(target)
        for entry in entries:
            attachment = clone_attachment_for_manual_paste(entry.get("attachment"))
            if attachment is None:
                continue
            paste_as_proof = self.form_version.get() == "USA" and str(entry.get("kind")) == "proof"
            if paste_as_proof:
                source_proof = entry.get("source_proof")
                source_item = entry.get("source_item")
                proof_item = BankStatementItem(
                    item_id=uuid.uuid4().hex,
                    path=str(attachment.get("path") or ""),
                    filename=str(attachment.get("filename") or Path(str(attachment.get("path") or "")).name),
                    source_path=str(attachment.get("source_path") or ""),
                    source_page=str(attachment.get("source_page") or ""),
                    date=str(getattr(source_proof, "date", "") or getattr(source_item, "date", "") or target.date),
                    amount=str(getattr(source_proof, "amount", "") or getattr(source_item, "amount", "") or target.amount),
                    place=str(getattr(source_proof, "place", "") or getattr(source_item, "place", "") or target.place),
                    matched_receipt_id=target_id,
                    status="Matched manually",
                    crop_box=copy.deepcopy(attachment.get("crop_box")),
                    rotation_degrees=normalize_rotation(attachment.get("rotation_degrees", 0)),
                )
                self.bank_items.append(proof_item)
                pasted_proof_ids.append(proof_item.item_id)
            else:
                target_images.append(attachment)
                pasted_receipt_ids.append(str(attachment.get("id") or ""))
                source_item = entry.get("source_item")
                fill_blank_receipt_fields(target, source_item if isinstance(source_item, ReceiptItem) else None)

            if mode == "cut":
                source_item_id = str(entry.get("source_item_id") or "")
                source_image_id = str(entry.get("source_image_id") or "")
                if source_item_id and source_image_id:
                    cut_source_images.setdefault(source_item_id, set()).add(source_image_id)
                source_proof_id = str(entry.get("source_proof_id") or "")
                if source_proof_id:
                    cut_source_proofs.add(source_proof_id)

        if not pasted_receipt_ids and not pasted_proof_ids:
            self.status_text.set(self.tr("clipboard_no_source"))
            return "break"

        if pasted_receipt_ids:
            target.receipt_images = target_images
            ensure_receipt_images(target)
            self.clear_attention_status_if_ready(target)

        if mode == "cut":
            remove_item_ids: set = set()
            for item in self.items:
                image_ids = cut_source_images.get(item.item_id)
                if not image_ids:
                    continue
                remaining = [image for image in ensure_receipt_images(item) if str(image.get("id") or "") not in image_ids]
                item.receipt_images = remaining
                if remaining:
                    ensure_receipt_images(item)
                else:
                    item.path = ""
                    remove_item_ids.add(item.item_id)
            if remove_item_ids:
                self.items = [item for item in self.items if item.item_id not in remove_item_ids]
            if cut_source_proofs:
                self.bank_items = [item for item in self.bank_items if item.item_id not in cut_source_proofs]
            self.image_clipboard = None

        id_to_index = {item.item_id: index for index, item in enumerate(self.items)}
        self.selected_index = id_to_index.get(target_id, 0 if self.items else None)
        if pasted_receipt_ids and self.selected_index is not None and 0 <= self.selected_index < len(self.items):
            images = ensure_receipt_images(self.items[self.selected_index])
            self.selected_attachment_kind = "receipt"
            self.selected_attachment_index = next(
                (index for index, image in enumerate(images) if str(image.get("id") or "") == pasted_receipt_ids[0]),
                max(0, len(images) - 1),
            )
        elif pasted_proof_ids and self.selected_index is not None:
            proofs = self.payment_proofs_for_receipt(self.items[self.selected_index])
            self.selected_attachment_kind = "proof"
            self.selected_attachment_index = next(
                (index for index, proof in enumerate(proofs) if proof.item_id == pasted_proof_ids[0]),
                max(0, len(proofs) - 1),
            )

        self.refresh_tree()
        self.refresh_bank_tree()
        if self.selected_index is not None and 0 <= self.selected_index < len(self.items):
            self.tree.selection_set(str(self.selected_index))
            self.tree.focus(str(self.selected_index))
        self.load_selected_into_fields()
        self.update_preview()
        if self.selected_index is not None and 0 <= self.selected_index < len(self.items):
            self.remember_receipt_item(self.items[self.selected_index], "manual")
        self.save_session()
        self._details_ready_for_export = False
        self.update_toolbar_recommendation()
        self.status_text.set(self.tr("clipboard_pasted", count=len(entries)))
        return "break"

    def sort_items_for_current_report(self) -> None:
        if not self.items:
            return
        selected_ids = {
            self.items[index].item_id
            for index in self.selected_indices()
            if 0 <= index < len(self.items)
        }
        focused_id = self.selected_item().item_id if self.selected_item() is not None else ""
        if self.form_version.get() == "USA":
            self.items, _removed_to_target, _merged_count = merge_same_usa_receipt_items(self.items)
            match_usa_payment_proof_items(self.items)
            report_rows = build_usa_report_rows(self.items, self.bank_items)
            ordered_ids: List[str] = []
            for row in report_rows:
                item = row.get("item")
                if isinstance(item, ReceiptItem):
                    ordered_ids.append(item.item_id)
                    ordered_ids.extend(
                        proof.item_id
                        for proof in self.items
                        if is_usa_payment_proof_item(proof) and proof.matched_receipt_id == item.item_id
                    )
                else:
                    proof_images = list(row.get("proof_images") or [])
                    proof_ids = {str(image.get("id") or "") for image in proof_images if isinstance(image, dict)}
                    ordered_ids.extend(
                        proof.item_id
                        for proof in self.items
                        if is_usa_payment_proof_item(proof)
                        and proof.item_id in proof_ids
                    )
            seen_ids: set[str] = set()
            id_to_item = {item.item_id: item for item in self.items}
            ordered: List[ReceiptItem] = []
            for item_id in ordered_ids:
                if item_id in id_to_item and item_id not in seen_ids:
                    ordered.append(id_to_item[item_id])
                    seen_ids.add(item_id)
            ordered.extend(item for item in self.items if item.item_id not in seen_ids)
            self.items = ordered
        else:
            self.items = sort_receipts_for_report(self.items, self.form_version.get())
        self.items = sort_receipt_list_for_review(self.items)
        id_to_index = {item.item_id: index for index, item in enumerate(self.items)}
        if focused_id in id_to_index:
            self.selected_index = id_to_index[focused_id]
        elif selected_ids:
            self.selected_index = id_to_index.get(next(iter(selected_ids)), 0)
        elif self.items:
            self.selected_index = 0
        else:
            self.selected_index = None
        selection = [str(id_to_index[item_id]) for item_id in selected_ids if item_id in id_to_index]
        self.refresh_tree()
        if selection:
            self.tree.selection_set(selection)
        elif self.selected_index is not None:
            self.tree.selection_set(str(self.selected_index))

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
        before = receipt_item_signature(item)
        for key, var in self.field_vars.items():
            value = var.get()
            if key == "category":
                value = self.category_key_from_display(value)
            elif key == "document_type":
                value = normalize_document_type(self.document_type_key_from_display(value), "unknown")
                if value != "payment_proof":
                    item.matched_receipt_id = ""
            elif key == "currency":
                value = normalize_currency(value, "USD" if self.form_version.get() == "USA" else "KRW")
            setattr(item, key, value)
        self.update_item_amount_fields(item, self._last_amount_source)
        self.clear_attention_status_if_ready(item)
        if receipt_item_signature(item) != before:
            self.remember_receipt_item(item, "manual")

    def load_selected_into_fields(self) -> None:
        item = self.selected_item()
        if item is None:
            self._loading_fields = True
            try:
                for var in self.field_vars.values():
                    var.set("")
            finally:
                self._loading_fields = False
            self.update_date_field_attention()
            return
        self._loading_fields = True
        try:
            for key, var in self.field_vars.items():
                value = getattr(item, key)
                if key == "category":
                    value = self.category_display_label(category_value_to_key(value, self.form_version.get()))
                elif key == "document_type":
                    value = self.document_type_display_label(normalize_document_type(value, "unknown"))
                var.set(value)
        finally:
            self._loading_fields = False
        self._sync_amount_fields("rate")
        self.update_date_field_attention()

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
        self.preview_canvas.configure(scrollregion=(0, 0, width, height))
        self.photo = None
        self.preview_photos = []
        self.preview_tiles = []
        self.preview_action_regions = {}
        self.proof_drop_region = None
        self._pressed_preview_tile = None
        self._dragging_preview_tile = None
        self._drag_start = (0, 0)
        self._drag_outline_id = None
        self._preview_image_bounds = (0, 0, 0, 0)
        self._preview_original_size = (0, 0)
        self._crop_handle_centers = {}
        if self.revert_crop_btn is not None:
            self.revert_crop_btn.configure(state="disabled")
        if self.rotate_left_btn is not None:
            self.rotate_left_btn.configure(state="disabled")
        if self.rotate_right_btn is not None:
            self.rotate_right_btn.configure(state="disabled")
        if self.delete_screenshot_btn is not None:
            self.delete_screenshot_btn.configure(state="disabled")
        if self.zoom_out_btn is not None:
            self.zoom_out_btn.configure(state="disabled")
        if self.zoom_in_btn is not None:
            self.zoom_in_btn.configure(state="disabled")
        self.update_zoom_label()
        if self.swap_proof_btn is not None:
            self.swap_proof_btn.configure(state="disabled")
        if self.unlink_proof_btn is not None:
            self.unlink_proof_btn.configure(state="disabled")

    def payment_proofs_for_receipt(self, receipt: ReceiptItem) -> List[BankStatementItem]:
        return [item for item in self.bank_items if item.matched_receipt_id == receipt.item_id]

    def selected_preview_tile(self) -> Optional[Dict[str, Any]]:
        for tile in self.preview_tiles:
            if (
                tile.get("kind") == self.selected_attachment_kind
                and tile.get("index") == self.selected_attachment_index
            ):
                return tile
        return None

    def preview_tile_at(self, x: float, y: float) -> Optional[Dict[str, Any]]:
        for tile in reversed(self.preview_tiles):
            left, top, right, bottom = tile.get("bbox", (0, 0, 0, 0))
            if left <= x <= right and top <= y <= bottom:
                return tile
        return None

    def selected_image_state(self) -> Optional[Tuple[Path, Any, int, Optional[Dict[str, Any]], Optional[BankStatementItem]]]:
        tile = self.selected_preview_tile()
        if tile is None:
            return None
        if tile.get("kind") == "proof":
            proof_item = tile.get("proof_item")
            if not isinstance(proof_item, BankStatementItem):
                return None
            return (
                Path(proof_item.path),
                proof_item.crop_box,
                normalize_rotation(proof_item.rotation_degrees),
                None,
                proof_item,
            )
        attachment = None
        receipt = self.selected_item()
        if receipt is not None:
            images = ensure_receipt_images(receipt)
            index = int(tile.get("index") or 0)
            if 0 <= index < len(images):
                attachment = images[index]
        if attachment is None:
            attachment = tile.get("attachment")
        if not isinstance(attachment, dict):
            return None
        return (
            Path(str(attachment.get("path") or "")),
            attachment.get("crop_box"),
            normalize_rotation(attachment.get("rotation_degrees", 0)),
            attachment,
            None,
        )

    def set_selected_image_state(self, crop_box: Any = None, rotation_degrees: Optional[int] = None) -> None:
        state = self.selected_image_state()
        if state is None:
            return
        _path, _crop, _rotation, attachment, proof_item = state
        if proof_item is not None:
            proof_item.crop_box = copy.deepcopy(crop_box)
            if rotation_degrees is not None:
                proof_item.rotation_degrees = normalize_rotation(rotation_degrees)
            return
        if attachment is None:
            return
        receipt = self.selected_item()
        if receipt is not None:
            images = ensure_receipt_images(receipt)
            index = self.selected_attachment_index
            if 0 <= index < len(images):
                attachment = images[index]
        attachment["crop_box"] = copy.deepcopy(crop_box)
        if rotation_degrees is not None:
            attachment["rotation_degrees"] = normalize_rotation(rotation_degrees)
        if receipt is not None:
            ensure_receipt_images(receipt)

    def update_preview(self) -> None:
        if self.preview_canvas is None:
            return
        receipt = self.selected_item()
        if receipt is None:
            self.clear_preview(self.tr("select_receipt_prompt"))
            return
        if Image is None or ImageTk is None:
            self.clear_preview("Image preview is unavailable. Install Pillow from requirements.txt.")
            return
        receipt_images = ensure_receipt_images(receipt)
        show_payment_proof = self.form_version.get() == "USA" and bool(self.bank_items)
        proof_items = self.payment_proofs_for_receipt(receipt) if show_payment_proof else []
        if self.selected_attachment_kind == "proof" and (not proof_items or not show_payment_proof):
            self.selected_attachment_kind = "receipt"
            self.selected_attachment_index = 0
        if self.selected_attachment_kind == "receipt" and self.selected_attachment_index >= len(receipt_images):
            self.selected_attachment_index = 0
        if self.selected_attachment_kind == "proof" and self.selected_attachment_index >= len(proof_items):
            self.selected_attachment_index = 0
        canvas_w = max(240, self.preview_canvas.winfo_width())
        canvas_h = max(240, self.preview_canvas.winfo_height())
        zoom = self.preview_zoom_value()
        layout_w = max(canvas_w, int(canvas_w * zoom))
        layout_h = max(canvas_h, int(canvas_h * zoom))
        self.preview_canvas.delete("all")
        self.preview_photos = []
        self.preview_tiles = []
        self.preview_action_regions = {}
        self.proof_drop_region = None
        self.photo = None
        self._preview_image_bounds = (0, 0, 0, 0)
        self._preview_original_size = (0, 0)
        self._crop_handle_centers = {}
        try:
            if show_payment_proof:
                gap = 46
                left_w = max(120, int((layout_w - gap) * 0.56))
                right_w = max(100, layout_w - left_w - gap)
                proof_x = left_w + gap
                self.proof_drop_region = (proof_x, 0, layout_w, layout_h)
                self.draw_preview_section(
                    self.tr("receipt_screenshots"),
                    0,
                    0,
                    left_w,
                    layout_h,
                    receipt_images,
                    "receipt",
                )
                self.draw_preview_section(
                    self.tr("payment_proof_section"),
                    proof_x,
                    0,
                    right_w,
                    layout_h,
                    proof_items,
                    "proof",
                )
                self.draw_preview_actions(left_w, gap, layout_h)
            else:
                self.draw_preview_section(
                    self.tr("receipt_screenshots"),
                    0,
                    0,
                    layout_w,
                    layout_h,
                    receipt_images,
                    "receipt",
                )
            if self.preview_tiles and self.selected_preview_tile() is None:
                first_tile = self.preview_tiles[0]
                self.selected_attachment_kind = str(first_tile.get("kind") or "receipt")
                self.selected_attachment_index = int(first_tile.get("index") or 0)
            self.draw_crop_overlay()
            bbox = self.preview_canvas.bbox("all")
            if bbox:
                self.preview_canvas.configure(
                    scrollregion=(
                        0,
                        0,
                        max(layout_w, int(bbox[2]) + 12),
                        max(layout_h, int(bbox[3]) + 12),
                    )
                )
            else:
                self.preview_canvas.configure(scrollregion=(0, 0, layout_w, layout_h))
            self.update_preview_buttons()
        except Exception as exc:
            log_exception("Receipt preview failed")
            self.clear_preview(f"Could not preview image:\n{exc}")

    def draw_preview_section(
        self,
        title: str,
        x: int,
        y: int,
        width: int,
        height: int,
        entries: Sequence[Any],
        kind: str,
    ) -> None:
        if self.preview_canvas is None:
            return
        self.preview_canvas.create_text(
            x + 8,
            y + 8,
            text=title,
            anchor="nw",
            fill="#222222",
            font=("Segoe UI", 10, "bold"),
        )
        if kind == "proof":
            self.preview_canvas.create_line(x - 7, y, x - 7, y + height, fill="#D8D8D8")
        if not entries:
            message = self.tr("proof_prompt") if kind == "proof" else self.tr("no_receipt_screenshot")
            self.preview_canvas.create_text(
                x + width / 2,
                y + height / 2,
                text=message,
                fill="#777777",
                width=max(120, width - 20),
                justify="center",
            )
            return
        pdf_source_paths = {
            str(entry.get("source_path") or "").strip().lower()
            for entry in entries
            if kind != "proof" and isinstance(entry, dict) and str(entry.get("source_path") or "").strip().lower().endswith(".pdf")
        }
        stack_pdf_pages = kind != "proof" and len(entries) > 1 and len(pdf_source_paths) == 1
        columns = 1 if width < 360 or len(entries) == 1 or stack_pdf_pages else 2
        rows = max(1, (len(entries) + columns - 1) // columns)
        gap = 10
        top = y + 34
        cell_w = max(72, int((width - gap * (columns - 1) - 8) / columns))
        cell_h = max(86, int((height - 42 - gap * (rows - 1)) / rows))
        for index, entry in enumerate(entries):
            col = index % columns
            row = index // columns
            cell_x = x + 4 + col * (cell_w + gap)
            cell_y = top + row * (cell_h + gap)
            bbox = (cell_x, cell_y, cell_x + cell_w, cell_y + cell_h)
            attachment: Optional[Dict[str, Any]] = None
            proof_item: Optional[BankStatementItem] = None
            if kind == "proof":
                proof_item = entry if isinstance(entry, BankStatementItem) else None
                if proof_item is None:
                    continue
                attachment = attachment_from_item(proof_item)
            else:
                attachment = entry if isinstance(entry, dict) else None
            if attachment is None:
                continue
            path = Path(str(attachment.get("path") or ""))
            is_selected = kind == self.selected_attachment_kind and index == self.selected_attachment_index
            outline = "#0078D7" if is_selected else "#D0D0D0"
            self.preview_canvas.create_rectangle(*bbox, outline=outline, width=2 if is_selected else 1)
            image_bounds = (cell_x + 6, cell_y + 26, 1, 1)
            original_size = (0, 0)
            filename = str(attachment.get("filename") or path.name)
            self.preview_canvas.create_text(
                cell_x + 8,
                cell_y + 6,
                text=filename,
                anchor="nw",
                fill="#444444",
                font=("Segoe UI", 8),
                width=max(60, cell_w - 16),
            )
            if path.exists():
                try:
                    image = oriented_image_from_path(path, attachment.get("rotation_degrees", 0))
                    original_size = (image.width, image.height)
                    max_w = max(1, cell_w - 14)
                    max_h = max(1, cell_h - 34)
                    image.thumbnail((max_w, max_h))
                    photo = ImageTk.PhotoImage(image)
                    self.preview_photos.append(photo)
                    image_x = int(cell_x + (cell_w - image.width) / 2)
                    image_y = int(cell_y + 28 + max(0, (max_h - image.height) / 2))
                    self.preview_canvas.create_image(image_x, image_y, anchor="nw", image=photo)
                    image_bounds = (image_x, image_y, image.width, image.height)
                except Exception:
                    log_exception(f"Could not draw preview tile: {path}")
                    self.preview_canvas.create_text(
                        cell_x + cell_w / 2,
                        cell_y + cell_h / 2,
                        text="Preview failed",
                        fill="#777777",
                    )
            else:
                self.preview_canvas.create_text(
                    cell_x + cell_w / 2,
                    cell_y + cell_h / 2,
                    text="Missing file",
                    fill="#777777",
                )
            self.preview_tiles.append(
                {
                    "kind": kind,
                    "index": index,
                    "bbox": bbox,
                    "image_bounds": image_bounds,
                    "original_size": original_size,
                    "attachment": attachment,
                    "proof_item": proof_item,
                }
            )

    def draw_preview_actions(self, left_width: int, gap: int, canvas_height: int) -> None:
        if self.preview_canvas is None:
            return
        receipt = self.selected_item()
        proofs = self.payment_proofs_for_receipt(receipt) if receipt is not None else []
        has_proofs = bool(proofs)
        can_swap = self.form_version.get() == "USA" and has_proofs and len(self.bank_items) > len(proofs)
        center_x = left_width + gap // 2
        actions = [
            ("swap", "\u21C4", "Swap payment proof", can_swap),
        ]
        start_y = max(56, canvas_height // 2)
        self.preview_action_regions = {}
        for index, (name, icon, label, enabled) in enumerate(actions):
            y = start_y + index * 48
            bbox = (center_x - 17, y - 17, center_x + 17, y + 17)
            fill = "#2563EB" if enabled else "#E2E8F0"
            outline = "#1D4ED8" if enabled else "#CBD5E1"
            text_fill = "#FFFFFF" if enabled else "#94A3B8"
            self.preview_canvas.create_oval(*bbox, fill=fill, outline=outline, width=1)
            self.preview_canvas.create_text(
                center_x,
                y - 1,
                text=icon,
                fill=text_fill,
                font=("Segoe UI Symbol", 14, "bold"),
            )
            self.preview_action_regions[name] = {
                "bbox": bbox,
                "enabled": enabled,
                "label": label,
            }
        if self.proof_drop_region is not None and not has_proofs:
            left, top, right, bottom = self.proof_drop_region
            self.preview_canvas.create_rectangle(
                left + 4,
                top + 34,
                right - 4,
                bottom - 4,
                outline="#CBD5E1",
                dash=(4, 6),
                width=1,
            )

    def preview_action_at(self, x: float, y: float) -> Optional[str]:
        for action, meta in self.preview_action_regions.items():
            left, top, right, bottom = meta.get("bbox", (0, 0, 0, 0))
            if left <= x <= right and top <= y <= bottom:
                return action
        return None

    def point_in_proof_drop_region(self, x: float, y: float) -> bool:
        if self.proof_drop_region is None:
            return False
        left, top, right, bottom = self.proof_drop_region
        return left <= x <= right and top <= y <= bottom

    def run_preview_action(self, action: str) -> None:
        meta = self.preview_action_regions.get(action)
        if not meta or not meta.get("enabled"):
            return
        if action == "swap":
            self.swap_payment_proof()
        elif action == "unlink":
            self.unlink_payment_proof()

    def update_preview_buttons(self) -> None:
        has_tile = self.selected_preview_tile() is not None
        state = "normal" if has_tile else "disabled"
        if not has_tile and self.revert_crop_btn is not None:
            self.revert_crop_btn.configure(state="disabled")
        if self.zoom_out_btn is not None:
            self.zoom_out_btn.configure(state=state)
        if self.zoom_in_btn is not None:
            self.zoom_in_btn.configure(state=state)
        self.update_zoom_label()
        if self.rotate_left_btn is not None:
            self.rotate_left_btn.configure(state=state)
        if self.rotate_right_btn is not None:
            self.rotate_right_btn.configure(state=state)
        if self.delete_screenshot_btn is not None:
            self.delete_screenshot_btn.configure(state=state)
        receipt = self.selected_item()
        has_proofs = bool(receipt and self.payment_proofs_for_receipt(receipt))
        can_use_proofs = self.form_version.get() == "USA" and receipt is not None and bool(self.bank_items)
        if self.swap_proof_btn is not None:
            self.swap_proof_btn.configure(state="normal" if can_use_proofs else "disabled")
        if self.unlink_proof_btn is not None:
            self.unlink_proof_btn.configure(state="normal" if has_proofs else "disabled")

    def current_preview_crop_points(self) -> List[Tuple[float, float]]:
        state = self.selected_image_state()
        width, height = self._preview_original_size
        crop_box = state[1] if state is not None else None
        crop = normalized_crop_points(crop_box, width, height)
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
        tile = self.selected_preview_tile()
        state = self.selected_image_state()
        if tile is None or state is None:
            self.update_preview_buttons()
            return
        self._preview_image_bounds = tuple(tile.get("image_bounds", (0, 0, 0, 0)))  # type: ignore[assignment]
        self._preview_original_size = tuple(tile.get("original_size", (0, 0)))  # type: ignore[assignment]
        original_w, original_h = self._preview_original_size
        image_x, image_y, display_w, display_h = self._preview_image_bounds
        if original_w <= 0 or original_h <= 0 or display_w <= 0 or display_h <= 0:
            self.update_preview_buttons()
            return
        self.preview_canvas.delete("crop")
        crop = self.current_preview_crop_points()
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
            has_changes = normalized_crop_points(state[1], original_w, original_h) is not None or state[2] != 0
            self.revert_crop_btn.configure(state="normal" if has_changes else "disabled")
        self.preview_canvas.tag_raise("crop")

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
        self._pressed_preview_tile = None
        self._dragging_preview_tile = None
        x, y = self.preview_event_coords(event)
        self._drag_start = (int(x), int(y))
        if self._drag_outline_id is not None and self.preview_canvas is not None:
            self.preview_canvas.delete(self._drag_outline_id)
            self._drag_outline_id = None
        action = self.preview_action_at(x, y)
        if action is not None:
            self.run_preview_action(action)
            return
        tile = self.preview_tile_at(x, y)
        if tile is not None:
            kind = str(tile.get("kind") or "receipt")
            index = int(tile.get("index") or 0)
            if kind != self.selected_attachment_kind or index != self.selected_attachment_index:
                self.selected_attachment_kind = kind
                self.selected_attachment_index = index
                self.update_preview()
                tile = self.selected_preview_tile()
            self._pressed_preview_tile = tile
        self._dragging_crop_handle = self.nearest_crop_handle(x, y)
        if self._dragging_crop_handle is not None:
            self._pressed_preview_tile = None

    def _on_crop_drag(self, event: Any) -> None:
        handle = self._dragging_crop_handle
        original_w, original_h = self._preview_original_size
        x, y = self.preview_event_coords(event)
        if handle is None and self._pressed_preview_tile is not None:
            if self.form_version.get() != "USA" or self._pressed_preview_tile.get("kind") != "receipt":
                return
            dx = x - self._drag_start[0]
            dy = y - self._drag_start[1]
            if self._dragging_preview_tile is None and (dx * dx + dy * dy) < 64:
                return
            self._dragging_preview_tile = self._pressed_preview_tile
            if self.preview_canvas is not None:
                if self._drag_outline_id is not None:
                    self.preview_canvas.delete(self._drag_outline_id)
                left, top, right, bottom = self._dragging_preview_tile.get("bbox", (0, 0, 0, 0))
                width = right - left
                height = bottom - top
                x1 = int(x - width / 2)
                y1 = int(y - height / 2)
                self._drag_outline_id = self.preview_canvas.create_rectangle(
                    x1,
                    y1,
                    x1 + width,
                    y1 + height,
                    outline="#2563EB" if self.point_in_proof_drop_region(x, y) else "#64748B",
                    dash=(5, 4),
                    width=2,
                )
                self.preview_canvas.tag_raise(self._drag_outline_id)
            self.status_text.set("Release in Payment proof to move the screenshot.")
            return
        if self.selected_image_state() is None or handle is None or original_w <= 0 or original_h <= 0:
            return
        points = self.current_preview_crop_points()
        x, y = self.canvas_to_original(x, y)
        handle_indexes = {"nw": 0, "ne": 1, "se": 2, "sw": 3}
        points[handle_indexes[handle]] = (x, y)
        state = self.selected_image_state()
        rotation = state[2] if state is not None else 0
        self.set_selected_image_state(flatten_crop_points(points), rotation)
        self.draw_crop_overlay()

    def _on_crop_release(self, event: Any) -> None:
        x, y = self.preview_event_coords(event)
        if self._dragging_preview_tile is not None:
            tile = self._dragging_preview_tile
            dropped = self.point_in_proof_drop_region(x, y)
            if self._drag_outline_id is not None and self.preview_canvas is not None:
                self.preview_canvas.delete(self._drag_outline_id)
            self._drag_outline_id = None
            self._dragging_preview_tile = None
            self._pressed_preview_tile = None
            if dropped:
                self.move_receipt_screenshot_to_payment_proof(tile)
            else:
                self.status_text.set("Screenshot move cancelled.")
            return
        self._pressed_preview_tile = None
        if self._dragging_crop_handle is None:
            return
        self._dragging_crop_handle = None
        self.save_session()
        self.status_text.set("Crop updated for selected screenshot.")

    def revert_crop(self) -> None:
        state = self.selected_image_state()
        if state is None:
            return
        self.set_selected_image_state(None, 0)
        self.update_preview()
        self.save_session()
        self.status_text.set("Selected screenshot reverted to original.")

    def rotate_selected(self, delta_degrees: int) -> None:
        state = self.selected_image_state()
        if state is None:
            return
        path, crop_box, rotation_degrees, _attachment, _proof_item = state
        try:
            image = oriented_image_from_path(path, rotation_degrees)
            transformed_crop = rotated_crop_points(
                crop_box,
                image.width,
                image.height,
                delta_degrees,
            )
        except Exception:
            log_exception("Could not transform crop during rotation")
            transformed_crop = None
        new_rotation = normalize_rotation(normalize_rotation(rotation_degrees) + delta_degrees)
        self.set_selected_image_state(transformed_crop, new_rotation)
        self.update_preview()
        self.save_session()
        direction = "right" if delta_degrees > 0 else "left"
        self.status_text.set(f"Rotated selected screenshot {direction}.")

    def move_receipt_screenshot_to_payment_proof(self, tile: Dict[str, Any]) -> None:
        if self.form_version.get() != "USA":
            return
        receipt = self.selected_item()
        if receipt is None:
            return
        index = int(tile.get("index") or 0)
        images = ensure_receipt_images(receipt)
        if index < 0 or index >= len(images):
            return
        attachment = copy.deepcopy(images[index])
        filename = str(attachment.get("filename") or Path(str(attachment.get("path") or "")).name)
        removing_row = len(images) == 1
        if removing_row and messagebox and not messagebox.askyesno(
            "Move to payment proof",
            f"Move this screenshot to Payment proof and remove the receipt row?\n\n{filename}",
        ):
            return
        proof_item = BankStatementItem(
            item_id=uuid.uuid4().hex,
            path=str(attachment.get("path") or ""),
            filename=filename,
            source_path=str(attachment.get("source_path") or ""),
            source_page=str(attachment.get("source_page") or ""),
            date=receipt.date,
            amount=receipt.amount,
            place=receipt.place,
            matched_receipt_id="" if removing_row else receipt.item_id,
            status="Payment proof" if removing_row else "Matched manually",
            crop_box=copy.deepcopy(attachment.get("crop_box")),
            rotation_degrees=normalize_rotation(attachment.get("rotation_degrees", 0)),
        )
        self.bank_items.append(proof_item)
        del images[index]
        if not removing_row:
            receipt.receipt_images = images
            ensure_receipt_images(receipt)
            self.selected_attachment_kind = "proof"
            self.selected_attachment_index = len(self.payment_proofs_for_receipt(receipt)) - 1
            self.refresh_tree()
            self.refresh_bank_tree()
            self.update_preview()
            self.save_session()
            self.status_text.set(f"Moved {filename} to Payment proof.")
            self._details_ready_for_export = False
            self.update_toolbar_recommendation()
            return

        remove_index = self.selected_index
        if remove_index is not None and 0 <= remove_index < len(self.items):
            del self.items[remove_index]
        self.selected_index = None
        self.selected_attachment_kind = "receipt"
        self.selected_attachment_index = 0
        self.refresh_tree()
        self.refresh_bank_tree()
        if self.items:
            self.select_index(min(remove_index or 0, len(self.items) - 1))
        else:
            self.load_selected_into_fields()
            self.update_preview()
        self.save_session()
        self.status_text.set("Moved screenshot to Payment proof.")
        self._details_ready_for_export = False
        self.update_toolbar_recommendation()

    def delete_selected_screenshot(self) -> None:
        tile = self.selected_preview_tile()
        receipt = self.selected_item()
        if tile is None or receipt is None:
            return
        kind = str(tile.get("kind") or "receipt")
        index = int(tile.get("index") or 0)
        if kind == "proof":
            proof_item = tile.get("proof_item")
            if not isinstance(proof_item, BankStatementItem):
                return
            if messagebox and not messagebox.askyesno(
                "Delete screenshot",
                f"Remove this payment proof image?\n\n{proof_item.filename}",
            ):
                return
            self.bank_items = [item for item in self.bank_items if item.item_id != proof_item.item_id]
            self.selected_attachment_kind = "receipt"
            self.selected_attachment_index = 0
            self.refresh_bank_tree()
            self.update_preview()
            self.save_session()
            self.status_text.set("Deleted selected payment proof screenshot.")
            self._details_ready_for_export = False
            self.update_toolbar_recommendation()
            return

        images = ensure_receipt_images(receipt)
        if index < 0 or index >= len(images):
            return
        filename = str(images[index].get("filename") or Path(str(images[index].get("path") or "")).name)
        if messagebox and not messagebox.askyesno(
            "Delete screenshot",
            f"Remove this receipt screenshot?\n\n{filename}",
        ):
            return
        del images[index]
        if images:
            receipt.receipt_images = images
            ensure_receipt_images(receipt)
            self.selected_attachment_index = max(0, min(index, len(images) - 1))
            self.refresh_tree()
            self.update_preview()
            self.save_session()
            self.status_text.set("Deleted selected receipt screenshot.")
            return

        removed_id = receipt.item_id
        remove_index = self.selected_index
        if remove_index is not None and 0 <= remove_index < len(self.items):
            del self.items[remove_index]
        for proof_item in self.bank_items:
            if proof_item.matched_receipt_id == removed_id:
                proof_item.matched_receipt_id = ""
                proof_item.status = "Payment proof"
        self.selected_index = None
        self.refresh_tree()
        if self.items:
            self.select_index(min(remove_index or 0, len(self.items) - 1))
        else:
            self.load_selected_into_fields()
            self.update_preview()
        self.save_session()
        self.status_text.set("Deleted receipt row because its last screenshot was removed.")

    def swap_payment_proof(self) -> None:
        receipt = self.selected_item()
        if receipt is None:
            return
        if self.form_version.get() != "USA":
            self.show_info("Payment proof is only used for the USA form.")
            return
        if not self.bank_items:
            self.show_info("Select payment proof files first.")
            return
        current_indexes = [
            index for index, proof_item in enumerate(self.bank_items)
            if proof_item.matched_receipt_id == receipt.item_id
        ]
        start = current_indexes[0] if current_indexes else -1
        chosen_index: Optional[int] = None
        for offset in range(1, len(self.bank_items) + 1):
            candidate_index = (start + offset) % len(self.bank_items)
            if candidate_index not in current_indexes:
                chosen_index = candidate_index
                break
        if chosen_index is None:
            self.show_info("No other payment proof is loaded.")
            return
        for index in current_indexes:
            self.bank_items[index].matched_receipt_id = ""
            self.bank_items[index].status = "Payment proof"
        chosen = self.bank_items[chosen_index]
        chosen.matched_receipt_id = receipt.item_id
        chosen.status = "Matched manually"
        self.selected_attachment_kind = "proof"
        self.selected_attachment_index = 0
        self.refresh_bank_tree()
        self.update_preview()
        self.save_session()
        self.status_text.set(f"Swapped payment proof to {chosen.filename}.")
        self._details_ready_for_export = False
        self.update_toolbar_recommendation()

    def unlink_payment_proof(self) -> None:
        receipt = self.selected_item()
        if receipt is None:
            return
        changed = 0
        for proof_item in self.bank_items:
            if proof_item.matched_receipt_id == receipt.item_id:
                proof_item.matched_receipt_id = ""
                proof_item.status = "Payment proof"
                changed += 1
        if changed:
            self.selected_attachment_kind = "receipt"
            self.selected_attachment_index = 0
            self.refresh_bank_tree()
            self.update_preview()
            self.save_session()
            self.status_text.set("Unlinked payment proof from selected receipt.")
            self._details_ready_for_export = False
            self.update_toolbar_recommendation()

    def remove_selected(self, _event: Any = None) -> str:
        indices = self.selected_indices()
        if not indices:
            return "break"
        if messagebox and not messagebox.askyesno(
            "Remove receipts",
            f"Remove {len(indices)} selected receipt row(s)?",
        ):
            return "break"
        removed_ids = {self.items[index].item_id for index in indices if 0 <= index < len(self.items)}
        for index in sorted(indices, reverse=True):
            del self.items[index]
        for bank_item in self.bank_items:
            if bank_item.matched_receipt_id in removed_ids:
                bank_item.matched_receipt_id = ""
                bank_item.status = "Payment proof"
        self.selected_index = None
        self.refresh_tree()
        self.refresh_bank_tree()
        if self.items:
            next_index = min(indices[0], len(self.items) - 1)
            self._selection_anchor_index = next_index
            self.select_index(next_index)
        else:
            self._selection_anchor_index = None
            self.load_selected_into_fields()
            self.update_preview()
        self.status_text.set(f"Removed {len(indices)} selected receipt(s).")
        self._details_ready_for_export = False
        self.update_toolbar_recommendation()
        self.save_session()
        return "break"

    def clear_all(self) -> None:
        if not self.items and not self.bank_items and not self.exchange_rate_items:
            return
        if messagebox and not messagebox.askyesno("Clear receipts", "Remove all uploaded receipt rows and support images?"):
            return
        self.items.clear()
        self.bank_items.clear()
        self.exchange_rate_items.clear()
        self.selected_index = None
        self.selected_bank_index = None
        self._selection_anchor_index = None
        self.refresh_tree()
        self.refresh_bank_tree()
        self.load_selected_into_fields()
        self.update_preview()
        self.status_text.set("Cleared receipt list.")
        self._details_ready_for_export = False
        self.update_toolbar_recommendation()
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
            "document_type": "document_type",
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
            elif dest == "document_type":
                text = normalize_document_type(text, item.document_type if self.form_version.get() == "USA" else "receipt")
            elif dest == "currency":
                text = normalize_currency(text, "USD" if self.form_version.get() == "USA" else item.currency or "KRW")
            elif dest == "amount":
                hinted_currency = normalize_currency(text, "")
                if hinted_currency:
                    item.currency = hinted_currency
            setattr(item, dest, text)
        if self.form_version.get() == "USA":
            item.currency = "USD"
            item.document_type = normalize_document_type(item.document_type, "unknown")
            apply_primary_receipt_rotation_hint(item, data.get("rotation_degrees"))
            locked_document_type = normalize_document_type(item.locked_document_type, "unknown")
            if locked_document_type != "unknown":
                item.document_type = locked_document_type
            else:
                auto_classify_usa_document_type(
                    item,
                    data.get("vendor"),
                    data.get("confidence_notes"),
                    data.get("document_type"),
                )
            if item.document_type == "unknown" and looks_like_usa_payment_proof_text(
                item.filename,
                item.place,
                item.details,
                item.payment_method,
                item.receipt_label,
                data.get("vendor"),
                data.get("confidence_notes"),
            ):
                item.document_type = "payment_proof"
            if locked_document_type != "unknown":
                item.document_type = locked_document_type
            if item.document_type == "receipt":
                normalize_usa_fuel_receipt_fields(item, data.get("vendor"), data.get("confidence_notes"))
                normalize_usa_parking_receipt_fields(item, data.get("vendor"), data.get("confidence_notes"))
            elif item.document_type == "payment_proof":
                normalize_usa_parking_receipt_fields(item, data.get("vendor"), data.get("confidence_notes"))
        else:
            item.currency = normalize_currency(item.currency, "KRW")
            item.document_type = "receipt"
            item.matched_receipt_id = ""
            apply_primary_receipt_rotation_hint(item, data.get("rotation_degrees"))
        self.update_item_amount_fields(item, "amount")
        item.status = "AI filled"

    def apply_ai_data_to_bank_item(self, item: BankStatementItem, data: Dict[str, Any]) -> None:
        for source_key, dest_key in {
            "date": "date",
            "amount": "amount",
            "place": "place",
            "vendor": "place",
            "details": "place",
        }.items():
            value = data.get(source_key)
            if value is None:
                continue
            text = str(value).strip()
            if text:
                setattr(item, dest_key, text)
        item.status = payment_proof_missing_status(item) or "Needs match"

    def same_expense_date(self, left: str, right: str) -> bool:
        return same_expense_date_values(left, right)

    def match_bank_item(self, bank_item: BankStatementItem) -> None:
        match_bank_statement_to_receipt(
            bank_item,
            [item for item in self.items if is_usa_receipt_item(item)],
        )

    def match_all_bank_items(self) -> None:
        for bank_item in self.bank_items:
            self.match_bank_item(bank_item)
        self.resolve_payment_proof_conflicts()
        self.refresh_bank_tree()

    def resolve_payment_proof_conflicts(self) -> None:
        grouped: Dict[str, List[BankStatementItem]] = {}
        for bank_item in self.bank_items:
            if bank_item.matched_receipt_id:
                grouped.setdefault(bank_item.matched_receipt_id, []).append(bank_item)
        for proofs in grouped.values():
            if len(proofs) <= 1:
                continue
            keeper = proofs[0]
            keeper.status = "Matched"
            for proof_item in proofs[1:]:
                proof_item.matched_receipt_id = ""
                proof_item.status = "Multiple possible matches"

    def receipt_merge_key(self, item: ReceiptItem) -> Optional[Tuple[str, float]]:
        return usa_receipt_merge_key(item)

    def merge_same_usa_receipts(self) -> int:
        if self.form_version.get() != "USA":
            return 0
        merged_items, removed_to_target, merged_count = merge_same_usa_receipt_items(self.items)
        if not merged_count:
            return 0
        self.items = merged_items
        for proof_item in self.bank_items:
            if proof_item.matched_receipt_id in removed_to_target:
                proof_item.matched_receipt_id = removed_to_target[proof_item.matched_receipt_id]
        if self.selected_index is not None and self.selected_index >= len(self.items):
            self.selected_index = max(0, len(self.items) - 1) if self.items else None
        return merged_count

    def _submit_excel_image_prewarm(
        self,
        attachments: List[Dict[str, Any]],
        max_width: int,
        max_height: int,
        allow_upscale: bool = False,
        stretch_tiles: bool = False,
        trim_whitespace: bool = False,
        compact_whitespace: bool = False,
        max_columns: int = 2,
    ) -> None:
        if Image is None:
            return
        normalized = [copy.deepcopy(attachment) for attachment in (normalize_attachment(raw) for raw in attachments) if attachment]
        if not normalized:
            return
        cache_key = attachment_contact_sheet_cache_key(
            normalized,
            max_width,
            max_height,
            allow_upscale,
            stretch_tiles,
            trim_whitespace,
            compact_whitespace,
            max_columns,
        )
        if get_cached_contact_sheet(cache_key) is not None:
            return
        with self._excel_prewarm_lock:
            self._excel_prewarm_futures = [pending for pending in self._excel_prewarm_futures if not pending.done()]
            if cache_key is not None and cache_key in self._excel_prewarm_keys:
                return
            if cache_key is not None:
                self._excel_prewarm_keys.add(cache_key)
        try:
            future = self._excel_prewarm_executor.submit(
                prewarm_attachment_contact_sheet,
                normalized,
                max_width,
                max_height,
                allow_upscale,
                stretch_tiles,
                trim_whitespace,
                compact_whitespace,
                max_columns,
            )
        except RuntimeError:
            if cache_key is not None:
                with self._excel_prewarm_lock:
                    self._excel_prewarm_keys.discard(cache_key)
            return
        if cache_key is not None:
            def clear_prewarm_key(_future: Any, key: Tuple[Any, ...] = cache_key) -> None:
                with self._excel_prewarm_lock:
                    self._excel_prewarm_keys.discard(key)
            future.add_done_callback(clear_prewarm_key)
        with self._excel_prewarm_lock:
            self._excel_prewarm_futures.append(future)

    def _queue_usa_report_row_prewarm(self, row_data: Dict[str, Any]) -> None:
        image_padding = USA_SUPPORT_IMAGE_PADDING_PX
        receipt_image_width = max(1, USA_RECEIPT_IMAGE_MAX_WIDTH - image_padding * 2)
        receipt_image_height = max(1, USA_RECEIPT_IMAGE_MAX_HEIGHT - image_padding * 2)
        proof_image_width = max(1, USA_PROOF_IMAGE_MAX_WIDTH - image_padding * 2)
        proof_image_height = max(1, USA_PROOF_IMAGE_MAX_HEIGHT - image_padding * 2)
        receipt_images = list(row_data.get("receipt_images") or [])
        if receipt_images:
            self._submit_excel_image_prewarm(
                receipt_images,
                receipt_image_width,
                receipt_image_height,
                allow_upscale=True,
            )
        proof_images = list(row_data.get("proof_images") or [])
        if proof_images:
            self._submit_excel_image_prewarm(
                proof_images,
                proof_image_width,
                proof_image_height,
                allow_upscale=True,
                trim_whitespace=True,
                compact_whitespace=True,
                max_columns=1,
            )

    def _queue_korea_receipt_block_prewarm(self, block: Dict[str, Any]) -> None:
        attachments = list(block.get("attachments") or [])
        if not attachments:
            return
        mode = str(block.get("mode") or "half")
        slot = korea_receipt_slot(0, wide=mode == "wide", full_page=mode == "full")
        self._submit_excel_image_prewarm(
            attachments,
            int(slot["max_width"]),
            int(slot["max_height"]),
            allow_upscale=True,
            stretch_tiles=bool(block.get("stretch_tiles", False)),
        )

    def queue_excel_image_prewarm_for_item(self, item: ReceiptItem, form_version: Optional[str] = None) -> None:
        form = normalized_form_version(form_version or self.form_version.get())
        item_snapshot = copy.deepcopy(item)
        if form == "USA":
            attachments = ensure_receipt_images(item_snapshot)
            if not attachments:
                return
            if is_usa_payment_proof_item(item_snapshot):
                self._queue_usa_report_row_prewarm({"receipt_images": [], "proof_images": attachments})
            elif is_usa_receipt_item(item_snapshot):
                self._queue_usa_report_row_prewarm({"receipt_images": attachments, "proof_images": []})
            return
        for block in build_korea_receipt_blocks([item_snapshot], []):
            self._queue_korea_receipt_block_prewarm(block)

    def queue_excel_image_prewarm_for_bank_item(self, bank_item: BankStatementItem) -> None:
        self._queue_usa_report_row_prewarm({"receipt_images": [], "proof_images": [attachment_from_item(bank_item)]})

    def queue_excel_export_prewarm(self, form_version: Optional[str] = None) -> None:
        form = normalized_form_version(form_version or self.form_version.get())
        try:
            if form == "USA":
                report_rows = build_usa_report_rows(copy.deepcopy(self.items), copy.deepcopy(self.bank_items))
                for row_data in report_rows:
                    self._queue_usa_report_row_prewarm(row_data)
            else:
                sorted_items = sort_receipts_for_report(copy.deepcopy(self.items), "Korea")
                exchange_rate_attachments = [attachment_from_item(item) for item in copy.deepcopy(self.exchange_rate_items)]
                for block in build_korea_receipt_blocks(sorted_items, exchange_rate_attachments):
                    self._queue_korea_receipt_block_prewarm(block)
        except Exception:
            LOGGER.debug("Could not queue background Excel image preparation.", exc_info=True)

    def generate_empty_details(self) -> None:
        if self._receipt_import_active:
            self.show_info("Receipt files are still being added. Please wait a moment.")
            return
        self.save_current_fields()
        items = [item for item in self.items if receipt_needs_ai(item, self.form_version.get())]
        if not items:
            self.status_text.set(self.tr("no_empty_rows"))
            return
        self._run_ai_for_items(items, reload_selected=True, include_bank=True)

    def generate_selected_details(self) -> None:
        self.generate_empty_details()

    def generate_all_or_cancel(self) -> None:
        if self._ai_cancel_event is not None:
            self.cancel_generate_all_details()
            return
        self.generate_all_details()

    def cancel_generate_all_details(self) -> None:
        if self._ai_cancel_event is None:
            return
        self._ai_cancel_event.set()
        self.status_text.set(self.tr("cancel_requested"))
        try:
            self.generate_all_btn.configure(text=self.tr("cancel_generation"), state="disabled")
        except Exception:
            pass

    def generate_all_details(self) -> None:
        if self._receipt_import_active:
            self.show_info("Receipt files are still being added. Please wait a moment.")
            return
        self.save_current_fields()
        if not self.items:
            self.show_info("Upload receipts first.")
            return
        self._run_ai_for_items(list(self.items), reload_selected=True, include_bank=True)

    def _run_ai_for_items(
        self,
        items: List[ReceiptItem],
        reload_selected: bool,
        include_bank: bool = False,
    ) -> None:
        if self._busy:
            self.show_info("A batch is already running.")
            return
        bank_items = list(self.bank_items) if include_bank and self.form_version.get() == "USA" else []
        total = len(items) + len(bank_items)
        if total <= 0:
            return
        form_version = self.form_version.get()
        ai_concurrency = self.ai_concurrency_value()
        if not self.confirm_ai_concurrency(ai_concurrency, len(items), len(bank_items)):
            return
        if include_bank:
            self._details_ready_for_export = False
        cancel_event = threading.Event()
        self._ai_cancel_event = cancel_event
        self.set_busy(True)
        try:
            self.generate_all_btn.configure(text=self.tr("cancel_generation"), state="normal")
        except Exception:
            pass
        self.set_progress(0, total, f"0/{total}")
        self.status_text.set(f"Generating details for {total} image(s)...")

        def worker() -> None:
            successes = 0
            failures: List[str] = []
            completed = 0
            cancelled_result = "__cancelled__"
            progress_lock = threading.Lock()

            def next_done() -> int:
                nonlocal completed
                with progress_lock:
                    completed += 1
                    return completed

            def current_done() -> int:
                with progress_lock:
                    return completed

            def process_receipt(index: int, item: ReceiptItem) -> Optional[str]:
                if cancel_event.is_set():
                    return cancelled_result
                display_index = index + 1
                self.root.after(
                    0,
                    lambda i=item, n=display_index: self.mark_item_status(i, f"Processing {n}/{total}", current_done(), total),
                )
                self.set_status(f"Generating details {display_index}/{len(items)}: {item.filename}")
                try:
                    data = call_openai_receipt_extraction(
                        Path(item.path),
                        form_version,
                        item,
                        progress=self.set_status,
                    )
                    if cancel_event.is_set():
                        self.root.after(
                            0,
                            lambda i=item: self.mark_item_status(i, "Needs regen", current_done(), total),
                        )
                        return cancelled_result
                    self.apply_ai_data_to_item(item, data)
                    self.remember_receipt_item(item, "ai")
                    self.queue_excel_image_prewarm_for_item(item, form_version)
                    self.move_item_to_processed(item)
                    done = next_done()
                    self.root.after(
                        0,
                        lambda i=item, d=done: self.mark_item_status(i, "AI filled", d, total),
                    )
                    return None
                except Exception as exc:
                    setup_logging()
                    LOGGER.exception("AI extraction failed for %s", item.filename)
                    done = next_done()
                    self.root.after(
                        0,
                        lambda i=item, d=done: self.mark_item_status(i, "Failed", d, total),
                    )
                    return f"{item.filename}: {exc}"

            receipt_workers = min(ai_concurrency, max(1, len(items)))
            with ThreadPoolExecutor(max_workers=receipt_workers) as executor:
                futures = [executor.submit(process_receipt, idx, item) for idx, item in enumerate(items)]
                for future in as_completed(futures):
                    if cancel_event.is_set():
                        for pending in futures:
                            pending.cancel()
                    try:
                        failure = future.result()
                    except Exception as exc:
                        if cancel_event.is_set():
                            continue
                        failures.append(str(exc))
                        continue
                    if failure == cancelled_result:
                        continue
                    if failure:
                        failures.append(failure)
                    else:
                        successes += 1
            if include_bank and form_version == "USA" and not cancel_event.is_set():
                merged_count = self.merge_same_usa_receipts()
                match_usa_payment_proof_items(self.items)
                if merged_count:
                    self.root.after(0, self.refresh_tree)
                    self.set_status(f"Merged {merged_count} duplicate receipt screenshot(s).")
                self.queue_excel_export_prewarm(form_version)
            elif form_version == "Korea" and not cancel_event.is_set():
                self.queue_excel_export_prewarm(form_version)
            def process_bank_proof(index: int, bank_item: BankStatementItem) -> Optional[str]:
                if cancel_event.is_set():
                    return cancelled_result
                display_index = index + 1
                self.root.after(
                    0,
                    lambda i=bank_item, n=display_index: self.mark_bank_item_status(
                        i, f"Processing proof {n}/{len(bank_items)}", current_done(), total
                    ),
                )
                self.set_status(f"Reading payment proof {display_index}/{len(bank_items)}: {bank_item.filename}")
                try:
                    temp = ReceiptItem(
                        item_id=bank_item.item_id,
                        path=bank_item.path,
                        filename=bank_item.filename,
                        source_path=bank_item.source_path,
                        source_page=bank_item.source_page,
                        date=bank_item.date,
                        place=bank_item.place,
                        amount=bank_item.amount,
                        currency="USD",
                        document_type="payment_proof",
                        crop_box=bank_item.crop_box,
                        rotation_degrees=bank_item.rotation_degrees,
                    )
                    data = call_openai_receipt_extraction(
                        Path(bank_item.path),
                        "USA",
                        temp,
                        progress=self.set_status,
                    )
                    if cancel_event.is_set():
                        self.root.after(
                            0,
                            lambda i=bank_item: self.mark_bank_item_status(i, "Needs AI", current_done(), total),
                        )
                        return cancelled_result
                    self.apply_ai_data_to_bank_item(bank_item, data)
                    self.queue_excel_image_prewarm_for_bank_item(bank_item)
                    done = next_done()
                    self.root.after(
                        0,
                        lambda i=bank_item, d=done: self.mark_bank_item_status(i, i.status, d, total),
                    )
                    return None
                except Exception as exc:
                    setup_logging()
                    LOGGER.exception("Bank statement extraction failed for %s", bank_item.filename)
                    bank_item.status = "Missing date/amount"
                    done = next_done()
                    self.root.after(
                        0,
                        lambda i=bank_item, d=done: self.mark_bank_item_status(
                            i, "Missing date/amount", d, total
                        ),
                    )
                    return f"{bank_item.filename}: {exc}"
            if bank_items and not cancel_event.is_set():
                bank_workers = min(ai_concurrency, max(1, len(bank_items)))
                with ThreadPoolExecutor(max_workers=bank_workers) as executor:
                    futures = [executor.submit(process_bank_proof, idx, bank_item) for idx, bank_item in enumerate(bank_items)]
                    for future in as_completed(futures):
                        if cancel_event.is_set():
                            for pending in futures:
                                pending.cancel()
                        try:
                            failure = future.result()
                        except Exception as exc:
                            if cancel_event.is_set():
                                continue
                            failures.append(str(exc))
                            continue
                        if failure == cancelled_result:
                            continue
                        if failure:
                            failures.append(failure)
                        else:
                            successes += 1
            if bank_items and not cancel_event.is_set():
                for bank_item in bank_items:
                    if bank_item.status != "Missing date/amount":
                        self.match_bank_item(bank_item)
                self.resolve_payment_proof_conflicts()
                self.root.after(0, self.refresh_bank_tree)
                self.queue_excel_export_prewarm(form_version)
            final_completed = current_done()
            self.root.after(
                0,
                lambda: self.after_ai_complete(
                    successes,
                    failures,
                    total,
                    reload_selected,
                    include_bank,
                    cancelled=cancel_event.is_set(),
                    completed=final_completed,
                ),
            )

        threading.Thread(target=worker, daemon=True).start()

    def after_ai_complete(
        self,
        successes: int,
        failures: List[str],
        total: int,
        reload_selected: bool,
        ready_for_export: bool = False,
        cancelled: bool = False,
        completed: Optional[int] = None,
    ) -> None:
        done_count = total if completed is None else completed
        self._ai_cancel_event = None
        self.set_busy(False)
        try:
            self.generate_all_btn.configure(text=self.tr("generate_all"), state="normal")
        except Exception:
            pass
        if ready_for_export and successes > 0:
            self._details_ready_for_export = True
        self.set_progress(done_count, total, f"{done_count}/{total}")
        if ready_for_export and successes > 0:
            self.sort_items_for_current_report()
        else:
            self.refresh_tree()
        self.refresh_bank_tree()
        if reload_selected:
            self.load_selected_into_fields()
            self.update_preview()
        self.save_session()
        if cancelled:
            self.status_text.set(self.tr("generation_cancelled", done=done_count, total=total))
            self.update_toolbar_recommendation()
            return
        if failures:
            self.status_text.set(f"AI finished: {successes} succeeded, {len(failures)} failed.")
            if messagebox:
                shown = "\n".join(failures[:5])
                if len(failures) > 5:
                    shown += f"\n...and {len(failures) - 5} more."
                messagebox.showwarning(
                    self.tr("generate_empty"),
                    f"Finished with failures.\n\n{shown}\n\nLogged to:\n{LOG_DIR / 'app.log'}",
                )
        else:
            self.status_text.set("AI details generated. Review before exporting.")
        self.update_toolbar_recommendation()

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
                if (
                    button is self.generate_all_btn
                    and busy
                    and self._ai_cancel_event is not None
                    and not self._ai_cancel_event.is_set()
                ):
                    button.configure(state="normal", text=self.tr("cancel_generation"))
                else:
                    button.configure(state=state)
            except Exception:
                pass
        if not busy:
            try:
                self.generate_all_btn.configure(text=self.tr("generate_all"))
            except Exception:
                pass
        if self.select_payment_proof_btn is not None:
            proof_state = "normal" if not busy else "disabled"
            try:
                self.select_payment_proof_btn.configure(state=proof_state)
            except Exception:
                pass
        self.update_toolbar_recommendation()

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

    def mark_bank_item_status(self, item: BankStatementItem, status: str, done: int, total: int) -> None:
        item.status = status
        self.refresh_bank_tree()
        self.set_progress(done, total, f"{done}/{total}")

    def move_item_to_processed(self, item: ReceiptItem) -> bool:
        source = Path(item.path)
        if not source.exists() or not same_folder(source, UNPROCESSED_DIR):
            return False
        try:
            ensure_runtime_folders()
            destination = unique_path(PROCESSED_DIR, source.name)
            shutil.move(str(source), str(destination))
            self.apply_source_rename_to_item(item, source, destination)
            item.path = str(destination)
            item.filename = destination.name if len(ensure_receipt_images(item)) <= 1 else item.filename
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
        if self._receipt_import_active:
            self.show_info("Receipt files are still being added. Please wait a moment.")
            return
        self.save_current_fields()
        if not self.items:
            self.show_info("Upload at least one receipt before generating Excel.")
            return
        form_version = self.form_version.get()
        if form_version == "USA":
            for item in self.items:
                auto_classify_usa_document_type(item)
            unknown_items = [
                item
                for item in self.items
                if normalize_document_type(item.document_type, "unknown") == "unknown"
            ]
            if unknown_items:
                files = ", ".join(item.filename for item in unknown_items[:5])
                if len(unknown_items) > 5:
                    files += f", +{len(unknown_items) - 5}"
                self.status_text.set(self.tr("unknown_document_type", count=len(unknown_items), files=files))
                self.refresh_tree()
                try:
                    first_index = self.items.index(unknown_items[0])
                    self.selected_index = first_index
                    self.tree.selection_set(str(first_index))
                    self.tree.focus(str(first_index))
                    self.tree.see(str(first_index))
                    self.load_selected_into_fields()
                    self.update_preview()
                except Exception:
                    pass
                if messagebox:
                    shown_files = "\n".join(item.filename for item in unknown_items[:10])
                    if len(unknown_items) > 10:
                        shown_files += f"\n...and {len(unknown_items) - 10} more."
                    messagebox.showwarning(
                        self.tr("manual_review_title"),
                        self.tr(
                            "unknown_document_type_detail",
                            count=len(unknown_items),
                            files=shown_files,
                        ),
                    )
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
        if form_version == "USA":
            self.merge_same_usa_receipts()
            match_usa_payment_proof_items(self.items)
            self.refresh_tree()
        items_snapshot = copy.deepcopy(self.items)
        bank_snapshot = copy.deepcopy(self.bank_items)
        exchange_snapshot = copy.deepcopy(self.exchange_rate_items)
        exchange_rate = safe_float(self.exchange_rate.get()) or safe_float(DEFAULT_USD_TO_RMB_RATE) or 6.8
        krw_rate = safe_float(self.krw_to_rmb_rate.get()) or safe_float(DEFAULT_KRW_TO_RMB_RATE) or 0.0046
        usd_krw_rate = safe_float(self.usd_to_krw_rate.get()) or safe_float(DEFAULT_USD_TO_KRW_RATE) or 1548.86
        self.set_busy(True)
        self.set_progress(0, 1, "0/1")
        self.status_text.set("Generating Excel workbook...")

        def export_progress(done: int, total: int, message: str) -> None:
            label = f"{done}/{max(1, total)}"

            def apply() -> None:
                self.set_progress(done, total, label)
                self.status_text.set(message)

            self.root.after(0, apply)

        def worker() -> None:
            try:
                if form_version == "USA":
                    export_usa(items_snapshot, output_path, exchange_rate, bank_snapshot, progress=export_progress)
                else:
                    export_korea(
                        items_snapshot,
                        output_path,
                        krw_rate,
                        exchange_rate,
                        usd_krw_rate,
                        exchange_snapshot,
                        progress=export_progress,
                    )
            except Exception as exc:
                log_path = log_exception("Excel export failed")
                error_message = str(exc)

                def fail() -> None:
                    self.set_busy(False)
                    self.set_progress(0, 1, "")
                    if messagebox:
                        messagebox.showerror(
                            "Generate Excel",
                            f"Could not generate workbook:\n{error_message}\n\nLogged to:\n{log_path}",
                        )
                    self.status_text.set("Export failed.")

                self.root.after(0, fail)
                return

            def done() -> None:
                self.set_busy(False)
                self.set_progress(1, 1, "1/1")
                self.last_generated_output_path = output_path
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

            self.root.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    def open_output_folder(self) -> None:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        target = self.last_generated_output_path
        if target is None or not target.exists():
            workbooks = list(OUTPUT_DIR.glob("*.xlsx"))
            target = max(workbooks, key=lambda path: path.stat().st_mtime, default=None)
        if target is not None and target.exists():
            try:
                reveal_in_file_explorer(target)
                return
            except Exception:
                log_exception("Could not reveal output workbook")
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
            "usd_to_krw_rate": self.usd_to_krw_rate.get(),
            "krw_to_rmb_rate": self.krw_to_rmb_rate.get(),
            "selected_index": self.selected_index,
            "items": [asdict(item) for item in self.items],
            "bank_items": [asdict(item) for item in self.bank_items],
            "exchange_rate_items": [asdict(item) for item in self.exchange_rate_items],
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
        raw_bank_items = data.get("bank_items")
        raw_exchange_rate_items = data.get("exchange_rate_items")
        if not isinstance(raw_items, list):
            raw_items = []
        if not isinstance(raw_bank_items, list):
            raw_bank_items = []
        if not isinstance(raw_exchange_rate_items, list):
            raw_exchange_rate_items = []
        if not raw_items and not raw_bank_items and not raw_exchange_rate_items:
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
            if not isinstance(values.get("receipt_images"), list):
                values["receipt_images"] = []
            values["rotation_degrees"] = normalize_rotation(values.get("rotation_degrees", 0))
            values["document_type"] = normalize_document_type(values.get("document_type"), "receipt")
            locked = normalize_document_type(values.get("locked_document_type"), "unknown")
            values["locked_document_type"] = locked if locked != "unknown" else ""
            values["matched_receipt_id"] = str(values.get("matched_receipt_id") or "")
            item = ReceiptItem(**values)
            ensure_receipt_images(item)
            restored.append(item)
        if not restored:
            restored = []
        bank_fields = set(BankStatementItem.__dataclass_fields__.keys())
        restored_bank: List[BankStatementItem] = []
        for raw in raw_bank_items:
            if not isinstance(raw, dict):
                continue
            values = {key: raw.get(key, "") for key in bank_fields}
            if not values.get("item_id"):
                values["item_id"] = uuid.uuid4().hex
            if not isinstance(values.get("crop_box"), list):
                values["crop_box"] = None
            values["rotation_degrees"] = normalize_rotation(values.get("rotation_degrees", 0))
            restored_bank.append(BankStatementItem(**values))
        restored_exchange: List[BankStatementItem] = []
        for raw in raw_exchange_rate_items:
            if not isinstance(raw, dict):
                continue
            values = {key: raw.get(key, "") for key in bank_fields}
            if not values.get("item_id"):
                values["item_id"] = uuid.uuid4().hex
            if not isinstance(values.get("crop_box"), list):
                values["crop_box"] = None
            values["rotation_degrees"] = normalize_rotation(values.get("rotation_degrees", 0))
            restored_exchange.append(BankStatementItem(**values))
        self.form_version.set(normalized_form_version(data.get("form_version") or self.form_version.get()))
        self.exchange_rate.set(str(data.get("usa_exchange_rate") or DEFAULT_USD_TO_RMB_RATE))
        self.usd_to_krw_rate.set(str(data.get("usd_to_krw_rate") or DEFAULT_USD_TO_KRW_RATE))
        self.krw_to_rmb_rate.set(str(data.get("krw_to_rmb_rate") or DEFAULT_KRW_TO_RMB_RATE))
        self.items = restored
        self.bank_items = restored_bank
        self.exchange_rate_items = restored_exchange
        self._last_form_version = self.form_version.get()
        try:
            selected = int(data.get("selected_index", 0) or 0)
        except (TypeError, ValueError):
            selected = 0
        self.selected_index = None
        self._on_form_version_changed()
        self.refresh_tree()
        self.refresh_bank_tree()
        if self.items:
            self.select_index(max(0, min(selected, len(self.items) - 1)))
        self.status_text.set("Previous reimbursement session restored.")

    def on_close(self) -> None:
        self.settings["usa_exchange_rate"] = self.exchange_rate.get()
        self.settings["usd_to_krw_rate"] = self.usd_to_krw_rate.get()
        self.settings["krw_to_rmb_rate"] = self.krw_to_rmb_rate.get()
        self.settings["ai_concurrency"] = self.ai_concurrency_value()
        self.settings["last_form_version"] = normalized_form_version(self.form_version.get())
        self.settings["language"] = normalize_language(self.language.get())
        self.settings["usa_upload_mode"] = normalize_usa_upload_mode(self.usa_upload_mode.get())
        save_user_settings(self.settings)
        self.save_session()
        try:
            self._excel_prewarm_executor.shutdown(wait=False, cancel_futures=True)
        except Exception:
            pass
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
        safe_float(DEFAULT_USD_TO_KRW_RATE) or 1548.86,
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
