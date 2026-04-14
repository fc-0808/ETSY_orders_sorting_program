#!/usr/bin/env python3
"""
Simple desktop UI for the shopping-route script.

Plain language for anyone new — no Etsy jargon required.
Uses tkinter (included with Python on Windows). Double-click run_ui.ps1 or run from project folder.
UI language: defaults to 中文; use the language button for English.
Layout: title + quick-open row, then the gray step strip, then tabbed panels (PDF / Charms drop zones mirror each other).
"""

from __future__ import annotations

import os
import queue
import re
import secrets
import shutil
import subprocess
import sys
import threading
from datetime import date
from collections.abc import Callable
from io import BytesIO
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

PROJECT_ROOT = Path(__file__).resolve().parent.parent
GENERATOR = PROJECT_ROOT / "src" / "generate_shopping_route.py"

try:
    from generate_shopping_route import (
        CATALOG_SHEET,
        CHARM_LIBRARY_SHEET,
        CHARM_SHOPS_SHEET,
        ProductMapPickerRow,
        extract_photos_from_xlsx,
        list_product_map_rows_for_picker,
        load_charm_library,
        load_charm_shops,
    mark_product_map_discontinued_by_row,
    reorder_charm_library_rows,
    update_product_map_cells,
    update_product_map_photo,
)
except ImportError:
    CATALOG_SHEET = "Product Map"  # type: ignore[assignment, misc]
    CHARM_LIBRARY_SHEET = "Charm Library"  # type: ignore[assignment, misc]
    CHARM_SHOPS_SHEET = "Charm Shops"  # type: ignore[assignment, misc]
    ProductMapPickerRow = None  # type: ignore[assignment, misc]
    extract_photos_from_xlsx = None  # type: ignore[assignment, misc]
    list_product_map_rows_for_picker = None  # type: ignore[assignment, misc]
    load_charm_library = None  # type: ignore[assignment, misc]
    load_charm_shops = None  # type: ignore[assignment, misc]
    mark_product_map_discontinued_by_row = None  # type: ignore[assignment, misc]
    reorder_charm_library_rows = None  # type: ignore[assignment, misc]
    update_product_map_cells = None  # type: ignore[assignment, misc]
    update_product_map_photo = None  # type: ignore[assignment, misc]

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None  # type: ignore[assignment, misc]
    ImageTk = None  # type: ignore[assignment, misc]

try:
    from supplier_catalog_backup import (
        backup_supplier_catalog_before_write,
        catalog_backup_dir,
        list_supplier_catalog_backups,
        restore_supplier_catalog,
    )
except ImportError:
    backup_supplier_catalog_before_write = None  # type: ignore[assignment, misc]
    catalog_backup_dir = None  # type: ignore[assignment, misc]
    list_supplier_catalog_backups = None  # type: ignore[assignment, misc]
    restore_supplier_catalog = None  # type: ignore[assignment, misc]

# Same layout as generate_shopping_route.py with --project-dir .
DATA_DIR = PROJECT_ROOT / "data"
OUTPUT_DIR = PROJECT_ROOT / "output"
CHARM_IMAGES_DIR_NAME = "charm_images"
DEFAULT_CHARM_IMAGES_DIR = DATA_DIR / CHARM_IMAGES_DIR_NAME
# Staged copies use this prefix; import matches them via charm_import_pattern_argv() (must not match CH-*.png).
CHARM_INCOMING_PREFIX = "__incoming__"
CHARM_INCOMING_PATTERN = f"{CHARM_INCOMING_PREFIX}*"
CHARM_IMAGE_EXTS = frozenset({".png", ".jpg", ".jpeg", ".webp"})


def charm_import_pattern_argv() -> list[str]:
    """Globs for --import-charm-images: UI-staged copies and classic Screenshot*.png files."""
    return [
        "--import-charm-pattern",
        CHARM_INCOMING_PATTERN,
        "--import-charm-pattern",
        "Screenshot*.png",
    ]
ORDER_PDF_EXT = ".pdf"
DEFAULT_ORDER_INPUT_DIR = PROJECT_ROOT / "input"
DEFAULT_BACKUP_DIR = PROJECT_ROOT / "backup"
_ORDER_PDF_MMdd_SUFFIX = re.compile(r"_(\d{4})\.pdf$", re.IGNORECASE)
# Match Charms tab drop target height for a consistent look.
DROP_ZONE_H = 92

FILE_SUPPLIER_CATALOG = DATA_DIR / "supplier_catalog.xlsx"
FILE_SHOPPING_ROUTE = OUTPUT_DIR / "shopping_route.xlsx"
FILE_SHOPPING_ROUTE_ZH = OUTPUT_DIR / "shopping_route_zh.xlsx"
FILE_SHOPPING_HTML = OUTPUT_DIR / "shopping_route.html"
FILE_SHOPPING_HTML_ZH = OUTPUT_DIR / "shopping_route_zh.html"

windnd = None
if sys.platform == "win32":
    try:
        import windnd  # type: ignore[import-untyped, unused-ignore]
    except ImportError:
        windnd = None

# Readable, calm palette (works with ttk “clam” on Windows).
COLORS = {
    "app": "#eef2f7",
    "card": "#ffffff",
    "hero": "#1d4ed8",
    "accent": "#1e40af",
    "accent_soft": "#dbeafe",
    "run": "#047857",
    "run_hover": "#065f46",
    "text": "#0f172a",
    "muted": "#64748b",
    "border": "#cbd5e1",
    "strip": "#dbeafe",
    "strip_text": "#1e3a8a",
    "strip_accent": "#2563eb",
    "log_bg": "#f8fafc",
    "drop_zone": "#eff6ff",
    "drop_border": "#93c5fd",
    "drop_hint": "#3b82f6",
    "separator": "#e2e8f0",
}

# Window layout (logical pixels; OS display scaling still applies).
UI_DEFAULT_W = 1240
UI_DEFAULT_H = 1020
UI_MIN_W = 1020
UI_MIN_H = 760
UI_MARGIN_X = 16
UI_LOG_LINES = 10

Lang = str  # "en" | "zh"

# Main tasks: id -> (title, blurb) per language
JOB_TEXT: dict[str, dict[Lang, tuple[str, str]]] = {
    "new_batch": {
        "en": (
            "Process new order PDFs",
            "Use when you have new Etsy order PDFs. Add them with the drop zone above or the input folder — "
            "only files never processed before are read. Then run this job (or use «Run: process new order PDFs» "
            "in that panel). This updates your shopping list files.",
        ),
        "zh": (
            "处理新的订单 PDF",
            "有新的 Etsy 订单 PDF 时使用。可用上方拖放区或 input 文件夹加入文件——只会处理从未处理过的 PDF。"
            "然后选本项并点绿色「运行」，或点拖放区旁的「运行：处理新订单 PDF」。会更新采购清单相关文件。",
        ),
    },
    "refresh_catalog": {
        "en": (
            "Rebuild the shopping list (no new PDFs)",
            "Use after editing the supplier product spreadsheet (supplier_catalog.xlsx). "
            "Uses saved order data — you do not need new PDFs.",
        ),
        "zh": (
            "重新生成采购清单（不需要新 PDF）",
            "在修改供应商商品表（supplier_catalog.xlsx）后使用。沿用已保存的订单数据，不需要新的 PDF。",
        ),
    },
    "purge_purchased": {
        "en": (
            "Clean up after shopping",
            "Use after you opened the shopping list spreadsheet and marked what you already bought. "
            "This removes finished items from the list and updates saved data.",
        ),
        "zh": (
            "购物后清理清单",
            "在采购清单表格里标记「已买」等项目后使用。会从清单中移除已完成项并更新保存的数据。",
        ),
    },
    "rebuild_catalog": {
        "en": (
            "Re-sort the supplier spreadsheet only",
            "Reorders rows in supplier_catalog.xlsx by floor/stall. "
            "If you also need new shopping list files, run “Rebuild the shopping list” afterward.",
        ),
        "zh": (
            "仅重新排序供应商表格",
            "按楼层/摊位重排 supplier_catalog.xlsx。若还需要新的采购清单文件，请之后再选「重新生成采购清单」。",
        ),
    },
    "reset": {
        "en": (
            "Advanced: start over from PDFs only",
            "Ignores saved progress and rebuilds from whatever PDFs are in the input folder now. "
            "Only use this if your manager asked — it is stronger than a normal refresh.",
        ),
        "zh": (
            "高级：仅从 PDF 重新开始",
            "忽略已保存进度，只根据 input 文件夹里现有的 PDF 重建。除非主管要求，否则不要随意使用。",
        ),
    },
}

JOB_ORDER = list(JOB_TEXT.keys())

# Charm: one clear step at a time (avoids mixing incompatible flags). Internal id -> argv flags.
CHARM_MODE_FLAGS: dict[str, list[str]] = {
    "skip": [],
    "setup_excel": ["--init-charm-shops"],
    "try_add_photos": ["--import-charm-images", "--import-charm-dry-run"],
    "add_photos": ["--import-charm-images"],
    "export_list": ["--export-charm-manifest"],
    "try_fill_codes": ["--fill-charm-sku-dry-run"],
    "fill_codes": ["--fill-charm-sku"],
    "try_renumber": ["--renumber-charms", "--renumber-charms-dry-run"],
    "renumber": ["--renumber-charms"],
}

# Order shown in the UI
CHARM_MODE_ORDER = [
    "skip",
    "setup_excel",
    "try_add_photos",
    "add_photos",
    "export_list",
    "try_fill_codes",
    "fill_codes",
    "try_renumber",
    "renumber",
]

# id -> (radio title, one-line hint) per language — keep titles short; details live in charm_run_note / section A.
CHARM_MODE_TEXT: dict[str, dict[Lang, tuple[str, str]]] = {
    "skip": {
        "en": (
            "None (default)",
            "Green Run on tab 1 runs only the main job you selected there.",
        ),
        "zh": (
            "不用（默认）",
            "绿色「运行」只执行第 1 页已选的主要任务。",
        ),
    },
    "setup_excel": {
        "en": (
            "One-time: add Charm sheets to the workbook",
            "Creates Charm Shops + Charm Library if missing. Does not import images.",
        ),
        "zh": (
            "一次性：在商品表里添加挂饰工作表",
            "若无则创建 Charm Shops、Charm Library；不导入图片。",
        ),
    },
    "try_add_photos": {
        "en": (
            "Preview charm import (dry run)",
            "Prints planned renames only — no changes to disk or Excel.",
        ),
        "zh": (
            "预览挂饰导入（试运行）",
            "仅在输出中列出将重命名的文件，不写磁盘、不改 Excel。",
        ),
    },
    "add_photos": {
        "en": (
            "Import charm images (via Run)",
            "Same result as Import into workbook above; this path uses green Run instead.",
        ),
        "zh": (
            "导入挂饰图（通过「运行」）",
            "与上方「导入到工作簿」效果相同，只是改用绿色「运行」执行。",
        ),
    },
    "export_list": {
        "en": (
            "Export charm index (JSON)",
            "Sidecar file for web or scripts — not the shopping list.",
        ),
        "zh": (
            "导出挂饰索引（JSON）",
            "供网页或脚本用的辅助文件，不是采购清单。",
        ),
    },
    "try_fill_codes": {
        "en": (
            "Preview AI SKU fill (dry run)",
            "Lists target rows; no API calls or saves.",
        ),
        "zh": (
            "预览 AI 填写 SKU（试运行）",
            "列出将处理的行；不调接口、不保存。",
        ),
    },
    "fill_codes": {
        "en": (
            "Fill empty SKUs in Charm Library (AI)",
            "Needs CHARM_VISION_* / OPENAI_* (or local base URL) set on this PC.",
        ),
        "zh": (
            "用 AI 补全 Charm Library 的空 SKU",
            "需本机已配置 CHARM_VISION_* / OPENAI_* 或本机视觉接口地址。",
        ),
    },
    "try_renumber": {
        "en": (
            "Preview renumber charm codes (dry run)",
            "Shows the planned old → new code mapping without making any changes.",
        ),
        "zh": (
            "预览重新编号挂饰代码（试运行）",
            "仅显示旧代码→新代码的对应关系，不写入任何修改。",
        ),
    },
    "renumber": {
        "en": (
            "Renumber charm codes (apply)",
            "Reassigns codes by row order, updates Product Map references and renames image files.",
        ),
        "zh": (
            "重新编号挂饰代码（应用）",
            "按行顺序重新分配代码，自动更新商品表引用并重命名图片文件。",
        ),
    },
}

# Fixed chrome strings
CHROME: dict[Lang, dict[str, str]] = {
    "en": {
        "win_title": "Shopping list builder",
        "lang_btn": "中文",
        "header_title": "Turn order PDFs into shopping lists and store maps.",
        "tab_orders": "1 — Orders & outputs",
        "tab_charms": "2 — Charms (optional)",
        "step_strip": (
            "  ①  Add PDFs (drop zone / Browse / «Open input folder» on this tab)   →   ②  Pick ONE main job (left)   →   "
            "③  Green Run below (or «Run: process new order PDFs» after adding PDFs)"
        ),
        "footer_run_hint": "Read the messages in the output box when it finishes.",
        "quick_hint": (
            "Quick open — opens in Excel or your browser when the file exists; run the tool first if not. "
            "Top group: supplier catalog file and the data folder it lives in. "
            "Bottom group: shopping list outputs (use Chinese / HTML checkboxes before Run when needed)."
        ),
        "quick_group_supplier": "Supplier — product data (catalog lives here)",
        "quick_group_route": "Shopping list — generated outputs",
        "quick_catalog": "Open supplier workbook",
        "btn_catalog_backups": "Catalog backups…",
        "cat_backup_title": "Supplier catalog backups",
        "cat_backup_hint": (
            "Timestamped copies are saved automatically before this program changes "
            "supplier_catalog.xlsx. Close Excel before restoring."
        ),
        "cat_backup_open_folder": "Open backup folder",
        "cat_backup_snapshot": "Snapshot now",
        "cat_backup_snapshot_ok": "Snapshot saved:\n{path}",
        "cat_backup_snapshot_no_file": "supplier_catalog.xlsx was not found — nothing to copy.",
        "cat_backup_restore": "Restore selected",
        "cat_backup_none": "(No backups yet — run any catalog action to create one.)",
        "cat_backup_restore_confirm": (
            "Replace the current supplier_catalog.xlsx with this backup?\n\n{path}\n\n"
            "Close Excel first. The current file will be copied to the backup folder first."
        ),
        "cat_backup_restore_ok": "Catalog restored from backup.",
        "cat_backup_restore_pick": "Select a backup in the list first.",
        "cat_backup_close": "Close",
        "btn_edit_products": "Edit products…",
        "edit_title": "Edit Product Map",
        "edit_heading": "Assign shop, stall, charm code — or mark as discontinued",
        "edit_intro": (
            "Select a product row, then use the fields on the right to assign or change "
            "the Shop Name, Stall, Charm Shop, and Charm Code — or mark the product as "
            "discontinued to remove it from future shopping lists."
        ),
        "edit_search_label": "Search",
        "edit_search_tip": "Matches title, shop, or stall",
        "edit_col_photo": "Photo",
        "edit_col_row": "Row",
        "edit_col_title": "Title (short)",
        "edit_col_shop": "Shop",
        "edit_col_stall": "Stall",
        "edit_col_charm_shop": "Charm Shop",
        "edit_col_charm_code": "Charm Code",
        "edit_lbl_shop": "Shop Name",
        "edit_lbl_stall": "Stall",
        "edit_lbl_charm_shop": "Charm Shop",
        "edit_lbl_charm_code": "Charm Code",
        "edit_btn_save": "Save",
        "edit_btn_close": "Close",
        "edit_saved": "Saved — {title}",
        "edit_no_selection": "Select a product row in the table first.",
        "edit_no_import": "Catalog helpers are unavailable (check installation).",
        "edit_empty": "No product rows on the Product Map sheet.",
        "edit_preview_title": "Edit — selected product",
        "edit_preview_placeholder": "Click a row to edit its fields.",
        "edit_hover_tip": "Tip: rest the pointer on a thumbnail in the Photo column — a larger photo pops up next to the cursor.",
        "edit_btn_upload_photo": "Upload new photo\u2026",
        "edit_upload_photo_title": "Choose product photo",
        "edit_upload_photo_pending": "\u2605 New photo ready — click Save to apply",
        "edit_photo_saved": "Photo updated — {title}",
        "edit_photo_save_fail": "Could not save photo: {err}",
        "edit_charm_pick": "Pick…",
        "edit_charm_unknown": "Code not found in Charm Library.",
        "charm_picker_title": "Choose a charm",
        "charm_picker_heading": "Click a charm to select it",
        "charm_picker_search": "Filter",
        "charm_picker_search_tip": "Matches code or SKU",
        "charm_picker_col_photo": "Photo",
        "charm_picker_col_code": "Code",
        "charm_picker_col_sku": "SKU",
        "charm_picker_col_shop": "Default Shop",
        "charm_picker_none": "(none — clear charm code)",
        "charm_picker_hover_tip": "Tip: hover over the Photo column — a larger image pops up next to the cursor.",
        "edit_danger_zone": "Danger zone",
        "edit_danger_note": "This action moves the product permanently out of the active catalog.",
        "edit_btn_discontinue": "Mark as discontinued\u2026",
        "edit_discontinue_confirm_title": "Mark as discontinued",
        "edit_discontinue_confirm": (
            "Move this product to the Discontinued Products sheet?\n\n"
            "{title}\n\n"
            "It will no longer appear on future shopping lists.\n"
            "The record is kept for reference — it is not deleted."
        ),
        "edit_discontinue_done": (
            "Moved to Discontinued Products sheet.\n"
            "Run \u00abRebuild the shopping list\u00bb next to update the route."
        ),
        "discontinue_title": "Mark product as discontinued",
        "discontinue_heading": "Stop matching a catalog product",
        "discontinue_intro": (
            "When a supplier no longer sells an item, mark it here. The product row is moved "
            "from Product Map to the Discontinued Products sheet — it stays for your records "
            "but new orders will not match it."
        ),
        "discontinue_step1": "1. Use the photo plus Shop and Stall to confirm the correct row.",
        "discontinue_step2": "2. Click a row, check the preview, then «Mark selected» (or double-click). The product is moved to a Discontinued Products sheet.",
        "discontinue_step3": "3. Run «Rebuild the shopping list» on the main tab so the route updates.",
        "discontinue_search_label": "Search",
        "discontinue_search_tip": "Matches title, shop, or stall",
        "discontinue_hover_tip": "Tip: rest the pointer on a thumbnail in the Image column — a larger photo pops up next to the cursor.",
        "discontinue_preview_title": "Preview — selected product",
        "discontinue_preview_placeholder": "Click a row to see the full title and a larger photo.",
        "discontinue_col_thumb": "Photo",
        "discontinue_col_row": "Row",
        "discontinue_col_shop": "Shop",
        "discontinue_col_stall": "Stall",
        "discontinue_col_title": "Title (short)",
        "discontinue_apply": "Mark selected as discontinued",
        "discontinue_done": "Product moved to Discontinued Products sheet. Regenerate the shopping list next.",
        "discontinue_empty": "No product rows on the Product Map sheet.",
        "discontinue_close": "Close",
        "discontinue_no_import": "Catalog helpers are unavailable (check installation).",
        "discontinue_no_pillow": (
            "Install Pillow for product thumbnails:  pip install Pillow  (then restart this app)."
        ),
        "discontinue_no_selection": "Select a product row in the table first.",
        "quick_data_folder": "Open data folder",
        "quick_route": "Excel (default)",
        "quick_route_zh": "Excel — Chinese",
        "quick_html": "Web page",
        "quick_html_zh": "Web — Chinese",
        "file_missing_title": "File not found",
        "file_missing_body": "This file is not there yet:\n{path}\nRun the tool above first, or ask a supervisor. For Chinese or web files, tick those options before running.",
        "file_open_fail_title": "Could not open",
        "file_open_fail_body": "{err}",
        "main_frame": "Main job (choose one)",
        "also_frame": "Also create",
        "chk_excel_zh": "Chinese shopping list (Excel)",
        "chk_html": "Phone-friendly web page of the list",
        "chk_no_catalog": "Do not add new products to the supplier spreadsheet this run",
        "charm_drop_section": "A — Add a new charm picture to the catalog",
        "charm_how_body": (
            "What actually happens\n"
            "• Drag or Browse copies your file into the charm images folder (Open charm folder) under a temporary name.\n"
            "• Import into workbook renames it on disk to the next free code (e.g. CH-00012.png) and adds one new row on the "
            "Charm Library sheet in supplier_catalog.xlsx, with the picture embedded. That is the master list of charms.\n"
            "• It does not add rows to Charm Shops — that sheet is only your list of physical market stalls (setup once).\n"
            "• First-time shop setup: run One-time: add Charm sheets… in section B once (supervisor), then imports work.\n"
            "Formats: PNG, JPG/JPEG, WebP."
        ),
        "charm_drop_blurb": "Drop or browse, then click Import into workbook — that is the normal path for new photos.",
        "charm_run_section": "B — Optional: charm task for green Run",
        "charm_run_note": (
            "Choose one option below, or leave None. If you pick any charm task, the next green Run performs only that task "
            "(tab 1’s main job is skipped for that run). To add new charm photos in the usual way, use section A and "
            "Import into workbook — you rarely need section B for that."
        ),
        "charm_smart": "Also use AI for SKU column (C) when Run does a charm import — PC must be configured by a supervisor",
        "charm_drop_hint_dnd": "Drop image files here",
        "charm_drop_hint_no_dnd": "Drag-and-drop: on Windows, install “windnd” (pip install windnd) and restart this app. You can always use Browse.",
        "charm_drop_browse": "Browse for photos…",
        "charm_drop_browse_title": "Choose charm photos",
        "charm_drop_import": "Import into workbook (rename + Excel)",
        "charm_drop_open_folder": "Open charm images folder",
        "charm_drop_vision": (
            "Suggest SKU with AI when importing (env CHARM_VISION_* / OPENAI_*; "
            "or CHARM_VISION_BASE_URL for Ollama, OpenRouter, LM Studio, …)"
        ),
        "charm_msg_title": "Charm photos",
        "charm_drop_no_valid": "None of those files are supported images (PNG, JPG/JPEG, WebP).",
        "charm_drop_nothing_to_import": "No staged photos to import. Drop files or use Browse first.",
        "charm_drop_staged": "Staged {n} photo(s). Click «Import into workbook» to rename and update Excel.\n",
        "charm_import_start": "\n--- Charm image import ---\n",
        "charm_reorder_section": "C — Reorder charms",
        "charm_reorder_body": (
            "Drag charms up or down in the panel to change their order — "
            "no need to touch the Excel file.  "
            "After reordering, codes are reassigned CH-00001, CH-00002 … by position "
            "and all Product Map references update automatically."
        ),
        "charm_reorder_btn_open": "Open charm reorder panel\u2026",
        "charm_reorder_start_preview": "\n--- Renumber charm codes (preview) ---\n",
        "charm_reorder_start_apply": "\n--- Renumber charm codes ---\n",
        "charm_reorder_nothing": "Charm Library not found or empty — nothing to reorder.",
        "charm_reorder_no_import": "Catalog helpers are unavailable (check installation).",
        # Reorder dialog strings
        "reorder_title": "Reorder Charm Library",
        "reorder_heading": "Drag and drop to reorder",
        "reorder_intro": (
            "Drag rows up or down, or select a row and use the \u2191 \u2193 buttons.  "
            "The \"New Code\" column shows the code each charm will receive after applying.  "
            "Click Apply to save — the Excel file and Product Map are updated automatically."
        ),
        "reorder_col_photo": "Photo",
        "reorder_col_code": "Current Code",
        "reorder_col_new": "New Code",
        "reorder_col_sku": "SKU",
        "reorder_col_shop": "Default Shop",
        "reorder_btn_up": "\u2191  Move up",
        "reorder_btn_down": "\u2193  Move down",
        "reorder_btn_top": "\u21d1  To top",
        "reorder_btn_bottom": "\u21d3  To bottom",
        "reorder_btn_apply": "Apply reorder & renumber",
        "reorder_btn_close": "Close",
        "reorder_confirm_title": "Apply reorder",
        "reorder_confirm_body": (
            "This will rewrite the Charm Library rows in the new order,\n"
            "reassign codes CH-00001, CH-00002 \u2026 by position,\n"
            "and update all Product Map references.\n\n"
            "Make sure supplier_catalog.xlsx is closed in Excel first.\n\n"
            "Continue?"
        ),
        "reorder_busy": "Applying\u2026 please wait.",
        "reorder_done": "Charm Library reordered and saved.\nRun \u00abRebuild the shopping list\u00bb to update the route.",
        "reorder_empty": "No charms found in the Charm Library.",
        "reorder_no_catalog": "supplier_catalog.xlsx not found.\nRun \u00abOne-time: add Charm sheets\u00bb first.",
        "pdf_drop_frame": "Order PDFs",
        "pdf_drop_blurb": (
            "Add Etsy order PDFs to the project input folder — same idea as the Charms photo box on the other tab. "
            "You can drop files or whole folders (PDFs are found inside). Then choose «Process new order PDFs» "
            "and the green Run, or click «Run: process new order PDFs» here to run that job immediately "
            "(uses the checkboxes on the right for Chinese / HTML / catalog options)."
        ),
        "pdf_drop_hint_dnd": "Drop PDF files or folders here",
        "pdf_drop_hint_no_dnd": (
            "Drag-and-drop: on Windows, install “windnd” (pip install windnd) and restart this app. "
            "You can always use Browse."
        ),
        "pdf_drop_browse": "Browse for PDFs…",
        "pdf_drop_browse_title": "Choose order PDFs",
        "pdf_drop_run": "Run: process new order PDFs",
        "pdf_drop_open_folder": "Open input folder",
        "pdf_drop_move_backup": "Move PDFs to backup…",
        "pdf_drop_no_valid": "None of those paths contained a .pdf file.",
        "pdf_drop_copied": "Copied {n} PDF file(s) to the input folder.\n",
        "pdf_new_batch_start": "\n--- Process new order PDFs ---\n",
        "pdf_backup_empty": "There are no PDF files in the input folder.",
        "pdf_backup_confirm_title": "Move to backup",
        "pdf_backup_confirm_body": (
            "Move {n} PDF file(s) from the input folder into the backup folder?\n\n"
            "Each file goes under backup\\MMDD\\ (month+day). If the name ends with _MMDD.pdf "
            "and that date is valid, that MMDD is used; otherwise today’s date is used.\n\n"
            "{path}"
        ),
        "pdf_backup_log_start": "\n--- Move order PDFs to backup ---\n",
        "pdf_backup_log_line": "  {src}  →  backup\\{mmdd}\\{dest}\n",
        "pdf_backup_done": "Moved {n} PDF file(s) to backup.\n",
        "pdf_backup_errors": "Could not move {n} file(s) (see log).\n",
        "opts_frame": "Optional fields (usually leave empty)",
        "threshold_hint": "Match strictness (0–100, higher = stricter). Empty = default.",
        "zh_exclude_hint": "Exclude shops from Chinese file (comma-separated). Only if Chinese Excel is checked.",
        "charm_dir_hint": "Custom charm photos folder (full path). Empty = default folder.",
        "run": "Run",
        "output_frame": "Output",
        "log_ready": (
            "Ready. Tab 1: order PDFs + main job + green Run. Tab 2: section A adds charm photos to the catalog; "
            "section B is only for charm steps triggered by Run.\n"
        ),
        "msg_busy_title": "Busy",
        "msg_busy": "Already running — wait for it to finish.",
        "msg_missing_title": "Missing file",
        "msg_missing": "Cannot find:\n",
        "log_start": "\n--- Starting ---\n",
        "log_messages": "\n--- Messages ---\n",
        "log_finished": "\n--- Finished (exit code {code}) ---\n",
        "log_error": "\nError: {e}\n",
    },
    "zh": {
        "win_title": "采购清单生成工具",
        "lang_btn": "English",
        "header_title": "把订单 PDF 转成采购清单和卖场路线图。",
        "tab_orders": "1 — 订单与输出",
        "tab_charms": "2 — 挂饰（可选）",
        "step_strip": (
            "  ①  加入 PDF（本页拖放 / 浏览 /「打开输入文件夹」）   →   ②  左侧选「一项」主要任务   →   "
            "③  下方绿色「运行」（或加完 PDF 后点「运行：处理新订单 PDF」）"
        ),
        "footer_run_hint": "完成后请看下方输出框里的提示信息。",
        "quick_hint": (
            "快速打开 — 文件已存在时用 Excel 或浏览器打开；没有时请先在下方「运行」生成。"
            "上面一组是供应商商品表及其所在文件夹（含挂饰图等）；下面一组是采购清单生成结果（要中文版或网页请先勾选再运行）。"
        ),
        "quick_group_supplier": "供应商 — 商品数据（商品表在此文件夹中）",
        "quick_group_route": "采购清单 — 生成结果",
        "quick_catalog": "打开商品表（Excel）",
        "btn_catalog_backups": "商品表备份…",
        "cat_backup_title": "供应商商品表备份",
        "cat_backup_hint": (
            "在程序修改 supplier_catalog.xlsx 之前会自动保存带时间戳的副本。"
            "恢复前请先关闭 Excel。"
        ),
        "cat_backup_open_folder": "打开备份文件夹",
        "cat_backup_snapshot": "立即快照",
        "cat_backup_snapshot_ok": "快照已保存：\n{path}",
        "cat_backup_snapshot_no_file": "未找到 supplier_catalog.xlsx，无法复制。",
        "cat_backup_restore": "恢复所选备份",
        "cat_backup_none": "（尚无备份——执行任意会写入商品表的操作后即可生成。）",
        "cat_backup_restore_confirm": (
            "用此备份替换当前的 supplier_catalog.xlsx？\n\n{path}\n\n"
            "请先关闭 Excel。当前文件会先复制到备份文件夹后再替换。"
        ),
        "cat_backup_restore_ok": "已从备份恢复商品表。",
        "cat_backup_restore_pick": "请先在列表中选择一项备份。",
        "cat_backup_close": "关闭",
        "btn_edit_products": "编辑商品…",
        "edit_title": "编辑商品表",
        "edit_heading": "分配店铺、摊位与挂饰代码，或标记为停售",
        "edit_intro": (
            "选择一行商品，在右侧面板中分配或更改店铺、摊位、挂饰店和挂饰代码，"
            "或将其标记为停售以将其从后续采购清单中移除。"
        ),
        "edit_search_label": "搜索",
        "edit_search_tip": "匹配标题、店铺或摊位",
        "edit_col_photo": "照片",
        "edit_col_row": "行",
        "edit_col_title": "标题（简短）",
        "edit_col_shop": "店铺",
        "edit_col_stall": "摊位",
        "edit_col_charm_shop": "挂饰店",
        "edit_col_charm_code": "挂饰代码",
        "edit_lbl_shop": "店铺名称",
        "edit_lbl_stall": "摊位",
        "edit_lbl_charm_shop": "挂饰店",
        "edit_lbl_charm_code": "挂饰代码",
        "edit_btn_save": "保存",
        "edit_btn_close": "关闭",
        "edit_saved": "已保存 — {title}",
        "edit_no_selection": "请先在表格中选择一个商品行。",
        "edit_no_import": "商品表辅助功能不可用（请检查安装）。",
        "edit_empty": "Product Map 工作表中没有商品行。",
        "edit_preview_title": "编辑 — 已选商品",
        "edit_preview_placeholder": "点击一行以编辑其字段。",
        "edit_hover_tip": "提示：将鼠标悬停在照片列的缩略图上——光标旁边会弹出更大的照片。",
        "edit_btn_upload_photo": "上传新照片\u2026",
        "edit_upload_photo_title": "选择商品照片",
        "edit_upload_photo_pending": "\u2605 新照片已就绪——点击「保存」以应用",
        "edit_photo_saved": "照片已更新 — {title}",
        "edit_photo_save_fail": "无法保存照片：{err}",
        "edit_charm_pick": "选择…",
        "edit_charm_unknown": "挂饰库中未找到此代码。",
        "charm_picker_title": "选择挂饰",
        "charm_picker_heading": "点击一个挂饰以选择",
        "charm_picker_search": "筛选",
        "charm_picker_search_tip": "匹配代码或 SKU",
        "charm_picker_col_photo": "照片",
        "charm_picker_col_code": "代码",
        "charm_picker_col_sku": "SKU",
        "charm_picker_col_shop": "默认挂饰店",
        "charm_picker_none": "（无——清除挂饰代码）",
        "charm_picker_hover_tip": "提示：将鼠标悬停在照片栏上，旁边会弹出较大的图片。",
        "edit_danger_zone": "危险操作",
        "edit_danger_note": "此操作将商品永久移出活跃商品目录。",
        "edit_btn_discontinue": "标记为停售\u2026",
        "edit_discontinue_confirm_title": "标记为停售",
        "edit_discontinue_confirm": (
            "将此商品移至「Discontinued Products」工作表？\n\n"
            "{title}\n\n"
            "它将不再出现在后续采购清单中。\n"
            "记录将保留以供参考——不会删除。"
        ),
        "edit_discontinue_done": (
            "已移至「Discontinued Products」工作表。\n"
            "请接着在主界面运行「重新生成采购清单」以更新路线。"
        ),
        "discontinue_title": "标记商品为停售",
        "discontinue_heading": "从目录中排除某款商品",
        "discontinue_intro": (
            "供应商不再供货时在此标记。商品行会从「Product Map」移至「Discontinued Products」工作表——"
            "记录仍会保留，但新订单不会再自动匹配到它。"
        ),
        "discontinue_step1": "1. 对照左侧缩略图，并结合「店名 / 摊位」确认是哪一款。",
        "discontinue_step2": "2. 单击一行查看详情，再点「标记所选」（或双击）。商品会被移到「Discontinued Products」工作表。",
        "discontinue_step3": "3. 回到主界面执行「重新生成采购清单」，路线图才会更新。",
        "discontinue_search_label": "搜索",
        "discontinue_search_tip": "匹配标题、店名或摊位",
        "discontinue_hover_tip": "提示：将鼠标停留在左侧「图片」列的缩略图上片刻，光标旁会弹出放大商品图。",
        "discontinue_preview_title": "预览 — 当前选中",
        "discontinue_preview_placeholder": "点击表格中的一行，可查看完整标题和放大图片。",
        "discontinue_col_thumb": "图片",
        "discontinue_col_row": "行",
        "discontinue_col_shop": "店名",
        "discontinue_col_stall": "摊位",
        "discontinue_col_title": "标题（简）",
        "discontinue_apply": "将所选标记为停售",
        "discontinue_done": "商品已移至「Discontinued Products」工作表。请接着在主界面重新生成采购清单。",
        "discontinue_empty": "Product Map 工作表中没有商品行。",
        "discontinue_close": "关闭",
        "discontinue_no_import": "无法加载商品表功能（请检查安装）。",
        "discontinue_no_pillow": (
            "安装 Pillow 后可显示商品缩略图：  pip install Pillow  （装好后重启本程序）"
        ),
        "discontinue_no_selection": "请先在表格中选择一行商品。",
        "quick_data_folder": "打开商品数据文件夹",
        "quick_route": "Excel（默认）",
        "quick_route_zh": "Excel — 中文版",
        "quick_html": "网页版",
        "quick_html_zh": "网页 — 中文版",
        "file_missing_title": "还没有这个文件",
        "file_missing_body": "目前找不到：\n{path}\n请先在上方运行一次，或询问主管。若要中文版或网页，运行前请勾选对应选项。",
        "file_open_fail_title": "无法打开",
        "file_open_fail_body": "{err}",
        "main_frame": "主要任务（选一项）",
        "also_frame": "同时生成",
        "chk_excel_zh": "中文版采购清单（Excel）",
        "chk_html": "手机友好的网页版清单",
        "chk_no_catalog": "本次运行不向供应商表添加新商品行",
        "charm_drop_section": "A — 把新挂饰照片加入商品表",
        "charm_how_body": (
            "具体会做什么\n"
            "• 拖放或浏览会把文件复制到挂饰图片文件夹（见「打开挂饰照片文件夹」），文件名暂时带前缀。\n"
            "• 点「导入到工作簿」后：磁盘上的文件会重命名为下一个可用编码（如 CH-00012.png），并在 supplier_catalog.xlsx 的 Charm Library "
            "工作表新增一行，图片嵌在表里。这里是「每个挂饰一条」的主表。\n"
            "• 不会往 Charm Shops 表加行——那个表只是卖场摊位清单，一般在 setup 时填好即可。\n"
            "• 第一次使用：请主管在 B 区执行一次「一次性：在商品表里添加挂饰工作表」，之后才能正常导入。\n"
            "支持格式：PNG、JPG/JPEG、WebP。"
        ),
        "charm_drop_blurb": "拖放或浏览选图后，点「导入到工作簿」——这是添加新挂饰图的常规做法。",
        "charm_run_section": "B — 可选：挂饰任务（配合绿色「运行」）",
        "charm_run_note": (
            "以下任选一项，或保持「不用」。若选了某项挂饰任务，下一次绿色「运行」将只执行该任务，"
            "不会同时跑第 1 页的主要任务。日常添加新挂饰图请用上方 A 区并点「导入到工作簿」；多数情况不必动 B 区。"
        ),
        "charm_smart": "在「运行」执行挂饰导入时，同时用 AI 填 SKU（C 列）；需主管在本机配置接口",
        "charm_drop_hint_dnd": "将图片文件拖放到此处",
        "charm_drop_hint_no_dnd": "拖放：在 Windows 上请安装 windnd（pip install windnd）并重启本程序。也可始终使用「浏览」按钮。",
        "charm_drop_browse": "浏览照片…",
        "charm_drop_browse_title": "选择挂饰照片",
        "charm_drop_import": "导入到工作簿（重命名并写 Excel）",
        "charm_drop_open_folder": "打开挂饰照片文件夹",
        "charm_drop_vision": (
            "导入时用 AI 建议 SKU（环境变量 CHARM_VISION_* / OPENAI_*；"
            "本机或其他服务可设 CHARM_VISION_BASE_URL，如 Ollama、OpenRouter 等）"
        ),
        "charm_msg_title": "挂饰照片",
        "charm_drop_no_valid": "所选文件里没有支持的图片格式（PNG、JPG/JPEG、WebP）。",
        "charm_drop_nothing_to_import": "没有待导入的暂存照片。请先拖放文件或使用浏览。",
        "charm_drop_staged": "已暂存 {n} 张照片。请点击「导入到工作簿」以重命名并更新 Excel。\n",
        "charm_import_start": "\n--- 挂饰图片导入 ---\n",
        "charm_reorder_section": "C — 重排挂饰顺序",
        "charm_reorder_body": (
            "在面板中上下拖动挂饰行即可更改顺序，无需打开 Excel 文件。"
            "应用后，代码将按位置重新分配为 CH-00001、CH-00002\u2026，"
            "商品表中所有引用也自动更新。"
        ),
        "charm_reorder_btn_open": "打开挂饰排序面板\u2026",
        "charm_reorder_start_preview": "\n--- 重新编号挂饰代码（预览）---\n",
        "charm_reorder_start_apply": "\n--- 重新编号挂饰代码 ---\n",
        "charm_reorder_nothing": "未找到 Charm Library 或为空——无需重排。",
        "charm_reorder_no_import": "商品表辅助功能不可用（请检查安装）。",
        # Reorder dialog strings
        "reorder_title": "挂饰排序",
        "reorder_heading": "拖放调整顺序",
        "reorder_intro": (
            "上下拖动行，或选中行后点击 \u2191 \u2193 按钮。"
            "「新代码」列显示应用后各挂饰将获得的代码。"
            "点击「应用」保存——Excel 文件与商品表均自动更新。"
        ),
        "reorder_col_photo": "照片",
        "reorder_col_code": "当前代码",
        "reorder_col_new": "新代码",
        "reorder_col_sku": "SKU",
        "reorder_col_shop": "默认挂饰店",
        "reorder_btn_up": "\u2191 上移",
        "reorder_btn_down": "\u2193 下移",
        "reorder_btn_top": "\u21d1 移到顶部",
        "reorder_btn_bottom": "\u21d3 移到底部",
        "reorder_btn_apply": "应用排序并重新编号",
        "reorder_btn_close": "关闭",
        "reorder_confirm_title": "应用排序",
        "reorder_confirm_body": (
            "将按新顺序重写 Charm Library 行，\n"
            "按位置重新分配 CH-00001、CH-00002\u2026，\n"
            "并更新所有商品表引用。\n\n"
            "请先确保 supplier_catalog.xlsx 在 Excel 中已关闭。\n\n"
            "继续？"
        ),
        "reorder_busy": "应用中\u2026请稍候。",
        "reorder_done": "Charm Library 已重排并保存。\n请运行「重新生成采购清单」以更新路线。",
        "reorder_empty": "Charm Library 中未找到挂饰。",
        "reorder_no_catalog": "未找到 supplier_catalog.xlsx。\n请先运行「一次性：在商品表里添加挂饰工作表」。",
        "pdf_drop_frame": "订单 PDF",
        "pdf_drop_blurb": (
            "把 Etsy 订单 PDF 放进项目的 input 文件夹——用法与「挂饰」标签页的照片拖放区一致。"
            "可拖文件或整个文件夹（会自动查找其中的 PDF）。然后左侧选「处理新的订单 PDF」再点绿色「运行」，"
            "或在本区点击「运行：处理新订单 PDF」立刻执行该任务（右侧勾选的中文版 / 网页 / 商品表选项会一并生效）。"
        ),
        "pdf_drop_hint_dnd": "将 PDF 文件或文件夹拖放到此处",
        "pdf_drop_hint_no_dnd": (
            "拖放：在 Windows 上请安装 windnd（pip install windnd）并重启本程序。也可始终使用「浏览」按钮。"
        ),
        "pdf_drop_browse": "浏览 PDF…",
        "pdf_drop_browse_title": "选择订单 PDF",
        "pdf_drop_run": "运行：处理新订单 PDF",
        "pdf_drop_open_folder": "打开输入文件夹",
        "pdf_drop_move_backup": "将 PDF 移至备份…",
        "pdf_drop_no_valid": "这些路径里没有找到 .pdf 文件。",
        "pdf_drop_copied": "已将 {n} 个 PDF 复制到 input 文件夹。\n",
        "pdf_new_batch_start": "\n--- 处理新订单 PDF ---\n",
        "pdf_backup_empty": "input 文件夹中没有 PDF 文件。",
        "pdf_backup_confirm_title": "移至备份",
        "pdf_backup_confirm_body": (
            "是否将 input 文件夹中的 {n} 个 PDF 移到 backup 文件夹？\n\n"
            "每个文件会放入 backup\\MMDD\\（月+日，共四位）。若文件名以 _MMDD.pdf 结尾且日期有效，"
            "则使用该 MMDD；否则使用今天的日期。\n\n"
            "{path}"
        ),
        "pdf_backup_log_start": "\n--- 将订单 PDF 移至备份 ---\n",
        "pdf_backup_log_line": "  {src}  →  backup\\{mmdd}\\{dest}\n",
        "pdf_backup_done": "已将 {n} 个 PDF 移到备份。\n",
        "pdf_backup_errors": "有 {n} 个文件未能移动（见日志）。\n",
        "opts_frame": "可选项（通常留空）",
        "threshold_hint": "匹配严格程度（0–100，越大越严）。留空则用默认。",
        "zh_exclude_hint": "中文版要排除的店铺（英文逗号分隔）。仅在选择中文版 Excel 时有效。",
        "charm_dir_hint": "自定义挂饰照片文件夹（完整路径）。留空则用默认目录。",
        "run": "运行",
        "output_frame": "输出",
        "log_ready": (
            "就绪。第 1 页：订单 PDF、主要任务与绿色「运行」。第 2 页：A 区把挂饰照片写入商品表；"
            "B 区仅在使用「运行」做挂饰专项时需要。\n"
        ),
        "msg_busy_title": "请稍候",
        "msg_busy": "正在运行中，请等待完成。",
        "msg_missing_title": "缺少文件",
        "msg_missing": "找不到文件：\n",
        "log_start": "\n--- 开始 ---\n",
        "log_messages": "\n--- 提示信息 ---\n",
        "log_finished": "\n--- 结束（退出码 {code}）---\n",
        "log_error": "\n错误：{e}\n",
    },
}


def flag_for_job(job_id: str) -> str | None:
    return {
        "new_batch": "--new-batch",
        "refresh_catalog": "--refresh-catalog",
        "purge_purchased": "--purge-purchased",
        "rebuild_catalog": "--rebuild-catalog",
        "reset": "--reset",
    }.get(job_id)


def _decode_windnd_paths(files: object) -> list[Path]:
    out: list[Path] = []
    seq = list(files) if files is not None else []
    for f in seq:
        if isinstance(f, str):
            if f:
                out.append(Path(f))
            continue
        if isinstance(f, bytes):
            text: str | None = None
            for enc in ("utf-8", "mbcs"):
                try:
                    text = f.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            if text is None:
                text = f.decode("utf-8", errors="replace")
        else:
            text = str(f)
        if text:
            out.append(Path(text))
    return out


def _expand_pdf_paths(paths: list[Path]) -> list[Path]:
    """Collect .pdf files from dropped files and from directories (recursive)."""
    out: list[Path] = []
    for p in paths:
        try:
            rp = p if isinstance(p, Path) else Path(p)
            rp = rp.expanduser()
        except (TypeError, ValueError):
            continue
        try:
            if rp.is_file():
                if rp.suffix.lower() == ORDER_PDF_EXT:
                    out.append(rp.resolve())
            elif rp.is_dir():
                for child in rp.rglob("*"):
                    if child.is_file() and child.suffix.lower() == ORDER_PDF_EXT:
                        out.append(child.resolve())
        except OSError:
            continue
    return out


def _unique_pdf_dest(input_dir: Path, src: Path) -> Path:
    """Pick input_dir / name that does not overwrite an existing file."""
    input_dir.mkdir(parents=True, exist_ok=True)
    stem = re.sub(r"[^\w\-.]+", "_", src.stem, flags=re.UNICODE).strip("._") or "order"
    stem = stem[:120]
    ext = ORDER_PDF_EXT
    candidate = input_dir / f"{stem}{ext}"
    if not candidate.exists():
        return candidate
    for i in range(2, 10_000):
        alt = input_dir / f"{stem} ({i}){ext}"
        if not alt.exists():
            return alt
    return input_dir / f"{stem}_{secrets.token_hex(4)}{ext}"


def _valid_mmdd(token: str) -> bool:
    """True if token is MMDD for a real calendar day (uses a leap year so Feb 29 is allowed)."""
    if len(token) != 4 or not token.isdigit():
        return False
    month = int(token[:2])
    day = int(token[2:])
    try:
        date(2000, month, day)
    except ValueError:
        return False
    return True


def _mmdd_folder_for_order_pdf(path: Path) -> str:
    """Pick backup subfolder name MMDD: from filename *_MMDD.pdf if valid, else today."""
    m = _ORDER_PDF_MMdd_SUFFIX.search(path.name)
    if m:
        tok = m.group(1)
        if _valid_mmdd(tok):
            return tok
    today = date.today()
    return f"{today.month:02d}{today.day:02d}"


def _unique_path_in_dir(directory: Path, filename: str) -> Path:
    """Pick directory / filename that does not overwrite an existing file."""
    directory.mkdir(parents=True, exist_ok=True)
    candidate = directory / filename
    if not candidate.exists():
        return candidate
    stem = Path(filename).stem
    ext = Path(filename).suffix
    for i in range(2, 10_000):
        alt = directory / f"{stem} ({i}){ext}"
        if not alt.exists():
            return alt
    return directory / f"{stem}_{secrets.token_hex(4)}{ext}"


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self._lang: Lang = "zh"
        self.minsize(UI_MIN_W, UI_MIN_H)
        self._wrap_sync_id: int | None = None

        self._job_var = tk.StringVar(value=JOB_ORDER[0])
        self._chinese = tk.BooleanVar(value=True)
        self._html = tk.BooleanVar(value=True)
        self._no_catalog = tk.BooleanVar(value=False)
        self._charm_mode_var = tk.StringVar(value="skip")
        self._charm_smart_import = tk.BooleanVar(value=False)
        self._charm_drop_vision = tk.BooleanVar(value=False)

        self._log_q: queue.Queue[str] = queue.Queue()
        self._run_busy = False

        # Widgets updated on language change
        self._w_title: ttk.Label | None = None
        self._w_quick_hint: ttk.Label | None = None
        self._w_quick_grp_sup: ttk.Label | None = None
        self._w_quick_grp_route: ttk.Label | None = None
        self._quick_file_btns: list[tuple[ttk.Button, str]] = []
        self._btn_edit_products: ttk.Button | None = None
        self._btn_catalog_backups: ttk.Button | None = None
        self._w_btn_data: ttk.Button | None = None
        self._w_lang: ttk.Button | None = None
        self._main_frame: ttk.LabelFrame | None = None
        self._job_rbs: list[tuple[ttk.Radiobutton, str]] = []
        self._job_blurbs: list[tuple[ttk.Label, str]] = []
        self._also_frame: ttk.LabelFrame | None = None
        self._chk_zh: ttk.Checkbutton | None = None
        self._chk_html: ttk.Checkbutton | None = None
        self._chk_nc: ttk.Checkbutton | None = None
        self._charm_run_frame: ttk.LabelFrame | None = None
        self._w_charm_how: ttk.Label | None = None
        self._w_charm_run_note: ttk.Label | None = None
        self._charm_mode_rbs: list[tuple[ttk.Radiobutton, str]] = []
        self._charm_mode_blurbs: list[tuple[ttk.Label, str]] = []
        self._charm_smart_cb: ttk.Checkbutton | None = None
        self._opts_frame: ttk.LabelFrame | None = None
        self._w_th_lbl: ttk.Label | None = None
        self._w_zx_lbl: ttk.Label | None = None
        self._w_cd_lbl: ttk.Label | None = None
        self._run_btn: tk.Button | None = None
        self._w_run_hint: tk.Label | None = None
        self._log_frame: ttk.LabelFrame | None = None
        self._notebook: ttk.Notebook | None = None
        self._tab_main: ttk.Frame | None = None
        self._tab_charms: ttk.Frame | None = None
        self._main_inner: ttk.Frame | None = None
        self._charm_inner: ttk.Frame | None = None
        self._w_step_strip: tk.Label | None = None
        self._charm_canvas: tk.Canvas | None = None
        self._charm_vsb: ttk.Scrollbar | None = None
        self._main_canvas: tk.Canvas | None = None
        self._main_vsb: ttk.Scrollbar | None = None
        self._charm_drop_frame: ttk.LabelFrame | None = None
        self._w_charm_drop_blurb: ttk.Label | None = None
        self._w_charm_drop_hint: tk.Label | None = None
        self._charm_drop_zone: tk.Frame | None = None
        self._btn_charm_browse: ttk.Button | None = None
        self._btn_charm_import: ttk.Button | None = None
        self._btn_charm_open_folder: ttk.Button | None = None
        self._charm_drop_vision_cb: ttk.Checkbutton | None = None
        self._charm_reorder_frame: ttk.LabelFrame | None = None
        self._w_charm_reorder_body: ttk.Label | None = None
        self._btn_charm_reorder_open: ttk.Button | None = None
        self._pdf_drop_frame: ttk.LabelFrame | None = None
        self._w_pdf_drop_blurb: ttk.Label | None = None
        self._pdf_drop_zone: tk.Frame | None = None
        self._w_pdf_drop_hint: tk.Label | None = None
        self._btn_pdf_browse: ttk.Button | None = None
        self._btn_pdf_run_new: ttk.Button | None = None
        self._btn_pdf_open_folder: ttk.Button | None = None
        self._btn_pdf_move_backup: ttk.Button | None = None

        self._build()
        self._apply_language()
        self._place_initial_window()
        self.after(200, self._drain_log)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _place_initial_window(self) -> None:
        self.update_idletasks()
        w, h = UI_DEFAULT_W, UI_DEFAULT_H
        sw = max(self.winfo_screenwidth(), w + 1)
        sh = max(self.winfo_screenheight(), h + 1)
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2 - 24)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _schedule_wrap_sync(self, event: tk.Event) -> None:
        if event.widget is not self:
            return
        if self._wrap_sync_id is not None:
            try:
                self.after_cancel(self._wrap_sync_id)
            except tk.TclError:
                pass
        self._wrap_sync_id = self.after(120, self._sync_text_wraps)

    def _sync_text_wraps(self) -> None:
        self._wrap_sync_id = None
        self.update_idletasks()
        w = self.winfo_width()
        if w < 80:
            return
        edge = UI_MARGIN_X * 2 + 32
        wrap_full = max(360, w - edge)
        wrap_strip = max(300, w - edge)
        half = max(260, (w - edge) // 2 - 56)
        run_hint = max(260, w - 300)
        charm = max(400, w - edge - 24)
        if self._w_quick_hint:
            self._w_quick_hint.config(wraplength=wrap_full)
        if self._w_quick_grp_sup:
            self._w_quick_grp_sup.config(wraplength=wrap_full)
        if self._w_quick_grp_route:
            self._w_quick_grp_route.config(wraplength=wrap_full)
        if self._w_step_strip:
            self._w_step_strip.config(wraplength=wrap_strip)
        if self._w_run_hint:
            self._w_run_hint.config(wraplength=run_hint)
        for lbl, _jid in self._job_blurbs:
            lbl.config(wraplength=half)
        for lbl, _mid in self._charm_mode_blurbs:
            lbl.config(wraplength=max(260, charm - 24))
        if self._w_charm_how:
            self._w_charm_how.config(wraplength=max(320, charm - 8))
        if self._w_charm_run_note:
            self._w_charm_run_note.config(wraplength=max(320, charm - 8))
        if self._w_charm_reorder_body:
            self._w_charm_reorder_body.config(wraplength=max(320, charm - 8))
        if self._w_charm_drop_blurb:
            self._w_charm_drop_blurb.config(wraplength=max(280, charm - 24))
        if self._w_charm_drop_hint:
            self._w_charm_drop_hint.config(wraplength=max(240, w - 200))
        if self._w_pdf_drop_blurb:
            self._w_pdf_drop_blurb.config(wraplength=max(280, wrap_full - 40))
        if self._w_pdf_drop_hint:
            self._w_pdf_drop_hint.config(wraplength=max(240, w - 200))
        self.after_idle(self._sync_main_scroll)
        self.after_idle(self._sync_charm_scroll)

    def _on_close(self) -> None:
        if self._wrap_sync_id is not None:
            try:
                self.after_cancel(self._wrap_sync_id)
            except tk.TclError:
                pass
        try:
            self.unbind_all("<MouseWheel>")
            self.unbind_all("<Button-4>")
            self.unbind_all("<Button-5>")
        except tk.TclError:
            pass
        self.destroy()

    def _t(self, key: str, **fmt: object) -> str:
        s = CHROME[self._lang][key]
        return s.format(**fmt) if fmt else s

    def _hook_windnd_drop(self, widget: tk.Misc | None, handler: Callable[[list[Path]], None]) -> None:
        """Register *handler* for file/folder drops on *widget* (Windows + windnd only)."""
        if windnd is None or widget is None:
            return

        def _on_drop(files: object, _obj: object | None = None) -> None:
            paths = _decode_windnd_paths(files)
            self.after(0, lambda p=list(paths): handler(p))

        try:
            windnd.hook_dropfiles(widget, _on_drop, force_unicode=True)
        except (tk.TclError, OSError, AttributeError):
            pass

    def _reveal_directory(self, path: Path) -> None:
        """Open a directory in the OS file manager (cross-platform)."""
        path.mkdir(parents=True, exist_ok=True)
        target = str(path.resolve())
        try:
            if sys.platform == "win32":
                os.startfile(target)
            elif sys.platform == "darwin":
                subprocess.run(["open", target], check=False)
            else:
                subprocess.run(["xdg-open", target], check=False)
        except OSError as e:
            messagebox.showerror(self._t("file_open_fail_title"), self._t("file_open_fail_body", err=e))

    def _toggle_language(self) -> None:
        self._lang = "zh" if self._lang == "en" else "en"
        self._apply_language()

    def _apply_language(self) -> None:
        self.title(CHROME[self._lang]["win_title"])
        if self._w_lang:
            # Button label = the language you switch TO
            self._w_lang.config(text=CHROME[self._lang]["lang_btn"])
        if self._w_title:
            self._w_title.config(text=CHROME[self._lang]["header_title"])
        if self._w_quick_hint:
            self._w_quick_hint.config(text=CHROME[self._lang]["quick_hint"])
        if self._w_quick_grp_sup:
            self._w_quick_grp_sup.config(text=CHROME[self._lang]["quick_group_supplier"])
        if self._w_quick_grp_route:
            self._w_quick_grp_route.config(text=CHROME[self._lang]["quick_group_route"])
        for btn, key in self._quick_file_btns:
            btn.config(text=CHROME[self._lang][key])
        if self._btn_edit_products:
            self._btn_edit_products.config(text=CHROME[self._lang]["btn_edit_products"])
        if self._btn_catalog_backups:
            self._btn_catalog_backups.config(text=CHROME[self._lang]["btn_catalog_backups"])
        if self._w_btn_data:
            self._w_btn_data.config(text=CHROME[self._lang]["quick_data_folder"])
        if self._main_frame:
            self._main_frame.config(text=CHROME[self._lang]["main_frame"])
        for rb, job_id in self._job_rbs:
            title, _ = JOB_TEXT[job_id][self._lang]
            rb.config(text=title)
        for lbl, job_id in self._job_blurbs:
            _, blurb = JOB_TEXT[job_id][self._lang]
            lbl.config(text=blurb)
        if self._also_frame:
            self._also_frame.config(text=CHROME[self._lang]["also_frame"])
        if self._chk_zh:
            self._chk_zh.config(text=CHROME[self._lang]["chk_excel_zh"])
        if self._chk_html:
            self._chk_html.config(text=CHROME[self._lang]["chk_html"])
        if self._chk_nc:
            self._chk_nc.config(text=CHROME[self._lang]["chk_no_catalog"])
        if self._charm_drop_frame:
            self._charm_drop_frame.config(text=CHROME[self._lang]["charm_drop_section"])
        if self._w_charm_how:
            self._w_charm_how.config(text=CHROME[self._lang]["charm_how_body"])
        if self._charm_run_frame:
            self._charm_run_frame.config(text=CHROME[self._lang]["charm_run_section"])
        if self._w_charm_run_note:
            self._w_charm_run_note.config(text=CHROME[self._lang]["charm_run_note"])
        for rb, mode_id in self._charm_mode_rbs:
            title, _ = CHARM_MODE_TEXT[mode_id][self._lang]
            rb.config(text=title)
        for lbl, mode_id in self._charm_mode_blurbs:
            _, blurb = CHARM_MODE_TEXT[mode_id][self._lang]
            lbl.config(text=blurb)
        if self._charm_smart_cb:
            self._charm_smart_cb.config(text=CHROME[self._lang]["charm_smart"])
        if self._w_charm_drop_blurb:
            self._w_charm_drop_blurb.config(text=CHROME[self._lang]["charm_drop_blurb"])
        if self._w_charm_drop_hint:
            hk = "charm_drop_hint_dnd" if windnd is not None else "charm_drop_hint_no_dnd"
            self._w_charm_drop_hint.config(text=CHROME[self._lang][hk])
        if self._btn_charm_browse:
            self._btn_charm_browse.config(text=CHROME[self._lang]["charm_drop_browse"])
        if self._btn_charm_import:
            self._btn_charm_import.config(text=CHROME[self._lang]["charm_drop_import"])
        if self._btn_charm_open_folder:
            self._btn_charm_open_folder.config(text=CHROME[self._lang]["charm_drop_open_folder"])
        if self._charm_drop_vision_cb:
            self._charm_drop_vision_cb.config(text=CHROME[self._lang]["charm_drop_vision"])
        if self._charm_reorder_frame:
            self._charm_reorder_frame.config(text=CHROME[self._lang]["charm_reorder_section"])
        if self._w_charm_reorder_body:
            self._w_charm_reorder_body.config(text=CHROME[self._lang]["charm_reorder_body"])
        if self._btn_charm_reorder_open:
            self._btn_charm_reorder_open.config(
                text=CHROME[self._lang]["charm_reorder_btn_open"]
            )
        if self._pdf_drop_frame:
            self._pdf_drop_frame.config(text=CHROME[self._lang]["pdf_drop_frame"])
        if self._w_pdf_drop_blurb:
            self._w_pdf_drop_blurb.config(text=CHROME[self._lang]["pdf_drop_blurb"])
        if self._w_pdf_drop_hint:
            hk = "pdf_drop_hint_dnd" if windnd is not None else "pdf_drop_hint_no_dnd"
            self._w_pdf_drop_hint.config(text=CHROME[self._lang][hk])
        if self._btn_pdf_browse:
            self._btn_pdf_browse.config(text=CHROME[self._lang]["pdf_drop_browse"])
        if self._btn_pdf_run_new:
            self._btn_pdf_run_new.config(text=CHROME[self._lang]["pdf_drop_run"])
        if self._btn_pdf_open_folder:
            self._btn_pdf_open_folder.config(text=CHROME[self._lang]["pdf_drop_open_folder"])
        if self._btn_pdf_move_backup:
            self._btn_pdf_move_backup.config(text=CHROME[self._lang]["pdf_drop_move_backup"])
        if self._opts_frame:
            self._opts_frame.config(text=CHROME[self._lang]["opts_frame"])
        if self._w_th_lbl:
            self._w_th_lbl.config(text=CHROME[self._lang]["threshold_hint"])
        if self._w_zx_lbl:
            self._w_zx_lbl.config(text=CHROME[self._lang]["zh_exclude_hint"])
        if self._w_cd_lbl:
            self._w_cd_lbl.config(text=CHROME[self._lang]["charm_dir_hint"])
        if self._run_btn:
            self._run_btn.config(text=CHROME[self._lang]["run"])
        if self._w_run_hint:
            self._w_run_hint.config(text=CHROME[self._lang]["footer_run_hint"])
        if self._log_frame:
            self._log_frame.config(text=CHROME[self._lang]["output_frame"])
        if self._notebook and self._tab_main and self._tab_charms:
            self._notebook.tab(self._tab_main, text=CHROME[self._lang]["tab_orders"])
            self._notebook.tab(self._tab_charms, text=CHROME[self._lang]["tab_charms"])
        if self._w_step_strip:
            self._w_step_strip.config(text=CHROME[self._lang]["step_strip"])
        self.after_idle(self._sync_charm_scroll)
        self.after_idle(self._sync_main_scroll)
        self.after_idle(self._sync_text_wraps)

    def _sync_main_scroll(self) -> None:
        if self._main_canvas is None:
            return
        self.update_idletasks()
        box = self._main_canvas.bbox("all")
        if box:
            self._main_canvas.configure(scrollregion=box)
        self._maybe_toggle_main_vsb()

    def _maybe_toggle_main_vsb(self) -> None:
        c = self._main_canvas
        v = self._main_vsb
        if c is None or v is None:
            return
        try:
            c.update_idletasks()
            bbox = c.bbox("all")
            if not bbox:
                return
            content_h = bbox[3] - bbox[1]
            view_h = c.winfo_height()
        except tk.TclError:
            return
        if content_h <= view_h + 12:
            v.pack_forget()
        elif not v.winfo_ismapped():
            v.pack(side=tk.RIGHT, fill=tk.Y)

    def _sync_charm_scroll(self) -> None:
        if self._charm_canvas is None:
            return
        self.update_idletasks()
        box = self._charm_canvas.bbox("all")
        if box:
            self._charm_canvas.configure(scrollregion=box)
        self._maybe_toggle_charm_vsb()

    def _maybe_toggle_charm_vsb(self) -> None:
        c = self._charm_canvas
        v = self._charm_vsb
        if c is None or v is None:
            return
        try:
            c.update_idletasks()
            bbox = c.bbox("all")
            if not bbox:
                return
            content_h = bbox[3] - bbox[1]
            view_h = c.winfo_height()
        except tk.TclError:
            return
        if content_h <= view_h + 12:
            v.pack_forget()
        elif not v.winfo_ismapped():
            v.pack(side=tk.RIGHT, fill=tk.Y)

    def _setup_styles(self) -> None:
        s = ttk.Style()
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        C = COLORS
        s.configure(".", background=C["app"])
        s.configure("TFrame", background=C["app"])
        s.configure("App.TFrame", background=C["app"])
        s.configure("Card.TFrame", background=C["card"])
        s.configure("TNotebook", background=C["app"], borderwidth=0)
        s.configure("TNotebook.Tab", padding=[20, 11], font=("Segoe UI", 10, "bold"))
        s.map("TNotebook.Tab",
              background=[("selected", C["card"]), ("!selected", C["app"])],
              foreground=[("selected", C["accent"]), ("!selected", C["muted"])])
        s.configure("Card.TLabelframe", background=C["card"], relief="solid", borderwidth=1,
                    bordercolor=C["border"])
        s.configure("Card.TLabelframe.Label", background=C["card"], foreground=C["accent"],
                    font=("Segoe UI", 11, "bold"))
        s.configure("TLabel", background=C["app"], foreground=C["text"], font=("Segoe UI", 10))
        s.configure("Card.TLabel", background=C["card"], foreground=C["text"], font=("Segoe UI", 10))
        s.configure("Muted.TLabel", background=C["card"], foreground=C["muted"], font=("Segoe UI", 10))
        s.configure("Title.TLabel", background=C["app"], foreground=C["text"], font=("Segoe UI", 16, "bold"))
        s.configure("Sub.TLabel", background=C["app"], foreground=C["muted"], font=("Segoe UI", 10))
        s.configure("TRadiobutton", background=C["card"], foreground=C["text"], font=("Segoe UI", 10))
        s.configure("TCheckbutton", background=C["card"], foreground=C["text"], font=("Segoe UI", 10))
        s.configure("Tool.TButton", font=("Segoe UI", 10), padding=(11, 7))
        s.map("Tool.TButton",
              background=[("active", C["accent_soft"])],
              foreground=[("active", C["accent"])])
        s.configure("Run.TButton", background=C["run"], foreground="#ffffff",
                    font=("Segoe UI", 12, "bold"), padding=(28, 12))
        s.map("Run.TButton", background=[("active", C["run_hover"]), ("disabled", "#94a3b8")])

    def _build(self) -> None:
        self.configure(bg=COLORS["app"])
        self._setup_styles()
        # Root layout must use grid (not pack) so the notebook cannot swallow vertical space
        # reserved for Run + log. On Windows, ttk.Notebook + pack(expand=True) often leaves
        # siblings clipped or painted under the native tab control; grid + explicit row weights fixes it.
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        hero = tk.Frame(self, bg=COLORS["hero"], height=5, highlightthickness=0)
        hero.grid(row=0, column=0, sticky="ew")
        self.grid_rowconfigure(0, minsize=5)

        pad = {"padx": UI_MARGIN_X, "pady": 6}
        # Header + strip only (no expand): keeps the notebook + Run bar from fighting for height inside one frame.
        top_block = ttk.Frame(self, style="App.TFrame")
        top_block.grid(row=1, column=0, sticky="ew")

        header = ttk.Frame(top_block, style="App.TFrame")
        header.pack(fill=tk.X, **pad)

        top_row = ttk.Frame(header, style="App.TFrame")
        top_row.pack(fill=tk.X)
        self._w_lang = ttk.Button(top_row, text="", command=self._toggle_language, width=10, style="Tool.TButton")
        self._w_lang.pack(side=tk.RIGHT)

        self._w_title = ttk.Label(top_row, text="", style="Title.TLabel")
        self._w_title.pack(anchor=tk.W, fill=tk.X, expand=True)

        self._w_quick_hint = ttk.Label(header, text="", style="Sub.TLabel", wraplength=1080)
        self._w_quick_hint.pack(anchor=tk.W, pady=(6, 4))

        quick_rows = ttk.Frame(header, style="App.TFrame")
        quick_rows.pack(anchor=tk.W, fill=tk.X)

        def _mk_open_btn(parent: ttk.Frame, path: Path, chrome_key: str) -> None:
            b = ttk.Button(parent, text="", command=lambda p=path: self._open_file_path(p), style="Tool.TButton")
            b.pack(side=tk.LEFT, padx=(0, 6), pady=(0, 4))
            self._quick_file_btns.append((b, chrome_key))

        grp_sup = ttk.Frame(quick_rows, style="App.TFrame")
        grp_sup.pack(anchor=tk.W, fill=tk.X, pady=(0, 2))
        self._w_quick_grp_sup = ttk.Label(grp_sup, text="", style="Sub.TLabel", wraplength=1060)
        self._w_quick_grp_sup.pack(anchor=tk.W, pady=(0, 4))
        row_sup = ttk.Frame(grp_sup, style="App.TFrame")
        row_sup.pack(anchor=tk.W, fill=tk.X)
        _mk_open_btn(row_sup, FILE_SUPPLIER_CATALOG, "quick_catalog")
        self._btn_edit_products = ttk.Button(
            row_sup, text="", command=self._open_edit_products_dialog, style="Tool.TButton"
        )
        self._btn_edit_products.pack(side=tk.LEFT, padx=(0, 6), pady=(0, 4))
        self._btn_catalog_backups = ttk.Button(
            row_sup, text="", command=self._open_catalog_backups_dialog, style="Tool.TButton"
        )
        self._btn_catalog_backups.pack(side=tk.LEFT, padx=(0, 6), pady=(0, 4))
        self._w_btn_data = ttk.Button(row_sup, text="", command=self._open_data_folder, style="Tool.TButton")
        self._w_btn_data.pack(side=tk.LEFT, padx=(0, 6), pady=(0, 4))

        ttk.Separator(quick_rows, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(6, 8))

        grp_route = ttk.Frame(quick_rows, style="App.TFrame")
        grp_route.pack(anchor=tk.W, fill=tk.X, pady=(0, 0))
        self._w_quick_grp_route = ttk.Label(grp_route, text="", style="Sub.TLabel", wraplength=1060)
        self._w_quick_grp_route.pack(anchor=tk.W, pady=(0, 4))
        row_route_xlsx = ttk.Frame(grp_route, style="App.TFrame")
        row_route_xlsx.pack(anchor=tk.W, fill=tk.X, pady=(0, 4))
        _mk_open_btn(row_route_xlsx, FILE_SHOPPING_ROUTE, "quick_route")
        _mk_open_btn(row_route_xlsx, FILE_SHOPPING_ROUTE_ZH, "quick_route_zh")
        row_route_web = ttk.Frame(grp_route, style="App.TFrame")
        row_route_web.pack(anchor=tk.W, fill=tk.X)
        _mk_open_btn(row_route_web, FILE_SHOPPING_HTML, "quick_html")
        _mk_open_btn(row_route_web, FILE_SHOPPING_HTML_ZH, "quick_html_zh")

        strip = tk.Frame(top_block, bg=COLORS["strip"], highlightthickness=0)
        strip.pack(fill=tk.X, pady=(8, 0))
        strip_bar = tk.Frame(strip, bg=COLORS["strip_accent"], width=5, highlightthickness=0)
        strip_bar.pack(side=tk.LEFT, fill=tk.Y)
        self._w_step_strip = tk.Label(
            strip,
            text="",
            bg=COLORS["strip"],
            fg=COLORS["strip_text"],
            font=("Segoe UI", 10, "bold"),
            wraplength=1060,
            justify=tk.LEFT,
            padx=14,
            pady=11,
        )
        self._w_step_strip.pack(side=tk.LEFT, anchor=tk.W, fill=tk.X, expand=True)

        self._notebook = ttk.Notebook(self)
        self._notebook.grid(row=2, column=0, sticky="nsew", padx=(UI_MARGIN_X, UI_MARGIN_X), pady=(10, 0))

        self._tab_main = ttk.Frame(self._notebook, style="Card.TFrame")
        self._tab_charms = ttk.Frame(self._notebook, style="Card.TFrame")
        self._notebook.add(self._tab_main, text=" ")
        self._notebook.add(self._tab_charms, text=" ")
        self._notebook.select(0)

        # Notebook panes clip children; tall content needs an inner canvas (same idea as Charms tab).
        main_wrap = ttk.Frame(self._tab_main, style="Card.TFrame")
        main_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self._main_canvas = tk.Canvas(main_wrap, bg=COLORS["card"], highlightthickness=0, borderwidth=0)
        self._main_vsb = ttk.Scrollbar(main_wrap, orient=tk.VERTICAL, command=self._main_canvas.yview)
        self._main_canvas.configure(yscrollcommand=self._main_vsb.set)
        main_inner = ttk.Frame(self._main_canvas, style="Card.TFrame")
        self._main_inner = main_inner
        main_win = self._main_canvas.create_window((0, 0), window=main_inner, anchor=tk.NW)

        def _main_cw(event: tk.Event) -> None:
            self._main_canvas.itemconfigure(main_win, width=event.width)
            self.after_idle(self._maybe_toggle_main_vsb)

        def _main_configure(_: tk.Event | None = None) -> None:
            self._main_canvas.configure(scrollregion=self._main_canvas.bbox("all"))
            self.after_idle(self._maybe_toggle_main_vsb)

        self._main_canvas.bind("<Configure>", _main_cw)
        main_inner.bind("<Configure>", _main_configure)
        self._main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._main_vsb.pack(side=tk.RIGHT, fill=tk.Y)

        wrap_main = max(360, 1020)
        self._pdf_drop_frame = ttk.LabelFrame(main_inner, text="", style="Card.TLabelframe", padding=14)
        self._pdf_drop_frame.pack(fill=tk.X, padx=10, pady=(10, 10))
        self._w_pdf_drop_blurb = ttk.Label(
            self._pdf_drop_frame,
            text="",
            style="Muted.TLabel",
            wraplength=wrap_main - 24,
            justify=tk.LEFT,
        )
        self._w_pdf_drop_blurb.pack(anchor=tk.W)
        self._pdf_drop_zone = tk.Frame(
            self._pdf_drop_frame,
            bg=COLORS["drop_zone"],
            highlightthickness=2,
            highlightbackground=COLORS["drop_border"],
            height=DROP_ZONE_H,
            highlightcolor=COLORS["accent"],
        )
        self._pdf_drop_zone.pack(fill=tk.X, pady=(8, 10))
        self._pdf_drop_zone.pack_propagate(False)
        self._w_pdf_drop_hint = tk.Label(
            self._pdf_drop_zone,
            text="",
            bg=COLORS["drop_zone"],
            fg=COLORS["drop_hint"],
            font=("Segoe UI", 10),
            wraplength=760,
            justify=tk.CENTER,
        )
        self._w_pdf_drop_hint.pack(expand=True)
        pdf_btns = ttk.Frame(self._pdf_drop_frame, style="Card.TFrame")
        pdf_btns.pack(fill=tk.X)
        self._btn_pdf_browse = ttk.Button(
            pdf_btns,
            text="",
            command=self._pdf_browse,
            style="Tool.TButton",
        )
        self._btn_pdf_browse.pack(side=tk.LEFT, padx=(0, 8))
        self._btn_pdf_run_new = ttk.Button(
            pdf_btns,
            text="",
            command=self._pdf_run_new_batch,
            style="Tool.TButton",
        )
        self._btn_pdf_run_new.pack(side=tk.LEFT, padx=(0, 8))
        self._btn_pdf_open_folder = ttk.Button(
            pdf_btns,
            text="",
            command=self._open_input,
            style="Tool.TButton",
        )
        self._btn_pdf_open_folder.pack(side=tk.LEFT, padx=(0, 8))
        self._btn_pdf_move_backup = ttk.Button(
            pdf_btns,
            text="",
            command=self._pdf_move_to_backup,
            style="Tool.TButton",
        )
        self._btn_pdf_move_backup.pack(side=tk.LEFT, padx=(0, 8))

        self._hook_windnd_drop(self._pdf_drop_zone, self._pdf_on_paths_dropped)

        cols = ttk.Frame(main_inner, style="Card.TFrame")
        cols.pack(fill=tk.X, padx=10, pady=(0, 10))
        left_col = ttk.Frame(cols, style="Card.TFrame")
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))
        right_col = ttk.Frame(cols, style="Card.TFrame")
        right_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0))

        self._main_frame = ttk.LabelFrame(left_col, text="", style="Card.TLabelframe", padding=14)
        self._main_frame.pack(fill=tk.X, pady=(0, 10))

        n_jobs = len(JOB_ORDER)
        for i, job_id in enumerate(JOB_ORDER):
            title, blurb = JOB_TEXT[job_id]["en"]
            rb = ttk.Radiobutton(self._main_frame, text=title, value=job_id, variable=self._job_var)
            rb.pack(anchor=tk.W, pady=(0, 2))
            self._job_rbs.append((rb, job_id))
            lb = ttk.Label(self._main_frame, text=blurb, style="Muted.TLabel", wraplength=480)
            lb.pack(anchor=tk.W, padx=(22, 0), pady=(0, 10 if i < n_jobs - 1 else 0))
            self._job_blurbs.append((lb, job_id))

        self._also_frame = ttk.LabelFrame(left_col, text="", style="Card.TLabelframe", padding=14)
        self._also_frame.pack(fill=tk.X, pady=(0, 0))
        self._chk_zh = ttk.Checkbutton(self._also_frame, text="", variable=self._chinese)
        self._chk_zh.pack(anchor=tk.W)
        self._chk_html = ttk.Checkbutton(self._also_frame, text="", variable=self._html)
        self._chk_html.pack(anchor=tk.W)
        self._chk_nc = ttk.Checkbutton(self._also_frame, text="", variable=self._no_catalog)
        self._chk_nc.pack(anchor=tk.W)

        self._opts_frame = ttk.LabelFrame(right_col, text="", style="Card.TLabelframe", padding=14)
        self._opts_frame.pack(fill=tk.X)
        self._w_th_lbl = ttk.Label(self._opts_frame, text="", style="Card.TLabel")
        self._w_th_lbl.pack(anchor=tk.W)
        self._threshold = ttk.Entry(self._opts_frame, width=14)
        self._threshold.pack(anchor=tk.W, pady=(2, 8))
        self._w_zx_lbl = ttk.Label(self._opts_frame, text="", style="Card.TLabel")
        self._w_zx_lbl.pack(anchor=tk.W)
        self._zh_exclude = ttk.Entry(self._opts_frame, width=52)
        self._zh_exclude.pack(anchor=tk.W, pady=(2, 8))
        self._w_cd_lbl = ttk.Label(self._opts_frame, text="", style="Card.TLabel")
        self._w_cd_lbl.pack(anchor=tk.W)
        self._charm_dir = ttk.Entry(self._opts_frame, width=52)
        self._charm_dir.pack(anchor=tk.W, pady=(2, 0))

        self.after(120, _main_configure)

        ch_wrap = ttk.Frame(self._tab_charms, style="Card.TFrame")
        ch_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self._charm_canvas = tk.Canvas(ch_wrap, bg=COLORS["card"], highlightthickness=0, borderwidth=0)
        self._charm_vsb = ttk.Scrollbar(ch_wrap, orient=tk.VERTICAL, command=self._charm_canvas.yview)
        self._charm_canvas.configure(yscrollcommand=self._charm_vsb.set)
        ch_inner = ttk.Frame(self._charm_canvas, style="Card.TFrame")
        self._charm_inner = ch_inner
        ch_win = self._charm_canvas.create_window((0, 0), window=ch_inner, anchor=tk.NW)

        def _charm_cw(event: tk.Event) -> None:
            self._charm_canvas.itemconfigure(ch_win, width=event.width)
            self.after_idle(self._maybe_toggle_charm_vsb)

        def _charm_configure(_: tk.Event | None = None) -> None:
            self._charm_canvas.configure(scrollregion=self._charm_canvas.bbox("all"))
            self.after_idle(self._maybe_toggle_charm_vsb)

        self._charm_canvas.bind("<Configure>", _charm_cw)
        ch_inner.bind("<Configure>", _charm_configure)
        self._charm_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._charm_vsb.pack(side=tk.RIGHT, fill=tk.Y)

        def _pointer_in_canvas(c: tk.Canvas | None, event: tk.Event) -> bool:
            if c is None:
                return False
            try:
                rx, ry = c.winfo_rootx(), c.winfo_rooty()
                rw, rh = c.winfo_width(), c.winfo_height()
            except tk.TclError:
                return False
            return rw > 2 and rh > 2 and rx <= event.x_root < rx + rw and ry <= event.y_root < ry + rh

        def _canvas_for_wheel_widget(widget: tk.Misc | None) -> tk.Canvas | None:
            """Map the widget under the pointer to the notebook tab canvas that should scroll."""
            if widget is None:
                return None
            mi, ci = self._main_inner, self._charm_inner
            w: tk.Misc | None = widget
            for _ in range(64):
                if w is None or w is self:
                    break
                if ci is not None and w is ci:
                    return self._charm_canvas
                if mi is not None and w is mi:
                    return self._main_canvas
                try:
                    parent = w.winfo_parent()
                    w = self.nametowidget(parent) if parent else None
                except tk.TclError:
                    break
            return None

        def _canvas_for_wheel_event(event: tk.Event) -> tk.Canvas | None:
            c = _canvas_for_wheel_widget(event.widget)
            if c is not None:
                return c
            try:
                sel = self._notebook.select() if self._notebook else ""
            except tk.TclError:
                sel = ""
            if sel == str(self._tab_charms):
                order = (self._charm_canvas, self._main_canvas)
            elif sel == str(self._tab_main):
                order = (self._main_canvas, self._charm_canvas)
            else:
                order = (self._main_canvas, self._charm_canvas)
            for cand in order:
                if _pointer_in_canvas(cand, event):
                    return cand
            return None

        def _on_mousewheel(event: tk.Event) -> str | None:
            delta = int(-1 * (event.delta / 120)) if getattr(event, "delta", 0) else 0
            if not delta:
                return None
            target = _canvas_for_wheel_event(event)
            if target is None:
                return None
            target.yview_scroll(delta, "units")
            return "break"

        def _on_linux_scroll(event: tk.Event) -> str | None:
            delta = -1 if event.num == 4 else 1
            target = _canvas_for_wheel_event(event)
            if target is None:
                return None
            target.yview_scroll(delta, "units")
            return "break"

        self.bind_all("<MouseWheel>", _on_mousewheel)
        self.bind_all("<Button-4>", _on_linux_scroll)
        self.bind_all("<Button-5>", _on_linux_scroll)

        def _nb_refresh_scroll(_: tk.Event) -> None:
            self.after_idle(self._sync_main_scroll)
            self.after_idle(self._sync_charm_scroll)

        self._notebook.bind("<<NotebookTabChanged>>", _nb_refresh_scroll)

        wrap_ch = 1020
        self._charm_drop_frame = ttk.LabelFrame(ch_inner, text="", style="Card.TLabelframe", padding=14)
        self._charm_drop_frame.pack(fill=tk.X)
        self._w_charm_how = ttk.Label(
            self._charm_drop_frame,
            text="",
            style="Muted.TLabel",
            wraplength=wrap_ch - 16,
            justify=tk.LEFT,
        )
        self._w_charm_how.pack(anchor=tk.W, pady=(0, 10))
        self._w_charm_drop_blurb = ttk.Label(
            self._charm_drop_frame,
            text="",
            style="Muted.TLabel",
            wraplength=wrap_ch - 24,
            justify=tk.LEFT,
        )
        self._w_charm_drop_blurb.pack(anchor=tk.W)

        self._charm_drop_zone = tk.Frame(
            self._charm_drop_frame,
            bg=COLORS["drop_zone"],
            highlightthickness=2,
            highlightbackground=COLORS["drop_border"],
            height=DROP_ZONE_H,
            highlightcolor=COLORS["accent"],
        )
        self._charm_drop_zone.pack(fill=tk.X, pady=(8, 10))
        self._charm_drop_zone.pack_propagate(False)
        self._w_charm_drop_hint = tk.Label(
            self._charm_drop_zone,
            text="",
            bg=COLORS["drop_zone"],
            fg=COLORS["drop_hint"],
            font=("Segoe UI", 10),
            wraplength=760,
            justify=tk.CENTER,
        )
        self._w_charm_drop_hint.pack(expand=True)

        drop_btns = ttk.Frame(self._charm_drop_frame, style="Card.TFrame")
        drop_btns.pack(fill=tk.X)
        self._btn_charm_browse = ttk.Button(
            drop_btns,
            text="",
            command=self._charm_browse_photos,
            style="Tool.TButton",
        )
        self._btn_charm_browse.pack(side=tk.LEFT, padx=(0, 8))
        self._btn_charm_import = ttk.Button(
            drop_btns,
            text="",
            command=self._charm_import_workbook,
            style="Tool.TButton",
        )
        self._btn_charm_import.pack(side=tk.LEFT, padx=(0, 8))
        self._btn_charm_open_folder = ttk.Button(
            drop_btns,
            text="",
            command=self._open_charm_images_folder,
            style="Tool.TButton",
        )
        self._btn_charm_open_folder.pack(side=tk.LEFT, padx=(0, 8))

        self._charm_drop_vision_cb = ttk.Checkbutton(
            self._charm_drop_frame,
            text="",
            variable=self._charm_drop_vision,
        )
        self._charm_drop_vision_cb.pack(anchor=tk.W, pady=(10, 0))

        self._hook_windnd_drop(self._charm_drop_zone, self._charm_on_paths_dropped)

        self._charm_run_frame = ttk.LabelFrame(ch_inner, text="", style="Card.TLabelframe", padding=14)
        self._charm_run_frame.pack(fill=tk.X, pady=(14, 0))
        self._w_charm_run_note = ttk.Label(
            self._charm_run_frame,
            text="",
            style="Muted.TLabel",
            wraplength=wrap_ch - 16,
            justify=tk.LEFT,
        )
        self._w_charm_run_note.pack(anchor=tk.W, pady=(0, 12))

        for mode_id in CHARM_MODE_ORDER:
            title, blurb = CHARM_MODE_TEXT[mode_id]["en"]
            rb = ttk.Radiobutton(
                self._charm_run_frame,
                text=title,
                value=mode_id,
                variable=self._charm_mode_var,
            )
            rb.pack(anchor=tk.W)
            self._charm_mode_rbs.append((rb, mode_id))
            lb = ttk.Label(self._charm_run_frame, text=blurb, style="Muted.TLabel", wraplength=wrap_ch - 24)
            lb.pack(anchor=tk.W, padx=(22, 0), pady=(0, 4))
            self._charm_mode_blurbs.append((lb, mode_id))

        self._charm_smart_cb = ttk.Checkbutton(
            self._charm_run_frame,
            text="",
            variable=self._charm_smart_import,
        )
        self._charm_smart_cb.pack(anchor=tk.W, pady=(8, 0))

        def _sync_charm_smart_state(*_: object) -> None:
            m = self._charm_mode_var.get()
            ok = m in ("try_add_photos", "add_photos")
            if ok:
                self._charm_smart_cb.state(["!disabled"])
            else:
                self._charm_smart_cb.state(["disabled"])
                self._charm_smart_import.set(False)

        self._charm_mode_var.trace_add("write", lambda *_: _sync_charm_smart_state())
        _sync_charm_smart_state()

        # ---- Section C: Reorder / Renumber charm codes ----
        self._charm_reorder_frame = ttk.LabelFrame(
            ch_inner, text="", style="Card.TLabelframe", padding=14
        )
        self._charm_reorder_frame.pack(fill=tk.X, pady=(14, 0))
        self._w_charm_reorder_body = ttk.Label(
            self._charm_reorder_frame,
            text="",
            style="Muted.TLabel",
            wraplength=wrap_ch - 16,
            justify=tk.LEFT,
        )
        self._w_charm_reorder_body.pack(anchor=tk.W, pady=(0, 10))
        reorder_btns = ttk.Frame(self._charm_reorder_frame, style="Card.TFrame")
        reorder_btns.pack(fill=tk.X)
        self._btn_charm_reorder_open = ttk.Button(
            reorder_btns,
            text="",
            command=self._open_charm_reorder_dialog,
            style="Tool.TButton",
        )
        self._btn_charm_reorder_open.pack(side=tk.LEFT, padx=(0, 8))

        self.after(120, _charm_configure)

        foot_sep = tk.Frame(self, bg=COLORS["separator"], height=1, highlightthickness=0)
        foot_sep.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        foot = tk.Frame(self, bg=COLORS["accent_soft"], highlightthickness=0)
        foot.grid(row=4, column=0, sticky="ew")
        fin = tk.Frame(foot, bg=COLORS["accent_soft"])
        fin.pack(fill=tk.X, padx=UI_MARGIN_X, pady=14)
        self._run_btn = tk.Button(
            fin,
            text="",
            command=self._run,
            bg=COLORS["run"],
            fg="#ffffff",
            activebackground=COLORS["run_hover"],
            activeforeground="#ffffff",
            font=("Segoe UI", 12, "bold"),
            padx=28,
            pady=10,
            cursor="hand2",
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=0,
        )
        self._run_btn.pack(side=tk.LEFT)
        self._w_run_hint = tk.Label(
            fin,
            text="",
            bg=COLORS["accent_soft"],
            fg=COLORS["muted"],
            font=("Segoe UI", 10),
            wraplength=760,
            justify=tk.LEFT,
        )
        self._w_run_hint.pack(side=tk.LEFT, padx=(18, 0), anchor=tk.NW)

        self._log_frame = ttk.LabelFrame(self, text="", style="Card.TLabelframe", padding=10)
        self._log_frame.grid(row=5, column=0, sticky="ew", padx=UI_MARGIN_X, pady=(4, 14))
        self._log = scrolledtext.ScrolledText(
            self._log_frame,
            height=UI_LOG_LINES,
            wrap=tk.WORD,
            font=("Consolas", 10),
            bg=COLORS["log_bg"],
            fg=COLORS["text"],
            insertbackground=COLORS["text"],
            selectbackground=COLORS["accent_soft"],
            selectforeground=COLORS["text"],
            relief=tk.FLAT,
            borderwidth=0,
        )
        self._log.pack(fill=tk.BOTH, expand=True)

        self._append_log(CHROME[self._lang]["log_ready"])

        def _stack_run_and_log_above_notebook() -> None:
            # Native ttk.Notebook can paint above later siblings; keep Run + output on top.
            try:
                foot_sep.lift(self._notebook)
                foot.lift(foot_sep)
                self._log_frame.lift(foot)
            except tk.TclError:
                pass

        self.after_idle(_stack_run_and_log_above_notebook)

        self.bind("<Configure>", self._schedule_wrap_sync, add="+")

    def _append_log(self, text: str) -> None:
        self._log.insert(tk.END, text)
        self._log.see(tk.END)

    def _drain_log(self) -> None:
        try:
            while True:
                self._append_log(self._log_q.get_nowait())
        except queue.Empty:
            pass
        self.after(200, self._drain_log)

    def _open_input(self) -> None:
        self._reveal_directory(DEFAULT_ORDER_INPUT_DIR)

    def _list_input_order_pdfs(self) -> list[Path]:
        """Top-level .pdf files in the order input folder (non-recursive)."""
        d = DEFAULT_ORDER_INPUT_DIR
        if not d.is_dir():
            return []
        out: list[Path] = []
        try:
            for child in d.iterdir():
                if child.is_file() and child.suffix.lower() == ORDER_PDF_EXT:
                    out.append(child)
        except OSError:
            return []
        out.sort(key=lambda p: p.name.lower())
        return out

    def _pdf_move_to_backup(self) -> None:
        if self._run_busy:
            messagebox.showinfo(self._t("msg_busy_title"), self._t("msg_busy"))
            return
        pdfs = self._list_input_order_pdfs()
        if not pdfs:
            messagebox.showinfo(self._t("pdf_drop_frame"), self._t("pdf_backup_empty"))
            return
        in_path = str(DEFAULT_ORDER_INPUT_DIR.resolve())
        if not messagebox.askyesno(
            self._t("pdf_backup_confirm_title"),
            self._t("pdf_backup_confirm_body", n=len(pdfs), path=in_path),
            parent=self,
        ):
            return
        self._append_log(self._t("pdf_backup_log_start"))
        ok = 0
        err = 0
        for src in pdfs:
            mmdd = _mmdd_folder_for_order_pdf(src)
            dest_dir = DEFAULT_BACKUP_DIR / mmdd
            try:
                dest = _unique_path_in_dir(dest_dir, src.name)
                shutil.move(str(src), str(dest))
                ok += 1
                self._append_log(
                    self._t("pdf_backup_log_line", src=src.name, mmdd=mmdd, dest=dest.name),
                )
            except OSError:
                err += 1
                self._append_log(f"  ! failed: {src.name}\n")
        if ok:
            self._append_log(self._t("pdf_backup_done", n=ok))
        if err:
            self._append_log(self._t("pdf_backup_errors", n=err))

    def _collect_generator_args(self, job_id: str, *, include_charm_steps: bool = True) -> list[str]:
        args: list[str] = ["--project-dir", str(PROJECT_ROOT)]
        main_flag = flag_for_job(job_id)
        if main_flag:
            args.append(main_flag)
        if self._no_catalog.get():
            args.append("--no-catalog-update")
        if self._chinese.get():
            args.append("--chinese")
        if self._html.get():
            args.append("--html")
        th = self._threshold.get().strip()
        if th:
            args.extend(["--threshold", th])
        zx = self._zh_exclude.get().strip()
        if zx:
            args.extend(["--chinese-exclude-shops", zx])
        cd = self._charm_dir.get().strip()
        if cd:
            args.extend(["--charm-images-dir", cd])
        if include_charm_steps:
            mode = self._charm_mode_var.get()
            args.extend(CHARM_MODE_FLAGS.get(mode, []))
            if mode in ("try_add_photos", "add_photos"):
                args.extend(charm_import_pattern_argv())
            if self._charm_smart_import.get() and mode in ("try_add_photos", "add_photos"):
                args.append("--import-charm-vision-sku")
        return args

    def _spawn_generator(self, gen_args: list[str], *, log_intro: str) -> None:
        self._run_busy = True
        self._set_chrome_busy(True)
        self._append_log(log_intro)
        cmd = [sys.executable, str(GENERATOR), *gen_args]
        self._append_log(" ".join(cmd) + "\n\n")

        def work() -> None:
            try:
                proc = subprocess.run(
                    cmd,
                    cwd=str(PROJECT_ROOT),
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                )
                out = proc.stdout or ""
                err = proc.stderr or ""
                self._log_q.put(out)
                if err:
                    self._log_q.put(self._t("log_messages") + err)
                self._log_q.put(self._t("log_finished", code=proc.returncode))
            except Exception as e:
                self._log_q.put(self._t("log_error", e=e))
            finally:

                def done() -> None:
                    self._run_busy = False
                    self._set_chrome_busy(False)

                self.after(0, done)

        threading.Thread(target=work, daemon=True).start()

    def _copy_pdfs_to_input(self, paths: list[Path]) -> int:
        input_dir = DEFAULT_ORDER_INPUT_DIR
        pdfs = _expand_pdf_paths(paths)
        n = 0
        for src in pdfs:
            try:
                dest = _unique_pdf_dest(input_dir, src)
                shutil.copy2(src, dest)
                n += 1
            except OSError:
                continue
        return n

    def _pdf_on_paths_dropped(self, paths: list[Path]) -> None:
        if self._run_busy:
            messagebox.showinfo(self._t("msg_busy_title"), self._t("msg_busy"))
            return
        if not paths:
            return
        n = self._copy_pdfs_to_input(paths)
        if n == 0:
            messagebox.showinfo(self._t("pdf_drop_frame"), self._t("pdf_drop_no_valid"))
            return
        self._append_log(self._t("pdf_drop_copied", n=n))

    def _pdf_browse(self) -> None:
        if self._run_busy:
            messagebox.showinfo(self._t("msg_busy_title"), self._t("msg_busy"))
            return
        picked = filedialog.askopenfilenames(
            parent=self,
            filetypes=[("PDF", "*.pdf"), ("All", "*.*")],
            title=self._t("pdf_drop_browse_title"),
        )
        if not picked:
            return
        self._pdf_on_paths_dropped([Path(s) for s in picked])

    def _pdf_run_new_batch(self) -> None:
        if self._run_busy:
            messagebox.showinfo(self._t("msg_busy_title"), self._t("msg_busy"))
            return
        if not GENERATOR.is_file():
            messagebox.showerror(self._t("msg_missing_title"), self._t("msg_missing") + str(GENERATOR))
            return
        self._spawn_generator(
            self._collect_generator_args("new_batch", include_charm_steps=False),
            log_intro=self._t("pdf_new_batch_start"),
        )

    def _open_edit_products_dialog(self) -> None:
        if update_product_map_cells is None or list_product_map_rows_for_picker is None:
            messagebox.showerror(
                self._t("msg_missing_title"), self._t("edit_no_import"),
            )
            return
        if not FILE_SUPPLIER_CATALOG.exists():
            messagebox.showinfo(
                self._t("msg_missing_title"),
                self._t("msg_missing") + str(FILE_SUPPLIER_CATALOG),
            )
            return
        try:
            all_rows = list_product_map_rows_for_picker(FILE_SUPPLIER_CATALOG)
        except OSError as e:
            messagebox.showerror(self._t("file_open_fail_title"), str(e))
            return
        if not all_rows:
            messagebox.showinfo(self._t("edit_title"), self._t("edit_empty"))
            return

        row_photos: dict[int, bytes] = {}
        try:
            if extract_photos_from_xlsx is not None:
                row_photos = extract_photos_from_xlsx(
                    FILE_SUPPLIER_CATALOG, sheet_name=CATALOG_SHEET, photo_col_idx=0,
                )
        except Exception:
            pass

        supplier_shops: list[str] = []
        supplier_stalls: list[str] = []
        try:
            import openpyxl as _xl
            _wb = _xl.load_workbook(FILE_SUPPLIER_CATALOG, read_only=True, data_only=True)
            if "Suppliers" in _wb.sheetnames:
                _ws_s = _wb["Suppliers"]
                _shop_ci = _stall_ci = None
                for ci in range(1, 20):
                    h = str(_ws_s.cell(1, ci).value or "").strip().lower()
                    if h == "shop name":
                        _shop_ci = ci
                    elif h == "stall":
                        _stall_ci = ci
                for r in _ws_s.iter_rows(min_row=2, values_only=False):
                    if _shop_ci is not None:
                        v = str(r[_shop_ci - 1].value or "").strip()
                        if v and v not in supplier_shops:
                            supplier_shops.append(v)
                    if _stall_ci is not None:
                        v = str(r[_stall_ci - 1].value or "").strip()
                        if v and v not in supplier_stalls:
                            supplier_stalls.append(v)
            _wb.close()
        except Exception:
            pass

        charm_shop_names: list[str] = []
        charm_entries: dict = {}
        try:
            if load_charm_shops is not None:
                for cs in load_charm_shops(FILE_SUPPLIER_CATALOG):
                    if cs.shop_name and cs.shop_name not in charm_shop_names:
                        charm_shop_names.append(cs.shop_name)
            if load_charm_library is not None:
                charm_entries = load_charm_library(FILE_SUPPLIER_CATALOG)
        except Exception:
            pass

        _ProductMapEditorDialog(
            self,
            all_rows,
            row_photos,
            supplier_shops=supplier_shops,
            supplier_stalls=supplier_stalls,
            charm_shop_names=charm_shop_names,
            charm_entries=charm_entries,
        )

    def _open_discontinue_dialog(self) -> None:
        if (
            list_product_map_rows_for_picker is None
            or mark_product_map_discontinued_by_row is None
            or extract_photos_from_xlsx is None
        ):
            messagebox.showerror(self._t("msg_missing_title"), self._t("discontinue_no_import"))
            return
        if not FILE_SUPPLIER_CATALOG.exists():
            messagebox.showinfo(self._t("msg_missing_title"), self._t("msg_missing") + str(FILE_SUPPLIER_CATALOG))
            return
        try:
            all_rows = list_product_map_rows_for_picker(FILE_SUPPLIER_CATALOG)
        except OSError as e:
            messagebox.showerror(self._t("file_open_fail_title"), str(e))
            return
        if not all_rows:
            messagebox.showinfo(self._t("msg_missing_title"), self._t("discontinue_empty"))
            return

        pil_ok = Image is not None and ImageTk is not None
        try:
            row_photos: dict[int, bytes] = extract_photos_from_xlsx(
                FILE_SUPPLIER_CATALOG,
                sheet_name=CATALOG_SHEET,
                photo_col_idx=0,
            )
        except Exception:
            row_photos = {}

        d = tk.Toplevel(self)
        d.title(self._t("discontinue_title"))
        d.transient(self)
        d.grab_set()
        d.configure(bg=COLORS["app"])
        d.minsize(1020, 620)
        d.geometry("1320x820")

        tk_img_refs: list[object] = []
        row_by_iid: dict[str, ProductMapPickerRow] = {}
        preview_photo_ref: list[object] = []

        def bytes_to_thumb(b: bytes | None, px: int) -> object | None:
            if not pil_ok or Image is None or ImageTk is None:
                return None
            if not b:
                im = Image.new("RGB", (px, px), (226, 232, 240))
            else:
                try:
                    im = Image.open(BytesIO(b))
                except Exception:
                    im = Image.new("RGB", (px, px), (226, 232, 240))
                else:
                    if im.mode not in ("RGB", "RGBA"):
                        im = im.convert("RGBA") if "A" in im.mode else im.convert("RGB")
                    if im.mode == "RGBA":
                        bg = Image.new("RGB", im.size, (255, 255, 255))
                        bg.paste(im, mask=im.split()[3])
                        im = bg
                    im.thumbnail((px, px), Image.Resampling.LANCZOS)
            ph = ImageTk.PhotoImage(im)
            tk_img_refs.append(ph)
            return ph

        # ---- root layout: grid (hero, search, hover hint, body, buttons) ----
        d.grid_columnconfigure(0, weight=1)
        d.grid_rowconfigure(3, weight=1)  # main content row expands

        # Row 0 — hero banner (fixed height)
        hero = tk.Frame(d, bg=COLORS["hero"], highlightthickness=0)
        hero.grid(row=0, column=0, sticky="ew")
        tk.Label(
            hero, text=self._t("discontinue_heading"),
            font=("Segoe UI", 12, "bold"), fg="#ffffff", bg=COLORS["hero"],
        ).pack(anchor=tk.W, padx=14, pady=(10, 4))
        for key in ("discontinue_intro", "discontinue_step1", "discontinue_step2", "discontinue_step3"):
            tk.Label(
                hero, text=self._t(key), wraplength=1240, justify=tk.LEFT,
                font=("Segoe UI", 9), fg="#e2e8f0", bg=COLORS["hero"],
            ).pack(anchor=tk.W, padx=14, pady=(0, 2))
        tk.Label(hero, text="", bg=COLORS["hero"], height=0).pack(pady=(0, 6))

        # Row 1 — search bar (fixed height)
        search_frame = tk.Frame(d, bg=COLORS["app"])
        search_frame.grid(row=1, column=0, sticky="ew", padx=14, pady=(8, 4))
        tk.Label(
            search_frame, text=self._t("discontinue_search_label"),
            font=("Segoe UI", 10, "bold"), bg=COLORS["app"], fg=COLORS["text"],
        ).pack(side=tk.LEFT)
        filt = tk.StringVar()
        ef = tk.Entry(search_frame, textvariable=filt, font=("Segoe UI", 10))
        ef.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8))
        tk.Label(
            search_frame, text=self._t("discontinue_search_tip"),
            font=("Segoe UI", 8), bg=COLORS["app"], fg=COLORS["muted"],
        ).pack(side=tk.LEFT)
        tk.Label(
            d,
            text=self._t("discontinue_hover_tip"),
            font=("Segoe UI", 8),
            bg=COLORS["app"],
            fg=COLORS["muted"],
            wraplength=1220,
            justify=tk.LEFT,
        ).grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 2))

        # Row 3 — main content (tree + preview, expands)
        body = tk.Frame(d, bg=COLORS["app"])
        body.grid(row=3, column=0, sticky="nsew", padx=14, pady=(4, 0))
        body.grid_columnconfigure(0, weight=3)
        body.grid_columnconfigure(1, weight=1, minsize=360)
        body.grid_rowconfigure(0, weight=1)

        # ---- Left: Treeview ----
        tree_frame = tk.Frame(body, bg=COLORS["app"])
        tree_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        sty = ttk.Style(d)
        row_h = 80 if pil_ok else 24
        try:
            sty.configure("Disc.Treeview", rowheight=row_h, font=("Segoe UI", 9))
            sty.configure("Disc.Treeview.Heading", font=("Segoe UI", 9, "bold"))
            sty.map("Disc.Treeview", background=[("selected", "#dbeafe")])
        except tk.TclError:
            pass

        cols = ("row", "shop", "stall", "title")
        tree = ttk.Treeview(
            tree_frame, columns=cols, show="tree headings",
            selectmode="browse", style="Disc.Treeview",
        )
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        tree.heading("#0", text=self._t("discontinue_col_thumb"))
        tree.column("#0", width=88 if pil_ok else 20, stretch=False, anchor=tk.CENTER)
        tree.heading("row", text=self._t("discontinue_col_row"))
        tree.column("row", width=40, stretch=False, anchor=tk.CENTER)
        tree.heading("shop", text=self._t("discontinue_col_shop"))
        tree.column("shop", width=100, stretch=False)
        tree.heading("stall", text=self._t("discontinue_col_stall"))
        tree.column("stall", width=62, stretch=False, anchor=tk.CENTER)
        tree.heading("title", text=self._t("discontinue_col_title"))
        tree.column("title", width=380, stretch=True)

        # ---- Right: Preview panel ----
        preview = tk.Frame(body, bg=COLORS["card"], highlightthickness=1, highlightbackground=COLORS["border"])
        preview.grid(row=0, column=1, sticky="nsew")
        preview.grid_rowconfigure(3, weight=1)
        preview.grid_columnconfigure(0, weight=1)

        tk.Label(
            preview, text=self._t("discontinue_preview_title"),
            font=("Segoe UI", 10, "bold"), bg=COLORS["card"], fg=COLORS["text"],
            anchor=tk.W,
        ).grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 2))

        preview_meta = tk.Label(
            preview, text="", bg=COLORS["card"], fg=COLORS["muted"],
            font=("Segoe UI", 9), anchor=tk.W, justify=tk.LEFT, wraplength=340,
        )
        preview_meta.grid(row=1, column=0, sticky="ew", padx=10)

        preview_img = tk.Label(preview, bg="#f1f5f9", borderwidth=0, highlightthickness=0)
        preview_img.grid(row=2, column=0, padx=10, pady=(8, 4))

        preview_title = tk.Text(
            preview, wrap=tk.WORD, font=("Segoe UI", 10),
            bg=COLORS["card"], fg=COLORS["text"], relief=tk.FLAT,
            highlightthickness=0, padx=6, pady=6,
        )
        preview_title.grid(row=3, column=0, sticky="nsew", padx=10, pady=(4, 10))
        preview_title.configure(state=tk.DISABLED)

        def set_preview(r: ProductMapPickerRow | None) -> None:
            preview_photo_ref.clear()
            preview_img.configure(image="", text="")
            preview_title.configure(state=tk.NORMAL)
            preview_title.delete("1.0", tk.END)
            if r is None:
                preview_meta.configure(text=self._t("discontinue_preview_placeholder"))
                preview_title.configure(state=tk.DISABLED)
                return
            shop_str = r.shop_name or "—"
            stall_str = r.stall or "—"
            lines = [
                f"Row {r.row_num}",
                f"{self._t('discontinue_col_shop')}: {shop_str}",
                f"{self._t('discontinue_col_stall')}: {stall_str}",
            ]
            preview_meta.configure(text="\n".join(lines), fg=COLORS["text"])
            preview_title.insert(tk.END, r.title)
            preview_title.configure(state=tk.DISABLED)
            big = bytes_to_thumb(row_photos.get(r.row_num), 240)
            if big is not None:
                preview_img.configure(image=big, text="")
                preview_photo_ref.append(big)
            else:
                no_photo = "(no photo)" if self._lang == "en" else "（无图）"
                preview_img.configure(image="", text=no_photo, fg=COLORS["muted"], font=("Segoe UI", 9))

        set_preview(None)

        # Hover zoom popup (large image near cursor; debounced per row)
        _hover_after_id: list[object | None] = [None]
        _hover_tip_win: list[tk.Toplevel | None] = [None]
        _hover_photo: list[object] = []
        _hover_active_iid: list[str | None] = [None]
        HOVER_DELAY_MS = 220
        # Display cap: upscale/downscale to fit; shrink on small displays (≤ half the shorter screen side, max 720)
        try:
            _sh_scr = int(d.winfo_screenheight())
            _sw_scr = int(d.winfo_screenwidth())
        except tk.TclError:
            _sh_scr, _sw_scr = 1080, 1920
        HOVER_MAX_PX = min(720, max(300, min(_sh_scr, _sw_scr) // 2))

        def _hide_product_hover() -> None:
            if _hover_after_id[0] is not None:
                try:
                    d.after_cancel(_hover_after_id[0])
                except tk.TclError:
                    pass
                _hover_after_id[0] = None
            _hover_active_iid[0] = None
            if _hover_tip_win[0] is not None:
                try:
                    _hover_tip_win[0].destroy()
                except tk.TclError:
                    pass
                _hover_tip_win[0] = None
            _hover_photo.clear()

        def _photo_for_hover(raw: bytes) -> object | None:
            if not pil_ok or Image is None or ImageTk is None:
                return None
            try:
                im = Image.open(BytesIO(raw))
            except Exception:
                return None
            if im.mode not in ("RGB", "RGBA"):
                im = im.convert("RGBA") if "A" in im.mode else im.convert("RGB")
            if im.mode == "RGBA":
                bg = Image.new("RGB", im.size, (255, 255, 255))
                bg.paste(im, mask=im.split()[3])
                im = bg
            w, h = im.size
            if w >= 1 and h >= 1:
                scale = min(HOVER_MAX_PX / w, HOVER_MAX_PX / h)
                nw = max(1, int(round(w * scale)))
                nh = max(1, int(round(h * scale)))
                if (nw, nh) != (w, h):
                    im = im.resize((nw, nh), Image.Resampling.LANCZOS)
            ph = ImageTk.PhotoImage(im)
            _hover_photo.append(ph)
            return ph

        def _show_hover_for_iid(iid: str) -> None:
            _hover_after_id[0] = None
            px, py = tree.winfo_pointerx(), tree.winfo_pointery()
            lx = px - tree.winfo_rootx()
            ly = py - tree.winfo_rooty()
            if lx < 0 or ly < 0 or lx >= tree.winfo_width() or ly >= tree.winfo_height():
                _hide_product_hover()
                return
            if tree.identify_region(lx, ly) != "tree":
                _hover_active_iid[0] = None
                return
            if tree.identify_row(ly) != iid:
                _hover_active_iid[0] = None
                return
            r = row_by_iid.get(iid)
            if r is None:
                _hover_active_iid[0] = None
                return
            raw = row_photos.get(r.row_num)
            if not raw:
                _hover_active_iid[0] = None
                return
            ph = _photo_for_hover(raw)
            if ph is None:
                _hover_active_iid[0] = None
                return
            if _hover_tip_win[0] is not None:
                try:
                    _hover_tip_win[0].destroy()
                except tk.TclError:
                    pass
                _hover_tip_win[0] = None
            tip = tk.Toplevel(d)
            tip.overrideredirect(True)
            try:
                tip.attributes("-topmost", True)
            except tk.TclError:
                pass
            tip.configure(bg="#ffffff", bd=0, highlightthickness=0)
            tk.Label(
                tip,
                image=ph,
                bg="#ffffff",
                bd=0,
                highlightthickness=0,
                relief=tk.FLAT,
            ).pack()
            tip.update_idletasks()
            tw, th = tip.winfo_reqwidth(), tip.winfo_reqheight()
            sw, sh = tip.winfo_screenwidth(), tip.winfo_screenheight()
            x = min(max(12, px + 20), sw - tw - 12)
            y = min(max(12, py + 20), sh - th - 12)
            tip.geometry(f"+{x}+{y}")
            _hover_tip_win[0] = tip

        def on_tree_motion(_e: tk.Event | None = None) -> None:
            if not pil_ok:
                return
            px, py = tree.winfo_pointerx(), tree.winfo_pointery()
            lx = px - tree.winfo_rootx()
            ly = py - tree.winfo_rooty()
            if lx < 0 or ly < 0 or lx >= tree.winfo_width() or ly >= tree.winfo_height():
                _hide_product_hover()
                return
            if tree.identify_region(lx, ly) != "tree":
                _hide_product_hover()
                return
            iid = tree.identify_row(ly)
            if not iid:
                _hide_product_hover()
                return
            r = row_by_iid.get(iid)
            if r is None or not row_photos.get(r.row_num):
                _hide_product_hover()
                return
            if iid == _hover_active_iid[0]:
                return
            _hide_product_hover()
            _hover_active_iid[0] = iid
            _hover_after_id[0] = d.after(HOVER_DELAY_MS, lambda i=iid: _show_hover_for_iid(i))

        def on_tree_leave(_e: tk.Event | None = None) -> None:
            _hide_product_hover()

        tree.bind("<Motion>", on_tree_motion)
        tree.bind("<Leave>", on_tree_leave)

        def row_matches_query(r: ProductMapPickerRow, q: str) -> bool:
            if not q:
                return True
            blob = " ".join((r.title, r.shop_name, r.stall)).lower()
            return q in blob

        def refill_tree() -> None:
            _hide_product_hover()
            tree.delete(*tree.get_children())
            row_by_iid.clear()
            tk_img_refs.clear()
            q = filt.get().strip().lower()
            for r in all_rows:
                if not row_matches_query(r, q):
                    continue
                iid = f"r{r.row_num}"
                row_by_iid[iid] = r
                t_short = r.title if len(r.title) <= 80 else r.title[:77] + "…"
                dash = "—"
                raw = row_photos.get(r.row_num)
                thumb = bytes_to_thumb(raw, 72) if pil_ok else None
                kw: dict = dict(
                    values=(r.row_num, r.shop_name or dash, r.stall or dash, t_short),
                )
                if thumb is not None:
                    kw["image"] = thumb
                    kw["text"] = ""
                else:
                    kw["text"] = dash
                tree.insert("", tk.END, iid=iid, **kw)

        def on_tree_select(_evt: object | None = None) -> None:
            sel = tree.selection()
            if not sel:
                set_preview(None)
                return
            set_preview(row_by_iid.get(sel[0]))

        filt.trace_add("write", lambda *_: refill_tree())
        refill_tree()
        tree.bind("<<TreeviewSelect>>", on_tree_select)

        def do_mark() -> None:
            sel = tree.selection()
            if not sel:
                messagebox.showinfo(
                    self._t("discontinue_title"), self._t("discontinue_no_selection"), parent=d,
                )
                return
            r = row_by_iid.get(sel[0])
            if r is None:
                return
            mark_btn.config(state=tk.DISABLED)
            close_btn.config(state=tk.DISABLED)
            prev_cursor = d.cget("cursor")
            d.config(cursor="watch")
            d.update_idletasks()
            try:
                mark_product_map_discontinued_by_row(FILE_SUPPLIER_CATALOG, r.row_num)
            except Exception as e:
                messagebox.showerror(self._t("file_open_fail_title"), str(e), parent=d)
                return
            finally:
                # Always restore cursor and re-enable buttons regardless of success or failure,
                # so the dialog is never left in a locked state on unexpected exceptions.
                d.config(cursor=prev_cursor)
                mark_btn.config(state=tk.NORMAL)
                close_btn.config(state=tk.NORMAL)
            nonlocal all_rows
            try:
                row_photos.clear()
                row_photos.update(
                    extract_photos_from_xlsx(FILE_SUPPLIER_CATALOG, sheet_name=CATALOG_SHEET, photo_col_idx=0)
                )
            except Exception:
                pass
            all_rows = list_product_map_rows_for_picker(FILE_SUPPLIER_CATALOG)
            refill_tree()
            set_preview(None)
            # Redraw list before the modal so the removed row is not visible behind the message
            d.update_idletasks()
            d.update()
            messagebox.showinfo(
                self._t("discontinue_title"), self._t("discontinue_done"), parent=d,
            )

        def _on_tree_double_click(evt: tk.Event) -> None:
            # Only trigger on actual data rows/thumbnails, not column headings or separators.
            region = tree.identify_region(evt.x, evt.y)
            if region not in ("tree", "cell"):
                return
            do_mark()

        tree.bind("<Double-Button-1>", _on_tree_double_click)

        # Row 4 — buttons (fixed height)
        btn_bar = tk.Frame(d, bg=COLORS["app"])
        btn_bar.grid(row=4, column=0, sticky="ew", padx=14, pady=(8, 12))
        mark_btn = tk.Button(
            btn_bar, text=self._t("discontinue_apply"), font=("Segoe UI", 10, "bold"),
            bg="#dc2626", fg="#ffffff", activebackground="#b91c1c", activeforeground="#ffffff",
            relief=tk.FLAT, padx=18, pady=6, cursor="hand2", command=do_mark,
        )
        mark_btn.pack(side=tk.LEFT)
        def _close_discontinue_dialog() -> None:
            _hide_product_hover()
            d.destroy()

        close_btn = tk.Button(
            btn_bar, text=self._t("discontinue_close"), font=("Segoe UI", 10),
            bg=COLORS["strip"], fg=COLORS["text"], relief=tk.FLAT, padx=14, pady=6,
            cursor="hand2", command=_close_discontinue_dialog,
        )
        close_btn.pack(side=tk.RIGHT)
        d.protocol("WM_DELETE_WINDOW", _close_discontinue_dialog)

        ef.focus_set()

    def _open_catalog_backups_dialog(self) -> None:
        if (
            list_supplier_catalog_backups is None
            or restore_supplier_catalog is None
            or catalog_backup_dir is None
            or backup_supplier_catalog_before_write is None
        ):
            messagebox.showinfo(self._t("msg_missing_title"), self._t("edit_no_import"))
            return

        dlg = tk.Toplevel(self)
        dlg.title(self._t("cat_backup_title"))
        dlg.minsize(420, 320)
        dlg.geometry("520x380")
        dlg.transient(self)
        dlg.configure(bg=COLORS["app"])

        outer = ttk.Frame(dlg, style="App.TFrame", padding=12)
        outer.pack(fill=tk.BOTH, expand=True)

        ttk.Label(outer, text=self._t("cat_backup_hint"), wraplength=480).pack(
            anchor=tk.W, pady=(0, 8)
        )

        lb_frame = ttk.Frame(outer)
        lb_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        scroll = ttk.Scrollbar(lb_frame)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        listbox = tk.Listbox(
            lb_frame,
            height=12,
            font=("Segoe UI", 10),
            yscrollcommand=scroll.set,
            selectmode=tk.SINGLE,
        )
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.config(command=listbox.yview)

        paths: list[Path] = []

        def refresh() -> None:
            nonlocal paths
            listbox.delete(0, tk.END)
            paths = list_supplier_catalog_backups(FILE_SUPPLIER_CATALOG)
            if not paths:
                listbox.insert(tk.END, self._t("cat_backup_none"))
                return
            for p in paths:
                listbox.insert(tk.END, p.name)

        refresh()

        btn_row = ttk.Frame(outer)
        btn_row.pack(fill=tk.X, pady=(4, 0))

        def open_folder() -> None:
            bdir = catalog_backup_dir(FILE_SUPPLIER_CATALOG)
            bdir.mkdir(parents=True, exist_ok=True)
            self._open_file_path(bdir)

        def snapshot() -> None:
            if not FILE_SUPPLIER_CATALOG.is_file():
                messagebox.showinfo(
                    self._t("msg_missing_title"),
                    self._t("cat_backup_snapshot_no_file"),
                )
                return
            dest = backup_supplier_catalog_before_write(
                FILE_SUPPLIER_CATALOG, "manual_snapshot"
            )
            if dest:
                messagebox.showinfo(
                    self._t("cat_backup_title"),
                    self._t("cat_backup_snapshot_ok", path=str(dest)),
                )
                refresh()

        def restore() -> None:
            if not paths:
                return
            sel = listbox.curselection()
            if not sel:
                messagebox.showinfo(
                    self._t("cat_backup_title"),
                    self._t("cat_backup_restore_pick"),
                )
                return
            idx = int(sel[0])
            if idx < 0 or idx >= len(paths):
                return
            chosen = paths[idx]
            if not messagebox.askyesno(
                self._t("cat_backup_title"),
                self._t("cat_backup_restore_confirm", path=chosen.name),
            ):
                return
            try:
                restore_supplier_catalog(FILE_SUPPLIER_CATALOG, chosen)
            except OSError as exc:
                messagebox.showerror(self._t("cat_backup_title"), str(exc))
                return
            messagebox.showinfo(
                self._t("cat_backup_title"), self._t("cat_backup_restore_ok")
            )
            refresh()

        ttk.Button(
            btn_row, text=self._t("cat_backup_open_folder"), command=open_folder
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            btn_row, text=self._t("cat_backup_snapshot"), command=snapshot
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            btn_row, text=self._t("cat_backup_restore"), command=restore
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            btn_row, text=self._t("cat_backup_close"), command=dlg.destroy
        ).pack(side=tk.RIGHT)

        dlg.focus_set()

    def _open_data_folder(self) -> None:
        self._reveal_directory(DATA_DIR)

    def _open_file_path(self, path: Path) -> None:
        if not path.exists():
            messagebox.showinfo(self._t("file_missing_title"), self._t("file_missing_body", path=str(path)))
            return
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.run(["open", str(path)], check=False)
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except OSError as e:
            messagebox.showerror(self._t("file_open_fail_title"), self._t("file_open_fail_body", err=e))

    def _resolve_charm_images_dir(self) -> Path:
        raw = self._charm_dir.get().strip() if self._charm_dir else ""
        if raw:
            return Path(raw).expanduser().resolve()
        return DEFAULT_CHARM_IMAGES_DIR.resolve()

    def _set_chrome_busy(self, busy: bool) -> None:
        if busy:
            ttk_state = ["disabled"]
            if self._run_btn:
                self._run_btn.config(state=tk.DISABLED, bg="#94a3b8", fg="#f1f5f9")
        else:
            ttk_state = ["!disabled"]
            if self._run_btn:
                self._run_btn.config(state=tk.NORMAL, bg=COLORS["run"], fg="#ffffff")
        for b in (
            self._btn_edit_products,
            self._btn_catalog_backups,
            self._btn_charm_browse,
            self._btn_charm_import,
            self._btn_charm_open_folder,
            self._btn_charm_reorder_open,
            self._btn_pdf_browse,
            self._btn_pdf_run_new,
            self._btn_pdf_open_folder,
            self._btn_pdf_move_backup,
        ):
            if b is not None:
                b.state(ttk_state)
        if self._charm_drop_vision_cb is not None:
            self._charm_drop_vision_cb.state(ttk_state)

    def _charm_stage_files(self, paths: list[Path]) -> int:
        dest = self._resolve_charm_images_dir()
        dest.mkdir(parents=True, exist_ok=True)
        n = 0
        for src in paths:
            try:
                p = src if isinstance(src, Path) else Path(src)
                if not p.is_file():
                    continue
                ext = p.suffix.lower()
                if ext not in CHARM_IMAGE_EXTS:
                    continue
                stem_clean = re.sub(r"[^\w\-.]+", "_", p.stem, flags=re.UNICODE)
                stem_clean = stem_clean.strip("._") or "photo"
                stem_clean = stem_clean[:72]
                name = f"{CHARM_INCOMING_PREFIX}{secrets.token_hex(4)}_{stem_clean}{ext}"
                shutil.copy2(p, dest / name)
                n += 1
            except OSError:
                continue
        return n

    def _charm_on_paths_dropped(self, paths: list[Path]) -> None:
        if self._run_busy:
            messagebox.showinfo(self._t("msg_busy_title"), self._t("msg_busy"))
            return
        if not paths:
            return
        n = self._charm_stage_files(paths)
        if n == 0:
            messagebox.showinfo(self._t("charm_msg_title"), self._t("charm_drop_no_valid"))
            return
        self._append_log(self._t("charm_drop_staged", n=n))

    def _charm_browse_photos(self) -> None:
        if self._run_busy:
            messagebox.showinfo(self._t("msg_busy_title"), self._t("msg_busy"))
            return
        picked = filedialog.askopenfilenames(
            parent=self,
            filetypes=[
                ("Images", "*.png *.jpg *.jpeg *.webp"),
                ("PNG", "*.png"),
                ("JPEG", "*.jpg *.jpeg"),
                ("WebP", "*.webp"),
                ("All", "*.*"),
            ],
            title=self._t("charm_drop_browse_title"),
        )
        if not picked:
            return
        self._charm_on_paths_dropped([Path(s) for s in picked])

    def _open_charm_images_folder(self) -> None:
        self._reveal_directory(self._resolve_charm_images_dir())

    def _charm_import_workbook(self) -> None:
        if self._run_busy:
            messagebox.showinfo(self._t("msg_busy_title"), self._t("msg_busy"))
            return
        if not GENERATOR.is_file():
            messagebox.showerror(self._t("msg_missing_title"), self._t("msg_missing") + str(GENERATOR))
            return
        dest = self._resolve_charm_images_dir()
        dest.mkdir(parents=True, exist_ok=True)
        staged = [
            p
            for p in dest.iterdir()
            if p.is_file()
            and p.name.startswith(CHARM_INCOMING_PREFIX)
            and p.suffix.lower() in CHARM_IMAGE_EXTS
        ]
        if not staged:
            messagebox.showinfo(self._t("charm_msg_title"), self._t("charm_drop_nothing_to_import"))
            return

        args: list[str] = [
            sys.executable,
            str(GENERATOR),
            "--project-dir",
            str(PROJECT_ROOT),
            "--import-charm-images",
            *charm_import_pattern_argv(),
        ]
        cd = self._charm_dir.get().strip()
        if cd:
            args.extend(["--charm-images-dir", cd])
        if self._charm_drop_vision.get():
            args.append("--import-charm-vision-sku")

        self._run_busy = True
        self._set_chrome_busy(True)
        self._append_log(self._t("charm_import_start"))
        self._append_log(" ".join(args) + "\n\n")

        def work() -> None:
            try:
                proc = subprocess.run(
                    args,
                    cwd=str(PROJECT_ROOT),
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                )
                out = proc.stdout or ""
                err = proc.stderr or ""
                self._log_q.put(out)
                if err:
                    self._log_q.put(self._t("log_messages") + err)
                self._log_q.put(self._t("log_finished", code=proc.returncode))
            except Exception as e:
                self._log_q.put(self._t("log_error", e=e))
            finally:

                def done() -> None:
                    self._run_busy = False
                    self._set_chrome_busy(False)

                self.after(0, done)

        threading.Thread(target=work, daemon=True).start()

    def _open_charm_reorder_dialog(self) -> None:
        if reorder_charm_library_rows is None or load_charm_library is None:
            messagebox.showerror(
                self._t("msg_missing_title"),
                self._t("charm_reorder_no_import"),
            )
            return
        if not FILE_SUPPLIER_CATALOG.exists():
            messagebox.showinfo(
                self._t("msg_missing_title"),
                self._t("reorder_no_catalog"),
            )
            return
        try:
            entries = load_charm_library(FILE_SUPPLIER_CATALOG)
        except Exception as e:
            messagebox.showerror(self._t("file_open_fail_title"), str(e))
            return
        if not entries:
            messagebox.showinfo(
                self._t("reorder_title"), self._t("reorder_empty")
            )
            return
        charm_images_dir = self._resolve_charm_images_dir()
        _CharmReorderDialog(self, entries, FILE_SUPPLIER_CATALOG, charm_images_dir)

    def _run(self) -> None:
        if self._run_busy:
            messagebox.showinfo(self._t("msg_busy_title"), self._t("msg_busy"))
            return
        if not GENERATOR.is_file():
            messagebox.showerror(self._t("msg_missing_title"), self._t("msg_missing") + str(GENERATOR))
            return
        job = self._job_var.get()
        self._spawn_generator(
            self._collect_generator_args(job),
            log_intro=self._t("log_start"),
        )


class _ProductMapEditorDialog:
    """
    Modal dialog for editing Shop Name, Stall, Charm Shop, and Charm Code
    on Product Map rows — with photo preview, search filtering, and
    dropdown-style comboboxes populated from the Suppliers and Charm sheets.
    """

    _THUMB = 68
    _ROW_H = 80

    def __init__(
        self,
        parent: App,
        all_rows: list,
        row_photos: dict[int, bytes],
        *,
        supplier_shops: list[str],
        supplier_stalls: list[str],
        charm_shop_names: list[str],
        charm_entries: dict,
    ) -> None:
        self._parent         = parent
        self._all_rows       = all_rows
        self._row_photos     = row_photos
        self._lang           = parent._lang
        self._sup_shops      = supplier_shops
        self._sup_stalls     = supplier_stalls
        self._charm_shops    = charm_shop_names
        self._charm_entries  = charm_entries
        self._charm_codes    = list(charm_entries.keys())
        self._tk_img_refs: list[object] = []
        self._row_by_iid: dict[str, object] = {}
        self._selected_row: object | None = None
        self._preview_photo_ref: list[object] = []
        self._pending_photo_bytes: bytes | None = None

        pil_ok = Image is not None and ImageTk is not None
        self._pil_ok = pil_ok

        d = tk.Toplevel(parent)
        self._d = d
        d.title(self._t("edit_title"))
        d.transient(parent)
        d.grab_set()
        d.configure(bg=COLORS["app"])
        d.geometry("1380x820")
        d.minsize(1020, 600)

        d.grid_columnconfigure(0, weight=1)
        d.grid_rowconfigure(3, weight=1)

        # ── Row 0: hero ──
        hero = tk.Frame(d, bg=COLORS["hero"], highlightthickness=0)
        hero.grid(row=0, column=0, sticky="ew")
        tk.Label(
            hero, text=self._t("edit_heading"),
            font=("Segoe UI", 12, "bold"), fg="#ffffff", bg=COLORS["hero"],
        ).pack(anchor=tk.W, padx=14, pady=(10, 2))
        tk.Label(
            hero, text=self._t("edit_intro"),
            wraplength=1300, justify=tk.LEFT,
            font=("Segoe UI", 9), fg="#e2e8f0", bg=COLORS["hero"],
        ).pack(anchor=tk.W, padx=14, pady=(0, 8))

        # ── Row 1: search ──
        sf = tk.Frame(d, bg=COLORS["app"])
        sf.grid(row=1, column=0, sticky="ew", padx=14, pady=(8, 4))
        tk.Label(sf, text=self._t("edit_search_label"),
                 font=("Segoe UI", 10, "bold"), bg=COLORS["app"], fg=COLORS["text"],
                 ).pack(side=tk.LEFT)
        self._filt = tk.StringVar()
        ef = tk.Entry(sf, textvariable=self._filt, font=("Segoe UI", 10))
        ef.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8))
        tk.Label(sf, text=self._t("edit_search_tip"),
                 font=("Segoe UI", 8), bg=COLORS["app"], fg=COLORS["muted"],
                 ).pack(side=tk.LEFT)

        # ── Row 2: hover tip ──
        tk.Label(
            d, text=self._t("edit_hover_tip"),
            font=("Segoe UI", 8), bg=COLORS["app"], fg=COLORS["muted"],
            wraplength=1300, justify=tk.LEFT,
        ).grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 2))

        # ── Row 3: body (tree + edit panel) ──
        body = tk.Frame(d, bg=COLORS["app"])
        body.grid(row=3, column=0, sticky="nsew", padx=14, pady=(4, 0))
        body.grid_columnconfigure(0, weight=3)
        body.grid_columnconfigure(1, weight=1, minsize=360)
        body.grid_rowconfigure(0, weight=1)

        # ── Left: Treeview ──
        tf = tk.Frame(body, bg=COLORS["app"])
        tf.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        tf.grid_rowconfigure(0, weight=1)
        tf.grid_columnconfigure(0, weight=1)

        sty = ttk.Style(d)
        row_h = self._ROW_H if pil_ok else 24
        try:
            sty.configure("Edit.Treeview", rowheight=row_h, font=("Segoe UI", 9))
            sty.configure("Edit.Treeview.Heading", font=("Segoe UI", 9, "bold"))
            sty.map("Edit.Treeview", background=[("selected", "#dbeafe")])
        except tk.TclError:
            pass

        cols = ("row", "shop", "stall", "charm_shop", "charm_code", "title")
        tree = ttk.Treeview(
            tf, columns=cols, show="tree headings",
            selectmode="browse", style="Edit.Treeview",
        )
        vsb = ttk.Scrollbar(tf, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self._tree = tree

        tree.heading("#0", text=self._t("edit_col_photo"))
        tree.column("#0", width=88 if pil_ok else 20, stretch=False, anchor=tk.CENTER)
        tree.heading("row",         text=self._t("edit_col_row"))
        tree.column ("row",         width=40, stretch=False, anchor=tk.CENTER)
        tree.heading("shop",        text=self._t("edit_col_shop"))
        tree.column ("shop",        width=100, stretch=False)
        tree.heading("stall",       text=self._t("edit_col_stall"))
        tree.column ("stall",       width=60, stretch=False, anchor=tk.CENTER)
        tree.heading("charm_shop",  text=self._t("edit_col_charm_shop"))
        tree.column ("charm_shop",  width=100, stretch=False)
        tree.heading("charm_code",  text=self._t("edit_col_charm_code"))
        tree.column ("charm_code",  width=100, stretch=False, anchor=tk.CENTER)
        tree.heading("title",       text=self._t("edit_col_title"))
        tree.column ("title",       width=300, stretch=True)

        # ── Right: edit panel ──
        panel = tk.Frame(body, bg=COLORS["card"],
                         highlightthickness=1, highlightbackground=COLORS["border"])
        panel.grid(row=0, column=1, sticky="nsew")
        panel.grid_columnconfigure(0, weight=1)

        tk.Label(
            panel, text=self._t("edit_preview_title"),
            font=("Segoe UI", 10, "bold"), bg=COLORS["card"], fg=COLORS["text"],
        ).grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 2))

        # Photo preview + upload controls (all in one frame at row 1)
        photo_area = tk.Frame(panel, bg=COLORS["card"])
        photo_area.grid(row=1, column=0, sticky="ew", padx=12, pady=(4, 2))
        photo_area.grid_columnconfigure(0, weight=1)

        self._preview_img_label = tk.Label(
            photo_area, bg=COLORS["card"], bd=0,
            highlightthickness=0, relief=tk.FLAT,
        )
        self._preview_img_label.grid(row=0, column=0, pady=(0, 4))

        self._upload_photo_btn = tk.Button(
            photo_area,
            text=self._t("edit_btn_upload_photo"),
            command=self._upload_photo,
            bg=COLORS["accent_soft"], fg=COLORS["accent"],
            activebackground="#bfdbfe", activeforeground=COLORS["accent"],
            font=("Segoe UI", 8, "bold"),
            padx=10, pady=4,
            cursor="hand2",
            relief=tk.FLAT, borderwidth=0, highlightthickness=0,
            state=tk.DISABLED,
        )
        self._upload_photo_btn.grid(row=1, column=0, sticky="ew", pady=(0, 2))

        self._upload_status_lbl = tk.Label(
            photo_area, text="",
            font=("Segoe UI", 8, "bold"), bg=COLORS["card"], fg="#047857",
            wraplength=310, justify=tk.CENTER,
        )
        self._upload_status_lbl.grid(row=2, column=0, sticky="ew")

        # Title label
        self._preview_title = tk.Label(
            panel, text="", wraplength=320, justify=tk.LEFT,
            font=("Segoe UI", 9), bg=COLORS["card"], fg=COLORS["text"],
        )
        self._preview_title.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 8))

        # Form fields
        form = tk.Frame(panel, bg=COLORS["card"])
        form.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 4))
        form.grid_columnconfigure(1, weight=1)

        self._fields: dict[str, ttk.Combobox | tk.Entry] = {}
        combo_defs = [
            ("shop_name",  "edit_lbl_shop",       self._sup_shops),
            ("stall",      "edit_lbl_stall",      self._sup_stalls),
            ("charm_shop", "edit_lbl_charm_shop",  self._charm_shops),
        ]
        for i, (key, lbl_key, values) in enumerate(combo_defs):
            tk.Label(
                form, text=self._t(lbl_key),
                font=("Segoe UI", 9, "bold"), bg=COLORS["card"], fg=COLORS["text"],
                anchor=tk.W,
            ).grid(row=i, column=0, sticky="w", padx=(0, 8), pady=3)
            cb = ttk.Combobox(form, values=values, font=("Segoe UI", 9), width=24)
            cb.grid(row=i, column=1, sticky="ew", pady=3)
            self._fields[key] = cb

        # Charm Code — entry + visual browse button
        cc_row = len(combo_defs)
        tk.Label(
            form, text=self._t("edit_lbl_charm_code"),
            font=("Segoe UI", 9, "bold"), bg=COLORS["card"], fg=COLORS["text"],
            anchor=tk.W,
        ).grid(row=cc_row, column=0, sticky="w", padx=(0, 8), pady=3)

        cc_frame = tk.Frame(form, bg=COLORS["card"])
        cc_frame.grid(row=cc_row, column=1, sticky="ew", pady=3)
        cc_frame.grid_columnconfigure(0, weight=1)

        cc_entry = tk.Entry(cc_frame, font=("Segoe UI", 9))
        cc_entry.grid(row=0, column=0, sticky="ew")
        self._fields["charm_code"] = cc_entry

        self._cc_pick_btn = ttk.Button(
            cc_frame, text=self._t("edit_charm_pick"),
            command=self._open_charm_picker,
            style="Tool.TButton", width=5,
        )
        self._cc_pick_btn.grid(row=0, column=1, padx=(4, 0))

        # Charm preview thumbnail (shows next to charm code entry after selection)
        self._charm_preview_frame = tk.Frame(form, bg=COLORS["card"])
        self._charm_preview_frame.grid(
            row=cc_row + 1, column=0, columnspan=2, sticky="ew", pady=(2, 0)
        )
        self._charm_preview_img = tk.Label(
            self._charm_preview_frame, bg=COLORS["card"], bd=0,
            highlightthickness=0, relief=tk.FLAT,
        )
        self._charm_preview_img.pack(side=tk.LEFT, padx=(0, 8))
        self._charm_preview_text = tk.Label(
            self._charm_preview_frame, text="",
            font=("Segoe UI", 8), fg=COLORS["muted"], bg=COLORS["card"],
            wraplength=200, justify=tk.LEFT, anchor=tk.W,
        )
        self._charm_preview_text.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._charm_preview_photo_ref: list[object] = []

        cc_entry.bind("<KeyRelease>", lambda _e: self._update_charm_preview())
        cc_entry.bind("<FocusOut>", lambda _e: self._update_charm_preview())

        # Save button
        self._save_btn = tk.Button(
            panel, text=self._t("edit_btn_save"),
            command=self._save_current,
            bg=COLORS["run"], fg="#ffffff",
            activebackground=COLORS["run_hover"],
            activeforeground="#ffffff",
            font=("Segoe UI", 10, "bold"),
            padx=18, pady=6,
            cursor="hand2",
            relief=tk.FLAT, borderwidth=0, highlightthickness=0,
        )
        self._save_btn.grid(row=4, column=0, padx=12, pady=(10, 4), sticky="ew")

        placeholder = tk.Label(
            panel, text=self._t("edit_preview_placeholder"),
            font=("Segoe UI", 8), fg=COLORS["muted"], bg=COLORS["card"],
            wraplength=320, justify=tk.LEFT,
        )
        placeholder.grid(row=5, column=0, sticky="ew", padx=12, pady=(0, 6))
        self._placeholder_lbl = placeholder

        # ── Danger zone ──
        ttk.Separator(panel, orient=tk.HORIZONTAL).grid(
            row=6, column=0, sticky="ew", padx=12, pady=(10, 0)
        )
        dz_hdr = tk.Frame(panel, bg=COLORS["card"])
        dz_hdr.grid(row=7, column=0, sticky="ew", padx=12, pady=(6, 2))
        tk.Label(
            dz_hdr,
            text=self._t("edit_danger_zone"),
            font=("Segoe UI", 8, "bold"),
            fg="#b91c1c",
            bg=COLORS["card"],
            anchor=tk.W,
        ).pack(side=tk.LEFT)
        tk.Label(
            panel,
            text=self._t("edit_danger_note"),
            font=("Segoe UI", 8),
            fg=COLORS["muted"],
            bg=COLORS["card"],
            wraplength=310,
            justify=tk.LEFT,
            anchor=tk.W,
        ).grid(row=8, column=0, sticky="ew", padx=12, pady=(0, 4))
        self._disc_btn = tk.Button(
            panel,
            text=self._t("edit_btn_discontinue"),
            command=self._mark_discontinued,
            bg="#dc2626",
            fg="#ffffff",
            activebackground="#b91c1c",
            activeforeground="#ffffff",
            disabledforeground="#fca5a5",
            font=("Segoe UI", 9, "bold"),
            padx=14,
            pady=5,
            cursor="hand2",
            relief=tk.FLAT,
            borderwidth=0,
            highlightthickness=0,
            state=tk.DISABLED,
        )
        self._disc_btn.grid(row=9, column=0, padx=12, pady=(2, 14), sticky="ew")

        # ── Row 4: footer ──
        foot = tk.Frame(d, bg=COLORS["app"])
        foot.grid(row=4, column=0, sticky="ew", padx=14, pady=(8, 12))
        ttk.Button(
            foot, text=self._t("edit_btn_close"),
            command=d.destroy, style="Tool.TButton",
        ).pack(side=tk.RIGHT)

        # ── Populate + bindings ──
        self._refill_tree()
        self._filt.trace_add("write", lambda *_: self._refill_tree())
        tree.bind("<<TreeviewSelect>>", self._on_select)

        # Hover zoom
        self._hover_tip_win: list[tk.Toplevel | None] = [None]
        self._hover_active_iid: list[str | None] = [None]
        self._hover_after_id: list[str | None] = [None]
        self._hover_photo: list[object] = []
        tree.bind("<Motion>", self._on_tree_motion)
        tree.bind("<Leave>", self._on_tree_leave)

    # ── helpers ──

    def _t(self, key: str, **kw: object) -> str:
        s = CHROME.get(self._lang, CHROME["en"]).get(key, key)
        return s.format(**kw) if kw else s

    def _thumb(self, raw: bytes | None) -> object | None:
        if not self._pil_ok or not raw or Image is None or ImageTk is None:
            return None
        try:
            im = Image.open(BytesIO(raw))
            if im.mode not in ("RGB", "RGBA"):
                im = im.convert("RGBA" if "A" in im.mode else "RGB")
            if im.mode == "RGBA":
                bg = Image.new("RGB", im.size, (255, 255, 255))
                bg.paste(im, mask=im.split()[3])
                im = bg
            im.thumbnail((self._THUMB, self._THUMB), Image.Resampling.LANCZOS)
            ph = ImageTk.PhotoImage(im)
            self._tk_img_refs.append(ph)
            return ph
        except Exception:
            return None

    # ── tree ──

    def _refill_tree(self) -> None:
        self._tree.delete(*self._tree.get_children())
        self._row_by_iid.clear()
        self._tk_img_refs.clear()
        q = self._filt.get().strip().lower()
        for r in self._all_rows:
            if q:
                blob = " ".join((r.title, r.shop_name, r.stall)).lower()
                if q not in blob:
                    continue
            iid = f"r{r.row_num}"
            self._row_by_iid[iid] = r
            t_short = r.title if len(r.title) <= 60 else r.title[:57] + "…"
            raw = self._row_photos.get(r.row_num)
            thumb = self._thumb(raw) if self._pil_ok else None
            kw: dict = dict(
                values=(
                    r.row_num,
                    r.shop_name or "—",
                    r.stall or "—",
                    r.charm_shop or "—",
                    r.charm_code or "—",
                    t_short,
                ),
            )
            if thumb is not None:
                kw["image"] = thumb
                kw["text"] = ""
            else:
                kw["text"] = "—"
            self._tree.insert("", tk.END, iid=iid, **kw)

    def _on_select(self, _evt: object = None) -> None:
        sel = self._tree.selection()
        if not sel:
            self._set_edit_panel(None)
            return
        self._set_edit_panel(self._row_by_iid.get(sel[0]))

    def _set_edit_panel(self, r: object | None) -> None:
        self._selected_row = r
        self._pending_photo_bytes = None
        self._upload_status_lbl.config(text="")
        if r is None:
            self._preview_title.config(text="")
            self._preview_img_label.config(image="")
            self._preview_photo_ref.clear()
            for w in self._fields.values():
                if isinstance(w, ttk.Combobox):
                    w.set("")
                    w.config(state="disabled")
                else:
                    w.delete(0, tk.END)
                    w.config(state="disabled")
            self._cc_pick_btn.config(state=tk.DISABLED)
            self._upload_photo_btn.config(state=tk.DISABLED)
            self._save_btn.config(state=tk.DISABLED)
            self._disc_btn.config(state=tk.DISABLED)
            self._clear_charm_preview()
            self._placeholder_lbl.config(text=self._t("edit_preview_placeholder"))
            return

        self._placeholder_lbl.config(text="")
        self._preview_title.config(text=r.title)

        raw = self._row_photos.get(r.row_num)
        self._preview_photo_ref.clear()
        if raw and self._pil_ok and Image is not None and ImageTk is not None:
            try:
                im = Image.open(BytesIO(raw))
                if im.mode not in ("RGB", "RGBA"):
                    im = im.convert("RGBA" if "A" in im.mode else "RGB")
                if im.mode == "RGBA":
                    bg = Image.new("RGB", im.size, (255, 255, 255))
                    bg.paste(im, mask=im.split()[3])
                    im = bg
                im.thumbnail((240, 240), Image.Resampling.LANCZOS)
                ph = ImageTk.PhotoImage(im)
                self._preview_photo_ref.append(ph)
                self._preview_img_label.config(image=ph)
            except Exception:
                self._preview_img_label.config(image="")
        else:
            self._preview_img_label.config(image="")

        for w in self._fields.values():
            if isinstance(w, ttk.Combobox):
                w.config(state="normal")
            else:
                w.config(state="normal")
        self._cc_pick_btn.config(state="normal")
        self._upload_photo_btn.config(state=tk.NORMAL)
        self._fields["shop_name"].set(r.shop_name)
        self._fields["stall"].set(r.stall)
        self._fields["charm_shop"].set(r.charm_shop)
        cc_w = self._fields["charm_code"]
        cc_w.delete(0, tk.END)
        cc_w.insert(0, r.charm_code)
        self._update_charm_preview()
        self._save_btn.config(state=tk.NORMAL)
        self._disc_btn.config(state=tk.NORMAL)

    # ── charm code preview + picker ──

    def _clear_charm_preview(self) -> None:
        self._charm_preview_photo_ref.clear()
        self._charm_preview_img.config(image="")
        self._charm_preview_text.config(text="")

    def _update_charm_preview(self) -> None:
        code = self._fields["charm_code"].get().strip()
        entry = self._charm_entries.get(code)
        if not entry:
            self._clear_charm_preview()
            if code and code not in self._charm_entries:
                self._charm_preview_text.config(
                    text=self._t("edit_charm_unknown"),
                    fg="#b91c1c",
                )
            return
        # Show thumbnail + label
        self._charm_preview_photo_ref.clear()
        if entry.photo_bytes and self._pil_ok and Image is not None and ImageTk is not None:
            try:
                im = Image.open(BytesIO(entry.photo_bytes))
                if im.mode not in ("RGB", "RGBA"):
                    im = im.convert("RGBA" if "A" in im.mode else "RGB")
                if im.mode == "RGBA":
                    bg = Image.new("RGB", im.size, (255, 255, 255))
                    bg.paste(im, mask=im.split()[3])
                    im = bg
                im.thumbnail((52, 52), Image.Resampling.LANCZOS)
                ph = ImageTk.PhotoImage(im)
                self._charm_preview_photo_ref.append(ph)
                self._charm_preview_img.config(image=ph)
            except Exception:
                self._charm_preview_img.config(image="")
        else:
            self._charm_preview_img.config(image="")
        parts = [code]
        if entry.sku:
            parts.append(entry.sku)
        if entry.default_charm_shop:
            parts.append(f"({entry.default_charm_shop})")
        self._charm_preview_text.config(text="  —  ".join(parts), fg=COLORS["text"])

    def _open_charm_picker(self) -> None:
        if not self._charm_entries:
            return
        _CharmPickerPopup(
            self._d,
            self._charm_entries,
            lang=self._lang,
            callback=self._on_charm_picked,
        )

    def _on_charm_picked(self, code: str) -> None:
        w = self._fields["charm_code"]
        w.config(state="normal")
        w.delete(0, tk.END)
        w.insert(0, code)
        self._update_charm_preview()

    # ── photo upload ──

    def _upload_photo(self) -> None:
        """Open a file dialog to choose a new product photo and stage it for saving."""
        path = filedialog.askopenfilename(
            parent=self._d,
            title=self._t("edit_upload_photo_title"),
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.webp *.bmp *.gif"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        if Image is None or ImageTk is None:
            messagebox.showerror(
                self._t("file_open_fail_title"),
                "Pillow is required for photo upload.  pip install Pillow",
                parent=self._d,
            )
            return

        try:
            im = Image.open(path)
            if im.mode not in ("RGB", "RGBA"):
                im = im.convert("RGBA" if "A" in im.mode else "RGB")
            if im.mode == "RGBA":
                bg = Image.new("RGB", im.size, (255, 255, 255))
                bg.paste(im, mask=im.split()[3])
                im = bg
            # Encode as JPEG for compact storage in the workbook
            buf = BytesIO()
            im.save(buf, format="JPEG", quality=90)
            self._pending_photo_bytes = buf.getvalue()
        except Exception as exc:
            messagebox.showerror(
                self._t("file_open_fail_title"),
                self._t("edit_photo_save_fail", err=str(exc)),
                parent=self._d,
            )
            return

        # Update the preview in the panel with the new image
        try:
            im_prev = Image.open(BytesIO(self._pending_photo_bytes))
            im_prev.thumbnail((240, 240), Image.Resampling.LANCZOS)
            ph = ImageTk.PhotoImage(im_prev)
            self._preview_photo_ref.clear()
            self._preview_photo_ref.append(ph)
            self._preview_img_label.config(image=ph)
        except Exception:
            pass

        self._upload_status_lbl.config(text=self._t("edit_upload_photo_pending"))

    # ── save ──

    def _save_current(self) -> None:
        r = self._selected_row
        if r is None:
            messagebox.showinfo(
                self._t("edit_title"), self._t("edit_no_selection"), parent=self._d,
            )
            return

        new_shop       = self._fields["shop_name"].get().strip()
        new_stall      = self._fields["stall"].get().strip()
        new_charm_shop = self._fields["charm_shop"].get().strip()
        new_charm_code = self._fields["charm_code"].get().strip()
        pending_photo  = self._pending_photo_bytes

        prev_cursor = self._d.cget("cursor")
        self._d.config(cursor="watch")
        self._save_btn.config(state=tk.DISABLED)
        self._disc_btn.config(state=tk.DISABLED)
        self._upload_photo_btn.config(state=tk.DISABLED)
        self._d.update_idletasks()

        try:
            update_product_map_cells(
                FILE_SUPPLIER_CATALOG,
                r.row_num,
                shop_name=new_shop,
                stall=new_stall,
                charm_shop=new_charm_shop,
                charm_code=new_charm_code,
            )
            if pending_photo and update_product_map_photo is not None:
                update_product_map_photo(FILE_SUPPLIER_CATALOG, r.row_num, pending_photo)
        except Exception as e:
            messagebox.showerror(
                self._t("file_open_fail_title"), str(e), parent=self._d,
            )
            return
        finally:
            self._d.config(cursor=prev_cursor)
            self._save_btn.config(state=tk.NORMAL)
            self._disc_btn.config(state=tk.NORMAL)
            self._upload_photo_btn.config(state=tk.NORMAL)

        # Clear pending photo state
        self._pending_photo_bytes = None
        self._upload_status_lbl.config(text="")

        # Reload photos + rows so the treeview thumbnail reflects the new photo
        if pending_photo:
            try:
                if extract_photos_from_xlsx is not None:
                    self._row_photos = extract_photos_from_xlsx(
                        FILE_SUPPLIER_CATALOG, sheet_name=CATALOG_SHEET, photo_col_idx=0,
                    )
            except Exception:
                pass
        try:
            self._all_rows = list_product_map_rows_for_picker(FILE_SUPPLIER_CATALOG)
        except Exception:
            pass
        cur_iid = f"r{r.row_num}"
        self._refill_tree()
        if self._tree.exists(cur_iid):
            self._tree.selection_set(cur_iid)
            self._tree.see(cur_iid)
            self._set_edit_panel(self._row_by_iid.get(cur_iid))

        short = r.title if len(r.title) <= 60 else r.title[:57] + "…"
        if pending_photo:
            messagebox.showinfo(
                self._t("edit_title"),
                self._t("edit_photo_saved", title=short),
                parent=self._d,
            )
        else:
            messagebox.showinfo(
                self._t("edit_title"),
                self._t("edit_saved", title=short),
                parent=self._d,
            )

    # ── mark discontinued ──

    def _mark_discontinued(self) -> None:
        r = self._selected_row
        if r is None:
            messagebox.showinfo(
                self._t("edit_title"), self._t("edit_no_selection"), parent=self._d,
            )
            return
        if mark_product_map_discontinued_by_row is None:
            messagebox.showerror(
                self._t("msg_missing_title"), self._t("edit_no_import"), parent=self._d,
            )
            return

        short = r.title if len(r.title) <= 80 else r.title[:77] + "…"
        if not messagebox.askyesno(
            self._t("edit_discontinue_confirm_title"),
            self._t("edit_discontinue_confirm", title=short),
            icon="warning",
            parent=self._d,
        ):
            return

        prev_cursor = self._d.cget("cursor")
        self._d.config(cursor="watch")
        self._save_btn.config(state=tk.DISABLED)
        self._disc_btn.config(state=tk.DISABLED)
        self._d.update_idletasks()

        try:
            mark_product_map_discontinued_by_row(FILE_SUPPLIER_CATALOG, r.row_num)
        except Exception as e:
            messagebox.showerror(
                self._t("file_open_fail_title"), str(e), parent=self._d,
            )
            return
        finally:
            self._d.config(cursor=prev_cursor)

        # Reload the product list — discontinued row is now gone from Product Map
        try:
            self._all_rows = list_product_map_rows_for_picker(FILE_SUPPLIER_CATALOG)
            self._row_photos = {}
            if extract_photos_from_xlsx is not None:
                self._row_photos = extract_photos_from_xlsx(
                    FILE_SUPPLIER_CATALOG, sheet_name=CATALOG_SHEET, photo_col_idx=0,
                )
        except Exception:
            pass

        self._set_edit_panel(None)
        self._refill_tree()

        messagebox.showinfo(
            self._t("edit_discontinue_confirm_title"),
            self._t("edit_discontinue_done"),
            parent=self._d,
        )

    # ── hover zoom (same pattern as discontinue dialog) ──

    def _hide_hover(self) -> None:
        aid = self._hover_after_id[0]
        if aid is not None:
            try:
                self._d.after_cancel(aid)
            except (tk.TclError, ValueError):
                pass
            self._hover_after_id[0] = None
        self._hover_active_iid[0] = None
        tip = self._hover_tip_win[0]
        if tip is not None:
            try:
                tip.destroy()
            except tk.TclError:
                pass
            self._hover_tip_win[0] = None

    def _on_tree_motion(self, _e: tk.Event | None = None) -> None:
        if not self._pil_ok:
            return
        tree = self._tree
        px, py = tree.winfo_pointerx(), tree.winfo_pointery()
        lx = px - tree.winfo_rootx()
        ly = py - tree.winfo_rooty()
        if lx < 0 or ly < 0 or lx >= tree.winfo_width() or ly >= tree.winfo_height():
            self._hide_hover()
            return
        if tree.identify_region(lx, ly) != "tree":
            self._hide_hover()
            return
        iid = tree.identify_row(ly)
        if not iid:
            self._hide_hover()
            return
        r = self._row_by_iid.get(iid)
        if r is None or not self._row_photos.get(r.row_num):
            self._hide_hover()
            return
        if iid == self._hover_active_iid[0]:
            return
        self._hide_hover()
        self._hover_active_iid[0] = iid
        self._hover_after_id[0] = self._d.after(
            350, lambda i=iid: self._show_hover(i)
        )

    def _show_hover(self, iid: str) -> None:
        self._hover_after_id[0] = None
        if Image is None or ImageTk is None:
            return
        tree = self._tree
        px, py = tree.winfo_pointerx(), tree.winfo_pointery()
        lx, ly = px - tree.winfo_rootx(), py - tree.winfo_rooty()
        if lx < 0 or ly < 0 or lx >= tree.winfo_width() or ly >= tree.winfo_height():
            return
        if tree.identify_row(ly) != iid:
            return
        r = self._row_by_iid.get(iid)
        if r is None:
            return
        raw = self._row_photos.get(r.row_num)
        if not raw:
            return
        try:
            im = Image.open(BytesIO(raw))
            if im.mode not in ("RGB", "RGBA"):
                im = im.convert("RGBA" if "A" in im.mode else "RGB")
            if im.mode == "RGBA":
                bg = Image.new("RGB", im.size, (255, 255, 255))
                bg.paste(im, mask=im.split()[3])
                im = bg
            w, h = im.size
            max_dim = 320
            if w > max_dim or h > max_dim:
                ratio = min(max_dim / w, max_dim / h)
                im = im.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)
            ph = ImageTk.PhotoImage(im)
            self._hover_photo.clear()
            self._hover_photo.append(ph)
        except Exception:
            return
        tip = tk.Toplevel(self._d)
        tip.overrideredirect(True)
        try:
            tip.attributes("-topmost", True)
        except tk.TclError:
            pass
        tip.configure(bg="#ffffff", bd=0, highlightthickness=0)
        tk.Label(tip, image=ph, bg="#ffffff", bd=0, highlightthickness=0, relief=tk.FLAT).pack()
        tip.update_idletasks()
        tw, th = tip.winfo_reqwidth(), tip.winfo_reqheight()
        sw, sh = tip.winfo_screenwidth(), tip.winfo_screenheight()
        x = min(max(12, px + 20), sw - tw - 12)
        y = min(max(12, py + 20), sh - th - 12)
        tip.geometry(f"+{x}+{y}")
        self._hover_tip_win[0] = tip

    def _on_tree_leave(self, _e: tk.Event | None = None) -> None:
        self._hide_hover()


class _CharmPickerPopup:
    """
    Visual charm selector — shows every Charm Library entry with its photo,
    code, and SKU in a scrollable Treeview.  Click a row (or double-click) to
    select that charm and close the popup.  The first row is "(none)" to let
    the user clear the charm code.
    """

    _THUMB = 64
    _ROW_H = 72

    def __init__(
        self,
        parent: tk.Toplevel | tk.Tk,
        charm_entries: dict,
        *,
        lang: str = "en",
        callback: object = None,
    ) -> None:
        self._entries  = charm_entries
        self._lang     = lang
        self._callback = callback
        self._tk_img_refs: list[object] = []

        pil_ok = Image is not None and ImageTk is not None
        self._pil_ok = pil_ok

        d = tk.Toplevel(parent)
        self._d = d
        d.title(self._t("charm_picker_title"))
        d.transient(parent)
        d.grab_set()
        d.configure(bg=COLORS["app"])
        d.geometry("680x620")
        d.minsize(520, 400)

        d.grid_columnconfigure(0, weight=1)
        d.grid_rowconfigure(3, weight=1)

        # Hero
        hero = tk.Frame(d, bg=COLORS["hero"], highlightthickness=0)
        hero.grid(row=0, column=0, sticky="ew")
        tk.Label(
            hero, text=self._t("charm_picker_heading"),
            font=("Segoe UI", 11, "bold"), fg="#ffffff", bg=COLORS["hero"],
        ).pack(anchor=tk.W, padx=14, pady=(8, 6))

        # Search
        sf = tk.Frame(d, bg=COLORS["app"])
        sf.grid(row=1, column=0, sticky="ew", padx=14, pady=(6, 4))
        tk.Label(sf, text=self._t("charm_picker_search"),
                 font=("Segoe UI", 9, "bold"), bg=COLORS["app"], fg=COLORS["text"],
                 ).pack(side=tk.LEFT)
        self._filt = tk.StringVar()
        ef = tk.Entry(sf, textvariable=self._filt, font=("Segoe UI", 9))
        ef.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8))
        tk.Label(sf, text=self._t("charm_picker_search_tip"),
                 font=("Segoe UI", 8), bg=COLORS["app"], fg=COLORS["muted"],
                 ).pack(side=tk.LEFT)

        # Hover tip (only when PIL is available so the hover actually works)
        if pil_ok:
            tk.Label(
                d, text=self._t("charm_picker_hover_tip"),
                font=("Segoe UI", 8), bg=COLORS["app"], fg=COLORS["muted"],
                wraplength=640, justify=tk.LEFT,
            ).grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 2))

        # Treeview
        tf = tk.Frame(d, bg=COLORS["app"])
        tf.grid(row=3, column=0, sticky="nsew", padx=14, pady=(2, 10))
        tf.grid_rowconfigure(0, weight=1)
        tf.grid_columnconfigure(0, weight=1)

        sty = ttk.Style(d)
        row_h = self._ROW_H if pil_ok else 24
        try:
            sty.configure("CharmPick.Treeview", rowheight=row_h, font=("Segoe UI", 9))
            sty.configure("CharmPick.Treeview.Heading", font=("Segoe UI", 9, "bold"))
            sty.map("CharmPick.Treeview", background=[("selected", COLORS["accent_soft"])])
        except tk.TclError:
            pass

        cols = ("code", "sku", "shop")
        tree = ttk.Treeview(
            tf, columns=cols, show="tree headings",
            selectmode="browse", style="CharmPick.Treeview",
        )
        vsb = ttk.Scrollbar(tf, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self._tree = tree

        tree.heading("#0",   text=self._t("charm_picker_col_photo"))
        tree.column ("#0",   width=80 if pil_ok else 20, stretch=False, anchor=tk.CENTER)
        tree.heading("code", text=self._t("charm_picker_col_code"))
        tree.column ("code", width=120, stretch=False, anchor=tk.CENTER)
        tree.heading("sku",  text=self._t("charm_picker_col_sku"))
        tree.column ("sku",  width=260, stretch=True)
        tree.heading("shop", text=self._t("charm_picker_col_shop"))
        tree.column ("shop", width=130, stretch=False)

        # Hover zoom state — must be initialised before _refill() which calls _hide_charm_hover()
        self._hover_after_id: list[object | None] = [None]
        self._hover_tip_win: list[tk.Toplevel | None] = [None]
        self._hover_photo: list[object] = []
        self._hover_active_iid: list[str | None] = [None]

        self._refill()
        self._filt.trace_add("write", lambda *_: self._refill())
        tree.bind("<Double-1>", lambda _e: self._pick())
        tree.bind("<Return>",   lambda _e: self._pick())

        if pil_ok:
            tree.bind("<Motion>", self._on_charm_hover_motion)
            tree.bind("<Leave>",  self._on_charm_hover_leave)

    def _t(self, key: str) -> str:
        return CHROME.get(self._lang, CHROME["en"]).get(key, key)

    def _thumb(self, raw: bytes | None) -> object | None:
        if not self._pil_ok or not raw or Image is None or ImageTk is None:
            return None
        try:
            im = Image.open(BytesIO(raw))
            if im.mode not in ("RGB", "RGBA"):
                im = im.convert("RGBA" if "A" in im.mode else "RGB")
            if im.mode == "RGBA":
                bg = Image.new("RGB", im.size, (255, 255, 255))
                bg.paste(im, mask=im.split()[3])
                im = bg
            im.thumbnail((self._THUMB, self._THUMB), Image.Resampling.LANCZOS)
            ph = ImageTk.PhotoImage(im)
            self._tk_img_refs.append(ph)
            return ph
        except Exception:
            return None

    def _refill(self) -> None:
        self._hide_charm_hover()
        self._tree.delete(*self._tree.get_children())
        self._tk_img_refs.clear()
        q = self._filt.get().strip().lower()

        # "(none)" row — lets the user clear the charm code
        if not q:
            self._tree.insert(
                "", tk.END, iid="__none__",
                text="—", values=("", self._t("charm_picker_none"), ""),
            )

        for code, entry in self._entries.items():
            if q:
                blob = f"{code} {entry.sku} {entry.default_charm_shop}".lower()
                if q not in blob:
                    continue
            thumb = self._thumb(entry.photo_bytes)
            kw: dict = dict(
                values=(code, entry.sku or "", entry.default_charm_shop or ""),
            )
            if thumb is not None:
                kw["image"] = thumb
                kw["text"]  = ""
            else:
                kw["text"] = "—"
            self._tree.insert("", tk.END, iid=f"ch_{code}", **kw)

    def _pick(self) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        iid = sel[0]
        if iid == "__none__":
            code = ""
        else:
            code = iid.removeprefix("ch_")
        self._hide_charm_hover()
        if self._callback and callable(self._callback):
            self._callback(code)
        self._d.destroy()

    # ── hover zoom ──

    def _hide_charm_hover(self) -> None:
        aid = self._hover_after_id[0]
        if aid is not None:
            try:
                self._d.after_cancel(aid)
            except (tk.TclError, ValueError):
                pass
            self._hover_after_id[0] = None
        self._hover_active_iid[0] = None
        tip = self._hover_tip_win[0]
        if tip is not None:
            try:
                tip.destroy()
            except tk.TclError:
                pass
            self._hover_tip_win[0] = None

    def _on_charm_hover_motion(self, _e: tk.Event | None = None) -> None:
        if not self._pil_ok:
            return
        tree = self._tree
        px, py = tree.winfo_pointerx(), tree.winfo_pointery()
        lx = px - tree.winfo_rootx()
        ly = py - tree.winfo_rooty()
        if lx < 0 or ly < 0 or lx >= tree.winfo_width() or ly >= tree.winfo_height():
            self._hide_charm_hover()
            return
        # Only activate when the cursor is over the Photo (tree icon) column.
        if tree.identify_column(lx) != "#0":
            self._hide_charm_hover()
            return
        iid = tree.identify_row(ly)
        if not iid or iid == "__none__":
            self._hide_charm_hover()
            return
        code = iid.removeprefix("ch_")
        entry = self._entries.get(code)
        if entry is None or not entry.photo_bytes:
            self._hide_charm_hover()
            return
        if iid == self._hover_active_iid[0]:
            return
        self._hide_charm_hover()
        self._hover_active_iid[0] = iid
        self._hover_after_id[0] = self._d.after(
            220, lambda i=iid: self._show_charm_hover(i)
        )

    def _show_charm_hover(self, iid: str) -> None:
        self._hover_after_id[0] = None
        if Image is None or ImageTk is None:
            return
        tree = self._tree
        px, py = tree.winfo_pointerx(), tree.winfo_pointery()
        lx, ly = px - tree.winfo_rootx(), py - tree.winfo_rooty()
        if lx < 0 or ly < 0 or lx >= tree.winfo_width() or ly >= tree.winfo_height():
            return
        if tree.identify_row(ly) != iid:
            return
        code = iid.removeprefix("ch_")
        entry = self._entries.get(code)
        if entry is None:
            return
        raw = entry.photo_bytes
        if not raw:
            return
        try:
            im = Image.open(BytesIO(raw))
            if im.mode not in ("RGB", "RGBA"):
                im = im.convert("RGBA" if "A" in im.mode else "RGB")
            if im.mode == "RGBA":
                bg = Image.new("RGB", im.size, (255, 255, 255))
                bg.paste(im, mask=im.split()[3])
                im = bg
            w, h = im.size
            max_dim = 400
            if w > max_dim or h > max_dim:
                ratio = min(max_dim / w, max_dim / h)
                im = im.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)
            ph = ImageTk.PhotoImage(im)
            self._hover_photo.clear()
            self._hover_photo.append(ph)
        except Exception:
            return
        tip = tk.Toplevel(self._d)
        tip.overrideredirect(True)
        try:
            tip.attributes("-topmost", True)
        except tk.TclError:
            pass
        tip.configure(bg="#ffffff", bd=2, highlightthickness=1,
                      highlightbackground="#cccccc")
        tk.Label(tip, image=ph, bg="#ffffff", bd=0,
                 highlightthickness=0, relief=tk.FLAT).pack(padx=2, pady=2)
        tip.update_idletasks()
        tw, th = tip.winfo_reqwidth(), tip.winfo_reqheight()
        sw, sh = tip.winfo_screenwidth(), tip.winfo_screenheight()
        x = min(max(12, px + 24), sw - tw - 12)
        y = min(max(12, py + 24), sh - th - 12)
        tip.geometry(f"+{x}+{y}")
        self._hover_tip_win[0] = tip

    def _on_charm_hover_leave(self, _e: tk.Event | None = None) -> None:
        self._hide_charm_hover()


class _CharmReorderDialog:
    """
    Modal dialog that lets the user drag-and-drop (or use arrow buttons) to
    reorder charms.  Calls ``reorder_charm_library_rows`` on Apply.

    Layout
    ------
    Row 0 — hero banner (title + instructions)
    Row 1 — main body: Treeview (left) + arrow buttons (right)
    Row 2 — footer: Close / Apply buttons
    """

    _ROW_H   = 80   # Treeview row height in px when Pillow is available
    _THUMB   = 68   # thumbnail size for Treeview column #0

    def __init__(
        self,
        parent: App,
        entries: dict,   # dict[str, CharmLibraryEntry]
        catalog_path: Path,
        charm_images_dir: Path,
    ) -> None:
        self._parent         = parent
        self._entries        = entries
        self._catalog_path   = catalog_path
        self._charm_images   = charm_images_dir
        self._lang           = parent._lang
        self._tk_img_refs: list[object] = []

        pil_ok = Image is not None and ImageTk is not None
        self._pil_ok = pil_ok

        d = tk.Toplevel(parent)
        self._d = d
        d.title(self._t("reorder_title"))
        d.transient(parent)
        d.grab_set()
        d.configure(bg=COLORS["app"])
        d.geometry("1020x680")
        d.minsize(760, 500)

        d.grid_columnconfigure(0, weight=1)
        d.grid_rowconfigure(1, weight=1)

        # ── Row 0: hero banner ────────────────────────────────────────────
        hero = tk.Frame(d, bg=COLORS["hero"], highlightthickness=0)
        hero.grid(row=0, column=0, sticky="ew")
        tk.Label(
            hero,
            text=self._t("reorder_heading"),
            font=("Segoe UI", 12, "bold"),
            fg="#ffffff", bg=COLORS["hero"],
        ).pack(anchor=tk.W, padx=14, pady=(10, 2))
        tk.Label(
            hero,
            text=self._t("reorder_intro"),
            wraplength=940, justify=tk.LEFT,
            font=("Segoe UI", 9), fg="#e2e8f0", bg=COLORS["hero"],
        ).pack(anchor=tk.W, padx=14, pady=(0, 8))

        # ── Row 1: body (tree + arrow buttons) ───────────────────────────
        body = tk.Frame(d, bg=COLORS["app"])
        body.grid(row=1, column=0, sticky="nsew", padx=14, pady=(10, 0))
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        tree_frame = tk.Frame(body, bg=COLORS["app"])
        tree_frame.grid(row=0, column=0, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Treeview styling
        sty = ttk.Style(d)
        row_h = self._ROW_H if pil_ok else 26
        try:
            sty.configure("Reorder.Treeview",
                          rowheight=row_h, font=("Segoe UI", 9))
            sty.configure("Reorder.Treeview.Heading",
                          font=("Segoe UI", 9, "bold"))
            sty.map("Reorder.Treeview",
                    background=[("selected", COLORS["accent_soft"])])
        except tk.TclError:
            pass

        cols = ("code", "new_code", "sku", "shop")
        tree = ttk.Treeview(
            tree_frame,
            columns=cols,
            show="tree headings",
            selectmode="browse",
            style="Reorder.Treeview",
        )
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self._tree = tree

        thumb_w = self._THUMB + 8 if pil_ok else 20
        tree.heading("#0",       text=self._t("reorder_col_photo"))
        tree.column ("#0",       width=thumb_w, stretch=False, anchor=tk.CENTER)
        tree.heading("code",     text=self._t("reorder_col_code"))
        tree.column ("code",     width=110, stretch=False, anchor=tk.CENTER)
        tree.heading("new_code", text=self._t("reorder_col_new"))
        tree.column ("new_code", width=110, stretch=False, anchor=tk.CENTER)
        tree.heading("sku",      text=self._t("reorder_col_sku"))
        tree.column ("sku",      width=220, stretch=True)
        tree.heading("shop",     text=self._t("reorder_col_shop"))
        tree.column ("shop",     width=130, stretch=False)

        # Arrow-button column (right of tree)
        btn_col = tk.Frame(body, bg=COLORS["app"])
        btn_col.grid(row=0, column=1, sticky="ns", padx=(10, 0), pady=(0, 0))

        btn_cfg = dict(style="Tool.TButton", width=14)
        ttk.Button(btn_col, text=self._t("reorder_btn_top"),
                   command=self._move_top, **btn_cfg).pack(pady=(0, 4))
        ttk.Button(btn_col, text=self._t("reorder_btn_up"),
                   command=self._move_up,  **btn_cfg).pack(pady=(0, 4))
        ttk.Button(btn_col, text=self._t("reorder_btn_down"),
                   command=self._move_down, **btn_cfg).pack(pady=(0, 4))
        ttk.Button(btn_col, text=self._t("reorder_btn_bottom"),
                   command=self._move_bottom, **btn_cfg).pack(pady=(0, 0))

        # ── Row 2: footer buttons ─────────────────────────────────────────
        foot = tk.Frame(d, bg=COLORS["app"])
        foot.grid(row=2, column=0, sticky="ew", padx=14, pady=12)
        close_btn = ttk.Button(
            foot, text=self._t("reorder_btn_close"),
            command=d.destroy, style="Tool.TButton",
        )
        close_btn.pack(side=tk.LEFT)
        self._apply_btn = tk.Button(
            foot,
            text=self._t("reorder_btn_apply"),
            command=self._apply,
            bg=COLORS["run"], fg="#ffffff",
            activebackground=COLORS["run_hover"],
            activeforeground="#ffffff",
            font=("Segoe UI", 10, "bold"),
            padx=18, pady=6,
            cursor="hand2",
            relief=tk.FLAT, borderwidth=0, highlightthickness=0,
        )
        self._apply_btn.pack(side=tk.RIGHT)

        # ── Populate tree ─────────────────────────────────────────────────
        self._iid_to_code: dict[str, str] = {}
        self._populate()

        # ── Drag-and-drop bindings ────────────────────────────────────────
        self._drag: dict[str, object] = {"item": None}
        tree.bind("<ButtonPress-1>",   self._on_drag_start)
        tree.bind("<B1-Motion>",       self._on_drag_motion)
        tree.bind("<ButtonRelease-1>", self._on_drag_release)

    # ------------------------------------------------------------------ helpers

    def _t(self, key: str, **kw: object) -> str:
        s = CHROME.get(self._lang, CHROME["en"]).get(key, key)
        return s.format(**kw) if kw else s

    def _thumb(self, raw: bytes | None) -> object | None:
        if not self._pil_ok or not raw or Image is None or ImageTk is None:
            return None
        try:
            im = Image.open(BytesIO(raw))
            if im.mode not in ("RGB", "RGBA"):
                im = im.convert("RGBA" if "A" in im.mode else "RGB")
            if im.mode == "RGBA":
                bg = Image.new("RGB", im.size, (255, 255, 255))
                bg.paste(im, mask=im.split()[3])
                im = bg
            im.thumbnail((self._THUMB, self._THUMB), Image.Resampling.LANCZOS)
            ph = ImageTk.PhotoImage(im)
            self._tk_img_refs.append(ph)
            return ph
        except Exception:
            return None

    def _codes_in_order(self) -> list[str]:
        return [self._iid_to_code[iid] for iid in self._tree.get_children()]

    def _new_code_for_pos(self, pos: int, total: int) -> str:
        """Sequential code string for 1-based *pos* out of *total*."""
        import re as _re
        all_codes = list(self._entries.keys())
        pfx = "CH-"
        for code in all_codes:
            m = _re.match(r"^([A-Za-z]+-+)(\d+)$", code.strip())
            if m:
                pfx = m.group(1).upper()
                break
        # Compute width
        widths = []
        for code in all_codes:
            m = _re.match(rf"^{_re.escape(pfx)}(\d+)$", code.strip(), _re.IGNORECASE)
            if m:
                widths.append(len(m.group(1)))
        w = max(widths) if widths else 5
        w = max(w, len(str(total)), 5)
        return f"{pfx}{pos:0{w}d}"

    def _refresh_new_codes(self) -> None:
        """Update the «New Code» column for every row based on current position."""
        items = self._tree.get_children()
        total = len(items)
        for pos, iid in enumerate(items, start=1):
            nc = self._new_code_for_pos(pos, total)
            self._tree.set(iid, "new_code", nc)

    def _populate(self) -> None:
        self._tree.delete(*self._tree.get_children())
        self._iid_to_code.clear()
        self._tk_img_refs.clear()
        total = len(self._entries)
        for pos, (code, entry) in enumerate(self._entries.items(), start=1):
            iid   = f"charm_{code}"
            thumb = self._thumb(entry.photo_bytes)
            nc    = self._new_code_for_pos(pos, total)
            kw: dict = dict(
                values=(code, nc, entry.sku or "", entry.default_charm_shop or ""),
            )
            if thumb is not None:
                kw["image"] = thumb
                kw["text"]  = ""
            else:
                kw["text"] = "—"
            self._tree.insert("", tk.END, iid=iid, **kw)
            self._iid_to_code[iid] = code

    # ------------------------------------------------------------------ movement

    def _move_up(self) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        item = sel[0]
        idx  = self._tree.index(item)
        if idx > 0:
            self._tree.move(item, "", idx - 1)
            self._refresh_new_codes()

    def _move_down(self) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        item = sel[0]
        idx  = self._tree.index(item)
        last = len(self._tree.get_children()) - 1
        if idx < last:
            self._tree.move(item, "", idx + 1)
            self._refresh_new_codes()

    def _move_top(self) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        self._tree.move(sel[0], "", 0)
        self._refresh_new_codes()

    def _move_bottom(self) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        last = len(self._tree.get_children()) - 1
        self._tree.move(sel[0], "", last)
        self._refresh_new_codes()

    # ------------------------------------------------------------------ drag-drop

    def _on_drag_start(self, event: tk.Event) -> None:
        item = self._tree.identify_row(event.y)
        if item:
            self._tree.selection_set(item)
            self._drag["item"] = item

    def _on_drag_motion(self, event: tk.Event) -> None:
        item = self._drag.get("item")
        if not item:
            return
        target = self._tree.identify_row(event.y)
        if not target or target == item:
            return
        drag_idx   = self._tree.index(item)
        target_idx = self._tree.index(target)
        # Determine insert-before or insert-after from mouse Y within target row
        bbox = self._tree.bbox(target, "#0")
        if bbox:
            mid = bbox[1] + bbox[3] // 2
            if event.y < mid:
                dest = target_idx if drag_idx > target_idx else target_idx - 1
            else:
                dest = target_idx if drag_idx < target_idx else target_idx + 1
        else:
            dest = target_idx
        self._tree.move(item, "", max(0, dest))
        self._refresh_new_codes()

    def _on_drag_release(self, _event: tk.Event) -> None:
        self._drag["item"] = None

    # ------------------------------------------------------------------ apply

    def _apply(self) -> None:
        if not messagebox.askyesno(
            self._t("reorder_confirm_title"),
            self._t("reorder_confirm_body"),
            parent=self._d,
        ):
            return

        new_order = self._codes_in_order()

        self._apply_btn.config(state=tk.DISABLED)
        prev_cursor = self._d.cget("cursor")
        self._d.config(cursor="watch")
        self._d.update_idletasks()

        try:
            n, lines = reorder_charm_library_rows(
                self._catalog_path,
                new_order,
                charm_images_dir=self._charm_images,
            )
            summary = "\n".join(lines)
            msg = f"{self._t('reorder_done')}\n\n{summary}"
            messagebox.showinfo(self._t("reorder_title"), msg, parent=self._d)
            # Refresh the tree to reflect the now-applied order (codes updated)
            try:
                updated = load_charm_library(self._catalog_path)
                self._entries = updated
            except Exception:
                pass
            self._populate()
        except Exception as e:
            messagebox.showerror(self._t("file_open_fail_title"), str(e), parent=self._d)
        finally:
            self._d.config(cursor=prev_cursor)
            self._apply_btn.config(state=tk.NORMAL)


def main() -> None:
    if not GENERATOR.is_file():
        print(f"Missing {GENERATOR}", file=sys.stderr)
        sys.exit(1)
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
