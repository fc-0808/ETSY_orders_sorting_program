#!/usr/bin/env python3
"""
migrate_to_sqlite.py

One-time migration: seeds ``data/etsy_orders.db`` from:
  • data/supplier_catalog.xlsx  (Product Map + Charm Library + Charm Shops)
  • cache/orders_cache.json     (resolved orders + processed PDF filenames)

Run once from the project root:
    python src/migrate_to_sqlite.py
    python src/migrate_to_sqlite.py --project-dir /path/to/project

Options:
    --project-dir DIR   Project root (default: cwd)
    --db PATH           Override database path
                        (default: <project-dir>/data/etsy_orders.db)
    --no-catalog        Skip catalog / charm tables
    --no-cache          Skip orders / processed_pdfs tables
    --rebuild-fts       Force-rebuild the FTS5 index after migration
                        (safe to run any time; equivalent to
                        INSERT INTO catalog_fts(catalog_fts) VALUES('rebuild'))

The script is idempotent: re-running it uses INSERT … ON CONFLICT DO UPDATE
for catalog/charm rows and INSERT OR IGNORE for orders/items/pdfs so
duplicate runs do not corrupt data.
"""

from __future__ import annotations

import argparse
import base64
import json
import logging
import sqlite3
import sys
from pathlib import Path

import openpyxl

# ---------------------------------------------------------------------------
# Make generate_shopping_route importable so we can reuse
# extract_photos_from_xlsx and the sheet-name constants.
# ---------------------------------------------------------------------------
_SRC = Path(__file__).parent
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from generate_shopping_route import (  # noqa: E402  (after sys.path patch)
    CATALOG_SHEET,
    CHARM_LIBRARY_SHEET,
    CHARM_SHOPS_SHEET,
    extract_photos_from_xlsx,
)

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Schema
# ---------------------------------------------------------------------------

_SCHEMA_SQL = """\
-- ── Catalog (Product Map) ──────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS catalog (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    product_title TEXT    UNIQUE NOT NULL,
    shop_name     TEXT    NOT NULL DEFAULT '',
    stall         TEXT    NOT NULL DEFAULT '',
    price         TEXT    NOT NULL DEFAULT '',
    charm_shop    TEXT    NOT NULL DEFAULT '',
    charm_code    TEXT    NOT NULL DEFAULT '',
    notes         TEXT    NOT NULL DEFAULT '',
    photo         BLOB
);

-- ── Charm Library ─────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS charm_library (
    code               TEXT PRIMARY KEY,
    sku                TEXT NOT NULL DEFAULT '',
    default_charm_shop TEXT NOT NULL DEFAULT '',
    notes              TEXT NOT NULL DEFAULT '',
    photo              BLOB
);

-- ── Charm Shops reference ─────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS charm_shops (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    shop_name TEXT UNIQUE NOT NULL,
    stall     TEXT NOT NULL DEFAULT '',
    notes     TEXT NOT NULL DEFAULT ''
);

-- ── Orders (header) ───────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS orders (
    order_number     TEXT PRIMARY KEY,
    etsy_shop        TEXT NOT NULL DEFAULT '',
    buyer_name       TEXT NOT NULL DEFAULT '',
    buyer_username   TEXT NOT NULL DEFAULT '',
    ship_to_name     TEXT NOT NULL DEFAULT '',
    ship_to_country  TEXT NOT NULL DEFAULT '',
    order_date       TEXT NOT NULL DEFAULT '',
    private_notes    TEXT NOT NULL DEFAULT ''
);

-- ── Order line items ──────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS order_items (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    order_number TEXT    NOT NULL REFERENCES orders(order_number) ON DELETE CASCADE,
    title        TEXT    NOT NULL,
    quantity     INTEGER NOT NULL DEFAULT 1,
    phone_model  TEXT    NOT NULL DEFAULT '',
    style        TEXT    NOT NULL DEFAULT '',
    photo_bytes  BLOB
);
CREATE INDEX IF NOT EXISTS idx_order_items_order ON order_items(order_number);

-- ── Processed PDFs (new-batch dedup) ─────────────────────────────────────
CREATE TABLE IF NOT EXISTS processed_pdfs (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    filename     TEXT UNIQUE NOT NULL,
    processed_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- ── FTS5 trigram index for fast product-title candidate lookup ────────────
CREATE VIRTUAL TABLE IF NOT EXISTS catalog_fts USING fts5(
    product_title,
    content='catalog',
    content_rowid='id',
    tokenize='trigram'
);

-- ── Triggers: keep catalog_fts in sync with catalog ──────────────────────
CREATE TRIGGER IF NOT EXISTS catalog_ai AFTER INSERT ON catalog BEGIN
    INSERT INTO catalog_fts(rowid, product_title)
        VALUES (new.id, new.product_title);
END;

CREATE TRIGGER IF NOT EXISTS catalog_ad AFTER DELETE ON catalog BEGIN
    INSERT INTO catalog_fts(catalog_fts, rowid, product_title)
        VALUES ('delete', old.id, old.product_title);
END;

CREATE TRIGGER IF NOT EXISTS catalog_au AFTER UPDATE ON catalog BEGIN
    INSERT INTO catalog_fts(catalog_fts, rowid, product_title)
        VALUES ('delete', old.id, old.product_title);
    INSERT INTO catalog_fts(rowid, product_title)
        VALUES (new.id, new.product_title);
END;
"""


# ---------------------------------------------------------------------------
# Database initialisation
# ---------------------------------------------------------------------------

def init_db(db_path: Path) -> sqlite3.Connection:
    """Create (or open) the database and apply the schema.

    Safe to call repeatedly — all DDL statements use IF NOT EXISTS / OR IGNORE.
    Returns an open Connection with WAL mode and foreign-key enforcement active.
    """
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    # executescript commits any open transaction before running, then auto-
    # commits after each statement — safe for DDL-only scripts.
    conn.executescript(_SCHEMA_SQL)
    conn.commit()
    log.info("Database ready at %s", db_path)
    return conn


# ---------------------------------------------------------------------------
# Catalog migration  (Product Map sheet)
# ---------------------------------------------------------------------------

def migrate_catalog(conn: sqlite3.Connection, catalog_path: Path) -> int:
    """Read every non-header row from the Product Map sheet and upsert into
    the ``catalog`` table.  Embedded product photos (column A) are stored as
    BLOB.

    Returns the number of rows inserted or updated.
    """
    if not catalog_path.exists():
        log.warning("Catalog not found at %s — skipping product migration", catalog_path)
        return 0

    # --- Extract embedded photos via the ZIP/XML approach ----------------
    try:
        row_photos: dict[int, bytes] = extract_photos_from_xlsx(
            catalog_path, sheet_name=CATALOG_SHEET, photo_col_idx=0
        )
        log.info("Extracted %d photo(s) from Product Map", len(row_photos))
    except Exception as exc:
        log.warning("Product Map photo extraction skipped: %s", exc)
        row_photos = {}

    # --- Read cell data ---------------------------------------------------
    wb = openpyxl.load_workbook(catalog_path, read_only=True, data_only=True)
    if CATALOG_SHEET not in wb.sheetnames:
        wb.close()
        log.warning("Sheet '%s' not found in %s", CATALOG_SHEET, catalog_path.name)
        return 0

    ws = wb[CATALOG_SHEET]

    # Detect layout: 8-column (current) vs. legacy with a CATEGORY column
    h3 = str(ws.cell(1, 3).value or "").strip().lower()
    h7 = str(ws.cell(1, 7).value or "").strip().lower()
    has_category       = h3 == "category"
    legacy_notes_first = has_category and h7 == "notes"

    cur = conn.cursor()
    rows_affected = 0

    for r_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        # Column B (index 1) is the product title
        title = row[1]
        if not title or not isinstance(title, str):
            continue
        title = title.strip()
        if title.startswith("TOTAL:") or title == "Unknown Product":
            continue

        if has_category:
            # Legacy layout: A photo, B title, C category, D shop, E stall,
            #                F price, G charm_shop|notes, H charm_code|charm_shop, ...
            shop_v  = str(row[3]).strip() if len(row) > 3 and row[3] is not None else ""
            stall_v = str(row[4]).strip() if len(row) > 4 and row[4] is not None else ""
            price_v = str(row[5]).strip() if len(row) > 5 and row[5] is not None else ""
            if legacy_notes_first:
                notes      = str(row[6]).strip() if len(row) > 6 and row[6] is not None else ""
                charm_shop = str(row[7]).strip() if len(row) > 7 and row[7] is not None else ""
                charm_code = str(row[8]).strip() if len(row) > 8 and row[8] is not None else ""
            else:
                charm_shop = str(row[6]).strip() if len(row) > 6 and row[6] is not None else ""
                charm_code = str(row[7]).strip() if len(row) > 7 and row[7] is not None else ""
                notes      = str(row[8]).strip() if len(row) > 8 and row[8] is not None else ""
        else:
            # Current 8-column layout:
            # A photo | B title | C shop | D stall | E price | F charm_shop | G charm_code | H notes
            shop_v     = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
            stall_v    = str(row[3]).strip() if len(row) > 3 and row[3] is not None else ""
            price_v    = str(row[4]).strip() if len(row) > 4 and row[4] is not None else ""
            charm_shop = str(row[5]).strip() if len(row) > 5 and row[5] is not None else ""
            charm_code = str(row[6]).strip() if len(row) > 6 and row[6] is not None else ""
            notes      = str(row[7]).strip() if len(row) > 7 and row[7] is not None else ""

        photo: bytes | None = row_photos.get(r_num)

        cur.execute(
            """
            INSERT INTO catalog
                (product_title, shop_name, stall, price,
                 charm_shop, charm_code, notes, photo)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(product_title) DO UPDATE SET
                shop_name  = excluded.shop_name,
                stall      = excluded.stall,
                price      = excluded.price,
                charm_shop = excluded.charm_shop,
                charm_code = excluded.charm_code,
                notes      = excluded.notes,
                -- Keep existing photo if the new row has no photo
                photo      = COALESCE(excluded.photo, catalog.photo)
            """,
            (title, shop_v, stall_v, price_v, charm_shop, charm_code, notes, photo),
        )
        rows_affected += 1

    wb.close()
    conn.commit()
    log.info("Product Map: %d row(s) upserted", rows_affected)
    return rows_affected


# ---------------------------------------------------------------------------
# Charm Library migration
# ---------------------------------------------------------------------------

def migrate_charm_library(conn: sqlite3.Connection, catalog_path: Path) -> int:
    """Read the Charm Library sheet and upsert into ``charm_library``.

    Column layout: A photo | B code | C SKU | D default_charm_shop | E notes
    """
    if not catalog_path.exists():
        return 0

    try:
        row_photos: dict[int, bytes] = extract_photos_from_xlsx(
            catalog_path, sheet_name=CHARM_LIBRARY_SHEET, photo_col_idx=0
        )
        log.info("Extracted %d photo(s) from Charm Library", len(row_photos))
    except Exception as exc:
        log.warning("Charm Library photo extraction skipped: %s", exc)
        row_photos = {}

    wb = openpyxl.load_workbook(catalog_path, read_only=True, data_only=True)
    if CHARM_LIBRARY_SHEET not in wb.sheetnames:
        wb.close()
        log.info("No '%s' sheet — skipping charm library", CHARM_LIBRARY_SHEET)
        return 0

    ws = wb[CHARM_LIBRARY_SHEET]
    cur = conn.cursor()
    rows_affected = 0

    for r_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        code = str(row[1]).strip() if len(row) > 1 and row[1] else ""
        if not code or code.lower() in ("charm code", "code"):
            continue

        sku      = str(row[2]).strip() if len(row) > 2 and row[2] else ""
        def_shop = str(row[3]).strip() if len(row) > 3 and row[3] else ""
        notes    = str(row[4]).strip() if len(row) > 4 and row[4] else ""
        photo: bytes | None = row_photos.get(r_num)

        cur.execute(
            """
            INSERT INTO charm_library (code, sku, default_charm_shop, notes, photo)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(code) DO UPDATE SET
                sku                = excluded.sku,
                default_charm_shop = excluded.default_charm_shop,
                notes              = excluded.notes,
                photo              = COALESCE(excluded.photo, charm_library.photo)
            """,
            (code, sku, def_shop, notes, photo),
        )
        rows_affected += 1

    wb.close()
    conn.commit()
    log.info("Charm Library: %d charm(s) upserted", rows_affected)
    return rows_affected


# ---------------------------------------------------------------------------
# Charm Shops migration
# ---------------------------------------------------------------------------

def migrate_charm_shops(conn: sqlite3.Connection, catalog_path: Path) -> int:
    """Read the Charm Shops sheet and upsert into ``charm_shops``.

    Column layout: A shop_name | B stall | C notes
    Rows without both name AND stall are skipped (instructional rows).
    """
    if not catalog_path.exists():
        return 0

    wb = openpyxl.load_workbook(catalog_path, read_only=True, data_only=True)
    if CHARM_SHOPS_SHEET not in wb.sheetnames:
        wb.close()
        log.info("No '%s' sheet — skipping charm shops", CHARM_SHOPS_SHEET)
        return 0

    ws = wb[CHARM_SHOPS_SHEET]
    cur = conn.cursor()
    rows_affected = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        name  = str(row[0] or "").strip() if row[0]               else ""
        stall = str(row[1] or "").strip() if len(row) > 1 and row[1] else ""
        notes = str(row[2] or "").strip() if len(row) > 2 and row[2] else ""
        if not name or not stall:
            continue

        cur.execute(
            """
            INSERT INTO charm_shops (shop_name, stall, notes)
            VALUES (?, ?, ?)
            ON CONFLICT(shop_name) DO UPDATE SET
                stall = excluded.stall,
                notes = excluded.notes
            """,
            (name, stall, notes),
        )
        rows_affected += 1

    wb.close()
    conn.commit()
    log.info("Charm Shops: %d shop(s) upserted", rows_affected)
    return rows_affected


# ---------------------------------------------------------------------------
# Orders cache migration  (orders_cache.json)
# ---------------------------------------------------------------------------

def migrate_orders_cache(
    conn: sqlite3.Connection, cache_path: Path
) -> tuple[int, int, int]:
    """Read ``orders_cache.json`` and insert into ``orders``, ``order_items``,
    and ``processed_pdfs``.

    The JSON has this top-level structure (from ``_resolved_to_dict``)::

        {
            "processed_pdfs": ["file1.pdf", ...],
            "items": [
                {
                    "order_number": "...",
                    "etsy_shop": "...",
                    "buyer_name": "...",
                    ...
                    "title": "...",
                    "quantity": 1,
                    "phone_model": "...",
                    "style": "...",
                    "photo_b64": "...",   ← base64-encoded JPEG or null
                    ...
                },
                ...
            ]
        }

    Each dict in ``items`` is a *flattened* ResolvedItem: one row per
    order-line-item, order header fields repeated for each item.

    Returns ``(orders_inserted, items_inserted, pdfs_inserted)``.
    """
    if not cache_path.exists():
        log.info("No cache file at %s — skipping orders migration", cache_path)
        return 0, 0, 0

    try:
        data = json.loads(cache_path.read_text(encoding="utf-8"))
    except Exception as exc:
        log.warning("Could not read cache (%s) — skipping", exc)
        return 0, 0, 0

    cur = conn.cursor()

    # --- 1. Processed PDF filenames ----------------------------------------
    pdfs_inserted = 0
    for filename in data.get("processed_pdfs", []):
        if not filename:
            continue
        cur.execute(
            "INSERT OR IGNORE INTO processed_pdfs (filename) VALUES (?)",
            (str(filename),),
        )
        pdfs_inserted += cur.rowcount

    # --- 2. Orders (header) and line items ---------------------------------
    orders_seen: set[str] = set()
    orders_inserted = 0
    items_inserted = 0

    for rec in data.get("items", []):
        order_number = str(rec.get("order_number") or "").strip()
        if not order_number:
            log.debug("Skipping cache record with no order_number: %r", rec)
            continue

        # Insert the order header once per unique order_number
        if order_number not in orders_seen:
            cur.execute(
                """
                INSERT INTO orders
                    (order_number, etsy_shop, buyer_name, buyer_username,
                     ship_to_name, ship_to_country, order_date, private_notes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(order_number) DO NOTHING
                """,
                (
                    order_number,
                    rec.get("etsy_shop",        ""),
                    rec.get("buyer_name",        ""),
                    rec.get("buyer_username",    ""),
                    rec.get("ship_to_name",      ""),
                    rec.get("ship_to_country",   ""),
                    rec.get("order_date",        ""),
                    rec.get("private_notes",     ""),
                ),
            )
            if cur.rowcount:
                orders_inserted += 1
            orders_seen.add(order_number)

        # Decode the product photo from base64 → raw bytes for BLOB storage
        photo_b64: str | None = rec.get("photo_b64")
        photo_bytes: bytes | None = None
        if photo_b64:
            try:
                photo_bytes = base64.b64decode(photo_b64)
            except Exception as exc:
                log.debug("Photo decode failed for order %s: %s", order_number, exc)

        cur.execute(
            """
            INSERT INTO order_items
                (order_number, title, quantity, phone_model, style, photo_bytes)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                order_number,
                rec.get("title",       ""),
                rec.get("quantity",    1),
                rec.get("phone_model", ""),
                rec.get("style",       ""),
                photo_bytes,
            ),
        )
        items_inserted += 1

    conn.commit()
    log.info(
        "Cache migration: %d order(s), %d item(s), %d PDF filename(s)",
        orders_inserted, items_inserted, pdfs_inserted,
    )
    return orders_inserted, items_inserted, pdfs_inserted


# ---------------------------------------------------------------------------
# FTS5 index rebuild (idempotent, safe any time)
# ---------------------------------------------------------------------------

def rebuild_fts_index(conn: sqlite3.Connection) -> None:
    """Force a full rebuild of the catalog_fts trigram index.

    The triggers keep the index in sync during normal INSERT/UPDATE/DELETE, but
    running this once after a bulk migration ensures the index is pristine and
    any edge-case gaps are closed.
    """
    log.info("Rebuilding FTS5 trigram index …")
    conn.execute("INSERT INTO catalog_fts(catalog_fts) VALUES('rebuild')")
    conn.commit()
    log.info("FTS5 index rebuild complete")


# ---------------------------------------------------------------------------
# Smoke-test: verify row counts and a sample FTS query
# ---------------------------------------------------------------------------

def _smoke_test(conn: sqlite3.Connection) -> None:
    cur = conn.cursor()
    tables = ["catalog", "charm_library", "charm_shops", "orders", "order_items", "processed_pdfs"]
    print("\n── Row counts ──────────────────────────────────────")
    for tbl in tables:
        (n,) = cur.execute(f"SELECT COUNT(*) FROM {tbl}").fetchone()  # noqa: S608
        print(f"  {tbl:<20} {n:>6} rows")

    # FTS5 sanity check: retrieve top 3 catalog titles
    print("\n── FTS5 sample (top 3 product titles) ─────────────")
    rows = cur.execute(
        "SELECT product_title FROM catalog_fts LIMIT 3"
    ).fetchall()
    if rows:
        for (title,) in rows:
            print(f"  {title}")
    else:
        print("  (catalog is empty or FTS index has no rows)")
    print()


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument(
        "--project-dir", metavar="DIR", default=".",
        help="Project root directory (default: current working directory).",
    )
    p.add_argument(
        "--db", metavar="PATH",
        help=(
            "Override the database path "
            "(default: <project-dir>/data/etsy_orders.db)."
        ),
    )
    p.add_argument(
        "--no-catalog", action="store_true",
        help="Skip Product Map, Charm Library, and Charm Shops migration.",
    )
    p.add_argument(
        "--no-cache", action="store_true",
        help="Skip orders_cache.json migration.",
    )
    p.add_argument(
        "--rebuild-fts", action="store_true",
        help="Force-rebuild the FTS5 trigram index after migration.",
    )
    p.add_argument(
        "--smoke-test", action="store_true",
        help="Print row counts and a sample FTS query after migration.",
    )
    return p


def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        datefmt="%H:%M:%S",
    )

    args = _build_parser().parse_args()
    project_dir = Path(args.project_dir).resolve()
    db_path     = (
        Path(args.db).resolve() if args.db
        else project_dir / "data" / "etsy_orders.db"
    )
    catalog_path = project_dir / "data" / "supplier_catalog.xlsx"
    cache_path   = project_dir / "cache" / "orders_cache.json"

    log.info("──────────────────────────────────────────")
    log.info("Project dir  : %s", project_dir)
    log.info("Database     : %s", db_path)
    log.info("Catalog      : %s", catalog_path)
    log.info("Cache        : %s", cache_path)
    log.info("──────────────────────────────────────────")

    conn = init_db(db_path)

    if not args.no_catalog:
        migrate_catalog(conn, catalog_path)
        migrate_charm_library(conn, catalog_path)
        migrate_charm_shops(conn, catalog_path)

    if not args.no_cache:
        migrate_orders_cache(conn, cache_path)

    if args.rebuild_fts:
        rebuild_fts_index(conn)

    if args.smoke_test:
        _smoke_test(conn)

    conn.close()
    log.info("Migration complete. Database written to %s", db_path)


if __name__ == "__main__":
    main()
