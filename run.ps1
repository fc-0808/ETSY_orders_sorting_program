# Run the shopping route generator from project root.
# Usage: .\run.ps1 [args]
#
# Simple desktop UI (for staff): .\run_ui.ps1
#
# === COMMANDS ===
#   --new-batch                     New PDFs from input/ (skips already processed)
#   --new-batch file1.pdf ...       Specific PDFs
#   --purge-purchased               After shopping: purge purchased / OOP (see script help)
#   --refresh-catalog               After editing supplier_catalog.xlsx — rebuild route
#   --rebuild-catalog               Re-sort catalog by floor / stall (also syncs Suppliers)
#   --sync-suppliers                Append unique Shop+Stall from Product Map → Suppliers sheet
#   --clear-product-map-charm-codes Clear Product Map col H (wrong entries; re-pick from Charm Library)
#   --list-catalog-backups          List timestamped copies under data\supplier_catalog_backups\
#   --restore-latest-catalog-backup Restore supplier_catalog.xlsx from newest backup (Excel closed)
#   --restore-catalog-backup PATH   Restore from a specific backup file (path or filename in backups folder)
#   --init-charm-shops              Add Charm Shops + Charm Library sheets; Product Map G / H / I
#
# === CHARM LIBRARY + PRODUCT MAP (professional workflow) ===
#
#   CHARM LIBRARY (master list of bead straps / charms)
#   -----------------------------------------------------
#   Sheet "Charm Library":  Photo (A), Charm Code (B), SKU (C), Default Charm Shop (D).
#   One row per *physical charm design* you carry.  Use C (SKU) for how *you* sort and
#   search (e.g. CHM-SANRIO-HK-...); B stays the stable ID for files and column H (Charm Code).
#
#   Disk photos (optional, recommended):  data\charm_images\<Charm Code>.png
#   These override embedded Excel photos when building the route.  New imports use
#   zero-padding (5+ digits when starting fresh; matches width of existing CH-xxx).
#
#   PRODUCT MAP (link each *product variant* to a charm + shop)
#   ------------------------------------------------------------
#   Discontinuing a product: run_ui.ps1 → «Mark product discontinued», or CLI:
#       .\run.ps1 --mark-product-discontinued "Title"
#       The product is moved to the «Discontinued Products» sheet (with timestamp + photo)
#       and fully removed from Product Map. Run --refresh-catalog after.
#   Column G "Charm Shop"  — which charm-market stall supplies this product's charm
#       (dropdown from "Charm Shops").  Often matches Library column D for that code.
#   Column H "Charm Code"  — pick the same code as Charm Library column B when this
#       listing uses that charm.  Then the shopping route + HTML *Charm* section shows
#       your library/disk photo (accurate strap reference).  Leave H empty if you prefer
#       the PDF listing photo only, or the charm is not catalogued yet.
#
#   After editing Library or Product Map:  .\run.ps1 --refresh-catalog
#   Override charm folder:  --charm-images-dir PATH
#
#   Import new screenshots (default pattern Screenshot*.png):
#     .\run.ps1 --import-charm-dry-run
#     .\run.ps1 --import-charm-images
#
#   AI SKU labels (OpenAI-compatible API: CHARM_VISION_API_KEY / OPENAI_API_KEY, or
#   local Ollama/LM Studio via CHARM_VISION_BASE_URL, e.g. http://127.0.0.1:11434/v1):
#     .\run.ps1 --import-charm-images --import-charm-vision-sku
#     .\run.ps1 --fill-charm-sku
#     .\run.ps1 --fill-charm-sku --fill-charm-sku-dry-run
#
#   JSON index (library + disk):
#     .\run.ps1 --export-charm-manifest
#
#   After catalog or image changes: .\run.ps1 --refresh-catalog  (or a normal route run)
#
# === CHINESE / HTML ===
#   --chinese                       shopping_route_zh.xlsx
#   --chinese-exclude-shops A,B
#   --html                          shopping_route.html (mobile-friendly)
#
# === OTHER ===
#   --threshold N                   Fuzzy match cutoff (default 65)
#   --no-catalog-update             Do not append new rows to catalog
#   --reset                         Rebuild from PDFs only (ignore cache)
#
# === EXAMPLES ===
#   .\run.ps1 --new-batch --chinese --html
#   .\run.ps1 --refresh-catalog --chinese --html
#   .\run.ps1 --init-charm-shops
#   .\run.ps1 --export-charm-manifest

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot
python src/generate_shopping_route.py --project-dir . @args
