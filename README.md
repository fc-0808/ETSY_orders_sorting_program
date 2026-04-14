# Etsy Orders Sorting Program

Generates shopping routes from Etsy order PDFs, organized by supplier location for efficient in-person shopping.

## Quick Start

1. **Drop new PDFs** into `input/`
2. **Run** `.\run.ps1 --new-batch`
3. **Open** `output/shopping_route.xlsx`

## Folder Layout

| Folder | Purpose |
|--------|---------|
| `src/` | Python scripts |
| `data/` | `supplier_catalog.xlsx` — edit supplier info here |
| `input/` | Order PDFs to process |
| `output/` | Generated `shopping_route.xlsx`, `.html` |
| `cache/` | Runtime caches (orders, translations, OOP log) |
| `docs/` | Instructions |

## Common Commands

```powershell
.\run.ps1 --new-batch              # Process new PDFs from input/
.\run.ps1 --purge-purchased        # Clean up after shopping trip
.\run.ps1 --refresh-catalog        # Rebuild after editing supplier_catalog.xlsx
.\run.ps1 --refresh-catalog --chinese --html   # Chinese + HTML versions
```

See `docs/instructions.txt` for the full workflow.

## Version control (Git)

This repository tracks **source code and docs** only. Paths listed in `.gitignore` stay on your machine: order PDFs (`input/` and `backup/`), supplier workbook and other Excel under `data/`, charm image binaries, `cache/`, `output/`, and `data/charm_manifest.json` (regenerated; may contain local absolute paths).

After cloning on another PC, copy in your own `supplier_catalog.xlsx`, charm images, and PDFs, then run as usual.
