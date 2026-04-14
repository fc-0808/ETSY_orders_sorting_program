import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import pdfplumber
from pathlib import Path

for fname in sorted(Path(".").glob("*.pdf")):
    print(f"\n=== {fname} ===")
    with pdfplumber.open(fname) as pdf:
        for pi, page in enumerate(pdf.pages):
            images = page.images
            print(f"  Page {pi}: {len(images)} image(s)")
            for ii, img in enumerate(images):
                print(f"    [{ii}] x0={img['x0']:.1f} y0={img['y0']:.1f} "
                      f"x1={img['x1']:.1f} y1={img['y1']:.1f} "
                      f"w={img['width']} h={img['height']} "
                      f"colorspace={img.get('colorspace')} "
                      f"bits={img.get('bits')} "
                      f"filter={img.get('filter')} "
                      f"name={img.get('name')}")
