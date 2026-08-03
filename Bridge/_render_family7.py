"""Render Family 7 PDF pages for visual table confirmation."""
from __future__ import annotations

import pathlib
import fitz

ROOT = pathlib.Path(__file__).resolve().parent.parent
CAP = ROOT / "Bridge" / "captures"


def render(pdf_name: str, zoom: float = 1.5):
    doc = fitz.open(str(CAP / pdf_name))
    stem = pdf_name.replace(".pdf", "").replace("-", "")
    for i, pg in enumerate(doc):
        pix = pg.get_pixmap(matrix=fitz.Matrix(zoom, zoom), annots=False)
        out = CAP / f"sheet_{stem}_p{i+1}.png"
        pix.save(str(out))
        print(f"saved {out.name} {pix.width}x{pix.height} rot={pg.rotation}")


if __name__ == "__main__":
    render("619-110.pdf", zoom=2.0)
    render("619-111.pdf", zoom=1.5)
    render("619-112.pdf", zoom=1.5)
