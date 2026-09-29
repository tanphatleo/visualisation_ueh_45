"""
QR code for the public deck repository, for the title slide.

    python build_qr.py

Error correction is M (15%), not H. H would push this URL to version 6 — 41
modules across — and at the size the title slide gives it, the modules land
below what a phone camera resolves. M keeps it to version 4 (33 modules), so
each module is a third larger. Drawn in the deck's accent blue on white, with
the standard four-module quiet zone intact: crop that and phones stop seeing it.
"""

from pathlib import Path

import qrcode
from qrcode.constants import ERROR_CORRECT_M

URL = "https://github.com/tanphatleo/visualisation_ueh_45"
ACCENT = "#0F5499"
HERE = Path(__file__).resolve().parent
OUT = HERE / "images" / "repo_qr.png"


def main() -> None:
    qr = qrcode.QRCode(error_correction=ERROR_CORRECT_M, box_size=24, border=4)
    qr.add_data(URL)
    qr.make(fit=True)
    img = qr.make_image(fill_color=ACCENT, back_color="white").convert("RGB")
    OUT.parent.mkdir(parents=True, exist_ok=True)
    img.save(OUT)
    print(f"{OUT.relative_to(HERE)}  {img.size[0]}x{img.size[1]}  version {qr.version}")


if __name__ == "__main__":
    main()
