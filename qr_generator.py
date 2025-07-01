"""
qr_generator.py
Core utilities for producing QR codes and dropping them into a
Microsoft Word document (.docx).

You can import `create_qr_doc()` from another script (GUI, tests, etc.)
or treat this file as a CLI:

    python qr_generator.py 1000 1200

That would create “QR_1000_1200.docx”.
"""

from __future__ import annotations

import argparse
from io import BytesIO
from typing import Union
from pathlib import Path
from typing import Iterable

import qrcode
from qrcode.constants import ERROR_CORRECT_L
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Mm, Pt, Inches


# ──────────────────────────────────────────────────────────────────────────────
# Low-level helpers
# ──────────────────────────────────────────────────────────────────────────────
def _qr_png_stream(data: Union[int, str]) -> BytesIO:
    """Return an in-memory *PNG* stream for a single QR code.

    The QSize is hard-coded to match the original `box_size=5, border=2` preset.
    Feel free to expose these as parameters if you need more flexibility.
    """
    qr = qrcode.QRCode(
        version=1,
        error_correction=ERROR_CORRECT_L,
        box_size=5,
        border=2,
    )

    qr.add_data(data)
    qr.make(fit=False)

    img = qr.make_image(fill_color="black", back_color="white")
    buffer = BytesIO()

    img.save(buffer, format="PNG")
    buffer.seek(0)

    return buffer


def _chunk_range(
    start: int, end_exclusive: int, size: int = 100
) -> Iterable[tuple[int, int]]:
    """
    Yield (chunk_start, chunk_end_inclusive) for `range(start, end_exclusive)`
    in blocks of `size`. Works even if the last chunk is < size.
    """
    cur = start
    while cur < end_exclusive:
        # The number printed *inside* the QR is cur … min(cur+size-1, end-1)
        yield cur, min(cur + size - 1, end_exclusive - 1)
        cur += size


def _new_doc_a4_margins() -> Document:
    """Return a fresh A4 document with 0.5″ margins."""
    doc = Document()
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    for sec in doc.sections:
        margin = Inches(0.5)
        sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = margin
    return doc


# ──────────────────────────────────────────────────────────────────────────────
# Public API
# ──────────────────────────────────────────────────────────────────────────────
def add_qr_block(doc: Document, start_num: int, end_num: int, cols: int = 17) -> None:
    """Insert QR codes for *start_num … end_num* (inclusive) into *doc*.

    The codes are laid out in a single table with *cols* columns.  An alignment
    label (e.g. "101-200") is placed to the immediate right of the last code
    when room permits; otherwise it shares that cell.
    """
    if end_num < start_num:
        raise ValueError("end_num must be ≥ start_num")

    total = end_num - start_num + 1
    rows = -(-total // cols)  # ceiling division without math.ceil()

    table = doc.add_table(rows=rows, cols=cols)
    table.autofit = False

    for idx, number in enumerate(range(start_num, end_num + 1)):
        r, c = divmod(idx, cols)
        cell = table.cell(r, c)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Generate QR picture *once* and hand the stream straight to python‑docx
        p.add_run().add_picture(_qr_png_stream(number), width=Mm(9))

    # ── trailing range label ────────────────────────────────────────────────
    r, c = divmod(total - 1, cols)
    label_cell = table.cell(r, c + 1 if c < cols - 1 else c)
    label_para = label_cell.paragraphs[0]
    label_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = label_para.add_run(f"{start_num}-{end_num}")
    run.font.size = Pt(8)

    # spacer lines after each block
    doc.add_paragraph()
    doc.add_paragraph()


def create_qr_doc(start: int, end_exclusive: int, out_path: str | Path) -> Path:
    """
    Build an A4 .docx file with QR codes for *all* numbers in
    `range(start, end_exclusive)`.  Returns the path of the file created.
    """
    doc = _new_doc_a4_margins()
    for chunk_start, chunk_end in _chunk_range(start, end_exclusive):
        add_qr_block(doc, chunk_start, chunk_end)

    out_path = Path(out_path)
    doc.save(out_path)
    return out_path


# ──────────────────────────────────────────────────────────────────────────────
# CLI entry-point – optional, keeps parity with your original script
# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate QR-code .docx files.")
    parser.add_argument("start", type=int, help="First integer (inclusive).")
    parser.add_argument("end", type=int, help="Last integer (exclusive).")

    args = parser.parse_args()
    fname = f"QR_{args.start}_{args.end}.docx"

    create_qr_doc(args.start, args.end, fname)
    print(f"Saved → {fname}")
