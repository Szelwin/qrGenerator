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
import os
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
def _generate_qr_png(data: int | str, out_path: Path) -> None:
    """Create a tiny temporary PNG for *one* QR code."""
    qr = qrcode.QRCode(version=1,
                       error_correction=ERROR_CORRECT_L,
                       box_size=5,
                       border=2)
    qr.add_data(data)
    qr.make(fit=False)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save(out_path)


def _chunk_range(start: int, end_exclusive: int, size: int = 100) -> Iterable[tuple[int, int]]:
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
    """
    Insert QR codes for the numbers start_num … end_num (inclusive) into `doc`
    as one table.  Works for any count ≤ 100; the GUI/CLI use 100-sized blocks.
    """
    total = end_num - start_num + 1
    rows = (total + cols - 1) // cols  # ceiling division
    table = doc.add_table(rows=rows, cols=cols)
    table.autofit = False

    tmp_png = Path("__qr_tmp.png")  # reused & deleted each loop
    for idx, number in enumerate(range(start_num, end_num + 1)):
        _generate_qr_png(number, tmp_png)
        r, c = divmod(idx, cols)
        cell = table.cell(r, c)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(tmp_png), width=Mm(9))
    tmp_png.unlink(missing_ok=True)  # clean up

    # Small “###-###” label to the immediate right of the last QR
    r, c = divmod(total - 1, cols)
    if c < cols - 1:
        label_cell = table.cell(r, c + 1)
    else:
        label_cell = table.cell(r, c)
    run = label_cell.paragraphs[0].add_run(f"{start_num}-{end_num}")
    run.font.size = Pt(8)
    label_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

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
