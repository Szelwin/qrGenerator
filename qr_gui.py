"""
gui_app.py
==========

Tiny Tkinter shell around qr_document.build_and_save().
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from qr_generator import create_qr_doc


class QRApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("QR Code Batch Generator")
        self._make_widgets()
        self.resizable(False, False)

    # ──────────────────────────────────────────────────────────────────────
    # UI layout
    # ──────────────────────────────────────────────────────────────────────
    def _make_widgets(self):
        paddings = dict(padx=8, pady=4)

        ttk.Label(self, text="Start number (inclusive)").grid(
            row=0, column=0, **paddings, sticky="w"
        )
        self.start_entry = ttk.Entry(self, width=20)
        self.start_entry.grid(row=0, column=1, **paddings)

        ttk.Label(self, text="End number (exclusive)").grid(
            row=1, column=0, **paddings, sticky="w"
        )
        self.end_entry = ttk.Entry(self, width=20)
        self.end_entry.grid(row=1, column=1, **paddings)

        self.generate_btn = ttk.Button(self, text="Generate", command=self._on_generate)
        self.generate_btn.grid(row=2, column=0, columnspan=2, pady=(10, 4))

    # ──────────────────────────────────────────────────────────────────────
    # Callbacks
    # ──────────────────────────────────────────────────────────────────────
    def _on_generate(self):
        try:
            start = int(self.start_entry.get())
            end = int(self.end_entry.get())
            if end <= start:
                raise ValueError

        except ValueError:
            messagebox.showerror(
                "Invalid input", "Please enter two integers where end > start."
            )
            return

        out_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx")],
            initialfile=f"QR_{start}_{end}.docx",
            title="Save QR document as…",
        )
        if not out_path:
            return  # user cancelled

        # Run the long job (it’s fast, so we skip threading/progress bars)
        try:
            create_qr_doc(start, end, Path(out_path))
        except Exception as exc:
            messagebox.showerror("Generation failed", str(exc))
            return

        messagebox.showinfo("Done", f"Document saved to:\n{out_path}")


if __name__ == "__main__":
    QRApp().mainloop()
