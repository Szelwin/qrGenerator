"""
gui_app.py
==========

Tiny Tkinter shell around qr_document.build_and_save().
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
import threading
import queue
from qr_generator import create_qr_doc, chunk_range


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

        # Progress bar (initially hidden)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            self, variable=self.progress_var, maximum=100, length=300
        )
        self.progress_bar.grid(row=3, column=0, columnspan=2, pady=(4, 8), sticky="ew")
        self.progress_bar.grid_remove()  # Hide initially

        # Status label for progress
        self.status_label = ttk.Label(self, text="")
        self.status_label.grid(row=4, column=0, columnspan=2, **paddings)
        self.status_label.grid_remove()  # Hide initially

        # Queue for thread communication
        self.progress_queue = queue.Queue()
        self.generation_thread = None

    # ──────────────────────────────────────────────────────────────────────
    # Callbacks
    # ──────────────────────────────────────────────────────────────────────
    def _on_generate(self):
        # Prevent multiple simultaneous generations
        if self.generation_thread and self.generation_thread.is_alive():
            return

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
        self._start_generation(start, end, Path(out_path))

    def _start_generation(self, start, end, out_path):
        """Start QR generation in background thread with progress tracking."""
        # Show progress UI
        self.progress_bar.grid()
        self.status_label.grid()
        self.generate_btn.config(state="disabled", text="Generating...")
        self.progress_var.set(0)
        self.status_label.config(text="Starting generation...")

        # Clear the queue
        while not self.progress_queue.empty():
            try:
                self.progress_queue.get_nowait()
            except queue.Empty:
                break

        # Start background thread
        self.generation_thread = threading.Thread(
            target=self._generate_qr_codes,
            args=(start, end, out_path),
            daemon=True
        )
        self.generation_thread.start()

        # Start monitoring progress
        self._check_progress()

    def _generate_qr_codes(self, start, end, out_path):
        """Background worker that generates QR codes with progress updates."""
        try:
            from qr_generator import create_document, add_qr_block, QRConfig, DocumentConfig
            
            qr_config = QRConfig()
            doc_config = DocumentConfig()
            
            # Calculate total chunks for progress
            total_chunks = len(list(chunk_range(start, end, doc_config.chunk_size)))
            
            self.progress_queue.put(("status", f"Creating document with {end-start} QR codes..."))
            
            doc = create_document(doc_config)
            chunk_count = 0
            
            for chunk_start, chunk_end in chunk_range(start, end, doc_config.chunk_size):
                chunk_count += 1
                chunk_size = chunk_end - chunk_start + 1
                
                self.progress_queue.put((
                    "status", 
                    f"Processing chunk {chunk_count}/{total_chunks} ({chunk_start}-{chunk_end})"
                ))
                
                add_qr_block(doc, chunk_start, chunk_end, qr_config, doc_config)
                
                # Update progress
                progress = (chunk_count / total_chunks) * 100
                self.progress_queue.put(("progress", progress))
            
            self.progress_queue.put(("status", "Saving document..."))
            doc.save(out_path)
            
            self.progress_queue.put(("success", str(out_path)))
            
        except Exception as exc:
            self.progress_queue.put(("error", str(exc)))

    def _check_progress(self):
        """Check for updates from the background thread."""
        try:
            while True:
                msg_type, msg_data = self.progress_queue.get_nowait()
                
                if msg_type == "progress":
                    self.progress_var.set(msg_data)
                elif msg_type == "status":
                    self.status_label.config(text=msg_data)
                elif msg_type == "success":
                    self._generation_complete(msg_data)
                    return
                elif msg_type == "error":
                    self._generation_error(msg_data)
                    return
                    
        except queue.Empty:
            pass
        
        # Continue checking
        self.after(100, self._check_progress)

    def _generation_complete(self, out_path):
        """Handle successful completion of QR generation."""
        self.progress_bar.grid_remove()
        self.status_label.grid_remove()
        self.generate_btn.config(state="normal", text="Generate")
        messagebox.showinfo("Done", f"Document saved to:\n{out_path}")

    def _generation_error(self, error_msg):
        """Handle error during QR generation."""
        self.progress_bar.grid_remove()
        self.status_label.grid_remove()
        self.generate_btn.config(state="normal", text="Generate")
        messagebox.showerror("Generation failed", error_msg)


if __name__ == "__main__":
    QRApp().mainloop()
