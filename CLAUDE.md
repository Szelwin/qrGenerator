# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a QR code batch generator that creates Microsoft Word documents (.docx) filled with QR codes. The project has two main interfaces:
- CLI: `qr_generator.py` - Core logic with command-line interface  
- GUI: `qr_gui.py` - Tkinter-based graphical interface

## Architecture

The codebase follows a clean separation of concerns:

- **Core Logic (`qr_generator.py`)**: Contains all QR generation and document creation functionality. Key functions:
  - `create_qr_doc(start, end_exclusive, out_path)` - Main API for creating QR documents
  - `add_qr_block(doc, start_num, end_num, cols=17)` - Adds QR code table blocks to documents
  - `_qr_png_stream(data)` - Generates individual QR code PNG streams
  - `_chunk_range(start, end, size=100)` - Splits large ranges into manageable chunks

- **GUI Interface (`qr_gui.py`)**: Minimal Tkinter wrapper around the core logic. Uses `QRApp` class with simple input validation and file dialog integration.

## Development Commands

### Setup
```bash
pip install -r requirements.txt
```

### Running the Application
```bash
# CLI usage
python qr_generator.py 1000 1200

# GUI usage  
python qr_gui.py
```

### Building Executables
```bash
# Install PyInstaller
pip install pyinstaller

# Build GUI executable (cross-platform)
pyinstaller --onefile --windowed --hidden-import=PIL qr_gui.py

# Windows (no console)
pyinstaller --noconsole --onefile --name qr_gui --hidden-import=PIL qr_gui.py

# The executable will be in dist/ directory
```

## Key Dependencies

- `qrcode[pil]` - QR code generation
- `python-docx` - Word document manipulation  
- `Pillow` - Image processing (required for QR code images)
- `tkinter` - GUI framework (built into Python)

## GitHub Actions

The project includes automated builds for Windows and macOS executables on every push to main. The workflow:
1. Sets up Python 3.11 environment
2. Installs dependencies and PyInstaller
3. Builds platform-specific executables
4. Updates the "Latest" release with new binaries

## Technical Notes

- QR codes are generated with `box_size=5, border=2` and low error correction
- Documents use A4 format with 0.5" margins
- QR codes are arranged in 17-column tables with 9mm width
- Large ranges are automatically chunked into blocks of 100 for better document organization
- Range labels are added to the right of each QR block