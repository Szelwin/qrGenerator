# QR Batch Generator

This tool creates Word documents filled with QR codes. The GUI version is
implemented in `qr_gui.py` with Tkinter while `qr_generator.py` provides the
core logic and a small CLI.

## Building executables

A GitHub Actions workflow builds standalone executables for Windows and macOS
using [PyInstaller](https://pyinstaller.org/). Each build runs on its native OS
and uploads the resulting binary as an artifact.

You can also build locally:

```bash
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --onefile --windowed qr_gui.py
```

The executable will be placed in the `dist/` directory.
