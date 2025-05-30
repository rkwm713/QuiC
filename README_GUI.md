# SPIDA ↔ Katapult Desktop GUI

This folder contains the standalone Tkinter application for comparing SPIDAcalc and Katapult Pro JSON exports.

## Quick start

```bash
# 1‒ Install dependencies (into a venv is recommended)
python -m pip install -r requirements.txt

# 2‒ Launch the GUI
python -m gui.main        # preferred (runs as package)
# or
python gui/main.py        # works too
```

## Folder layout

```
<project>/
├── gui/
│   ├── __init__.py
│   ├── main.py          # application window
│   └── editable_tree.py # inline-edit Treeview widget
├── compare.py           # core comparison engine (imported by gui/main.py)
├── spida_writer.py      # helper to patch SPIDA JSON during save
└── requirements.txt
```

* `compare.py` must provide `compare()` and `haversine_m()`
* `spida_writer.py` must provide `apply_edit()`

If you place them elsewhere, adjust the imports in `gui/main.py` accordingly.

## Notes

* The map view relies on `tkintermapview` which in turn needs Tk 8.6+.  Standard Python on Windows/macOS already ships with Tk.  On minimal Linux distros you may need:
  ```bash
  sudo apt-get install python3-tk
  ```
* Data never leaves your machine; all comparison and JSON editing is local. 