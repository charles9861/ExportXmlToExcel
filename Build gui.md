Nice — below is a complete, copy-ready guide (step-by-step) plus three files you can drop into your GitHub repo.
It builds a simple Tkinter GUI launcher that:
	•	Lets you pick an XML file
	•	Exports to Excel (using your existing Excel exporter) or Word
	•	Remembers the last save folder in last_path.txt and uses it as the default next time
	•	Lets you optionally show a Save As dialog before actually saving

I’ll give you:
	1.	clash_exporter_gui.py — the GUI launcher (main file)
	2.	export_xml_to_word.py — a small example Word exporter (you can replace/extend with your project’s Word exporter)
	3.	README.md — step-by-step instructions (install, run, package with PyInstaller)

⸻

1) clash_exporter_gui.py

Save this file in the root of your project (same folder as your Excel/Word exporter modules).
Edit the EXPORT_EXCEL_MODULE / EXPORT_WORD_MODULE config values if your module filenames are different.

``` python
#!/usr/bin/env python3
"""
clash_exporter_gui.py
Simple Tkinter GUI that lets you pick an XML and export to Excel or Word.
It expects two functions to be importable from modules:
 - export_to_excel(xml_path: str, output_path: str)
 - export_to_word(xml_path: str, output_path: str)

Configuration at top of file (module names, function names) can be adjusted.
"""

import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import importlib
import sys
import os
import traceback

# ========== CONFIG - change these if your exporter module names differ ==========
EXPORT_EXCEL_MODULE = "export_xml_to_excel_v8_complete"  # module filename without .py
EXPORT_EXCEL_FUNC = "export_to_excel"

EXPORT_WORD_MODULE = "export_xml_to_word"                # module filename without .py
EXPORT_WORD_FUNC = "export_to_word"

LAST_PATH_FILE = "last_path.txt"   # saved in same folder as this script
# ==============================================================================

# Ensure script directory is first on sys.path so local modules can be imported
SCRIPT_DIR = Path(__file__).parent.resolve()
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))


def load_last_path():
    p = SCRIPT_DIR / LAST_PATH_FILE
    if p.exists():
        try:
            txt = p.read_text(encoding="utf-8").strip()
            if txt:
                return Path(txt)
        except Exception:
            return None
    return None


def save_last_path(folder_path: Path):
    try:
        (SCRIPT_DIR / LAST_PATH_FILE).write_text(str(folder_path), encoding="utf-8")
    except Exception:
        pass


def import_callable(module_name: str, func_name: str):
    try:
        module = importlib.import_module(module_name)
    except Exception as e:
        raise ImportError(f"Could not import module '{module_name}': {e}")
    try:
        func = getattr(module, func_name)
    except Exception:
        raise ImportError(f"Module '{module_name}' has no function '{func_name}'")
    return func


# Build the GUI
root = tk.Tk()
root.title("Clash XML Exporter")
root.geometry("560x200")
root.resizable(False, False)

# --- Widgets ---
frame_top = tk.Frame(root)
frame_top.pack(padx=12, pady=12, fill="x")

tk.Label(frame_top, text="Selected XML file:").grid(row=0, column=0, sticky="w")
xml_entry = tk.Entry(frame_top, width=62)
xml_entry.grid(row=1, column=0, columnspan=3, pady=(4, 8), sticky="w")

def browse_xml():
    initial = str(load_last_path() or SCRIPT_DIR)
    f = filedialog.askopenfilename(title="Select Clash XML file",
                                   initialdir=initial,
                                   filetypes=[("XML files", "*.xml"), ("All files", "*.*")])
    if f:
        xml_entry.delete(0, tk.END)
        xml_entry.insert(0, f)

tk.Button(frame_top, text="Browse...", command=browse_xml, width=12).grid(row=1, column=3, padx=(8,0))

# Buttons and options
frame_opts = tk.Frame(root)
frame_opts.pack(padx=12, pady=(0,8), fill="x")

ask_save_var = tk.BooleanVar(value=False)
tk.Checkbutton(frame_opts, text="Ask where to save (Save As)...", variable=ask_save_var).grid(row=0, column=0, sticky="w")

status_var = tk.StringVar(value=f"Last save folder: {str(load_last_path() or '(none)')}")
status_label = tk.Label(root, textvariable=status_var, anchor="w", fg="grey")
status_label.pack(fill="x", padx=12)

frame_buttons = tk.Frame(root)
frame_buttons.pack(padx=12, pady=(6,12), fill="x")
frame_buttons.columnconfigure((0,1,2), weight=1)

def do_export(kind: str):
    xml_path = xml_entry.get().strip()
    if not xml_path:
        messagebox.showerror("Error", "Please select an XML file first.")
        return
    xml_p = Path(xml_path)
    if not xml_p.exists():
        messagebox.showerror("Error", f"XML file not found:\n{xml_path}")
        return

    # Choose exporter function
    if kind == "excel":
        module_name = EXPORT_EXCEL_MODULE
        func_name = EXPORT_EXCEL_FUNC
        ext = ".xlsx"
        filetypes = [("Excel workbook", "*.xlsx")]
    else:
        module_name = EXPORT_WORD_MODULE
        func_name = EXPORT_WORD_FUNC
        ext = ".docx"
        filetypes = [("Word document", "*.docx")]

    try:
        exporter = import_callable(module_name, func_name)
    except ImportError as e:
        messagebox.showerror("Import error", str(e))
        return

    last = load_last_path()
    default_dir = last if last and last.exists() else xml_p.parent
    default_name = xml_p.stem + ext
    output_path = Path(default_dir) / default_name

    # If user wants to be asked where to save -> Save As
    if ask_save_var.get():
        chosen = filedialog.asksaveasfilename(title="Save As",
                                              initialdir=str(default_dir),
                                              initialfile=default_name,
                                              defaultextension=ext,
                                              filetypes=filetypes)
        if not chosen:
            return
        output_path = Path(chosen)

    # Try to run exporter (it should accept (xml_path, output_path) both strings or Paths)
    try:
        # call exporter
        exporter(str(xml_p), str(output_path))
        # update last save directory
        save_last_path(output_path.parent)
        status_var.set(f"Last save folder: {str(output_path.parent)}")
        messagebox.showinfo("Saved", f"Saved: {output_path}")
    except Exception as ex:
        tb = traceback.format_exc()
        print(tb)
        messagebox.showerror("Export failed", f"An error occurred while exporting:\n{ex}")

tk.Button(frame_buttons, text="Export to Excel", width=18, command=lambda: do_export("excel")).grid(row=0, column=0, padx=6)
tk.Button(frame_buttons, text="Export to Word", width=18, command=lambda: do_export("word")).grid(row=0, column=1, padx=6)
tk.Button(frame_buttons, text="Close", width=12, command=root.quit).grid(row=0, column=2, padx=6)

root.mainloop()

``` 
⸻

2) export_xml_to_word.py (example)

If you already have a Word export function, replace this file or rename it and update EXPORT_WORD_MODULE in the GUI config.
This example uses python-docx to create a simple report and includes images if they’re referenced in the XML (basic behavior — adapt to your formatting).

``` python

#!/usr/bin/env python3
"""
export_xml_to_word.py
A small example Word exporter exposing export_to_word(xml_file, output_file).
Install dependency: pip install python-docx Pillow
"""

import xml.etree.ElementTree as ET
from pathlib import Path
from docx import Document
from docx.shared import Inches
from PIL import Image as PILImage
import tempfile, os

def find_image_file(href_raw: str, xml_path: Path):
    if not href_raw:
        return None
    href = href_raw.replace("\\", "/").strip()
    candidates = []
    p = Path(href)
    if p.is_absolute():
        candidates.append(p)
    xml_dir = xml_path.parent
    candidates.append(xml_dir / href)
    fname = Path(href).name
    candidates.append(xml_dir / fname)
    candidates.append(xml_dir / "ELV_files" / fname)
    for cand in candidates:
        try:
            cand_resolved = cand.resolve()
        except Exception:
            cand_resolved = cand
        if cand_resolved.exists():
            return Path(cand_resolved)
    return None

def export_to_word(xml_file, output_file):
    xml_p = Path(xml_file)
    tree = ET.parse(xml_p)
    root = tree.getroot()

    doc = Document()
    doc.add_heading("Navisworks Clash Report", level=1)
    doc.add_paragraph(f"Source XML: {xml_file}")

    # Find clashes (support both clashresult and clashgroup)
    clashes = root.findall(".//clashresult")
    if not clashes:
        clashes = root.findall(".//clashgroup")

    for i, clash in enumerate(clashes, start=1):
        # Determine test/group names (basic heuristics similar to Excel exporter)
        test_name = "Unknown Test"
        group_name = clash.get("name", "None")
        # If clashresult, find its parent's test name
        if clash.tag.lower().endswith("clashresult"):
            guid = clash.get("guid")
            for test in root.findall(".//clashtest"):
                for cr in test.findall(".//clashresult"):
                    if guid and cr.get("guid") == guid:
                        test_name = test.get("name", "Unknown Test")
                        break
                if test_name != "Unknown Test":
                    break
        else:
            # For clashgroup, its parent may be clashtest -> climb up
            parent_test = clash.find("../..")
            if parent_test is not None and parent_test.get("name"):
                test_name = parent_test.get("name")

        # Basic details
        clash_name = clash.get("name", f"Clash{i}")
        pos = clash.find(".//pos3f")
        coords = ""
        if pos is not None:
            try:
                x = float(pos.get("x") or 0.0)
                y = float(pos.get("y") or 0.0)
                z = float(pos.get("z") or 0.0)
                coords = f"{x:.3f}, {y:.3f}, {z:.3f}"
            except Exception:
                coords = ""

        p = doc.add_paragraph()
        p.add_run(f"{i}. ").bold = True
        p.add_run(f"Test: {test_name}  |  Group: {group_name}\n").italic = True
        p.add_run(f"Clash: {clash_name}\n")
        if coords:
            p.add_run(f"Coords: {coords}\n")

        # Insert referenced image if available
        href_raw = clash.get("href") or ""
        img_path = find_image_file(href_raw, xml_p)
        if img_path:
            try:
                # ensure image fits reasonably
                doc.add_paragraph()
                doc.add_picture(str(img_path), width=Inches(4))
            except Exception:
                doc.add_paragraph(f"[Image present but couldn't embed: {img_path}]")

        doc.add_paragraph()  # small spacer

    doc.save(output_file)
    print(f"Saved Word report: {output_file}")

```

⸻

3) README.md (copy into your repo)

A short step-by-step README you can commit:


``` python
# Clash Exporter GUI (Tkinter)

Small GUI to run XML → Excel / Word exporters from a simple desktop window.

## Files
- `clash_exporter_gui.py` - Main Tkinter GUI.
- `export_xml_to_word.py`  - Example Word exporter (replace/extend if you already have your own).
- Your Excel exporter (e.g. `export_xml_to_excel_v8_complete.py`) should be in the same folder.
- `last_path.txt` will be created automatically to remember the last save folder.

## Requirements
Install the dependencies used by your exporters and (for the example Word exporter):
```bash
pip install openpyxl pillow python-docx

```

Usage (development)
	1.	Put clash_exporter_gui.py in the same folder as your Excel exporter module.
	2.	Edit the top of clash_exporter_gui.py to set:
	•	EXPORT_EXCEL_MODULE to your Excel exporter module name (filename without .py)
	•	EXPORT_WORD_MODULE to your Word exporter module name (if different)
	3.	Run:
``` bash
python clash_exporter_gui.py
``` 
	4.	Click Browse… to select an XML file, then Export to Excel or Export to Word.
	•	By default, the output will be saved into the last-used folder (or the XML’s folder).
	•	Check Ask where to save to show a Save As dialog instead.

Packaging to a single executable (Windows example)

Install PyInstaller:

``` bash
pip install pyinstaller
``` 

Build:

``` bash
pyinstaller --onefile --noconsole clash_exporter_gui.py
```

The generated executable will be in dist/.

Notes
	•	Exporter functions must follow signature: export_to_excel(xml_path, output_path) and export_to_word(xml_path, output_path).
	•	If your export functions have different names, either:
	•	rename the functions, or
	•	edit the EXPORT_*_FUNC values in clash_exporter_gui.py.

---


``` python
## Quick checklist for you (on PC)
1. Put the three files in your repo root.  
2. Ensure your **Excel exporter** module is present and exposes `export_to_excel(xml, out)`. If it has a different name, edit `EXPORT_EXCEL_MODULE` at top of the GUI file.  
3. `pip install` dependencies: `openpyxl pillow python-docx` (and any other libs your exporter uses).  
4. Run `python clash_exporter_gui.py` and test with an XML.  
5. If everything works, package with PyInstaller (optional) to share an `.exe`.

---
``` pyhton
## Optional polish ideas (later)
- Add a small progress bar during export.  
- Add a small preview pane for the first clash image.  
- Add drag-and-drop of XML onto the window.  
- Replace Tkinter with PySimpleGUI or PyQt for a nicer UI.

---

If you want, I can:
- Tailor the GUI to *exactly* call your current module names (tell me the filenames) — I’ll update the `EXPORT_...` defaults before you commit.
- Produce a one-file `export_xml_to_word.py` that matches your Word formatting (if you paste your Word export code here I will convert it).

Want me to update the GUI config now to use the exact filenames from your project (tell me the filenames)?
``` 
