#!/usr/bin/env python3
"""
python clash_exporter_gui.py
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
EXPORT_EXCEL_MODULE = "export_xml_to_excel_v8"  # module filename without .py
EXPORT_EXCEL_FUNC = "export_to_excel"

EXPORT_WORD_MODULE = "export_xml_to_word_v4"                # module filename without .py
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
