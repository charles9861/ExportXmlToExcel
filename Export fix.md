fix line shading and header tex size


Perfect 🚀 glad the `short_item_name` issue is solved.

Now, let’s tweak two things:

1. **Alternate row shading:**
   Right now, your code does this:

```python
if (current_row - start_data_row) % 2 == 0:
    # even offset -> fill
```

That actually fills **even-numbered rows** and leaves odd ones blank — but because you offset by `start_data_row`, it can feel flipped.
If you want **even rows filled, odd rows no fill**, then we’ll invert the condition.

---

2. **Header text size = 18pt:**
   You’re currently setting:

```python
header_font = Font(color=HEADER_FONT_COLOR, bold=True)
```

Just add `size=18`.

---

### ✅ Updated fixes

#### Change header font setup:

```python
header_font = Font(color=HEADER_FONT_COLOR, bold=True, size=18)
```

---

#### Change alternate row fill loop:

Replace this block:

```python
# Alternate row fill color every second row (data rows)
if (current_row - start_data_row) % 2 == 0:
    # even offset -> fill
    for c in range(1, 8):
        cell = ws.cell(row=current_row, column=c)
        # apply background fill
        cell.fill = PatternFill(start_color=ALT_ROW_FILL, end_color=ALT_ROW_FILL, fill_type="solid")
        # set alignment and wrap
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.font = Font(color="000000")
else:
    for c in range(1, 8):
        cell = ws.cell(row=current_row, column=c)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.font = Font(color="000000")
```

with this:

```python
# Alternate row fill: even rows filled, odd rows no fill
for c in range(1, 8):
    cell = ws.cell(row=current_row, column=c)
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    cell.font = Font(color="000000")

    if (current_row % 2 == 0):  # even row -> apply fill
        cell.fill = PatternFill(start_color=ALT_ROW_FILL, end_color=ALT_ROW_FILL, fill_type="solid")
```

---

This way:

* Row 2, 4, 6… get filled.
* Row 3, 5, 7… remain no fill.
* Header row (row 1) already styled separately.
* Header font is now size 18.

---

👉 Do you want me to give you the **entire corrected script block** for `export_to_excel` so you can just copy-paste cleanly?
