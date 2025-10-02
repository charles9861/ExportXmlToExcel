import config
import xml.etree.ElementTree as ET
import os
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage

def export_to_excel(xml_file, output_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    wb = Workbook()
    ws = wb.active
    ws.title = "Clash Report"

    # -----------------------
    # Column headers
    # -----------------------
    headers = ["#", "Clash Details", "Item 1", "Item 2", "Clash Image", "User Images", "Comments"]
    ws.append(headers)

    # -----------------------
    # Find all clash results
    # -----------------------
    clashes = root.findall(".//clashresult")
    for i, clash in enumerate(clashes, start=1):

        # --- Get Clash Group (from parent <clashtest>)
        clash_group = "Unknown Group"
        parent = clash.find("..")  # ElementTree has no getparent(), workaround below
        # Instead: search backwards to find the clashtest ancestor
        clashtest = clash.find("ancestor::clashtest")  # lxml supports, ET does not
        if clashtest is None:
            # fallback: manual loop
            for test in root.findall(".//clashtest"):
                if clash in test.findall(".//clashresult"):
                    clash_group = test.get("name", "Unknown Group")
                    break
        else:
            clash_group = clashtest.get("name", "Unknown Group")

        # --- Clash basic info
        clash_name = clash.get("name", f"Clash {i}")
        distance = clash.get("distance", "N/A")
        pos = clash.find(".//pos3f")
        coords = ""
        if pos is not None:
            coords = f"{float(pos.get('x')):.3f}m,\n{float(pos.get('y')):.3f}m,\n{float(pos.get('z')):.3f}m"

        # --- Extract Item details
        objs = clash.findall(".//clashobject")
        item1 = get_item_details(objs[0]) if len(objs) > 0 else "Unknown"
        item2 = get_item_details(objs[1]) if len(objs) > 1 else "Unknown"

        # -----------------------
        # Build Clash Details (Column 2)
        # -----------------------
        clash_details = (
            f"Clash Group: {clash_group}\n"
            f"Between: {short_item_name(item1)} and {short_item_name(item2)}\n"
            f"{clash_name}\n"
            f"Distance: {distance}m\n"
            f"Clash Point:\n{coords}"
        )

        # -----------------------
        # Add row to sheet
        # -----------------------
        row = [i, clash_details, item1, item2, "", "", ""]
        ws.append(row)

        # --- Add clash screenshot (Column 5)
        img_href = clash.get("href", "").replace("\\", "/")
        if img_href and os.path.exists(img_href):
            img = XLImage(img_href)
            img.width, img.height = 200, 160
            ws.add_image(img, f"E{ws.max_row}")

    wb.save(output_file)
    print(f"Exported: {output_file}")


# -----------------------
# Extract Item details for columns 3 & 4
# -----------------------
def get_item_details(clash_object):
    smarttags = clash_object.findall(".//smarttag")
    data = {}
    for tag in smarttags:
        name = tag.findtext("name", "").strip()
        value = tag.findtext("value", "").strip()
        if not name or not value:
            continue
        if name == "Item Name":
            data["itemName"] = value
        elif name == "Civil3D General:Network name":
            data["network"] = value
        elif name == "Civil3D General:Part Size Name":
            data["partSize"] = value
        elif name == "Item Type":
            data["itemType"] = value
        elif name == "Civil3D General:Inner Diameter or Width":
            data["inner"] = value
        elif name == "Civil3D General:Outer Diameter or Width":
            data["outer"] = value

    type_line1 = " ".join(filter(None, [data.get("partSize"), data.get("itemType")]))
    type_line2 = ""
    if data.get("inner") or data.get("outer"):
        type_line2 = f"Pipe {data.get('inner','')} x {data.get('outer','')}".strip()

    lines = []
    if data.get("itemName"):
        lines.append(f"Item Name: {data['itemName']}")
    if data.get("network"):
        lines.append(f"Network: {data['network']}")
    if type_line1:
        lines.append(f"Item Type: {type_line1}")
    if type_line2:
        lines.append(type_line2)

    return "\n".join(lines)


# -----------------------
# Helper: Short name for "Between" line
# (just the Item Name if available)
# -----------------------
def short_item_name(item_details):
    for line in item_details.splitlines():
        if line.startswith("Item Name:"):
            return line.replace("Item Name:", "").strip()
    return item_details.splitlines()[0] if item_details else "Unknown"


if __name__ == "__main__":
    export_to_excel(config.XML_FILE, config.OUTPUT_FILE)
