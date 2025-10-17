#!/usr/bin/env python3
"""
discover_xml_fields.py
--------------------------------
Scans an XML file and extracts all possible element and attribute field paths.

Usage:
    python discover_xml_fields.py path/to/input.xml

What it does:
--------------
1. Loads the XML.
2. Traverses all elements and attributes recursively.
3. Builds unique "field paths" like:
   - clashresult/@guid
   - clashresult/@name
   - clashobject/smarttag/value
   - clashpoint/pos3f/@x
4. Saves the discovered paths into `available_fields.json`.

Output Example:
---------------
{
  "elements": [
    "exchange/batchtest/clashtests/clashtest/summary/testtype",
    "exchange/batchtest/clashtests/clashtest/summary/teststatus",
    ...
  ],
  "attributes": [
    "exchange/batchtest/clashtests/clashtest/@name",
    "exchange/batchtest/clashtests/clashtest/clashresults/clashresult/@guid",
    ...
  ]
}
"""

import sys
import json
import xml.etree.ElementTree as ET
from pathlib import Path

def discover_fields(xml_path: Path):
    tree = ET.parse(xml_path)
    root = tree.getroot()

    elements = set()
    attributes = set()

    def walk(elem, path=""):
        tag_path = f"{path}/{elem.tag}" if path else elem.tag
        # Add attributes for this element
        for attr in elem.attrib:
            attributes.add(f"{tag_path}/@{attr}")
        # If element has text, record the tag itself
        if (elem.text and elem.text.strip()) or len(elem):
            elements.add(tag_path)
        # Recurse
        for child in elem:
            walk(child, tag_path)

    walk(root)

    return sorted(elements), sorted(attributes)


def main():
    if len(sys.argv) < 2:
        print("Usage: python discover_xml_fields.py path/to/input.xml")
        sys.exit(1)

    xml_file = Path(sys.argv[1])
    if not xml_file.exists():
        print(f"❌ File not found: {xml_file}")
        sys.exit(1)

    print(f"🔍 Scanning XML: {xml_file.name}")
    elements, attributes = discover_fields(xml_file)

    output_data = {"elements": elements, "attributes": attributes}
    output_file = xml_file.with_name("available_fields.json")

    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(output_data, f, indent=2)

    print(f"✅ Done! Found {len(elements)} elements and {len(attributes)} attributes.")
    print(f"📁 Saved to: {output_file}")

    # Optional preview
    print("\nSample element paths:")
    for e in elements[:10]:
        print("  ", e)
    print("...")

    print("\nSample attribute paths:")
    for a in attributes[:10]:
        print("  ", a)
    print("...")


if __name__ == "__main__":
    main()