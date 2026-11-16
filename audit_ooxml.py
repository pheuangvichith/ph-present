#!/usr/bin/env python3
"""Strict structural OOXML + ph-present audit for .pptx files."""

import io
import os
import posixpath
import sys
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "p159": "http://schemas.microsoft.com/office/powerpoint/2015/09/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}


def audit(pptx_path: str) -> int:
    errors = []
    with zipfile.ZipFile(pptx_path, "r") as zf:
        names = set(zf.namelist())
        print(f"Package entries: {len(names)}")

        # 1. Parse every XML & .rels entry
        for name in sorted(names):
            if name.endswith(".xml") or name.endswith(".rels"):
                try:
                    ET.fromstring(zf.read(name))
                except ET.ParseError as e:
                    errors.append(f"XML parse error in {name}: {e}")

        # 2. Verify all .rels targets resolve
        for name in sorted(names):
            if name.endswith(".rels"):
                rel_dir = posixpath.dirname(posixpath.dirname(name)) if "_rels/" in name else ""
                if name == "_rels/.rels":
                    rel_dir = ""
                root = ET.fromstring(zf.read(name))
                for rel in root:
                    target = rel.attrib.get("Target", "")
                    if target.startswith("http"):
                        continue
                    resolved = posixpath.normpath(posixpath.join(rel_dir, target))
                    if resolved not in names:
                        errors.append(f"Broken rel in {name}: {target} -> {resolved}")

        # 3. Validate embedded xlsx
        xlsx_entries = [n for n in names if n.endswith(".xlsx")]
        for xe in xlsx_entries:
            try:
                with zipfile.ZipFile(io.BytesIO(zf.read(xe)), "r") as xzf:
                    for xn in xzf.namelist():
                        if xn.endswith(".xml") or xn.endswith(".rels"):
                            ET.fromstring(xzf.read(xn))
                print(f"Embedded XLSX OK: {xe} ({len(xzf.namelist())} parts)")
            except Exception as e:
                errors.append(f"Invalid embedded XLSX {xe}: {e}")

        # 4. Per-slide ph-present checks
        slide_files = sorted(
            [n for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml")],
            key=lambda s: int(s.split("slide")[-1].split(".")[0])
        )
        print(f"Slides found: {len(slide_files)}")
        for sf in slide_files:
            root = ET.fromstring(zf.read(sf))
            c_nv_prs = root.findall(".//p:cNvPr", NS)
            all_names = [el.attrib.get("name", "") for el in c_nv_prs if el.attrib.get("name", "")]
            bang_names = [n for n in all_names if n.startswith("!!")]
            dups = [k for k, v in Counter(bang_names).items() if v > 1]
            if dups:
                errors.append(f"Duplicate !! names on {sf}: {dups}")

            has_morph = root.find(".//p159:morph", NS) is not None
            rot_count = len([el for el in root.findall(".//a:xfrm", NS) if "rot" in el.attrib])
            cust_count = len(root.findall(".//a:custGeom", NS))
            print(
                f"  {os.path.basename(sf)}: !!anchors={len(bang_names)} (dups={len(dups)}), "
                f"morph={has_morph}, rotated={rot_count}, custGeom={cust_count}"
            )

    if errors:
        print("\nAUDIT FAILED:")
        for e in errors:
            print("  ERROR:", e)
        return 1
    print("\nOOXML AUDIT PASSED: 0 errors, 0 duplicate !! anchors.")
    return 0


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "Inside_the_Agent_Loop.pptx")
    sys.exit(audit(path))
