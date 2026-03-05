#!/usr/bin/env python3
"""
vba-inject.py — Inject VBA modules into an xlsx/xlsm file.

Usage:
    # Inject VBA into existing xlsx (converts to xlsm):
    python3 vba-inject.py input.xlsx output.xlsm Module1.bas

    # Inject multiple modules:
    python3 vba-inject.py input.xlsx output.xlsm Module1.bas Module2.bas ThisWorkbook.cls

    # Inject into existing xlsm (replaces VBA project):
    python3 vba-inject.py input.xlsm output.xlsm Module1.bas

Module file naming convention:
    - *.bas → StdModule (standard VBA module)
    - *.cls → DocModule (document module, e.g. ThisWorkbook)

No external dependencies — uses vba_project_builder.py from this directory.
"""

import sys
import os
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET

# Import our own builder
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vba_project_builder import build_vba_project


def create_vba_project_bin(module_files: list, output_path: str,
                           code_page: int = 1251) -> None:
    """Create vbaProject.bin from VBA source files."""
    modules = []
    for filepath in module_files:
        basename = os.path.basename(filepath)
        name, ext = os.path.splitext(basename)
        ext = ext.lower()

        with open(filepath, "r", encoding="utf-8") as f:
            source = f.read()
        # Normalize line endings to CRLF (VBA requirement)
        source = source.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")

        if ext == ".cls":
            modules.append((name, "document", source))
        elif ext == ".bas":
            modules.append((name, "standard", source))
        else:
            print(f"Warning: skipping unknown file type: {filepath}",
                  file=sys.stderr)

    # Ensure ThisWorkbook exists
    has_this_workbook = any(
        name.lower() in ("thisworkbook", "этакнига", "этакнига")
        for name, _, _ in modules
    )
    if not has_this_workbook:
        modules.insert(0, ("ThisWorkbook", "document", ""))

    build_vba_project(modules, output_path, code_page=code_page)


def inject_vba_into_xlsx(input_path: str, output_path: str,
                         vba_bin_path: str) -> None:
    """Inject vbaProject.bin into xlsx/xlsm, producing a valid xlsm."""
    with tempfile.TemporaryDirectory() as tmpdir:
        # Extract original zip
        with zipfile.ZipFile(input_path, "r") as zin:
            zin.extractall(tmpdir)

        # 1. Copy vbaProject.bin into xl/
        xl_dir = os.path.join(tmpdir, "xl")
        shutil.copy2(vba_bin_path, os.path.join(xl_dir, "vbaProject.bin"))

        # 2. Update [Content_Types].xml
        ct_path = os.path.join(tmpdir, "[Content_Types].xml")
        ns = {"ct": "http://schemas.openxmlformats.org/package/2006/content-types"}
        ET.register_namespace("", ns["ct"])
        tree = ET.parse(ct_path)
        root = tree.getroot()

        # Change workbook ContentType to macroEnabled
        for override in root.findall("ct:Override", ns):
            if override.get("PartName") == "/xl/workbook.xml":
                override.set(
                    "ContentType",
                    "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                )

        # Add vbaProject.bin override if not present
        vba_override_exists = any(
            o.get("PartName") == "/xl/vbaProject.bin"
            for o in root.findall("ct:Override", ns)
        )
        if not vba_override_exists:
            ET.SubElement(
                root,
                "Override",
                PartName="/xl/vbaProject.bin",
                ContentType="application/vnd.ms-office.vbaProject",
            )

        tree.write(ct_path, xml_declaration=True, encoding="UTF-8")

        # 3. Add relationship in xl/_rels/workbook.xml.rels
        rels_path = os.path.join(xl_dir, "_rels", "workbook.xml.rels")
        ns_rels = {
            "r": "http://schemas.openxmlformats.org/package/2006/relationships"
        }
        ET.register_namespace("", ns_rels["r"])
        rels_tree = ET.parse(rels_path)
        rels_root = rels_tree.getroot()

        # Check if vbaProject relationship already exists
        vba_rel_exists = any(
            rel.get("Target") == "vbaProject.bin"
            for rel in rels_root.findall("r:Relationship", ns_rels)
        )
        if not vba_rel_exists:
            existing_ids = [
                int(rel.get("Id", "rId0").replace("rId", ""))
                for rel in rels_root.findall("r:Relationship", ns_rels)
                if rel.get("Id", "").startswith("rId")
            ]
            next_id = max(existing_ids, default=0) + 1

            ET.SubElement(
                rels_root,
                "Relationship",
                Id=f"rId{next_id}",
                Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject",
                Target="vbaProject.bin",
            )

        rels_tree.write(rels_path, xml_declaration=True, encoding="UTF-8")

        # 4. Re-zip everything into output xlsm
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for dirpath, dirnames, filenames in os.walk(tmpdir):
                for fname in filenames:
                    full_path = os.path.join(dirpath, fname)
                    arcname = os.path.relpath(full_path, tmpdir)
                    zout.write(full_path, arcname)

    print(f"Created: {output_path}")


def main():
    if len(sys.argv) < 4:
        print("Usage: python3 vba-inject.py <input.xlsx|xlsm> <output.xlsm>"
              " <module1.bas> [module2.cls ...]")
        print()
        print("Module types by extension:")
        print("  *.bas  Standard VBA module")
        print("  *.cls  Document module (ThisWorkbook, Sheet1, etc.)")
        print()
        print("No external dependencies required.")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]
    module_files = sys.argv[3:]

    if not os.path.exists(input_path):
        print(f"Error: input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    for mf in module_files:
        if not os.path.exists(mf):
            print(f"Error: module file not found: {mf}", file=sys.stderr)
            sys.exit(1)

    with tempfile.TemporaryDirectory() as tmpdir:
        vba_bin = os.path.join(tmpdir, "vbaProject.bin")

        print(f"Creating vbaProject.bin from {len(module_files)} module(s)...")
        create_vba_project_bin(module_files, vba_bin)

        print(f"Injecting VBA into {input_path}...")
        inject_vba_into_xlsx(input_path, output_path, vba_bin)


if __name__ == "__main__":
    main()
