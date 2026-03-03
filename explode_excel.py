import sys
import csv
import struct
import base64
import zipfile
import re
import xml.etree.ElementTree as ET
from pathlib import Path
import openpyxl
from openpyxl.cell.cell import Cell, MergedCell


def export_sheets_with_formulas(xlsx_path: Path, output_dir: Path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=False, keep_links=False)
    output_dir.mkdir(parents=True, exist_ok=True)

    for sheet in wb.worksheets:

        def get_formula_or_value(cell: Cell | MergedCell) -> str:
            val = cell.value
            if cell.data_type == "f":
                return val if isinstance(val, str) else str(val.text)  # type: ignore
            else:
                return str(val) if val is not None else ""

        csv_path = output_dir / f"{sheet.title}.csv"

        # Collect all rows
        rows = [
            [get_formula_or_value(cell) for cell in row]
            for row in sheet.iter_rows(values_only=False)
        ]

        # Trim trailing empty rows
        while rows and all(
            cell is None or str(cell).strip() == "" for cell in rows[-1]
        ):
            rows.pop()

        # Write trimmed rows to CSV
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerows(rows)


def extract_fast_data_load(xlsx_path: Path) -> dict[str, bool]:
    """Extract EnableFastDataLoad per query from the DataMashup metadata.

    EnableFastDataLoad is stored as the inverse of the BufferNextRefresh
    entry in the LocalPackageMetadataFile inside the DataMashup binary
    (customXml/item*.xml).

    Returns a dict mapping query name (e.g. "customers") to True/False.
    """
    result: dict[str, bool] = {}

    with zipfile.ZipFile(xlsx_path) as z:
        # Find the DataMashup customXml item
        mashup_xml = None
        for name in z.namelist():
            if not re.match(r"customXml/item\d+\.xml", name):
                continue
            raw = z.read(name)
            try:
                text = raw.decode("utf-8")
            except UnicodeDecodeError:
                text = raw.decode("utf-16")
            if "DataMashup" in text:
                mashup_xml = text
                break

        if mashup_xml is None:
            return result

        root = ET.fromstring(mashup_xml)
        if root.text is None:
            return result
        raw = base64.b64decode(root.text.strip())

        # Binary layout:
        #   version(4) + pkg_len(4) + ZIP(pkg_len)
        #   + perm_len(4) + perm(perm_len)
        #   + meta_len(4) + meta(meta_len)
        offset = 4  # skip version
        pkg_len = struct.unpack_from("<I", raw, offset)[0]
        offset += 4 + pkg_len  # skip ZIP
        perm_len = struct.unpack_from("<I", raw, offset)[0]
        offset += 4 + perm_len  # skip permissions
        meta_len = struct.unpack_from("<I", raw, offset)[0]
        offset += 4
        meta_data = raw[offset : offset + meta_len]

        # Metadata has an 8-byte sub-header (version u32 + xml_length u32)
        xml_len = struct.unpack_from("<I", meta_data, 4)[0]
        meta_xml = meta_data[8 : 8 + xml_len].decode("utf-8-sig")
        meta_root = ET.fromstring(meta_xml)

        items_elem = list(meta_root)[0]  # <Items>
        for item in items_elem:
            loc_elem = item.find("ItemLocation")
            if loc_elem is None:
                continue
            item_path = loc_elem.find("ItemPath")
            if item_path is None or not item_path.text:
                continue
            # ItemPath looks like "Section1/QueryName"
            query_name = item_path.text.strip().split("/")[-1]

            stable = item.find("StableEntries")
            if stable is None:
                continue
            for entry in stable:
                if entry.attrib.get("Type") == "BufferNextRefresh":
                    # BufferNextRefresh "l1" = True → FastDataLoad = False
                    # BufferNextRefresh "l0" = False → FastDataLoad = True
                    buffer_val = entry.attrib.get("Value", "")
                    result[query_name] = buffer_val == "l0"
                    break

    return result


def patch_m_files_with_fast_data_load(
    queries_dir: Path, fast_data_load: dict[str, bool]
):
    """Inject EnableFastDataLoad into exported .m files.

    Looks for the connection-properties comment block and appends the
    EnableFastDataLoad line after RefreshWithRefreshAll (or replaces an
    existing one).
    """
    for m_file in queries_dir.glob("*.m"):
        query_name = m_file.stem
        if query_name not in fast_data_load:
            continue

        enabled = fast_data_load[query_name]
        fdl_line = f"//   EnableFastDataLoad:    {enabled}\n"

        content = m_file.read_text(encoding="utf-8")

        # Replace existing EnableFastDataLoad line if present
        if "//   EnableFastDataLoad:" in content:
            content = re.sub(r"//   EnableFastDataLoad:.*\n", fdl_line, content)
        # Otherwise insert after RefreshWithRefreshAll line
        elif "//   RefreshWithRefreshAll:" in content:
            content = re.sub(
                r"(//   RefreshWithRefreshAll:.*\n)",
                r"\1" + fdl_line,
                content,
            )
        else:
            # No connection properties header; skip
            continue

        m_file.write_text(content, encoding="utf-8")
        print(f"  [+] Patched EnableFastDataLoad={enabled} in {query_name}.m")


def delete_existing():
    path = "exploded/00 TUS cert/sheets"
    for child in Path(path).glob("*"):
        if child.is_file():
            child.unlink()
        elif child.is_dir():
            for subchild in child.glob("**/*"):
                if subchild.is_file():
                    subchild.unlink()
            child.rmdir()


def main():
    delete_existing()
    for path_str in sys.argv[1:]:
        path = Path(path_str)
        if path.suffix.lower() not in {".xlsx", ".xltm", ".xlsm", ".xltx", ".xlsb"}:
            continue

        exploded_root = Path("exploded") / path.stem
        sheets_dir = exploded_root / "sheets"

        print(f"[+] Processing: {path.name}")
        export_sheets_with_formulas(path, sheets_dir)

        # Patch .m query files with EnableFastDataLoad from DataMashup
        queries_dir = exploded_root / "queries"
        if queries_dir.is_dir():
            fast_data_load = extract_fast_data_load(path)
            if fast_data_load:
                patch_m_files_with_fast_data_load(queries_dir, fast_data_load)


if __name__ == "__main__":
    main()
