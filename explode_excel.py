from __future__ import annotations

import sys
import csv
from pathlib import Path
import openpyxl
from openpyxl.cell.cell import Cell, MergedCell


def export_sheets_with_formulas(wb: openpyxl.Workbook, output_dir: Path):
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


def export_named_ranges(wb: openpyxl.Workbook, output_dir: Path):
    """Export all defined names (Name Manager) to a TSV file."""
    tsv_path = output_dir / "names.tsv"
    sheet_names = wb.sheetnames

    # Skip Excel internal names (_xlpm.* = LET/LAMBDA params,
    # _xleta.* = internal function aliases)
    INTERNAL_PREFIXES = ("_xlpm.", "_xleta.")

    rows: list[list[str]] = []
    for defn in sorted(wb.defined_names.values(), key=lambda d: d.name.lower()):
        if any(defn.name.startswith(p) for p in INTERNAL_PREFIXES):
            continue
        if defn.localSheetId is not None:
            scope = sheet_names[defn.localSheetId]
        else:
            scope = "Workbook"
        comment = defn.comment or ""
        rows.append([defn.name, defn.value, scope, comment])

    with open(tsv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter="\t")
        writer.writerow(["Name", "Refers To", "Scope", "Comment"])
        writer.writerows(rows)

    print(f"    Exported {len(rows)} named ranges to {tsv_path}")


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
        wb = openpyxl.load_workbook(path, data_only=False, keep_links=False)
        export_sheets_with_formulas(wb, sheets_dir)
        export_named_ranges(wb, exploded_root)


if __name__ == "__main__":
    main()
