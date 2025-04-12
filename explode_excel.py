import sys
import csv
from pathlib import Path
import openpyxl
from openpyxl.cell.cell import Cell


def export_sheets_with_formulas(xlsx_path: Path, output_dir: Path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=False, keep_links=False)
    output_dir.mkdir(parents=True, exist_ok=True)

    for sheet in wb.worksheets:
        def get_formula_or_value(cell: Cell) -> str:
            val = cell.value
            if cell.data_type == 'f':
                return val if isinstance(val, str) else val.text
            else:
                return val

        csv_path = output_dir / f"{sheet.title}.csv"
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            for row in sheet.iter_rows(values_only=False):
                writer.writerow([get_formula_or_value(cell) for cell in row])


def main():
    for path_str in sys.argv[1:]:
        path = Path(path_str)
        if path.suffix.lower() not in {
                ".xlsx", ".xltm", ".xlsm", ".xltx", ".xlsb"}:
            continue

        exploded_root = Path("exploded") / path.stem
        sheets_dir = exploded_root / "sheets"

        print(f"[+] Processing: {path.name}")
        export_sheets_with_formulas(path, sheets_dir)


if __name__ == "__main__":
    main()
