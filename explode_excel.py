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
            rows = [
                [get_formula_or_value(cell) for cell in row]
                for row in sheet.iter_rows(values_only=False)
            ]
            writer.writerows(rows)

        # After writing, check if all rows are just commas
        with open(csv_path, 'r+', encoding='utf-8') as f:
            content = f.read()
            if all(
                not any(cell.strip() for cell in line.split(','))
                for line in content.splitlines()
            ):
                f.seek(0)
                f.truncate()

def delete_existing():
    path = 'exploded/00 TUS cert/sheets'
    for child in Path(path).glob('*'):
        if child.is_file():
            child.unlink()
        elif child.is_dir():
            for subchild in child.glob('**/*'):
                if subchild.is_file():
                    subchild.unlink()
            child.rmdir()

def main():
    delete_existing()
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
