import sys
import zipfile
import csv
import xml.etree.ElementTree as ET
from pathlib import Path
from openpyxl import load_workbook

def extract_m_scripts(xlsx_path: Path, output_dir: Path):
    with zipfile.ZipFile(xlsx_path, 'r') as zipf:
        try:
            queries_xml = zipf.read('xl/powerquery/queries.xml')
        except KeyError:
            print(f"[info] No Power Query found in {xlsx_path.name}")
            return

        output_dir.mkdir(parents=True, exist_ok=True)
        root = ET.fromstring(queries_xml)
        ns = {'pq': 'http://schemas.microsoft.com/office/powerquery/'}  # namespace

        for query in root.findall('pq:Query', ns):
            name = query.attrib.get('Name', 'UnnamedQuery')
            formula_elem = query.find('pq:Formula', ns)
            if formula_elem is not None:
                m_code = formula_elem.text or ''
                m_path = output_dir / f"{name}.m"
                with open(m_path, 'w', encoding='utf-8') as f:
                    f.write(m_code)

def export_sheets_with_formulas(xlsx_path: Path, output_dir: Path):
    wb = load_workbook(xlsx_path, data_only=False, keep_links=False)
    output_dir.mkdir(parents=True, exist_ok=True)

    for sheet in wb.worksheets:
        csv_path = output_dir / f"{sheet.title}.csv"
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            for row in sheet.iter_rows(values_only=False):
                writer.writerow([
                    cell.value if cell.data_type == 'f' else cell.value
                    for cell in row
                ])

def main():
    for path_str in sys.argv[1:]:
        path = Path(path_str)
        if path.suffix.lower() not in {".xlsx", ".xltm", ".xlsm", ".xltx", ".xlsb"}:
            continue

        exploded_root = Path("exploded") / path.stem
        queries_dir = exploded_root / "queries"
        sheets_dir = exploded_root / "sheets"

        print(f"[+] Processing: {path.name}")
        extract_m_scripts(path, queries_dir)
        export_sheets_with_formulas(path, sheets_dir)

if __name__ == "__main__":
    main()
