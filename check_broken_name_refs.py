from __future__ import annotations

import csv
import sys
from pathlib import Path


def main() -> int:
    names_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("exploded/TUS/names.tsv")

    if not names_path.exists():
        print(f"ERROR: names file not found: {names_path}")
        return 2

    broken: list[tuple[str, str, str]] = []

    with names_path.open("r", newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle, delimiter="\t")
        for row in reader:
            name = (row.get("Name") or "").strip()
            refers_to = (row.get("Refers To") or "").strip()
            scope = (row.get("Scope") or "").strip()
            if "#REF!" in refers_to:
                broken.append((name, refers_to, scope))

    if not broken:
        print(f"OK: no broken Name Manager references found in {names_path}")
        return 0

    print(f"ERROR: found {len(broken)} broken Name Manager references in {names_path}:")
    for name, refers_to, scope in broken:
        print(f"  - {name}\t{refers_to}\t{scope}")

    return 1


if __name__ == "__main__":
    raise SystemExit(main())
