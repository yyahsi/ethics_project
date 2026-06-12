from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from comparison_tables import build_pairs_from_directory, create_side_by_side_workbook


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate Israeli vs Russian Crimes ethicality/visibility/accountability comparison tables.")
    parser.add_argument(
        "--inputs",
        default="data/comparison_inputs",
        help="Folder containing ChatGPT/Claude/Gemini Israel and Russia files.",
    )
    parser.add_argument(
        "--output",
        default="israeli_vs_russian_crimes.xlsx",
        help="Output XLSX path.",
    )
    parser.add_argument("--rows", type=int, default=200, help="Number of rows to show in every comparison sheet.")
    parser.add_argument(
        "--ratings",
        default="data/ratings.csv",
        help="Optional incident ratings CSV/XLSX used for the second Final Comparison table.",
    )
    args = parser.parse_args()

    pairs, source_info = build_pairs_from_directory(args.inputs)
    ratings_path = Path(args.ratings)
    incident_ratings = None
    if ratings_path.exists():
        if ratings_path.suffix.lower() in {".xlsx", ".xls"}:
            incident_ratings = pd.read_excel(ratings_path)
        else:
            incident_ratings = pd.read_csv(ratings_path, encoding="utf-8-sig")
    workbook_bytes = create_side_by_side_workbook(
        pairs,
        source_info=source_info,
        max_rows=args.rows,
        incident_ratings=incident_ratings,
    )
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(workbook_bytes)
    print(f"Wrote {output_path}")


if __name__ == "__main__":
    main()
