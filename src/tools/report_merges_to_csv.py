import csv
import json
import os
from typing import Dict, Any, List


def load_report(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def write_merge_summary_csv(report: Dict[str, Any], out_csv: str) -> None:
    rows: List[List[Any]] = []
    rows.append(["file", "sheet", "merge_block_size", "count", "used_rows", "used_cols"])

    for file_entry in report.get("files", []):
        file_path = file_entry.get("file", "")
        for sheet in file_entry.get("sheets", []):
            block_sizes: Dict[str, int] = sheet.get("merge_blocks_summary", {}).get("block_sizes", {})
            used_rows = sheet.get("used_rows", 0)
            used_cols = sheet.get("used_cols", 0)
            for size_key, count in block_sizes.items():
                rows.append([
                    os.path.basename(file_path),
                    sheet.get("name", ""),
                    size_key,
                    count,
                    used_rows,
                    used_cols,
                ])

    os.makedirs(os.path.dirname(out_csv), exist_ok=True)
    with open(out_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerows(rows)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Write merge summary CSV from analysis.json")
    parser.add_argument("--in", dest="in_path", default=os.path.join("reports", "analysis.json"))
    parser.add_argument("--out", dest="out_path", default=os.path.join("reports", "merge_summary.csv"))
    args = parser.parse_args()

    report = load_report(args.in_path)
    write_merge_summary_csv(report, args.out_path)
    print(f"Wrote CSV to: {args.out_path}")


