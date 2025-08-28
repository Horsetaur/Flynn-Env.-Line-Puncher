import argparse
import os
from analyzer.excel_pattern_analyzer import ExcelPatternAnalyzer, write_json_report
from excel_connector import ExcelConnector
from row_inserter import RowInserter
from gui.gui_interface import LinePuncherGUI


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Analyze Excel merge/border patterns")
    parser.add_argument(
        "--dir",
        dest="directory",
        default="Base Case Files",
        help="Directory containing Excel files (.xls/.xlsx/.xlsm)",
    )
    parser.add_argument(
        "--out",
        dest="out_path",
        default=os.path.join("reports", "analysis.json"),
        help="Output JSON report path",
    )
    parser.add_argument(
        "--max-cells",
        dest="max_cells",
        type=int,
        default=2000,
        help="Maximum sampled cells per sheet",
    )
    parser.add_argument(
        "--borders",
        dest="include_borders",
        action="store_true",
        help="Include border and font extraction (slower)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
    target_dir = args.directory
    if not os.path.isabs(target_dir):
        target_dir = os.path.join(repo_root, target_dir)

    out_path = args.out_path
    if not os.path.isabs(out_path):
        out_path = os.path.join(repo_root, out_path)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    # Analysis mode when directory exists
    analyzer = ExcelPatternAnalyzer(directory_path=target_dir, include_borders=args.include_borders)
    results = analyzer.analyze(max_cells_per_sheet=args.max_cells)
    write_json_report(results, out_path)
    print(f"Wrote report to: {out_path}")

    # Minimal GUI wiring (optional run)
    try:
        conn = ExcelConnector()
        inserter = RowInserter()

        def on_add_row() -> None:
            _, ws, cell = conn.get_active_cell()
            inserter.add_row_to_category(ws, int(cell.Row))

        def on_add_category() -> None:
            _, ws, cell = conn.get_active_cell()
            inserter.add_new_category(ws, int(cell.Row))

        # Comment out next line if you don't want GUI auto-run
        # LinePuncherGUI(on_add_row, on_add_category).run()
    except Exception:
        pass


if __name__ == "__main__":
    main()


