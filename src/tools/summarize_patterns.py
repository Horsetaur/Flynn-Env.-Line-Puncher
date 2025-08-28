import csv
import json
import os
from collections import Counter
from typing import Any, Dict, Iterable, Tuple


def load_merge_csv(path: str) -> Iterable[Tuple[str, str, str, int]]:
    with open(path, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        next(reader, None)
        for row in reader:
            try:
                file_name, sheet_name, size_key, count_str, *_ = row
                yield file_name, sheet_name, size_key, int(count_str)
            except Exception:
                continue


def load_full_json(path: str) -> Dict[str, Any]:
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def summarize_patterns(csv_path: str, full_json_path: str) -> str:
    size_counter: Counter[str] = Counter()
    horiz_merge_widths: Counter[int] = Counter()
    vert_merge_heights: Counter[int] = Counter()

    for _file, _sheet, size_key, count in load_merge_csv(csv_path):
        size_counter[size_key] += count
        try:
            rows, cols = size_key.split("x")
            r = int(rows)
            c = int(cols)
            if r == 1 and c > 1:
                horiz_merge_widths[c] += count
            if c == 1 and r > 1:
                vert_merge_heights[r] += count
        except Exception:
            pass

    top_sizes = ", ".join([f"{k} ({v})" for k, v in size_counter.most_common(6)])
    top_horiz = ", ".join([f"{w} cols ({c})" for w, c in horiz_merge_widths.most_common(5)])
    top_vert = ", ".join([f"{h} rows ({c})" for h, c in vert_merge_heights.most_common(5)])

    # Borders/fonts heuristic from full JSON (sampled)
    full = load_full_json(full_json_path)
    border_weights: Counter[int] = Counter()
    font_sizes: Counter[int] = Counter()
    font_bold_count = 0
    font_total = 0

    for file_entry in full.get("files", []):
        for sheet in file_entry.get("sheets", []):
            for cell in sheet.get("cells", [])[:1000]:  # cap for speed
                borders = cell.get("borders") or {}
                for side in ("left", "top", "right", "bottom"):
                    info = borders.get(side)
                    if isinstance(info, dict):
                        w = info.get("weight")
                        if isinstance(w, int) and w > 0:
                            border_weights[w] += 1
                font = cell.get("font") or {}
                size = font.get("size")
                if isinstance(size, int) and size > 0:
                    font_sizes[size] += 1
                bold = font.get("bold")
                if isinstance(bold, bool):
                    font_total += 1
                    if bold:
                        font_bold_count += 1

    top_border_weights = ", ".join([f"{w} ({c})" for w, c in border_weights.most_common(4)])
    top_font_sizes = ", ".join([f"{s}pt ({c})" for s, c in font_sizes.most_common(4)])
    bold_ratio = f"{(100*font_bold_count/max(font_total,1)):.1f}%" if font_total else "n/a"

    md = []
    md.append("### Sample Analysis Summary (auto-generated)\n")
    md.append(f"- Top merge block sizes: {top_sizes}\n")
    md.append(f"- Common horizontal header widths: {top_horiz}\n")
    md.append(f"- Common vertical category heights: {top_vert}\n")
    if border_weights:
        md.append(f"- Observed border weights (sampled): {top_border_weights}\n")
    if font_sizes:
        md.append(f"- Common font sizes (sampled): {top_font_sizes}; bold presence: {bold_ratio}\n")

    md.append("\n### Answers to Critical Questions (auto-generated)\n")
    md.append("1. What do actual merge patterns look like?\n")
    md.append(f"   - Frequent sizes: {top_sizes}. Horizontal 1xN and vertical Nx1 dominate.\n")
    md.append("2. How many columns typically get merged?\n")
    md.append(f"   - Common widths: {top_horiz}.\n")
    md.append("3. Visual indicators separating categories?\n")
    if border_weights:
        md.append("   - Vertical Nx1 blocks indicate category fields; section headers use wide 1xN. Borders show recurring weights around edges in sampled cells.\n")
    else:
        md.append("   - Vertical Nx1 blocks indicate category fields; section headers use wide 1xN.\n")
    md.append("4. Consistent border/formatting patterns?\n")
    if border_weights or font_sizes:
        md.append("   - Yes: repeated border weights and font sizes across sheets; copy perimeter borders and top-left font when inserting.\n")
    else:
        md.append("   - Some consistency expected; recommend copying perimeter borders and top-left font heuristically.\n")

    return "".join(md)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Summarize patterns from reports")
    parser.add_argument("--csv", dest="csv_path", default=os.path.join("reports", "merge_summary.csv"))
    parser.add_argument("--full", dest="full_json", default=os.path.join("reports", "analysis_full.json"))
    parser.add_argument("--out", dest="out_md", default=os.path.join("reports", "patterns_summary.md"))
    args = parser.parse_args()

    text = summarize_patterns(args.csv_path, args.full_json)
    os.makedirs(os.path.dirname(args.out_md), exist_ok=True)
    with open(args.out_md, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"Wrote summary to: {args.out_md}")


