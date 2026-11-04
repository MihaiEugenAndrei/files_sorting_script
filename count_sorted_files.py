r"""Statistics and audit for the sorted dataset.

Scans TARGET_DIR and writes a CSV report with: general summary, stats by extension,
by main category, by subfolder (subject), and the top 20 files by size.

Usage:
    python count_sorted_files.py                          # Z:\Sorted -> Z:\Sorted\statistici_fisiere.csv
    python count_sorted_files.py --target D:\Sorted --output D:\report.csv
"""

import argparse
import csv
import os
from collections import defaultdict
from pathlib import Path

from sorter_utils import strip_long_path_prefix, win_long_path

DEFAULT_TARGET = r"Z:\Sorted"
# Existing artifact filenames on the sorted drive — kept as-is for compatibility.
IGNORE_FILES = {"sorting_log.csv", "sorting_errors.csv", "statistici_fisiere.csv"}
TOP_N = 20


def format_size_mb(size_bytes: int) -> float:
    return round(size_bytes / (1024 * 1024), 2)


def format_size_gb(size_bytes: int) -> float:
    return round(size_bytes / (1024 * 1024 * 1024), 3)


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Statistics for the sorted folder.")
    p.add_argument("--target", default=DEFAULT_TARGET, help="Folder to analyze (default: %(default)s)")
    p.add_argument("--output", default=None, help="Output CSV (default: <target>/statistici_fisiere.csv)")
    return p.parse_args(argv)


def scan_target(target_dir: str):
    root_long = win_long_path(Path(target_dir))
    if not os.path.exists(root_long):
        raise FileNotFoundError(f"Folder {target_dir} not found!")

    ext_stats = defaultdict(lambda: {"count": 0, "size": 0})
    cat_stats = defaultdict(lambda: {"count": 0, "size": 0, "extensions": set()})
    sub_stats = defaultdict(lambda: {"count": 0, "size": 0})
    biggest = []  # (size, relative path)

    total_files = 0
    total_size = 0

    for root, dirs, files in os.walk(root_long):
        rel_root = os.path.relpath(root, root_long)
        parts = [] if rel_root == "." else rel_root.split(os.sep)
        category = parts[0] if len(parts) >= 1 else "Root (No Category)"
        subject = parts[1] if len(parts) >= 2 else "(Direct in category)"

        for file in files:
            if file.lower() in IGNORE_FILES:
                continue
            full_path = os.path.join(root, file)
            try:
                file_size = os.path.getsize(full_path)
            except OSError:
                file_size = 0

            ext = os.path.splitext(file)[1].lower() or "(no extension)"
            total_files += 1
            total_size += file_size
            ext_stats[ext]["count"] += 1
            ext_stats[ext]["size"] += file_size
            cat_stats[category]["count"] += 1
            cat_stats[category]["size"] += file_size
            cat_stats[category]["extensions"].add(ext)
            sub_stats[(category, subject)]["count"] += 1
            sub_stats[(category, subject)]["size"] += file_size
            rel_file = os.path.join(rel_root, file) if rel_root != "." else file
            biggest.append((file_size, strip_long_path_prefix(rel_file)))

    biggest.sort(reverse=True)
    return {
        "total_files": total_files,
        "total_size": total_size,
        "ext_stats": ext_stats,
        "cat_stats": cat_stats,
        "sub_stats": sub_stats,
        "biggest": biggest[:TOP_N],
    }


def write_report(stats: dict, output_csv: str):
    total_files = stats["total_files"]
    total_size = stats["total_size"]

    with open(output_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)

        writer.writerow(["=== GENERAL SUMMARY ==="])
        writer.writerow(["Total Sorted Files", "Distinct File Types (Extensions)", "Total Size (MB)", "Total Size (GB)"])
        writer.writerow([total_files, len(stats["ext_stats"]), format_size_mb(total_size), format_size_gb(total_size)])
        writer.writerow([])

        writer.writerow(["=== 1. STATS BY FILE TYPE (EXTENSION) ==="])
        writer.writerow(["Extension / File Type", "File Count", "Percent of Total (%)", "Size (MB)", "Size (GB)"])
        for ext, data in sorted(stats["ext_stats"].items(), key=lambda x: x[1]["count"], reverse=True):
            pct = round(data["count"] / total_files * 100, 2) if total_files else 0
            writer.writerow([ext, data["count"], pct, format_size_mb(data["size"]), format_size_gb(data["size"])])
        writer.writerow([])

        writer.writerow(["=== 2. STATS BY MAIN CATEGORY ==="])
        writer.writerow(["Category", "File Count", "Percent of Total (%)", "Size (MB)", "Size (GB)", "Contained Extensions"])
        for cat, data in sorted(stats["cat_stats"].items(), key=lambda x: x[1]["count"], reverse=True):
            pct = round(data["count"] / total_files * 100, 2) if total_files else 0
            ext_list = ", ".join(sorted(data["extensions"]))
            writer.writerow([cat, data["count"], pct, format_size_mb(data["size"]), format_size_gb(data["size"]), ext_list])
        writer.writerow([])

        writer.writerow(["=== 3. DETAILED STATS BY SUBFOLDER ==="])
        writer.writerow(["Category", "Subfolder (Subject)", "File Count", "Size (MB)"])
        for (cat, sub), data in sorted(stats["sub_stats"].items(), key=lambda x: (x[0][0], -x[1]["count"])):
            writer.writerow([cat, sub, data["count"], format_size_mb(data["size"])])
        writer.writerow([])

        writer.writerow([f"=== 4. TOP {TOP_N} FILES BY SIZE ==="])
        writer.writerow(["Relative Path", "Size (MB)", "Size (GB)"])
        for size, rel in stats["biggest"]:
            writer.writerow([rel, format_size_mb(size), format_size_gb(size)])


def main(argv=None) -> int:
    args = parse_args(argv)
    output_csv = args.output or str(Path(args.target) / "statistici_fisiere.csv")

    print(f"Scanning all files in {args.target}...")
    stats = scan_target(args.target)
    write_report(stats, output_csv)

    print("=" * 60)
    print(f"DONE! Analyzed {stats['total_files']} files ({format_size_gb(stats['total_size'])} GB).")
    print(f"Found {len(stats['ext_stats'])} distinct file types (extensions).")
    print(f"Categories: {len(stats['cat_stats'])}, subfolders: {len(stats['sub_stats'])}.")
    print(f"CSV report saved to: {output_csv}")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
