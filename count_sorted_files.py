r"""Statistici și audit pentru datasetul sortat.

Scanează DEST_DIR și generează un CSV cu: sumar general, statistici pe extensii,
pe categorii principale, pe subfoldere (subiecte) și top 20 fișiere ca dimensiune.

Utilizare:
    python count_sorted_files.py                          # Z:\\Sorted -> Z:\\Sorted\\statistici_fisiere.csv
    python count_sorted_files.py --target D:\Sorted --output D:\raport.csv
"""

import argparse
import csv
import os
from collections import defaultdict
from pathlib import Path

from sorter_utils import strip_long_path_prefix, win_long_path

DEFAULT_TARGET = r"Z:\Sorted"
IGNORE_FILES = {"sorting_log.csv", "sorting_errors.csv", "statistici_fisiere.csv"}
TOP_N = 20


def format_size_mb(size_bytes: int) -> float:
    return round(size_bytes / (1024 * 1024), 2)


def format_size_gb(size_bytes: int) -> float:
    return round(size_bytes / (1024 * 1024 * 1024), 3)


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Statistici pentru folderul sortat.")
    p.add_argument("--target", default=DEFAULT_TARGET, help="Folderul de analizat (default: %(default)s)")
    p.add_argument("--output", default=None, help="CSV de ieșire (default: <target>/statistici_fisiere.csv)")
    return p.parse_args(argv)


def scan_target(target_dir: str):
    root_long = win_long_path(Path(target_dir))
    if not os.path.exists(root_long):
        raise FileNotFoundError(f"Folderul {target_dir} nu a fost găsit!")

    ext_stats = defaultdict(lambda: {"count": 0, "size": 0})
    cat_stats = defaultdict(lambda: {"count": 0, "size": 0, "extensions": set()})
    sub_stats = defaultdict(lambda: {"count": 0, "size": 0})
    biggest = []  # (size, cale_relativa)

    total_files = 0
    total_size = 0

    for root, dirs, files in os.walk(root_long):
        rel_root = os.path.relpath(root, root_long)
        parts = [] if rel_root == "." else rel_root.split(os.sep)
        categorie = parts[0] if len(parts) >= 1 else "Root (Fără Categorie)"
        subiect = parts[1] if len(parts) >= 2 else "(Direct în categorie)"

        for file in files:
            if file.lower() in IGNORE_FILES:
                continue
            full_path = os.path.join(root, file)
            try:
                file_size = os.path.getsize(full_path)
            except OSError:
                file_size = 0

            ext = os.path.splitext(file)[1].lower() or "(fără extensie)"
            total_files += 1
            total_size += file_size
            ext_stats[ext]["count"] += 1
            ext_stats[ext]["size"] += file_size
            cat_stats[categorie]["count"] += 1
            cat_stats[categorie]["size"] += file_size
            cat_stats[categorie]["extensions"].add(ext)
            sub_stats[(categorie, subiect)]["count"] += 1
            sub_stats[(categorie, subiect)]["size"] += file_size
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

        writer.writerow(["=== SUMAR GENERAL ==="])
        writer.writerow(["Total Fișiere Sortate", "Total Tipuri de Date (Extensii)", "Dimensiune Totală (MB)", "Dimensiune Totală (GB)"])
        writer.writerow([total_files, len(stats["ext_stats"]), format_size_mb(total_size), format_size_gb(total_size)])
        writer.writerow([])

        writer.writerow(["=== 1. STATISTICI PE TIPURI DE DATE (EXTENSII) ==="])
        writer.writerow(["Extensie / Tip Fișier", "Număr Fișiere", "Procent din Total (%)", "Dimensiune (MB)", "Dimensiune (GB)"])
        for ext, data in sorted(stats["ext_stats"].items(), key=lambda x: x[1]["count"], reverse=True):
            pct = round(data["count"] / total_files * 100, 2) if total_files else 0
            writer.writerow([ext, data["count"], pct, format_size_mb(data["size"]), format_size_gb(data["size"])])
        writer.writerow([])

        writer.writerow(["=== 2. STATISTICI PE CATEGORII PRINCIPALE ==="])
        writer.writerow(["Categorie", "Număr Fișiere", "Procent din Total (%)", "Dimensiune (MB)", "Dimensiune (GB)", "Tipuri de Extensii Conținute"])
        for cat, data in sorted(stats["cat_stats"].items(), key=lambda x: x[1]["count"], reverse=True):
            pct = round(data["count"] / total_files * 100, 2) if total_files else 0
            ext_list = ", ".join(sorted(data["extensions"]))
            writer.writerow([cat, data["count"], pct, format_size_mb(data["size"]), format_size_gb(data["size"]), ext_list])
        writer.writerow([])

        writer.writerow(["=== 3. STATISTICI DETALIATE PE SUBFOLDERE ==="])
        writer.writerow(["Categorie", "Subfolder (Subiect)", "Număr Fișiere", "Dimensiune (MB)"])
        for (cat, sub), data in sorted(stats["sub_stats"].items(), key=lambda x: (x[0][0], -x[1]["count"])):
            writer.writerow([cat, sub, data["count"], format_size_mb(data["size"])])
        writer.writerow([])

        writer.writerow([f"=== 4. TOP {TOP_N} FIȘIERE CA DIMENSIUNE ==="])
        writer.writerow(["Cale Relativă", "Dimensiune (MB)", "Dimensiune (GB)"])
        for size, rel in stats["biggest"]:
            writer.writerow([rel, format_size_mb(size), format_size_gb(size)])


def main(argv=None) -> int:
    args = parse_args(argv)
    output_csv = args.output or str(Path(args.target) / "statistici_fisiere.csv")

    print(f"Scanez rapid toate fișierele din {args.target}...")
    stats = scan_target(args.target)
    write_report(stats, output_csv)

    print("=" * 60)
    print(f"FINALIZAT! S-au analizat {stats['total_files']} fișiere ({format_size_gb(stats['total_size'])} GB).")
    print(f"Au fost identificate {len(stats['ext_stats'])} tipuri diferite de fișiere (extensii).")
    print(f"Categorii: {len(stats['cat_stats'])}, subfoldere: {len(stats['sub_stats'])}.")
    print(f"Raportul CSV a fost salvat în: {output_csv}")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
