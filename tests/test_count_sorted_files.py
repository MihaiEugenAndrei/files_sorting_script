import csv

import pytest

from count_sorted_files import main, scan_target


def build_tree(tmp_path):
    dest = tmp_path / "sorted"
    (dest / "Financiar" / "Facturi").mkdir(parents=True)
    (dest / "Financiar").mkdir(parents=True, exist_ok=True)
    (dest / "Unknown").mkdir(parents=True)
    (dest / "Financiar" / "Facturi" / "a.txt").write_bytes(b"x" * 1000)
    (dest / "Financiar" / "b.pdf").write_bytes(b"y" * 2048)
    (dest / "Unknown" / "c.txt").write_bytes(b"z" * 10)
    (dest / "sorting_log.csv").write_text("header\n", encoding="utf-8")  # trebuie ignorat
    return dest


def read_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return [row for row in csv.reader(f) if row]


def test_scan_and_report(tmp_path):
    dest = build_tree(tmp_path)
    out = tmp_path / "raport.csv"
    rc = main(["--target", str(dest), "--output", str(out)])
    assert rc == 0

    rows = read_csv(out)
    flat = {tuple(r[:2]) for r in rows}

    # sumar general: 3 fișiere (log-ul e ignorat), 2 extensii, 3058 bytes
    summary = next(r for r in rows if r and r[0] == "Total Sorted Files")
    assert summary is not None
    total_row = rows[rows.index(summary) + 1]
    assert total_row[0] == "3"
    assert total_row[1] == "2"

    # sectiunea extensii
    assert (".txt", "2") in flat or any(r[0] == ".txt" and r[1] == "2" for r in rows)
    assert any(r[0] == ".pdf" and r[1] == "1" for r in rows)

    # categorii
    assert any(r[0] == "Financiar" and r[1] == "2" for r in rows)
    assert any(r[0] == "Unknown" and r[1] == "1" for r in rows)

    # subfoldere
    assert any(r[0] == "Financiar" and r[1] == "Facturi" and r[2] == "1" for r in rows)

    # top fisiere: b.pdf (2048) primul
    top_header = next(r for r in rows if r and r[0].startswith("=== 4."))
    top_idx = rows.index(top_header)
    assert rows[top_idx + 2][0].endswith("b.pdf")


def test_scan_missing_target(tmp_path):
    with pytest.raises(FileNotFoundError):
        scan_target(str(tmp_path / "nu_exista"))


def test_default_output_name(tmp_path):
    dest = build_tree(tmp_path)
    rc = main(["--target", str(dest)])
    assert rc == 0
    assert (dest / "statistici_fisiere.csv").is_file()
