import os
import shutil

import pytest

from sorter_utils import (
    clean_name,
    normalize_path,
    safe_copy,
    strip_long_path_prefix,
    unique_destination,
    win_long_path,
)


class TestCleanName:
    def test_strips_illegal_chars(self):
        assert clean_name('a<b>:"c/d\\e|f?g*h') == "abcdefgh"

    def test_underscores_become_spaces(self):
        assert clean_name("factura_raiffeisen_2026") == "factura raiffeisen 2026"

    def test_collapses_whitespace_and_trims(self):
        assert clean_name("  a   b  ") == "a b"

    def test_max_len(self):
        assert len(clean_name("x" * 200, max_len=65)) == 65

    def test_trailing_dots_and_commas_removed(self):
        assert clean_name("nume...") == "nume"
        assert clean_name("nume, ,") == "nume"

    def test_empty(self):
        assert clean_name("") == ""
        assert clean_name(None) == ""

    def test_keeps_romanian_diacritics(self):
        assert clean_name("Factură Raiffeisen") == "Factură Raiffeisen"


class TestPaths:
    def test_normalize_path(self):
        assert normalize_path("Z:\\To be Sorted V1\\A.TXT") == os.path.normcase(
            os.path.normpath("Z:\\To be Sorted V1\\A.TXT")
        )

    @pytest.mark.skipif(os.name == "nt", reason="pe Linux nu se adaugă prefix")
    def test_win_long_path_posix_returns_abspath(self, tmp_path):
        assert win_long_path(tmp_path / "a.txt") == os.path.abspath(str(tmp_path / "a.txt"))

    def test_strip_long_path_prefix(self):
        assert strip_long_path_prefix("\\\\?\\Z:\\Sorted\\a.txt") == "Z:\\Sorted\\a.txt"
        assert strip_long_path_prefix("Z:\\Sorted\\a.txt") == "Z:\\Sorted\\a.txt"


class TestSafeCopy:
    def test_copies_content(self, tmp_path):
        src = tmp_path / "src.txt"
        src.write_text("continut", encoding="utf-8")
        out = safe_copy(src, tmp_path / "out", "nume.txt")
        assert out.name == "nume.txt"
        assert out.read_text(encoding="utf-8") == "continut"
        assert src.exists()  # original păstrat

    def test_collision_gets_counter(self, tmp_path):
        src = tmp_path / "src.txt"
        src.write_text("x", encoding="utf-8")
        (tmp_path / "out" / "nume.txt").parent.mkdir(parents=True)
        (tmp_path / "out" / "nume.txt").write_text("vechi", encoding="utf-8")
        out = safe_copy(src, tmp_path / "out", "nume.txt")
        assert out.name == "nume (1).txt"
        assert (tmp_path / "out" / "nume.txt").read_text(encoding="utf-8") == "vechi"

    def test_unique_destination_never_overwrites(self, tmp_path):
        folder = tmp_path / "d"
        folder.mkdir()
        (folder / "a.txt").touch()
        (folder / "a (1).txt").touch()
        assert unique_destination(folder, "a", ".txt").name == "a (2).txt"

    def test_fallback_short_name_on_failure(self, tmp_path, monkeypatch):
        src = tmp_path / "src.bin"
        src.write_bytes(b"date" * 10)
        real_copy = shutil.copy2
        calls = {"n": 0}

        def fake_copy(a, b):
            calls["n"] += 1
            if calls["n"] == 1:
                raise OSError("simulat: cale prea lungă")
            return real_copy(a, b)

        monkeypatch.setattr("sorter_utils.shutil.copy2", fake_copy)
        long_name = "nume foarte lung " + "x" * 100 + ".bin"
        out = safe_copy(src, tmp_path / "out", long_name)
        assert out.exists()
        assert out.name.endswith(".bin")
        assert len(out.stem) <= 35
        assert out.read_bytes() == b"date" * 10

    def test_illegal_chars_in_desired_name(self, tmp_path):
        src = tmp_path / "s.txt"
        src.write_text("ok", encoding="utf-8")
        out = safe_copy(src, tmp_path / "out", 'evil<>:"name.txt')
        assert out.exists()
        assert set(out.name) & set('<>:"') == set()
