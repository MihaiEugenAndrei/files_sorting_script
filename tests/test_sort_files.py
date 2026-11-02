import csv
import os
from types import SimpleNamespace

import pytest

import sort_files
from sort_files import junk_reason, load_processed, main, resolve_target_folder


def make_client_router():
    """Client fake care decide răspunsul după conținutul payload-ului."""
    calls = []

    def create(**kwargs):
        calls.append(kwargs)
        user = kwargs["messages"][1]["content"]
        text = user if isinstance(user, str) else user[0]["text"]
        if "Raiffeisen" in text:
            content = '{"categorie": "Financiar", "subiect": "Facturi", "nume_nou": "Factura Raiffeisen Utilitati", "pastreaza_original": false}'
        elif "Popescu" in text:
            content = '{"categorie": "Muncă și Carieră", "subiect": "CV", "nume_nou": "CV Ion Popescu", "pastreaza_original": false}'
        else:
            content = '{"categorie": "Unknown", "subiect": "", "nume_nou": "Necunoscut", "pastreaza_original": true}'
        msg = SimpleNamespace(content=content)
        return SimpleNamespace(choices=[SimpleNamespace(message=msg)])

    client = SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(create=create)), calls=calls)
    return client


def make_source(tmp_path):
    src = tmp_path / "in"
    (src / "node_modules" / "pkg").mkdir(parents=True)
    (src / "factura_scan.txt").write_text(
        "Factura Raiffeisen Bank nr 100, plata utilitati, suma 250 RON, scadenta 2026. "
        "Detalii suplimentare despre plata facturii de utilitati pentru luna curenta. " * 3,
        encoding="utf-8",
    )
    (src / "cv ion.txt").write_text(
        "Curriculum Vitae Ion Popescu, software developer cu 10 ani experienta in Python. "
        "Experienta in proiecte enterprise si echipe internationale. " * 3,
        encoding="utf-8",
    )
    (src / "junk.tmp").write_text("x", encoding="utf-8")           # extensie junk
    (src / "tiny.txt").write_text("hi", encoding="utf-8")          # < 200 bytes
    (src / "node_modules" / "pkg" / "index.js").write_text("console.log(1)", encoding="utf-8")  # dir ignorat
    return src


def read_log(log_file):
    with open(log_file, encoding="utf-8-sig", newline="") as f:
        return list(csv.reader(f))


class TestJunkReason:
    def test_junk_extension(self, tmp_path):
        f = tmp_path / "a.log"
        f.write_text("x" * 500)
        assert junk_reason(f) == "extensie_junk"

    def test_too_small(self, tmp_path):
        f = tmp_path / "a.txt"
        f.write_text("hi")
        assert junk_reason(f) == "prea_mic"

    def test_ignored_dir(self, tmp_path):
        d = tmp_path / "node_modules"
        d.mkdir()
        f = d / "a.js"
        f.write_text("x" * 500)
        assert junk_reason(f) == "director_ignorat"

    def test_normal_file(self, tmp_path):
        f = tmp_path / "a.txt"
        f.write_text("x" * 500)
        assert junk_reason(f) is None


class TestResolveTargetFolder:
    def test_category_and_subject(self, tmp_path):
        assert resolve_target_folder(tmp_path, {"categorie": "Financiar", "subiect": "Facturi"}) == tmp_path / "Financiar" / "Facturi"

    def test_empty_category_falls_back_to_unknown(self, tmp_path):
        assert resolve_target_folder(tmp_path, {"categorie": "", "subiect": ""}) == tmp_path / "Unknown"

    def test_illegal_chars_stripped(self, tmp_path):
        folder = resolve_target_folder(tmp_path, {"categorie": 'Fi<n:a', "subiect": "a/b"})
        assert "/" not in folder.name and "<" not in folder.name


class TestEndToEnd:
    def test_full_run_sort_junk_resume(self, tmp_path):
        src = make_source(tmp_path)
        dest = tmp_path / "out"
        client = make_client_router()

        rc = main(["--source", str(src), "--dest", str(dest), "--model", "fake"], client=client)
        assert rc == 0
        assert len(client.calls) == 2  # doar cele 2 fișiere reale trec prin AI

        facturi = dest / "Financiar" / "Facturi"
        assert (facturi / "Factura Raiffeisen Utilitati (factura scan).txt").is_file()
        cv_dir = dest / "Muncă și Carieră" / "CV"
        assert (cv_dir / "CV Ion Popescu (cv ion).txt").is_file()

        deleted = dest / "to be deleted"
        assert (deleted / "junk.tmp").is_file()
        assert (deleted / "tiny.txt").is_file()
        assert (deleted / "index.js").is_file()

        # originalele sunt intacte (non-destructiv)
        assert (src / "factura_scan.txt").is_file()
        assert (src / "cv ion.txt").is_file()

        # log: header + 5 intrări
        rows = read_log(dest / "sorting_log.csv")
        assert rows[0] == ["Cale_Originala", "Mod", "Categorie", "Subiect", "Cale_Noua"]
        assert len(rows) == 6

        # RESUME: a doua rulare nu mai apelează AI și nu mai copiază nimic
        client2 = make_client_router()
        rc = main(["--source", str(src), "--dest", str(dest), "--model", "fake"], client=client2)
        assert rc == 0
        assert len(client2.calls) == 0
        assert len(read_log(dest / "sorting_log.csv")) == 6

    def test_resume_with_different_casing(self, tmp_path):
        src = make_source(tmp_path)
        dest = tmp_path / "out"
        client = make_client_router()
        main(["--source", str(src), "--dest", str(dest), "--model", "fake"], client=client)

        processed = load_processed(str(dest / "sorting_log.csv"))
        # pe Windows normcase face match case-insensitive; verificăm normalizarea în sine
        some = next(iter(processed))
        assert some == os.path.normcase(os.path.normpath(some))

    def test_failed_copy_not_marked_processed_and_retried(self, tmp_path, monkeypatch):
        src = make_source(tmp_path)
        dest = tmp_path / "out"
        client = make_client_router()

        real_safe_copy = sort_files.safe_copy
        failing = {"armed": True}

        def flaky_copy(s, folder, name, **kw):
            if failing["armed"] and "factura" in str(s):
                raise OSError("disc plin simulat")
            return real_safe_copy(s, folder, name, **kw)

        monkeypatch.setattr(sort_files, "safe_copy", flaky_copy)
        rc = main(["--source", str(src), "--dest", str(dest), "--model", "fake"], client=client)
        assert rc == 0

        # fișierul eșuat NU e în log-ul de procesate, dar E în cel de erori
        log_rows = read_log(dest / "sorting_log.csv")
        assert not any("factura_scan" in r[0] for r in log_rows)
        with open(dest / "sorting_errors.csv", encoding="utf-8-sig", newline="") as f:
            err_rows = list(csv.reader(f))
        assert any("factura_scan" in r[0] for r in err_rows)

        # reîncercarea reușește după ce "discul" e reparat
        failing["armed"] = False
        client2 = make_client_router()
        main(["--source", str(src), "--dest", str(dest), "--model", "fake"], client=client2)
        assert (dest / "Financiar" / "Facturi" / "Factura Raiffeisen Utilitati (factura scan).txt").is_file()
        assert len(client2.calls) == 1  # doar fișierul nereușit a fost reîncercat

    def test_dry_run_copies_nothing(self, tmp_path, capsys):
        src = make_source(tmp_path)
        dest = tmp_path / "out"
        client = make_client_router()
        rc = main(["--source", str(src), "--dest", str(dest), "--model", "fake", "--dry-run"], client=client)
        assert rc == 0
        assert len(client.calls) == 2
        out = capsys.readouterr().out
        assert "DRY-RUN" in out
        assert not (dest / "sorting_log.csv").exists()
        copied = [p for p in dest.rglob("*") if p.is_file()]
        assert copied == []

    def test_limit(self, tmp_path):
        src = make_source(tmp_path)
        dest = tmp_path / "out"
        client = make_client_router()
        rc = main(["--source", str(src), "--dest", str(dest), "--model", "fake", "--limit", "1"], client=client)
        assert rc == 0
        # doar primul fișier neprocesat a fost procesat (junk sau AI, în ordinea rglob)
        rows = read_log(dest / "sorting_log.csv")
        assert len(rows) == 2  # header + 1
