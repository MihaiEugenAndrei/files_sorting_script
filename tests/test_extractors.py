import zipfile

import docx
import fitz
import openpyxl
from pptx import Presentation

from sort_files import (
    extract_file_payload,
    extract_strings_from_binary,
    parse_torrent_file,
    read_text_file,
)


def test_read_text_file_utf8(tmp_path):
    f = tmp_path / "a.txt"
    f.write_text("Factură Raiffeisen — ă î ș ț", encoding="utf-8")
    assert "Raiffeisen" in read_text_file(f)


def test_read_text_file_cp1252_fallback(tmp_path):
    f = tmp_path / "b.txt"
    f.write_bytes(b"Factur\xe3 pentru servicii prestate")  # 0xE3 invalid in utf-8, valid cp1252
    text = read_text_file(f)
    assert "Factur" in text


def test_extract_txt(tmp_path):
    f = tmp_path / "doc.txt"
    f.write_text("Contract de închiriere apartament", encoding="utf-8")
    mode, payload = extract_file_payload(f)
    assert mode == "text"
    assert "Contract" in payload


def test_extract_pdf_with_text(tmp_path):
    f = tmp_path / "factura.pdf"
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), "Factura Raiffeisen Bank nr 100 plata utilitati scadenta")
    doc.save(str(f))
    doc.close()
    mode, payload = extract_file_payload(f)
    assert mode == "text"
    assert "Raiffeisen" in payload


def test_extract_pdf_scan_becomes_image(tmp_path):
    f = tmp_path / "scan.pdf"
    doc = fitz.open()
    doc.new_page()  # pagina goala = fara text selectabil -> OCR vizual
    doc.save(str(f))
    doc.close()
    mode, payload = extract_file_payload(f)
    assert mode == "image"
    assert isinstance(payload, str) and len(payload) > 100  # base64


def test_extract_docx(tmp_path):
    f = tmp_path / "cv.docx"
    d = docx.Document()
    d.add_paragraph("Curriculum Vitae Ion Popescu software developer")
    d.save(str(f))
    mode, payload = extract_file_payload(f)
    assert mode == "text"
    assert "Popescu" in payload


def test_extract_pptx(tmp_path):
    f = tmp_path / "prezentare.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Prezentare Proiect Marketing Q3"
    prs.save(str(f))
    mode, payload = extract_file_payload(f)
    assert mode == "text"
    assert "Marketing" in payload


def test_extract_xlsx(tmp_path):
    f = tmp_path / "tabel.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Banca"
    ws["B1"] = "Raiffeisen"
    ws["A2"] = "Suma"
    ws["B2"] = 1500
    wb.save(str(f))
    mode, payload = extract_file_payload(f)
    assert mode == "text"
    assert "Raiffeisen" in payload


def test_extract_zip(tmp_path):
    f = tmp_path / "arhiva.zip"
    with zipfile.ZipFile(f, "w") as z:
        z.writestr("facturi/ianuarie.pdf", "x")
        z.writestr("facturi/februarie.pdf", "y")
    mode, payload = extract_file_payload(f)
    assert mode == "text"
    assert "ianuarie.pdf" in payload


def test_parse_torrent(tmp_path):
    f = tmp_path / "film.torrent"
    name = b"My Cool Movie 2026"
    data = b"d8:announce15:http://tracker4:name" + str(len(name)).encode() + b":" + name + b"e"
    f.write_bytes(data)
    text = parse_torrent_file(f)
    assert "My Cool Movie 2026" in text


def test_extract_binary_strings(tmp_path):
    f = tmp_path / "setup.exe"
    payload = b"\x00\x01MZ" + b"SomeInstallerName Setup" + b"\x00" + "Companie SRL".encode("utf-16le") + b"\x00" * 20
    f.write_bytes(payload)
    mode, text = extract_file_payload(f)
    assert mode == "text"
    assert "SomeInstallerName" in text
    assert "Companie" in text


def test_extract_strings_from_binary_filters_hex(tmp_path):
    f = tmp_path / "x.bin"
    f.write_bytes(b"deadbeef" * 10 + b"\x00\x01" + b"Real Useful String Here")
    text = extract_strings_from_binary(f)
    assert "deadbeef" not in text
    assert "Real Useful String" in text


def test_unknown_extension_no_strings(tmp_path):
    f = tmp_path / "mystery.xyz"
    f.write_bytes(bytes(range(0, 16)) * 50)
    mode, payload = extract_file_payload(f)
    assert mode == "unknown"
    assert payload is None
