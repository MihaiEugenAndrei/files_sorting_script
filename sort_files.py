r"""AI file sorting pipeline (non-destructive copy + local VLM renaming).

Citeste fisierele din SOURCE_DIR, extrage continutul (text/OCR/imagine), cere unui
VLM local (LM Studio / Bionic, API compatibil OpenAI) o clasificare JSON, si copiaza
fiecare fisier in DEST_DIR/<Categorie>/<Subiect>/<nume_nou> (<stem original>).ext.
Originalele nu sunt niciodata modificate sau sterse.

Siguranta la crash: fiecare fisier copiat cu succes este logat imediat in CSV.
La repornire, fisierele deja logate sunt sarite instant. Fisierele la care copierea
a esuat NU sunt marcate ca procesate (sunt in sorting_errors.csv si se reincearca).

Utilizare:
    python sort_files.py                          # configuratia implicita (Z:\\...)
    python sort_files.py --source D:\in --dest D:\out
    python sort_files.py --dry-run --limit 50     # doar propune mutarile, nu copiaza
"""

import argparse
import base64
import csv
import json
import os
import re
import sys
import zipfile
from io import BytesIO
from pathlib import Path

import openpyxl
from openai import OpenAI
from PIL import Image

try:
    import pymupdf as fitz
except ImportError:  # PyMuPDF mai vechi
    import fitz
import docx
from pptx import Presentation

from sorter_utils import clean_name, normalize_path, safe_copy, win_long_path

# =====================================================================
# 1. CONFIGURARE IMPLICITA (poate fi suprascrisa din CLI)
# =====================================================================
DEFAULT_SOURCE = r"Z:\To be sorted V1"
DEFAULT_DEST = r"Z:\Sorted"
DEFAULT_API_URL = "http://127.0.0.1:1235/v1"

IGNORED_DIRS = {
    "node_modules", ".git", ".venv", "venv", "__pycache__",
    ".vscode", ".idea", "AppData", "Temp", "cache", "dist",
    "build", "$Recycle.Bin", ".cache",
}

JUNK_EXTENSIONS = {".tmp", ".bak", ".crdownload", ".part", ".dmp", ".log", ".ini"}
MIN_FILE_SIZE = 200  # fisiere mai mici merg direct in "to be deleted"

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tif", ".tiff"}
TEXT_EXTENSIONS = {".txt", ".md", ".csv", ".json", ".xml", ".py", ".js", ".html", ".svg", ".yaml", ".yml"}
MEDIA_EXTENSIONS = {".mp4", ".mkv", ".avi", ".mov", ".mp3", ".flac", ".wav"}

MAX_TEXT_CHARS = 1200
MAX_TOKENS = 90
MAX_TOKENS_RETRY = 240

# =====================================================================
# 2. PROMPT SISTEM — GRUPARE LARGA (FARA SUBFOLDERE PER FIRMA)
# =====================================================================
SYSTEM_PROMPT = """Ești un arhivar digital expert. Clasifică fișierul logic și concis în limba română.

REGULI STRICTE PENTRU SUBIECT (subfolder):
- SUBIECTUL REPREZINTĂ DOAR TIPUL GENERAL DE DOCUMENT (ex: "Facturi", "Contracte", "Extrase", "Comenzi", "Adeverințe", "Studii de Caz", "Cursuri", "Prezentări", "Drivere", "Anime", "Filme", "Muzică", "Imagini").
- ESTE STRICT INTERZIS să creezi subfoldere cu numele firmei, băncii sau persoanei!
  (Grupare corectă: și factura Raiffeisen și factura ING merg în subiectul "Facturi", NU în foldere separate cu numele băncilor).
- Subiectul are maxim 2-3 cuvinte.

REGULI PENTRU NUME (nume_nou):
- AICI introduci numele firmei/băncii/persoanei identificate!
- Format scurt (maxim 5-6 cuvinte, fără extensie, fără caractere interzise): [Tip Document] [Firmă Sau Persoană] [Detalii Scurte]
  * "Factură Raiffeisen Comandă 2026"
  * "Factură ING Abonament"
  * "Prezentare Proiect Marketing"
  * "Studiu de Caz Mihai"
- Dacă numele original are deja sens, setează pastreaza_original: true.

CATEGORII MARI:
- Financiar
- Acte și Juridic
- Muncă și Carieră
- Educație și Cursuri
- Media și Divertisment
- Fotografii și Grafică
- Programe și Utilitare
- Unknown

Răspunde DOAR cu JSON brut:
{
  "categorie": "String",
  "subiect": "String",
  "nume_nou": "String",
  "pastreaza_original": true | false
}"""


# =====================================================================
# 3. EXTRACTOARE DE CONTINUT
# =====================================================================
def extract_strings_from_binary(filepath: Path, max_chars: int = MAX_TEXT_CHARS) -> str:
    """Extragere rapidă din binare (.doc/.ppt vechi, .exe, .msi)."""
    try:
        with open(filepath, "rb") as f:
            raw = f.read(250000)
        ascii_strings = re.findall(b"[\x20-\x7E]{4,}", raw)
        utf16_strings = re.findall(b"(?:[\x20-\x7E]\x00){4,}", raw)
        decoded = [s.decode("ascii", "ignore") for s in ascii_strings]
        decoded += [s.decode("utf-16le", "ignore") for s in utf16_strings]
        cuvinte_utile = [
            w.strip() for w in decoded
            if len(w.strip()) > 3 and not all(c in "0123456789abcdefABCDEF" for c in w)
        ]
        return " ".join(cuvinte_utile[:120])[:max_chars]
    except Exception:
        return ""


def parse_torrent_file(filepath: Path) -> str:
    try:
        with open(filepath, "rb") as f:
            data = f.read(150000)
        info_idx = data.find(b"4:name")
        nume_torrent = ""
        if info_idx != -1:
            start = info_idx + 6
            colon = data.find(b":", start)
            length = int(data[start:colon])
            nume_torrent = data[colon + 1: colon + 1 + length].decode("utf-8", "ignore")
        strings_gasite = re.findall(rb"[\x20-\x7E]{5,}", data)
        mostra = " | ".join(s.decode("utf-8", "ignore") for s in strings_gasite[:8])
        return f"Torrent: {nume_torrent}\nConținut: {mostra}"[:MAX_TEXT_CHARS]
    except Exception:
        return f"Fișier torrent: {filepath.name}"


def parse_zip_file(filepath: Path) -> str:
    try:
        with zipfile.ZipFile(filepath, "r") as z:
            return f"Arhivă ZIP: {filepath.name}\nFișiere: " + ", ".join(z.namelist()[:15])
    except Exception:
        return f"Arhivă ZIP: {filepath.name}"


def read_text_file(filepath: Path, max_chars: int = MAX_TEXT_CHARS) -> str:
    for enc in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(filepath, "r", encoding=enc) as f:
                return f.read(max_chars).strip()
        except (UnicodeDecodeError, UnicodeError):
            continue
        except OSError:
            break
    return ""


def encode_pil_image(img: Image.Image) -> str:
    """768px + JPEG q80 = inferență GPU ~2x mai rapidă."""
    if img.mode != "RGB":
        img = img.convert("RGB")
    img.thumbnail((768, 768))
    buffer = BytesIO()
    img.save(buffer, format="JPEG", quality=80)
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


def extract_file_payload(filepath: Path):
    """Intoarce (mode, payload): mode este 'image', 'text' sau 'unknown'."""
    ext = filepath.suffix.lower()
    try:
        if ext in IMAGE_EXTENSIONS:
            with Image.open(filepath) as img:
                return "image", encode_pil_image(img)

        elif ext == ".pdf":
            doc = fitz.open(filepath)
            if len(doc) == 0:
                return "unknown", None
            text = "\n".join(doc[i].get_text("text") for i in range(min(2, len(doc)))).strip()
            if len(text) > 40:
                doc.close()
                return "text", text[:MAX_TEXT_CHARS]
            pix = doc[0].get_pixmap(dpi=120)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()
            return "image", encode_pil_image(img)

        elif ext == ".pptx":
            prs = Presentation(filepath)
            slides_text = []
            for slide in list(prs.slides)[:5]:
                for shape in slide.shapes:
                    if shape.has_text_frame and shape.text_frame.text.strip():
                        slides_text.append(shape.text_frame.text.strip())
                    elif shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                if cell.text.strip():
                                    slides_text.append(cell.text.strip())
            text = "\n".join(slides_text).strip()
            return ("text", text[:MAX_TEXT_CHARS]) if len(text) > 10 else ("text", f"Prezentare: {filepath.stem}")

        elif ext == ".ppt":
            text = extract_strings_from_binary(filepath)
            return "text", text if len(text) > 20 else f"Prezentare veche: {filepath.stem}"

        elif ext == ".docx":
            d = docx.Document(filepath)
            text = "\n".join(p.text for p in d.paragraphs if p.text.strip())
            return ("text", text[:MAX_TEXT_CHARS]) if len(text) > 10 else ("text", f"Document Word: {filepath.stem}")

        elif ext == ".doc":
            text = extract_strings_from_binary(filepath)
            return "text", text if len(text) > 20 else f"Document Word vechi: {filepath.stem}"

        elif ext == ".xlsx":
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            sheet = wb.active
            rows = [
                " | ".join(str(c) for c in row if c is not None)
                for row in list(sheet.iter_rows(values_only=True))[:15]
            ]
            wb.close()
            return "text", "\n".join(rows)[:MAX_TEXT_CHARS]

        elif ext == ".zip" or (ext in {".jar", ".apk", ".war", ".whl"} and zipfile.is_zipfile(filepath)):
            return "text", parse_zip_file(filepath)

        elif ext == ".torrent":
            return "text", parse_torrent_file(filepath)

        elif ext in {".exe", ".msi"}:
            metadata = extract_strings_from_binary(filepath)
            return "text", f"Executabil: {filepath.name}\nDate: {metadata}"

        elif ext in MEDIA_EXTENSIONS:
            return "text", f"Fișier Media: {filepath.name}"

        elif ext in TEXT_EXTENSIONS:
            return "text", read_text_file(filepath)

        text = extract_strings_from_binary(filepath, max_chars=800)
        return ("text", f"Fișier {ext}: {filepath.name}\n{text}") if text else ("unknown", None)

    except Exception as e:
        print(f"  [!] Eroare extragere pe {filepath.name}: {e}")
        return "unknown", None


# =====================================================================
# 4. APEL CĂTRE AI LOCAL + PARSARE JSON ROBUSTĂ
# =====================================================================
def parse_ai_json(raw: str):
    """Extrage primul obiect JSON valid din raspunsul modelului. Intoarce None la esec."""
    if not raw:
        return None
    text = re.sub(r"^```(?:json)?|```$", "", raw.strip(), flags=re.MULTILINE).strip()
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        return None
    try:
        data = json.loads(text[start:end + 1])
    except (json.JSONDecodeError, ValueError):
        return None
    if not isinstance(data, dict) or "categorie" not in data:
        return None
    return data


def build_messages(mode: str, payload: str, original_filename: str, strict: bool = False):
    if mode == "image":
        user_content = [
            {"type": "text", "text": f"Fișier: '{original_filename}'. Identifică emitentul/tipul și returnează JSON scurt."},
            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{payload}"}},
        ]
    else:
        user_content = f"Fișier: '{original_filename}'\n\n mostră:\n{payload}\n\nReturnează JSON scurt conform regulilor."
    if strict:
        user_content += "\nRĂSPUNDE DOAR CU OBIECTUL JSON, FĂRĂ TEXT ÎN JUR, FĂRĂ ```, NEAPĂRAT COMPLET (toate cele 4 chei)."
    return [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_content},
    ]


def ask_local_ai(client, model: str, mode: str, payload: str, original_filename: str) -> dict:
    """Un apel + o reîncercare cu max_tokens mai mare dacă JSON-ul e trunchiat/invalid."""
    for strict, max_tokens in ((False, MAX_TOKENS), (True, MAX_TOKENS_RETRY)):
        try:
            response = client.chat.completions.create(
                model=model,
                messages=build_messages(mode, payload, original_filename, strict=strict),
                temperature=0.0,
                max_tokens=max_tokens,
            )
            raw = response.choices[0].message.content or ""
            data = parse_ai_json(raw)
            if data is not None:
                data["ai_ok"] = True
                return data
            print(f"  [!] JSON invalid la {original_filename} (încercarea {2 if strict else 1}), reiau...")
        except Exception as e:
            print(f"  [EROARE AI la {original_filename}]: {e}")
            break
    return {
        "categorie": "Unknown",
        "subiect": "",
        "nume_nou": original_filename,
        "pastreaza_original": True,
        "ai_ok": False,
    }


# =====================================================================
# 5. LOGICA DE CLASIFICARE / NUMIRE
# =====================================================================
def junk_reason(file_path: Path):
    """Intoarce motivul pentru care fișierul merge direct în 'to be deleted', altfel None."""
    if any(part in IGNORED_DIRS for part in file_path.parts):
        return "director_ignorat"
    if file_path.suffix.lower() in JUNK_EXTENSIONS:
        return "extensie_junk"
    try:
        if file_path.stat().st_size < MIN_FILE_SIZE:
            return "prea_mic"
    except OSError:
        return "stat_esuat"
    return None


def build_final_name(res: dict, orig_stem: str, orig_ext: str) -> str:
    nume_nou = clean_name(res.get("nume_nou", ""), max_len=55)
    keep_orig = bool(res.get("pastreaza_original", False))
    curat_orig_stem = clean_name(orig_stem, max_len=50)
    if keep_orig or not nume_nou or nume_nou.lower() == curat_orig_stem.lower():
        return f"{curat_orig_stem}{orig_ext}"
    return f"{nume_nou} ({curat_orig_stem}){orig_ext}"


def resolve_target_folder(dest_path: Path, res: dict) -> Path:
    cat = clean_name(res.get("categorie", ""), max_len=25) or "Unknown"
    sub = clean_name(res.get("subiect", ""), max_len=30)
    return (dest_path / cat / sub) if sub else (dest_path / cat)


def load_processed(log_file: str) -> set:
    processed = set()
    if os.path.exists(log_file):
        with open(log_file, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            next(reader, None)  # header
            for row in reader:
                if row and row[0]:
                    processed.add(normalize_path(row[0].strip()))
    return processed


# =====================================================================
# 6. BUCLA PRINCIPALĂ
# =====================================================================
def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Sortare AI non-destructivă a fișierelor cu VLM local.")
    p.add_argument("--source", default=DEFAULT_SOURCE, help="Directorul sursă (default: %(default)s)")
    p.add_argument("--dest", default=DEFAULT_DEST, help="Directorul destinație (default: %(default)s)")
    p.add_argument("--log", default=None, help="Fișier CSV log (default: <dest>/sorting_log.csv)")
    p.add_argument("--api-url", default=DEFAULT_API_URL, help="Endpoint OpenAI-compatibil (default: %(default)s)")
    p.add_argument("--model", default=None, help="Numele modelului (default: primul model de pe server)")
    p.add_argument("--limit", type=int, default=0, help="Procesează doar primele N fișiere neprocesate (0 = toate)")
    p.add_argument("--dry-run", action="store_true", help="Doar afișează deciziile AI, nu copiază nimic și nu scrie în log")
    return p.parse_args(argv)


def open_log_safely(log_file: str):
    try:
        return open(log_file, "a", newline="", encoding="utf-8-sig")
    except PermissionError:
        print(f"[!] Nu pot deschide {log_file} — este posibil să fie deschis în Excel sau să ruleze alt proces Python.")
        print("    Închide fișierul și reîncearcă.")
        raise SystemExit(1)


def main(argv=None, client=None):
    args = parse_args(argv)

    source_path = Path(args.source)
    dest_path = Path(args.dest)
    log_file = args.log or str(dest_path / "sorting_log.csv")
    errors_file = str(dest_path / "sorting_errors.csv")
    delete_path = dest_path / "to be deleted"

    if not source_path.is_dir():
        print(f"[!] Eroare: folderul sursă {args.source} nu există!")
        return 1

    os.makedirs(win_long_path(dest_path), exist_ok=True)
    os.makedirs(win_long_path(delete_path), exist_ok=True)

    if client is None:
        client = OpenAI(base_url=args.api_url, api_key="lm-studio")
    model = args.model
    if model is None:
        try:
            modele = client.models.list()
            model = modele.data[0].id if modele.data else "local-model"
            print(f"-> Conectat la {args.api_url}. Model activ: {model}")
        except Exception as e:
            print(f"Atenție la conexiunea cu serverul ({e}); folosesc 'local-model'.")
            model = "local-model"

    processed_files = load_processed(log_file)
    print(f"Reluare proces: {len(processed_files)} fișiere au fost deja sortate.")

    all_files = [p for p in source_path.rglob("*") if p.is_file()]
    total = len(all_files)
    print(f"S-au găsit {total} fișiere în {args.source}.\n")

    log_f = None
    writer = None
    errors_writer = None
    errors_f = None
    if not args.dry_run:
        log_f = open_log_safely(log_file)
        writer = csv.writer(log_f)
        if not processed_files and (not os.path.exists(log_file) or os.path.getsize(log_file) == 0):
            writer.writerow(["Cale_Originala", "Mod", "Categorie", "Subiect", "Cale_Noua"])
        errors_f = open(errors_file, "a", newline="", encoding="utf-8-sig")
        errors_writer = csv.writer(errors_f)
        if os.path.getsize(errors_file) == 0:
            errors_writer.writerow(["Cale_Originala", "Eroare"])

    ignored_counter = 0
    ai_counter = 0
    done = 0

    try:
        for idx, file_path in enumerate(all_files, start=1):
            if args.limit and done >= args.limit:
                print(f"\nLimită atinsă (--limit {args.limit}).")
                break

            str_path = str(file_path)
            if normalize_path(str_path) in processed_files:
                continue

            orig_stem = file_path.stem
            orig_ext = file_path.suffix

            reason = junk_reason(file_path)
            if reason:
                try:
                    if args.dry_run:
                        print(f"[{idx}/{total}] DRY-RUN junk ({reason}) -> to be deleted\\{file_path.name}")
                    else:
                        final = safe_copy(file_path, delete_path, file_path.name)
                        writer.writerow([str_path, f"Junk_{reason}", "to be deleted", "", str(final)])
                        log_f.flush()
                    ignored_counter += 1
                    done += 1
                except Exception as e:
                    print(f"  [!] Eroare junk {file_path.name}: {e}")
                    if errors_writer:
                        errors_writer.writerow([str_path, f"junk: {e}"])
                        errors_f.flush()
                if ignored_counter % 250 == 0:
                    print(f"[{idx}/{total}] Copiat junk în 'to be deleted' (Total: {ignored_counter})...")
                continue

            print(f"[{idx}/{total}] Procesez AI: {file_path.name}")
            mode, payload = extract_file_payload(file_path)

            if mode == "unknown" or not payload:
                target = dest_path / "Unknown"
                try:
                    if args.dry_run:
                        print(f"  -> DRY-RUN Unknown: {file_path.name}")
                    else:
                        final = safe_copy(file_path, target, file_path.name)
                        writer.writerow([str_path, "Necunoscut", "Unknown", "", str(final)])
                        log_f.flush()
                        print(f"  -> Trimis în Unknown: {final.name}")
                    done += 1
                except Exception as e:
                    print(f"  [!] Eroare necunoscut {file_path.name}: {e}")
                    if errors_writer:
                        errors_writer.writerow([str_path, f"unknown: {e}"])
                        errors_f.flush()
                continue

            res = ask_local_ai(client, model, mode, payload, file_path.name)
            final_name = build_final_name(res, orig_stem, orig_ext)
            target_folder = resolve_target_folder(dest_path, res)

            try:
                if args.dry_run:
                    rel = target_folder.relative_to(dest_path)
                    print(f"  -> DRY-RUN: {rel}\\{final_name}")
                else:
                    final = safe_copy(file_path, target_folder, final_name)
                    rel = target_folder.relative_to(dest_path)
                    writer.writerow([str_path, mode, clean_name(res.get('categorie', ''), 25) or 'Unknown',
                                     clean_name(res.get('subiect', ''), 30), str(final)])
                    log_f.flush()
                    print(f"  -> Sortat în: {rel} -> {final.name}")
                ai_counter += 1
                done += 1
                if ai_counter % 25 == 0:
                    print(f"  ... {ai_counter} fișiere procesate cu AI până acum.")
            except Exception as e:
                print(f"  [!] Eroare copiere {file_path.name}: {e}")
                if errors_writer:
                    errors_writer.writerow([str_path, f"copiere: {e}"])
                    errors_f.flush()

    except KeyboardInterrupt:
        print("\n[!] Oprit de utilizator. Relansează scriptul pentru a continua de unde a rămas (resume automat).")
        return 130
    finally:
        if log_f:
            log_f.close()
        if errors_f:
            errors_f.close()

    print(f"\nGata! {ai_counter} fișiere procesate cu AI, {ignored_counter} junk, din {total} gasite.")
    if not args.dry_run:
        print(f"Log: {log_file}")
        print(f"Erori (de reîncercat la următoarea rulare): {errors_file}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
