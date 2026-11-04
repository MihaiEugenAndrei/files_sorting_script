r"""AI file sorting pipeline (non-destructive copy + local VLM renaming).

Reads files from SOURCE_DIR, extracts content (text/OCR/image), asks a local VLM
(LM Studio / Bionic, OpenAI-compatible API) for a JSON classification, and copies
each file to DEST_DIR/<Category>/<Subject>/<new_name> (<original stem>).ext.
Originals are never modified or deleted.

Crash safety: every successfully copied file is logged immediately to CSV.
On restart, logged files are skipped instantly. Files whose copy FAILED are not
marked as processed (they go to sorting_errors.csv and are retried).

Usage:
    python sort_files.py                          # default config (Z:\...)
    python sort_files.py --source D:\in --dest D:\out
    python sort_files.py --dry-run --limit 50     # only show proposed moves, copy nothing

Language note: code, logs and CLI output are in English. The system prompt and the
JSON keys stay in Romanian on purpose — the model must classify in Romanian, so
category/subfolder names come out in Romanian. The "to be deleted" folder name and
the log CSV filenames are kept unchanged to stay compatible with already-sorted data.
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
except ImportError:  # older PyMuPDF
    import fitz
import docx
from pptx import Presentation

from sorter_utils import clean_name, normalize_path, safe_copy, win_long_path

# =====================================================================
# 1. DEFAULT CONFIG (can be overridden via CLI)
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
MIN_FILE_SIZE = 200  # smaller files go straight to "to be deleted"

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tif", ".tiff"}
TEXT_EXTENSIONS = {".txt", ".md", ".csv", ".json", ".xml", ".py", ".js", ".html", ".svg", ".yaml", ".yml"}
MEDIA_EXTENSIONS = {".mp4", ".mkv", ".avi", ".mov", ".mp3", ".flac", ".wav"}

MAX_TEXT_CHARS = 1200
MAX_TOKENS = 90
MAX_TOKENS_RETRY = 240

# =====================================================================
# 2. SYSTEM PROMPT — kept in Romanian on purpose: the model must classify
#    in Romanian (broad grouping, NO per-company subfolders).
#    JSON keys (categorie/subiect/nume_nou/pastreaza_original) are part of
#    this contract and must stay Romanian.
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
# 3. CONTENT EXTRACTORS
# =====================================================================
def extract_strings_from_binary(filepath: Path, max_chars: int = MAX_TEXT_CHARS) -> str:
    """Fast extraction from binaries (legacy .doc/.ppt, .exe, .msi)."""
    try:
        with open(filepath, "rb") as f:
            raw = f.read(250000)
        ascii_strings = re.findall(b"[\x20-\x7E]{4,}", raw)
        utf16_strings = re.findall(b"(?:[\x20-\x7E]\x00){4,}", raw)
        decoded = [s.decode("ascii", "ignore") for s in ascii_strings]
        decoded += [s.decode("utf-16le", "ignore") for s in utf16_strings]
        useful = [
            w.strip() for w in decoded
            if len(w.strip()) > 3 and not all(c in "0123456789abcdefABCDEF" for c in w)
        ]
        return " ".join(useful[:120])[:max_chars]
    except Exception:
        return ""


def parse_torrent_file(filepath: Path) -> str:
    try:
        with open(filepath, "rb") as f:
            data = f.read(150000)
        info_idx = data.find(b"4:name")
        torrent_name = ""
        if info_idx != -1:
            start = info_idx + 6
            colon = data.find(b":", start)
            length = int(data[start:colon])
            torrent_name = data[colon + 1: colon + 1 + length].decode("utf-8", "ignore")
        strings_found = re.findall(rb"[\x20-\x7E]{5,}", data)
        sample = " | ".join(s.decode("utf-8", "ignore") for s in strings_found[:8])
        return f"Torrent: {torrent_name}\nContents: {sample}"[:MAX_TEXT_CHARS]
    except Exception:
        return f"Torrent file: {filepath.name}"


def parse_zip_file(filepath: Path) -> str:
    try:
        with zipfile.ZipFile(filepath, "r") as z:
            return f"ZIP archive: {filepath.name}\nFiles: " + ", ".join(z.namelist()[:15])
    except Exception:
        return f"ZIP archive: {filepath.name}"


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
    """768px + JPEG q80 = ~2x faster GPU inference."""
    if img.mode != "RGB":
        img = img.convert("RGB")
    img.thumbnail((768, 768))
    buffer = BytesIO()
    img.save(buffer, format="JPEG", quality=80)
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


def extract_file_payload(filepath: Path):
    """Return (mode, payload): mode is 'image', 'text' or 'unknown'."""
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
            return ("text", text[:MAX_TEXT_CHARS]) if len(text) > 10 else ("text", f"Presentation: {filepath.stem}")

        elif ext == ".ppt":
            text = extract_strings_from_binary(filepath)
            return "text", text if len(text) > 20 else f"Legacy presentation: {filepath.stem}"

        elif ext == ".docx":
            d = docx.Document(filepath)
            text = "\n".join(p.text for p in d.paragraphs if p.text.strip())
            return ("text", text[:MAX_TEXT_CHARS]) if len(text) > 10 else ("text", f"Word document: {filepath.stem}")

        elif ext == ".doc":
            text = extract_strings_from_binary(filepath)
            return "text", text if len(text) > 20 else f"Legacy Word document: {filepath.stem}"

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
            return "text", f"Executable: {filepath.name}\nMetadata: {metadata}"

        elif ext in MEDIA_EXTENSIONS:
            return "text", f"Media file: {filepath.name}"

        elif ext in TEXT_EXTENSIONS:
            return "text", read_text_file(filepath)

        text = extract_strings_from_binary(filepath, max_chars=800)
        return ("text", f"File {ext}: {filepath.name}\n{text}") if text else ("unknown", None)

    except Exception as e:
        print(f"  [!] Extraction error on {filepath.name}: {e}")
        return "unknown", None


# =====================================================================
# 4. LOCAL AI CALL + ROBUST JSON PARSING
# =====================================================================
def parse_ai_json(raw: str):
    """Extract the first valid JSON object from the model output. None on failure."""
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
            {"type": "text", "text": f"File: '{original_filename}'. Identify the issuer/type and return short JSON."},
            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{payload}"}},
        ]
    else:
        user_content = f"File: '{original_filename}'\n\nSample:\n{payload}\n\nReturn short JSON per the rules."
    if strict:
        user_content += "\nRESPOND ONLY WITH THE JSON OBJECT, NO TEXT AROUND IT, NO ```, MUST BE COMPLETE (all 4 keys)."
    return [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_content},
    ]


def ask_local_ai(client, model: str, mode: str, payload: str, original_filename: str) -> dict:
    """One call + one retry with more tokens if the JSON came out truncated/invalid."""
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
            print(f"  [!] Invalid JSON for {original_filename} (attempt {2 if strict else 1}), retrying...")
        except Exception as e:
            print(f"  [AI ERROR on {original_filename}]: {e}")
            break
    return {
        "categorie": "Unknown",
        "subiect": "",
        "nume_nou": original_filename,
        "pastreaza_original": True,
        "ai_ok": False,
    }


# =====================================================================
# 5. CLASSIFICATION / NAMING LOGIC
# =====================================================================
def junk_reason(file_path: Path):
    """Return the reason a file goes straight to 'to be deleted', else None."""
    if any(part in IGNORED_DIRS for part in file_path.parts):
        return "ignored_dir"
    if file_path.suffix.lower() in JUNK_EXTENSIONS:
        return "junk_extension"
    try:
        if file_path.stat().st_size < MIN_FILE_SIZE:
            return "too_small"
    except OSError:
        return "stat_failed"
    return None


def build_final_name(res: dict, orig_stem: str, orig_ext: str) -> str:
    nume_nou = clean_name(res.get("nume_nou", ""), max_len=55)
    keep_orig = bool(res.get("pastreaza_original", False))
    clean_orig_stem = clean_name(orig_stem, max_len=50)
    if keep_orig or not nume_nou or nume_nou.lower() == clean_orig_stem.lower():
        return f"{clean_orig_stem}{orig_ext}"
    return f"{nume_nou} ({clean_orig_stem}){orig_ext}"


def resolve_target_folder(dest_path: Path, res: dict) -> Path:
    cat = clean_name(res.get("categorie", ""), max_len=25) or "Unknown"
    sub = clean_name(res.get("subiect", ""), max_len=30)
    return (dest_path / cat / sub) if sub else (dest_path / cat)


def load_processed(log_file: str) -> set:
    processed = set()
    if os.path.exists(log_file):
        with open(log_file, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            next(reader, None)  # header row is skipped, so old Romanian-header logs still resume fine
            for row in reader:
                if row and row[0]:
                    processed.add(normalize_path(row[0].strip()))
    return processed


# =====================================================================
# 6. MAIN LOOP
# =====================================================================
def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Non-destructive AI file sorting with a local VLM.")
    p.add_argument("--source", default=DEFAULT_SOURCE, help="Source directory (default: %(default)s)")
    p.add_argument("--dest", default=DEFAULT_DEST, help="Destination directory (default: %(default)s)")
    p.add_argument("--log", default=None, help="CSV log file (default: <dest>/sorting_log.csv)")
    p.add_argument("--api-url", default=DEFAULT_API_URL, help="OpenAI-compatible endpoint (default: %(default)s)")
    p.add_argument("--model", default=None, help="Model name (default: first model on the server)")
    p.add_argument("--limit", type=int, default=0, help="Process only the first N unprocessed files (0 = all)")
    p.add_argument("--pause-every", type=int, default=0, help="Pause for Enter after every N AI-processed files (0 = never); Ctrl+C still stops safely and resumes next run")
    p.add_argument("--dry-run", action="store_true", help="Only print AI decisions; copy nothing and write no log")
    return p.parse_args(argv)


def open_log_safely(log_file: str):
    try:
        return open(log_file, "a", newline="", encoding="utf-8-sig")
    except PermissionError:
        print(f"[!] Cannot open {log_file} — it may be open in Excel or another Python process is running.")
        print("    Close the file and try again.")
        raise SystemExit(1)


def main(argv=None, client=None):
    args = parse_args(argv)

    source_path = Path(args.source)
    dest_path = Path(args.dest)
    log_file = args.log or str(dest_path / "sorting_log.csv")
    errors_file = str(dest_path / "sorting_errors.csv")
    # Folder name kept in English on purpose: existing Sorted output already uses it.
    delete_path = dest_path / "to be deleted"

    if not source_path.is_dir():
        print(f"[!] Error: source folder {args.source} does not exist!")
        return 1

    os.makedirs(win_long_path(dest_path), exist_ok=True)
    os.makedirs(win_long_path(delete_path), exist_ok=True)

    if client is None:
        client = OpenAI(base_url=args.api_url, api_key="lm-studio")
    model = args.model
    if model is None:
        try:
            models = client.models.list()
            model = models.data[0].id if models.data else "local-model"
            print(f"-> Connected to {args.api_url}. Active model: {model}")
        except Exception as e:
            print(f"Warning: could not reach the server ({e}); using 'local-model'.")
            model = "local-model"

    processed_files = load_processed(log_file)
    print(f"Resume: {len(processed_files)} files already sorted.")

    all_files = [p for p in source_path.rglob("*") if p.is_file()]
    total = len(all_files)
    print(f"Found {total} files in {args.source}.\n")

    log_f = None
    writer = None
    errors_writer = None
    errors_f = None
    if not args.dry_run:
        log_f = open_log_safely(log_file)
        writer = csv.writer(log_f)
        if not processed_files and (not os.path.exists(log_file) or os.path.getsize(log_file) == 0):
            writer.writerow(["Source_Path", "Mode", "Category", "Subject", "New_Path"])
        errors_f = open(errors_file, "a", newline="", encoding="utf-8-sig")
        errors_writer = csv.writer(errors_f)
        if os.path.getsize(errors_file) == 0:
            errors_writer.writerow(["Source_Path", "Error"])

    ignored_counter = 0
    ai_counter = 0
    done = 0

    try:
        for idx, file_path in enumerate(all_files, start=1):
            if args.limit and done >= args.limit:
                print(f"\nLimit reached (--limit {args.limit}).")
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
                    print(f"  [!] Junk error on {file_path.name}: {e}")
                    if errors_writer:
                        errors_writer.writerow([str_path, f"junk: {e}"])
                        errors_f.flush()
                if ignored_counter % 250 == 0:
                    print(f"[{idx}/{total}] Copied junk to 'to be deleted' (total: {ignored_counter})...")
                continue

            print(f"[{idx}/{total}] AI processing: {file_path.name}")
            mode, payload = extract_file_payload(file_path)

            if mode == "unknown" or not payload:
                target = dest_path / "Unknown"
                try:
                    if args.dry_run:
                        print(f"  -> DRY-RUN Unknown: {file_path.name}")
                    else:
                        final = safe_copy(file_path, target, file_path.name)
                        writer.writerow([str_path, "Unknown", "Unknown", "", str(final)])
                        log_f.flush()
                        print(f"  -> Sent to Unknown: {final.name}")
                    done += 1
                except Exception as e:
                    print(f"  [!] Unknown-file error on {file_path.name}: {e}")
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
                    print(f"  -> Sorted into: {rel} -> {final.name}")
                ai_counter += 1
                done += 1
                if ai_counter % 25 == 0:
                    print(f"  ... {ai_counter} AI-processed files so far.")
                if args.pause_every and ai_counter % args.pause_every == 0:
                    input(f"\n[PAUSE] {ai_counter} AI-processed files so far. Press Enter to continue, Ctrl+C to stop (auto-resume on next run)...")
            except Exception as e:
                print(f"  [!] Copy error on {file_path.name}: {e}")
                if errors_writer:
                    errors_writer.writerow([str_path, f"copy: {e}"])
                    errors_f.flush()

    except KeyboardInterrupt:
        print("\n[!] Interrupted by user. Re-run the script to continue from where it left off (automatic resume).")
        return 130
    finally:
        if log_f:
            log_f.close()
        if errors_f:
            errors_f.close()

    print(f"\nDone! {ai_counter} AI-processed, {ignored_counter} junk, out of {total} found.")
    if not args.dry_run:
        print(f"Log: {log_file}")
        print(f"Errors (retried on next run): {errors_file}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
