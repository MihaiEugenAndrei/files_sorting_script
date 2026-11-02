"""Helperi partajate intre sort_files.py si count_sorted_files.py."""

import os
import re
import shutil
from pathlib import Path

INVALID_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
LONG_PATH_PREFIX = "\\\\?\\"


def clean_name(text, max_len: int = 65) -> str:
    """Curata un string pentru a fi folosit ca nume de fisier/folder pe Windows."""
    if not text:
        return ""
    text = INVALID_CHARS_RE.sub("", str(text))
    text = text.replace("_", " ")
    text = re.sub(r"\s+", " ", text.strip())
    text = text.rstrip(". ,")
    return text[:max_len].strip()


def normalize_path(p) -> str:
    """Normalizare pentru comparatii exacte de cai (case-insensitive pe Windows)."""
    return os.path.normcase(os.path.normpath(str(p)))


def win_long_path(p) -> str:
    """Prepend \\?\\ pe Windows pentru cai > 260 caractere."""
    abs_str = os.path.abspath(str(p))
    if os.name == "nt" and not abs_str.startswith(LONG_PATH_PREFIX):
        return LONG_PATH_PREFIX + abs_str
    return abs_str


def strip_long_path_prefix(p: str) -> str:
    if p.startswith(LONG_PATH_PREFIX):
        return p[len(LONG_PATH_PREFIX):]
    return p


def unique_destination(target_folder: Path, stem: str, ext: str) -> Path:
    """Gaseste un nume de destinatie nefolosit; nu suprascrie niciodata."""
    candidate = target_folder / f"{stem}{ext}"
    counter = 1
    while os.path.exists(win_long_path(candidate)):
        candidate = target_folder / f"{stem} ({counter}){ext}"
        counter += 1
        if counter > 9999:
            raise RuntimeError(f"Prea multe coliziuni de nume in {target_folder}")
    return candidate


def safe_copy(src: Path, target_folder: Path, desired_filename: str, max_stem: int = 85) -> Path:
    """Copie non-destructiva cu suport cai lungi si fallback pe nume scurt.

    Nu suprascrie fisiere existente; intoarce calea efectiva de destinatie.
    """
    target_long = win_long_path(target_folder)
    os.makedirs(target_long, exist_ok=True)

    stem = clean_name(Path(desired_filename).stem, max_len=max_stem) or "fara_nume"
    ext = Path(desired_filename).suffix.lower()[:10]

    dest_path = unique_destination(target_folder, stem, ext)
    try:
        shutil.copy2(win_long_path(src), win_long_path(dest_path))
        return dest_path
    except Exception:
        partial = win_long_path(dest_path)
        if os.path.exists(partial):
            try:
                os.remove(partial)
            except OSError:
                pass
        fallback = unique_destination(target_folder, stem[:35], ext)
        shutil.copy2(win_long_path(src), win_long_path(fallback))
        return fallback
