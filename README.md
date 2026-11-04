# AI File Sorting & Archiving Pipeline

Non-destructive sorting of a large personal dataset with a local vision-language model
(LM Studio / Bionic, OpenAI-compatible API). The script **copies** files into a
destination folder, classifies them with the AI, and renames them descriptively —
**originals are never modified or deleted**.

## Project structure

```
sort_files.py            # main pipeline (AI classification + copy + CSV log)
count_sorted_files.py    # stats & audit report for the sorted folder (CSV)
sorter_utils.py          # shared helpers (clean_name, safe_copy, Windows long paths)
requirements.txt         # dependencies
tests/                   # pytest suite (runs without an AI server, mocked client)
```

## Setup

```bat
pip install -r requirements.txt
```

Local AI server: LM Studio / Bionic on `http://127.0.0.1:1235/v1`, model
`qwen2.5-vl-7b-instruct`, context 4096–8192, Flash Attention ON, 100% GPU offload.

## Usage

```bat
:: full run with default paths (Z:\To be sorted V1 -> Z:\Sorted)
python sort_files.py

:: custom paths
python sort_files.py --source "D:\in" --dest "D:\out"

:: only the first 50 unprocessed files (useful for checking quality)
python sort_files.py --limit 50

:: dry-run: prints the AI decisions, copies nothing, writes no log
python sort_files.py --dry-run --limit 100

:: statistics for the sorted folder
python count_sorted_files.py --target "Z:\Sorted"
```

Options: `--source --dest --log --api-url --model --limit --dry-run` (see `--help`).

## Sorting rules

- **Non-destructive**: only `shutil.copy2` from source to destination.
- **Automatic junk**: system folders (`node_modules`, `.git`, `AppData`, `Temp`...), extensions
  `.tmp/.log/.ini/.bak...` and files < 200 bytes go straight to `Sorted\to be deleted`, without AI.
- **Naming**: `[Document Type] [Company/Person] [Short Details] ([original stem]).ext`, with spaces.
- **Anti-fragmentation**: the subfolder (`subiect`) is only the broad document type
  (`Financiar\Facturi`, not one folder per bank). The company/bank goes into the **file name**.
- **Crash-safe resume**: every copied file is logged immediately to `sorting_log.csv`
  (paths normalized, case-insensitive); on restart, logged files are skipped instantly.
  Files whose copy failed are **not** marked as processed — they go to `sorting_errors.csv`
  and are retried automatically on the next run.
- **Long Windows paths**: handled via the `\\?\` prefix (up to 32,767 chars) + short-name
  fallback; existing files are never overwritten (` (1)`, ` (2)`... suffixes are added).

## Content extraction before the AI call

PDF (text from the first 2 pages, or visual OCR at 120 DPI if scanned), `.docx`, `.pptx`,
`.xlsx` (first 15 rows), images (768px JPEG q80 thumbnail), `.doc/.ppt/.exe/.msi`
(ASCII/UTF-16 strings from the binary), `.torrent` (bencode `4:name`), `.zip/.jar/.apk`
(file listing), plain text (utf-8 → cp1252 fallback). Payload capped at 1200 chars,
`max_tokens=90`, `temperature=0.0` (+ one retry with 240 tokens if the JSON is truncated).

## Tests

```bat
pip install pytest
python -m pytest tests/ -q
```

Tests run locally, without an AI server (mocked OpenAI client) and without `Z:\` paths.

## Troubleshooting: "Import could not be resolved" in VS Code

Pylance shows `reportMissingImports` for `openpyxl` / `openai` / `PIL` / `pymupdf` / `docx` / `pptx`
when the selected interpreter is not the one where the dependencies were installed
(e.g. a stray `python.exe` picked from somewhere on disk). These are editor warnings only —
the script still runs fine from an interpreter that has the packages. To fix:

1. `Ctrl+Shift+P` → **Python: Select Interpreter** → choose your global Python 3.13
   (find its path in a terminal: `py -3.13 -c "import sys; print(sys.executable)"`).
2. Or install the deps into the currently selected interpreter instead:
   `python -m pip install -r requirements.txt`
3. Verify: `python -c "import openai, pymupdf, docx, pptx, openpyxl, PIL; print('ok')"`
4. If warnings persist after switching: `Ctrl+Shift+P` → **Developer: Reload Window**.

## Language convention

Code, comments, console output, CSV log/report headers and commit messages are in English.
Kept in Romanian on purpose (context-dependent): the AI system prompt and the JSON keys in
`sort_files.py` (the model must classify in Romanian, so category/subfolder names come out
in Romanian), the `to be deleted` folder name, and the existing CSV filenames
(`sorting_log.csv`, `sorting_errors.csv`, `statistici_fisiere.csv`) so already-sorted data
on the drive keeps matching on resume.

## Changelog (vs. the Gemini-chat version)

1. Failed copies are no longer logged as "processed" — they are retried on the next run
   (before, they were skipped forever, losing files).
2. Robust JSON parsing (extracts the object between the first `{` and last `}`) + one retry
   with `max_tokens=240` when the answer is truncated; before, they silently landed in `Unknown`.
3. `safe_copy` can never overwrite an existing file (the old fallback silently overwrote).
4. Empty/illegal category or subfolder from the AI → fallback to `Unknown`; `\\?\` prefix
   also used in `os.path.exists` checks.
5. CLI with `--source/--dest/--log/--model/--limit/--dry-run` (dry-run is useful for prompt
   tuning before the full run).
6. Clear error when `sorting_log.csv` is locked (open in Excel / zombie process).
7. More extensions handled (`.gif/.tif/.svg/.yaml`, zip-based `.jar/.apk`), text reading with
   encoding fallback, `pymupdf` imported with fallback to `fitz`.
8. Stats: relative paths robust to the `\\?\` prefix, new "Top 20 files by size" section,
   log/error CSVs excluded from counts.
