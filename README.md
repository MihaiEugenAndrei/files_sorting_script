# AI File Sorting & Archiving Pipeline

Sortare non-destructivă a unui dataset personal mare cu un VLM local (LM Studio / Bionic,
API compatibil OpenAI). Scriptul **copiază** fișierele într-un folder destinație, le
clasifică cu AI și le redenumește descriptiv — **originalele nu sunt niciodată modificate
sau șterse**.

## Structură

```
sort_files.py            # pipeline-ul principal de sortare (AI + copiere + log CSV)
count_sorted_files.py    # statistici & audit pentru folderul sortat (CSV raport)
sorter_utils.py          # helperi partajați (clean_name, safe_copy, căi lungi Windows)
requirements.txt         # dependențe
tests/                   # suite pytest (rulează fără server AI, cu client mock)
```

## Setup

```bat
pip install -r requirements.txt
```

Server local: LM Studio / Bionic pe `http://127.0.0.1:1235/v1`, model `qwen2.5-vl-7b-instruct`,
context 4096–8192, Flash Attention ON, 100% GPU offload.

## Utilizare

```bat
:: rulare completă cu căile implicite (Z:\To be sorted V1 -> Z:\Sorted)
python sort_files.py

:: căi custom
python sort_files.py --source "D:\in" --dest "D:\out"

:: doar primele 50 de fișiere neprocesate (util pentru testarea calității)
python sort_files.py --limit 50

:: dry-run: afișează deciziile AI fără să copieze și fără să scrie în log
python sort_files.py --dry-run --limit 100

:: statistici pentru folderul sortat
python count_sorted_files.py --target "Z:\Sorted"
```

Opțiuni: `--source --dest --log --api-url --model --limit --dry-run` (vezi `--help`).

## Reguli de sortare

- **Non-destructiv**: doar `shutil.copy2` din sursă în destinație.
- **Junk automat**: foldere de sistem (`node_modules`, `.git`, `AppData`, `Temp`...), extensii
  `.tmp/.log/.ini/.bak...` și fișiere < 200 bytes merg direct în `Sorted\to be deleted`, fără AI.
- **Nume**: `[Tip Document] [Firmă/Persoană] [Detalii] ([stem original]).ext`, cu spații.
- **Anti-fragmentare**: subfolderul (`subiect`) e doar tipul general de document
  (`Financiar\Facturi`, nu câte un folder per bancă). Firma/banca intră în **numele fișierului**.
- **Resume la crash**: fiecare fișier copiat e logat imediat în `sorting_log.csv`
  (căi normalizate case-insensitive); la relansare, cele logate sunt sărite instant.
  Fișierele la care copierea a eșuat **nu** sunt marcate ca procesate — sunt în
  `sorting_errors.csv` și se reîncearcă automat la următoarea rulare.
- **Căi lungi Windows**: suport `\\?\` (până la 32.767 caractere) + fallback pe nume scurt;
  nu se suprascrie niciodată un fișier existent (se adaugă ` (1)`, ` (2)`...).

## Extracție de conținut înainte de AI

PDF (text din primele 2 pagini, sau OCR vizual la 120 DPI dacă e scan), `.docx`, `.pptx`,
`.xlsx` (primele 15 rânduri), imagini (thumbnail 768px JPEG q80), `.doc/.ppt/.exe/.msi`
(stringuri ASCII/UTF-16 din binar), `.torrent` (bencode `4:name`), `.zip/.jar/.apk`
(lista de fișiere), text (utf-8 → cp1252 fallback). Payload limitat la 1200 caractere,
`max_tokens=90`, `temperature=0.0` (+ o reîncercare cu 240 tokeni dacă JSON-ul e trunchiat).

## Teste

```bat
pip install pytest
python -m pytest tests/ -q
```

Testele rulează local, fără server AI (client OpenAI mock) și fără căi `Z:\`.

## Istoric (față de varianta din chat-ul Gemini)

1. Copierile eșuate nu mai sunt logate ca "procesate" — se reîncearcă la relansare
   (înainte erau sărite definitiv, pierzând fișiere).
2. Parsare JSON robustă (extrage obiectul dintre prima `{` și ultima `}`) + reîncercare
   cu `max_tokens=240` când răspunsul e trunchiat; înainte ajungeau silențios în `Unknown`.
3. `safe_copy` nu mai poate suprascrie un fișier existent (fallback-ul vechi suprascria).
4. Categorii/subiecte goale → fallback `Unknown`; verificări `\\?\` și la `os.path.exists`.
5. CLI cu `--source/--dest/--log/--model/--limit/--dry-run` (dry-run util pentru tuning de prompt).
6. Mesaj clar când `sorting_log.csv` e blocat (deschis în Excel / proces zombie).
7. Extensii suplimentare (`.gif/.tif/.svg/.yaml`, arhive zip-based `.jar/.apk`), citire text
   cu fallback de encoding, `pymupdf` importat cu fallback la `fitz`.
8. Statistici: cale relativă robustă la prefixul `\\?\`, secțiune nouă "Top 20 fișiere ca
   dimensiune", log-urile și erorile excluse din numărătoare.
