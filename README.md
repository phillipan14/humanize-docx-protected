# humanize-docx-protected

Paragraph-by-paragraph DOCX humanizer via roundtrip translation
(English → Simplified Chinese → Korean → English) with automatic
protection of **proper nouns, citations, URLs, references, and
years** so they come back unchanged.

Designed for academic papers, theses, and reports where stylistic
paraphrasing is wanted but names, citations, and numbers must be
preserved exactly.

## What it does

Each translatable paragraph goes through:

```
English  →  Simplified Chinese  →  Korean  →  English
```

via Google Translate (`deep-translator`). This paraphrases wording
while preserving meaning.

**Before translation**, the script replaces protected strings with
placeholder tokens (`§1§`, `§2§`, …) that survive Google Translate.
**After translation**, placeholders are restored to the originals.

## What's protected

**Automatic (regex-based):**

- URLs and DOIs
- APA citations — `(Author, 2025)` or `(Author et al., 2025)`
- Interview citations — `(author interview, April 5, 2026)`
- Dollar amounts — `$100–200/hr`
- Years — `2024`, `2025`, etc.

**Named list (customize in the script):**

- Team member names, advisors, school names
- Framework terminology (e.g., `H-S-D`, `Universal Manipulation Interface`)
- Company names (Chinese and foreign)
- Investor / institution names
- Government agencies (IFR, NBS, SCIO, USTR, BIS, etc.)
- Geographic places
- Tech product / model names (GPT-4, Claude, DeepSeek, etc.)

See `PROTECTED_TERMS` in `humanize_protected.py` and extend it for
your own document.

## What's skipped entirely

- Empty / whitespace-only paragraphs
- Headings (Word style contains "Heading" or "Title")
- Anything inside sections listed in `SKIP_HEADINGS` — references,
  bibliography, appendices of interview summaries, Chinese abstract,
  resumes, list-of-figures, etc.
- Paragraphs shorter than 20 words
- Paragraphs that are primarily URLs
- Table / figure caption stubs
- Paragraphs containing `[placeholder]` markers

## Install

```bash
pip install deep-translator python-docx
```

## Run

```bash
python3 humanize_protected.py your_document.docx
```

Output: `your_document_humanized.docx` alongside the input.
Original is untouched.

Interrupted by a rate-limit? Resume:

```bash
python3 humanize_protected.py your_document.docx --resume
```

Progress is saved every 5 paragraphs so you don't lose work.

## Behavior and caveats

- Google Translate free tier will rate-limit after ~50–100 rapid
  requests. The script backs off and retries automatically; if
  blocked longer, run with `--resume` later.
- This is a style pass. **Always review output** — meaning can
  drift, numbers inside sentences (not protected by regex) can
  shift, and some phrasings become awkward.
- Extend `PROTECTED_TERMS` with any proper nouns unique to your
  document before running, or expect some to be paraphrased
  ("Skylarq" might become "Skylark", for instance, if not in the
  protected list).
- The script preserves the first-run formatting of each paragraph
  and clears subsequent runs; complex intra-paragraph formatting
  (colored text, mid-sentence bold, hyperlinks) will be flattened.

## License

MIT. Built for research use; no warranty.
