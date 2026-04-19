"""
Humanize a DOCX paragraph-by-paragraph via roundtrip translation, while
protecting proper nouns, citations, references, URLs, and named entities.

HOW IT WORKS
------------
Each translatable paragraph is processed through:

    English  →  Simplified Chinese  →  Korean  →  English

via Google Translate (free, unofficial, via the `deep-translator` library).
This produces a paraphrased version that preserves meaning but varies
wording — useful for style homogenization / AI-text-detection resistance.

PROTECTIONS
-----------
Before translation, the script replaces the following with placeholder tokens
(§1§, §2§, ...) that survive Google Translate unchanged. After translation,
placeholders are restored to the original strings.

  - URLs and DOIs
  - APA citations: (Author, 2025) or (Author et al., 2025)
  - Interview citations: (author interview, April 5, 2026)
  - A maintained list of proper nouns (PROTECTED_TERMS) — team names,
    advisor, school, framework name, company names, institutions,
    investors, places, etc. Extend for your document.
  - Dollar amounts and years

PARAGRAPHS SKIPPED ENTIRELY
---------------------------
  - Empty / whitespace-only
  - Headings (Word style contains "Heading" or "Title")
  - Anything inside sections named in SKIP_HEADINGS
    (References, Bibliography, Appendix D interview compendium,
     Chinese Abstract, List of Figures, RESUME, 声明, etc.)
  - Very short paragraphs (< 20 words)
  - Paragraphs dominated by URLs
  - Paragraphs containing "[placeholder]" / "[To be" markers
  - Table / figure caption stubs

PROGRESS
--------
Saves every 5 paragraphs so work isn't lost on rate-limit errors.
Use --resume to continue after interruption.

USAGE
-----
    pip install deep-translator python-docx
    python3 humanize_protected.py "your_document.docx"
    python3 humanize_protected.py "your_document.docx" --resume

OUTPUT
------
Writes `your_document_humanized.docx` alongside the input. Original untouched.

NOTES
-----
  - Google Translate unofficial may rate-limit after ~50-100 calls.
    The script will back off and retry; if blocked, run with --resume later.
  - Expand PROTECTED_TERMS with any additional proper nouns unique to your
    document to avoid mangling.
  - This is a style-paraphrase pass, not an edit. Always review output.
"""
import sys, time, random, re
from pathlib import Path
from deep_translator import GoogleTranslator
from docx import Document

MAX_CHARS = 4500
SAVE_EVERY = 5
DELAY_MIN = 1.8
DELAY_MAX = 3.2

CHAIN = [("en", "zh-CN"), ("zh-CN", "ko"), ("ko", "en")]

# ------------------------------------------------------------------------
# Proper nouns / names / institutions / framework terms to PROTECT from
# translation. Extend this list for your document.
# ------------------------------------------------------------------------
PROTECTED_TERMS = [
    # EXAMPLE TERMS — REPLACE / EXTEND WITH YOUR OWN PROPER NOUNS
    # Team & advisor
    "Phillip Guangning An", "Phillip An", "Phillip",
    "Angelo Mok", "Angelo",
    "Boran Cui", "Boran",
    "Professor David Q. Pan", "Professor Pan", "David Pan", "David Q. Pan",
    "潘庆中", "Pan Qingzhong",
    "Jack",
    # Schools & programs
    "Schwarzman College", "Schwarzman Scholar", "Schwarzman Scholars Program", "Schwarzman Scholars", "Schwarzman",
    "Tsinghua University", "Tsinghua",
    "Master of Global Affairs", "MGA",
    # Framework & key terminology
    "Hardware–Software–Data/Service", "Hardware-Software-Data/Service", "H-S-D framework", "H-S-D",
    # Partner org & group capstone terms
    "Skylarq AI", "Skylarq",
    "Capstone Project", "Capstone", "Individual Reflection",
    # Company names (Chinese)
    "Geekplus", "Hai Robotics", "Quicktron", "Libiao", "Mushiny",
    "Pudu Robotics", "Pudu", "KEENON", "Keenon", "Gaussian Robotics", "Gaussian", "Reeman", "Orion Star",
    "AUBO Robotics", "AUBO", "JAKA Robotics", "JAKA", "Elite Robots", "Elite", "Dobot", "Flexiv",
    "Estun Automation", "Estun", "SIASUN", "Efort", "Inovance", "GSK CNC", "ROKAE",
    "MicroPort MedBot", "MicroPort", "Tinavi", "Edge Medical", "Wego",
    "Fourier Intelligence", "Fourier",
    "Unitree Robotics", "Unitree", "Galaxy Robotics", "Galaxy", "AGI Bot", "Agibot",
    "LimX Dynamics", "LimX", "Kepler", "Galbot", "Deep Robotics", "CloudMinds",
    "IO-AI Tech", "IO-AI",
    "BitRobot Network", "BitRobot", "FrodoBox Labs", "FrodoBox",
    "Dexta Robotics", "Dexta", "Manus", "Rococo", "Mantis",
    # Foreign benchmarks
    "Universal Robots", "FANUC", "ABB", "KUKA",
    "Looking Glass XR", "Physical Intelligence", "Skild",
    "Agility Robotics", "Agility", "Figure AI", "Boston Dynamics", "SenseGlove",
    # Investors / VCs
    "Hill House Capital", "Hillhouse Capital", "Hillhouse",
    "Andreessen Horowitz", "Andreessen", "Porsche", "Lenovo", "Volkswagen", "SK Hynix", "IDG",
    "Mercor", "Simovian Intelligence", "Simovian", "Boundless VC", "Boundless",
    # Government / Institutional
    "International Federation of Robotics", "IFR",
    "National Bureau of Statistics of China", "NBS", "SCIO", "State Council Information Office",
    "USTR", "Bureau of Industry and Security", "BIS", "Federal Reserve Board",
    "U.S. Bureau of Labor Statistics", "BLS", "Eurostat",
    "JETRO", "KOTRA", "ITIF", "CSIS", "CSET",
    "HKEX", "Hong Kong Exchanges and Clearing",
    "FDA", "CE MDR", "CE", "ISO/TS 15066",
    "European Commission", "European Union", "Office of Foreign Assets Control", "OFAC",
    "Tsinghua University Graduate Thesis Writing Guide",
    # Tech / model names
    "GPT-4", "Claude", "DeepSeek", "ChatGPT", "Gemini",
    "VLA", "UMI", "Universal Manipulation Interface",
    "Teleoperation", "AMR", "SLAM",
    "Ceres", "Neuromecca", "Dexmed", "ByteDance", "Google DeepMind", "DeepMind", "Google", "OpenAI", "Anthropic",
    "Bosch", "Imperial College London", "Momenta",
    "Tencent", "Alibaba", "BMW", "Mercedes", "Xiaomi", "GE", "Nvidia", "NVIDIA",
    # Past capstone precedents
    "Xu Jia", "Rui Kaili", "Zou Yujia", "Wang Ziqi",
    # Document structure terms
    "Thuthesis", "Tsinghua Guide to Thesis Writing",
    "Executive Summary", "Table of Contents", "Abstract", "Acknowledgements",
    "Appendix A", "Appendix B", "Appendix C", "Appendix D", "Appendix E",
    "Chapter 1", "Chapter 2", "Chapter 3", "Chapter 4", "Chapter 5", "Chapter 6", "Chapter 7", "Chapter 8",
    "Section 4.3", "Section 4.4", "Section 5.14", "Section 6.3.4", "Section 6.5.1",
    # Places
    "Beijing", "Shanghai", "Shenzhen", "Suzhou", "Hangzhou", "Guangzhou", "Huzhou",
    "China", "Chinese", "United States", "U.S.", "US",
    "Japan", "Korea", "South Korea", "Singapore",
    "Germany", "United Kingdom", "UK", "Switzerland", "Poland",
    "India", "Southeast Asia", "SEA", "Vietnam", "Thailand", "Philippines",
    "Russia", "Europe", "European", "Macau",
    # Chinese characters that should not be translated
    "摘要", "声明",
    # Agency abbreviations (duplicates above but kept to handle standalone)
    "FTC", "EU",
]

# Sort by length descending so longer matches replace first
PROTECTED_TERMS = sorted(set(PROTECTED_TERMS), key=lambda x: -len(x))

# Citation / URL / structural patterns protected automatically
CITATION_PATTERN = re.compile(r'\([A-Z][A-Za-z0-9 .,&\'\-]+?,\s*(?:19|20)\d{2}[a-z]?\)')
INTERVIEW_PATTERN = re.compile(r'\([A-Za-z ,]+interview[^)]*?202\d\)', re.IGNORECASE)
URL_PATTERN = re.compile(r'https?://\S+')
NUMERIC_BRACKET = re.compile(r'\$[\d,]+(?:–[\d,]+)?(?:/hr)?')
YEAR_PATTERN = re.compile(r'\b20\d{2}\b|\b19\d{2}\b')

# Sections skipped entirely
SKIP_HEADINGS = [
    "References", "Bibliography", "Appendix A", "Appendix B", "Appendix C", "Appendix D", "Appendix E",
    "LIST OF FIGURES", "LIST OF SYMBOLS", "TABLE OF CONTENTS", "Generative AI Use Disclosure",
    "声明", "RESUME", "Interview D.", "COMMENTS FROM", "RESOLUTION OF",
    "摘要", "Chinese Abstract",
    "LIST OF FIGURES AND TABLES", "LIST OF SYMBOLS AND ACRONYMS",
    "About Schwarzman College Group Capstone",
    "Responsibility Description",
]

def is_heading(para):
    if para.style and para.style.name:
        return 'Heading' in para.style.name or 'Title' in para.style.name
    return False

def is_skip_paragraph(para, current_section):
    text = para.text.strip()
    if not text:
        return True
    if is_heading(para):
        return True
    if current_section and any(h.lower() in current_section.lower() for h in SKIP_HEADINGS):
        return True
    if len(text.split()) < 20:
        return True
    if URL_PATTERN.search(text) and len(URL_PATTERN.findall(text)[0]) > len(text) * 0.3:
        return True
    if '[To be' in text or '[resume placeholder' in text or '[Title of' in text:
        return True
    if (text.startswith('Table ') or text.startswith('Figure ')) and len(text.split()) < 30:
        return True
    if text.isupper() and len(text.split()) < 10:
        return True
    return False

def protect(text):
    mapping = {}
    counter = [0]
    def tok():
        counter[0] += 1
        return f"§{counter[0]}§"

    def sub(pat, s):
        def rep(m):
            t = tok()
            mapping[t] = m.group(0)
            return t
        return pat.sub(rep, s)

    text = sub(URL_PATTERN, text)
    text = sub(CITATION_PATTERN, text)
    text = sub(INTERVIEW_PATTERN, text)

    for term in PROTECTED_TERMS:
        if term in text:
            t = tok()
            mapping[t] = term
            text = text.replace(term, t)

    text = sub(NUMERIC_BRACKET, text)
    text = sub(YEAR_PATTERN, text)
    return text, mapping

def restore(text, mapping):
    for _ in range(3):
        for token, original in mapping.items():
            if token in text:
                text = text.replace(token, original)
    for token, original in mapping.items():
        for bad in [token.replace('§', '#'), token.replace('§', '$'), token.replace('§', '')]:
            if bad and bad != token and bad in text:
                text = text.replace(bad, original)
    return text

def translate_chunk(text, src, tgt, retries=3):
    if not text.strip():
        return text
    for attempt in range(retries):
        try:
            return GoogleTranslator(source=src, target=tgt).translate(text)
        except Exception as e:
            if attempt < retries - 1:
                wait = (attempt + 1) * 10 + random.uniform(0, 5)
                print(f"    rate-limit / error — waiting {wait:.0f}s...", file=sys.stderr)
                time.sleep(wait)
            else:
                raise

def translate_text(text, src, tgt):
    if len(text) <= MAX_CHARS:
        return translate_chunk(text, src, tgt)
    sents = text.replace('. ', '.\n').split('\n')
    chunks, cur = [], ""
    for s in sents:
        if len(cur) + len(s) + 1 > MAX_CHARS:
            if cur: chunks.append(cur)
            cur = s
        else:
            cur = (cur + ' ' + s).strip() if cur else s
    if cur: chunks.append(cur)
    out = []
    for c in chunks:
        out.append(translate_chunk(c, src, tgt))
        time.sleep(random.uniform(0.5, 1.2))
    return ' '.join(out)

def roundtrip(text):
    cur = text
    for src, tgt in CHAIN:
        cur = translate_text(cur, src, tgt)
        time.sleep(random.uniform(0.8, 1.5))
    return cur

def main(docx_path, resume=False):
    input_path = Path(docx_path)
    output_path = input_path.with_stem(input_path.stem + "_humanized")

    if resume and output_path.exists():
        doc = Document(str(output_path))
        original_doc = Document(str(input_path))
        done = set()
        for i, (o, n) in enumerate(zip(original_doc.paragraphs, doc.paragraphs)):
            if o.text.strip() and o.text != n.text:
                done.add(i)
        print(f"Resuming: {len(done)} paragraphs already humanized.\n")
    else:
        doc = Document(str(input_path))
        done = set()

    current_section = ""
    all_paras = list(enumerate(doc.paragraphs))
    total = len(all_paras)
    processed = 0

    for i, para in all_paras:
        if is_heading(para) and para.text.strip():
            current_section = para.text.strip()
            continue
        if i in done:
            continue
        if is_skip_paragraph(para, current_section):
            continue

        text = para.text.strip()
        label = text[:80].replace('\n', ' ')
        print(f"[§ {current_section[:40]} | para {i}/{total}] {label}...")

        try:
            protected, mapping = protect(text)
            translated = roundtrip(protected)
            restored = restore(translated, mapping)
        except Exception as e:
            print(f"  ERROR: {e} — keeping original")
            doc.save(str(output_path))
            continue

        if para.runs:
            para.runs[0].text = restored
            for r in para.runs[1:]:
                r.text = ""
        else:
            para.text = restored

        processed += 1
        print(f"  done ({len(text)} → {len(restored)} chars, {len(mapping)} protections)")

        if processed % SAVE_EVERY == 0:
            doc.save(str(output_path))
            print(f"  [progress: {processed} paragraphs humanized, saved]")

        time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

    doc.save(str(output_path))
    print(f"\nCompleted. Output: {output_path}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python3 humanize_protected.py <document.docx> [--resume]")
        sys.exit(1)
    resume = "--resume" in sys.argv
    main(sys.argv[1], resume=resume)
