# Declaration: Code generated using Anthropic Claude (Opus 5.5)
import hashlib
import json
import re
import shutil
import statistics
import subprocess
import sys
import tempfile
from functools import lru_cache
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import ollama
import pymupdf

TARGET_WORDS = 1500
MAX_CHUNK_FACTOR = 1.2
FOOTNOTE_SIZE_RATIO = 0.9
MARGIN_FRACTION = 0.08
SUPERSCRIPT_SIZE_RATIO = 0.85
INDENT_POINTS = 10
SHORT_LINE_POINTS = 30
PARAGRAPH_GAP_FACTOR = 1.3
HEADING_SIZE_FACTOR = 1.05
TEMPERATURE = 0.4
MAX_REPLY_TOKENS = 3000
CONTEXT_OPTIONS = [8192, 12288, 16384, 24576, 32768]
TOKENS_PER_WORD = 1.6
REPLY_ALLOWANCE_TOKENS = 2000
THINK_REPLY_ALLOWANCE_TOKENS = 6000
CONTEXT_FULL_FRACTION = 0.98
TIMINGS_FILE = "timings.json"
WORD_LIST_PATH = "/usr/share/dict/words"
LIGATURE_GUESSES = ["ff", "fi", "fl", "ffi", "ffl"]
RESUME = "Resume"
OVERWRITE = "Overwrite"
KEEP_BOTH = "Keep both"
WORD_SUFFIXES_ACCEPTED = (".doc", ".docx")
WORD_APP = Path("/Applications/Microsoft Word.app")
WORD_SANDBOX_TMP = Path.home() / "Library/Containers/com.microsoft.Word/Data/tmp"
LIBREOFFICE_PATHS = ["/Applications/LibreOffice.app/Contents/MacOS/soffice", "soffice", "libreoffice"]
CONVERSION_TIMEOUT_SECONDS = 300

WORD_SCRIPT = """
on run argv
    set inputPath to item 1 of argv
    set outputPath to item 2 of argv
    set wasRunning to application "Microsoft Word" is running
    tell application "Microsoft Word"
        open file name inputPath
        set theDocument to document 1
        save as theDocument file name outputPath file format format PDF
        close theDocument saving no
        if not wasRunning then quit
    end tell
end run
"""

DEFAULT_PROMPT = 'Check spelling, grammar, syntax, and objectively inconsistent style conventions. Do not assess punctuation placement in relation to quotation marks; treat the punctuation and quotation marks in the source as outside the scope of this check. Do not recommend stylistic improvements or rewrites. For quoted material, report only clear spelling or grammar errors, and flag uncertainty where another language convention may apply.'

HEADER_RULE = "-" * 60
EXTRAS_RULE = "=" * 60

REPLY_FORMAT = (
    "Reply only with a list of the errors you find, one per line, in this format:\n"
    "- \"exact quoted text\" → suggested correction (spelling / grammar / syntax / punctuation / consistency)\n\n"
    "Quote the words exactly as they appear, and only the words concerned. You may add a brief reason "
    "of no more than 12 words after the brackets. If the same error occurs several times, list it once. "
    "If you find no errors, reply exactly: No errors found.\n"
    "Do not list anything that is not an error. Do not rewrite the passage. Do not comment on anything else. "
    "Do not return a corrected version of the full text."
)

OUTPUT_INSTRUCTIONS = (
    "The text below is one section of a longer document, extracted from a PDF. Paragraphs are complete "
    "and separated by blank lines; headings stand on their own lines. Footnote reference numbers have "
    "been removed, so do not flag missing references.\n\n" + REPLY_FORMAT
)

FOOTNOTE_INSTRUCTIONS = (
    "The text below contains the footnotes for one section of a longer document, extracted from a PDF. "
    "Each line is one footnote and begins with its number. Most footnotes are citations: check them for "
    "errors in spelling and punctuation, and for inconsistencies in citation style between these footnotes "
    "(for example in how authors, titles, publishers, dates and page ranges are given). "
    "Do not flag the footnote numbers themselves.\n\n" + REPLY_FORMAT
)

ABBREVIATIONS = {
    "mr.", "mrs.", "ms.", "dr.", "st.", "prof.", "rev.", "capt.", "col.", "gen.",
    "lt.", "sgt.", "jr.", "sr.", "esq.", "viz.", "ca.", "c.", "fl.", "cf.",
    "e.g.", "i.e.", "etc.", "ibid.", "p.", "pp.", "vol.", "vols.", "no.", "nos.",
    "ed.", "eds.", "trans.", "repr.", "fig.", "ch.", "chap.", "esp.", "op.", "cit.",
}

LIGATURES = {"\ufb00": "ff", "\ufb01": "fi", "\ufb02": "fl", "\ufb03": "ffi", "\ufb04": "ffl", "\ufb05": "st", "\ufb06": "st"}
WORD_SUFFIXES = ("s", "es", "ed", "d", "ing", "ly", "er", "ers", "est", "'s", "’s")

MARKER = re.compile(r"⟦(/i|i|p|n)([^⟧]*)⟧")
NOTE_MARKER = re.compile(r"⟦n[^⟧]*⟧")
SPLIT_ITALICS = re.compile(r"⟦/i⟧(\s*(?:⟦[pn][^⟧]*⟧)?\s*)⟦i⟧")
TAG = re.compile(r"</?i>", re.I)
NOTE_START = re.compile(r"^⟦n([^⟧]+)⟧\s*")
NOTE_REFERENCE_TEXT = re.compile(r"\d{1,4}(?:\s*[,–-]\s*\d{1,4})*|[*†‡§]+")
QUOTED_TEXT = re.compile(r"[\"“”]([^\"“”]+)[\"“”]")
SENTENCE_BOUNDARY = re.compile(r"[.!?][\"”’)\]]*(?:⟦[^⟧]*⟧)*\s+(?=(?:⟦[^⟧]*⟧)*[\"“‘(\[]*[A-Z])")
NON_SUGGESTION = re.compile(
    r"\bno (?:error|change|correction|issue)s?\b|\bcorrect as (?:written|is)\b|\bno changes? (?:needed|required)\b",
    re.I,
)
MATCH_TABLE = str.maketrans({
    "“": '"', "”": '"', "„": '"', "‘": "'", "’": "'", "ʼ": "'", "ʾ": "'", "ʿ": "'",
    "–": "-", "—": "-", "‑": "-", " ": " ",
})


def strip_markers(text):
    return MARKER.sub("", text)


def count_words(text):
    return len(strip_markers(text).split())


def estimate_tokens(text):
    return int(len(text.split()) * TOKENS_PER_WORD)


def normalise_margin_text(text):
    return re.sub(r"\d+", "#", strip_markers(text).lower().strip())


def is_page_number(text):
    return re.fullmatch(r"[\-–—\s]*(\d+|[ivxlcdm]+)[\-–—\s]*", strip_markers(text).strip(), re.I) is not None


def ends_sentence(text):
    return re.search(r"[.!?:;][\"”’)\]]*$", strip_markers(text).rstrip()) is not None


def join_text(previous, following, marker=""):
    if re.search(r"\S[-–—/](?:⟦/i⟧)?$", previous):
        return previous + marker + following
    return previous + " " + marker + following


def is_word_file(filename):
    return Path(filename).suffix.lower() in WORD_SUFFIXES_ACCEPTED


def conversion_error(error):
    return (getattr(error, "stderr", "") or str(error)).strip()


def convert_with_word(data, suffix):
    WORD_SANDBOX_TMP.mkdir(parents=True, exist_ok=True)
    folder = Path(tempfile.mkdtemp(prefix="proofreader-", dir=WORD_SANDBOX_TMP))
    try:
        source = folder / f"proofreader-document{suffix}"
        target = folder / "proofreader-document.pdf"
        source.write_bytes(data)
        subprocess.run(
            ["osascript", "-", str(source), str(target)], input=WORD_SCRIPT,
            text=True, capture_output=True, timeout=CONVERSION_TIMEOUT_SECONDS, check=True,
        )
        if not target.exists():
            raise OSError("Word did not produce a PDF.")
        return target.read_bytes()
    finally:
        shutil.rmtree(folder, ignore_errors=True)


def find_libreoffice():
    for candidate in LIBREOFFICE_PATHS:
        if Path(candidate).exists():
            return candidate
        found = shutil.which(candidate)
        if found:
            return found
    return None


def convert_with_libreoffice(data, suffix):
    with tempfile.TemporaryDirectory(prefix="proofreader-") as folder:
        source = Path(folder) / f"document{suffix}"
        source.write_bytes(data)
        profile = (Path(folder) / "profile").as_uri()
        subprocess.run(
            [find_libreoffice(), f"-env:UserInstallation={profile}", "--headless",
             "--convert-to", "pdf", "--outdir", folder, str(source)],
            text=True, capture_output=True, timeout=CONVERSION_TIMEOUT_SECONDS, check=True,
        )
        target = Path(folder) / "document.pdf"
        if not target.exists():
            raise OSError("LibreOffice did not produce a PDF.")
        return target.read_bytes()


def convert_word_to_pdf(data, filename):
    suffix = Path(filename).suffix.lower()
    problems = []
    if sys.platform == "darwin" and WORD_APP.exists():
        try:
            return convert_with_word(data, suffix), "Microsoft Word"
        except (subprocess.SubprocessError, OSError) as error:
            problems.append(f"Microsoft Word: {conversion_error(error)}")
    if find_libreoffice():
        try:
            return convert_with_libreoffice(data, suffix), "LibreOffice"
        except (subprocess.SubprocessError, OSError) as error:
            problems.append(f"LibreOffice: {conversion_error(error)}")
    if not problems:
        problems.append("neither Microsoft Word nor LibreOffice was found")
    raise RuntimeError(
        "Could not convert the Word document to PDF (" + "; ".join(problems) + "). "
        "Save it as a PDF from Word (File → Save As → PDF) and upload that instead."
    )


@lru_cache(maxsize=1)
def load_word_list():
    try:
        with open(WORD_LIST_PATH, encoding="utf-8", errors="ignore") as handle:
            return frozenset(line.strip().lower() for line in handle if line.strip())
    except OSError:
        return frozenset()


def is_known_word(word, known):
    word = word.lower()
    if word in known:
        return True
    return any(word.endswith(suffix) and word[:-len(suffix)] in known for suffix in WORD_SUFFIXES)


def origin_key(origin):
    return (round(origin[0], 1), round(origin[1], 1))


def trace_words(doc):
    words = []
    for page_index, page in enumerate(doc):
        for span in page.get_texttrace():
            current = []
            for unicode_value, glyph, origin, bbox in span["chars"]:
                character = chr(unicode_value)
                if character.isalpha() or character in "'’":
                    current.append((character, glyph, page_index, origin_key(origin)))
                elif current:
                    words.append((span["font"], current))
                    current = []
            if current:
                words.append((span["font"], current))
    return words


def find_misread_glyphs(doc):
    known = load_word_list()
    if not known:
        return {}, Counter()
    groups = defaultdict(list)
    for font, letters in trace_words(doc):
        for character, glyph, page_index, key in letters:
            others = [letter for letter in letters if letter[0] != character and letter[0].isalpha()]
            if character.isascii() and character.isalpha() and others:
                groups[(font, glyph, character)].append(letters)

    repairs = defaultdict(dict)
    summary = Counter()
    for (font, glyph, character), words in groups.items():
        spelled = ["".join(letter[0] for letter in word) for word in words]
        before = sum(is_known_word(word, known) for word in spelled) / len(spelled)
        if before >= 0.5:
            continue
        best_guess = None
        best_score = 0
        for guess in LIGATURE_GUESSES:
            repaired = [
                "".join(guess if (letter[0], letter[1]) == (character, glyph) else letter[0] for letter in word)
                for word in words
            ]
            score = sum(is_known_word(word, known) for word in repaired) / len(repaired)
            if score > best_score:
                best_guess, best_score = guess, score
        if best_guess and best_score >= 0.6 and best_score - before >= 0.4:
            for word in words:
                for letter_character, letter_glyph, page_index, key in word:
                    if (letter_character, letter_glyph) == (character, glyph):
                        repairs[page_index][key] = best_guess
                        summary[f"{character} → {best_guess}"] += 1
    return repairs, summary


def read_page_rows(page, repairs=None):
    repairs = repairs or {}
    raw = []
    for block in page.get_text("rawdict")["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                span["text"] = "".join(
                    repairs.get(origin_key(char["origin"]), LIGATURES.get(char["c"], char["c"]))
                    for char in span["chars"]
                )
            if any(span["text"].strip() for span in line["spans"]):
                raw.append({"y0": line["bbox"][1], "y1": line["bbox"][3], "spans": list(line["spans"])})
    raw.sort(key=lambda line: (line["y0"] + line["y1"]) / 2)
    rows = []
    for line in raw:
        centre = (line["y0"] + line["y1"]) / 2
        if rows:
            row = rows[-1]
            row_centre = (row["y0"] + row["y1"]) / 2
            if row["y0"] <= centre <= row["y1"] or line["y0"] <= row_centre <= line["y1"]:
                row["spans"].extend(line["spans"])
                row["y0"] = min(row["y0"], line["y0"])
                row["y1"] = max(row["y1"], line["y1"])
                continue
        rows.append(line)
    return rows


def is_italic(span):
    return bool(span["flags"] & 2) or "Italic" in span["font"] or "Oblique" in span["font"]


def merge_italics(text):
    return SPLIT_ITALICS.sub(r"\1", text)


def build_row(raw_row, page_label):
    spans = sorted(raw_row["spans"], key=lambda span: span["bbox"][0])
    weights = Counter()
    for span in spans:
        if span["text"].strip():
            weights[round(span["size"], 1)] += len(span["text"].strip())
    size = weights.most_common(1)[0][0]
    main_spans = [s for s in spans if round(s["size"], 1) == size and s["text"].strip()]
    baseline = max(span["origin"][1] for span in main_spans)
    main_chars = sum(len(span["text"]) for span in main_spans)
    bold_chars = sum(len(span["text"]) for span in main_spans if span["flags"] & 16 or "Bold" in span["font"])

    parts = []
    for span in spans:
        stripped = span["text"].strip()
        raised = span["origin"][1] < baseline - 1
        small = span["size"] < size * SUPERSCRIPT_SIZE_RATIO
        is_superscript = span["flags"] & 1 or (small and raised)
        if stripped and is_superscript and NOTE_REFERENCE_TEXT.fullmatch(stripped):
            number = re.sub(r"\s+", "", stripped)
            parts.append(f"⟦n{number}⟧")
        elif stripped and is_italic(span):
            text = span["text"]
            lead = text[:len(text) - len(text.lstrip())]
            trail = text[len(text.rstrip()):]
            parts.append(f"{lead}⟦i⟧{stripped}⟦/i⟧{trail}")
        else:
            parts.append(span["text"])
    visible = [span for span in spans if span["text"].strip()]
    return {
        "page": page_label,
        "x0": min(span["bbox"][0] for span in visible),
        "x1": max(span["bbox"][2] for span in visible),
        "y0": raw_row["y0"],
        "y1": raw_row["y1"],
        "size": size,
        "bold": bold_chars > 0.6 * main_chars,
        "text": merge_italics(re.sub(r"\s+", " ", "".join(parts)).strip()),
    }


def in_margin(row, height):
    return row["y1"] <= height * MARGIN_FRACTION or row["y0"] >= height * (1 - MARGIN_FRACTION)


def find_repeated_margin_text(pages):
    counts = Counter()
    for label, height, rows in pages:
        for text in {normalise_margin_text(row["text"]) for row in rows if in_margin(row, height)}:
            counts[text] += 1
    return {text for text, number in counts.items() if number >= 3}


def is_heading(row, body_size):
    return row["bold"] or row["size"] > body_size * HEADING_SIZE_FACTOR


def starts_paragraph(row, previous, layout):
    if previous is None:
        return True
    if is_heading(row, layout["body_size"]) or is_heading(previous, layout["body_size"]):
        return True
    same_page = row["page"] == previous["page"]
    if same_page and layout["pitch"] and row["y0"] - previous["y0"] > layout["pitch"] * PARAGRAPH_GAP_FACTOR:
        return True
    indented = row["x0"] > layout["left"] + INDENT_POINTS
    previous_indented = previous["x0"] > layout["left"] + INDENT_POINTS
    if indented and not (previous_indented and abs(row["x0"] - previous["x0"]) < 3):
        return True
    if previous["x1"] < layout["right"] - SHORT_LINE_POINTS and ends_sentence(previous["text"]):
        return True
    return False


def build_paragraphs(rows, body_size):
    if not rows:
        return []
    gaps = [b["y0"] - a["y0"] for a, b in zip(rows, rows[1:]) if a["page"] == b["page"] and b["y0"] > a["y0"]]
    layout = {
        "body_size": body_size,
        "left": Counter(round(row["x0"]) for row in rows).most_common(1)[0][0],
        "right": Counter(round(row["x1"]) for row in rows).most_common(1)[0][0],
        "pitch": statistics.median(gaps) if gaps else 0,
    }
    paragraphs = []
    previous = None
    for row in rows:
        if starts_paragraph(row, previous, layout):
            paragraphs.append({
                "text": row["text"],
                "start": row["page"],
                "end": row["page"],
                "heading": is_heading(row, body_size),
                "notes": [],
            })
        else:
            paragraph = paragraphs[-1]
            marker = f"⟦p{row['page']}⟧" if row["page"] != paragraph["end"] else ""
            paragraph["text"] = join_text(paragraph["text"], row["text"], marker)
            paragraph["end"] = row["page"]
        previous = row
    return paragraphs


def build_notes(rows):
    notes = []
    for row in rows:
        text = row["text"]
        number = None
        start = NOTE_START.match(text)
        if start:
            number = start.group(1)
            text = text[start.end():]
        else:
            text = NOTE_MARKER.sub("", text)
            fallback = re.match(r"^(\d{1,4})[.\s]\s*", text)
            last = notes[-1]["number"] if notes else None
            if fallback and (last is None or (last.isdigit() and int(fallback.group(1)) == int(last) + 1)):
                number = fallback.group(1)
                text = text[fallback.end():]
        text = NOTE_MARKER.sub("", text).strip()
        if number is not None or not notes:
            notes.append({"number": number or "?", "page": row["page"], "text": text})
        else:
            notes[-1]["text"] = join_text(notes[-1]["text"], text)
    for note in notes:
        note["text"] = merge_italics(note["text"])
        note["words"] = count_words(note["text"])
    return notes


def referenced_numbers(text):
    numbers = []
    for kind, content in MARKER.findall(text):
        if kind == "n":
            numbers.extend(re.findall(r"\d+|[*†‡§]", content))
    return numbers


def fallback_paragraph(paragraphs, page):
    candidates = [i for i, p in enumerate(paragraphs) if p["start"] <= page <= p["end"]]
    if candidates:
        return candidates[-1]
    earlier = [i for i, p in enumerate(paragraphs) if p["start"] <= page]
    return earlier[-1] if earlier else 0


def attach_notes(paragraphs, notes):
    owner = {}
    for index, paragraph in enumerate(paragraphs):
        for number in referenced_numbers(paragraph["text"]):
            owner.setdefault(number, index)
    unlinked = 0
    for note in notes:
        index = owner.get(note["number"])
        if index is None:
            unlinked += 1
            index = fallback_paragraph(paragraphs, note["page"])
        paragraphs[index]["notes"].append(note)
    return unlinked


def extract_document(pdf_bytes, page_offset=0, footnote_ratio=FOOTNOTE_SIZE_RATIO):
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    repairs, repaired = find_misread_glyphs(doc)
    pages = []
    size_chars = Counter()
    for index, page in enumerate(doc):
        label = index + 1 + page_offset
        rows = [build_row(raw_row, label) for raw_row in read_page_rows(page, repairs.get(index))]
        for row in rows:
            size_chars[row["size"]] += len(strip_markers(row["text"]))
        pages.append((label, page.rect.height, rows))
    doc.close()

    info = {"body_size": 0, "paragraphs": 0, "notes": 0, "unlinked": 0, "repaired": dict(repaired)}
    if not size_chars:
        return [], info
    body_size = size_chars.most_common(1)[0][0]
    repeated = find_repeated_margin_text(pages)

    body_rows = []
    note_rows = []
    for label, height, rows in pages:
        content = []
        for row in rows:
            furniture = is_page_number(row["text"]) or normalise_margin_text(row["text"]) in repeated
            if in_margin(row, height) and furniture:
                continue
            content.append(row)
        body_bottoms = [row["y1"] for row in content if row["size"] >= body_size * footnote_ratio]
        last_body = max(body_bottoms) if body_bottoms else 0
        for row in content:
            if row["size"] < body_size * footnote_ratio and row["y0"] >= last_body - 1:
                note_rows.append(row)
            else:
                body_rows.append(row)

    paragraphs = build_paragraphs(body_rows, body_size)
    for paragraph in paragraphs:
        paragraph["text"] = merge_italics(paragraph["text"])
    notes = build_notes(note_rows)
    unlinked = attach_notes(paragraphs, notes) if paragraphs else len(notes)
    info = {
        "body_size": body_size, "paragraphs": len(paragraphs), "notes": len(notes),
        "unlinked": unlinked, "repaired": dict(repaired),
    }
    return paragraphs, info


def split_sentences(text):
    sentences = []
    start = 0
    for match in SENTENCE_BOUNDARY.finditer(text):
        before = strip_markers(text[start:match.start() + 1]).split()
        last_word = before[-1].lower().lstrip("\"“‘([") if before else ""
        if last_word in ABBREVIATIONS or len(last_word) <= 2:
            continue
        sentences.append(text[start:match.end()].strip())
        start = match.end()
    if text[start:].strip():
        sentences.append(text[start:].strip())
    return sentences


def unit_words(unit):
    return count_words(unit["text"]) + sum(note["words"] for note in unit["notes"])


def split_paragraph(paragraph, target_words):
    pieces = []
    current = None
    page = paragraph["start"]
    used = set()
    for sentence in split_sentences(paragraph["text"]):
        lead = re.match(r"^⟦p(-?\d+)⟧", sentence)
        sentence_page = int(lead.group(1)) if lead else page
        numbers = referenced_numbers(sentence)
        sentence_notes = [n for n in paragraph["notes"] if id(n) not in used and n["number"] in numbers]
        used.update(id(note) for note in sentence_notes)
        words = count_words(sentence) + sum(note["words"] for note in sentence_notes)
        if current and current["words"] + words > target_words:
            pieces.append(current)
            current = None
        if current is None:
            current = {"text": sentence, "start": sentence_page, "end": sentence_page,
                       "heading": False, "notes": [], "words": 0}
        else:
            current["text"] += " " + sentence
        current["notes"].extend(sentence_notes)
        current["words"] += words
        for marker_page in re.findall(r"⟦p(-?\d+)⟧", sentence):
            page = int(marker_page)
        current["end"] = page
    if current:
        pieces.append(current)
    pieces[-1]["notes"].extend(n for n in paragraph["notes"] if id(n) not in used)
    return pieces


def new_render():
    return {"tagged": "", "plain": "", "italic": [], "breakpoints": [], "in_italic": False}


def add_plain(render, piece):
    if render["plain"] == "" or render["plain"].endswith("\n") or render["plain"].endswith(" "):
        piece = piece.lstrip(" ")
    render["plain"] += piece
    render["tagged"] += piece
    render["italic"].extend([render["in_italic"]] * len(piece))


def add_marked(render, text, page_point=None):
    position = 0
    for match in MARKER.finditer(text):
        add_plain(render, text[position:match.start()])
        kind = match.group(1)
        if kind == "p" and page_point:
            page_point(int(match.group(2)))
        elif kind == "i" and not render["in_italic"]:
            render["tagged"] += "<i>"
            render["in_italic"] = True
        elif kind == "/i" and render["in_italic"]:
            render["tagged"] += "</i>"
            render["in_italic"] = False
        position = match.end()
    add_plain(render, text[position:])
    if render["in_italic"]:
        render["tagged"] += "</i>"
        render["in_italic"] = False


def finish_render(render):
    plain = render["plain"].rstrip()
    return render["tagged"].rstrip(), plain, render["italic"][:len(plain)], render["breakpoints"]


def render_main(units):
    render = new_render()
    for unit in units:
        if render["plain"]:
            add_plain(render, "\n\n")
        render["breakpoints"].append((len(render["plain"]), unit["start"]))
        add_marked(render, unit["text"], lambda page: render["breakpoints"].append((len(render["plain"]), page)))
    return finish_render(render)


def render_notes(notes):
    render = new_render()
    for note in notes:
        if render["plain"]:
            add_plain(render, "\n")
        render["breakpoints"].append((len(render["plain"]), note["page"], note["number"]))
        add_plain(render, f"{note['number']}. ")
        add_marked(render, note["text"])
    return finish_render(render)


def build_chunk(number, units):
    main_text, main_plain, main_italic, main_map = render_main(units)
    notes = [note for unit in units for note in unit["notes"]]
    notes_text, notes_plain, notes_italic, notes_map = render_notes(notes)
    pages = [page for offset, page in main_map]
    body_words = count_words(main_plain)
    note_words = sum(note["words"] for note in notes)
    return {
        "number": number,
        "first_page": min(pages),
        "last_page": max(pages),
        "body_words": body_words,
        "note_words": note_words,
        "words": body_words + note_words,
        "note_count": len(notes),
        "main_text": main_plain,
        "main_plain": main_plain,
        "main_italic": main_italic,
        "main_map": main_map,
        "notes_text": notes_plain,
        "notes_plain": notes_plain,
        "notes_italic": notes_italic,
        "notes_map": notes_map,
    }


def make_chunks(paragraphs, target_words=TARGET_WORDS, include_notes=True):
    if not include_notes:
        paragraphs = [dict(paragraph, notes=[]) for paragraph in paragraphs]
    limit = int(target_words * MAX_CHUNK_FACTOR)
    units = []
    for paragraph in paragraphs:
        if unit_words(paragraph) > limit and not paragraph["heading"]:
            units.extend(split_paragraph(paragraph, target_words))
        else:
            units.append(paragraph)

    groups = []
    current = []
    current_words = 0
    for unit in units:
        words = unit_words(unit)
        if current:
            combined = current_words + words
            nearer = abs(combined - target_words) < abs(current_words - target_words)
            if not (combined <= limit and nearer):
                carried = []
                if len(current) > 1 and current[-1]["heading"]:
                    carried = [current.pop()]
                groups.append(current)
                current = carried
                current_words = sum(unit_words(u) for u in carried)
        current.append(unit)
        current_words += words
    if current:
        groups.append(current)

    if len(groups) > 1:
        last_words = sum(unit_words(u) for u in groups[-1])
        previous_words = sum(unit_words(u) for u in groups[-2])
        if last_words < 0.25 * target_words and previous_words + last_words <= limit:
            last_group = groups.pop()
            groups[-1] = groups[-1] + last_group

    return [build_chunk(number, group) for number, group in enumerate(groups, start=1)]


def page_range_label(chunk):
    if chunk["first_page"] == chunk["last_page"]:
        return f"p. {chunk['first_page']}"
    return f"pp. {chunk['first_page']}–{chunk['last_page']}"


def page_range_title(chunk):
    if chunk["first_page"] == chunk["last_page"]:
        return f"Page {chunk['first_page']}"
    return f"Pages {chunk['first_page']}–{chunk['last_page']}"


def chunk_filename(chunk):
    if chunk["first_page"] == chunk["last_page"]:
        pages = f"p{chunk['first_page']}"
    else:
        pages = f"pp{chunk['first_page']}-{chunk['last_page']}"
    return f"chunk_{chunk['number']:02d}_{pages}.txt"


def chunk_signature(chunk, model, prompt, num_ctx, think, temperature=TEMPERATURE):
    material = f"v5|{model}|{prompt}|{num_ctx}|{bool(think)}|{float(temperature)}|{chunk['main_text']}|{chunk['notes_text']}"
    return hashlib.sha1(material.encode("utf-8")).hexdigest()[:12]


def context_needed(chunk, prompt, think):
    allowance = THINK_REPLY_ALLOWANCE_TOKENS if think else REPLY_ALLOWANCE_TOKENS
    needed = estimate_tokens(prompt + " " + OUTPUT_INSTRUCTIONS + " " + chunk["main_text"])
    if chunk["notes_text"]:
        needed = max(needed, estimate_tokens(prompt + " " + FOOTNOTE_INSTRUCTIONS + " " + chunk["notes_text"]))
    return needed + allowance


def choose_context(chunks, prompt, think):
    needed = max((context_needed(chunk, prompt, think) for chunk in chunks), default=0)
    for size in CONTEXT_OPTIONS:
        if size >= needed:
            return size
    return CONTEXT_OPTIONS[-1]


def model_name_of(entry):
    try:
        return entry["model"]
    except KeyError:
        return entry["name"]


def list_models():
    return sorted(model_name_of(entry) for entry in ollama.list()["models"])


def get_model_info(model):
    info = {"name": model, "parameter_size": "", "quantisation": "", "digest": "", "supports_thinking": False}
    try:
        shown = ollama.show(model)
        details = getattr(shown, "details", None)
        info["parameter_size"] = getattr(details, "parameter_size", "") or ""
        info["quantisation"] = getattr(details, "quantization_level", "") or ""
        info["supports_thinking"] = "thinking" in (getattr(shown, "capabilities", None) or [])
        for entry in ollama.list()["models"]:
            if model_name_of(entry) == model:
                info["digest"] = (getattr(entry, "digest", "") or "")[:12]
    except Exception:
        pass
    return info


def describe_model(info):
    parts = ["Ollama"]
    if info["parameter_size"]:
        parts.append(info["parameter_size"] + " parameters")
    if info["quantisation"]:
        parts.append(info["quantisation"])
    if info["digest"]:
        parts.append("digest " + info["digest"])
    return f"{info['name']} ({', '.join(parts)})"


def think_setting(model, think):
    if "gpt-oss" in model.lower():
        return "medium" if think else "low"
    return bool(think)


def stream_reply(model, messages, options, think_value, on_progress):
    reply = ""
    thinking = ""
    last = None
    for part in ollama.chat(model=model, messages=messages, stream=True, options=options, think=think_value):
        message = part["message"]
        reply += message["content"] or ""
        thinking += getattr(message, "thinking", None) or ""
        last = part
        if on_progress:
            on_progress(reply, thinking)
    return reply, thinking, last


def ask_model(model, system_text, user_text, num_ctx, think=False, on_progress=None, temperature=TEMPERATURE):
    room = max(num_ctx - estimate_tokens(system_text + " " + user_text), 1024)
    max_tokens = min(room, THINK_REPLY_ALLOWANCE_TOKENS if think else MAX_REPLY_TOKENS)
    messages = [
        {"role": "system", "content": system_text},
        {"role": "user", "content": user_text},
    ]
    options = {"num_ctx": num_ctx, "temperature": temperature, "num_predict": max_tokens}
    try:
        reply, thinking, last = stream_reply(model, messages, options, think_setting(model, think), on_progress)
    except ollama.ResponseError as error:
        if "think" not in str(error).lower():
            raise
        reply, thinking, last = stream_reply(model, messages, options, None, on_progress)

    inline = re.findall(r"<think>(.*?)</think>", reply, flags=re.S)
    reply = re.sub(r"<think>.*?</think>", "", reply, flags=re.S).strip()
    prompt_tokens = getattr(last, "prompt_eval_count", None) or 0
    return {
        "reply": reply,
        "thinking": (thinking + "\n".join(inline)).strip(),
        "cut_off": getattr(last, "done_reason", None) == "length",
        "context_full": prompt_tokens >= num_ctx * CONTEXT_FULL_FRACTION,
        "prompt_tokens": prompt_tokens,
        "max_tokens": max_tokens,
    }


def run_pass(model, system_text, user_text, num_ctx, think, on_progress, temperature=TEMPERATURE):
    answer = ask_model(model, system_text, user_text, num_ctx, think, on_progress, temperature)
    fallback = False
    if think and (answer["cut_off"] or not answer["reply"]):
        answer = ask_model(model, system_text, user_text, num_ctx, False, on_progress, temperature)
        fallback = True
    if answer["context_full"]:
        raise RuntimeError(
            f"The text sent used {answer['prompt_tokens']} of the {num_ctx} tokens in the context window, "
            "so part of it was probably cut off. Choose a larger context size or a smaller chunk size."
        )
    if answer["cut_off"]:
        raise RuntimeError(f"The reply was cut off at the limit of {answer['max_tokens']} tokens.")
    if not answer["reply"]:
        raise RuntimeError("The model returned an empty reply. Try again, or try a different model.")
    answer["fallback"] = fallback
    return answer


def collapse_repeats(report):
    counts = {}
    order = []
    for line in report.splitlines():
        line = line.strip()
        if line not in counts:
            order.append(line)
            counts[line] = 0
        counts[line] += 1
    lines = []
    for line in order:
        if counts[line] > 1 and line:
            lines.append(f"{line} (×{counts[line]})")
        else:
            lines.append(line)
    return "\n".join(lines)


def normalise_for_match(text):
    return re.sub(r"\s+", " ", text.translate(MATCH_TABLE)).strip()


def same_text(first, second):
    quotes = str.maketrans({"“": '"', "”": '"', "‘": "'", "’": "'"})
    first = re.sub(r"\s+", " ", first.translate(quotes)).strip(" \"'")
    second = re.sub(r"\s+", " ", second.translate(quotes)).strip(" \"'")
    return first == second


def all_occurrences(haystack, needle):
    positions = []
    start = haystack.find(needle)
    while start != -1:
        positions.append(start)
        start = haystack.find(needle, start + 1)
    return positions


def find_positions(haystack, quote):
    needle = normalise_for_match(quote)
    if not needle:
        return []
    positions = all_occurrences(haystack, needle) or all_occurrences(haystack.lower(), needle.lower())
    if positions:
        return positions
    fragments = [part.strip() for part in re.split(r"…|\.\.\.", needle) if part.strip()]
    if len(fragments) > 1 and all(fragment.lower() in haystack.lower() for fragment in fragments):
        return all_occurrences(haystack.lower(), fragments[0].lower())
    return []


def location_at(breakpoints, offset):
    current = breakpoints[0]
    for point in breakpoints:
        if point[0] <= offset:
            current = point
        else:
            break
    return current


def location_label(breakpoints, positions, is_notes):
    places = []
    for position in positions:
        place = location_at(breakpoints, position)[1:]
        if place not in places:
            places.append(place)
    if not is_notes:
        pages = [str(place[0]) for place in places]
        return ("p. " if len(pages) == 1 else "pp. ") + ", ".join(pages)
    if len(places) == 1:
        return f"p. {places[0][0]}, n. {places[0][1]}"
    return "; ".join(f"n. {number} (p. {page})" for page, number in places)


def italics_verdict(quote_raw, correction_raw, tail, italic, position):
    if TAG.search(correction_raw) or TAG.search(quote_raw):
        wants_italic = "<i>" in correction_raw.lower()
    elif re.search(r"italic", tail, re.I):
        wants_italic = not re.search(r"\broman\b|\bremove\b|\bnot (?:be )?(?:in )?italic", tail, re.I)
    else:
        return "identical"
    if position is None:
        return "keep"
    length = len(normalise_for_match(TAG.sub("", quote_raw)))
    flags = italic[position:position + length]
    if wants_italic and flags and all(flags):
        return "already in italics"
    if not wants_italic and flags and not any(flags):
        return "already roman"
    return "keep"


def check_reply(reply, text, italic, breakpoints, is_notes):
    haystack = text.translate(MATCH_TABLE)
    kept = []
    removed = []
    for line in collapse_repeats(reply).splitlines():
        if not line or line.lower().startswith("no errors found"):
            continue
        bullet = re.match(r"^(?:[-*•]|\d+[.)])\s+", line)
        if bullet:
            body = line[bullet.end():]
        elif "→" in line or "->" in line:
            body = line
        else:
            removed.append(line)
            continue
        body = re.sub(r"^(?:pp?\.\s*[\d,\s–-]+|n\.\s*\d+|page\s*\d+)\s*[:,]\s*", "", body, flags=re.I)
        quote = QUOTED_TEXT.search(body)
        tail = body[quote.end():] if quote else body
        if NON_SUGGESTION.search(tail):
            removed.append(line)
            continue
        positions = []
        if quote:
            quote_raw = quote.group(1)
            positions = find_positions(haystack, TAG.sub("", quote_raw))
            correction = re.match(r"\s*(?:→|->|=>)\s*[\"“]?([^\"”(]*)", tail)
            correction_raw = correction.group(1) if correction else ""
            if correction and same_text(TAG.sub("", correction_raw), TAG.sub("", quote_raw)):
                verdict = italics_verdict(quote_raw, correction_raw, tail, italic, positions[0] if positions else None)
                if verdict == "identical":
                    removed.append(line)
                    continue
                if verdict != "keep":
                    removed.append(f"{line} [{verdict}]")
                    continue
        if positions:
            kept.append((positions[0], f"- {location_label(breakpoints, positions, is_notes)}: {body}"))
        else:
            kept.append((len(text) + 1, f"- page unknown: {body}"))
    kept.sort(key=lambda pair: pair[0])
    return [line for position, line in kept], removed


def proofread_chunk(chunk, model, prompt, num_ctx, think=False, on_progress=None, temperature=TEMPERATURE):
    passes = [("Main text", OUTPUT_INSTRUCTIONS, chunk["main_text"], chunk["main_plain"],
               chunk["main_italic"], chunk["main_map"], False)]
    if chunk["notes_text"]:
        passes.append(("Footnotes", FOOTNOTE_INSTRUCTIONS, chunk["notes_text"], chunk["notes_plain"],
                       chunk["notes_italic"], chunk["notes_map"], True))
    result = {"Main text": [], "Footnotes": [], "removed": [], "thinking": [], "fallback": False,
              "has_notes": bool(chunk["notes_text"])}
    for stage, instructions, text, plain, italic, breakpoints, is_notes in passes:
        callback = None
        if on_progress:
            callback = lambda reply, thinking, stage=stage: on_progress(stage, reply, thinking)
        answer = run_pass(model, prompt.strip() + "\n\n" + instructions, text, num_ctx, think, callback, temperature)
        lines, removed = check_reply(answer["reply"], plain, italic, breakpoints, is_notes)
        result[stage] = lines
        result["removed"].extend(f"{stage}: {line}" for line in removed)
        if answer["thinking"]:
            result["thinking"].append((stage, answer["thinking"]))
        result["fallback"] = result["fallback"] or answer["fallback"]
    return result


def full_prompt(prompt, language):
    language = language.strip()
    if not language:
        return prompt.strip()
    return (
        prompt.strip() + f"\n\nThe text is written in {language}. "
        f"Judge spelling, grammar and punctuation by the conventions of {language}."
    )


def underline(title, character="-"):
    return f"{title}\n{character * len(title)}"


def format_chunk_body(result):
    sections = []
    if result["fallback"]:
        sections.append("Note: the model thought for longer than its limit, so this chunk was redone with thinking off.")
    stages = ["Main text", "Footnotes"] if result["has_notes"] else ["Main text"]
    for stage in stages:
        lines = result[stage] or ["No errors found."]
        sections.append(stage.upper() + "\n\n" + "\n".join(lines))
    return "\n\n".join(sections)


def read_report_body(content):
    body = content.split("\n" + HEADER_RULE + "\n", 1)[1]
    return body.split("\n" + EXTRAS_RULE + "\n", 1)[0].strip()


def load_saved_report(path, signature):
    if not path.exists():
        return None
    content = path.read_text(encoding="utf-8")
    if f"Signature: {signature}" not in content:
        return None
    return read_report_body(content)


def save_chunk_report(path, chunk, total_chunks, model_text, prompt, num_ctx, think, signature, result,
                      temperature=TEMPERATURE):
    if not think:
        thinking_mode = "off"
    elif result["fallback"]:
        thinking_mode = "on, but switched off after a cut-off"
    else:
        thinking_mode = "on"
    header = (
        underline(f"Chunk {chunk['number']} of {total_chunks}: {page_range_label(chunk)}", "=") + "\n\n"
        f"Words: {chunk['words']} ({chunk['body_words']} main text, "
        f"{chunk['note_words']} in {chunk['note_count']} footnotes)\n"
        f"Model: {model_text}\n"
        f"Thinking mode: {thinking_mode}\n"
        f"Temperature: {temperature}\n"
        f"Context size: {num_ctx}\n"
        f"Prompt: {prompt.strip()}\n"
        f"Signature: {signature}\n"
    )
    text = header + "\n" + HEADER_RULE + "\n\n" + format_chunk_body(result) + "\n"
    extras = []
    if result["removed"]:
        extras.append("LINES REMOVED BY CHECKS\n\n" + "\n".join(f"- {line}" for line in result["removed"]))
    for stage, thinking in result["thinking"]:
        extras.append(f"MODEL'S THINKING ({stage.upper()})\n\n{thinking}")
    if extras:
        text += "\n" + EXTRAS_RULE + "\n\n" + "\n\n".join(extras) + "\n"
    path.write_text(text, encoding="utf-8")


def count_flagged(body):
    return sum(1 for line in body.splitlines() if line.startswith("- "))


def italic_runs(plain, italic):
    runs = []
    start = None
    for index, flag in enumerate(list(italic) + [False]):
        if flag and start is None:
            start = index
        elif not flag and start is not None:
            runs.append(plain[start:index])
            start = None
    return runs


def italic_term(run, known):
    term = run.strip().strip(" ,.;:()[]“”‘’\"'")
    words = term.split()
    if len(words) >= 2 and len(term) >= 5 and any(word[:1].isupper() for word in words):
        return term
    letters = sum(character.isalpha() for character in term)
    if len(words) == 1 and letters >= 3 and not term.isupper() and not is_known_word(term.strip("ʿʾ'’"), known):
        return term
    return None


def inside_quotation(plain, position):
    line_start = plain.rfind("\n", 0, position) + 1
    before = plain[line_start:position]
    if before.count("“") > before.count("”"):
        return True
    return before.count("‘") > len(re.findall(r"’(?!\w)", before))


def describe_places(places):
    labels = []
    for place in places:
        label = f"p. {place[0]}" if len(place) == 1 else f"p. {place[0]}, n. {place[1]}"
        if label not in labels:
            labels.append(label)
    return "; ".join(labels)


def italics_consistency(chunks):
    known = load_word_list()
    segments = []
    for chunk in chunks:
        segments.append((chunk["main_plain"], chunk["main_italic"], chunk["main_map"], False))
        if chunk["notes_plain"]:
            segments.append((chunk["notes_plain"], chunk["notes_italic"], chunk["notes_map"], True))

    terms = []
    for plain, italic, breakpoints, is_notes in segments:
        for run in italic_runs(plain, italic):
            term = italic_term(run, known)
            if term and term not in terms:
                terms.append(term)

    findings = []
    for term in terms:
        pattern = re.compile(r"(?<!\w)" + re.escape(term) + r"(?!\w)")
        found = {"italic": [], "roman": [], "partly italic": []}
        italic_in = set()
        other_in = set()
        first = None
        for segment_index, (plain, italic, breakpoints, is_notes) in enumerate(segments):
            for match in pattern.finditer(plain):
                flags = [italic[i] for i in range(match.start(), match.end()) if not plain[i].isspace()]
                if all(flags):
                    kind = "italic"
                    italic_in.add(is_notes)
                elif inside_quotation(plain, match.start()):
                    continue
                else:
                    kind = "roman" if not any(flags) else "partly italic"
                    other_in.add(is_notes)
                found[kind].append(location_at(breakpoints, match.start())[1:])
                if first is None:
                    first = (segment_index, match.start())
        if italic_in & other_in:
            parts = [f"{kind} {len(places)}× ({describe_places(places)})" for kind, places in found.items() if places]
            findings.append((first, f"- \"{term}\": " + "; ".join(parts)))
    findings.sort(key=lambda pair: pair[0])
    return [line for position, line in findings]


def build_full_report(chunks, out_dir, source_name, model_text, think, prompt, target_words, num_ctx,
                      converted_by=None, include_notes=True, language="", temperature=TEMPERATURE):
    out_dir = Path(out_dir)
    sections = []
    contents = [f"{'Chunk':<8}{'Pages':<16}{'Items flagged':>13}"]
    total_items = 0
    for chunk in chunks:
        content = (out_dir / chunk_filename(chunk)).read_text(encoding="utf-8")
        body = read_report_body(content)
        items = count_flagged(body)
        total_items += items
        contents.append(f"{chunk['number']:<8}{page_range_label(chunk):<16}{items:>13}")
        title = f"{page_range_title(chunk)} (chunk {chunk['number']} of {len(chunks)})"
        sections.append(underline(title, "=") + "\n\n" + body + "\n")
    contents.append(f"{'Total':<24}{total_items:>13}")

    source_line = f"Source file: {source_name}"
    if converted_by:
        source_line += f" (converted to PDF with {converted_by}; page numbers follow that layout)"
    lines = [
        underline(f"Proofreading report: {source_name}", "="),
        "",
        f"Proofread with {model_text}, a large language model run on this computer through Ollama. "
        f"Thinking mode: {'on' if think else 'off'}. "
        "Its suggestions can be wrong or invented and need checking by a human reader. "
        "Page and footnote numbers were added by the app, by finding each quoted passage in the text; "
        "'page unknown' means the quoted words could not be found exactly as the model gave them.",
        "",
        source_line,
        f"Date: {datetime.now():%d %B %Y, %H:%M}",
        f"Language: {language.strip() or 'not specified'}",
        f"Target chunk size: {target_words} words" + (", including footnotes" if include_notes else ""),
        f"Footnotes: {'proofread' if include_notes else 'not included'}",
        f"Temperature: {temperature}",
        f"Context size: {num_ctx}",
        f"Prompt: {prompt.strip()}",
        "",
        underline("Contents"),
        "",
    ]
    italics = italics_consistency(chunks) or ["No inconsistencies found."]
    italics_section = [
        underline("Italics consistency (checked by the app, not the model)"),
        "",
        "Titles and non-English words that are italic in some places but not in others. "
        "Check each roman or partly italic occurrence.",
        "",
    ] + italics + ["", ""]
    text = "\n".join(lines + contents + ["", ""] + italics_section + sections)
    (out_dir / full_report_filename(source_name)).write_text(text, encoding="utf-8")
    return text


def full_report_filename(source_name):
    return f"{Path(source_name).stem}_full_report.txt"


def existing_reports(folder):
    return sorted(list(Path(folder).glob("chunk_*.txt")) + list(Path(folder).glob("chunk_*.md")))


def choose_output_folder(base, name, mode):
    folder = Path(base) / name
    if mode == KEEP_BOTH and existing_reports(folder):
        number = 2
        while (Path(base) / f"{name}_{number}").exists():
            number += 1
        folder = Path(base) / f"{name}_{number}"
    elif mode == OVERWRITE:
        for path in existing_reports(folder):
            path.unlink()
        for path in list(folder.glob("*full_report.txt")) + list(folder.glob("*full_report.md")):
            path.unlink()
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def timing_key(model, think):
    return f"{model}|think={bool(think)}"


def load_timings(root):
    try:
        return json.loads((Path(root) / TIMINGS_FILE).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def update_timing(root, key, seconds, words):
    if words <= 0:
        return
    timings = load_timings(root)
    rate = seconds / words
    old = timings.get(key)
    timings[key] = rate if old is None else 0.7 * old + 0.3 * rate
    try:
        Path(root).mkdir(parents=True, exist_ok=True)
        (Path(root) / TIMINGS_FILE).write_text(json.dumps(timings, indent=2), encoding="utf-8")
    except OSError:
        pass
