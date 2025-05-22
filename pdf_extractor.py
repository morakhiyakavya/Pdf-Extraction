import fitz  # PyMuPDF
import pandas as pd
import re

PDF_FILE = "DDx Tabelle.pdf"
OUTPUT_EXCEL = "structured_output_final_full_table.xlsx"

KNOWN_SECTIONS = ["Ursachen", "Untersuchungen", "Wichtige Hinweise", "Alarmsignale"]
SUBSUB_HEADINGS = ["häufig", "gelegentlich", "selten"]

def is_heading(text):
    text = text.strip()
    if text.startswith("•"):
        return False
    if len(text.split()) > 5:
        return False
    return bool(re.match(r"^[\wäöüÄÖÜß() \-]+$", text))

def is_subsub(text_lines, spans):
    if not text_lines:
        return False
    first_line = text_lines[0].strip().lower()
    if first_line not in SUBSUB_HEADINGS:
        return False
    if not spans:
        return False
    first_span = spans[0]
    return (
        abs(first_span.get("size", 0) - 8.0) < 0.2 and
        "Bold" in first_span.get("font", "")
    )

def is_bullet(text):
    return text.strip().startswith("•")

def is_section(text):
    return text.strip() in KNOWN_SECTIONS

def is_possible_chapter(text, font_size):
    return is_heading(text) and font_size >= 12

def is_noise_block(text_lines):
    joined = " ".join(text_lines).strip().lower()
    words = joined.split()
    if all(word in {"ja", "nein"} or word.isdigit() for word in words):
        return True
    if len(text_lines) <= 3 and not re.search(r"\b[a-zäöü]{5,}\b", joined, re.IGNORECASE):
        return True
    return False

def looks_like_table_row(line):
    words = line.strip().split()
    score = sum(1 for w in words if w.lower() in {"ja", "nein", "eventuell", "kein", "keine", "nicht", "evtl."})
    return len(words) > 8 and score >= 3

def is_table_block(block):
    count_7_5 = 0
    total_spans = 0
    words_7_5 = []

    for line in block.get("lines", []):
        for span in line.get("spans", []):
            text = span.get("text", "").strip()
            size = span.get("size", 0)
            total_spans += 1

            if abs(size - 7.5) < 0.2 and text:
                count_7_5 += 1
                words_7_5.append(text)

    # Print all words that were size 7.5
    if words_7_5:
        print(f"🟡 Page suspected table content (size 7.5): {words_7_5}")

    # Only consider block as table if 7+ 7.5-sized words AND they make up most of the block
    is_table = count_7_5 >= 5 and (count_7_5 / max(total_spans, 1)) >= 0.6

    print(f"📊 Block {block.get('number', '?')} | Font size 7.5 count: {count_7_5} | Total spans: {total_spans} | Is table: {is_table}")
    return is_table


def is_header_or_footer(text_lines, page_num, current_chapter):
    if isinstance(text_lines, str):
        lines = [text_lines]
    else:
        lines = text_lines

    for line in lines:
        cleaned = re.sub(r"\s+", " ", line.strip())

        # Detect lines like "3 8 Allgemeinsymptome bei Erwachsenen"
        if re.match(r"^\d+ \d+ [\wÄÖÜäöüß\s\-]{3,}$", cleaned):
            return True

        if re.fullmatch(r"(\d\s*){1,5}", cleaned):
            return True

        if re.fullmatch(r"(flush\s*)?\d{1,3}", cleaned, flags=re.IGNORECASE):
            return True

        if current_chapter:
            current_chap_clean = re.sub(r"\s+", " ", current_chapter).strip()
            if cleaned == current_chap_clean:
                return True
            if re.fullmatch(rf"{re.escape(current_chap_clean)}(\s+\d\s*)*", cleaned):
                return True

        if re.match(r"^\d{1,3} [A-ZÄÖÜa-zäöüß\- ]{3,}$", cleaned):
            return True

        lower_words = cleaned.lower().split()
        if lower_words:
            matrix_terms = {"ja", "nein", "eventuell"}
            matrix_count = sum(w in matrix_terms for w in lower_words)
            if matrix_count >= len(lower_words) * 0.6 and len(lower_words) >= 3:
                return True

        if len(cleaned) < 10 and any(c.isdigit() for c in cleaned):
            return True

    return False


def flush(chapter, section, subsection, buffer, page_num, final=False):
    if not buffer.strip() or not section:
        return

    global current_buffer_data
    lines = buffer.splitlines()
    bullet_lines = []
    current_bullet = ""

    for line in lines:
        stripped = line.strip()
        if stripped.startswith("•"):
            if current_bullet:
                bullet_lines.append(current_bullet.strip())
            current_bullet = stripped
        else:
            if stripped:
                current_bullet += " " + stripped


    if current_bullet:
        bullet_lines.append(current_bullet.strip())

    page_tag = f"\n(Page {page_num + 1})"
    bullet_text = "\n".join(bullet_lines) + page_tag if bullet_lines else buffer.strip() + page_tag

    if current_buffer_data is None:
        label = f"From Section: {section}"
        if subsection:
            label += f"\nSubsection: {subsection}"
        content = f"{label}\n{bullet_text}"
        current_buffer_data = {
            "Chapter": chapter,
            "SectionContent": content
        }
    else:
        current_buffer_data["SectionContent"] += f"\n{bullet_text}"

    if final:
        data_rows.append(current_buffer_data)
        print(f"✅ Flushed FINAL: {chapter} | {section}" + (f" | {subsection}" if subsection else ""))
        current_buffer_data = None

# State
doc = fitz.open(PDF_FILE)
data_rows = []

current_chapter = None
current_section = None
current_subsection = None
current_buffer_data = None
buffer = ""

# Main loop
for page_num,page in enumerate(doc):
    print(f"\n=== Processing Page {page_num} ===")
    page = doc.load_page(page_num)
    blocks = page.get_text("dict")["blocks"]

    for block_index, block in enumerate(blocks):
        text_lines = []
        max_font_size = 0


        for line in block.get("lines", []):
            line_text = " ".join([span["text"] for span in line["spans"]]).strip()
            text_lines.append(line_text)
            max_font_size = max(max_font_size, *(span["size"] for span in line["spans"]))

        filtered_lines = [line for line in text_lines if not is_header_or_footer(line, page_num, current_chapter)]
        if not filtered_lines:
            continue

        full_text = "\n".join(filtered_lines).strip()
        if not full_text:
            continue

        # ➤ Detect new chapter
        if is_possible_chapter(full_text, max_font_size) and full_text not in KNOWN_SECTIONS + SUBSUB_HEADINGS:
            if current_chapter != full_text:
                flush(current_chapter, current_section, current_subsection, buffer, page_num, final=True)
                if current_buffer_data:
                    data_rows.append(current_buffer_data)
                    current_buffer_data = None
                print(f"📘 New chapter detected: {full_text} (page {page_num})")
                current_chapter = full_text
                current_section = None
                current_subsection = None
                buffer = ""
            continue

        # ➤ Skip noise or repeated headers
        if is_noise_block(text_lines) or is_header_or_footer(text_lines, page_num, current_chapter):
            continue
        
       
        # ➤ Detect new section
        if is_section(full_text):
            if current_section != full_text:  # Only flush if it's a real new section
                flush(current_chapter, current_section, current_subsection, buffer, page_num, final=True)
                if current_buffer_data:
                    data_rows.append(current_buffer_data)
                    current_buffer_data = None
                print(f"--- Section detected: {full_text}")
                current_section = full_text
                current_subsection = None
                buffer = ""
            continue

        # ➤ Detect sub-subsection
        spans = block["lines"][0]["spans"] if block.get("lines") and block["lines"][0].get("spans") else []
        if is_subsub(text_lines, spans):
            if current_subsection != text_lines[0].strip():  # Only flush on real change
                flush(current_chapter, current_section, current_subsection, buffer, page_num, final=True)
                if current_buffer_data:
                    data_rows.append(current_buffer_data)
                    current_buffer_data = None
                current_subsection = text_lines[0].strip()
                print(f"--- Sub-subsection detected: {current_subsection}")
                buffer = ""
                for line in text_lines[1:]:
                    buffer += line + "\n"
            continue

        # ➤ Accumulate content if under section/subsection
        if current_section or current_subsection:
            print(f"[Block {block_index}] Adding content under Section: '{current_section}', Subsection: '{current_subsection}'")
            for line in filtered_lines:

                is_bullet_line = is_bullet(line)
                has_7_5_size = any(
                    abs(span.get("size", 0) - 7.5) < 0.2 and span.get("text", "").strip() in line
                    for l in block.get("lines", [])
                    for span in l.get("spans", [])
                )
                print(f"  Line: {line} | Bullet: {is_bullet_line} | Font size 7.5: {has_7_5_size}")
                SPECIAL_WORDS = {"häufig", "gelegentlich", "selten", "wichtige hinweise", "alarmsignale", "untersuchungen", "ursachen"}
                if not is_bullet_line and has_7_5_size:
                    lower_line = line.lower().strip()
                    if not any(word in lower_line for word in SPECIAL_WORDS):
                        print(f"⚠️ Skipping non-bullet line with font size 7.5 (no special keywords): {line}")
                        continue
                if is_bullet(line):
                    print(f"  Bullet found: {line}")
                    buffer += line + "\n"
                else:
                    buffer += line + "\n\n"

# ❗️Only flush once, after loop ends — to save final accumulated content
flush(current_chapter, current_section, current_subsection, buffer, page_num, final=True)
if current_buffer_data:
    data_rows.append(current_buffer_data)


# Export to Excel
df = pd.DataFrame(data_rows)
df.to_excel(OUTPUT_EXCEL, index=False)
print(f"\n✅ Extraction complete.")
print(f"📄 Text saved to: {OUTPUT_EXCEL}")
