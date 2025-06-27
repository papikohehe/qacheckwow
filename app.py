import streamlit as st
import pandas as pd
from docx import Document
import io
from difflib import SequenceMatcher
import re

# ---------------------------------------------------------------------------
# Helper: highlight differences between two strings
# ---------------------------------------------------------------------------

def get_highlighted_diff(text1: str, text2: str) -> str:
    """Return *text1* with segments that do **not** occur in *text2* highlighted
    yellow using inline HTML.  We compare *text1* against *text2* so we can show
    the missing / altered parts visually to the user.
    """
    matcher = SequenceMatcher(None, text2, text1, autojunk=False)
    pieces: list[str] = []
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        fragment = text1[j1:j2]
        if tag == "equal":
            pieces.append(fragment)
        else:  # "replace", "insert" → highlight what is *only* in text1
            pieces.append(
                f"<span style=\"background-color:#fdd835;\">{fragment}</span>"
            )
    return "".join(pieces)


# ---------------------------------------------------------------------------
# DOCX parsing
# ---------------------------------------------------------------------------

def parse_docx(file_content: bytes):
    """Read a .docx file and build two maps:

    * *doc_data* maps a **full key** like ``L65:T3`` to the cleaned paragraph
      text (only for **non-empty** paragraphs).
    * *line_to_key_map* maps a **line number only** (65) to the key that exists
      *for that exact line* – including empty paragraphs.  This preserves the
      original Word line numbering so that Excel references stay aligned.
    """
    doc_data: dict[str, str] = {}
    line_to_key_map: dict[int, str] = {}

    doc = Document(io.BytesIO(file_content))

    for line_no, para in enumerate(doc.paragraphs, start=1):
        tab_count = para.text.count("\t")
        key = f"L{line_no}:T{tab_count}"
        line_to_key_map[line_no] = key
        clean_text = para.text.strip()
        if clean_text:
            doc_data[key] = clean_text

    return doc_data, line_to_key_map


# ---------------------------------------------------------------------------
# Location parsing (Excel → real keys)
# ---------------------------------------------------------------------------

def parse_location_string(
    location_str: str,
    line_to_key_map: dict[int, str],
    doc_data: dict[str, str],
):
    """Translate a location string from Excel into a list of keys.

    *Single* location – we respect the exact ``T`` value when present, so
    ``L65:T3`` matches **only** that precise paragraph.  If the user gives
    ``L65:C`` (sometimes they note check-columns) we fall back to the default
    paragraph for that line (whatever *Word* actually has).

    *Range* – we keep the previous behaviour: use only the line numbers so
    "L21:T0 - L24:T3" means *lines* 21–24, regardless of their actual tab
    indentation.
    """
    location_str = location_str.strip()

    # ------------------------------------------------------------------
    # Range handling
    # ------------------------------------------------------------------
    if " - " in location_str:
        try:
            start_loc, end_loc = (p.strip() for p in location_str.split(" - ", 1))
            line_re = r"L(\d+):(?:T\d+|C)"
            start_m, end_m = re.match(line_re, start_loc), re.match(line_re, end_loc)
            if not (start_m and end_m):
                return []
            start_ln, end_ln = int(start_m.group(1)), int(end_m.group(1))
            return [
                line_to_key_map.get(ln)
                for ln in range(start_ln, end_ln + 1)
                if line_to_key_map.get(ln)
            ]
        except Exception:
            return []

    # ------------------------------------------------------------------
    # Single location
    # ------------------------------------------------------------------
    single_re = r"L(?P<line>\d+):(?:(?:T(?P<tab>\d+))|C)"
    m = re.match(single_re, location_str)
    if not m:
        return []

    line_num = int(m.group("line"))
    tab_val = m.group("tab")

    if tab_val is not None:  # Explicit T… given – require exact match
        key = f"L{line_num}:T{int(tab_val)}"
        return [key] if key in doc_data else []

    # "C" shorthand – use whatever tab count the doc actually has
    key = line_to_key_map.get(line_num)
    return [key] if key else []


# ---------------------------------------------------------------------------
# Main checker
# ---------------------------------------------------------------------------

def run_checker(df: pd.DataFrame, doc_data: dict, line_to_key_map: dict[int, str]):
    """Validate each Excel row and build a list of result dictionaries."""
    results: list[dict] = []

    if len(df.columns) < 6:
        st.error("Error: The Excel file must have at least 6 columns.")
        return None

    sentence_col = df.columns[3]  # 4th column
    location_col = df.columns[5]  # 6th column

    for idx, row in df.iterrows():
        sentence = str(row.get(sentence_col, "")).strip()
        location_raw = str(row.get(location_col, "")).strip()
        if not sentence or not location_raw or sentence.lower() == "nan" or location_raw.lower() == "nan":
            continue  # skip blank lines

        result = {
            "excel_row": idx + 2,  # +2 (header row + zero-based index)
            "location": location_raw,
            "sentence": sentence,
            "status": "",
            "details": "",
        }

        loc_keys = parse_location_string(location_raw, line_to_key_map, doc_data)
        if not loc_keys:
            result["status"] = "❌ Error"
            result["details"] = (
                f"The location `{location_raw}` could not be resolved. "
                "Check the format (e.g. `L21:T0`, `L24:C`, or `L21:T0 - L24:T3`)."
            )
            results.append(result)
            continue

        doc_texts = [doc_data[key] for key in loc_keys if ke
