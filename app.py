import streamlit as st
import pandas as pd
from docx import Document
import io
from difflib import SequenceMatcher
import re

"""
Streamlit app ‑ *Extractive Sentence Checker*
============================================

Excel files produced by the client refer to locations in the meeting‑minutes
Word document using **the labels that already exist inside the DOCX itself** –
lines are explicitly written like

    L104:T4: วันพุธที่ ๑๒ มิถุนายน …

To avoid any off‑by‑one problems, we therefore **parse those labels directly**
instead of inventing our own numbering.  This guarantees a 1‑to‑1 match
between Excel references (e.g. ``L104:T4``) and the internal dictionary keys.
"""

# ---------------------------------------------------------------------------
# Helper: highlight textual differences
# ---------------------------------------------------------------------------

def get_highlighted_diff(source: str, target: str) -> str:
    """Return *source* with segments that are **not** present in *target*
    wrapped in a yellow span (for Streamlit markdown).
    """
    matcher = SequenceMatcher(None, target, source, autojunk=False)
    out: list[str] = []
    for tag, _, _, j1, j2 in matcher.get_opcodes():
        frag = source[j1:j2]
        if tag == "equal":
            out.append(frag)
        else:  # insert / replace – highlight
            out.append(f"<span style=\"background-color:#fdd835;\">{frag}</span>")
    return "".join(out)


# ---------------------------------------------------------------------------
# DOCX parsing – use the *actual* labels already in the file
# ---------------------------------------------------------------------------

def parse_docx(file_content: bytes):
    """Return two dictionaries based on the in‑document ``L…:T…:`` labels.

    * **doc_data** maps a key like ``L104:T4`` → *paragraph text without the
      leading label*.
    * **line_to_key_map** maps only the *line number* (104) → ``L104:T4`` so
      that ranges can be resolved quickly.
    """
    label_re = re.compile(r"^L(?P<line>\d+):T(?P<tab>\d+):\s*(?P<text>.*)")

    doc_data: dict[str, str] = {}
    line_to_key_map: dict[int, str] = {}

    document = Document(io.BytesIO(file_content))
    for para in document.paragraphs:
        m = label_re.match(para.text)
        if not m:
            # Paragraphs that do not start with a label are ignored for matching
            # They cannot be referenced from Excel anyway.
            continue

        line_no = int(m["line"])
        tab_no = int(m["tab"])
        key = f"L{line_no}:T{tab_no}"
        line_to_key_map[line_no] = key
        doc_data[key] = m["text"].strip()

    return doc_data, line_to_key_map


# ---------------------------------------------------------------------------
# Parse Excel‑style location strings → list of keys
# ---------------------------------------------------------------------------

def parse_location_string(location: str, line_to_key_map: dict[int, str], doc_data: dict[str, str]):
    location = location.strip()

    # -------- Ranges (e.g. "L104:T4 - L105:T4") --------------------------
    if " - " in location:
        try:
            start, end = (p.strip() for p in location.split(" - ", 1))
            num_re = r"L(\d+):(?:T\d+|C)"
            s_m, e_m = re.match(num_re, start), re.match(num_re, end)
            if not (s_m and e_m):
                return []
            s_ln, e_ln = int(s_m.group(1)), int(e_m.group(1))
            return [line_to_key_map.get(ln) for ln in range(s_ln, e_ln + 1) if line_to_key_map.get(ln)]
        except Exception:
            return []

    # -------- Single location --------------------------------------------
    single_re = r"L(?P<line>\d+):(?:(?:T(?P<tab>\d+))|C)"
    m = re.match(single_re, location)
    if not m:
        return []

    ln = int(m["line"])
    tab_val = m["tab"]

    # *Exact* tab requested → require exact key
    if tab_val is not None:
        key = f"L{ln}:T{int(tab_val)}"
        return [key] if key in doc_data else []

    # "C" shorthand: return whatever key exists for that line number
    key = line_to_key_map.get(ln)
    return [key] if key else []


# ---------------------------------------------------------------------------
# Checker core
# ---------------------------------------------------------------------------

def run_checker(df: pd.DataFrame, doc_data: dict[str, str], line_to_key_map: dict[int, str]):
    results: list[dict] = []

    if len(df.columns) < 6:
        st.error("The Excel file must contain at least 6 columns (header row included).")
        return None

    sentence_col, location_col = df.columns[3], df.columns[5]

    for idx, row in df.iterrows():
        sentence = str(row.get(sentence_col, "")).strip()
        location = str(row.get(location_col, "")).strip()
        if not sentence or not location or sentence.lower() == "nan" or location.lower() == "nan":
            continue  # skip empty rows

        res = {
            "excel_row": idx + 2,  # DataFrame row 0 == Excel row 2
            "location": location,
            "sentence": sentence,
            "status": "",
            "details": "",
        }

        keys = parse_location_string(location, line_to_key_map, doc_data)
        if not keys:
            res["status"] = "❌ Error"
            res["details"] = (
                f"Location `{location}` could not be resolved – check the format "
                "or confirm that the corresponding label exists in the DOCX."
            )
            results.append(res)
            continue

        # Build two parallel lists so we can both test and show nicely later
        texts = []  # label‑stripped → matching test
        raw_blocks = []  # label+text → display
        for k in keys:
            if k in doc_data:
                texts.append(doc_data[k])
                raw_blocks.append(f"{k}: {doc_data[k]}")

        # ---------------- exact match test --------------------------------
        if any(sentence in t for t in texts):
            res["status"] = "✅ Correct"
            res["details"] = "Sentence found exactly at the specified location."
        else:
            combined = " ".join(texts)
            res["status"] = "❌ Incorrect"
            res["details"] = (
                "Sentence **not** found at that location. Differences versus the "
                "text in that block are highlighted below:"
            )
            res["highlighted"] = get_highlighted_diff(sentence, combined)
            res["doc_text"] = " │ ".join(raw_blocks)

        results.append(res)

    return results


# ---------------------------------------------------------------------------
# Streamlit user interface
# ---------------------------------------------------------------------------

st.set_page_config(layout="wide")
st.title("📄 Extractive Sentence Checker Tool")

st.info(
    """
    **How to use this tool:**

    1. Upload the meeting‑minutes Word file (`.docx`) *which already contains* line
       labels (e.g. `L104:T4:`).
    2. Upload the corresponding Excel file. Column‑4 should contain the sentence,
       column‑6 the location string (e.g. `L104:T4` or `L104:T4 - L105:T4`).
    3. Results for each row will appear below.
    """
)

col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Upload DOCX File (บันทึกการประชุม)")
    docx_file = st.file_uploader("Upload .docx", type=["docx"], key="docx")
with col2:
    st.subheader("2. Upload Excel File")
    xlsx_file = st.file_uploader("Upload .xlsx", type=["xlsx"], key="xlsx")

if docx_file and xlsx_file:
    st.markdown("---")
    st.header("Results")

    try:
        doc_data, line_to_key_map = parse_docx(docx_file.getvalue())
        df = pd.read_excel(xlsx_file, header=0)
        results = run_checker(df, doc_data, line_to_key_map)

        if results:
            for r in results:
                with st.expander(
                    f"**Row {r['excel_row']} | Location: {r['location']} | Status: {r['status']}**"
                ):
                    st.markdown("**Sentence to Check:**")
                    st.markdown(f"> {r['sentence']}")
                    st.markdown("**Details:** " + r["details"])

                    if r.get("highlighted"):
                        st.markdown("**Highlighted Differences:**")
                        st.markdown(r["highlighted"], unsafe_allow_html=True)
                        st.markdown("**Original Text from Document:**")
                        st.markdown(f"> {r['doc_text']}")

    except Exception as e:
        st.error(f"Unexpected error: {e}")
        st.error("Please ensure both files are valid and follow the expected formats.")
