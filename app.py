import streamlit as st
import pandas as pd
from docx import Document
import io
from difflib import SequenceMatcher
import re


def get_highlighted_diff(text1: str, text2: str) -> str:
    """Return *text1* with portions that do not occur in *text2* wrapped in a yellow
    highlight span (HTML).
    """
    matcher = SequenceMatcher(None, text2, text1, autojunk=False)
    highlighted_text = []

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        fragment = text1[j1:j2]
        if tag == "equal":
            highlighted_text.append(fragment)
        else:  # "insert" or "replace" both mean the token is *only* in text1
            highlighted_text.append(
                f"<span style=\"background-color:#fdd835;\">{fragment}</span>"
            )

    return "".join(highlighted_text)


# ---------------------------------------------------------------------------
# DOCX helpers
# ---------------------------------------------------------------------------

def parse_docx(file_content: bytes):
    """Parse *file_content* (raw bytes of a .docx file).

    A Word *.docx* is internally a list of paragraphs ― *including* empty ones.
    The meeting‑minutes conventions referenced by the Excel file number lines
    exactly the same way Word does, so we must **not** skip empty paragraphs
    when assigning the L‑numbers.

    However, we still want to ignore empty paragraphs when we later search for
    sentences.  Therefore we:

    * map **every** paragraph index (starting at 1) to a key of the form
      ``L<n>:T<tab_count>`` and save that in *line_to_key_map*;
    * only add *non‑empty* paragraphs to *doc_data* (key → paragraph text).
    """
    doc_data: dict[str, str] = {}
    line_to_key_map: dict[int, str] = {}

    doc = Document(io.BytesIO(file_content))

    for idx, para in enumerate(doc.paragraphs, start=1):
        tab_count = para.text.count("\t")
        key = f"L{idx}:T{tab_count}"
        line_to_key_map[idx] = key  # always record the mapping so numbering is exact

        clean_text = para.text.strip()
        if clean_text:  # only keep non‑empty paragraphs for later text lookup
            doc_data[key] = clean_text

    return doc_data, line_to_key_map


# ---------------------------------------------------------------------------
# Excel‑location helpers
# ---------------------------------------------------------------------------

def parse_location_string(location_str: str, line_to_key_map: dict[int, str]):
    """Convert a location like ``"L21:T0"`` or ``"L21:T0 - L24:T3"`` into a list
    of real keys that exist in *doc_data*.
    """
    location_str = location_str.strip()
    range_regex = r"L(\d+):[TC]\d*"  # capture the *line* number only

    # Range – e.g. "L21:T0 - L24:T3"
    if " - " in location_str:
        try:
            start_loc, end_loc = (part.strip() for part in location_str.split(" - ", 1))
            start_match, end_match = re.match(range_regex, start_loc), re.match(range_regex, end_loc)
            if not (start_match and end_match):
                return []
            start_l, end_l = int(start_match.group(1)), int(end_match.group(1))
            return [line_to_key_map.get(ln) for ln in range(start_l, end_l + 1) if line_to_key_map.get(ln)]
        except Exception:
            return []

    # Single location
    match = re.match(range_regex, location_str)
    if not match:
        return []
    line_num = int(match.group(1))
    key = line_to_key_map.get(line_num)
    return [key] if key else []


# ---------------------------------------------------------------------------
# Checker core
# ---------------------------------------------------------------------------

def run_checker(df: pd.DataFrame, doc_data: dict, line_to_key_map: dict[int, str]):
    """Iterate through *df* rows and validate each sentence against *doc_data*."""
    results = []

    # Safety‑net: the spec says at least six columns
    if len(df.columns) < 6:
        st.error("Error: The Excel file must have at least 6 columns.")
        return None

    sentence_col_name = df.columns[3]  # 4th column (0‑based index 3)
    location_col_name = df.columns[5]  # 6th column (0‑based index 5)

    for index, row in df.iterrows():
        # Skip blank / NaN rows early so that errors later are real ones
        sentence_to_check = str(row.get(sentence_col_name, "")).strip()
        location_str = str(row.get(location_col_name, "")).strip()
        if not sentence_to_check or not location_str or sentence_to_check.lower() == "nan" or location_str.lower() == "nan":
            continue

        result_item: dict[str, str | int] = {
            "excel_row": index + 2,  # +2 because DataFrame index 0 == Excel row 2 (header row is 1)
            "location": location_str,
            "sentence": sentence_to_check,
            "status": "",  # will be updated below
            "details": "",
        }

        location_keys = parse_location_string(location_str, line_to_key_map)

        if not location_keys:
            result_item["status"] = "❌ Error"
            result_item["details"] = (
                f"The specified location `{location_str}` could not be resolved. "
                "Please check the line numbers and format (e.g. `L21:T0`, `L24:C`)."
            )
            results.append(result_item)
            continue

        # Gather all (non‑empty) document texts that fall within the location(s)
        doc_texts = [doc_data[key] for key in location_keys if key in doc_data]

        # --------------------------------------------------------
        # EXACT MATCH?
        # --------------------------------------------------------
        if any(sentence_to_check in text for text in doc_texts):
            result_item["status"] = "✅ Correct"
            result_item["details"] = (
                "The sentence was found exactly as stated in the document "
                f"within the specified location `{location_str}`."
            )
        else:
            # Build a combined string from the location block for diffing/highlight
            full_doc_text = " ".join(doc_texts)
            result_item["status"] = "❌ Incorrect"
            result_item["details"] = (
                "The sentence was **not** found in the specified location. "
                "Differences compared to the text in that block are highlighted below:"  # noqa: E501
            )
            result_item["highlighted"] = get_highlighted_diff(sentence_to_check, full_doc_text)
            result_item["doc_text"] = full_doc_text

        results.append(result_item)

    return results


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

st.set_page_config(layout="wide")
st.title("📄 Extractive Sentence Checker Tool")

st.info(
    """
    **How to use this tool:**
    1. Upload the meeting minutes as a `.docx` file (บันทึกการประชุม).
    2. Upload the corresponding Excel file with sentences to check. The Excel file should have a header row.
    3. The tool will verify if each sentence from the Excel file exists at the specified location in the Word document.
    """
)

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload DOCX File (บันทึกการประชุม)")
    docx_file = st.file_uploader("Upload your .docx file", type=["docx"], key="docx")

with col2:
    st.subheader("2. Upload Excel File")
    xlsx_file = st.file_uploader("Upload your .xlsx file", type=["xlsx"], key="xlsx")


if docx_file is not None and xlsx_file is not None:
    st.markdown("---")
    st.header("Results")

    try:
        # ----------------------------------------------------
        # Parse the Word doc and the Excel workbook
        # ----------------------------------------------------
        doc_data, line_to_key_map = parse_docx(docx_file.getvalue())
        df = pd.read_excel(xlsx_file, header=0)

        # ----------------------------------------------------
        # Run the checker
        # ----------------------------------------------------
        results = run_checker(df, doc_data, line_to_key_map)

        # ----------------------------------------------------
        # Display results
        # ----------------------------------------------------
        if results:
            for res in results:
                expander_title = (
                    f"**Row {res['excel_row']} | Location: {res['location']} | "
                    f"Status: {res['status']}**"
                )
                with st.expander(expander_title):
                    st.markdown("**Sentence to Check:**")
                    st.markdown(f"> {res['sentence']}")
                    st.markdown("**Details:** " + res["details"])

                    if res.get("highlighted"):
                        st.markdown("**Highlighted Differences:**")
                        st.markdown(res["highlighted"], unsafe_allow_html=True)

                        st.markdown("**Original Text from Document:**")
                        st.markdown(f"> {res['doc_text']}")

    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")
        st.error(
            "Please ensure the uploaded files are valid and not corrupted. "
            "The Excel file must have a header row."
        )
