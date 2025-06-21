import streamlit as st
import pandas as pd
from docx import Document
import io
from difflib import SequenceMatcher
import re

def get_highlighted_diff(text1, text2):
    """
    Compares two strings and returns an HTML string with the parts of text1 
    that are not in text2 highlighted in yellow.
    """
    matcher = SequenceMatcher(None, text2, text1, autojunk=False)
    
    opcodes = matcher.get_opcodes()
    
    highlighted_text = []
    for tag, i1, i2, j1, j2 in opcodes:
        if tag == 'equal':
            highlighted_text.append(text1[j1:j2])
        elif tag == 'insert':
            highlighted_text.append(f'<span style="background-color: #fdd835;">{text1[j1:j2]}</span>')
        elif tag == 'replace':
            highlighted_text.append(f'<span style="background-color: #fdd835;">{text1[j1:j2]}</span>')
        
    return "".join(highlighted_text)


def parse_docx(file_content):
    """
    Parses the uploaded .docx file content.
    Returns a dictionary mapping 'Ln:Tn' to the paragraph text.
    """
    doc_data = {}
    doc = Document(io.BytesIO(file_content))
    
    for i, para in enumerate(doc.paragraphs):
        if not para.text.strip():
            continue
            
        line_number = i + 1
        tab_count = para.text.count('\t')
        
        clean_text = para.text.strip()
        
        key = f"L{line_number}:T{tab_count}"
        doc_data[key] = clean_text
        
    return doc_data

def parse_location_string(location_str):
    """
    Parses a location string, which can be a single location or a range.
    Returns a list of location keys.
    e.g., "L2:T2" -> ["L2:T2"]
    e.g., "L2:T2 - L4:T2" -> ["L2:T2", "L3:T2", "L4:T2"]
    """
    location_str = location_str.strip()
    if " - " in location_str:
        try:
            start_loc, end_loc = location_str.split(" - ")
            start_match = re.match(r"L(\d+):T(\d+)", start_loc.strip())
            end_match = re.match(r"L(\d+):T(\d+)", end_loc.strip())
            
            if not start_match or not end_match:
                return [location_str] 
                
            start_l, start_t = int(start_match.group(1)), int(start_match.group(2))
            end_l, end_t = int(end_match.group(1)), int(end_match.group(2))

            if start_t != end_t:
                 return [location_str]

            return [f"L{i}:T{start_t}" for i in range(start_l, end_l + 1)]
        except Exception:
            return [location_str]
    else:
        return [location_str]


def run_checker(df, doc_data):
    """
    Checks each row of the DataFrame against the parsed document data.
    Handles single locations and ranges (e.g., "L2:T2 - L4:T2").
    Returns a list of dictionaries with the results.
    """
    results = []
    
    # Expect sentence in 4th col (D) and location in 6th col (F)
    if len(df.columns) < 6:
        st.error("Error: The Excel file must have at least 6 columns.")
        return None
        
    sentence_col_name = df.columns[3]
    location_col_name = df.columns[5]
    
    for index, row in df.iterrows():
        try:
            sentence_to_check = str(row[sentence_col_name]).strip()
            location_str = str(row[location_col_name]).strip()
            
            if not sentence_to_check or not location_str or sentence_to_check.lower() == 'nan' or location_str.lower() == 'nan':
                continue

        except (KeyError, IndexError) as e:
            st.error(f"Error accessing columns in Excel file. Details: {e}")
            return None

        result_item = {
            "excel_row": index + 2, # Excel rows are 1-based, plus header
            "location": location_str,
            "sentence": sentence_to_check,
            "status": "",
            "details": ""
        }
        
        location_keys = parse_location_string(location_str)
        
        doc_texts = [doc_data.get(key) for key in location_keys]

        if None in doc_texts:
            missing_keys = [key for i, key in enumerate(location_keys) if doc_texts[i] is None]
            result_item["status"] = "❌ Error"
            result_item["details"] = f"The specified location(s) `{', '.join(missing_keys)}` were not found in the DOCX file."
        else:
            full_doc_text = " ".join(doc_texts)
            
            if sentence_to_check in full_doc_text:
                result_item["status"] = "✅ Correct"
                result_item["details"] = f"The sentence was found exactly as stated in the document within `{location_str}`."
            else:
                result_item["status"] = "❌ Incorrect"
                highlighted_diff = get_highlighted_diff(sentence_to_check, full_doc_text)
                result_item["details"] = f"The sentence was **not** found as stated. Differences are highlighted below:"
                result_item["highlighted"] = highlighted_diff
                result_item["doc_text"] = full_doc_text

        results.append(result_item)
        
    return results


# --- Streamlit App UI ---

st.set_page_config(layout="wide")
st.title("📄 Extractive Sentence Checker Tool")

st.info("""
    **How to use this tool:**
    1.  Upload the meeting minutes as a `.docx` file (บันทึกการประชุม).
    2.  Upload the corresponding Excel file with sentences to check. The Excel file should have a header row.
    3.  The tool will verify if each sentence from the Excel file exists at the specified location in the Word document.
""")

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
        docx_content = docx_file.getvalue()
        doc_data = parse_docx(docx_content)
        
        # Use header=0 to treat the first row as the header
        df = pd.read_excel(xlsx_file, header=0) 

        results = run_checker(df, doc_data)

        if results:
            for res in results:
                with st.expander(f"**Row {res['excel_row']} | Location: {res['location']} | Status: {res['status']}**"):
                    st.markdown(f"**Sentence to Check:**")
                    st.markdown(f"> {res['sentence']}")
                    st.markdown(f"**Details:** {res['details']}")
                    
                    if res["status"] == "❌ Incorrect":
                        st.markdown("**Highlighted Differences:**")
                        st.markdown(res["highlighted"], unsafe_allow_html=True)
                        st.markdown("**Original Text from Document:**")
                        st.markdown(f"> {res['doc_text']}")

    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")
        st.error("Please ensure the uploaded files are valid and not corrupted. The Excel file must have a header row.")
