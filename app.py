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
    Parses the uploaded .docx file content, numbering ALL paragraphs sequentially.
    Returns two dictionaries:
    - doc_data: maps 'Ln:Tn' to the paragraph text.
    - line_to_key_map: maps a line number (int) to its full key 'Ln:Tn'.
    """
    doc_data = {}
    line_to_key_map = {}
    doc = Document(io.BytesIO(file_content))
    
    # Initialize a counter that increments for *every* paragraph, regardless of content
    true_line_number_counter = 0 
    
    for para in doc.paragraphs:
        true_line_number_counter += 1 # Increment for every paragraph encountered
        
        tab_count = para.text.count('\t')
        clean_text = para.text.strip() # Still strip whitespace for comparison later
        
        # The key uses the 'true_line_number_counter' for 'L'
        key = f"L{true_line_number_counter}:T{tab_count}"
        
        # Store the cleaned text. If it was an empty paragraph, clean_text will be empty.
        doc_data[key] = clean_text
        
        # Map this true line number to its key
        line_to_key_map[true_line_number_counter] = key
        
    return doc_data, line_to_key_map

def parse_location_string(location_str, line_to_key_map):
    """
    Parses a location string (e.g., "L21:C", "L21:T0 - L24:T3") and returns a list of valid doc_data keys.
    Uses line_to_key_map to find the actual key for a given line number.
    """
    location_str = location_str.strip()
    
    # Regex to capture the line number from formats like L21:T0, L24:C, etc.
    range_regex = r"L(\d+):[TC]\d*"
    
    # Handle ranges
    if " - " in location_str:
        try:
            start_loc, end_loc = location_str.split(" - ")
            start_match = re.match(range_regex, start_loc.strip())
            end_match = re.match(range_regex, end_loc.strip())

            if not start_match or not end_match:
                return [] 

            start_l = int(start_match.group(1))
            end_l = int(end_match.group(1))
            
            keys = []
            for line_num in range(start_l, end_l + 1):
                key = line_to_key_map.get(line_num)
                if key: # Only add if a corresponding key exists (i.e., that line number was parsed)
                    keys.append(key)
            return keys
        except Exception:
             return []
    # Handle single location
    else:
        try:
            match = re.match(range_regex, location_str)
            if not match:
                return []
            
            line_num = int(match.group(1))
            key = line_to_key_map.get(line_num)
            return [key] if key else []
        except Exception:
            return []


def run_checker(df, doc_data, line_to_key_map):
    """
    Checks each row of the DataFrame against the parsed document data.
    Handles single locations and ranges.
    For ranges, checks if the sentence exists in any single line within the range.
    Returns a list of dictionaries with the results.
    """
    results = []
    
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
            "excel_row": index + 2, # This shows the Excel row number (1-based, accounting for header)
            "location": location_str,
            "sentence": sentence_to_check,
            "status": "",
            "details": ""
        }
        
        location_keys = parse_location_string(location_str, line_to_key_map)
        
        if not location_keys:
            result_item["status"] = "❌ Error"
            result_item["details"] = f"The specified location {location_str} could not be found or was in an invalid format. Please check line numbers and format (e.g., L21:T0, L24:C). Ensure the line number exists in the document."
        else:
            doc_texts = [doc_data.get(key) for key in location_keys if doc_data.get(key) is not None]
            
            is_found = False
            for text in doc_texts:
                if sentence_to_check in text:
                    is_found = True
                    break # Found a match, no need to check other lines in the range

            if is_found:
                result_item["status"] = "✅ Correct"
                result_item["details"] = f"The sentence was found exactly as stated in the document within the range {location_str}."
            else:
                # If not found, combine text for highlighting context
                full_doc_text = " ".join(doc_texts)
                result_item["status"] = "❌ Incorrect"
                # If the combined text is empty (e.g., trying to check an empty line)
                if not full_doc_text.strip():
                    result_item["details"] = f"The sentence was **not** found in the specified location {location_str}. The document content at this location appears to be empty."
                else:
                    highlighted_diff = get_highlighted_diff(sentence_to_check, full_doc_text)
                    result_item["details"] = f"The sentence was **not** found in any single line within the specified range. Differences compared to the combined text from the range are highlighted below:"
                    result_item["highlighted"] = highlighted_diff
                    result_item["doc_text"] = full_doc_text

        results.append(result_item)
        
    return results


# --- Streamlit App UI ---

st.set_page_config(layout="wide")
st.title("📄 Extractive Sentence Checker Tool")

st.info("""
    **How to use this tool:**
    1. Upload the meeting minutes as a .docx file (บันทึกการประชุม).
    2. Upload the corresponding Excel file with sentences to check. The Excel file should have a header row.
    3. The tool will verify if each sentence from the Excel file exists at the specified location in the Word document.
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
        doc_data, line_to_key_map = parse_docx(docx_content)
        
        df = pd.read_excel(xlsx_file, header=0) 

        results = run_checker(df, doc_data, line_to_key_map)

        if results:
            for res in results:
                with st.expander(f"**Row {res['excel_row']} | Location: {res['location']} | Status: {res['status']}**"):
                    st.markdown(f"**Sentence to Check:**")
                    st.markdown(f"> {res['sentence']}")
                    st.markdown(f"**Details:** {res['details']}")
                    
                    if res.get("highlighted"):
                        st.markdown("**Highlighted Differences:**")
                        st.markdown(res["highlighted"], unsafe_allow_html=True)
                        st.markdown("**Original Text from Document:**")
                        st.markdown(f"> {res['doc_text']}")
                    elif res.get("doc_text") == "": # Specific case for empty line in document
                         st.markdown("**Original Text from Document:** (Empty at this location)")


    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")
        st.error("Please ensure the uploaded files are valid and not corrupted. The Excel file must have a header row.")
