import streamlit as st
import pandas as pd
from docx import Document
import io
from difflib import SequenceMatcher

def get_highlighted_diff(text1, text2):
    """
    Compares two strings and returns an HTML string with the parts of text1 
    that are not in text2 highlighted in yellow.
    """
    matcher = SequenceMatcher(None, text2, text1)
    
    # Get the operations needed to transform text2 into text1
    opcodes = matcher.get_opcodes()
    
    highlighted_text = []
    for tag, i1, i2, j1, j2 in opcodes:
        if tag == 'equal':
            # This part of text1 is present in text2
            highlighted_text.append(text1[j1:j2])
        elif tag == 'insert':
            # This part of text1 is NOT present in text2, so highlight it
            highlighted_text.append(f'<span style="background-color: #fdd835;">{text1[j1:j2]}</span>')
        elif tag == 'replace':
            # Part of text1 replaces something in text2, highlight the replacement
            highlighted_text.append(f'<span style="background-color: #fdd835;">{text1[j1:j2]}</span>')
        # 'delete' tag is ignored as we only care about text present in text1
        
    return "".join(highlighted_text)


def parse_docx(file_content):
    """
    Parses the uploaded .docx file content.
    Returns a dictionary mapping 'Ln:Tn' to the paragraph text.
    """
    doc_data = {}
    doc = Document(io.BytesIO(file_content))
    
    for i, para in enumerate(doc.paragraphs):
        # Skip empty paragraphs
        if not para.text.strip():
            continue
            
        line_number = i + 1
        # A tab in python-docx is represented by a '\t' character.
        tab_count = para.text.count('\t')
        
        # Strip leading/trailing whitespace and tabs for clean text content
        clean_text = para.text.strip()
        
        key = f"L{line_number}:T{tab_count}"
        doc_data[key] = clean_text
        
    return doc_data


def run_checker(df, doc_data):
    """
    Checks each row of the DataFrame against the parsed document data.
    Returns a list of dictionaries with the results.
    """
    results = []
    
    for index, row in df.iterrows():
        # Assuming sentence is in column D and location is in column F
        # Adjust 'D' and 'F' if your columns have names
        try:
            sentence_to_check = str(row['D']).strip()
            location = str(row['F']).strip()
        except KeyError as e:
            st.error(f"Error: Missing expected column in Excel file. Please ensure you have columns 'D' and 'F'. Details: {e}")
            return None

        result_item = {
            "excel_row": index + 2, # Excel rows are 1-based, plus header
            "location": location,
            "sentence": sentence_to_check,
            "status": "",
            "details": ""
        }

        # Check if the location from Excel exists in the parsed DOCX data
        if location in doc_data:
            doc_text = doc_data[location]
            
            # Check if the sentence from Excel is a substring of the text in the DOCX
            if sentence_to_check in doc_text:
                result_item["status"] = "✅ Correct"
                result_item["details"] = f"The sentence was found exactly as stated in the document at `{location}`."
            else:
                result_item["status"] = "❌ Incorrect"
                # Generate highlighted difference
                highlighted_diff = get_highlighted_diff(sentence_to_check, doc_text)
                result_item["details"] = f"The sentence was **not** found as stated. Differences are highlighted below:"
                result_item["highlighted"] = highlighted_diff
                result_item["doc_text"] = doc_text
        else:
            result_item["status"] = "❌ Error"
            result_item["details"] = f"The specified location `{location}` was not found in the DOCX file."

        results.append(result_item)
        
    return results


# --- Streamlit App UI ---

st.set_page_config(layout="wide")
st.title("📄 Extractive Sentence Checker Tool")

st.info("""
    **How to use this tool:**
    1.  Upload the meeting minutes as a `.docx` file (บันทึกการประชุม).
    2.  Upload the corresponding Excel file with sentences to check.
    3.  The tool will verify if each sentence from the Excel file exists at the specified location in the Word document.
""")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload DOCX File (บันทึกการประชุม)")
    docx_file = st.file_uploader("Upload your .docx file", type=["docx"])

with col2:
    st.subheader("2. Upload Excel File")
    xlsx_file = st.file_uploader("Upload your .xlsx file", type=["xlsx"])


if docx_file is not None and xlsx_file is not None:
    
    st.markdown("---")
    st.header("Results")

    try:
        # Read file contents
        docx_content = docx_file.getvalue()
        
        # Parse the documents
        doc_data = parse_docx(docx_content)
        df = pd.read_excel(xlsx_file, header=None, names=['A', 'B', 'C', 'D', 'E', 'F', 'G'])

        # Run the checker
        results = run_checker(df, doc_data)

        if results:
            # Display results
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
        st.error("Please ensure the uploaded files are in the correct format and not corrupted.")
