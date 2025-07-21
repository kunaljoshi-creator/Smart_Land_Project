# Smart Legal Document Generator with Streamlit and OpenAI

import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO
import os
import google.generativeai as genai
from zipfile import ZipFile

# Initialize Gemini
genai.configure(api_key="please enter your api key")
model = genai.GenerativeModel('gemini-2.0-flash')

# App title
st.title("📄 Smart Legal Document Generator for Land Department")

# Session states
if 'excel_uploaded' not in st.session_state:
    st.session_state.excel_uploaded = False
if 'doc_uploaded' not in st.session_state:
    st.session_state.doc_uploaded = False

# Step 1: Excel Upload
st.sidebar.header("1️⃣ Upload Excel Data")
uploaded_excel = st.sidebar.file_uploader("Upload Excel Data", type=["xlsx"], key="excel")
if uploaded_excel is not None and st.sidebar.button("✅ Confirm Excel Upload"):
    st.session_state.df = pd.read_excel(uploaded_excel)
    st.session_state.df = st.session_state.df.astype(str)  # ✅ Convert all columns to string
    st.session_state.excel_uploaded = True
    st.sidebar.success("Excel file uploaded successfully!")

# Step 2: Word Template Upload
if st.session_state.excel_uploaded:
    st.sidebar.header("2️⃣ Upload Word Template")
    uploaded_docx = st.sidebar.file_uploader("Upload Word Template", type=["docx"], key="docx")
    if uploaded_docx is not None and st.sidebar.button("✅ Confirm Template Upload"):
        try:
            temp_docx = BytesIO(uploaded_docx.getvalue())
            doc = DocxTemplate(temp_docx)

            try:
                _ = doc.get_docx()
                st.session_state.doc = doc
                st.session_state.doc_buffer = temp_docx
                st.session_state.doc_uploaded = True
                st.sidebar.success("Word template uploaded successfully!")
            except Exception:
                st.sidebar.error("Template verification failed. Please ensure it's a valid Word document.")
                st.session_state.doc_uploaded = False

        except Exception:
            st.sidebar.error("Please upload a valid Word document (.docx)")
            st.session_state.doc_uploaded = False

# Main processing
if st.session_state.excel_uploaded and st.session_state.doc_uploaded:
    st.subheader("📊 Excel Data Preview")
    df_preview = st.session_state.df.astype(str)  # ✅ Safe conversion for preview
    st.dataframe(df_preview.head())

    try:
        def extract_placeholders(docx_template):
            placeholders = set()
            for para in docx_template.paragraphs:
                if '{{' in para.text and '}}' in para.text:
                    parts = para.text.split('{{')
                    for part in parts[1:]:
                        placeholder = part.split('}}')[0].strip()
                        placeholders.add(placeholder)
            return list(placeholders)

        placeholders = extract_placeholders(st.session_state.doc)
        if placeholders:
            st.success(f"Found {len(placeholders)} placeholders in template.")
        else:
            st.warning("No placeholders found in the template. Please check for {{placeholder_name}} format.")

        if st.button("🤖 Use GenAI to Map Fields and Generate Documents"):
            with st.spinner("Thinking with Gemini..."):
                prompt = f"""
You are an expert system for mapping document fields. Your task is to create precise matches between placeholders and Excel columns.

Given placeholders: {placeholders}
Excel columns available: {list(st.session_state.df.columns)}

Requirements:
1. Map fields exactly as specified:
   - survey_no -> "Survey No."
   - actual_area -> "Area(Ha)"
   - payment -> "Payment"
   - acquired_area -> "Acquired Area Sq.M"
   - award_rate -> "Rate as per the award"
   - rate_sq_mtr -> "Demanded Rate per Sq.M"
   - case_no -> "Case No."
   - village -> "Name of Village"
   - award_date -> "Award Date"
   - applicant_name -> "Applicant Name"
   - tahsil -> "Tahsil"
   - district -> "District"
   - date_of_notification -> "Date of Notification 3 (A)"

2. Pay special attention to the applicant_name field as it's critical for document naming.
3. Return a valid JSON object with exact matches.
4. Ensure all placeholder names match exactly as given.

Return only the JSON mapping without any additional text or explanation.
"""
                response = model.generate_content(prompt)
                response_text = response.text.strip()

                required_mappings = {
                        "applicant_name": "Applicant Name",
                        "actual_area": "Area(Ha)",
                        "payment": "Payment",
                        "survey_no": "Survey No.",
                        "acquired_area": "Acquired Area Sq.M",
                        "award_rate": "Rate as per the award",
                        "rate_sq_mtr": "Demanded Rate per Sq.M",
                        "case_no": "Case No.",
                        "village": "Name of Village",
                        "award_date": "Award Date",
                        "tahsil": "Tahsil",
                        "district": "District",
                        "date_of_notification": "Date of Notification 3 (A)"
                    }

                try:
                    if '```' in response_text:
                        response_text = response_text.split('```')[1]
                        if response_text.startswith('json'):
                            response_text = response_text[4:]
                    response_text = response_text.strip().replace('null', 'None')

                    mapping = eval(response_text.strip())
                    mapping = {k.strip(): v.strip() for k, v in mapping.items()}

                    for key, value in required_mappings.items():
                        if key not in mapping or not mapping[key]:
                            mapping[key] = value

                except Exception:
                    mapping = required_mappings
                    st.info("Using default mapping for document generation")

                st.write("🔗 Placeholder Mapping:", mapping)

            # Document Generation
            zip_buffer = BytesIO()
            with ZipFile(zip_buffer, 'a') as zip_file:
                for i, row in st.session_state.df.iterrows():
                    doc_instance = DocxTemplate(BytesIO(st.session_state.doc_buffer.getvalue()))
                    context = {}
                    for placeholder, column in mapping.items():
                        context[placeholder] = str(row[column]) if column in row else ""
                    doc_instance.render(context)

                    file_buffer = BytesIO()
                    case_num = str(row[mapping['case_no']]).replace('/', ' ').replace('-', ' ')
                    filename = f"{case_num}.docx"
                    doc_instance.save(file_buffer)
                    zip_file.writestr(filename, file_buffer.getvalue())

            st.success("✅ Documents generated successfully!")
            st.download_button(
                label="⬇️ Download All Documents as ZIP",
                data=zip_buffer.getvalue(),
                file_name="generated_documents.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"Error processing template: {str(e)}")

else:
    st.warning("Please upload both Excel data and Word template to proceed.")

# Reset button
if st.sidebar.button("🔄 Reset Upload Process"):
    st.session_state.excel_uploaded = False
    st.session_state.doc_uploaded = False
    st.experimental_rerun()
