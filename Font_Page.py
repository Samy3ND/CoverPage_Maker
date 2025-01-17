import streamlit as st
from docx import Document
from datetime import datetime
from io import BytesIO
import os
import pypandoc
from docx.shared import Pt
import tempfile

# Student list for mapping
roll_number_to_name = {
    '022BIM001': 'Aabha Kumhal', 
    '022BIM003': 'Aarchi Palikhel', 
    '022BIM004': 'Aarohan Shakya', 
    '022BIM005': 'Aayush Ghimire', 
    '022BIM006': 'Abhilasha Adhikari',
    # Add the rest of the student mappings here...
}

# Subject Teacher names
subject_to_teacher = {
    "System Design & Development [IT 242]": "Er. Sanjay Kumar Yadav",
    "Python": "Mr Ramesh Shahi [IT 243]",
    "Artificial Intelligence [IT 288]": "Er Nischal Shrestha",
    "Information Security [IT 244]": "Er. Saroj Shahi",
}

def replace_in_paragraph(paragraph, replacements):
    for run in paragraph.runs:
        for key, value in replacements.items():
            if key in run.text:
                run.text = run.text.replace(key, value)
                run.font.name = "Times New Roman"  # Set font to Times New Roman
                run.font.size = Pt(12)  # Set font size to 12 pt
    return paragraph

def replace_placeholders(doc, name, roll_number, lab_report_number, subject, teacher):
    current_date = datetime.now().strftime("%d-%m-%Y")
    replacements = {
        "{Name}": name,
        "{RollNumber}": roll_number,
        "{LabReportNumber}": lab_report_number,
        "{Date}": current_date,
        "{Subject}": subject,
        "{Teacher}": teacher,
    }

    # Replace placeholders in paragraphs
    for paragraph in doc.paragraphs:
        replace_in_paragraph(paragraph, replacements)

    # Replace placeholders in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_in_paragraph(paragraph, replacements)

    return doc

def convert_docx_to_pdf(doc, output_name):
    try:
        # Save the DOCX file to a temporary location using tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_doc_file:
            temp_doc_path = temp_doc_file.name
            doc.save(temp_doc_path)  
            print(f"Saved DOCX to: {temp_doc_path}")

        # Set the output PDF file path
        pdf_file_path = temp_doc_path.replace(".docx", ".pdf")
        print(f"PDF will be saved to: {pdf_file_path}") 

        # Use pypandoc to convert DOCX to PDF
        output = pypandoc.convert_file(temp_doc_path, 'pdf', outputfile=pdf_file_path)
        print(f"PDF saved to: {pdf_file_path}")

    except Exception as e:
        print(f"Error during DOCX to PDF conversion: {e}")
        raise Exception(f"Error during DOCX to PDF conversion: {e}")
    finally:
        # Cleanup temporary DOCX file
        if os.path.exists(temp_doc_path):
            os.remove(temp_doc_path)
            print(f"Temporary DOCX file removed: {temp_doc_path}")

    # Read the PDF into memory
    with open(pdf_file_path, "rb") as pdf_file:
        pdf_bytes = BytesIO(pdf_file.read())
    os.remove(pdf_file_path)

    return pdf_bytes

# Streamlit UI
st.title("Lab Report Cover Page Maker")
st.header("Only Applicable for 5th Sem of SXC")

roll_number_input = st.text_input("Roll Number:").strip()

if roll_number_input:
    name = roll_number_to_name.get(roll_number_input, "")
    if name:
        st.text_input("Name:", value=name, disabled=True)  
    else:
        st.warning(f"No name found for Roll Number {roll_number_input}. Please verify the input.")

subject = st.selectbox(
    "Subject:",
    ["System Design & Development [IT 242]", "Python [IT 243]", "Artificial Intelligence [IT 288]", "Information Security [IT 244]"]
)

teacher = subject_to_teacher.get(subject, "Teacher not assigned")
st.text_input("Teacher:", value=teacher, disabled=True) 

lab_report_number = st.number_input("Lab Report Number:", min_value=1, step=1)

if st.button("Generate Cover Page"):
    if name and roll_number_input and lab_report_number:
        try:
            with st.spinner("Generating your cover page..."):
                doc = Document()  # Create a new document instead of loading from file

                updated_doc = replace_placeholders(doc, name, roll_number_input, str(lab_report_number), subject, teacher)

                pdf_bytes = convert_docx_to_pdf(updated_doc, "cover_page")
            
                st.success("Cover page generated successfully!")
                st.download_button(
                    label="Download Cover Page",
                    data=pdf_bytes,
                    file_name="Lab_Report_Cover.pdf",
                    mime="application/pdf",
                )
        except Exception as e:
            st.error(f"Error: {e}")
    else:
        st.error("Please fill out all the fields!")
