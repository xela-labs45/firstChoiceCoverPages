import streamlit as st
import os
import io
from docx import Document
from docxcompose.composer import Composer
from docx.shared import Pt
from datetime import datetime

# --- Configuration & Setup ---
st.set_page_config(page_title="Student Cover Page Generator", layout="wide")

# --- Helper Functions ---

def replace_placeholder(doc, placeholder, replacement):
    """
    Replaces a placeholder in a python-docx Document object.
    Robust version: Tries run-level replacement first, falls back to paragraph-level.
    """
    replacement = str(replacement)
    
    # 1. Search in paragraphs
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            # Attempt run-level replacement to preserve formatting
            replaced_in_run = False
            for run in paragraph.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, replacement)
                    replaced_in_run = True
            
            # Fallback: If not replaced in runs (likely split across runs), replace entire text
            # This might reset some specific formatting but ensures correct text
            if not replaced_in_run:
                paragraph.text = paragraph.text.replace(placeholder, replacement)

    # 2. Search in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if placeholder in paragraph.text:
                        replaced_in_run = False
                        for run in paragraph.runs:
                            if placeholder in run.text:
                                run.text = run.text.replace(placeholder, replacement)
                                replaced_in_run = True
                        if not replaced_in_run:
                            paragraph.text = paragraph.text.replace(placeholder, replacement)

def generate_single_document(template_path, student_data, subjects):
    """
    Generates a single Word document containing all cover pages.
    Uses docxcompose to append documents, ensuring formatting is preserved.
    """
    master_doc = None
    composer = None

    for i, subject in enumerate(subjects):
        # 1. Create a fresh document from template for this subject
        # We need to reload the template file each time. 
        # Since template_path is a path string here (based on current logic), it works directly.
        temp_doc = Document(template_path)
        
        # 2. Prepare replacements
        replacements = {
            "{{Name}}": student_data.get("Name", ""),
            "{{Surname}}": student_data.get("Surname", ""),
            "{{Class}}": student_data.get("Class", ""),
            "{{Year}}": student_data.get("Year", ""),
            "{{Subject}}": subject
        }
        
        # 3. Perform replacements in temp_doc
        for placeholder, value in replacements.items():
            replace_placeholder(temp_doc, placeholder, value)
            
        if master_doc is None:
            # First subject sets the base for the master document
            master_doc = temp_doc
            composer = Composer(master_doc)
        else:
            # Add a clear page break to the master document before appending
            master_doc.add_page_break()
            # Append the new document
            composer.append(temp_doc)

    # Save to memory
    doc_buffer = io.BytesIO()
    if master_doc:
        master_doc.save(doc_buffer)
    doc_buffer.seek(0)
    return doc_buffer

# --- Main App Interface ---

def main():
    st.title("🎓 First Choice Student Cover Page Generator")
    st.markdown("""
    Generate personalized cover pages for students based on a Word Document template.
    """)

    # --- Sidebar Inputs ---
    with st.sidebar:
        st.header("Student Details")
        name = st.text_input("Name", placeholder="Enter First Name")
        surname = st.text_input("Surname", placeholder="Enter Surname")
        student_class = st.text_input("Class", placeholder="e.g., Grade 10A")
        
        current_year = datetime.now().year
        year = st.number_input("Year", min_value=2000, max_value=2100, value=current_year, step=1)

    # --- Main Area ---
    
    # 1. Template Verification
    st.subheader("1. School Standard Template")
    
    # Check for School Standard Template (template.docx)
    standard_template_path = "template.docx"
    has_standard_template = os.path.exists(standard_template_path)
    
    if has_standard_template:
        st.success("✅ School Standard Template detected and ready for use.")
        template_file = standard_template_path
    else:
        st.error("❌ CRITICAL ERROR: 'template.docx' (School Standard Template) not found!")
        st.warning("Please ensure the standard template file is in the application directory to proceed.")
        template_file = None

    # 2. Subject Selection
    st.subheader("2. Select Subjects")
    
    # Pre-defined list of subjects
    DEFAULT_SUBJECTS = [
        "Mathematics", "Science", "English", "History", "Geography", 
        "Art", "Physics", "Chemistry", "Biology", "Computer Science",
        "Physical Education", "Music", "Drama", "Economics"
    ]
    
    selected_subjects = st.multiselect("Choose subjects to generate covers for:", options=DEFAULT_SUBJECTS)
    
    # Allow custom subjects just in case
    custom_subjects_input = st.text_input("Add custom subjects (comma separated)", placeholder="e.g., Robotics, Latin")
    if custom_subjects_input:
        custom_list = [s.strip() for s in custom_subjects_input.split(",") if s.strip()]
        selected_subjects.extend(custom_list)
        # Remove duplicates
        selected_subjects = list(set(selected_subjects))

    # 3. Generate Button
    st.subheader("3. Generate")
    
    if st.button("Generate Cover Pages", type="primary"):
        # Validation
        if not template_file:
            st.error("❌ No template selected. Please upload a template or ensure 'template.docx' exists locally.")
            return
            
        if not name or not surname:
            st.warning("⚠️ Please fill in at least Name and Surname.")
            return
            
        if not selected_subjects:
            st.warning("⚠️ No subjects selected.")
            return

        with st.spinner("Processing documents..."):
            try:
                # Prepare data
                student_data = {
                    "Name": name,
                    "Surname": surname,
                    "Class": student_class,
                    "Year": year
                }
                
                # Generate Single Document
                doc_buffer = generate_single_document(template_file, student_data, selected_subjects)
                
                st.success(f"✅ Successfully generated {len(selected_subjects)} cover pages in a single file!")
                
                # Download Button
                clean_name = "".join([c for c in f"{name}_{surname}" if c.isalnum() or c in (' ', '_', '-')]).strip().replace(' ', '_')
                clean_class = "".join([c for c in student_class if c.isalnum() or c in (' ', '_', '-')]).strip().replace(' ', '_')
                file_name = f"{clean_name}_{clean_class}_Cover_Pages.docx"
                
                st.download_button(
                    label="📥 Download Cover Pages (.docx)",
                    data=doc_buffer,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except Exception as e:
                st.error(f"An error occurred during generation: {e}")
                import traceback
                st.write(traceback.format_exc())

if __name__ == "__main__":
    main()