import streamlit as st
import os
import io
import zipfile
from docx import Document
from docx.shared import Pt
from datetime import datetime

# --- Configuration & Setup ---
st.set_page_config(page_title="Student Cover Page Generator", layout="wide")

# --- Helper Functions ---

def replace_placeholder(doc, placeholder, replacement):
    """
    Replaces a placeholder in a python-docx Document object.
    It searches in paragraphs and tables and applies specific styling:
    Cambria (Body), Size 20.
    """
    # 1. Search in paragraphs
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            for run in paragraph.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, str(replacement))
                    run.font.name = 'Cambria'
                    run.font.size = Pt(20)

    # 2. Search in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if placeholder in paragraph.text:
                        for run in paragraph.runs:
                            if placeholder in run.text:
                                run.text = run.text.replace(placeholder, str(replacement))
                                run.font.name = 'Cambria'
                                run.font.size = Pt(20)

def generate_cover_pages(template_file, student_data, subjects):
    """
    Generates a zip file containing a cover page for each subject.
    
    Args:
        template_file: The template file (path or file-like object).
        student_data (dict): Dictionary containing Name, Surname, Class, Year.
        subjects (list): List of selected subjects.
        
    Returns:
        BytesIO: A bytes buffer containing the zip file.
    """
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for subject in subjects:
            # Load the template FRESH for each subject to ensure clean slate
            # If template_file is a file-like object (UploadedFile), we must seek(0)
            if hasattr(template_file, 'seek'):
                template_file.seek(0)
            
            doc = Document(template_file)
            
            # Prepare replacements
            replacements = {
                "{{Name}}": student_data.get("Name", ""),
                "{{Surname}}": student_data.get("Surname", ""),
                "{{Class}}": student_data.get("Class", ""),
                "{{Year}}": student_data.get("Year", ""),
                "{{Subject}}": subject
            }
            
            # Perform replacements
            for placeholder, value in replacements.items():
                replace_placeholder(doc, placeholder, value)
            
            # Save the modified document to memory
            doc_buffer = io.BytesIO()
            doc.save(doc_buffer)
            doc_buffer.seek(0)
            
            # Create a filename
            safe_name = "".join([c for c in student_data.get("Name", "") if c.isalnum() or c in (' ', '_')]).strip()
            safe_subject = "".join([c for c in subject if c.isalnum() or c in (' ', '_')]).strip()
            filename = f"{safe_name}_{safe_subject}_Cover.docx"
            
            # Add to zip
            zip_file.writestr(filename, doc_buffer.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

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
    
    # 1. Template Selection
    st.subheader("1. Template Selection")
    
    # Check for School Standard Template (template.docx)
    standard_template_path = "template.docx"
    has_standard_template = os.path.exists(standard_template_path)
    
    template_file = None
    
    if has_standard_template:
        st.success("✅ School Standard Template detected.")
        use_standard = st.checkbox("Use School Standard Template?", value=True)
        if use_standard:
            template_file = standard_template_path
    
    if not template_file:
        with st.expander("ℹ️ Custom Template Requirements"):
            st.markdown("""
            If you upload your own template, ensure it contains the following placeholders exactly as shown (including double curly braces):
            - `{{Name}}`: Student's First Name
            - `{{Surname}}`: Student's Surname
            - `{{Class}}`: Student's Class/Grade
            - `{{Year}}`: Academic Year
            - `{{Subject}}`: The Name of the Subject (a separate file is created for each subject)
            """)
        template_file = st.file_uploader("Upload your Local/Custom Template (.docx)", type=["docx"])

    if template_file:
        st.info("Template loaded successfully.")
    else:
        st.warning("Please verify a template is available.")

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
                
                # Generate Zip
                zip_buffer = generate_cover_pages(template_file, student_data, selected_subjects)
                
                st.success(f"✅ Successfully generated {len(selected_subjects)} cover pages!")
                
                # Download Button
                st.download_button(
                    label="📥 Download Cover Pages (.zip)",
                    data=zip_buffer,
                    file_name="Cover_Pages.zip",
                    mime="application/zip"
                )
                
            except Exception as e:
                st.error(f"An error occurred during generation: {e}")
                import traceback
                st.write(traceback.format_exc())

if __name__ == "__main__":
    main()