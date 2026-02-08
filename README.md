# Student Cover Page Generator

A Streamlit application to generate personalized cover pages for students based on a Word Document template.

## Setup

1.  **Install Requirements:**
    ```bash
    pip install -r requirements.txt
    ```

2.  **Run the App:**
    ```bash
    streamlit run app.py
    ```

## Usage

1.  **Template:**
    - The app **requires** the **School Standard Template** (`template.docx`) to be present in this directory.
    - Custom template uploads are no longer supported to maintain brand consistency.
    - **Required Placeholders:** 
        - `{{Name}}` (Name)
        - `{{Surname}}` (Surname)
        - `{{Class}}` (Class)
        - `{{Year}}` (Year)
        - `{{Subject}}` (Subject)

2.  **Generate:**
    - Fill in student details.
    - Select subjects.
    - Click "Generate Cover Pages".
    - Download the resulting Word document (each subject on a new page).
