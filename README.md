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
    - Place a `template.docx` file in this directory.
    - OR use the upload feature in the app.
    - Ensure the template has placeholders: `{{Name}}`, `{{Surname}}`, `{{Class}}`, `{{Year}}`, `{{Subject}}`.

2.  **Generate:**
    - Fill in student details.
    - Select subjects.
    - Click "Generate Cover Pages".
    - Download the ZIP file.
