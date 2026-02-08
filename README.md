# Student Cover Page Generator (First Choice)

A professional Streamlit application designed to batch-generate personalized student cover pages from a single Word document template.

## Use Case
If a student needs cover pages for 10 different subjects, instead of manual typing, this app generates a single 10-page Word document with all student details and subject titles correctly filled in, ready for printing.

## Setup Instructions

### 1. Requirements
Ensure you have Python installed, then install the necessary libraries:
```bash
pip install streamlit python-docx docxcompose
```

### 2. The Template (`template.docx`)
The application requires a file named `template.docx` in its root folder. This template acts as the design foundation. 

**Required Placeholders:**
Use these exact tags in your Word document (including brackets):
- `{{Name}}` - Inserts student's first name
- `{{Surname}}` - Inserts student's surname
- `{{Class}}` - Inserts student's class
- `{{Year}}` - Inserts academic year
- `{{Subject}}` - Inserts the subject title (one page per subject)

### 3. Running the App
Launch the application by running:
```bash
streamlit run app.py
```

## Features
- **Batch Processing:** Select multiple subjects to generate a single document.
- **Style Preservation:** Automatically detects and re-applies the font style and size from your template placeholders.
- **Print-Ready:** Uses section breaks and page breaks to ensure each cover page starts on a new sheet.
- **Auto-Naming:** The downloaded file is automatically named using the student's name and class.
