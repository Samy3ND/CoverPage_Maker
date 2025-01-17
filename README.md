# Lab Report Cover Page Generator

A Streamlit-based application to generate lab report cover pages for students of the 5th Semester at St. Xavier's College. This tool streamlines the process by filling in placeholders in a Word document template with student and subject-specific information.

---

## Features

- Auto-fills student name based on roll number.
- Selects subject and assigns the respective teacher automatically.
- Generates a lab report cover page as a `.docx` file.
- Downloads the generated cover page directly.

---

## Requirements

To run this application, ensure you have the following installed:

- Python 3.7 or above
- Required Python libraries listed in `requirements.txt`

Install dependencies using:

```bash
pip install -r requirements.txt
```
## Usage

 1. Clone this repository

    ```bash
     git clone https://github.com/Samy3ND/CoverPage_Maker.git
     cd CoverPage_Maker
     ```

 2. Place the Font.docx file (template for the cover page) in the root directory.

 3. Run the application
     ```bash
     streamlit run app.py
     ```

 4. Enter your roll number, subject, and lab report number. The name and teacher will auto-populate.

 5. Click "Generate Cover Page" to create the .docx file.

 6. Download the generated file using the "Download Cover Page" button.

---

## File Structure

 - Font_Page.py: The main application script.
 - Font.docx: The Word template for the cover page (must include placeholders like {Name}, {RollNumber}, etc.).
 - requirements.txt: Lists the dependencies required for the application.
   
---

## Dependencies

 - streamlit: For creating the interactive UI.
 - python-docx: For manipulating Word documents.
 - datetime: For fetching the current date.
 - io: For managing in-memory file operations.

---
## Screenshot
![Screenshot 2025-01-17 195745](https://github.com/user-attachments/assets/1cf3ffa3-c032-49c9-a39b-d7bd2f0bb49d)

---
## Feedback
Feel free to open an issue or suggest improvements. Contributions are welcome!
