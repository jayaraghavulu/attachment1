# Email Migration Excel Attachment Comparison Plan

**Goal:** Replace the LibreOffice-based Excel-to-PDF conversion and subsequent PDF comparison with direct extraction of Excel content into structured plain text (Markdown tables) and comparison of this text by the LLM.

**Current Process:**

1.  Parsing `.msg` files using `process_email.py`.
2.  Extracting Excel attachments and saving them as temporary files.
3.  Converting these temporary Excel files to PDF using `convert_excel.py` (which uses LibreOffice).
4.  Comparing the generated PDFs using the Gemini API in `compare_emails.py`.

**Proposed Plan:**

Replace steps 3 and 4 for Excel attachments with:

1.  Extracting the content of the temporary Excel files into a structured plain text format (Markdown tables) within `process_email.py`.
2.  Comparing this plain text content directly using the Gemini API in `compare_emails.py`.

**Detailed Steps:**

1.  **Information Gathering:** Check the existing dependencies in `requirements.txt` to see if a library for reading Excel files (like `pandas` or `openpyxl`) is already included. If not, add `pandas`.
2.  **Modify `process_email.py`:**
    *   Add necessary imports for reading Excel files (e.g., `import pandas as pd`).
    *   Implement a new function, e.g., `extract_excel_tables_as_markdown(excel_path)`. This function will:
        *   Read the Excel file using `pandas`.
        *   Iterate through each sheet.
        *   For each sheet, identify distinct tables (contiguous blocks of data separated by empty rows/columns).
        *   Convert each identified table into a Markdown table string.
        *   Return a dictionary where keys are sheet names and values are lists of Markdown table strings found in that sheet.
    *   Update the `parse_msg_file` function:
        *   Modify the logic for handling `.xls` and `.xlsx` attachments.
        *   Instead of saving the raw attachment data to a temporary file and storing its path, call the new `extract_excel_tables_as_markdown` function with the path to the temporary Excel file.
        *   Store the structured Markdown table data returned by `extract_excel_tables_as_markdown` in the email data dictionary (e.g., under a new key like `"ExcelContentText"`).
        *   Ensure the temporary Excel file created from the attachment data is still cleaned up after its content is extracted.
3.  **Modify `compare_emails.py`:**
    *   Remove the import of `excel_to_pdf` from `convert_excel.py`.
    *   Remove the `compare_excel_pdfs` function entirely.
    *   Implement a new function, e.g., `compare_excel_text(before_excel_text_data, after_excel_text_data)`. This function will:
        *   Take the structured Excel text data (the dictionary of sheet names and Markdown tables) for both the 'before' and 'after' emails.
        *   Format this data into a clear prompt for the Gemini API, explaining the structure (sheet names, Markdown tables) and instructing the LLM to compare the content and structure of the tables.
        *   Define the desired JSON output structure for the table comparison report.
        *   Call the Gemini API with the prompt and receive the response.
        *   Parse the JSON response from the API.
        *   Return the parsed comparison report.
    *   Update the `main` function:
        *   Remove the calls to `excel_to_pdf`.
        *   Remove the calls to `compare_excel_pdfs`.
        *   Retrieve the extracted Excel text data using the new key (e.g., `"ExcelContentText"`) from the `before_email_data` and `after_email_data` dictionaries.
        *   Call the new `compare_excel_text` function with this data.
        *   Adjust the logic for incorporating the results from `compare_excel_text` into the `final_comparison_report`.
        *   Remove any remaining temporary file cleanup logic related to the PDF conversion process.
4.  **Update `README.md`:**
    *   Remove the section detailing the LibreOffice prerequisite and installation instructions.
    *   Update the "How the Code Works" section to accurately describe the new process for handling and comparing Excel attachments (extraction to Markdown text, direct LLM comparison).
    *   Update the "Setup and Running" section to remove any mention of LibreOffice and potentially add a note about the `pandas` dependency if it wasn't already listed.
5.  **Update `requirements.txt`:** Add `pandas` if it's not already present.

**Updated Process Flow:**

```mermaid
graph TD
    A[Start] --> B{Iterate .msg files in before_migration_emails};
    B --> C{Find corresponding file in after_migration_emails};
    C -- Found --> D[Parse before .msg file];
    C -- Not Found --> E[Skip file];
    D --> F[Parse after .msg file];
    F --> G[Extract Subject/Body Text];
    F --> H[Extract Excel Attachments as Markdown Text];
    G --> I[Compare Subject/Body Text with LLM];
    H --> J[Compare Excel Markdown Text with LLM];
    I --> K[Combine Comparison Results];
    J --> K;
    K --> L[Save Final Report to JSON];
    L --> B;
    B -- No more files --> M[End];