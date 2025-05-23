# Email Migration Comparison Plan

**Objective:** Modify the existing email processing and comparison scripts to convert Excel attachments to PDF, compare the PDFs using the Gemini API's document understanding, and include these results in the final JSON report.

**Plan:**

1.  **Modify `process_email.py`**:
    *   Update the `parse_msg_file()` function to iterate through attachments.
    *   Identify Excel files (`.xls`, `.xlsx`) based on their filename extension.
    *   Instead of reading them into pandas DataFrames, save the raw byte content of these Excel attachments to temporary files on the disk.
    *   Return the paths to these temporary Excel files in the dictionary, perhaps under a new key like `"ExcelAttachmentPaths"`.

2.  **Create a new module for conversion (`convert_excel.py`)**:
    *   Create a new Python file named `convert_excel.py`.
    *   Implement a function `excel_to_pdf(excel_path, pdf_path)` in this module.
    *   Inside `excel_to_pdf`, attempt to use `win32com` to open the Excel file at `excel_path` and save it as a PDF at `pdf_path`. This will involve interacting with the Microsoft Excel application object. Include error handling for potential issues (e.g., `win32com` not available, Excel not installed, conversion errors).
    *   If the `win32com` attempt fails or is not feasible (e.g., not on a Windows system), implement a fallback using `libreoffice`. This will involve executing a command-line command using Python's `subprocess` module to convert the Excel file to PDF via `libreoffice`. Ensure that `libreoffice` is installed and accessible from the system's PATH.

3.  **Modify `compare_emails.py`**:
    *   Import the `excel_to_pdf` function from the new `convert_excel.py` module.
    *   Update the `main()` function loop:
        *   After calling `parse_msg_file()` for both 'before' and 'after' emails, retrieve the list of temporary Excel file paths for attachments (using the new key added in step 1).
        *   Iterate through the list of Excel attachments. For each pair of corresponding 'before' and 'after' Excel files:
            *   Generate temporary PDF file paths (e.g., in the same temp directory).
            *   Call `convert_excel.excel_to_pdf` for both the 'before' and 'after' Excel files to create the temporary PDF files.
            *   Use the Gemini API (specifically the `client.files.upload` and `client.models.generate_content` methods as shown in the provided documentation excerpt for handling local files) to compare the content of the two generated PDF files. Craft a prompt that instructs Gemini to analyze the PDFs (which represent the Excel sheets) and report differences in data, formatting, and structure. The prompt should request the output in a structured format that can be parsed and integrated into the final JSON report's `table_comparison` section.
            *   Process the Gemini API response for the PDF comparison.
        *   Combine the results from the initial subject/body text comparison (using the existing `compare_email_content()` function) and the results from the PDF attachment comparison into a single dictionary that matches the desired JSON report structure.
        *   Save this combined dictionary as the final JSON report for the email pair.
        *   Add cleanup code to remove all temporary Excel and PDF files created during the process for the current email pair.

4.  **Refine JSON Report Structure**:
    *   Review and potentially adjust the `table_comparison` section of the JSON structure definition within the `compare_email_content()` function in `compare_emails.py`. The details provided by Gemini when comparing PDFs might be different from a text-based table comparison. Ensure the structure can accommodate the type of differences Gemini can identify (e.g., general formatting differences, data discrepancies within perceived tables).

5.  **Add Dependencies**:
    *   Update `requirements.txt` to include any new libraries needed (e.g., `pywin32` if using `win32com`).
    *   Add comments or documentation indicating the dependency on Microsoft Office (for `win32com`) or `libreoffice` for the conversion step.

Here is a Mermaid diagram illustrating the planned workflow:

```mermaid
graph TD
    A[Start] --> B{Process Email Pair};
    B --> C[Parse Before .msg];
    B --> D[Parse After .msg];
    C --> E[Extract Subject/Body];
    D --> F[Extract Subject/Body];
    C --> G[Extract Excel Attachments<br>Save to Temp .xlsx];
    D --> H[Extract Excel Attachments<br>Save to Temp .xlsx];
    E --> I{Compare Subject/Body<br>using Gemini};
    F --> I;
    G --> J[Convert Before .xlsx to Temp .pdf];
    H --> K[Convert After .xlsx to Temp .pdf];
    J --> L{Compare Before/After .pdf<br>using Gemini Document Understanding};
    K --> L;
    I --> M[Combine Comparison Results];
    L --> M;
    M --> N[Save Combined Report as JSON];
    N --> O[Cleanup Temp Files];
    O --> P[End];