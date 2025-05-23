# Automated Email Migration Content Comparison

This project provides a Python script to automate the comparison of email content before and after a migration using the Gemini API. It leverages a provided script to extract content from `.msg` files and a Jupyter notebook's logic for the LLM comparison.

## Project Structure

*   `before_migration_emails/`: This directory should contain the original `.msg` email files before migration.
*   `after_migration_emails/`: This directory should contain the corresponding `.msg` email files after migration.
*   `process_email.py`: A Python script to parse `.msg` files and extract email subject, body, and attachments.
*   `compare_emails.py`: The main orchestration script that iterates through email files, extracts content, calls the Gemini API for comparison, and saves the results.
*   `Email_Content_Comparison_using_Gemini_API.ipynb`: The original Jupyter notebook outlining the LLM comparison logic.
*   `requirements.txt`: Lists the necessary Python dependencies.
*   `comparison_results/`: This directory will be created by the script to store the JSON comparison reports.
*   `extracted_excel_text/`: This directory will be created by the script to store the extracted Markdown text content from Excel attachments.

## File Pairing Convention

For the script to correctly compare emails, corresponding files in the `before_migration_emails` and `after_migration_emails` folders **must have the exact same filename**.

**Example:**

*   `before_migration_emails/Daily Sales Report.msg` will be paired with `after_migration_emails/Daily Sales Report.msg`.
*   `before_migration_emails/Meeting Minutes.msg` will be paired with `after_migration_emails/Meeting Minutes.msg`.

Ensure that for every file in `before_migration_emails`, there is a corresponding file with the identical name in `after_migration_emails`. The script will skip any files in `before_migration_emails` that do not have a match in `after_migration_emails`.

## How the Code Works

1.  **`process_email.py`**: This script contains the `parse_msg_file` function, which takes the path to a `.msg` file as input. It uses the `extract-msg` library to read the email data, including subject, sender, recipients, date, and body (handling both HTML and RTF). It also extracts Excel attachments and converts their content into structured plain text using Markdown tables. The body content is cleaned by stripping HTML tags using `BeautifulSoup`.
2.  **`compare_emails.py`**: This is the main script that orchestrates the comparison process.
    *   It defines the paths for the 'before' and 'after' email directories, an output directory for results, and a directory for saving extracted Excel text.
    *   It reads the `GEMINI_API_KEY` from environment variables for secure API access.
    *   It iterates through each `.msg` file found in the `before_migration_emails` directory.
    *   For each file, it constructs the expected path for the corresponding file in the `after_migration_emails` directory based on the filename convention.
    *   If the corresponding 'after' file exists, it calls `process_email.parse_msg_file` for both the 'before' and 'after' files to get their content, including the extracted Excel text.
    *   The extracted Excel text content for each attachment is then saved to a JSON file in the `extracted_excel_text` directory.
    *   It then calls the `compare_email_content` function, passing the extracted subject and body text.
    *   It also calls the `compare_excel_text` function, passing the paths to the saved extracted Excel text files.
    *   These comparison functions construct detailed prompts for the Gemini API, including the desired JSON output structure, and send them to the configured Gemini model (`gemini-2.0-flash` or `gemini-1.5-flash` for Excel).
    *   It parses the JSON responses from the API, which contain the comparison reports.
    *   The main script combines the text and Excel comparison reports and saves the final JSON report to a file in the `comparison_results` directory, named after the original email file.
    *   Error handling is included for file parsing issues, missing corresponding files, API call failures, and JSON parsing errors.

## Setup and Running

1.  **Setting up a Virtual Environment**

It is highly recommended to use a virtual environment to manage project dependencies and avoid conflicts with other Python projects on your system.

### Windows

1.  Open Command Prompt or PowerShell.
2.  Navigate to the project directory:
    ```bash
    cd path/to/your/project
    ```
3.  Create a virtual environment:
    ```bash
    python -m venv .venv
    ```
4.  Activate the virtual environment:
    ```bash
    .venv\Scripts\activate
    ```

### macOS and Linux

1.  Open your terminal.
2.  Navigate to the project directory:
    ```bash
    cd path/to/your/project
    ```
3.  Create a virtual environment:
    ```bash
    python3 -m venv .venv
    ```
4.  Activate the virtual environment:
    ```bash
    source .venv/bin/activate
    ```

Once the virtual environment is activated, your terminal prompt should show `(.venv)` or similar, indicating that you are working within the isolated environment.

2.  **Install dependencies**: Navigate to the project directory in your terminal and run:
    ```bash
    pip install -r requirements.txt
    ```
    This project requires `pandas` for Excel file processing, which is listed in `requirements.txt`.
3.  **Place your email files**:
    *   Put your 'before migration' `.msg` files into the `before_migration_emails` directory.
    *   Put the corresponding 'after migration' `.msg` files into the `after_migration_emails` directory, ensuring they have the **exact same filenames**.
4.  **Set your Gemini API Key**: Create a file named `.env` in the project root directory and add the following line, replacing `'your_key'` with your actual Gemini API key:
    ```
    GEMINI_API_KEY='your_key'
    ```
5.  **Run the script**: Execute the main script from the project directory:
    ```bash
    python compare_emails.py
    ```

## Results

The script will print progress to the console. Upon completion:
*   JSON files containing the extracted Markdown text content from Excel attachments will be saved in the `extracted_excel_text` directory.
*   JSON files containing the overall comparison reports for each email pair will be saved in the `comparison_results` directory.

You can inspect these JSON files to see the extracted Excel content and the detailed comparison status, including differences found in the subject, body text, and Excel tables (based on the Markdown text comparison), as reported by the Gemini API.

Feel free to modify the scripts to adapt them to your specific needs, such as changing the Gemini model, adjusting the prompt, or altering the output format.