import os
import json
import sys
import tempfile
import shutil
from dotenv import load_dotenv # Import load_dotenv
# Add the directory containing process_email.py to the Python path
sys.path.append('.')
from process_email import parse_msg_file # Assuming process_email.py is in the same directory
# Remove the import of excel_to_pdf
# from convert_excel import excel_to_pdf # Import the conversion function

# Load environment variables from .env file
load_dotenv()

# Import and configure Gemini API client based on the notebook
from google import genai
from google.genai import types # Import types for file upload
import httpx # Import httpx for file upload

API_KEY = os.getenv("GEMINI_API_KEY")
if not API_KEY:
    print("Error: GEMINI_API_KEY environment variable not set. Create a .env file with GEMINI_API_KEY=YOUR_API_KEY")
    sys.exit(1)
client = genai.Client(api_key=API_KEY)

# Define directories
BEFORE_DIR = 'before_migration_emails'
AFTER_DIR = 'after_migration_emails'
OUTPUT_DIR = 'comparison_results'
ATTACHMENT_PDFS_DIR = 'attachment_pdfs' # New directory for converted PDFs
# New directory for extracted Excel text files
EXTRACTED_EXCEL_TEXT_DIR = 'extracted_excel_text'

# Ensure output and attachment PDF directories exist
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(ATTACHMENT_PDFS_DIR, exist_ok=True)
# Ensure the new directory exists
os.makedirs(EXTRACTED_EXCEL_TEXT_DIR, exist_ok=True)

def compare_email_content(old_content, new_content):
    """
    Constructs the prompt and calls the Gemini API for comparison of text content.
    """
    # Define the desired JSON structure for text comparison
    json_structure_definition = """
{
  "subject_comparison": {
    "status": "Match" | "Difference",
    "expected_subject": "...",
    "actual_subject": "...",
    "details": "..."
  },
  "body_text_comparison": {
    "status": "Match" | "Differences Found",
    "details": [
      {
        "type": "Missing Text" | "Extra Text" | "Text Mismatch" | "Formatting/Spacing Difference",
        "location": "e.g., paragraph number, line number",
        "expected_snippet": "...",
        "actual_snippet": "...",
        "description": "..."
      }
    ]
  }
}
"""

    prompt = f"""
You are a Quality Assurance (QA) assistant specializing in comparing email content after a migration.
Your task is to compare two versions of an email (an 'Expected' version and an 'Actual' version) and report any differences in the Subject and Body text. Do NOT compare tables in this step.

Focus on identifying even subtle differences, such as variations in wording, numbers, or spacing.

Provide your comparison report in the following JSON format:
{json_structure_definition}

Here are the two email versions:

---
Expected Email:
{old_content}
---

---
Actual Email:
{new_content}
---

Please analyze the content and generate the JSON comparison report. Ensure the JSON is valid and strictly follows the defined structure.
"""

    print("Sending text comparison prompt to Gemini API...")
    try:
        model = 'gemini-2.0-flash' # Or other suitable model
        response = client.models.generate_content(model=model, contents=prompt)
        print("Received response from API.")
        response_text = response.text.strip()

        json_start = response_text.find('{')
        json_end = response_text.rfind('}') + 1

        if json_start != -1 and json_end != -1:
            json_string = response_text[json_start:json_end]
            try:
                comparison_report = json.loads(json_string)
                print("Successfully parsed JSON response.")
                return comparison_report
            except json.JSONDecodeError as e:
                print(f"\nError decoding JSON from API response: {e}")
                print("Raw API response text:")
                print(response_text)
                return {"error": "JSON decoding failed", "details": str(e), "raw_response": response_text}
        else:
            print("\nCould not find valid JSON in the API response.")
            print("Raw API response text:")
            print(response_text)
            return {"error": "No JSON found in response", "raw_response": response_text}

    except Exception as e:
        print(f"\nAn error occurred during API call or processing: {e}")
        return {"error": "API call or processing failed", "details": str(e)}

# Remove the compare_excel_pdfs function

def compare_excel_text(before_text_path, after_text_path):
    """
    Compares the extracted Markdown text content of Excel attachments using Gemini API.
    Reads content from the provided file paths.
    """
    try:
        with open(before_text_path, 'r') as f:
            before_excel_text_data = json.load(f)
        with open(after_text_path, 'r') as f:
            after_excel_text_data = json.load(f)
    except FileNotFoundError as e:
        print(f"Error reading extracted Excel text files: {e}")
        return {"error": "Failed to read extracted Excel text files", "details": str(e)}
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON from extracted Excel text files: {e}")
        return {"error": "Failed to decode JSON from extracted Excel text files", "details": str(e)}
    except Exception as e:
        print(f"An unexpected error occurred while reading extracted Excel text files: {e}")
        return {"error": "Unexpected error reading extracted Excel text files", "details": str(e)}


    prompt = """
You are a Quality Assurance (QA) assistant specializing in comparing the content and structure of tables originating from Excel sheets, provided as Markdown text.
Your task is to compare two versions of extracted Excel content (an 'Expected' version and an 'Actual' version). The content is structured by sheet name, with each sheet containing one or more tables formatted in Markdown.

Analyze the content, structure, and formatting of the tables within this text.
Report any differences found in data values, missing or extra rows/columns, formatting inconsistencies, or structural variations.

Provide your comparison report for the tables in the following JSON format:
{
  "table_comparison": {
    "status": "Match" | "Differences Found",
    "details": [
      {
        "type": "Missing Sheet" | "Extra Sheet" | "Sheet Mismatch" | "Missing Table" | "Extra Table" | "Table Mismatch" | "Cell Value Mismatch" | "Formatting/Structure Difference" | "Other Difference",
        "location": "e.g., Sheet name, Table number within sheet, row number, column header",
        "expected_snippet": "...",
        "actual_snippet": "...",
        "description": "..."
      }
    ]
  }
}

Here is the extracted Excel content for comparison:

---
Expected Excel Content:
"""

    # Format the before_excel_text_data
    for sheet_name, tables in before_excel_text_data.items():
        prompt += f"\nSheet: {sheet_name}\n"
        for i, table in enumerate(tables):
            prompt += f"\nTable {i+1}:\n{table}\n"

    prompt += """
---

---
Actual Excel Content:
"""

    # Format the after_excel_text_data
    for sheet_name, tables in after_excel_text_data.items():
        prompt += f"\nSheet: {sheet_name}\n"
        for i, table in enumerate(tables):
            prompt += f"\nTable {i+1}:\n{table}\n"

    prompt += """
---

Analyze the provided Excel content (structured by sheet and Markdown tables) and generate the JSON comparison report for the tables. Ensure the JSON is valid and strictly follows the defined structure.
"""

    print("Sending Excel text comparison prompt to Gemini API...")
    try:
        # Use a model capable of handling longer contexts if needed, e.g., gemini-1.5-flash
        model = 'gemini-1.5-flash'
        response = client.models.generate_content(model=model, contents=prompt)
        print("Received response from API.")
        response_text = response.text.strip()

        json_start = response_text.find('{')
        json_end = response_text.rfind('}') + 1

        if json_start != -1 and json_end != -1:
            json_string = response_text[json_start:json_end]
            try:
                comparison_report = json.loads(json_string)
                print("Successfully parsed JSON response.")
                return comparison_report
            except json.JSONDecodeError as e:
                print(f"\nError decoding JSON from API response: {e}")
                print("Raw API response text:")
                print(response_text)
                return {"error": "JSON decoding failed", "details": str(e), "raw_response": response_text}
        else:
            print("\nCould not find valid JSON in the API response.")
            print("Raw API response text:")
            print(response_text)
            return {"error": "No JSON found in response", "raw_response": response_text}

    except Exception as e:
        print(f"\nAn error occurred during API call or processing for Excel text comparison: {e}")
        return {"error": "API call or processing failed", "details": str(e)}


def main():
    """
    Main function to iterate through emails and perform comparisons.
    """
    print(f"Processing emails in '{BEFORE_DIR}' and '{AFTER_DIR}'...")

    before_files = [f for f in os.listdir(BEFORE_DIR) if f.endswith('.msg')]

    for filename in before_files:
        before_path = os.path.join(BEFORE_DIR, filename)
        after_path = os.path.join(AFTER_DIR, filename) # Assuming same filename convention

        if not os.path.exists(after_path):
            print(f"Warning: Corresponding file not found in '{AFTER_DIR}' for '{filename}'. Skipping.")
            continue

        print(f"\nComparing email: {filename}")

        # No need for email-specific temp directory, PDFs go to ATTACHMENT_PDFS_DIR

        try:
            # Extract content from both emails
            before_email_data = parse_msg_file(before_path)
            after_email_data = parse_msg_file(after_path)

            # --- Step 1: Compare Subject and Body Text ---
            old_subject = before_email_data.get('Subject', '')
            old_body = before_email_data.get('Body', '')
            new_subject = after_email_data.get('Subject', '')
            new_body = after_email_data.get('Body', '')

            old_content = f"Subject: {old_subject}\n\nBody:\n{old_body}"
            new_content = f"Subject: {new_subject}\n\nBody:\n{new_body}"

            text_comparison_report = compare_email_content(old_content, new_content)

            # --- Step 2: Process and Compare Excel Attachments ---
            # --- Step 2: Process and Compare Excel Attachments ---
            excel_comparison_report = {"table_comparison": {"status": "No Attachments", "details": []}}
            before_excel_content = before_email_data.get('ExcelAttachmentContent', [])
            after_excel_content = after_email_data.get('ExcelAttachmentContent', [])

            if before_excel_content or after_excel_content:
                 excel_comparison_report["table_comparison"]["status"] = "Attachments Found"
                 excel_comparison_report["table_comparison"]["details"].append({"description": f"Found {len(before_excel_content)} Excel attachments in before email and {len(after_excel_content)} in after email."})

            # Compare the extracted Excel text content
            # A more robust solution would match attachments by filename before comparing content
            # For now, assuming attachments are in the same order
            min_attachments = min(len(before_excel_content), len(after_excel_content))

            for i in range(min_attachments):
                before_att_content = before_excel_content[i]
                after_att_content = after_excel_content[i]

                att_filename = before_att_content['filename'] # Use before filename as reference

                print(f"Processing attachment: {att_filename}")

                # Define paths for saving extracted text content
                # Use email filename as a subdirectory to avoid naming conflicts
                email_text_subdir = os.path.join(EXTRACTED_EXCEL_TEXT_DIR, os.path.splitext(filename)[0])
                os.makedirs(email_text_subdir, exist_ok=True) # Ensure subdirectory exists

                before_text_path = os.path.join(email_text_subdir, f"before_{att_filename}.json")
                after_text_path = os.path.join(email_text_subdir, f"after_{att_filename}.json")

                # Save the extracted Markdown content to JSON files
                try:
                    with open(before_text_path, 'w') as f:
                        json.dump(before_att_content['content'], f, indent=2)
                    print(f"Saved extracted text for '{att_filename}' (before) to '{before_text_path}'")
                except Exception as e:
                    print(f"Error saving extracted text for '{att_filename}' (before): {e}")
                    # Handle error, maybe skip comparison for this attachment

                try:
                    with open(after_text_path, 'w') as f:
                        json.dump(after_att_content['content'], f, indent=2)
                    print(f"Saved extracted text for '{att_filename}' (after) to '{after_text_path}'")
                except Exception as e:
                    print(f"Error saving extracted text for '{att_filename}' (after): {e}")
                    # Handle error, maybe skip comparison for this attachment

                # Call the new function to compare Excel text content, passing file paths
                # Need to modify compare_excel_text to accept paths and read content
                excel_text_comparison_result = compare_excel_text(before_text_path, after_text_path)

                # Append Excel text comparison details to the overall excel report
                if "table_comparison" in excel_text_comparison_result:
                     excel_comparison_report["table_comparison"]["details"].extend(excel_text_comparison_result["table_comparison"]["details"])
                     if excel_text_comparison_result["table_comparison"]["status"] == "Differences Found":
                         excel_comparison_report["table_comparison"]["status"] = "Differences Found"
                else:
                     excel_comparison_report["table_comparison"]["details"].append({"description": f"Could not get table comparison for {att_filename}", "details": excel_text_comparison_result})


            # --- Step 3: Combine Results and Save Report ---
            final_comparison_report = {
                "comparison_status": "Match", # Determine overall status later
                **text_comparison_report, # Include subject and body comparison
                **excel_comparison_report # Include excel/table comparison
            }

            # Determine overall comparison status
            if text_comparison_report.get("subject_comparison", {}).get("status") == "Difference" or \
               text_comparison_report.get("body_text_comparison", {}).get("status") == "Differences Found" or \
               excel_comparison_report.get("table_comparison", {}).get("status") == "Differences Found":
                final_comparison_report["comparison_status"] = "Differences Found"
            else:
                 final_comparison_report["comparison_status"] = "Match"


            # Add an overall summary
            final_comparison_report["overall_summary"] = f"Comparison of {filename} completed. Status: {final_comparison_report['comparison_status']}"


            # Save the final comparison report
            output_filename = os.path.splitext(filename)[0] + "_comparison.json"
            output_path = os.path.join(OUTPUT_DIR, output_filename)
            with open(output_path, 'w') as f:
                json.dump(final_comparison_report, f, indent=2)

            print(f"Final comparison report saved to '{output_path}'")

        except Exception as e:
            print(f"Error processing '{filename}': {e}")
        except Exception as e:
            print(f"Error processing '{filename}': {e}")
        # No finally block needed for temp dir cleanup as PDFs are persistent and Excel temps are deleted per attachment


    print("\nEmail comparison process finished.")
    # No main temporary directory cleanup needed as email-specific temp dirs are removed and Excel temps are deleted per attachment


if __name__ == "__main__":
    main()