import pandas as pd # Import pandas
import extract_msg
import io
import os
import tempfile
from bs4 import BeautifulSoup

def extract_clean_body(raw_body):
    """
    Convert HTML or RTF email body to plain text, preserving table structure.
    Handles bytes or string input.
    """
    if isinstance(raw_body, bytes):
        try:
            # Most Outlook HTML bodies are ISO-8859-1 encoded
            html = raw_body.decode('iso-8859-1', errors='ignore')
        except Exception as e:
            print("Decoding failed:", e)
            return None
    else:
        html = raw_body

    soup = BeautifulSoup(html, 'html.parser')

    # Find all tables and replace them with their text representation
    for table in soup.find_all('table'):
        table_text = format_table_as_text(table)
        # Create a new tag to hold the text representation
        text_node = soup.new_tag("pre") # Use pre to preserve formatting
        text_node.string = table_text
        table.replace_with(text_node)

    # Get the text of the modified soup, using newline as separator
    clean_text = soup.get_text(separator="\n", strip=True)

    return clean_text

def format_table_as_text(table_soup):
    """
    Formats a BeautifulSoup table element into a text-based representation.
    """
    rows = table_soup.find_all('tr')
    if not rows:
        return ""

    # Determine column widths
    col_widths = []
    for row in rows:
        cells = row.find_all(['td', 'th'])
        for i, cell in enumerate(cells):
            cell_text = cell.get_text(strip=True)
            if i >= len(col_widths):
                col_widths.append(len(cell_text))
            else:
                col_widths[i] = max(col_widths[i], len(cell_text))

    # Build the text table
    table_text = ""
    # Add top border
    separator_line = "+"
    for width in col_widths:
        separator_line += "-" * (width + 2) + "+" # +2 for padding spaces
    table_text += separator_line + "\n"

    for row in rows:
        cells = row.find_all(['td', 'th'])
        row_text = "|"
        for i, cell in enumerate(cells):
            cell_text = cell.get_text(strip=True)
            # Pad cell text
            padded_text = cell_text.ljust(col_widths[i])
            row_text += f" {padded_text} |"
        table_text += row_text + "\n"

        # Add separator line after each row
        separator_line = "+"
        for width in col_widths:
            separator_line += "-" * (width + 2) + "+" # +2 for padding spaces
        table_text += separator_line + "\n"

    return table_text

# Helper function to detect tables in a DataFrame sheet
def find_tables_in_sheet(df):
    """
    Identifies distinct tables within a single pandas DataFrame sheet.
    Assumes tables are separated by at least one empty row and column.
    Returns a list of DataFrames, each representing a detected table.
    """
    tables = []
    rows, cols = df.shape
    if rows == 0 or cols == 0:
        return tables

    # Identify blocks of non-null data
    # Create a boolean mask where True indicates a non-null cell
    mask = df.notna()

    # Find contiguous blocks of True values
    # This is a simplified approach; a more robust solution might use clustering or graph analysis
    # For now, let's assume a table is a rectangular block of non-nulls
    # We'll iterate through the mask to find the top-left corner of potential tables
    visited = set()

    for r in range(rows):
        for c in range(cols):
            if mask.iloc[r, c] and (r, c) not in visited:
                # Found a potential top-left corner of a table
                # Determine the bounds of this table
                r_end = r
                while r_end < rows and mask.iloc[r_end, c]:
                    r_end += 1
                r_end -= 1 # Adjust back to the last non-null row

                c_end = c
                while c_end < cols and mask.iloc[r, c_end]:
                    c_end += 1
                c_end -= 1 # Adjust back to the last non-null column

                # Check if this block is a valid table (more than just one cell)
                if r_end >= r and c_end >= c:
                     # Extract the potential table slice
                    table_slice = df.iloc[r:r_end+1, c:c_end+1]

                    # Simple check: ensure the entire slice is mostly non-null
                    # A more sophisticated check would verify connectivity
                    if table_slice.notna().sum().sum() > (r_end - r + 1) * (c_end - c + 1) * 0.5: # e.g., > 50% non-null
                         tables.append(table_slice)

                         # Mark all cells in this table as visited
                         for vr in range(r, r_end + 1):
                             for vc in range(c, c_end + 1):
                                 visited.add((vr, vc))

    return tables

def extract_excel_tables_as_markdown(excel_path):
    """
    Reads an Excel file, identifies tables in each sheet, and converts them to Markdown.
    Returns a dictionary where keys are sheet names and values are lists of Markdown table strings.
    """
    excel_content = {}
    try:
        with pd.ExcelFile(excel_path) as xls: # Use 'with' statement to ensure file is closed
            for sheet_name in xls.sheet_names:
                df = xls.parse(sheet_name)
                tables_in_sheet = find_tables_in_sheet(df)
                markdown_tables = [table.to_markdown(index=False) for table in tables_in_sheet]
                if markdown_tables:
                    excel_content[sheet_name] = markdown_tables
    except Exception as e:
        print(f"Error processing Excel file {excel_path}: {e}")
        # Return an empty dictionary or an error indicator if processing fails
        return {"error": f"Failed to process Excel file: {e}"}

    return excel_content


def parse_msg_file(msg_path):
    """
    Parses a .msg Outlook file and returns its content as a dictionary.
    Extracts subject, sender, recipients, body, and Excel attachments.
    """
    msg = extract_msg.Message(msg_path)

    # Extract metadata
    subject = msg.subject
    sender = msg.sender
    to = msg.to
    cc = msg.cc
    date = msg.date

    # Extract and clean body
    raw_body = msg.htmlBody or msg.rtfBody or msg.body
    clean_body = extract_clean_body(raw_body)

    # Extract Excel attachments and process their content
    excel_attachment_content = []
    # print(f"DEBUG: Found {len(msg.attachments)} attachments.") # Removed debug
    for i, att in enumerate(msg.attachments):
        # print(f"DEBUG: Processing attachment {i}: {att}") # Removed debug
        filename = att.longFilename or att.shortFilename
        # print(f"DEBUG: Attachment {i} filename: {filename}") # Removed debug

        # Remove potential trailing null bytes from the filename
        if filename:
            filename = filename.strip('\x00')
            # print(f"DEBUG: Cleaned Attachment {i} filename: '{filename}'") # Removed debug

        try:
            data_size = len(att.data) if att.data else 0
            # print(f"DEBUG: Attachment {i} data size: {data_size} bytes") # Removed debug
        except Exception as data_e:
            # print(f"DEBUG: Could not get data size for attachment {i}: {data_e}") # Removed debug
            data_size = 0 # Assume 0 if data access fails

        # print(f"DEBUG: Attachment {i} filename for check: '{filename}'") # Removed debug
        # print(f"DEBUG: Attachment {i} filename repr: {repr(filename)}") # Removed debug
        # print(f"DEBUG: Attachment {i} filename length: {len(filename) if filename else 0}") # Removed debug
        # print(f"DEBUG: Attachment {i} filename lower: '{filename.lower() if filename else None}'") # Removed debug
        # print(f"DEBUG: Attachment {i} endswith check: {filename.lower().endswith(('.xls', '.xlsx')) if filename else False}") # Removed debug

        if filename and filename.lower().endswith(('.xls', '.xlsx')):
            if data_size > 0:
                try:
                    file_data = att.data  # Raw bytes
                    # Create a temporary file to save the Excel data
                    # Use a suffix to retain the original file extension
                    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(filename)[1])
                    temp_file.write(file_data)
                    temp_file.close()

                    print(f"DEBUG: Attempting to extract Excel content from {temp_file.name}") # Log to confirm execution reaches here
                    # Extract content as Markdown tables
                    excel_markdown_content = extract_excel_tables_as_markdown(temp_file.name)

                    # Append the extracted content to the list
                    excel_attachment_content.append({
                        "filename": filename,
                        "content": excel_markdown_content
                    })

                    print(f"DEBUG: Successfully processed Excel attachment '{filename}'.")

                    # Clean up the temporary Excel file immediately after processing
                    try:
                        os.remove(temp_file.name)
                        print(f"DEBUG: Cleaned up temporary Excel file: {temp_file.name}")
                    except OSError as e:
                        print(f"DEBUG: Error removing temporary Excel file {temp_file.name}: {e}")

                except Exception as e:
                    print(f"DEBUG: Failed to process Excel attachment '{filename}': {e}")
            else:
                 # print(f"DEBUG: Skipping Excel attachment '{filename}' with no data.") # Removed debug
                 pass # Keep silent if no data
        else:
            # print(f"DEBUG: Skipping non-Excel attachment or attachment with no filename: {filename}") # Removed debug
            pass # Keep silent for non-Excel or no filename


    return {
        "Subject": subject,
        "Sender": sender,
        "To": to,
        "CC": cc,
        "Date": date,
        "Body": clean_body,
        "ExcelAttachmentContent": excel_attachment_content # Changed key and content here
    }

# Example usage
if __name__ == "__main__":
    # Test parsing a specific file reported to have issues with attachments
    msg_file_to_test = r"after_migration_emails\Example with Attachment.msg"
    print(f"Testing parsing of: {msg_file_to_test}")
    email_data = parse_msg_file(msg_file_to_test)

    print("Subject:", email_data["Subject"])
    print("From:", email_data["Sender"])
    print("To:", email_data["To"])
    print("Date:", email_data["Date"])
    print("Body:\n", email_data["Body"])

    # Example usage for new structure
    if email_data.get("ExcelAttachmentContent"):
        print("\nDetected Excel Attachments:")
        for attachment_info in email_data["ExcelAttachmentContent"]:
            print(f"- Filename: {attachment_info['filename']}")
            print(f"  Extracted Content:")
            # Print extracted content in a readable format
            for sheet_name, tables in attachment_info['content'].items():
                print(f"  Sheet: {sheet_name}")
                for i, table in enumerate(tables):
                    print(f"    Table {i+1}:\n{table}")
    else:
        print("\nNo Excel attachments detected.")
