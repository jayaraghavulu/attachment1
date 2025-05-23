import os
import subprocess
import sys
import tempfile

def excel_to_pdf(excel_path, pdf_path):
    """
    Converts an Excel file to PDF using libreoffice.
    Requires LibreOffice to be installed and in the system's PATH.
    """
    try:
        # Ensure the output directory exists
        pdf_dir = os.path.dirname(pdf_path)
        os.makedirs(pdf_dir, exist_ok=True)

        # libreoffice command to convert
        # -headless: don't open the LibreOffice GUI
        # -convert-to pdf: specify the output format
        # --outdir: specify the output directory
        command = [
            "soffice",
            "--headless",
            "--convert-to",
            "pdf:calc_pdf_Export:Scale=1",
            "--outdir",
            pdf_dir,
            os.path.abspath(excel_path)
        ]
        print(f"Attempting conversion using libreoffice: {' '.join(command)}")
        result = subprocess.run(command, capture_output=True, text=True, check=True)
        print("Libreoffice stdout:", result.stdout)
        print("Libreoffice stderr:", result.stderr)

        # Libreoffice creates the PDF in the output directory with the same base name
        expected_pdf_filename = os.path.splitext(os.path.basename(excel_path))[0] + ".pdf"
        actual_pdf_path_from_libreoffice = os.path.join(pdf_dir, expected_pdf_filename)

        # Check if the file was created and rename it if necessary
        if os.path.exists(actual_pdf_path_from_libreoffice):
             # If the target pdf_path is different from libreoffice's default output path, rename it
            if os.path.abspath(actual_pdf_path_from_libreoffice) != os.path.abspath(pdf_path):
                os.replace(actual_pdf_path_from_libreoffice, pdf_path)
            print(f"Successfully converted '{excel_path}' to '{pdf_path}' using libreoffice.")
            return True
        else:
            print(f"Libreoffice conversion failed: Expected output file '{actual_pdf_path_from_libreoffice}' not found.")
            return False

    except FileNotFoundError:
        print(f"Error: Command not found. Please ensure LibreOffice is installed and in your system's PATH. Attempted command: {command}")
        return False
    except subprocess.CalledProcessError as e:
        print(f"Error during libreoffice conversion of '{excel_path}': {e}")
        print("Libreoffice stdout:", e.stdout)
        print("Libreoffice stderr:", e.stderr)
        return False
    except Exception as e:
        print(f"An unexpected error occurred during libreoffice conversion of '{excel_path}': {e}")
        return False

if __name__ == '__main__':
    # Example Usage (requires test.xlsx and LibreOffice installed)
    # Create a dummy Excel file for testing
    try:
        import pandas as pd
        dummy_df = pd.DataFrame({'Col A': [1, 2], 'Col B': ['X', 'Y']})
        dummy_excel_path = "test_excel.xlsx"
        dummy_pdf_path = "test_excel.pdf"
        dummy_df.to_excel(dummy_excel_path, index=False)

        print(f"Created dummy Excel file: {dummy_excel_path}")

        if excel_to_pdf(dummy_excel_path, dummy_pdf_path):
            print(f"Conversion successful. PDF saved to {dummy_pdf_path}")
        else:
            print("Conversion failed.")

        # Clean up dummy files
        # os.remove(dummy_excel_path)
        # if os.path.exists(dummy_pdf_path):
        #     os.remove(dummy_pdf_path)

    except ImportError:
        print("pandas not found. Cannot create dummy Excel file for testing.")
    except Exception as e:
        print(f"An error occurred during example usage: {e}")