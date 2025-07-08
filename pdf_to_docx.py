import os
import sys
from pdf2docx import Converter

# Configure where to search if filename is given
SEARCH_BASE_DIR = os.path.expanduser("~")  # Change to "~" to search entire home directory

def resolve_input_file_path(input_str):
    """
    Resolves user input to a valid PDF path by:
    - Stripping quotes/spaces
    - Adding .pdf if missing
    - Accepting full paths
    - Trying relative paths
    - Searching base directory if name only
    """
    cleaned_input = input_str.strip().strip('"').strip("'")

    # Append .pdf if user left it out
    if not cleaned_input.lower().endswith('.pdf'):
        cleaned_input += '.pdf'

    # 1. Absolute or relative path
    abs_path = os.path.abspath(cleaned_input)
    if os.path.isfile(abs_path):
        return abs_path

    # 2. Search fallback directory
    return find_pdf_by_name(os.path.basename(cleaned_input), SEARCH_BASE_DIR)

def find_pdf_by_name(filename, search_dir):
    """
    Recursively search for a file in the specified directory.
    """
    target_name = filename.lower()
    for root, _, files in os.walk(search_dir):
        for f in files:
            if f.lower() == target_name:
                return os.path.join(root, f)
    return None

def convert_pdf_to_docx(pdf_path):
    abs_pdf_path = os.path.abspath(pdf_path)
    docx_path = os.path.splitext(abs_pdf_path)[0] + ".docx"

    cv = Converter(abs_pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()

    print(f"✅ Conversion complete! Saved as: {docx_path}")

# Prompt user
input_file = input("Enter the PDF file name or path (you can drag and drop, or type just the name): ").strip()

if not input_file:
    print("❌ No input provided.")
    sys.exit(1)

file_path = resolve_input_file_path(input_file)

if not file_path:
    print(f"❌ File not found for input: '{input_file}'")
    sys.exit(1)

print(f"✅ File found: {file_path}")
convert_pdf_to_docx(file_path)