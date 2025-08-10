**Overview**

This project processes an input .docx file, applies replacements using a JSON mapping file, and generates an updated .docx file as output.

**Files**

sample_input.docx — Example input file for testing.
mapping_data.json — JSON file containing the mapping rules for replacements.
output.docx — Generated file after processing the input.
main.py — Script that performs the processing.
requirements.txt — Python dependencies required to run the script.

**Steps to Run**

Place your input .docx file in the same directory as main.py.
Ensure mapping_data.json is in the same directory.
Install dependencies:
pip install -r requirements.txt
Run the script:
python main.py
When prompted, enter your input filename (e.g., sample_input.docx).
The processed file will be saved as output.docx in the same directory.
