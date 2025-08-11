from docx import Document
import json
import os
JSON_FILE = "mapping_data.json"  # Mapping jason file helper
OUTPUT_FILE = "output.docx"      # This will be the generated output file
def load_mapping(json_path):
    """Load mapping from your JSON file- value → mnemonic."""
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    mapping = {}
    if isinstance(data, list):
        for item in data:
            if "value" in item and "mnemonic" in item:
                mapping[str(item["value"])] = str(item["mnemonic"])
    return mapping
def replace_text_in_paragraph(paragraph, mapping):
    """Replacing text in a paragraph preserving formatting."""
    if not paragraph.runs:
        return
    full_text = ''.join(run.text for run in paragraph.runs)
    changed = False
    for old_val, new_val in mapping.items():
        if old_val in full_text:
            full_text = full_text.replace(old_val, new_val)
            changed = True
    if changed:
        style = paragraph.runs[0].style
        paragraph.clear()
        run = paragraph.add_run(full_text)
        run.style = style
def replace_in_tables(table, mapping):
    for row in table.rows:
        for cell in row.cells:
            process_block(cell, mapping)
def process_block(container, mapping):
    """Processing paragraphs and tables."""
    for para in container.paragraphs:
        replace_text_in_paragraph(para, mapping)
    for tbl in getattr(container, 'tables', []):
        replace_in_tables(tbl, mapping)

def main():
    input_file = input("Enter the file name (with .docx extension) : ").strip()
    if not os.path.isfile(input_file):
        print(f"File '{input_file}' not found!")
        return
    if not os.path.isfile(JSON_FILE):
        print(f"Mapping JSON file '{JSON_FILE}' not found!")
        return
    mapping = load_mapping(JSON_FILE)
    if not mapping:
        print("Mapping is empty.")
        return
    doc = Document(input_file)
    process_block(doc, mapping)
    for section in doc.sections:
        process_block(section.header, mapping)
        process_block(section.footer, mapping)
    doc.save(OUTPUT_FILE)
    print(f"Replacements done! Output saved as '{OUTPUT_FILE}'")
if __name__ == "__main__":
    main()
