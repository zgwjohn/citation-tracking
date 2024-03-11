import os
import pandas as pd
import refextract

# Set up pdftotext path
pdftotext_path = "\\poppler-23.11.0\\Library\\bin\\pdftotext.exe" # your location to the poppler library folder
os.environ['CFG_PATH_PDFTOTEXT'] = pdftotext_path
os.environ["PATH"] += os.pathsep + pdftotext_path

def sanitize_text(text):
    if not isinstance(text, str):
        return text
    return ''.join(char for char in text if char.isprintable())

def extract_references_from_folder(folder_path):
    all_references = []

    # Loop through each file in the folder
    for file in os.listdir(folder_path):
        if file.endswith('.pdf'):
            file_path = os.path.join(folder_path, file)

            # Extract references from each PDF
            try:
                references = refextract.extract_references_from_file(file_path)
                for reference in references:
                    # Construct a new dictionary with only linemarker, raw_ref, and source_file
                    cleaned_reference = {
                        'linemarker': sanitize_text(reference.get('linemarker')),
                        'raw_ref': sanitize_text(reference.get('raw_ref')),
                        'source_file': file
                    }
                    all_references.append(cleaned_reference)
            except Exception as e:
                print(f"Error extracting references from {file}: {e}")

    # Convert to DataFrame
    df = pd.DataFrame(all_references)

    # Save to Excel file
    output_file = os.path.join(folder_path, 'extracted_references.xlsx')
    try:
        df.to_excel(output_file, index=False)
    except Exception as e:
        print(f"Error writing to Excel: {e}")

    print(f"References extracted and saved to {output_file}")

# Output location
folder_path = '\\citationtracking\\'  # Replace with your folder path
extract_references_from_folder(folder_path)
