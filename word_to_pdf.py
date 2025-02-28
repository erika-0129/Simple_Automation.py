import os
from docx2pdf import convert

def convert_word_to_pdf(input_folder, output_folder):
    """
    Converts all Word documents in the input folder to PDF files
    and saves them in the output folder.

    Args:
        input_folder (str): Path to the folder containing Word documents.
        output_folder (str): Path to the folder where PDF files will be saved.
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for filename in os.listdir(input_folder):
        if filename.endswith((".docx", ".doc")):
            input_path = os.path.join(input_folder, filename)
            output_path = os.path.join(output_folder, os.path.splitext(filename)[0] + ".pdf")
            convert(input_path, output_path)
            print(f"Converted '{filename}' to '{os.path.splitext(filename)[0] + '.pdf'}'")

if __name__ == "__main__":
    input_directory = r"path\name"  # Replace with the actual path
    output_directory = r"path\name"    # Replace with the desired output path
    convert_word_to_pdf(input_directory, output_directory)