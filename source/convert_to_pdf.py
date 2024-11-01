import os
from docx2pdf import convert

#You will have to download the docx2pdf lib with this command in the cmd ( pip install python-docx docx2pdf  ) possibly as an administrator too

CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__)).replace("\\", "/")
OUTPUT_DIRECTORY = f"{CURRENT_DIRECTORY}/student_profiles_redacted"
PDF_OUTPUT_DIRECTORY = f"{CURRENT_DIRECTORY}/pdf_profiles"

if not os.path.exists(PDF_OUTPUT_DIRECTORY):
    os.mkdir(PDF_OUTPUT_DIRECTORY)

def convert_profiles_to_pdf():
    for filename in os.listdir(OUTPUT_DIRECTORY):
        if filename.endswith(".docx"):
            docx_path = os.path.join(OUTPUT_DIRECTORY, filename)
            pdf_filename = filename.replace(".docx", ".pdf")
            pdf_path = os.path.join(PDF_OUTPUT_DIRECTORY, pdf_filename)
            
            convert(docx_path, pdf_path)
            print(f"Converted {filename} to PDF.")

if __name__ == "__main__":
    convert_profiles_to_pdf()
