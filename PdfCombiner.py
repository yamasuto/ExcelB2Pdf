import sys
import os
import glob
from pypdf import PdfWriter

def main(source_pdf_folder_path, pdf_file_path):
    merger = PdfWriter()

    pdf_files = sorted(glob.glob(os.path.join(source_pdf_folder_path, "*.pdf")))
    for pdf in pdf_files:
        merger.append(pdf)
    
    merger.write(pdf_file_path)
    merger.close()
    print("Finished.")    

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python PdfCombiner.py <source_pdf_folder_path> (<pdf_file_path>)")
    if len(sys.argv) < 3:
        main(sys.argv[1], sys.argv[1]+".pdf")
    else:
        main(sys.argv[1], sys.argv[2])
