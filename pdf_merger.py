from tkinter import Tk
from tkinter.filedialog import askopenfilenames
from PyPDF2 import PdfMerger

# Create a Tkinter root window (hidden)
root = Tk()
root.withdraw()

# Prompt the user to select the PDF files
pdf_files = askopenfilenames(title="Select PDF files to merge", filetypes=[("PDF Files", "*.pdf")])

# Convert the selected files tuple to a list
pdfs = list(pdf_files)

# Check if any files were selected
if not pdfs:
    print("No files selected. Exiting.")
else:
    # Create the PdfMerger object
    merger = PdfMerger()

    # Append the PDFs to the merger
    for pdf in pdfs:
        merger.append(pdf)

    # Prompt the user for the output filename
    output_filename = input("Enter the filename for the merged PDF: ")

    # Write the merged PDF to a new file
    merger.write(output_filename)

    # Close the PdfMerger object
    merger.close()

    print(f"PDFs merged successfully into {output_filename}.")
