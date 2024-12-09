---
# PDF & Document Utility Toolkit

This toolkit provides a variety of functionalities for working with PDFs and image documents, as well as converting ODT files to PDFs. It features a Tkinter-based GUI to make these operations accessible to users without command-line experience.

## Features

1. **ODT to PDF Conversion**  
   Convert ODT documents to PDF using LibreOffice's headless mode.

2. **Image to High-Quality PDF**  
   Convert images (PNG, JPG, TIFF, etc.) into a high-quality, single-page PDF.  
   - The image is scaled proportionally to fit A4 or user-defined page sizes.  
   - White padding is added around the image if it doesn't fill the page.

3. **Batch PDF Resizing**  
   Resize multiple PDF files to standard page sizes (e.g., A4) or a custom dimension.

4. **Batch PDF Trimming**  
   Trim white spaces from the top and bottom of PDF pages to reduce unnecessary margins.

5. **Remove Specific Pages & (Optionally) Resize**  
   Remove a specified range of pages from a PDF, and optionally resize the remaining pages to a standard format.

6. **Merge PDFs and Images with Consistent Layout**  
   Merge multiple PDFs and images into a single PDF file, ensuring that all pages share a consistent layout and dimensions.

7. **Rotate PDF Pages**  
   Rotate all pages of a selected PDF either 90° clockwise or counterclockwise.

## Requirements

- **Python 3.6+**
- **Dependencies**:  
  - [LibreOffice](https://www.libreoffice.org/) (for ODT to PDF conversion)
  - `pip install Pillow pypdf PyMuPDF PyPDF2`
  
  Install dependencies:
  ```bash
  pip install Pillow pypdf PyMuPDF PyPDF2
  ```

- Make sure `libreoffice` is accessible in your system's PATH. If not, adjust your PATH or modify the script with the correct path to LibreOffice.

## Usage

1. **Clone or Download the Repository:**
   ```bash
   git clone https://github.com/yourusername/pdf-document-utility.git
   cd pdf-document-utility
   ```
   
2. **Run the Script:**
   ```bash
   python main.py
   ```
   This will launch a Tkinter GUI window.

3. **Main Window Options:**
   - **ODT to PDF:**  
     Opens a dialog to select an ODT file and convert it to PDF.

   - **Image to High-Quality PDF Page:**  
     Choose an image and specify page dimensions to convert it into a high-quality single-page PDF.

   - **Batch Resize PDFs:**  
     Select multiple PDF files and a target size to produce resized PDFs.

   - **Batch Trim White Spaces in PDFs:**  
     Choose multiple PDFs and specify how much to trim from top/bottom.

   - **Remove Pages & (Optionally) Resize PDF:**  
     Pick a PDF, specify a page range to remove, and optionally resize to A4.

   - **Merge PDFs/Images with Consistent Layout:**  
     Add multiple PDFs and images, reorder them, and merge into a single PDF with consistent page dimensions.

   - **Rotate PDF:**  
     Select a PDF and rotate all its pages either 90° clockwise or counterclockwise.

## Example Steps

- **To Convert ODT to PDF:**
  1. Click the "ODT to PDF" button.
  2. Select the ODT file.
  3. The converted PDF will be saved in the same directory as the ODT file.

- **To Merge PDFs and Images:**
  1. Click "Merge PDFs/Images with Consistent Layout."
  2. Add files by clicking "Add PDFs/Images."
  3. Use "Move Up" / "Move Down" to reorder.
  4. Click "Merge Files" and choose where to save the merged PDF.

- **To Rotate a PDF:**
  1. Click "Rotate PDF."
  2. Choose the input PDF.
  3. Select a name and location for the output file.
  4. When asked if you want to rotate right or left, click "Yes" for right (90°) or "No" for left (-90°).

## Customization

- **Page Sizes and DPI:**  
  Default page sizes are set to A4. You can modify these values in the code as needed.
  
- **Trimming White Spaces:**  
  The default trimming values can be changed by editing the `trim_top` and `trim_bottom` values.

- **ODT to PDF Conversion Path:**  
  If `libreoffice` is not on your PATH, edit the conversion code to include the full path to the `libreoffice` executable.

## Contributing

Feel free to open issues or submit pull requests. Contributions are welcome for:
- Adding support for more file types.
- Improving the user interface.
- Enhancing error handling.

## License

This project is licensed under the [MIT License](LICENSE).
