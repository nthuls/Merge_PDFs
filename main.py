import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tempfile import TemporaryDirectory
from pypdf import PdfReader, PdfWriter

# =========================
# ODT to PDF Converter
# =========================
def convert_odt_to_pdf(odt_file_path):
    if not os.path.exists(odt_file_path):
        messagebox.showerror("Error", f"The file {odt_file_path} does not exist.")
        return
    pdf_file_path = odt_file_path.replace(".odt", ".pdf")
    try:
        subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', odt_file_path],
            check=True
        )
        messagebox.showinfo("Success", f"Conversion successful! PDF saved as: {pdf_file_path}")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"An error occurred while converting the file: {e}")
    except Exception as ex:
        messagebox.showerror("Error", f"An unexpected error occurred: {ex}")

def select_odt_and_convert():
    odt_file_path = filedialog.askopenfilename(
        title="Select ODT file",
        filetypes=(("ODT files", "*.odt"), ("All files", "*.*"))
    )
    if odt_file_path:
        convert_odt_to_pdf(odt_file_path)


# =========================
# Image to High-Quality PDF Page
# =========================
def image_to_pdf_page_high_quality(image_path, target_width, target_height, temp_pdf_path, dpi=300):
    from PIL import Image
    with Image.open(image_path) as img:
        img = img.convert("RGB")
        # Adjust DPI
        img.info['dpi'] = (dpi, dpi)

        aspect_ratio = img.width / img.height
        scaled_width = target_width
        scaled_height = target_width / aspect_ratio

        if scaled_height > target_height:
            scaled_height = target_height
            scaled_width = target_height * aspect_ratio

        img = img.resize((int(scaled_width), int(scaled_height)), Image.LANCZOS)
        canvas = Image.new("RGB", (int(target_width), int(target_height)), (255, 255, 255))
        offset_x = (int(target_width) - img.width) // 2
        offset_y = (int(target_height) - img.height) // 2
        canvas.paste(img, (offset_x, offset_y))
        canvas.save(temp_pdf_path, "PDF", quality=95, optimize=True)

def select_image_and_convert_to_pdf():
    image_path = filedialog.askopenfilename(
        title="Select an image",
        filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.tif;*.bmp"), ("All files", "*.*")]
    )
    if not image_path:
        return
    width = simpledialog.askinteger("Page Width", "Enter page width in points (e.g. 595):", initialvalue=595)
    height = simpledialog.askinteger("Page Height", "Enter page height in points (e.g. 842):", initialvalue=842)
    if not width or not height:
        return
    output_pdf_path = filedialog.asksaveasfilename(
        title="Save PDF As",
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not output_pdf_path:
        return
    try:
        image_to_pdf_page_high_quality(image_path, width, height, output_pdf_path)
        messagebox.showinfo("Success", f"PDF created at {output_pdf_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert image to PDF: {e}")


# =========================
# PDF Resizing (Batch)
# =========================
def resize_pdf(input_pdf, output_pdf, target_size):
    import fitz
    target_width, target_height = target_size
    src = fitz.open(input_pdf)
    doc = fitz.open()

    for src_page in src:
        new_page = doc.new_page(width=target_width, height=target_height)
        rect = fitz.Rect(0, 0, target_width, target_height)
        new_page.show_pdf_page(rect, src, src_page.number)

    doc.save(output_pdf)
    doc.close()
    src.close()

def batch_resize_pdfs():
    width = simpledialog.askinteger("Page Width", "Enter target page width (points, e.g. 595):", initialvalue=595)
    height = simpledialog.askinteger("Page Height", "Enter target page height (points, e.g. 842):", initialvalue=842)
    if not width or not height:
        return

    input_files = filedialog.askopenfilenames(
        title="Select PDF files to resize",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not input_files:
        messagebox.showerror("Error", "No files selected.")
        return

    output_folder = filedialog.askdirectory(title="Select folder to save resized PDFs")
    if not output_folder:
        messagebox.showerror("Error", "No output folder selected.")
        return

    for input_file in input_files:
        try:
            output_file = os.path.join(output_folder, os.path.basename(input_file).replace(".pdf", "_resized.pdf"))
            resize_pdf(input_file, output_file, (width, height))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to resize {input_file}: {e}")

    messagebox.showinfo("Success", f"All files have been resized and saved to {output_folder}.")


# =========================
# PDF Trim White Spaces (Batch)
# =========================
def trim_whitespace(input_pdf, output_pdf, trim_top=190, trim_bottom=190):
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import RectangleObject
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        media_box = page.mediabox
        lower_left_x = media_box.lower_left[0]
        lower_left_y = media_box.lower_left[1] + trim_bottom
        upper_right_x = media_box.upper_right[0]
        upper_right_y = media_box.upper_right[1] - trim_top
        new_box = RectangleObject([lower_left_x, lower_left_y, upper_right_x, upper_right_y])
        page.mediabox = new_box
        page.cropbox = new_box
        writer.add_page(page)

    with open(output_pdf, "wb") as output_file:
        writer.write(output_file)

def batch_trim_whitespace():
    trim_top = simpledialog.askinteger("Trim Top", "Enter points to trim from top:", initialvalue=190)
    trim_bottom = simpledialog.askinteger("Trim Bottom", "Enter points to trim from bottom:", initialvalue=190)
    if trim_top is None or trim_bottom is None:
        return

    input_files = filedialog.askopenfilenames(
        title="Select PDF files to trim",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not input_files:
        messagebox.showerror("Error", "No files selected.")
        return

    output_folder = filedialog.askdirectory(title="Select folder to save trimmed PDFs")
    if not output_folder:
        messagebox.showerror("Error", "No output folder selected.")
        return

    for input_file in input_files:
        try:
            output_file = os.path.join(output_folder, os.path.basename(input_file).replace(".pdf", "_trimmed.pdf"))
            trim_whitespace(input_file, output_file, trim_top, trim_bottom)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to trim {input_file}: {e}")

    messagebox.showinfo("Success", f"All files have been trimmed and saved to {output_folder}.")


# =========================
# Remove Specific Pages & Resize
# =========================
def remove_and_resize_pages(input_pdf, output_pdf, remove_pages, resize_pages_to=None):
    from PyPDF2 import PdfReader, PdfWriter
    from PyPDF2.generic import RectangleObject

    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for page_num, page in enumerate(reader.pages):
        if page_num in remove_pages:
            continue
        if resize_pages_to:
            width, height = resize_pages_to
            page.mediabox = RectangleObject([0, 0, width, height])
        writer.add_page(page)

    with open(output_pdf, 'wb') as output_file:
        writer.write(output_file)

def remove_pages_from_pdf():
    input_pdf = filedialog.askopenfilename(
        title="Select PDF to remove pages from",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not input_pdf:
        return

    page_range_str = simpledialog.askstring("Pages to Remove", "Enter page range to remove (e.g. 9-25):")
    if not page_range_str:
        return

    try:
        start, end = page_range_str.split('-')
        start, end = int(start) - 1, int(end) - 1
        remove_pages_list = list(range(start, end + 1))
    except:
        messagebox.showerror("Error", "Invalid page range format.")
        return

    resize_answer = messagebox.askyesno("Resize?", "Do you want to resize the remaining pages to A4?")
    resize_pages_to = (595, 842) if resize_answer else None

    output_pdf = filedialog.asksaveasfilename(
        title="Save output PDF as",
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not output_pdf:
        return

    try:
        remove_and_resize_pages(input_pdf, output_pdf, remove_pages_list, resize_pages_to)
        messagebox.showinfo("Success", f"Pages removed and output saved to {output_pdf}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to remove pages: {e}")


# =========================
# Merge PDFs and Images with Consistent Page Layout
# =========================
def get_target_dimensions():
    # A4 size in points: ~8.27 x 11.69 inches at 72 dpi
    target_width = 595.276
    target_height = 841.890
    return target_width, target_height

def pdf_to_pdf_page(pdf_path, target_width, target_height, temp_pdf_path):
    import PyPDF2
    reader = PyPDF2.PdfReader(pdf_path)
    writer = PyPDF2.PdfWriter()

    for page in reader.pages:
        # Adjust mediabox
        page.mediabox.lower_left = (0, 0)
        page.mediabox.upper_right = (target_width, target_height)
        writer.add_page(page)

    with open(temp_pdf_path, "wb") as temp_pdf:
        writer.write(temp_pdf)

def merge_files(file_list, output_file):
    import PyPDF2
    if not file_list:
        messagebox.showerror("Error", "No files selected.")
        return

    target_width, target_height = get_target_dimensions()
    writer = PyPDF2.PdfWriter()

    with TemporaryDirectory() as tmpdir:
        for f in file_list:
            temp_pdf_path = os.path.join(tmpdir, "temp_page.pdf")
            lower_f = f.lower()

            if lower_f.endswith(".pdf"):
                try:
                    pdf_to_pdf_page(f, target_width, target_height, temp_pdf_path)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to process PDF {f}: {e}")
                    continue
            else:
                # Image file
                try:
                    image_to_pdf_page_high_quality(f, target_width, target_height, temp_pdf_path)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to process image {f}: {e}")
                    continue

            try:
                temp_reader = PyPDF2.PdfReader(temp_pdf_path)
                for page in temp_reader.pages:
                    writer.add_page(page)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add page from {temp_pdf_path}: {e}")

    try:
        with open(output_file, "wb") as final_pdf:
            writer.write(final_pdf)
        messagebox.showinfo("Success", f"Files merged successfully into {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to write merged PDF: {e}")

def open_merger_window():
    merger_window = tk.Toplevel(root)
    merger_window.title("PDF and Image Merger with Consistent Layout")

    pdf_listbox = tk.Listbox(merger_window, selectmode=tk.MULTIPLE, width=50, height=10)
    pdf_listbox.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

    def add_files():
        files = filedialog.askopenfilenames(
            filetypes=[("PDF and Image files", ("*.pdf", "*.png", "*.jpg", "*.jpeg", "*.tif", "*.bmp"))]
        )
        for file in files:
            pdf_listbox.insert(tk.END, file)

    def remove_selected():
        selected_items = pdf_listbox.curselection()
        for index in selected_items[::-1]:
            pdf_listbox.delete(index)

    def merge_selected_files():
        file_list = list(pdf_listbox.get(0, tk.END))
        if not file_list:
            messagebox.showerror("Error", "No files selected.")
            return
        output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if output_file:
            merge_files(file_list, output_file)

    def move_up():
        selected_items = pdf_listbox.curselection()
        for index in selected_items:
            if index > 0:
                item = pdf_listbox.get(index)
                pdf_listbox.delete(index)
                pdf_listbox.insert(index - 1, item)
                pdf_listbox.select_set(index - 1)

    def move_down():
        selected_items = pdf_listbox.curselection()
        for index in selected_items:
            if index < pdf_listbox.size() - 1:
                item = pdf_listbox.get(index)
                pdf_listbox.delete(index)
                pdf_listbox.insert(index + 1, item)
                pdf_listbox.select_set(index + 1)

    add_button = tk.Button(merger_window, text="Add PDFs/Images", command=add_files)
    add_button.grid(row=1, column=0, padx=10, pady=10)

    remove_button = tk.Button(merger_window, text="Remove Selected", command=remove_selected)
    remove_button.grid(row=1, column=1, padx=10, pady=10)

    merge_button = tk.Button(merger_window, text="Merge Files", command=merge_selected_files)
    merge_button.grid(row=1, column=2, padx=10, pady=10)

    move_up_button = tk.Button(merger_window, text="Move Up", command=move_up)
    move_up_button.grid(row=2, column=0, padx=10, pady=10)

    move_down_button = tk.Button(merger_window, text="Move Down", command=move_down)
    move_down_button.grid(row=2, column=1, padx=10, pady=10)


# =========================
# ROTATE PDFS
# =========================
def rotate_pdf(input_pdf, output_pdf, rotation_angle):
    """
    Rotate all pages in a PDF by the specified angle.

    Parameters:
    - input_pdf: Path to the input PDF file.
    - output_pdf: Path to save the rotated PDF.
    - rotation_angle: Angle to rotate pages (90 for right, -90 for left).
    """
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for page in reader.pages:
        # Rotate the page by the specified angle
        page.rotate(rotation_angle)
        writer.add_page(page)

    # Save the rotated PDF
    with open(output_pdf, "wb") as output_file:
        writer.write(output_file)

    print(f"Rotated PDF saved to {output_pdf}")
def rotate_pdf_gui():
    # Ask for input PDF
    input_pdf = filedialog.askopenfilename(
        title="Select a PDF file to rotate",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not input_pdf:
        messagebox.showerror("Error", "No file selected.")
        return

    # Ask for output PDF
    output_pdf = filedialog.asksaveasfilename(
        title="Save rotated PDF as",
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not output_pdf:
        messagebox.showerror("Error", "No output file specified.")
        return

    # Ask rotation direction
    rotation_choice = messagebox.askquestion("Rotate Direction", "Rotate right (Yes) or left (No)?")
    rotation_angle = 90 if rotation_choice == "yes" else -90

    # Perform rotation
    try:
        rotate_pdf(input_pdf, output_pdf, rotation_angle)
        messagebox.showinfo("Success", f"Rotated PDF saved to {output_pdf}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to rotate PDF: {e}")




# =========================
# Main GUI
# =========================
root = tk.Tk()
root.title("PDF & ODT Utility")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(fill="both", expand=True)

btn_odt_to_pdf = tk.Button(frame, text="ODT to PDF", command=select_odt_and_convert, width=40)
btn_odt_to_pdf.grid(row=0, column=0, pady=5)

btn_image_to_pdf = tk.Button(frame, text="Image to High-Quality PDF Page", command=select_image_and_convert_to_pdf, width=40)
btn_image_to_pdf.grid(row=1, column=0, pady=5)

btn_resize_pdf = tk.Button(frame, text="Batch Resize PDFs", command=batch_resize_pdfs, width=40)
btn_resize_pdf.grid(row=2, column=0, pady=5)

btn_trim_pdf = tk.Button(frame, text="Batch Trim White Spaces in PDFs", command=batch_trim_whitespace, width=40)
btn_trim_pdf.grid(row=3, column=0, pady=5)

btn_remove_pages = tk.Button(frame, text="Remove Pages & (Optionally) Resize PDF", command=remove_pages_from_pdf, width=40)
btn_remove_pages.grid(row=4, column=0, pady=5)

btn_merge = tk.Button(frame, text="Merge PDFs/Images with Consistent Layout", command=open_merger_window, width=40)
btn_merge.grid(row=5, column=0, pady=5)

btn_rotate_pdf = tk.Button(frame, text="Rotate PDF", command=rotate_pdf_gui, width=40)
btn_rotate_pdf.grid(row=6, column=0, pady=5)

root.mainloop()
