import PyPDF2
import tkinter as tk
from tkinter import filedialog, messagebox


def merge_pdfs(pdf_list, output_file):
    pdf_merger = PyPDF2.PdfMerger()
    for pdf in pdf_list:
        try:
            with open(pdf, 'rb') as pdf_file:
                pdf_merger.append(pdf_file)
        except FileNotFoundError:
            messagebox.showerror("Error", f"File {pdf} not found. Skipping.")

    with open(output_file, 'wb') as output_pdf:
        pdf_merger.write(output_pdf)
    messagebox.showinfo("Success", f"PDFs merged successfully into {output_file}")


def add_pdf():
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    for file in files:
        pdf_listbox.insert(tk.END, file)


def remove_pdf():
    selected_items = pdf_listbox.curselection()
    for index in selected_items[::-1]:
        pdf_listbox.delete(index)


def merge_selected_pdfs():
    pdf_files = list(pdf_listbox.get(0, tk.END))
    if not pdf_files:
        messagebox.showerror("Error", "No PDFs selected.")
        return
    output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if output_file:
        merge_pdfs(pdf_files, output_file)


def move_up():
    selected_items = pdf_listbox.curselection()
    for index in selected_items:
        if index > 0:
            pdf = pdf_listbox.get(index)
            pdf_listbox.delete(index)
            pdf_listbox.insert(index - 1, pdf)
            pdf_listbox.select_set(index - 1)


def move_down():
    selected_items = pdf_listbox.curselection()
    for index in selected_items:
        if index < pdf_listbox.size() - 1:
            pdf = pdf_listbox.get(index)
            pdf_listbox.delete(index)
            pdf_listbox.insert(index + 1, pdf)
            pdf_listbox.select_set(index + 1)


# GUI Setup
root = tk.Tk()
root.title("PDF Merger")

# Listbox to show added PDF files
pdf_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE, width=50, height=10)
pdf_listbox.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

# Buttons to add, remove, and merge PDFs
add_button = tk.Button(root, text="Add PDFs", command=add_pdf)
add_button.grid(row=1, column=0, padx=10, pady=10)

remove_button = tk.Button(root, text="Remove Selected", command=remove_pdf)
remove_button.grid(row=1, column=1, padx=10, pady=10)

merge_button = tk.Button(root, text="Merge PDFs", command=merge_selected_pdfs)
merge_button.grid(row=1, column=2, padx=10, pady=10)

# Buttons for moving items up and down
move_up_button = tk.Button(root, text="Move Up", command=move_up)
move_up_button.grid(row=2, column=0, padx=10, pady=10)

move_down_button = tk.Button(root, text="Move Down", command=move_down)
move_down_button.grid(row=2, column=1, padx=10, pady=10)

# Run the application
root.mainloop()
