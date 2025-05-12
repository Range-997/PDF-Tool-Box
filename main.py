import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
from pypdf import PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import ImageTk, Image
from tkinter import ttk
from tkinter import simpledialog

dragged_files = []
merge_file_data = []

def on_drop(event):
    global dragged_files
    files = root.tk.splitlist(event.data)
    pdf_files = [f for f in files if f.lower().endswith(".pdf")]
    if not pdf_files:
        messagebox.showwarning("Warning", "Please drag in PDF files only.")
        return
    dragged_files = pdf_files
    status_label.config(text=f"Dragged files: {', '.join([os.path.basename(f) for f in pdf_files])}")

def merge_pdfs():
    if not merge_file_data:
        messagebox.showwarning("Warning", "No PDF files added for merging.")
        return

    writer = PdfWriter()
    try:
        for item in merge_file_data:
            path = item['path']
            page_range = item['pages']
            reader = PdfReader(path)
            total_pages = len(reader.pages)

            if page_range.lower() == "all":
                pages = range(total_pages)
            elif "-" in page_range:
                start, end = map(int, page_range.split("-"))
                pages = range(start - 1, end)
            else:
                pages = [int(x.strip()) - 1 for x in page_range.split(",")]

            for i in pages:
                if 0 <= i < total_pages:
                    writer.add_page(reader.pages[i])
                else:
                    messagebox.showerror("Error", f"Page out of range in file {os.path.basename(path)}: {i + 1}")
                    return

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")],
                                                   title="Save merged PDF")
        if output_path:
            with open(output_path, "wb") as f:
                writer.write(f)
            messagebox.showinfo("Success", f"PDF merged successfully: {output_path}")
            status_label.config(text="Merge completed.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# def split_pdf():
#     global dragged_files
#     file = dragged_files[0] if dragged_files else filedialog.askopenfilename(title="Select a PDF to split",
#                                                                               filetypes=[("PDF Files", "*.pdf")])
#     if not file:
#         return
#
#     try:
#         reader = PdfReader(file)
#         total_pages = len(reader.pages)
#
#         page_range = page_entry.get().strip()
#         if not page_range:
#             messagebox.showwarning("Hint", "Enter page range like 1-3 or 1,3,5.")
#             return
#
#         writer = PdfWriter()
#         if "-" in page_range:
#             start, end = map(int, page_range.split("-"))
#             pages = list(range(start - 1, end))
#         else:
#             pages = [int(x.strip()) - 1 for x in page_range.split(",")]
#
#         for i in pages:
#             if 0 <= i < total_pages:
#                 writer.add_page(reader.pages[i])
#             else:
#                 messagebox.showerror("Error", f"Page out of range: {i + 1}")
#                 return
#
#         output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")],
#                                                    title="Save split PDF")
#         if output_path:
#             with open(output_path, "wb") as f:
#                 writer.write(f)
#             messagebox.showinfo("Success", f"PDF split successfully: {output_path}")
#             status_label.config(text="Split completed.")
#     except Exception as e:
#         messagebox.showerror("Error", str(e))

def extract_text_from_pdf():
    global dragged_files
    file = dragged_files[0] if dragged_files else filedialog.askopenfilename(
        title="Select a PDF to extract text from",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not file:
        return

    try:
        reader = PdfReader(file)
        total_pages = len(reader.pages)

        page_range = page_entry.get().strip()
        text_output = ""

        # If no page range is entered, extract the entire document
        if not page_range:
            for i in range(total_pages):
                text = reader.pages[i].extract_text()
                if text:
                    text_output += f"\n--- Page {i+1} ---\n{text}"
        else:
            if "-" in page_range:
                start, end = map(int, page_range.split("-"))
                pages = list(range(start - 1, end))
            else:
                pages = [int(x.strip()) - 1 for x in page_range.split(",")]

            for i in pages:
                if 0 <= i < total_pages:
                    text = reader.pages[i].extract_text()
                    if text:
                        text_output += f"\n--- Page {i+1} ---\n{text}"
                else:
                    messagebox.showerror("Error", f"Page out of range: {i + 1}")
                    return

        if not text_output.strip():
            messagebox.showinfo("Info", "No text found on the selected pages.")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt")],
            title="Save extracted text"
        )
        if output_path:
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(text_output)
            messagebox.showinfo("Success", f"Text extracted successfully: {output_path}")
            status_label.config(text="Text extraction completed.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def add_files_to_merge():
    global merge_file_data
    files = filedialog.askopenfilenames(title="Select PDF files", filetypes=[("PDF Files", "*.pdf")])
    for file in files:
        filepath = file
        filename = os.path.basename(file)
        merge_file_data.append({'path': filepath, 'pages': 'all'})
        tree.insert("", "end", values=(filename, 'all'))

def on_range_edit(event):
    selected = tree.focus()
    if selected:
        item = tree.item(selected)
        col = tree.identify_column(event.x)
        if col == "#2":  # 2nd column = Page Range
            new_range = simpledialog.askstring("Page Range", "Enter page range (e.g., 1-3, 1,3 or 'all'):")
            if new_range:
                tree.set(selected, column="Page Range", value=new_range)
                idx = tree.index(selected)
                merge_file_data[idx]['pages'] = new_range

def move_up():
    selected = tree.focus()
    if not selected:
        return
    idx = tree.index(selected)
    if idx > 0:
        swap_items(idx, idx - 1)

def move_down():
    selected = tree.focus()
    if not selected:
        return
    idx = tree.index(selected)
    if idx < len(merge_file_data) - 1:
        swap_items(idx, idx + 1)

def swap_items(i, j):
    merge_file_data[i], merge_file_data[j] = merge_file_data[j], merge_file_data[i]
    tree.delete(*tree.get_children())
    for item in merge_file_data:
        tree.insert("", "end", values=(os.path.basename(item['path']), item['pages']))


# Main window using TkinterDnD
root = TkinterDnD.Tk()
root.title("PDF Toolbox")
root.geometry("520x600")
root.resizable(False, False)

tk.Label(root, text="PDF Toolbox", font=("Arial", 16)).pack(pady=10)

drop_area = tk.Label(root, text="Drag PDF files here", relief="groove", width=60, height=4, bg="#f0f0f0")
drop_area.pack(pady=8)
drop_area.drop_target_register(DND_FILES)
drop_area.dnd_bind('<<Drop>>', on_drop)


# merge
tk.Label(root, text="Custom File List", font=("Arial", 12)).pack(pady=5)
tree = ttk.Treeview(root, columns=("Filename", "Page Range"), show="headings", height=5)
tree.heading("Filename", text="Filename")
tree.heading("Page Range", text="Page Range")
tree.column("Filename", width=260)
tree.column("Page Range", width=120)
tree.pack(pady=5)
tree.bind("<Double-1>", on_range_edit)
frame = tk.Frame(root)
frame.pack()
tk.Button(frame, text="Add PDFs", command=add_files_to_merge).grid(row=0, column=0, padx=5)
tk.Button(frame, text="Move Up", command=move_up).grid(row=0, column=1, padx=5)
tk.Button(frame, text="Move Down", command=move_down).grid(row=0, column=2, padx=5)
tk.Button(frame, text="Merge PDFs (Custom)", command=merge_pdfs).grid(row=0, column=3, padx=5)

# tk.Button(root, text="Split Specific Pages", font=("Arial", 12), command=split_pdf).pack(pady=8)
# tk.Label(root, text="Enter page numbers (e.g. 1-3 or 1,3,5):", font=("Arial", 10)).pack(pady=5)
page_entry = tk.Entry(root, font=("Arial", 12))
page_entry.pack(pady=5)


tk.Button(root, text="Extract Text from PDF", font=("Arial", 12), command=extract_text_from_pdf).pack(pady=8)



status_label = tk.Label(root, text="", font=("Arial", 10), fg="green")
status_label.pack(pady=5)

root.mainloop()
