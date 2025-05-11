import tkinter as tk
from tkinter import filedialog, messagebox
from pypdf import PdfReader, PdfWriter
import os

def merge_pdfs():
    files = filedialog.askopenfilenames(title="选择多个PDF文件合并", filetypes=[("PDF Files", "*.pdf")])
    if not files:
        return

    writer = PdfWriter()

    try:
        for file in files:
            reader = PdfReader(file)
            for page in reader.pages:
                writer.add_page(page)

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")], title="保存合并后的PDF")
        if output_path:
            with open(output_path, "wb") as f:
                writer.write(f)
            messagebox.showinfo("成功", f"PDF 合并完成：{output_path}")
    except Exception as e:
        messagebox.showerror("错误", str(e))

def split_pdf():
    file = filedialog.askopenfilename(title="选择一个PDF文件分割", filetypes=[("PDF Files", "*.pdf")])
    if not file:
        return

    try:
        reader = PdfReader(file)
        total_pages = len(reader.pages)

        # 获取用户输入页码范围
        page_range = page_entry.get().strip()
        if not page_range:
            messagebox.showwarning("提示", "请输入页码范围，如 1-3 或 1,3,5")
            return

        writer = PdfWriter()
        pages = []

        if "-" in page_range:
            start, end = map(int, page_range.split("-"))
            pages = list(range(start - 1, end))
        else:
            pages = [int(x.strip()) - 1 for x in page_range.split(",")]

        for i in pages:
            if 0 <= i < total_pages:
                writer.add_page(reader.pages[i])
            else:
                messagebox.showerror("错误", f"页码超出范围：{i + 1}")
                return

        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")], title="保存分割后的PDF")
        if output_path:
            with open(output_path, "wb") as f:
                writer.write(f)
            messagebox.showinfo("成功", f"PDF 分割完成：{output_path}")
    except Exception as e:
        messagebox.showerror("错误", str(e))

# 构建 GUI
root = tk.Tk()
root.title("PDF 合并与分割工具")
root.geometry("400x300")
root.resizable(False, False)

tk.Label(root, text="PDF 工具", font=("Arial", 16)).pack(pady=10)

tk.Button(root, text="📎 合并多个 PDF", font=("Arial", 12), command=merge_pdfs).pack(pady=10)

tk.Label(root, text="输入要提取的页码（如 1-3 或 1,3,5）:", font=("Arial", 10)).pack(pady=5)
page_entry = tk.Entry(root, font=("Arial", 12))
page_entry.pack(pady=5)

tk.Button(root, text="✂️ 分割指定页码", font=("Arial", 12), command=split_pdf).pack(pady=10)

root.mainloop()
