import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import comtypes.client
from pypdf import PdfWriter, PdfReader
from PIL import Image, ImageTk
import os
import sys

def convert_docx_to_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if file_path:
        output_path = file_path.replace('.docx', '.pdf')
        try:
            word = comtypes.client.CreateObject('Word.Application')
            doc = word.Documents.Open(file_path)
            doc.SaveAs(output_path, FileFormat=17)  # 17 is PDF format
            doc.Close()
            word.Quit()
            if os.path.exists(output_path):
                messagebox.showinfo("Success", f"Converted to {output_path}")
            else:
                messagebox.showerror("Error", "Conversion failed: Output file not created")
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")

def preview_merge():
    """Show preview of PDFs to be merged"""
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if not files:
        return

    # Create preview window
    preview_window = tk.Toplevel(root)
    preview_window.title("Merge Preview")
    preview_window.geometry("500x400")
    preview_window.configure(bg='#ffffff')
    preview_window.resizable(False, False)
    preview_window.transient(root)
    preview_window.grab_set()

    # Preview title
    title_label = tk.Label(preview_window, text="PDF Merge Preview", font=("Segoe UI", 16, "bold"),
                          bg='#ffffff', fg='#2d3748')
    title_label.pack(pady=(20, 10))

    # Frame for PDF list
    list_frame = tk.Frame(preview_window, bg='#ffffff')
    list_frame.pack(fill='both', expand=True, padx=20, pady=(0, 20))

    # Header
    header_label = tk.Label(list_frame, text="Selected PDFs:", font=("Segoe UI", 12, "bold"),
                           bg='#ffffff', fg='#4a5568')
    header_label.pack(anchor='w', pady=(0, 10))

    # PDF list with scrollbar
    canvas = tk.Canvas(list_frame, bg='#f7fafc', height=200)
    scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg='#f7fafc')

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # Add PDF info
    total_pages = 0
    for i, pdf_path in enumerate(files, 1):
        try:
            reader = PdfReader(pdf_path)
            page_count = len(reader.pages)
            total_pages += page_count
            filename = os.path.basename(pdf_path)

            # PDF item frame
            item_frame = tk.Frame(scrollable_frame, bg='#f7fafc')
            item_frame.pack(fill='x', pady=2)

            # PDF info
            info_text = f"{i}. {filename} - {page_count} pages"
            info_label = tk.Label(item_frame, text=info_text, font=("Segoe UI", 10),
                                bg='#f7fafc', fg='#2d3748', anchor='w')
            info_label.pack(fill='x', padx=10, pady=2)

        except Exception as e:
            error_label = tk.Label(scrollable_frame, text=f"Error reading {os.path.basename(pdf_path)}: {str(e)}",
                                 font=("Segoe UI", 9), bg='#fed7d7', fg='#c53030')
            error_label.pack(fill='x', padx=10, pady=2)

    # Summary
    summary_frame = tk.Frame(list_frame, bg='#ffffff')
    summary_frame.pack(fill='x', pady=(10, 0))

    summary_text = f"Total: {len(files)} files, {total_pages} pages"
    summary_label = tk.Label(summary_frame, text=summary_text, font=("Segoe UI", 11, "bold"),
                           bg='#ffffff', fg='#2d3748')
    summary_label.pack()

    # Buttons
    button_frame = tk.Frame(preview_window, bg='#ffffff')
    button_frame.pack(fill='x', padx=20, pady=(0, 20))

    def confirm_merge():
        preview_window.destroy()
        do_merge(files)

    def cancel_merge():
        preview_window.destroy()

    cancel_btn = tk.Button(button_frame, text="Cancel", command=cancel_merge,
                          bg='#f7fafc', fg='#718096', font=('Segoe UI', 10),
                          relief='flat', borderwidth=1, padx=15, pady=8,
                          activebackground='#e2e8f0', activeforeground='#4a5568')
    cancel_btn.pack(side='right', padx=(10, 0))

    confirm_btn = tk.Button(button_frame, text="Merge PDFs", command=confirm_merge,
                           bg='#0078d4', fg='white', font=('Segoe UI', 11, 'bold'),
                           relief='flat', borderwidth=0, padx=20, pady=10,
                           activebackground="#89898a", activeforeground='white')
    confirm_btn.pack(side='right')

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

def do_merge(files):
    """Perform the actual PDF merging"""
    writer = PdfWriter()
    try:
        for pdf_path in files:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                writer.add_page(page)
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if output_path:
            with open(output_path, "wb") as output_file:
                writer.write(output_file)
            messagebox.showinfo("Success", f"Merged to {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Merge failed: {str(e)}")

def merge_pdfs():
    """Entry point for PDF merging with preview"""
    preview_merge()

root = tk.Tk()
root.title("DOCX to PDF Converter and PDF Merger")
root.geometry("450x400")
root.configure(bg='#ffffff')
root.resizable(False, False)

# Modern styling
style = ttk.Style()

# Configure main buttons
style.configure('Modern.TButton',
                font=('Segoe UI', 11, 'bold'),
                padding=10,
                relief='flat',
                borderwidth=0)
style.map('Modern.TButton',
          background=[('active', "#000000"), ('pressed', "#000000"), ('!active', "#000000")],
          foreground=[('active', 'white'), ('pressed', 'white'), ('!active', 'white')])

# Configure secondary buttons
style.configure('Secondary.TButton',
                font=('Segoe UI', 10),
                padding=8,
                relief='flat',
                borderwidth=0)
style.map('Secondary.TButton',
          background=[('active', '#e2e8f0'), ('pressed', '#cbd5e0'), ('!active', '#f7fafc')],
          foreground=[('active', '#4a5568'), ('pressed', '#2d3748'), ('!active', '#718096')])

# Load logo image
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

logo_image = Image.open(resource_path("assets/logo vibia.png"))
logo_image = logo_image.resize((220, 75), Image.Resampling.LANCZOS)
logo_photo = ImageTk.PhotoImage(logo_image)

# Main container
main_frame = tk.Frame(root, bg='#ffffff')
main_frame.pack(fill='both', expand=True, padx=20, pady=20)

# Logo image
logo_label = tk.Label(main_frame, image=logo_photo, bg='#ffffff')
logo_label.pack(pady=(0, 15))

# Title with modern styling
title_frame = tk.Frame(main_frame, bg='#ffffff')
title_frame.pack(pady=(0, 25))

text_label = tk.Label(title_frame, text="PDF Tools", font=("Segoe UI", 24, "bold"), bg='#ffffff', fg='#2d3748')
text_label.pack()

subtitle_label = tk.Label(title_frame, text="Convert and merge your documents", font=("Segoe UI", 10), bg='#ffffff', fg='#718096')
subtitle_label.pack(pady=(5, 0))

# Buttons container
buttons_frame = tk.Frame(main_frame, bg='#ffffff')
buttons_frame.pack(pady=(0, 20))

# Convert button with icon-like styling
btn_convert = ttk.Button(buttons_frame, text="📄 Convert DOCX to PDF",
                        command=convert_docx_to_pdf, style='Modern.TButton')
btn_convert.pack(fill='x', pady=(0, 15), ipady=8)

# Merge button with icon-like styling
btn_merge = ttk.Button(buttons_frame, text="📑 Merge PDFs",
                      command=merge_pdfs, style='Modern.TButton')
btn_merge.pack(fill='x', ipady=8)

# Footer
footer_label = tk.Label(main_frame, text="Ready to process your files", font=("Segoe UI", 9), bg='#ffffff', fg='#a0aec0')
footer_label.pack(side='bottom', pady=(20, 0))

root.mainloop()
