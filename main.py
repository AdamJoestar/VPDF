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

def merge_pdfs():
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if files:
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

root = tk.Tk()
root.title("DOCX to PDF Converter and PDF Merger")
root.geometry("450x400")
root.configure(bg='#ffffff')
root.resizable(False, False)

# Modern styling
style = ttk.Style()
style.configure('Modern.TButton',
                font=('Segoe UI', 11, 'bold'),
                padding=10,
                relief='flat',
                borderwidth=0)
style.map('Modern.TButton',
          background=[('active', "#000000b0"), ('!active', "#000000")],
          foreground=[('active', 'white'), ('!active', 'white')])

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
