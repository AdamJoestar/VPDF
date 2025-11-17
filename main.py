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
root.geometry("400x300")
root.configure(bg='#f0f0f0')

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
logo_image = logo_image.resize((200, 70), Image.Resampling.LANCZOS)
logo_photo = ImageTk.PhotoImage(logo_image)

# Logo image
logo_label = tk.Label(root, image=logo_photo, bg='#f0f0f0')
logo_label.pack(pady=10)

# Text below logo
text_label = tk.Label(root, text="PDF Tools", font=("Arial", 20, "bold"), bg='#f0f0f0', fg='#333')
text_label.pack(pady=5)

# Frame for buttons
frame = tk.Frame(root, bg='#f0f0f0')
frame.pack(pady=20)

btn_convert = ttk.Button(frame, text="Convert DOCX to PDF", command=convert_docx_to_pdf, width=20)
btn_convert.pack(pady=10)

btn_merge = ttk.Button(frame, text="Merge PDFs", command=merge_pdfs, width=20)
btn_merge.pack(pady=10)

root.mainloop()
