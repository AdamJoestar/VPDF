# converter_logic.py
import os
import shutil
from docx2pdf import convert
from pypdf import PdfWriter, PdfReader

# PENTING: Untuk fungsi convert_docx_to_pdf, pastikan
# Microsoft Word (Windows) atau LibreOffice (Linux/macOS) sudah terinstal.

def convert_docx_to_pdf(input_path, output_dir, progress_callback=None):
    """Mengkonversi file DOCX tunggal ke PDF."""
    try:
        # Tentukan nama file output
        filename = os.path.basename(input_path)
        pdf_filename = filename.replace(".docx", ".pdf")
        output_path = os.path.join(output_dir, pdf_filename)

        # Panggil fungsi konversi dari pustaka docx2pdf
        convert(input_path, output_path)

        if progress_callback:
            progress_callback(100)  # Set progress to 100% on completion

        return True, f"Berhasil konversi ke: {output_path}"
    except Exception as e:
        # Check if the error is related to MS Word not being installed
        error_str = str(e).lower()
        if "word" in error_str or "office" in error_str or "libreoffice" in error_str:
            return False, "Konversi Gagal: Microsoft Word atau LibreOffice tidak terinstal. Silakan instal salah satu untuk konversi DOCX ke PDF."
        else:
            # Handle specific errors
            if "write" in error_str or "permission" in error_str:
                return False, f"Konversi Gagal: Tidak dapat menulis file output. Periksa izin folder dan pastikan file tidak sedang digunakan oleh aplikasi lain."
            else:
                return False, f"Konversi Gagal: {str(e)}"

def merge_pdfs(file_list, output_path):
    """Menggabungkan daftar file PDF menjadi satu file PDF."""
    if not file_list:
        return False, "Daftar file kosong."

    writer = PdfWriter()
    try:
        for filename in file_list:
            try:
                reader = PdfReader(filename)
                # Tambahkan semua halaman dari setiap file ke writer
                for page in reader.pages:
                    writer.add_page(page)
            except Exception as e:
                return False, f"Gagal membaca file {os.path.basename(filename)}: {str(e)}"

        # Tulis output yang digabungkan ke file
        try:
            with open(output_path, "wb") as f:
                writer.write(f)
        except Exception as e:
            error_str = str(e).lower()
            if "write" in error_str or "permission" in error_str:
                return False, f"Penggabungan Gagal: Tidak dapat menulis file output. Periksa izin folder dan pastikan file tidak sedang digunakan oleh aplikasi lain."
            else:
                return False, f"Penggabungan Gagal: {str(e)}"

        return True, f"Berhasil menggabungkan {len(file_list)} file ke: {output_path}"
    except Exception as e:
        return False, f"Penggabungan Gagal: {str(e)}"

def process_and_merge_mixed_files(file_list, output_path):
    """Menangani daftar file campuran (DOCX dan PDF) untuk digabungkan."""
    
    # 1. Siapkan folder sementara untuk file DOCX yang dikonversi
    temp_dir = "temp_pdf_files"
    os.makedirs(temp_dir, exist_ok=True)
    
    pdf_files_to_merge = []
    
    success = True
    messages = []
    
    try:
        for input_path in file_list:
            if input_path.lower().endswith(".pdf"):
                # Jika sudah PDF, langsung tambahkan
                pdf_files_to_merge.append(input_path)
            elif input_path.lower().endswith(".docx"):
                # Jika DOCX, konversi ke folder sementara
                docx_success, msg = convert_docx_to_pdf(input_path, temp_dir)
                if docx_success:
                    # Ambil path file PDF yang baru dikonversi
                    pdf_filename = os.path.basename(input_path).replace(".docx", ".pdf")
                    temp_pdf_path = os.path.join(temp_dir, pdf_filename)
                    pdf_files_to_merge.append(temp_pdf_path)
                else:
                    success = False
                    messages.append(f"Gagal konversi {os.path.basename(input_path)}: {msg}")
                    
        # 2. Lakukan Penggabungan PDF
        if success and pdf_files_to_merge:
            merge_success, merge_msg = merge_pdfs(pdf_files_to_merge, output_path)
            messages.append(merge_msg)
            success = merge_success
            
        elif not pdf_files_to_merge:
            messages.append("Tidak ada file yang valid untuk digabungkan.")
            success = False

    finally:
        # 3. Bersihkan folder sementara
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
            
    return success, "\n".join(messages)