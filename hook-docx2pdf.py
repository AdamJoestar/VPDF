# hook-docx2pdf.py
from PyInstaller.utils.hooks import collect_submodules, copy_metadata

# Beritahu PyInstaller untuk menyertakan semua modul tersembunyi dari win32com dan pythoncom
hiddenimports = collect_submodules('win32com') + collect_submodules('pythoncom')

# Sertakan metadata untuk docx2pdf untuk mengatasi PackageNotFoundError
datas = copy_metadata('docx2pdf')
