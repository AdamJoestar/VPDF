# hook-docx2pdf.py
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# Secara eksplisit sertakan docx2pdf dan semua dependensi COM yang tersembunyi
hiddenimports = collect_submodules('docx2pdf') + collect_submodules('win32com') + collect_submodules('pythoncom') + ['docx2pdf.main']

# Sertakan file data yang diperlukan
datas = collect_data_files('docx2pdf')
