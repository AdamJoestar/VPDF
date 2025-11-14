# main.py
import sys
from PyQt5.QtWidgets import QApplication
from gui import FileConverterApp

if __name__ == '__main__':
    # Pastikan PyQt5 tersedia
    try:
        app = QApplication(sys.argv)
        window = FileConverterApp()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        print(f"Error al ejecutar la aplicación: {e}")
        # Tambahkan notifikasi error jika perlu