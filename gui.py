# gui.py
import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QHBoxLayout, QPushButton, QLineEdit, QLabel,
    QFileDialog, QTabWidget, QListWidget, QListWidgetItem,
    QMessageBox, QProgressBar
)
from PyQt5.QtCore import QDir
from PyQt5.QtGui import QPixmap, QPalette, QBrush, QFont
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QCloseEvent
from converter_logic import convert_docx_to_pdf, process_and_merge_mixed_files

def resource_path(relative_path):
    """ Mendapatkan path absolut ke sumber daya, berfungsi untuk mode pengembangan dan PyInstaller """
    try:
        # PyInstaller membuat folder sementara dan menyimpan path di _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))

    return os.path.join(base_path, relative_path)

class ConversionWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(bool, str)

    def __init__(self, input_path, output_dir):
        super().__init__()
        self.input_path = input_path
        self.output_dir = output_dir

    def run(self):
        success, message = convert_docx_to_pdf(self.input_path, self.output_dir, self.progress.emit)
        self.finished.emit(success, message)

class MergeWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(bool, str)

    def __init__(self, file_paths, output_path):
        super().__init__()
        self.file_paths = file_paths
        self.output_path = output_path

    def run(self):
        success, message = process_and_merge_mixed_files(self.file_paths, self.output_path)
        self.finished.emit(success, message)

class FileConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VPDF - DOCX/PDF Utility")
        self.setGeometry(100, 100, 800, 600)

        # Atur latar belakang jendela
        self._set_background()

        # Widget Utama dan Layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(30, 30, 30, 30)
        self.main_layout.setSpacing(20)

        # Title with logo
        title_layout = QHBoxLayout()
        title_layout.setSpacing(15)

        # Title
        title_label = QLabel("VPDF - DOCX/PDF Utility")
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #333333;")
        title_layout.addWidget(title_label)

        title_layout.addStretch()

        # Logo
        self.logo_label = QLabel()
        logo_path = resource_path("assets/logo vibia.png")
        pixmap = QPixmap(logo_path)
        if not pixmap.isNull():
            scaled_pixmap = pixmap.scaled(90,90, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.logo_label.setPixmap(scaled_pixmap)
        else:
            self.logo_label.hide()
        title_layout.addWidget(self.logo_label)

        self.main_layout.addLayout(title_layout)

        # Tab Widget
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #ced4da;
                background: #ffffff;
                border-radius: 10px;
            }
            QTabBar::tab {
                background: #f8f9fa;
                border: 1px solid #ced4da;
                padding: 12px 24px;
                margin-right: 5px;
                border-radius: 8px;
                color: #333333;
                font-weight: bold;
                font-size: 14px;
            }
            QTabBar::tab:selected {
                background: #ffffff;
                border-bottom: 2px solid #6c757d;
            }
            QTabBar::tab:hover {
                background: #e9ecef;
            }
        """)
        self.main_layout.addWidget(self.tabs)

        # Buat halaman-halaman Tab
        self.tab_convert = QWidget()
        self.tab_merge = QWidget()

        self.tabs.addTab(self.tab_convert, "Convertir DOCX a PDF")
        self.tabs.addTab(self.tab_merge, "Combinar archivos en PDF")

        # Setup UI untuk setiap tab
        self._setup_convert_tab()
        self._setup_merge_tab()

        # Apply modern stylesheet
        self._apply_modern_stylesheet()

    def _set_background(self):
        # Set white background
        palette = QPalette()
        palette.setColor(QPalette.Window, Qt.white)
        self.setPalette(palette)
        self.setAutoFillBackground(True)

    def _add_logo_overlay(self):
        # Add logo to top-left
        self.logo_label = QLabel(self)
        pixmap = QPixmap("assets/logo vibia.png")
        if not pixmap.isNull():
            scaled_pixmap = pixmap.scaled(80, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.logo_label.setPixmap(scaled_pixmap)
            self.logo_label.setAlignment(Qt.AlignCenter)
            # Position logo at top-left
            self.logo_label.move(20, 20)
        else:
            self.logo_label.hide()

    def _apply_modern_stylesheet(self):
        # Modern light theme stylesheet with gray buttons
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
            }
            QWidget {
                background: #ffffff;
                color: #333333;
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                font-size: 14px;
            }
            QPushButton {
                background-color: #6c757d;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                color: white;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #5a6268;
            }
            QPushButton:pressed {
                background-color: #545b62;
            }
            QPushButton:disabled {
                background-color: #adb5bd;
            }
            QLineEdit {
                padding: 8px;
                border: 1px solid #ced4da;
                border-radius: 5px;
                background-color: #ffffff;
                color: #333333;
            }
            QLineEdit:focus {
                border-color: #6c757d;
            }
            QLabel {
                color: #333333;
                font-weight: bold;
            }
            QListWidget {
                background-color: #ffffff;
                border: 1px solid #ced4da;
                border-radius: 5px;
                padding: 5px;
                color: #333333;
            }
            QListWidget::item {
                padding: 5px;
                border-bottom: 1px solid #dee2e6;
            }
            QListWidget::item:selected {
                background-color: #6c757d;
                color: white;
            }
            QMessageBox {
                background-color: #ffffff;
                color: #333333;
            }
            QMessageBox QLabel {
                color: #333333;
            }
            QMessageBox QPushButton {
                background-color: #6c757d;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 3px;
            }
        """)

    # --------------------------------------------------------------------------
    # Tab Konversi DOCX ke PDF
    # --------------------------------------------------------------------------
    
    def _setup_convert_tab(self):
        layout = QVBoxLayout(self.tab_convert)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # Input File DOCX
        layout.addWidget(QLabel("Seleccionar archivo DOCX:"))
        h_layout_input = QHBoxLayout()
        self.convert_input_line = QLineEdit()
        self.convert_input_line.setPlaceholderText("Seleccionar archivo .docx...")
        self.convert_input_line.setReadOnly(True)
        h_layout_input.addWidget(self.convert_input_line)

        btn_browse = QPushButton("Examinar")
        btn_browse.clicked.connect(self._select_docx_file)
        h_layout_input.addWidget(btn_browse)

        layout.addLayout(h_layout_input)

        # Tombol Konversi
        self.btn_convert = QPushButton("Iniciar conversión")
        self.btn_convert.clicked.connect(self._start_conversion)
        layout.addWidget(self.btn_convert)

        # Progress Bar
        self.convert_progress_bar = QProgressBar()
        self.convert_progress_bar.setVisible(False)
        layout.addWidget(self.convert_progress_bar)

        # Status
        self.convert_status_label = QLabel("Estado: Listo")
        layout.addWidget(self.convert_status_label)

        layout.addStretch(1)

    def _select_docx_file(self):
        # Membuka dialog untuk memilih file DOCX
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Seleccionar archivo DOCX", 
            "", 
            "Documentos de Word (*.docx)"
        )
        if file_path:
            self.convert_input_line.setText(file_path)
            self.convert_status_label.setText("Estado: Archivo seleccionado")

    def _start_conversion(self):
        input_path = self.convert_input_line.text()

        if not input_path or not os.path.exists(input_path):
            QMessageBox.warning(self, "Advertencia", "Por favor, seleccione primero un archivo DOCX.")
            return

        # Ambil direktori tempat file output akan disimpan
        output_dir = os.path.dirname(input_path)

        self.convert_status_label.setText("Estado: Convirtiendo...")
        self.btn_convert.setEnabled(False)
        self.convert_progress_bar.setVisible(True)
        self.convert_progress_bar.setValue(0)

        # Start worker thread
        self.conversion_worker = ConversionWorker(input_path, output_dir)
        self.conversion_worker.progress.connect(self.convert_progress_bar.setValue)
        self.conversion_worker.finished.connect(self._on_conversion_finished)
        self.conversion_worker.start()

    def _on_conversion_finished(self, success, message):
        self.convert_progress_bar.setVisible(False)
        self.btn_convert.setEnabled(True)

        if success:
            QMessageBox.information(self, "Éxito", f"¡Conversión exitosa!\n{message}")
        else:
            QMessageBox.critical(self, "Fallido", f"¡Conversión fallida!\n{message}")

        self.convert_status_label.setText("Estado: Completado")

    # --------------------------------------------------------------------------
    # Tab Penggabungan DOCX/PDF ke PDF
    # --------------------------------------------------------------------------
    
    def _setup_merge_tab(self):
        layout = QVBoxLayout(self.tab_merge)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # Daftar File Input
        layout.addWidget(QLabel("Lista de archivos de entrada (.docx o .pdf):"))
        self.file_list_widget = QListWidget()
        self.file_list_widget.setMaximumHeight(200)
        layout.addWidget(self.file_list_widget)

        # Tombol Aksi untuk Daftar
        h_layout_actions = QHBoxLayout()
        self.btn_add_file = QPushButton("Agregar archivo")
        self.btn_add_file.clicked.connect(self._add_file_to_merge_list)
        h_layout_actions.addWidget(self.btn_add_file)

        self.btn_remove_file = QPushButton("Eliminar seleccionados")
        self.btn_remove_file.clicked.connect(self._remove_selected_file)
        h_layout_actions.addWidget(self.btn_remove_file)

        layout.addLayout(h_layout_actions)

        # Tombol Gabung
        self.btn_merge = QPushButton("Comenzar a combinar en PDF")
        self.btn_merge.clicked.connect(self._start_merging)
        layout.addWidget(self.btn_merge)

        # Progress Bar
        self.merge_progress_bar = QProgressBar()
        self.merge_progress_bar.setVisible(False)
        layout.addWidget(self.merge_progress_bar)

        # Status
        self.merge_status_label = QLabel("Estado: Listo")
        layout.addWidget(self.merge_status_label)

        layout.addStretch(1)
        
    def _add_file_to_merge_list(self):
        # Membuka dialog untuk memilih banyak file (DOCX atau PDF)
        files, _ = QFileDialog.getOpenFileNames(
            self, 
            "Elige un archivo DOCX o PDF para combinar", 
            "", 
            "Documentos (*.docx *.pdf)"
        )
        if files:
            for file_path in files:
                # Tambahkan jalur file ke QListWidget
                QListWidgetItem(file_path, self.file_list_widget)
            self.merge_status_label.setText(f"Estado: Agregado {len(files)} archivo")

    def _remove_selected_file(self):
        for item in self.file_list_widget.selectedItems():
            self.file_list_widget.takeItem(self.file_list_widget.row(item))
        self.merge_status_label.setText("Estado: Archivo eliminado")

    def _start_merging(self):
        # Kumpulkan semua jalur file dari QListWidget
        file_paths = [self.file_list_widget.item(i).text() for i in range(self.file_list_widget.count())]

        if not file_paths:
            QMessageBox.warning(self, "Advertencia", "Por favor, agregue un archivo para combinar.")
            return

        # Dialog untuk memilih lokasi dan nama file output
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar archivo PDF combinado como",
            os.path.expanduser("~/Combinación.pdf"),
            "PDF Files (*.pdf)"
        )

        if not output_path:
            return

        self.merge_status_label.setText("Estado: Procesando y Combinando...")
        self.btn_merge.setEnabled(False)
        self.merge_progress_bar.setVisible(True)
        self.merge_progress_bar.setValue(0)

        # Start worker thread
        self.merge_worker = MergeWorker(file_paths, output_path)
        self.merge_worker.progress.connect(self.merge_progress_bar.setValue)
        self.merge_worker.finished.connect(self._on_merge_finished)
        self.merge_worker.start()

    def _on_merge_finished(self, success, message):
        self.merge_progress_bar.setVisible(False)
        self.btn_merge.setEnabled(True)

        if success:
            QMessageBox.information(self, "Éxito", f"¡Fusión exitosa!\n{message}")
        else:
            QMessageBox.critical(self, "Fallido", f"¡Fusión fallida!\n{message}")

        self.merge_status_label.setText("Estado: Completado")

    def closeEvent(self, event: QCloseEvent):
        # Show confirmation dialog before closing
        reply = QMessageBox.question(
            self,
            "Confirmar salida",
            "¿Estás seguro de que quieres salir de la aplicación?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

if __name__ == '__main__':
    # Untuk pengujian cepat, tapi sebaiknya jalankan dari main.py
    app = QApplication(sys.argv)
    window = FileConverterApp()
    window.show()
    sys.exit(app.exec_())
