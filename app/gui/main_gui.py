from PyQt6.QtWidgets import QMainWindow, QFileDialog, QMessageBox, QProgressDialog
from PyQt6.QtCore import QThread, pyqtSignal
from gui.interface import Ui_MainWindow
import logic.main_logic as main_logic
from utils.resources import resource_path
import os


class Worker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, input_file: str, output_file: str):
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file

    def run(self):
        try:
            # Panggil fungsi utama
            message = main_logic.run_excel_process(self.input_file, self.output_file)
            self.finished.emit(message)
        except Exception as e:
            self.error.emit(str(e))


class MainApp(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(MainApp, self).__init__()
        self.setupUi(self)

        # Load external stylesheet
        style_path = resource_path("style/style.qss")
        if os.path.exists(style_path):
            with open(style_path, "r") as f:
                self.setStyleSheet(f.read())

        # Variabel untuk file
        self.input_file = None
        self.output_file = None

        # Start button awalnya disable
        self.start_btn.setEnabled(False)

        # Hubungkan tombol
        self.input_btn.clicked.connect(self.select_input_file)
        self.output_btn.clicked.connect(self.select_output_file)
        self.start_btn.clicked.connect(self.start_process)

    def select_input_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Input Excel File", "", "Excel Files (*.xlsx *.xlsm)"
        )
        if file_path:
            self.input_file = file_path
            self.input_line.setText(file_path)
            self.update_start_button_state()

    def select_output_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Output Excel File", "", "Excel Files (*.xlsx *.xlsm)"
        )
        if file_path:
            if not file_path.endswith(".xlsx") and not file_path.endswith(".xlsm"):
                file_path += ".xlsx"
            self.output_file = file_path
            self.output_line.setText(file_path)
            self.update_start_button_state()

    def update_start_button_state(self):
        """
        Aktifkan tombol Start hanya jika input & output sudah dipilih
        """
        self.start_btn.setEnabled(bool(self.input_file and self.output_file))

    def start_process(self):
        if not self.input_file or not self.output_file:
            QMessageBox.warning(self, "Warning", "Please select both input and output files!")
            return

        # Disable tombol Start saat proses berjalan
        self.start_btn.setEnabled(False)

        # Progress dialog
        self.progress = QProgressDialog("⏳ Please wait, processing data...", None, 0, 0, self)
        self.progress.setWindowTitle("Processing")
        self.progress.setCancelButton(None)
        self.progress.setMinimumDuration(0)
        self.progress.setAutoClose(False)
        self.progress.show()

        # Jalankan worker
        self.worker = Worker(self.input_file, self.output_file)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_finished(self, message: str):
        self.progress.close()
        QMessageBox.information(self, "✅ Process Completed", message)

        # Reset tombol Start
        self.start_btn.setEnabled(True)
        self.start_btn.setText("▶️ Start")

    def on_error(self, msg: str):
        self.progress.close()
        QMessageBox.critical(self, "Error", f"❌ Process Failed: {msg}")

        # Reset tombol Start
        self.start_btn.setEnabled(True)
        self.start_btn.setText("▶️ Start")