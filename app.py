import sys
import numpy as np
import os
import shutil
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QVBoxLayout,
    QPushButton,
    QLabel,
    QFileDialog,
    QListWidget,
    QWidget,
    QMessageBox,
)
from urllib.parse import unquote
import requests
import pandas as pd
from PySide6.QtGui import QColor, QMovie
from pdf_utils import process_pdf, pdf_check
from openpyxl import load_workbook

class AppInit(QThread):
    init_complete = Signal() 
    init_error = Signal(str) 
    init_success = Signal(pd.DataFrame, object, object) 

    def __init__(self, excel_path):
        super().__init__()
        self.excel_path = excel_path

    def run(self):
        try:
            mt_df = pd.read_excel(self.excel_path, sheet_name='Sheet1', engine='openpyxl')

            wb = load_workbook(self.excel_path)
            ws = wb['Sheet1']

            self.init_success.emit(mt_df, wb, ws)
            self.init_complete.emit()
        except Exception as e:
            self.init_error.emit(str(e))

    def clean_up(self):
        """Method to clean up after thread finishes."""
       
        if self.isRunning():
            self.quit() 
            self.wait()  
        self.deleteLater()  

class DownloaderThread(QThread):
    download_completed = Signal(str)
    download_finished = Signal()
    download_error = Signal(str, str)

    def __init__(self, file_link):
        super().__init__()
        self.file_link = file_link
    
    def run(self):
        try:
            for link in self.file_link:
                url = link.replace("?Web=1", "")
                file_name = unquote(url.split("/")[-1])
                output_path = os.path.join("local_storage", file_name)
                try:
                    response = requests.get(url)
                    response.raise_for_status()
                    with open(output_path, "wb") as file:
                        file.write(response.content)
                    self.download_completed.emit(output_path)

                except requests.exceptions.RequestException as e:
                    self.download_error.emit("download_err",str(e))
                    break
        except Exception as e:
            self.download_error.emit("thread_err", str(e))
        self.download_finished.emit()
    
    def clean_up(self):
        """Method to clean up after thread finishes."""
       
        if self.isRunning():
            self.quit() 
            self.wait()  
        self.deleteLater()  


class ScannerThread(QThread):
    """Thread for scanning PDFs."""
    scan_completed = Signal(str, dict)  
    scan_finished = Signal()
    scan_error = Signal(str)  

    def __init__(self, uploaded_files, poppler_path, mt_df, wb, ws):
        super().__init__()
        self.uploaded_files = uploaded_files
        self.poppler_path = poppler_path
        self.mt_df = mt_df
        self.wb = wb
        self.ws = ws

    def run(self):
        try:
            for file_path in self.uploaded_files:
                header_color = np.array([254, 0, 0])
                data_color = np.array([255, 192, 0])
                table, bottom_text = process_pdf(file_path, self.poppler_path, header_color, data_color)
                scan_result = pdf_check(self.mt_df, table, bottom_text, self.wb, self.ws)
                self.scan_completed.emit(file_path, scan_result)
            self.scan_finished.emit()
        except Exception as e:
            self.scan_error.emit(str(e))

    def clean_up(self):
        """Method to clean up after thread finishes."""
       
        if self.isRunning():
            self.quit() 
            self.wait()  
        self.deleteLater()  

class PDFScannerApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("BAA Scanner")
        self.setGeometry(100, 100, 600, 400)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout(self.central_widget)

        self.info_label = QLabel("Upload PDF files to scan:")
        self.info_label.setAlignment(Qt.AlignCenter)

        self.file_list_widget = QListWidget()

        self.upload_button = QPushButton("Upload PDF Files")
        self.upload_button.clicked.connect(self.upload_files)

        self.delete_button = QPushButton("Delete Selected File")
        self.delete_button.clicked.connect(self.delete_selected_file)
        self.delete_button.setEnabled(False)

        self.start_button = QPushButton("Start Scanning")
        self.start_button.clicked.connect(self.start_scanning)
        self.start_button.setEnabled(False)

        self.download_button = QPushButton("Download Files")
        self.download_button.clicked.connect(self.download_files)

        self.layout.addWidget(self.info_label)
        self.layout.addWidget(self.file_list_widget)
        self.layout.addWidget(self.upload_button)
        self.layout.addWidget(self.delete_button)
        self.layout.addWidget(self.start_button)
        self.layout.addWidget(self.download_button)

        self.loading_label = QLabel()
        self.layout.addWidget(self.loading_label)

        self.uploaded_files = {}

        self.active_threads = []

        self.poppler_path = r"poppler-24.07.0\Library\bin"

        self.mt_df = None
        self.wb = None
        self.ws = None

        self.show_loading_indicator(True)

        app_init = AppInit('mt_database.xlsx')
        app_init.init_success.connect(self.on_init_success)
        app_init.init_error.connect(self.on_init_error)
        self.active_threads.append(app_init) 
        app_init.start()  

        self.file_list_widget.itemSelectionChanged.connect(self.toggle_delete_button)

    def upload_files(self):
        file_dialog = QFileDialog(self)
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        file_dialog.setNameFilter("PDF Files (*.pdf)")

        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            local_storage_dir = "local_storage"
            os.makedirs(local_storage_dir, exist_ok=True)

            for file_path in selected_files:
                file_name = os.path.basename(file_path)
                destination_path = os.path.join(local_storage_dir, file_name)

                if destination_path not in self.uploaded_files:
                    shutil.copy(file_path, destination_path)
                    self.uploaded_files[destination_path] = None
                    self.file_list_widget.addItem(destination_path)

            self.start_button.setEnabled(bool(self.uploaded_files))

    def download_files(self):
        self.show_loading_indicator(True)
        file_link = self.mt_df[self.mt_df['BAA Link'].notna()]['BAA Link']
        downloader_thread = DownloaderThread(file_link) 
        downloader_thread.download_completed.connect(self.on_download_completed)
        downloader_thread.download_finished.connect(self.on_download_finished)
        downloader_thread.download_error.connect(self.on_download_error)
        self.active_threads.append(downloader_thread) 
        downloader_thread.start()

    def on_download_completed(self, file_path):
        if file_path not in self.uploaded_files:
            self.uploaded_files[file_path] = None
            self.file_list_widget.addItem(file_path)

    def on_download_finished(self):
        self.show_loading_indicator(False)
        for thread in self.active_threads:
            thread.quit()

    def on_download_error(self, err_type, error_message):
        if err_type == "download_err":
            QMessageBox.warning(self, "Error", f"An error occurred during download: {error_message}")
            self.show_loading_indicator(False)
        else:
            QMessageBox.critical(self, "Error", f"An error occurred during download: {error_message}")
            self.show_loading_indicator(False)

    def delete_selected_file(self):
        selected_items = self.file_list_widget.selectedItems()
        if not selected_items:
            return

        for item in selected_items:
            file_path = item.text()

            if os.path.exists(file_path):
                os.remove(file_path)

            self.file_list_widget.takeItem(self.file_list_widget.row(item))
            del self.uploaded_files[file_path]

        self.start_button.setEnabled(bool(self.uploaded_files))
        self.toggle_delete_button() 

    def start_scanning(self):
        if not self.uploaded_files:
            QMessageBox.warning(self, "Warning", "No files uploaded.")
            return

        self.show_loading_indicator(True)

        scanner_thread = ScannerThread(self.uploaded_files, self.poppler_path, self.mt_df, self.wb, self.ws)
        scanner_thread.scan_completed.connect(self.on_scan_completed)
        scanner_thread.scan_error.connect(self.on_scan_error)
        scanner_thread.scan_finished.connect(self.on_scan_finished)
        self.active_threads.append(scanner_thread) 
        scanner_thread.start()

    def toggle_delete_button(self):
        self.delete_button.setEnabled(bool(self.file_list_widget.selectedItems()))

    def show_loading_indicator(self, show):
        """Show or hide the loading indicator"""
        if show:
            movie = QMovie("assets/loading.gif") 
            self.loading_label.setMovie(movie)
            movie.start()
            self.loading_label.setAlignment(Qt.AlignCenter)
        else:
            self.loading_label.clear()

    def on_scan_finished(self):
        self.show_loading_indicator(False)
        for thread in self.active_threads:
            thread.quit()

    def on_scan_completed(self, file_path, scan_result):
        """Handle the completion of a scan."""
        try:
            for i in range(self.file_list_widget.count()):
                item = self.file_list_widget.item(i)
                if item.text() == file_path:
                    self.uploaded_files[file_path] = scan_result
                    if all(scan_result.values()):
                        item.setBackground(QColor("lightgreen"))
                    else:
                        item.setBackground(QColor("red"))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {e}")

    def on_scan_error(self, error_message):
        """Handle any errors during the scan."""
        QMessageBox.critical(self, "Error", f"An error occurred during scanning: {error_message}")
        self.show_loading_indicator(False)

    def on_init_success(self, mt_df, wb, ws):
        self.mt_df = mt_df
        self.wb = wb
        self.ws = ws
        self.show_loading_indicator(False)

    def on_init_error(self, error_message):
        QMessageBox.critical(self, "Error", f"An error occurred during init: {error_message}")
        self.show_loading_indicator(False)

    def closeEvent(self, event):
        """Ensure all threads are cleaned up before closing the application."""
        for thread in self.active_threads:
            thread.clean_up()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFScannerApp()
    window.show()
    sys.exit(app.exec())
