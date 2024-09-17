import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QListWidget, QPushButton, QVBoxLayout, QWidget, 
    QFileDialog, QMessageBox, QHBoxLayout, QProgressBar, QLabel, QGridLayout
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QPixmap, QDragEnterEvent, QDropEvent
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image
from io import BytesIO
from docx import Document
from docx.shared import Inches
from lxml import etree
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import tempfile
import comtypes.client

class FileProcessorThread(QThread):
    progress = pyqtSignal(int)
    error = pyqtSignal(str)

    def __init__(self, function, *args, **kwargs):
        super().__init__()
        self.function = function
        self.args = args
        self.kwargs = kwargs

    def run(self):
        try:
            self.function(*self.args, **self.kwargs)
        except Exception as e:
            self.error.emit(str(e))

class FileMerger(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dungerz All File Merger")
        self.setGeometry(100, 100, 800, 550)
        self.setAcceptDrops(True)  # Enable drag and drop for the entire window
        self.initUI()

    def initUI(self):
        # Main layout
        layout = QGridLayout()

        # Title and Description Labels
        titleLabel = QLabel("Dungerz All File Merger")
        titleLabel.setAlignment(Qt.AlignCenter)
        titleLabel.setStyleSheet("font-size: 24px; font-weight: bold; margin-bottom: 5px;")
        
        descriptionLabel = QLabel("Drag and drop files or use the buttons below to add files. You can merge PDFs, DOCX, XML, and images into a single document.")
        descriptionLabel.setAlignment(Qt.AlignCenter)
        descriptionLabel.setWordWrap(True)
        descriptionLabel.setStyleSheet("font-size: 14px; margin-bottom: 20px;")

        # List widget to display files
        self.listWidget = QListWidget()
        self.listWidget.setStyleSheet("""
            QListWidget {
                background-color: #f0f0f0;
                border: 2px dashed #cccccc;
                font-size: 14px;
                color: black;
            }
        """)
        self.listWidget.setDragEnabled(True)
        self.listWidget.setSelectionMode(QListWidget.MultiSelection)
        self.listWidget.setDragDropMode(QListWidget.DragDropMode.InternalMove)

        # Progress Bar
        self.progressBar = QProgressBar()
        self.progressBar.setTextVisible(True)
        self.progressBar.setMaximum(100)
        self.progressBar.setValue(0)
        self.progressBar.setStyleSheet("QProgressBar {height: 20px; border: 1px solid #000000;}")

        # Button Layout
        buttonLayout = QHBoxLayout()

        # Add File Button
        self.addButton = QPushButton("Add Files")
        self.addButton.setStyleSheet("""
            QPushButton {
                background-color: #007bff;
                color: white;
                font-size: 16px;
                padding: 10px;
                border-radius: 8px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """)
        self.addButton.clicked.connect(self.open_file_explorer)

        # Merge Button
        self.mergeButton = QPushButton("Merge and Save")
        self.mergeButton.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                font-size: 16px;
                padding: 10px;
                border-radius: 8px;
            }
            QPushButton:hover {
                background-color: #218838;
            }
        """)
        self.mergeButton.clicked.connect(self.merge_files)

        # Clear Button
        self.clearButton = QPushButton("Clear List")
        self.clearButton.setStyleSheet("""
            QPushButton {
                background-color: #dc3545;
                color: white;
                font-size: 16px;
                padding: 10px;
                border-radius: 8px;
            }
            QPushButton:hover {
                background-color: #c82333;
            }
        """)
        self.clearButton.clicked.connect(self.clear_list)

        # Remove Selected Button
        self.removeSelectedButton = QPushButton("Remove Selected")
        self.removeSelectedButton.setStyleSheet("""
            QPushButton {
                background-color: #ffc107;
                color: black;
                font-size: 16px;
                padding: 10px;
                border-radius: 8px;
            }
            QPushButton:hover {
                background-color: #e0a800;
            }
        """)
        self.removeSelectedButton.clicked.connect(self.remove_selected_items)

        # Add buttons to the button layout
        buttonLayout.addWidget(self.addButton)
        buttonLayout.addWidget(self.mergeButton)
        buttonLayout.addWidget(self.clearButton)
        buttonLayout.addWidget(self.removeSelectedButton)

        # Position elements in layout
        layout.addWidget(titleLabel, 0, 0, 1, 3)
        layout.addWidget(descriptionLabel, 1, 0, 1, 3)
        layout.addWidget(self.listWidget, 2, 0, 1, 3)
        layout.addWidget(self.progressBar, 3, 0, 1, 3)
        layout.addLayout(buttonLayout, 4, 0, 1, 3)

        # Set the main widget
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    # Drag Enter Event (for the entire window)
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        self.add_files(files)
    # File Explorer to open PDF, DOCX, Image, or XML files
    def open_file_explorer(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Files", "", 
            "All Files (*);;PDF Files (*.pdf);;Word Files (*.docx);;Image Files (*.jpg *.jpeg *.png *.bmp *.gif *.tiff);;XML Files (*.xml)"
        )
        self.add_files(files)

    def add_files(self, files):
        for file in files:
            if file.lower().endswith(('.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.docx', '.xml')):
                # Check if the file is already in the list
                items = self.listWidget.findItems(file, Qt.MatchExactly)
                if not items:
                    self.listWidget.addItem(file)
            else:
                QMessageBox.warning(self, "Unsupported File", f"The file {os.path.basename(file)} is not supported.")

    # Merge the selected files and export to desired format (PDF, DOCX, etc.)
    def merge_files(self):
        if self.listWidget.count() == 0:
            QMessageBox.information(self, "No Files Selected", "Please add files to merge.")
            return
        
        # Choose the export format
        export_formats = "PDF Files (*.pdf);;Word Files (*.docx);;All Files (*)"
        output_filename, selected_format = QFileDialog.getSaveFileName(self, "Save Merged File", "", export_formats)
        if not output_filename:
            return

        _, file_extension = os.path.splitext(output_filename)
        
        if file_extension.lower() == '.pdf':
            self.process_files(self.export_to_pdf, output_filename)
        elif file_extension.lower() == '.docx':
            self.process_files(self.export_to_docx, output_filename)
        else:
            QMessageBox.warning(self, "Unsupported Format", f"The format {file_extension} is not supported.")

    def process_files(self, export_function, output_filename):
        self.thread = FileProcessorThread(export_function, output_filename)
        self.thread.progress.connect(self.update_progress)
        self.thread.error.connect(self.show_error_message)
        self.thread.finished.connect(self.on_merge_finished)
        self.thread.start()
        self.mergeButton.setEnabled(False)

    def update_progress(self, value):
        self.progressBar.setValue(value)

    def show_error_message(self, message):
        QMessageBox.warning(self, "Merge Failed", f"An error occurred during the merge process: {message}")

    def on_merge_finished(self):
        self.progressBar.setValue(0)
        self.mergeButton.setEnabled(True)
        QMessageBox.information(self, "Merge Successful", f"Merged file saved successfully.")
        self.listWidget.clear()

    # Export merged files as PDF
    def export_to_pdf(self, output_filename):
        pdf_writer = PdfWriter()
        num_files = self.listWidget.count()
        for index in range(num_files):
            filepath = self.listWidget.item(index).text()
            self.add_file_to_pdf(pdf_writer, filepath)
            self.thread.progress.emit(int((index + 1) / num_files * 100))

        with open(output_filename, 'wb') as out:
            pdf_writer.write(out)

    def add_file_to_pdf(self, pdf_writer, filepath):
        _, file_extension = os.path.splitext(filepath)
        
        if file_extension.lower() == '.pdf':
            self.add_pdf_to_pdf(pdf_writer, filepath)
        elif file_extension.lower() in ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'):
            self.add_image_to_pdf(pdf_writer, filepath)
        elif file_extension.lower() == '.docx':
            self.add_docx_to_pdf(pdf_writer, filepath)
        elif file_extension.lower() == '.xml':
            self.add_xml_to_pdf(pdf_writer, filepath)
        else:
            raise ValueError(f"Unsupported file format: {file_extension}")

    def add_pdf_to_pdf(self, pdf_writer, filepath):
        pdf_reader = PdfReader(filepath)
        for page in pdf_reader.pages:
            pdf_writer.add_page(page)

    def add_image_to_pdf(self, pdf_writer, filepath):
        image = Image.open(filepath)
        pdf_bytes = BytesIO()
        image.convert("RGB").save(pdf_bytes, format='PDF')
        pdf_bytes.seek(0)
        temp_reader = PdfReader(pdf_bytes)
        for page in temp_reader.pages:
            pdf_writer.add_page(page)

    def add_docx_to_pdf(self, pdf_writer, filepath):
        # Convert DOCX to PDF using Word application
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(filepath)
        temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        temp_pdf.close()
        doc.SaveAs(temp_pdf.name, FileFormat=17)  # FileFormat=17 is for PDF
        doc.Close()
        word.Quit()

        # Add the temporary PDF to the PDF writer
        temp_reader = PdfReader(temp_pdf.name)
        for page in temp_reader.pages:
            pdf_writer.add_page(page)

        # Clean up the temporary file
        os.unlink(temp_pdf.name)

    def add_xml_to_pdf(self, pdf_writer, filepath):
        # Parse XML and create a PDF
        tree = etree.parse(filepath)
        root = tree.getroot()

        pdf_bytes = BytesIO()
        c = canvas.Canvas(pdf_bytes, pagesize=letter)
        width, height = letter
        y = height - 50  # Start from top of the page

        def add_element(element, indent=0):
            nonlocal y
            if y < 50:  # If near bottom of page, start a new page
                c.showPage()
                y = height - 50

            tag = element.tag
            text = element.text.strip() if element.text else ""
            c.drawString(50 + indent * 20, y, f"{tag}: {text}")
            y -= 20

            for child in element:
                add_element(child, indent + 1)

        add_element(root)
        c.save()

        pdf_bytes.seek(0)
        temp_reader = PdfReader(pdf_bytes)
        for page in temp_reader.pages:
            pdf_writer.add_page(page)

    def export_to_docx(self, output_filename):
        doc = Document()
        num_files = self.listWidget.count()
        for index in range(num_files):
            filepath = self.listWidget.item(index).text()
            self.add_file_to_docx(doc, filepath)
            self.thread.progress.emit(int((index + 1) / num_files * 100))

        doc.save(output_filename)

    def add_file_to_docx(self, doc, filepath):
        _, file_extension = os.path.splitext(filepath)
        
        if file_extension.lower() == '.pdf':
            self.add_pdf_to_docx(doc, filepath)
        elif file_extension.lower() in ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'):
            self.add_image_to_docx(doc, filepath)
        elif file_extension.lower() == '.docx':
            self.add_docx_to_docx(doc, filepath)
        elif file_extension.lower() == '.xml':
            self.add_xml_to_docx(doc, filepath)
        else:
            raise ValueError(f"Unsupported file format: {file_extension}")

    def add_pdf_to_docx(self, doc, filepath):
        pdf_reader = PdfReader(filepath)
        for page in pdf_reader.pages:
            doc.add_paragraph(page.extract_text())

    def add_image_to_docx(self, doc, filepath):
        doc.add_picture(filepath, width=Inches(6))

    def add_docx_to_docx(self, doc, filepath):
        src_doc = Document(filepath)
        for element in src_doc.element.body:
            doc.element.body.append(element)

    def add_xml_to_docx(self, doc, filepath):
        tree = etree.parse(filepath)
        root = tree.getroot()

        def add_element(element, indent=0):
            tag = element.tag
            text = element.text.strip() if element.text else ""
            doc.add_paragraph(f"{'  ' * indent}{tag}: {text}")

            for child in element:
                add_element(child, indent + 1)

        add_element(root)

    def clear_list(self):
        self.listWidget.clear()

    # Remove selected items from the list
    def remove_selected_items(self):
        for selectedItem in self.listWidget.selectedItems():
            self.listWidget.takeItem(self.listWidget.row(selectedItem))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileMerger()
    window.show()
    sys.exit(app.exec_())