import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QListWidget, QPushButton, QVBoxLayout, QWidget, 
    QFileDialog, QMessageBox, QHBoxLayout, QProgressBar, QLabel, QGridLayout, QListWidgetItem
)
from PyQt5.QtCore import Qt
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image
from io import BytesIO
from docx import Document  # For handling DOCX files
from lxml import etree  # For handling XML files
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

class PdfMerger(QMainWindow):
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
        
        descriptionLabel = QLabel("Drag and drop files or use the buttons below to add files. You can merge PDFs, images, DOCX, and XML into a single PDF file.")
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
        self.mergeButton = QPushButton("Merge Files")
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

    # Drop Event (for the entire window)
    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    file_path = url.toLocalFile()
                    if file_path.lower().endswith(('.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.docx', '.xml')):
                        self.listWidget.addItem(file_path)
            event.acceptProposedAction()
        else:
            event.ignore()

    # File Explorer to open PDF, DOCX, Image, or XML files
    def open_file_explorer(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Files", "", 
                                                "All Files (*);;PDF Files (*.pdf);;Word Files (*.docx);;Image Files (*.jpg *.jpeg *.png *.bmp *.gif *.tiff);;XML Files (*.xml)")
        if files:
            for file in files:
                if file.lower().endswith(('.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.docx', '.xml')):
                    self.listWidget.addItem(file)

    # Convert DOCX to PDF and return BytesIO
    def convert_docx_to_pdf(self, filepath):
        output = BytesIO()
        c = canvas.Canvas(output, pagesize=letter)
        doc = Document(filepath)
        width, height = letter
        text_object = c.beginText(40, height - 40)
        text_object.setFont("Helvetica", 12)

        # Read each paragraph from the DOCX and add it to the PDF
        for para in doc.paragraphs:
            text_object.textLine(para.text)
        c.drawText(text_object)
        c.showPage()
        c.save()

        output.seek(0)
        return output

    # Convert XML to PDF and return BytesIO
    def convert_xml_to_pdf(self, filepath):
        output = BytesIO()
        c = canvas.Canvas(output, pagesize=letter)
        tree = etree.parse(filepath)
        root = tree.getroot()
        width, height = letter
        text_object = c.beginText(40, height - 40)
        text_object.setFont("Helvetica", 12)

        # Recursively traverse and add XML data
        def traverse_xml(node):
            text_object.textLine(etree.tostring(node, pretty_print=True, encoding='unicode'))

        traverse_xml(root)
        c.drawText(text_object)
        c.showPage()
        c.save()

        output.seek(0)
        return output

    # Merge the selected PDFs, images, DOCX, and XML files
    def merge_files(self):
        if self.listWidget.count() == 0:
            QMessageBox.information(self, "No Files Selected", "Please add files to merge.")
            return
        
        output_filename = QFileDialog.getSaveFileName(self, "Save Merged PDF", "", "PDF Files (*.pdf)")[0]
        if not output_filename:
            return

        self.progressBar.setValue(0)
        try:
            pdf_writer = PdfWriter()
            num_files = self.listWidget.count()
            for index in range(num_files):
                filepath = self.listWidget.item(index).text()
                if filepath.lower().endswith('.pdf'):
                    # Handle PDF files
                    pdf_reader = PdfReader(filepath)
                    for page in pdf_reader.pages:
                        pdf_writer.add_page(page)
                elif filepath.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff')):
                    # Convert image to PDF and add to the writer
                    image = Image.open(filepath)
                    pdf_bytes = BytesIO()
                    image.convert("RGB").save(pdf_bytes, format='PDF')  # Save image as PDF in memory
                    pdf_bytes.seek(0)
                    temp_reader = PdfReader(pdf_bytes)
                    for page in temp_reader.pages:
                        pdf_writer.add_page(page)
                elif filepath.lower().endswith('.docx'):
                    # Convert DOCX to PDF
                    pdf_bytes = self.convert_docx_to_pdf(filepath)
                    temp_reader = PdfReader(pdf_bytes)
                    for page in temp_reader.pages:
                        pdf_writer.add_page(page)
                elif filepath.lower().endswith('.xml'):
                    # Convert XML to PDF
                    pdf_bytes = self.convert_xml_to_pdf(filepath)
                    temp_reader = PdfReader(pdf_bytes)
                    for page in temp_reader.pages:
                        pdf_writer.add_page(page)

                self.progressBar.setValue(int((index + 1) / num_files * 100))  # Update progress

            with open(output_filename, 'wb') as out:
                pdf_writer.write(out)

            QMessageBox.information(self, "Merge Successful", f"Merged PDF saved as: {output_filename}")
            self.listWidget.clear()  # Clear the list after successful merge

        except Exception as e:
            QMessageBox.warning(self, "Merge Failed", f"An error occurred during the merge process: {e}")

        finally:
            self.progressBar.setValue(0)  # Reset progress bar after merge is complete or fails

    # Clear the list of selected files
    def clear_list(self):
        self.listWidget.clear()

    # Remove selected items from the list
    def remove_selected_items(self):
        for selectedItem in self.listWidget.selectedItems():
            self.listWidget.takeItem(self.listWidget.row(selectedItem))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PdfMerger()
    window.show()
    sys.exit(app.exec_())
