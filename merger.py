import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QListWidget, QPushButton, QVBoxLayout, QWidget, QFileDialog, QMessageBox, QHBoxLayout)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor
from PyPDF2 import PdfReader, PdfWriter

class PdfMerger(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Merger")
        self.setGeometry(100, 100, 800, 400)  # Use a generous size for the window
        
        # Main layout
        mainLayout = QVBoxLayout()
        
        # List widget to display PDF files
        self.listWidget = QListWidget()
        self.listWidget.setStyleSheet("QListWidget {background-color: #f0f0f0; border: 1px solid grey;}")
        self.listWidget.setAcceptDrops(True)
        self.listWidget.setDragEnabled(True)
        self.listWidget.setSelectionMode(QListWidget.MultiSelection)
        
        # Enable the dropping functionality on the QListWidget
        self.listWidget.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        
        # Button layout
        buttonLayout = QHBoxLayout()
        
        # Merge button
        self.mergeButton = QPushButton("Merge")
        self.mergeButton.setStyleSheet("QPushButton {background-color: #4CAF50; color: white;}")
        self.mergeButton.clicked.connect(self.merge_pdfs)
        
        # Clear button
        self.clearButton = QPushButton("Clear")
        self.clearButton.setStyleSheet("QPushButton {background-color: #f44336; color: white;}")
        self.clearButton.clicked.connect(self.clear_list)
        
        # Remove selected button with color
        self.removeSelectedButton = QPushButton("Remove Selected")
        self.removeSelectedButton.setStyleSheet("QPushButton {background-color: #008CBA; color: white;}")
        self.removeSelectedButton.clicked.connect(self.remove_selected_items)
        
        buttonLayout.addWidget(self.mergeButton)
        buttonLayout.addWidget(self.clearButton)
        buttonLayout.addWidget(self.removeSelectedButton)
        
        mainLayout.addWidget(self.listWidget)
        mainLayout.addLayout(buttonLayout)
        
        container = QWidget()
        container.setLayout(mainLayout)
        self.setCentralWidget(container)
        
        # Enable dropping on the window
        self.setAcceptDrops(True)
        
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        for file in files:
            if file.lower().endswith('.pdf'):
                self.listWidget.addItem(file)
        event.acceptProposedAction()

    def merge_pdfs(self):
        if self.listWidget.count() == 0:
            QMessageBox.information(self, "No PDFs Selected", "Please add PDF files to merge.")
            return
        
        output_filename = QFileDialog.getSaveFileName(self, "Save Merged PDF", "", "PDF Files (*.pdf)")[0]
        if not output_filename:
            return
        
        try:
            pdf_writer = PdfWriter()
            for index in range(self.listWidget.count()):
                filepath = self.listWidget.item(index).text()
                pdf_reader = PdfReader(filepath)
                for page in range(len(pdf_reader.pages)):
                    pdf_writer.add_page(pdf_reader.pages[page])
            
            with open(output_filename, 'wb') as out:
                pdf_writer.write(out)
            
            QMessageBox.information(self, "Merge Successful", f"Merged PDF saved as: {output_filename}")
            self.clear_list()  # Clear the list after a successful merge
        except Exception as e:
            QMessageBox.warning(self, "Merge Failed", "An error occurred during the merge process.")

    def clear_list(self):
        # Clear all items from the list widget
        self.listWidget.clear()

    def remove_selected_items(self):
        for selectedItem in self.listWidget.selectedItems():
            self.listWidget.takeItem(self.listWidget.row(selectedItem))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PdfMerger()
    window.show()
    sys.exit(app.exec_())
