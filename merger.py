import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QListWidget, QPushButton, QVBoxLayout, QWidget, QFileDialog, QMessageBox, QHBoxLayout, QProgressBar)
from PyQt5.QtCore import Qt
from PyPDF2 import PdfReader, PdfWriter

class PdfMerger(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Merger")
        self.setGeometry(100, 100, 800, 450) 
        
        # Main layout
        mainLayout = QVBoxLayout()
        
        # List widget to display PDF files
        self.listWidget = QListWidget()
        self.listWidget.setStyleSheet("QListWidget {background-color: #f0f0f0; border: 1px solid grey;}")
        self.listWidget.setAcceptDrops(True)
        self.listWidget.setDragEnabled(True)
        self.listWidget.setSelectionMode(QListWidget.MultiSelection)
        self.listWidget.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        
        # Progress Bar
        self.progressBar = QProgressBar()
        self.progressBar.setStyleSheet("QProgressBar {border: 2px solid grey; border-radius: 5px; text-align: center;}")
        self.progressBar.setMaximum(100)
        self.progressBar.setValue(0)
        
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
        
        # Remove selected button
        self.removeSelectedButton = QPushButton("Remove Selected")
        self.removeSelectedButton.setStyleSheet("QPushButton {background-color: #008CBA; color: white;}")
        self.removeSelectedButton.clicked.connect(self.remove_selected_items)
        
        # Add buttons to the layout
        buttonLayout.addWidget(self.mergeButton)
        buttonLayout.addWidget(self.clearButton)
        buttonLayout.addWidget(self.removeSelectedButton)
        
        # Add widgets to the main layout
        mainLayout.addWidget(self.listWidget)
        mainLayout.addWidget(self.progressBar) 
        mainLayout.addLayout(buttonLayout)
        
        # Set the main widget
        container = QWidget()
        container.setLayout(mainLayout)
        self.setCentralWidget(container)
        
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
            self.progressBar.setValue(0)
            return

        self.progressBar.setValue(0)
        try:
            pdf_writer = PdfWriter()
            num_files = self.listWidget.count()
            for index, _ in enumerate(range(self.listWidget.count()), 1):
                filepath = self.listWidget.item(index - 1).text()
                pdf_reader = PdfReader(filepath)
                for page in pdf_reader.pages:
                    pdf_writer.add_page(page)
                self.progressBar.setValue(int((index / num_files) * 100))  # Update progress
                
            with open(output_filename, 'wb') as out:
                pdf_writer.write(out)
            
            QMessageBox.information(self, "Merge Successful", f"Merged PDF saved as: {output_filename}")
            self.progressBar.setValue(100)  # Complete progress
            self.clear_list()  # clear list after a successful merge
        except Exception as e:
            QMessageBox.warning(self, "Merge Failed", "An error occurred during the merge process.")
        finally:
            self.progressBar.setValue(0)  # Reset progress bar after merge is complete or fails    

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
