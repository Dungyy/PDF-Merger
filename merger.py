import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QListWidget, QPushButton, QVBoxLayout, QWidget, QFileDialog
from PyQt5.QtCore import Qt
from PyPDF2 import PdfReader, PdfWriter

class PdfMerger(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Merger")
        self.setGeometry(100, 100, 600, 300)
        
        # Main layout
        layout = QVBoxLayout()
        
        # List widget to display PDF files
        self.listWidget = QListWidget()
        self.listWidget.setAcceptDrops(True)
        self.listWidget.setDragEnabled(True)
        
        # Enable the dropping functionality on the QListWidget
        self.listWidget.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        
        # Merge button
        self.mergeButton = QPushButton("Merge")
        self.mergeButton.clicked.connect(self.merge_pdfs)
        
        # Clear button
        self.clearButton = QPushButton("Clear")
        self.clearButton.clicked.connect(self.clear_list)
        
        layout.addWidget(self.listWidget)
        layout.addWidget(self.mergeButton)
        layout.addWidget(self.clearButton)
        
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        
        # Enable dropping on the window
        self.setAcceptDrops(True)
        
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()  # accept the event to allow drop
        else:
            super().dragEnterEvent(event)

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        for file in files:
            if file.lower().endswith('.pdf'):
                self.listWidget.addItem(file)
        event.acceptProposedAction()

    def merge_pdfs(self):
        pdf_writer = PdfWriter()
        for index in range(self.listWidget.count()):
            filepath = self.listWidget.item(index).text()
            pdf_reader = PdfReader(filepath)
            for page in range(len(pdf_reader.pages)):
                pdf_writer.add_page(pdf_reader.pages[page])
        
        output_filename = QFileDialog.getSaveFileName(self, "Save Merged PDF", "", "PDF Files (*.pdf)")[0]
        if output_filename:
            with open(output_filename, 'wb') as out:
                pdf_writer.write(out)
            print(f"Merged PDF saved as: {output_filename}")

    def clear_list(self):
        # Clear all items from the list widget
        self.listWidget.clear()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PdfMerger()
    window.show()
    sys.exit(app.exec_())
