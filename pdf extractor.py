import sys
import os
from PySide6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTextEdit, \
    QFileDialog, QMessageBox, QComboBox
from PySide6.QtGui import QAction, QIcon, QPixmap
from PySide6.QtCore import QByteArray
import pytesseract
from PIL import Image
import fitz  # PyMuPDF
from google.cloud import vision
from google.oauth2 import service_account
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Specify the Tesseract executable path if it's not in the PATH environment variable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Update this path if necessary

# Base64-encoded image for logo (you can replace this with your actual base64 string)
base64_image = b'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'
  # Add your base64 string here


class PDFTextExtractor(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('PDF Text Extractor')
        self.setGeometry(100, 100, 800, 600)

        # Logo
        pixmap = QPixmap()
        pixmap.loadFromData(QByteArray.fromBase64(base64_image))
        self.setWindowIcon(QIcon(pixmap))

        # Create main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.layout = QVBoxLayout(central_widget)

        # Create Text Edit
        self.text_edit = QTextEdit(self)
        self.layout.addWidget(self.text_edit)

        # Create horizontal layout for dropdown and refresh button
        h_layout = QHBoxLayout()

        # Create Dropdown for method selection
        self.dropdown = QComboBox(self)
        self.dropdown.addItem("Fitz")
        self.dropdown.addItem("Pytesseract")
        self.dropdown.addItem("Google Cloud Vision")
        h_layout.addWidget(self.dropdown)

        # Create Refresh Button
        self.refresh_button = QPushButton('Refresh', self)
        self.refresh_button.setFixedWidth(80)  # Make the button smaller
        self.refresh_button.clicked.connect(self.refresh_application)
        h_layout.addWidget(self.refresh_button)

        self.layout.addLayout(h_layout)

        # Create Buttons
        self.open_button = QPushButton('Open PDF', self)
        self.open_button.clicked.connect(self.open_pdf)
        self.layout.addWidget(self.open_button)

        self.save_button = QPushButton('Save to PDF File', self)
        self.save_button.clicked.connect(self.save_text)
        self.layout.addWidget(self.save_button)

        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5; /* White background */
            }
            QLineEdit {
                background-color: white;
                border: 1px solid #ccc; /* Light Gray border */
                padding: 4px;
                width: 150px; /* Adjust as needed */
            }
            QPushButton {
                background-color: #0d6efd; /* Dark Blue button */
                color: white;
                border: none;
                padding: 6px 12px;
                border-radius: 3px;
                width: 150px;
            }
            QPushButton:hover {
                background-color: #0951ba; /* Slightly darker grey on hover */
            }
        """)

        # Create menu bar
        menubar = self.menuBar()
        file_menu = menubar.addMenu('File')
        open_action = QAction(QIcon(None), 'Open PDF', self)
        open_action.triggered.connect(self.open_pdf)
        file_menu.addAction(open_action)

        save_action = QAction(QIcon(None), 'Save to PDF File', self)
        save_action.triggered.connect(self.save_text)
        file_menu.addAction(save_action)

        refresh_action = QAction(QIcon(None), 'Refresh', self)
        refresh_action.triggered.connect(self.refresh_application)
        file_menu.addAction(refresh_action)
        # Add an About action to the Help menu
        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_about_dialog)
        menubar.addAction(about_action)

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close_application)
        menubar.addAction(exit_action)

    def show_about_dialog(self):
        about_text = "PDF Data Extractor\n\nVersion: 1.00\n\n"
        QMessageBox.about(self, "About PDF Data Extractor", about_text)
    def open_pdf(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Open PDF File", "", "PDF Files (*.pdf);;All Files (*)",
                                                   options=options)
        if file_name:
            self.extract_text_from_pdf(file_name)

    def close_application(self):
        # Close the socket when the application is closed
        self.close()
        sys.exit()
    def extract_text_from_pdf(self, file_name):
        selected_method = self.dropdown.currentText()
        try:
            if selected_method == "Fitz":
                self.extract_text_with_fitz(file_name)
            elif selected_method == "Pytesseract":
                self.extract_text_with_pytesseract(file_name)
            elif selected_method == "Google Cloud Vision":
                self.extract_text_with_google_cloud_vision(file_name)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to extract text: {e}")

    def extract_text_with_fitz(self, file_name):
        doc = fitz.open(file_name)
        text_per_page = []
        output_dir = self.create_output_directory(file_name)
        for page_num, page in enumerate(doc):
            text = page.get_text()
            text_per_page.append(text)
            self.save_page_text(output_dir, page_num, text)
        self.text_edit.setPlainText("\n\n".join(text_per_page))

    def extract_text_with_pytesseract(self, file_name):
        doc = fitz.open(file_name)
        extracted_text = []

        # Iterate through each page and extract text
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Use pytesseract to do OCR on the image of the page
            page_text = pytesseract.image_to_string(img)
            extracted_text.append(page_text)

            # Save the extracted text to a text file
            text_file_name = f"{os.path.splitext(file_name)[0]}_page_{page_num + 1}.txt"
            with open(text_file_name, "w", encoding="utf-8") as text_file:
                text_file.write(page_text)

        # Update the QTextEdit with extracted text
        self.text_edit.setPlainText("\n\n".join(extracted_text))

    def extract_text_with_google_cloud_vision(self, file_name):
        credentials = service_account.Credentials.from_service_account_file(
            r'F:\Everything\PDF_Extractor\eco-emissary- -p5-b215ef8ba65d.json')
        client = vision.ImageAnnotatorClient(credentials=credentials)

        doc = fitz.open(file_name)
        text_per_page = []
        output_dir = self.create_output_directory(file_name)
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img_byte_arr = img.tobytes()
            image = vision.Image(content=img_byte_arr)
            response = client.text_detection(image=image)
            text = response.full_text_annotation.text
            text_per_page.append(text)
            self.save_page_text(output_dir, page_num, text)
        self.text_edit.setPlainText("\n\n".join(text_per_page))

    def save_page_text(self, output_dir, page_num, text):
        file_path = os.path.join(output_dir, f"page_{page_num + 1}.txt")
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(text)

    def create_output_directory(self, file_name):
        base_name = os.path.basename(file_name)
        dir_name = os.path.splitext(base_name)[0]
        output_dir = os.path.join(os.path.dirname(file_name), dir_name)
        os.makedirs(output_dir, exist_ok=True)
        return output_dir

    def save_text(self):
        options = QFileDialog.Options()
        dir_name = QFileDialog.getExistingDirectory(self, "Save PDF File Directory", options=options)
        if dir_name:
            try:
                file_name = os.path.join(dir_name, "output.pdf")
                full_text = self.text_edit.toPlainText().split("\n\n")
                c = canvas.Canvas(file_name, pagesize=letter)
                width, height = letter
                for page_text in full_text:
                    text_object = c.beginText(40, height - 40)
                    for line in page_text.splitlines():
                        text_object.textLine(line)
                    c.drawText(text_object)
                    c.showPage()
                c.save()
                QMessageBox.information(self, "Success", "Text successfully saved.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not save file: {e}")

    def refresh_application(self):
        self.text_edit.clear()
        self.dropdown.setCurrentIndex(0)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PDFTextExtractor()
    ex.show()
    sys.exit(app.exec())
