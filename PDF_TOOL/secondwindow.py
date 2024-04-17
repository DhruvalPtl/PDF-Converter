import sys
import os
import shutil
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QPushButton, QWidget,
    QLabel, QFileDialog, QProgressBar, QMessageBox, QSpacerItem, QSizePolicy, QSpinBox, QCheckBox, QLineEdit
)
from PyQt6.QtGui import QIcon, QPixmap, QPainter, QPen, QFont, QColor, QRegularExpressionValidator
from PyQt6.QtCore import Qt, QRegularExpression

from pdfconvter import PDFConverter
class DragDropArea(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setAutoFillBackground(True)
        self.setFixedSize(600, 300)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setPen(QPen(QColor(Qt.GlobalColor.black), 2, Qt.PenStyle.DashLine))
        painter.drawRect(self.rect())
        
        text = "Choose a file or Drag it Here"
        font = QFont("Arial", 24)
        painter.setFont(font)
        text_rect = painter.boundingRect(self.rect(), Qt.AlignmentFlag.AlignCenter, text)
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, text)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            for url in urls:
                if url.isLocalFile():
                    file_extension = url.toLocalFile().split(".")[-1].lower()
                    if file_extension in ["pdf", "xlsx", "xls", "docx", "doc", "pptx", "ppt", "html", "htm", "jpg"]:
                        event.accept()
                        return
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            files = [url.toLocalFile() for url in event.mimeData().urls()]
            for file in files:
                destination_path = "D:/Python/project 1/Pdfconverter/TEMP/" + os.path.basename(file)
                shutil.copy(file, destination_path)
            event.accept()
        else:
            event.ignore()

class secondWindow(QMainWindow):
    def __init__(self,value):
        self.value = value
        self.obj = PDFConverter()
        super().__init__()
        self.setWindowTitle("File Format Converter")
        self.setGeometry(300, 150, 1500, 900)
        
        image = QPixmap("D:/Python/project 1/Pdfconverter/images/2.jpg")
        background_label = QLabel(self)
        background_label.setPixmap(image)
        background_label.setGeometry(0, 0, 1500, 1000)

        self.main_layout = QVBoxLayout()
        self.central_widget = QWidget()
        self.central_widget.setLayout(self.main_layout)
        self.setCentralWidget(self.central_widget)

        self.top_layout()
        self.middle_layout()
        self.middle1_layout()
        self.bottom_layout()

    def top_layout(self):
        self.top_layout_container = QWidget()
        self.top_layout_container.setStyleSheet("background-color: #f8f8f8;")

        top_layout = QHBoxLayout()

        self.pdf_label = QLabel("PDF Converter")
        self.pdf_label.setStyleSheet("padding-bottom:7px;font-size:35px;font-weight:600;color:#161616;")
        top_layout.addWidget(self.pdf_label)

        buttons = ["Compress PDF", "Merge PDF", "Split PDF"]
        for btn_text in buttons:
            button = QPushButton(btn_text)
            button.setStyleSheet("QPushButton { margin:20px; font-size:20px;font-weight:500;color:#161616;background-color: transparent;border: none;}QPushButton:hover {color: #e5322d;}")
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            button.setToolTip(btn_text)
            top_layout.addWidget(button)

        spacer_item = QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        top_layout.addItem(spacer_item)
        
        self.button4 = QPushButton("Login")
        self.button4.setStyleSheet("QPushButton { margin:20px; font-size:20px;font-weight:500;color:#161616;background-color: transparent;border: none;}QPushButton:hover {color: #e5322d;}") 
        self.button4.setCursor(Qt.CursorShape.PointingHandCursor)
        self.button4.setToolTip("Login")
        top_layout.addWidget(self.button4)
        
        sign_up_button = QPushButton("Sign Up")
        sign_up_button.setStyleSheet("QPushButton { padding:5px;font-size:20px;font-weight:600;color:white;background-color:#e5322d;border-radius:11px;}QPushButton:hover {background-color: #bd060a;}")
        top_layout.addWidget(sign_up_button)
        
        self.top_layout_container.setLayout(top_layout)
        self.main_layout.addWidget(self.top_layout_container)
    
    def middle_layout(self):
        self.middle = QVBoxLayout()
        self.title = QLabel()
        if self.value == "Merge":
            self.title.setText(f"<h1>Merge PDF files</h1><h6>Combine PDFs in the order you want with the easiest PDF merger available.</h6><br>")
        elif self.value == "Compression":
            self.title.setText(f"<h1>Compress PDF file</h1><h6>Reduce file size while optimizing for maximal PDF quality.</h6><br>")
        elif self.value == "Html2pdf":
            self.title.setText(f"<h1>HTML to PDF</h1><h6>Convert web pages to PDF documents with high accuracy</h6><br>")    
        elif self.value == "pdf2excel":
            self.title.setText(f"<h1>Convert PDF to EXCEL</h1><h6>Convert PDF Data to EXCEL Spreadsheets.</h6><br>")    
        elif self.value == "excel2pdf":
            self.title.setText(f"<h1>Convert EXCEL to PDF</h1><h6>Make EXCEL spreadsheets easy to read by converting them to PDF.</h6><br>")    
        elif self.value == "pdf2word":
            self.title.setText(f"<h1>PDF to WORD Converter</h1><h6>Convert your PDF to WORD documents with incredible accuracy.</h6><br>")    
        elif self.value == "word2pdf":
            self.title.setText(f"<h1>Convert WORD to PDF</h1><h6>Make DOC and DOCX files easy to read by converting them to PDF.</h6><br>")    
        elif self.value == "pdf2pptx":
            self.title.setText(f"<h1>Convert PDF to POWERPOINT</h1><h6>Convert your PDFs to POWERPOINT.</h6><br>")    
        elif self.value == "pptx2pdf":
            self.title.setText(f"<h1>Convert POWERPOINT to PDF</h1><h6>Make PPT and PPTX slideshows easy to view by converting them to PDF.</h6><br>")    
        elif self.value == "pdf2jpg":
            self.title.setText(f"<h1>PDF to JPG</h1><h6>Convert each PDF page into a JPG or extract all images contained in a PDF.</h6><br>")    
        elif self.value == "jpg2pdf":
            self.title.setText(f"<h1>JPG to PDF</h1><h6>Convert JPG images to PDF in seconds. Easily adjust orientation and margins.</h6><br>")    
        self.title.setStyleSheet("font-size:22px;font-weight: 600;color: #33333b;")
        self.title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.middle.addWidget(self.title)
        self.main_layout.addLayout(self.middle)
        
    def middle1_layout(self):
        middle_layout = QHBoxLayout()
        self.main_layout.addLayout(middle_layout)

        self.drag_drop_area = DragDropArea()
        middle_layout.addWidget(self.drag_drop_area)

    def bottom_layout(self):
        self.bottom_layout = QVBoxLayout()
        self.main_layout.addLayout(self.bottom_layout)

        title_label = QLabel("<h1>Or</h1>")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.bottom_layout.addWidget(title_label)

        self.input_button = QPushButton("Select File")
        self.input_button.setFixedWidth(1000)
        self.input_button.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                border: none;
                color: white;
                padding: 10px 24px;
                text-align: center;
                text-decoration: none;
                font-size: 40px;
                margin: 4px 0px 2px 500px;
                border-radius: 12px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        ''')
        self.input_button.clicked.connect(self.select_input_file)
        self.bottom_layout.addWidget(self.input_button)

        if self.value == "Split":
           self.split_layout()
            
        convert_layout = QHBoxLayout()

        self.convert_button = QPushButton("Convert File")
        self.convert_button.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                border: none;
                color: white;
                padding: 10px 24px;
                text-align: center;
                text-decoration: none;
                font-size: 20px;
                margin: 4px 0px 2px 10px;
                border-radius: 12px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        ''')
        self.convert_button.clicked.connect(self.convert_file)
        convert_layout.addWidget(self.convert_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setRange(0, 100)
        convert_layout.addWidget(self.progress_bar)

        self.bottom_layout.addLayout(convert_layout)

        download_layout = QHBoxLayout()
        
        self.download_button = QPushButton("Download File")
        self.download_button.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                border: none;
                color: white;
                padding: 10px 24px;
                text-align: center;
                text-decoration: none;
                font-size: 20px;
                margin: 4px 180px 70px 180px;
                border-radius: 12px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        ''')
        self.download_button.clicked.connect(self.download_file)
        download_layout.addWidget(self.download_button)
        
        self.quit_button = QPushButton("Quit", clicked=self.close)
        self.quit_button.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                border: none;
                color: white;
                padding: 10px 24px;
                text-align: center;
                text-decoration: none;
                font-size: 20px;
                margin: 4px 180px 70px 180px;
                border-radius: 12px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        ''')
        download_layout.addWidget(self.quit_button)

        self.bottom_layout.addLayout(download_layout)

    def split_layout(self):
            layout = QHBoxLayout()
            self.bottom_layout.addLayout(layout)
            layout.addSpacing(20)
            layout.setContentsMargins(60,30,60,30)

            # Add label indicating the purpose of the spin box
            label_from = QLabel("From page:")
            label_from.setStyleSheet("font-size:20px;font-weight: bold; color: #333;")
            layout.addWidget(label_from)

            # Create a spin box for the 'from' page
            self.spin_box_from = QSpinBox()
            self.spin_box_from.setValue(0)  # Set the initial value
            self.spin_box_from.setSingleStep(1)
            self.spin_box_from.setStyleSheet("font-size: 20px;")
            layout.addWidget(self.spin_box_from)

            # Add label indicating the range
            label_range = QLabel("          To")
            label_range.setStyleSheet("font-size: 20px;font-weight: bold; color: #333;")
            layout.addWidget(label_range)

            # Create a spin box for the 'to' page
            self.spin_box_to = QSpinBox()
            self.spin_box_to.setValue(0)  # Set the initial value
            self.spin_box_to.setSingleStep(1)  # Set the step value
            self.spin_box_to.setStyleSheet("font-size: 20px;")
            layout.addWidget(self.spin_box_to)

            title_label = QLabel("<h1>Or</h1>")
            title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(title_label)
            
            self.checkbox_single_page = QCheckBox("Single Page")
            self.checkbox_single_page.setStyleSheet("font-size:20px;font-weight: bold; color: #333;")
            layout.addWidget(self.checkbox_single_page)
            
            title_label = QLabel("<h1>Or</h1>")
            title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(title_label)
            
            # Add input field for pages to extract
            label_pages_to_extract = QLabel("Pages to extract:")
            label_pages_to_extract.setStyleSheet("font-size: 20px;font-weight: bold; color: #333;")
            layout.addWidget(label_pages_to_extract)

            digit_comma_regex = QRegularExpression("[0-9,]+")
            validator = QRegularExpressionValidator(digit_comma_regex)

            self.input_pages_to_extract = QLineEdit()
            self.input_pages_to_extract.setFixedWidth(150)
            self.input_pages_to_extract.setStyleSheet("font-size: 20px;padding: 5px; border: 1px solid #ccc; border-radius: 5px;")
            self.input_pages_to_extract.setValidator(validator)
            layout.addWidget(self.input_pages_to_extract)

    def convert_file(self):
        source_path = "D:/Python/project 1/Pdfconverter/TEMP1/"
        source_path1 = "D:/Python/project 1/Pdfconverter/TEMP/"
        if os.listdir(source_path):
            shutil.rmtree(source_path, ignore_errors=True)
        if not os.listdir(source_path1):
            self.display_message("No PDF files found select first")
        else:
            self.progress_bar.setValue(0)
            actions = {
                "Merge": (self.obj.pdfmerger, "The files have been merged successfully!"),
                "Compression": (self.obj.pdfcompress, "The files have been compressed successfully!"),
                "Split": (self.perform_split_action, "The files have been split successfully!"),
                "Html2pdf": (self.obj.html2pdf, "The files have been converted to PDF successfully!"),
                "pdf2excel": (self.obj.pdf2excel, "The files have been converted to Excel successfully!"),
                "excel2pdf": (self.obj.excel2pdf, "The files have been converted to PDF successfully!"),
                "pdf2word": (self.obj.pdf2word, "The files have been converted to Word successfully!"),
                "word2pdf": (self.obj.word2pdf, "The files have been converted to PDF successfully!"),
                "pdf2pptx": (self.obj.pdf2pptx, "The files have been converted to PowerPoint successfully!"),
                "pptx2pdf": (self.obj.pptx2pdf, "The files have been converted to PDF successfully!"),
                "pdf2jpg": (self.obj.pdf2images, "The files have been converted to JPG successfully!"),
                "jpg2pdf": (self.obj.image2pdf, "The files have been converted to PDF successfully!")
            }

            total_actions = len(actions)
            current_action = 0

            for action_name, (function, message) in actions.items():
                current_action += 1
                progress_percentage = (current_action / total_actions) * 100
                self.progress_bar.setValue(int(progress_percentage))

                if function:
                    if action_name == self.value:
                        function()
                        self.display_message(message)
                        self.obj.emptydir()
            self.progress_bar.setValue(100)
                
    def perform_split_action(self):
        from_page = self.spin_box_from.value()
        to_page = self.spin_box_to.value()
        input_text = self.input_pages_to_extract.text()
        input_list = input_text.split(',')
        input_list = [int(item.strip()) for item in input_list if item.strip()]
        if  self.checkbox_single_page.isChecked():
            self.obj.single_pages()
        if input_list:
            self.obj.extract_pages(input_list) 
        if (from_page and to_page) != 0:  
            self.obj.split_range(int(from_page), int(to_page))
        self.obj.emptydir()
                    
    def select_input_file(self):
        source_path = "D:/Python/project 1/Pdfconverter/TEMP/"
        file_dialog = QFileDialog()
        file_paths, _ = file_dialog.getOpenFileNames(self, "Select files to convert", "", "ALL(*.*);;PDF Files (*.pdf);;Excel Files (*.xlsx);;Word Files (*.docx);;PowerPoint Files (*.pptx);;HTML Files (*.html);;JPEG Files (*.jpg);;PNG Files(*.png)")
        if file_paths:
            for file_path in file_paths:
                destination_path = os.path.join(source_path, os.path.basename(file_path))
                shutil.copy(file_path, destination_path)

    def download_file(self):
        source_dir = "D:/Python/project 1/Pdfconverter/TEMP1"
        file_exte = [".pdf", ".xlsx", ".docx", ".pptx", ".jpg", ".png"]
        pdf_files = [i for i in os.listdir(source_dir) if os.path.splitext(i)[1] in file_exte]
        
        if not pdf_files:
            self.display_message(f"No files found. Try to convert first.")
            return

        if len(pdf_files) == 1:
            default_name = pdf_files[0]
            file_path, _ = QFileDialog.getSaveFileName(self, "Save File", default_name, "ALL Files (*.*);;PDF Files (*.pdf);;EXCEL Files (*.xlsx);;WORD Files (*.docx);;Power Point File(*.pptx);;JPG File(*.jpg)")
            if file_path:
                shutil.copy(os.path.join(source_dir, pdf_files[0]), file_path)
                self.display_message("Your file has been downloaded.")
                shutil.rmtree(source_dir, ignore_errors=True)
                os.makedirs(source_dir, exist_ok=True)
                self.progress_bar.setValue(0)
        else:
            default_name1 = "NewFile.zip"
            file_path, _ = QFileDialog.getSaveFileName(self, "Save File", default_name1, "ZIP Files (*.zip);")
            if file_path:
                shutil.make_archive(file_path[:-4], 'zip', source_dir)
                self.display_message("Your files have been downloaded as a ZIP archive.")
                shutil.rmtree(source_dir, ignore_errors=True)
                os.makedirs(source_dir, exist_ok=True)
                self.progress_bar.setValue(0)
        
        self.progress_bar.setValue(0)

        if self.value == "Split":
            self.spin_box_from.setValue(0)
            self.spin_box_to.setValue(0)
            self.checkbox_single_page.setChecked(False)
            self.input_pages_to_extract.clear()
        
    def display_message(self, message):
        msg_box = QMessageBox()
        msg_box.setText(message)
        msg_box.exec()

