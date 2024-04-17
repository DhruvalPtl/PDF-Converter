import sys
from PyQt6.QtWidgets import QLabel, QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QGridLayout, QPushButton, QWidget, QSpacerItem, QSizePolicy
from PyQt6.QtGui import QIcon, QPixmap
from PyQt6.QtCore import QSize, Qt
from secondwindow import secondWindow

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("File Format Converter")
        self.setGeometry(300, 150, 1500, 900)

        self.setup_ui()

    def setup_ui(self):
        self.setup_background()
        self.setup_layouts()

    def setup_background(self):
        image_path = "D:/Python/project 1/Pdfconverter/images/2.jpg"
        background_label = QLabel(self)
        background_label.setPixmap(QPixmap(image_path))
        background_label.setGeometry(0, 0, 1500, 1000)

    def setup_layouts(self):
        self.main_layout = QVBoxLayout()
        self.central_widget = QWidget()
        self.central_widget.setLayout(self.main_layout)
        self.setCentralWidget(self.central_widget)
        self.main_layout.setSpacing(10)

        self.create_top_layout()
        self.create_middle_layout()
        self.create_bottom_layout()  

    def create_top_layout(self):
        self.top_layout_container = QWidget()
        
        self.top_layout_container.setStyleSheet("background-color: #f8f8f8;")

        top_layout = QHBoxLayout()

        self.pdf_label = QLabel("PDF Converter")
        self.pdf_label.setStyleSheet("padding-bottom:7px;font-size:35px;font-weight:600;color:#161616;")
        top_layout.addWidget(self.pdf_label)

        buttons = ["Compress PDF", "Merge PDF", "Split PDF"]
        for index,btn_text in enumerate(buttons):
            button = QPushButton(btn_text)
            button.setStyleSheet("QPushButton { margin:20px; font-size:20px;font-weight:500;color:#161616;background-color: transparent;border: none;}QPushButton:hover {color: #e5322d;}")
            button.setCursor(Qt.CursorShape.PointingHandCursor)
            button.setToolTip(btn_text)
            button.clicked.connect(lambda _, name=buttons[index]: self.on_button_clicked(name))
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

    def create_middle_layout(self):
        self.sub_layout3 = QVBoxLayout()
        self.title = QLabel()
        self.title.setText(f"<h1>All the essential PDF tools consolidated into one convenient platform.</h1><h2>Every tool you need to use PDFs, at your fingertips. All are easy to use!<br> Compress, Merge, Split, Convert PDFs with just a few clicks.</h2><br><br>")
        self.title.setStyleSheet("font-size: 15px;font-weight: 600;color: #33333b;")
        self.title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.sub_layout3.addWidget(self.title)
        
        self.main_layout.addLayout(self.sub_layout3)

    def create_bottom_layout(self):
        self.sub_layout2 = QGridLayout()
        
        self.sub_layout2.setSpacing(26)
        
        icons = [
            "D:/Python/project 1/Pdfconverter/images/compress.png",
            "D:/Python/project 1/Pdfconverter/images/merge.png",
            "D:/Python/project 1/Pdfconverter/images/split.png",
            "D:/Python/project 1/Pdfconverter/images/htmltopdf.png",
            "D:/Python/project 1/Pdfconverter/images/pdftoexcel.png",
            "D:/Python/project 1/Pdfconverter/images/exceltopdf.png",
            "D:/Python/project 1/Pdfconverter/images/pdftoword.png",
            "D:/Python/project 1/Pdfconverter/images/wordtopdf.png",
            "D:/Python/project 1/Pdfconverter/images/pdftopp.png",
            "D:/Python/project 1/Pdfconverter/images/pptopdf.png",
            "D:/Python/project 1/Pdfconverter/images/pdftojpg.png",
            "D:/Python/project 1/Pdfconverter/images/jpgtopdf.png",
        ]
        button_names = ["Compress PDF","Merge PDF","Split PDF","HTML to PDF","PDF to Excel","Excel to PDF","PDF to Word","Word to PDF","PDF to PowerPoint","PowerPoint to PDF","PDF to JPG","JPG to PDF"]
        for i in range(3):
            for j in range(4):  # 4 columns
                index = i * 4 + j
                if index < len(icons):
                    button = QPushButton()
                    button.setIcon(QIcon(icons[index]))
                    button.setIconSize(QSize(65, 65))  # Set the size of the icon
                    button.setText(f"{button_names[index]}")
                    button.setFixedSize(330, 150)
                    button.setStyleSheet("font-weight: 600;font-size:25px;color: #33333b;")
                    button.setToolTip(button_names[index])
                    self.sub_layout2.addWidget(button, i, j)
                    button.clicked.connect(lambda _, name=button_names[index]: self.on_button_clicked(name))
                    self.sub_layout2.addWidget(button, i, j)
                   
        self.main_layout.addLayout(self.sub_layout2)
        
        spacer_item = QSpacerItem(20, 40, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)
        self.main_layout.addItem(spacer_item)

    def on_button_clicked(self, button_name):
        window_types = {
            "Merge PDF": "Merge",
            "Compress PDF": "Compression",
            "Split PDF": "Split",
            "HTML to PDF": "Html2pdf",
            "PDF to Excel": "pdf2excel",
            "Excel to PDF": "excel2pdf",
            "PDF to Word": "pdf2word",
            "Word to PDF": "word2pdf",
            "PDF to PowerPoint": "pdf2pptx",
            "PowerPoint to PDF": "pptx2pdf",
            "PDF to JPG": "pdf2jpg",
            "JPG to PDF": "jpg2pdf"
            ""
        }
        
        if button_name in window_types:
            window_type = window_types[button_name]
            self.view_page = secondWindow(window_type)
            self.view_page.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
